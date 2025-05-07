import os
from flask import Flask, request, jsonify
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import time

app = Flask(__name__)

# === CONFIGURAÇÃO DO GOOGLE SHEETS ===
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDENTIALS_FILE = os.getenv("CREDENTIALS_FILE", "credenciais.json")
PLANILHA_ID = os.getenv("PLANILHA_ID")

# === FUNÇÃO PARA EXTRAI OS ITENS ===
def extrair_itens(texto):
    palavras_chave = ['roteador', 'notebook', 'monitor', 'cabo', 'fonte']
    linhas = texto.lower().split('\n')
    itens = []
    for linha in linhas:
        for palavra in palavras_chave:
            if palavra in linha:
                itens.append(linha.strip())
    return itens

# === FUNÇÃO DE INICIALIZAÇÃO DO GOOGLE SHEETS ===
def inicializar_google_sheets():
    try:
        credenciais = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(credenciais)
        planilha = client.open_by_key(PLANILHA_ID)
        mes_atual = datetime.now().strftime("%m-%Y")

        abas_existentes = [aba.title for aba in planilha.worksheets()]
        if mes_atual not in abas_existentes:
            print(f"➕ Criando aba '{mes_atual}'...")
            ws_modelo = planilha.sheet1
            ws_nova = planilha.add_worksheet(title=mes_atual, rows="100", cols="10")
            ws_nova.update([ws_modelo.get_all_values()[0]])
        else:
            print(f"✅ Usando aba existente '{mes_atual}'")
            ws_nova = planilha.worksheet(mes_atual)

        # Garantindo que a aba tenha ao menos 100 linhas
        if ws_nova.row_count < 100:
            ws_nova.resize(rows=100)

        return ws_nova
    except Exception as e:
        print(f"❌ Erro ao conectar ao Google Sheets: {e}")
        exit()

# === ROTA DE VERIFICAÇÃO (GET) ===
@app.route('/webhook', methods=['GET'])
def verify_webhook():
    token_enviado = request.args.get('hub.verify_token')
    challenge_enviado = request.args.get('hub.challenge')

    if token_enviado == os.getenv("WHATSAPP_VERIFY_TOKEN"):
        return challenge_enviado, 200
    return 'Invalid verification token', 403

# === ROTA PRINCIPAL (POST) ===
@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.get_json()
    print(f"_PAYLOAD RECEBIDO: {data}")

    if not data or 'entry' not in data or not data['entry']:
        print("⚠️ Dados inválidos recebidos.")
        return jsonify({'status': 'error'}), 400

    entry = data['entry'][0]
    changes = entry.get('changes', [])
    if not changes:
        return jsonify({'status': 'ok'}), 200

    change = changes[0]
    value = change.get('value', {})
    messages = value.get('messages', [])

    contacts = value.get('contacts', [])

    for message in messages:
        msg_body = message.get('text', {}).get('body', '')
        from_wa_id = message.get('from', 'Sem número')

        contato = next((contact for contact in contacts if contact['wa_id'] == from_wa_id), None)
        nome_cliente = contato['profile']['name'] if contato and 'profile' in contato else 'Desconhecido'

        print(f"📩 Nova mensagem recebida: {nome_cliente} ({from_wa_id})")
        print(f"💬 Mensagem: {msg_body}")

        itens = extrair_itens(msg_body)
        if not itens:
            print(f"⚠️ Nenhum item encontrado na mensagem de {nome_cliente}.")
            continue

        ws = inicializar_google_sheets()
        agora = datetime.now()
        data_hora = agora.strftime("%d/%m/%Y %H:%M")

        for item in itens:
            # Encontre a primeira linha realmente vazia, começando da linha 2
            linha = 2
            while linha <= ws.row_count and ws.acell(f"A{linha}").value:
                linha += 1

            # Se ultrapassou as linhas existentes, use a próxima disponível
            if linha > ws.row_count:
                linha = ws.row_count + 1

            # Inserir a nova linha
            ws.insert_row([
                "",                         # Nº
                nome_cliente,               # Cliente
                "",                         # Email
                item,                       # Itens do Pedido
                data_hora.split()[0],       # Data
                data_hora.split()[1],       # Hora
                "WhatsApp",                 # Origem
                "",                         # ML Preço
                "",                         # ML Link
                ""                          # ML Prazo
            ], linha)

            print(f"✅ Pedido adicionado: {nome_cliente} - {item}")

    return jsonify({'status': 'ok'}), 200

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)))
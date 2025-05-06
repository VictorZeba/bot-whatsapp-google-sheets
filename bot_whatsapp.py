from flask import Flask, request
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import time

app = Flask(__name__)

# === CONFIGURAÇÃO DO GOOGLE SHEETS ===
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDENTIALS_FILE = "credenciais.json"
PLANILHA_ID = "1ToDWKQjr4w2El8JKvoD5Wbe5BA9_zccgbC5bbyqY-so"  # Exemplo: "1abcXYZ123..."

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
            ws_nova.update([ws_modelo.get_all_values()[0]])  # Copia cabeçalho
        else:
            print(f"✅ Usando aba existente '{mes_atual}'")
        return planilha.worksheet(mes_atual)
    except Exception as e:
        print(f"❌ Erro ao conectar ao Google Sheets: {e}")
        exit()

# === ROTA PRINCIPAL ===
@app.route('/webhook', methods=['POST'])
def webhook():
    # Verifica se a requisição tem Content-Type correto
    if not request.is_json:
        print("❌ Requisição sem JSON.")
        return 'Unsupported Media Type', 415

    data = request.get_json()
    
    if not data or 'text' not in data or 'message' not in data['text']:
        print("⚠️ Dados inválidos recebidos.")
        return 'OK', 200

    mensagem = data['text']['message']
    nome_cliente = data.get('senderName', 'Desconhecido')
    numero_cliente = data.get('phone', 'Sem número')

    print(f"📩 Nova mensagem recebida: {nome_cliente} ({numero_cliente})")
    print(f"💬 Mensagem: {mensagem}")

    itens = extrair_itens(mensagem)
    if not itens:
        print(f"⚠️ Nenhum item encontrado na mensagem de {nome_cliente}.")
        return 'OK', 200

    ws = inicializar_google_sheets()
    agora = datetime.now()
    data_hora = agora.strftime("%d/%m/%Y %H:%M")

    for item in itens:
        linha = ws.row_count + 1
        ws.insert_row([
            "",                         # Nº
            nome_cliente,               # Cliente
            "",                         # Email (vazio pois vem do WhatsApp)
            item,                       # Itens do Pedido
            data_hora.split()[0],       # Data
            data_hora.split()[1],       # Hora
            "WhatsApp",                 # Origem
            "",                         # ML Preço
            "",                         # ML Link
            ""                          # ML Prazo
        ], linha)

        print(f"✅ Pedido adicionado: {nome_cliente} - {item}")

    return 'OK', 200

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
"""
Script para EXCLUIR um webhook especifico do Agendor.

COMO USAR:
1. Preencha o TOKEN_AGENDOR abaixo
2. Confirme o WEBHOOK_ID (já preenchido com o ID do webhook antigo
   que aponta para o Make, encontrado via listar_webhooks.py)
3. Rode: python excluir_webhook.py
"""

import requests

# ===================== PREENCHA AQUI =====================

TOKEN_AGENDOR = "048682f9-b111-4dad-8b3b-8efe2e01c86f"

WEBHOOK_ID = 13455  # ID do webhook antigo, apontando para o Make

# ===========================================================

url_api = f"https://api.agendor.com.br/integrations/subscriptions/{WEBHOOK_ID}"

headers = {
    "Content-Type": "application/json",
    "Authorization": f"Token {TOKEN_AGENDOR}",
}

print(f"Excluindo webhook ID {WEBHOOK_ID}...")
response = requests.delete(url_api, headers=headers)

print(f"Status da resposta: {response.status_code}")
print("-" * 60)

if response.status_code in (200, 204):
    print("Webhook excluido com sucesso!")
else:
    print("Erro ao excluir webhook:")
    print(response.text)

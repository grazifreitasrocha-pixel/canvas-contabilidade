"""
Sistema Trello - Atualizador de Alvará
Canvas Contabilidade

Cole o texto com os campos e o card é criado ou atualizado automaticamente.

Formato aceito:
  Cliente: Nome do Cliente
  Serviço: Obtenção de Alvará
  Status: Em Andamento
  Pendência: Aguardando CNPJ
  Próximo passo: Enviar documentos
  Responsável: Grazielle

Uso:
  python trello_alvara.py          → cola o texto no terminal
  python trello_alvara.py texto.txt → lê de um arquivo .txt
"""

import re
import sys
import os
import requests
from datetime import datetime

# ─── CONFIGURAÇÃO ───────────────────────────────────────────────────────────────
TRELLO_API_KEY  = os.getenv("TRELLO_API_KEY",  "53ad51fb101155fdb3d5afc323e75d3b")
TRELLO_TOKEN    = os.getenv("TRELLO_TOKEN",    "c4638f138a5591657f058d3776a68c7d0021c29967b44ed4509c3e51debbf652")
TRELLO_BOARD_ID = os.getenv("TRELLO_BOARD_ID", "OgwLrimi")

BASE_URL = "https://api.trello.com/1"

# ─── MAPEAMENTO DE CAMPOS ───────────────────────────────────────────────────────
# Aceita variações de escrita (com/sem acento, maiúscula/minúscula)
CAMPOS_REGEX = {
    "cliente":       r"cliente\s*:",
    "servico":       r"servi[çc]o\s*:",
    "status":        r"status\s*:",
    "pendencia":     r"pend[êe]ncia\s*:",
    "proximo_passo": r"pr[óo]ximo\s+passo\s*:",
    "responsavel":   r"respons[áa]vel\s*:",
}

# ─── PARSE DO TEXTO ─────────────────────────────────────────────────────────────

def parse_texto(texto: str) -> dict:
    """
    Extrai os campos do texto colado.
    Aceita qualquer ordem, com ou sem acento, maiúscula ou minúscula.
    """
    dados = {}
    linhas = texto.strip().splitlines()

    for i, linha in enumerate(linhas):
        linha_strip = linha.strip()
        for chave, padrao in CAMPOS_REGEX.items():
            match = re.match(padrao, linha_strip, re.IGNORECASE)
            if match:
                # Valor pode estar na mesma linha ou nas linhas seguintes
                valor_inline = linha_strip[match.end():].strip()
                if valor_inline:
                    # Verifica se há continuação nas próximas linhas
                    # (linhas que não começam com outro campo)
                    j = i + 1
                    while j < len(linhas):
                        prox = linhas[j].strip()
                        eh_novo_campo = any(
                            re.match(p, prox, re.IGNORECASE)
                            for p in CAMPOS_REGEX.values()
                        )
                        if eh_novo_campo or not prox:
                            break
                        valor_inline += " " + prox
                        j += 1
                    dados[chave] = valor_inline.strip()
                break

    return dados


# ─── TRELLO HELPERS ─────────────────────────────────────────────────────────────

def _p(**extra):
    return {"key": TRELLO_API_KEY, "token": TRELLO_TOKEN, **extra}


def formatar_descricao(dados: dict) -> str:
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    return (
        f"**Cliente:** {dados.get('cliente', '—')}\n"
        f"**Serviço:** {dados.get('servico', '—')}\n"
        f"**Status:** {dados.get('status', '—')}\n\n"
        f"**Pendência:** {dados.get('pendencia', '—')}\n\n"
        f"**Próximo passo:** {dados.get('proximo_passo', '—')}\n\n"
        f"**Responsável:** {dados.get('responsavel', '—')}\n\n"
        f"---\n_Atualizado em: {agora}_"
    )


def formatar_nome_card(dados: dict) -> str:
    cliente = dados.get("cliente", "Cliente")
    servico = dados.get("servico", "Alvará")
    return f"{cliente} — {servico}"


def obter_ou_criar_lista(nome: str) -> str:
    r = requests.get(f"{BASE_URL}/boards/{TRELLO_BOARD_ID}/lists", params=_p())
    r.raise_for_status()
    for l in r.json():
        if l["name"].strip().lower() == nome.strip().lower():
            return l["id"]
    r2 = requests.post(f"{BASE_URL}/lists", params=_p(name=nome, idBoard=TRELLO_BOARD_ID, pos="bottom"))
    r2.raise_for_status()
    lista_id = r2.json()["id"]
    print(f"  📁 Lista criada: '{nome}'")
    return lista_id


def buscar_card_existente(cliente: str, servico: str = None) -> dict | None:
    """
    Busca card existente pelo nome do cliente (e serviço, se informado).
    Retorna o card mais recente encontrado ou None.
    """
    r = requests.get(f"{BASE_URL}/boards/{TRELLO_BOARD_ID}/cards", params=_p())
    r.raise_for_status()
    cards = r.json()

    for card in cards:
        nome = card["name"].lower()
        if cliente.lower() in nome:
            if servico is None or servico.lower() in nome:
                return card
    return None


def criar_card(dados: dict) -> dict:
    status   = dados.get("status", "Em Andamento")
    lista_id = obter_ou_criar_lista(status)

    r = requests.post(f"{BASE_URL}/cards", params=_p(
        name   = formatar_nome_card(dados),
        desc   = formatar_descricao(dados),
        idList = lista_id,
        pos    = "top",
    ))
    r.raise_for_status()
    card = r.json()
    print(f"  ✅ Card CRIADO: '{card['name']}'")
    print(f"     {card['url']}")
    return card


def atualizar_card(card_id: str, dados_novos: dict) -> dict:
    # Buscar dados atuais
    r = requests.get(f"{BASE_URL}/cards/{card_id}", params=_p())
    r.raise_for_status()
    card_atual = r.json()

    # Mesclar: mantém o que já estava, sobrescreve com o que veio no texto
    dados_base = extrair_dados_descricao(card_atual.get("desc", ""))
    dados_base.update(dados_novos)

    payload = dict(
        name = formatar_nome_card(dados_base),
        desc = formatar_descricao(dados_base),
    )

    if "status" in dados_novos:
        payload["idList"] = obter_ou_criar_lista(dados_novos["status"])

    r2 = requests.put(f"{BASE_URL}/cards/{card_id}", params=_p(**payload))
    r2.raise_for_status()
    card = r2.json()
    print(f"  🔄 Card ATUALIZADO: '{card['name']}'")
    print(f"     Status: {dados_base.get('status', '—')}")
    print(f"     {card['url']}")
    return card


def extrair_dados_descricao(desc: str) -> dict:
    mapa = {
        "cliente":       "**Cliente:**",
        "servico":       "**Serviço:**",
        "status":        "**Status:**",
        "pendencia":     "**Pendência:**",
        "proximo_passo": "**Próximo passo:**",
        "responsavel":   "**Responsável:**",
    }
    dados = {}
    for chave, prefixo in mapa.items():
        for linha in desc.splitlines():
            linha = linha.strip()
            if linha.startswith(prefixo):
                valor = linha[len(prefixo):].strip()
                if valor and valor != "—":
                    dados[chave] = valor
                break
    return dados


def adicionar_comentario(card_id: str, dados: dict):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    partes = []
    if dados.get("status"):
        partes.append(f"Status: {dados['status']}")
    if dados.get("pendencia"):
        partes.append(f"Pendência: {dados['pendencia']}")
    if dados.get("proximo_passo"):
        partes.append(f"Próximo passo: {dados['proximo_passo']}")
    if partes:
        texto = f"[{agora}] Atualização — " + " | ".join(partes)
        requests.post(
            f"{BASE_URL}/cards/{card_id}/actions/comments",
            params=_p(text=texto),
        )


# ─── FUNÇÃO PRINCIPAL ───────────────────────────────────────────────────────────

def processar_texto(texto: str):
    """
    Recebe o texto colado, faz o parse e cria ou atualiza o card no Trello.
    """
    print("\n─── Processando texto ───────────────────────────────")
    dados = parse_texto(texto)

    if not dados:
        print("  ❌ Nenhum campo reconhecido no texto.")
        print("     Use o formato: Campo: Valor")
        return

    # Mostrar o que foi identificado
    print("  Campos identificados:")
    for k, v in dados.items():
        print(f"    {k}: {v}")

    if not dados.get("cliente"):
        print("\n  ❌ Campo 'Cliente' é obrigatório para identificar o card.")
        return

    # Buscar card existente
    card = buscar_card_existente(dados["cliente"], dados.get("servico"))

    print()
    if card:
        atualizar_card(card["id"], dados)
        adicionar_comentario(card["id"], dados)
    else:
        novo = criar_card(dados)
        adicionar_comentario(novo["id"], dados)

    print("─────────────────────────────────────────────────────\n")


# ─── ENTRADA ────────────────────────────────────────────────────────────────────

def checar_credenciais():
    if TRELLO_API_KEY == "SUA_API_KEY_AQUI":
        print("\n⚠️  Configure suas credenciais do Trello!")
        print("   Edite as variáveis no início do arquivo ou use variáveis de ambiente:")
        print("     set TRELLO_API_KEY=sua_chave")
        print("     set TRELLO_TOKEN=seu_token")
        print("     set TRELLO_BOARD_ID=id_do_board")
        print("\n   Obtenha em: https://trello.com/app-key\n")
        sys.exit(1)


if __name__ == "__main__":
    checar_credenciais()

    # Modo arquivo: python trello_alvara.py meu_texto.txt
    if len(sys.argv) > 1:
        caminho = sys.argv[1]
        with open(caminho, encoding="utf-8") as f:
            texto = f.read()
        processar_texto(texto)

    # Modo terminal: cola o texto diretamente
    else:
        print("═" * 55)
        print("  CANVAS CONTABILIDADE — Atualizador de Alvará")
        print("═" * 55)
        print("  Cole o texto com os campos e pressione Enter duas")
        print("  vezes (linha em branco) para enviar ao Trello.")
        print("  Ctrl+C para sair.\n")
        print("  Formato:")
        print("    Cliente: Nome")
        print("    Serviço: Obtenção de Alvará")
        print("    Status: Em Andamento")
        print("    Pendência: ...")
        print("    Próximo passo: ...")
        print("    Responsável: ...")
        print("─" * 55)

        while True:
            linhas = []
            try:
                while True:
                    linha = input()
                    if linha == "" and linhas and linhas[-1] == "":
                        break
                    linhas.append(linha)
            except KeyboardInterrupt:
                print("\n\n  Até logo! 👋\n")
                break

            texto = "\n".join(linhas).strip()
            if texto:
                processar_texto(texto)
                print("  Pronto para o próximo. Cole outro texto:\n")

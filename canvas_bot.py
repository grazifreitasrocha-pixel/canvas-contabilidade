import os, re, shutil, urllib.parse, requests, socket, logging
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from flask import Flask, request, jsonify
import threading
# Log em arquivo para diagnóstico
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(Path(__file__).parent / "canvas_bot.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# Força IPv4 para evitar falhas de conexão via IPv6
_orig_getaddrinfo = socket.getaddrinfo
def _getaddrinfo_ipv4(host, port, _family=0, *args, **kwargs):
    return _orig_getaddrinfo(host, port, socket.AF_INET, *args, **kwargs)
socket.getaddrinfo = _getaddrinfo_ipv4
from docx import Document
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (Application, CommandHandler, ConversationHandler,
                           MessageHandler, filters, ContextTypes)

load_dotenv()

BOT_TOKEN  = os.environ["BOT_TOKEN"]
GROUP_ID   = int(os.environ["TELEGRAM_GROUP_ID"])
TRELLO_KEY = os.environ["TRELLO_KEY"]
TRELLO_TOK = os.environ["TRELLO_API_TOKEN"]
ASAAS_KEY  = os.environ.get("ASAAS_KEY", "")

# Card modelo com checklist de integração
TEMPLATE_CARD_INTEGRACAO = "KbA4dhyW",

# IDs das listas de entrada em cada quadro
ENTRY_LISTS = {
    "1": "694197174aa205b99b63d0be",  # Integração → 1. Venda Ganha
    "2": "69a2f7279d5270c9fdc3b434",  # Alvarás    → 🔵 Novas Solicitações
    "3": "69a594121db830f4646b9518",  # Serviços   → 1. ENTRADA DE SOLICITAÇÃO
    "7": "694197174aa205b99b63d0be",  # MEI        → Integração (mesma lista)
}
TIPOS = {
    "1": "Integração do Cliente",
    "2": "Alvarás e Licenças",
    "3": "Abertura de empresa",
    "4": "Certificado Digital",
    "5": "Serviços da Folha",
    "6": "Serviços do Fiscal",
    "7": "MEI",
}

# Tipos que NÃO têm Trello
SEM_TRELLO = {"4", "5", "6"}

CONTRACT_TMPL     = Path(r"C:\Users\grazi.RAPHAEL\Downloads\Nova pasta (2)\Contrato Assessoria.docx.docx")
CONTRACT_TMPL_MEI = Path(r"C:\Users\grazi.RAPHAEL\Downloads\Contrato MEI\Contrato_MEI_Canvas_Revisado.docx")
CONTRATOS_DIR = Path(__file__).parent / "contratos"
ASAAS_URL     = "https://api.asaas.com/v3/payments"

# ── Estados ────────────────────────────────────────────────────────────────────
(TIPO, EMPRESA, CNPJ, REGIME, DATA_INICIO, SERVICOS, FOLHA,
 HONORARIO, VENCIMENTO, RESPONSAVEL, OBS, REPRESENTANTE,
 CPF_REP, END_EMPRESA, END_REP, WHATSAPP, EMAIL) = range(17)

# Estados do fluxo /boleto
(BOL_EMPRESA, BOL_CNPJ, BOL_DESC, BOL_VALOR, BOL_VENC, BOL_WHATSAPP) = range(17, 23)


# ── Trello ─────────────────────────────────────────────────────────────────────

def trello_copy_checklists(card_id: str, template_card_id: str):
    r = requests.get(f"https://api.trello.com/1/cards/{template_card_id}/checklists",
                     params={"key": TRELLO_KEY, "token": TRELLO_TOK})
    r.raise_for_status()
    for checklist in r.json():
        requests.post("https://api.trello.com/1/checklists",
                      params={"key": TRELLO_KEY, "token": TRELLO_TOK},
                      json={"idCard": card_id, "idChecklistSource": checklist["id"]}).raise_for_status()

def trello_add_card(board_key: str, name: str, desc: str) -> str:
    list_id = ENTRY_LISTS[board_key]
    r = requests.post("https://api.trello.com/1/cards",
                      params={"key": TRELLO_KEY, "token": TRELLO_TOK},
                      json={"name": name, "desc": desc, "idList": list_id, "pos": "top"})
    r.raise_for_status()
    card = r.json()
    if board_key in ("1", "7"):
        trello_copy_checklists(card["id"], TEMPLATE_CARD_INTEGRACAO)
    return card["url"]


# ── Asaas ──────────────────────────────────────────────────────────────────────

def _asaas(method: str, path: str, **kwargs):
    headers = {"access_token": ASAAS_KEY, "Content-Type": "application/json"}
    r = requests.request(method, f"{ASAAS_URL}{path}", headers=headers, **kwargs)
    if not r.ok:
        try:
            erros = r.json().get("errors", [])
            detail = "; ".join(e.get("description", str(e)) for e in erros) if erros else r.text
        except Exception:
            detail = r.text
        raise Exception(f"Asaas {r.status_code}: {detail}")
    return r.json()

def asaas_customer(name: str, cpf_cnpj: str) -> str:
    clean = re.sub(r"\D", "", cpf_cnpj)
    found = _asaas("GET", "/customers", params={"cpfCnpj": clean}).get("data", [])
    if found:
        return found[0]["id"]
    return _asaas("POST", "/customers", json={"name": name, "cpfCnpj": clean})["id"]

def asaas_boleto(customer_id: str, value: float, due_day: int, desc: str) -> dict:
    import time
    now = datetime.now()
    m   = now.month + (1 if now.day >= due_day else 0)
    y   = now.year + (1 if m > 12 else 0)
    m   = m % 12 or 12
    due = datetime(y, m, due_day).strftime("%Y-%m-%d")
    d   = _asaas("POST", "/payments", json={
        "customer": customer_id, "billingType": "BOLETO",
        "value": value, "dueDate": due, "description": desc,
    })
    payment_id = d["id"]
    url = d.get("bankSlipUrl") or d.get("invoiceUrl", "")
    if not url:
        time.sleep(2)
        info = _asaas("GET", f"/payments/{payment_id}/bankSlipUrl")
        url  = info.get("bankSlipUrl", "")
    return {"id": payment_id, "url": url, "due": due}


# ── Contrato ───────────────────────────────────────────────────────────────────

MESES_PT = ["janeiro","fevereiro","março","abril","maio","junho",
            "julho","agosto","setembro","outubro","novembro","dezembro"]

def _apply_subst(doc, subst: dict, bracket_style: bool):
    def sub_para(para):
        full = "".join(r.text for r in para.runs)
        new  = full
        if bracket_style:
            for key, val in subst.items():
                new = new.replace(key, val)
        else:
            for key, val in subst.items():
                new = re.sub(r'\$\{"' + re.escape(key) + r'"\s*:\s*"[^"]*"\}', val, new)
        if new != full and para.runs:
            para.runs[0].text = new
            for run in para.runs[1:]:
                run.text = ""

    for para in doc.paragraphs:
        sub_para(para)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    sub_para(para)

def fill_contract_mei(d: dict, today, safe: str) -> Path:
    doc   = Document(CONTRACT_TMPL_MEI)
    subst = {
        "[Nome completo do Cliente]":              d["representante"],
        "[CPF ou CNPJ do Contratante]":            d["cpf_rep"],
        "[CNPJ do MEI]":                           d["cnpj"],
        "[Endereço completo com CEP]":             d["end_empresa"],
        "[E-mail do Contratante]":                 d.get("email", "-"),
        "[Telefone]":                              d["whatsapp"],
        "[Valor da mensalidade (ex: R$ 89,90)]":  f"R$ {d['honorario']}",
        "[valor por extenso]":                     d["honorario"],
        "[dia do vencimento (ex: 10)]":            d["vencimento"],
        "[data de início (dd/mm/aaaa)]":           d["data_inicio"],
        "[dia]":                                   today.strftime("%d"),
        "[mês]":                                   MESES_PT[today.month - 1],
        "[ano]":                                   today.strftime("%Y"),
        "[NOME COMPLETO DO CONTRATANTE]":          d["representante"].upper(),
    }
    _apply_subst(doc, subst, bracket_style=True)
    out = CONTRATOS_DIR / f"Contrato_MEI_{safe}_{today.strftime('%Y%m%d_%H%M')}.docx"
    doc.save(out)
    return out

def fill_contract_assessoria(d: dict, today, safe: str) -> Path:
    doc   = Document(CONTRACT_TMPL)
    subst = {
        "CONTRATANTE":                  d["empresa"],
        "CNPJ":                         d["cnpj"],
        "Endereço da empresa":          d["end_empresa"],
        "Representante legal":          d["representante"],
        "CPF":                          d["cpf_rep"],
        "Endereço contratante":         d["end_rep"],
        "Início contrato":              d["data_inicio"],
        "Valor do honorário":           f"R$ {d['honorario']}",
        "Data de vencimento":           f"dia {d['vencimento']}",
        "Data e local de assinatura":   f"Fortaleza, {today.strftime('%d/%m/%Y')}",
        "Contratante":                  d["empresa"],
    }
    _apply_subst(doc, subst, bracket_style=False)
    out = CONTRATOS_DIR / f"Contrato_{safe}_{today.strftime('%Y%m%d_%H%M')}.docx"
    doc.save(out)
    return out

def fill_contract(d: dict) -> Path:
    CONTRATOS_DIR.mkdir(exist_ok=True)
    today = datetime.now()
    safe  = re.sub(r"[^\w\s-]", "", d["empresa"])[:40].strip()
    if d["tipo"] == "7":
        return fill_contract_mei(d, today, safe)
    return fill_contract_assessoria(d, today, safe)


# ── Mensagem grupo ─────────────────────────────────────────────────────────────

def build_group_msg(d: dict, trello_url: str) -> str:
    label = TIPOS[d["tipo"]]
    return (
        f"🆕 NOVO CADASTRO\n"
        f"🔖 Tipo: {label}\n"
        f"{'─'*30}\n\n"
        f"🏢 Empresa: {d['empresa']}\n"
        f"📄 CNPJ/CPF: {d['cnpj']}\n"
        f"📱 WhatsApp: {d['whatsapp']}\n"
        f"💼 Regime tributário: {d['regime']}\n"
        f"📅 Início da responsabilidade: {d['data_inicio']}\n"
        f"📋 Serviços contratados: {d['servicos']}\n"
        f"👥 Folha de pagamento: {d['folha']}\n\n"
        f"💰 Honorário: R$ {d['honorario']} | Venc. {d['vencimento']}\n\n"
        f"📌 Status atual: Novo cadastro\n\n"
        f"👩‍💼 Responsável pela integração: {d['responsavel']}\n\n"
        f"📝 Obs: {d['obs']}\n\n"
        f"🔗 Ver cartão no Trello: {trello_url}\n\n"
        f"Canvas Contabilidade — Aqui a gente não para"
    )


# ── Handlers da conversa ───────────────────────────────────────────────────────

async def novo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text(
        "🆕 *NOVO CADASTRO*\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "Qual o tipo de serviço?\n\n"
        "1️⃣ — Integração do Cliente\n"
        "2️⃣ — Alvarás e Licenças\n"
        "3️⃣ — Abertura / Alteração / Baixa de Empresa\n"
        "4️⃣ — Certificado Digital\n"
        "5️⃣ — Serviços da Folha\n"
        "6️⃣ — Serviços do Fiscal\n"
        "7️⃣ — MEI\n\n"
        "Digite o número:",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(
            [["1"], ["2"], ["3"], ["4"], ["5"], ["6"], ["7"]],
            one_time_keyboard=True, resize_keyboard=True
        )
    )
    return TIPO

async def get_tipo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    t = update.message.text.strip()[0]
    if t not in TIPOS:
        await update.message.reply_text("Digite 1, 2 ou 3.")
        return TIPO
    ctx.user_data["tipo"] = t
    await update.message.reply_text("🏢 Nome da empresa:", reply_markup=ReplyKeyboardRemove())
    return EMPRESA

async def get_empresa(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["empresa"] = update.message.text.strip().upper()
    await update.message.reply_text("📄 CNPJ ou CPF:")
    return CNPJ

async def get_cnpj(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["cnpj"] = update.message.text.strip()
    await update.message.reply_text(
        "💼 Regime tributário:",
        reply_markup=ReplyKeyboardMarkup(
            [["MEI"], ["Simples Nacional"], ["Lucro Presumido"], ["Lucro Real"]],
            one_time_keyboard=True, resize_keyboard=True
        )
    )
    return REGIME

async def get_regime(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["regime"] = update.message.text.strip()
    await update.message.reply_text("📅 Início da responsabilidade (dd/mm/aaaa):", reply_markup=ReplyKeyboardRemove())
    return DATA_INICIO

async def get_data(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["data_inicio"] = update.message.text.strip()
    await update.message.reply_text("📋 Serviços contratados:")
    return SERVICOS

async def get_servicos(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["servicos"] = update.message.text.strip()
    await update.message.reply_text(
        "👥 Folha de pagamento?",
        reply_markup=ReplyKeyboardMarkup([["Sim"], ["Não"]], one_time_keyboard=True, resize_keyboard=True)
    )
    return FOLHA

async def get_folha(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["folha"] = update.message.text.strip()
    await update.message.reply_text("💰 Honorário mensal (ex: 200,00):", reply_markup=ReplyKeyboardRemove())
    return HONORARIO

async def get_honorario(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["honorario"] = update.message.text.strip()
    await update.message.reply_text("📅 Dia de vencimento (ex: 15):")
    return VENCIMENTO

async def get_vencimento(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["vencimento"] = update.message.text.strip()
    await update.message.reply_text("👩‍💼 Responsável pela integração:")
    return RESPONSAVEL

async def get_responsavel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["responsavel"] = update.message.text.strip()
    await update.message.reply_text("📝 Observações (ou 'nenhuma'):")
    return OBS

async def get_obs(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["obs"] = update.message.text.strip()
    await update.message.reply_text("👤 Representante legal (nome completo):")
    return REPRESENTANTE

async def get_representante(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["representante"] = update.message.text.strip()
    await update.message.reply_text("🪪 CPF do representante:")
    return CPF_REP

async def get_cpf_rep(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["cpf_rep"] = update.message.text.strip()
    await update.message.reply_text("📍 Endereço da empresa (completo):")
    return END_EMPRESA

async def get_end_empresa(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["end_empresa"] = update.message.text.strip()
    await update.message.reply_text("🏠 Endereço do representante (completo):")
    return END_REP

async def get_end_rep(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["end_rep"] = update.message.text.strip()
    await update.message.reply_text("📱 WhatsApp do cliente (com DDD, ex: 85999887766):")
    return WHATSAPP

async def get_whatsapp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["whatsapp"] = re.sub(r"\D", "", update.message.text.strip())
    await update.message.reply_text("📧 E-mail do cliente:")
    return EMAIL

async def get_email(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["email"] = update.message.text.strip()
    d = ctx.user_data

    await update.message.reply_text("⏳ Processando...", reply_markup=ReplyKeyboardRemove())

    erros      = []
    trello_url = "—"
    boleto_url = "—"

    # 1 ── Trello (apenas para tipos com quadro)
    if d["tipo"] not in SEM_TRELLO:
        try:
            desc = (
                f"**Empresa:** {d['empresa']}\n"
                f"**CNPJ/CPF:** {d['cnpj']}\n"
                f"**Regime:** {d['regime']}\n"
                f"**Inicio:** {d['data_inicio']}\n"
                f"**Servicos:** {d['servicos']}\n"
                f"**Folha:** {d['folha']}\n"
                f"**Honorario:** R$ {d['honorario']} | Venc. {d['vencimento']}\n"
                f"**Responsavel:** {d['responsavel']}\n"
                f"**Obs:** {d['obs']}"
            )
            trello_url = trello_add_card(
                d["tipo"],
                f"{d['empresa']} — {TIPOS[d['tipo']]}",
                desc
            )
        except Exception as e:
            erros.append(f"Trello: {e}")
    else:
        trello_url = "sem Trello para este servico"

    # 2 ── Mensagem no grupo Telegram
    try:
        await ctx.bot.send_message(chat_id=GROUP_ID, text=build_group_msg(d, trello_url))
    except Exception as e:
        erros.append(f"Grupo Telegram (chat_id={GROUP_ID}): {e}")

    # 3 ── Contrato
    contract_path = None
    try:
        contract_path = fill_contract(d)
    except Exception as e:
        erros.append(f"Contrato: {e}")

    # 4 ── Boleto Asaas
    if ASAAS_KEY and ASAAS_KEY != "PREENCHER_COM_CHAVE_ASAAS":
        try:
            valor = float(d["honorario"].replace(".", "").replace(",", "."))
            cid   = asaas_customer(d["empresa"], d["cnpj"])
            bol   = asaas_boleto(cid, valor, int(d["vencimento"]),
                                 f"Honorario {TIPOS[d['tipo']]} — {d['empresa']}")
            boleto_url = bol["url"]
        except Exception as e:
            erros.append(f"Asaas: {e}")
    else:
        erros.append("Asaas: chave nao configurada — boleto nao gerado")

    # 5 ── Informa caminho do contrato (sem upload para evitar timeout)
    if contract_path and contract_path.exists():
        erros_info = f"\n\nContrato salvo em:\n{contract_path}"
    else:
        erros_info = ""

    # 6 ── Link WhatsApp
    if d["tipo"] == "1":
        wa_text = (
            f"🎉 Seja muito bem-vindo(a) à Canvas Contabilidade!\n\n"
            f"Eu sou a Bruna, da equipe de implantação, e estarei com você durante todo o processo inicial da sua empresa conosco. 😊\n\n"
            f"Nosso objetivo é tornar essa integração simples, rápida e tranquila, garantindo que você tenha segurança, clareza e um atendimento próximo desde o primeiro dia.\n\n"
            f"Sempre que surgir qualquer dúvida, pode contar comigo! 🤝\n\n"
            f"⏰ Nosso horário de atendimento:\n"
            f"Segunda a sexta-feira\n"
            f"07h30 às 11h00 e das 13h00 às 17h30\n\n"
            f"Para darmos andamento à integração da sua empresa na Canvas Contabilidade, precisamos dos seguintes documentos:\n\n"
            f"👥 Dos sócios:\n"
            f"• Documento pessoal (RG e CPF ou CNH)\n"
            f"• Comprovante de endereço atualizado\n"
            f"• Certidão de casamento (caso seja casado)\n\n"
            f"🏢 Da empresa:\n"
            f"• Contrato social e alterações (preferencialmente em PDF)\n"
            f"• Cartão do CNPJ\n"
            f"• Comprovante de endereço atualizado\n"
            f"• Contrato de locação do imóvel (caso seja alugado)\n\n"
            f"Assim que recebermos a documentação, iniciaremos as próximas etapas da implantação. ✨\n\n"
            f"Será um prazer fazer parte dessa nova fase da sua empresa!\n\n"
            f"Link de pagamento do boleto de honorário:\n{boleto_url}"
        )
    elif d["tipo"] == "3":
        wa_text = (
            f"Olá, {d['empresa'].title()}! Tudo bem?\n\n"
            f"É com muita alegria que compartilhamos com você o Contrato Social e o Cartão CNPJ da sua empresa. 🎉\n\n"
            f"A partir de agora, você inicia uma nova fase como empresário, e queremos que saiba que estamos aqui para caminhar ao seu lado em cada passo.\n\n"
            f"Nos próximos dias, o nosso time de implantação entrará em contato para orientar sobre os primeiros procedimentos, obrigações e também para realizar um treinamento, garantindo que você tenha clareza e segurança nessa jornada.\n\n"
            f"Link de pagamento do boleto de honorário:\n{boleto_url}\n\n"
            f"Um grande abraço,\n\n"
            f"Equipe Canvas Contabilidade"
        )
    else:
        wa_text = (
            f"Olá, {d['empresa'].title()}! 😊\n\n"
            f"Seja bem-vindo(a) à Canvas Contabilidade!\n\n"
            f"Serviço contratado: {TIPOS[d['tipo']]}\n\n"
            f"Segue o link do seu boleto:\n{boleto_url}\n\n"
            f"Vencimento: dia {d['vencimento']}\n"
            f"Valor: R$ {d['honorario']}\n\n"
            f"Qualquer dúvida, estamos à disposição! 🟢\n"
            f"Canvas Contabilidade"
        )
    wa_link = f"https://wa.me/55{d['whatsapp']}?text={urllib.parse.quote(wa_text)}"

    # Resumo final
    status = "✅ Tudo certo!" if not erros else "⚠️ Concluído com avisos:\n" + "\n".join(f"• {e}" for e in erros)

    await update.message.reply_text(
        f"{status}\n\n"
        f"Trello: {trello_url}\n"
        f"Boleto: {boleto_url}\n\n"
        f"WhatsApp:\n{wa_link}"
        f"{erros_info}",
        disable_web_page_preview=True
    )

    return ConversationHandler.END

# ── Fluxo /boleto ──────────────────────────────────────────────────────────────

async def boleto_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text(
        "💰 *Boleto de Honorário Avulso*\n\n🏢 Nome da empresa ou cliente:",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove(),
    )
    return BOL_EMPRESA

async def bol_empresa(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["bol_empresa"] = update.message.text.strip().upper()
    await update.message.reply_text("📄 CNPJ ou CPF:")
    return BOL_CNPJ

async def bol_cnpj(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["bol_cnpj"] = update.message.text.strip()
    await update.message.reply_text("📋 Descrição do serviço (ex: Honorário maio/2026):")
    return BOL_DESC

async def bol_desc(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["bol_desc"] = update.message.text.strip()
    await update.message.reply_text("💰 Valor (ex: 350,00):")
    return BOL_VALOR

async def bol_valor(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["bol_valor"] = update.message.text.strip()
    await update.message.reply_text("📅 Dia de vencimento (ex: 10):")
    return BOL_VENC

async def bol_venc(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["bol_venc"] = update.message.text.strip()
    await update.message.reply_text("📱 WhatsApp do cliente (com DDD, ex: 85999887766):")
    return BOL_WHATSAPP

async def bol_whatsapp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["bol_whatsapp"] = re.sub(r"\D", "", update.message.text.strip())
    d = ctx.user_data

    await update.message.reply_text("⏳ Gerando boleto...")

    try:
        valor = float(d["bol_valor"].replace(".", "").replace(",", "."))
        cid   = asaas_customer(d["bol_empresa"], d["bol_cnpj"])
        bol   = asaas_boleto(cid, valor, int(d["bol_venc"]), d["bol_desc"])

        wa_text = (
            f"Olá, {d['bol_empresa'].title()}! 😊\n\n"
            f"Segue o boleto referente a:\n{d['bol_desc']}\n\n"
            f"💰 Valor: R$ {d['bol_valor']}\n"
            f"📅 Vencimento: {bol['due'].split('-')[2]}/{bol['due'].split('-')[1]}/{bol['due'].split('-')[0]}\n\n"
            f"🔗 {bol['url']}\n\n"
            f"Qualquer dúvida, estamos à disposição! 🟢\n"
            f"Canvas Contabilidade"
        )
        wa_link = f"https://wa.me/55{d['bol_whatsapp']}?text={urllib.parse.quote(wa_text)}"

        await update.message.reply_text(
            f"✅ *Boleto gerado!*\n\n"
            f"🏢 {d['bol_empresa']}\n"
            f"📋 {d['bol_desc']}\n"
            f"💰 R$ {d['bol_valor']}\n"
            f"📅 Venc: {bol['due'].split('-')[2]}/{bol['due'].split('-')[1]}/{bol['due'].split('-')[0]}\n\n"
            f"🔗 Boleto: {bol['url']}\n\n"
            f"📲 Enviar pelo WhatsApp:\n{wa_link}",
            parse_mode="Markdown",
            disable_web_page_preview=True,
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Erro ao gerar boleto: {e}")

    ctx.user_data.clear()
    return ConversationHandler.END


async def cancelar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ Cancelado.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Canvas Contabilidade 🟢\n\n"
        "/novo — cadastrar novo cliente\n"
        "/boleto — gerar boleto de honorário avulso\n"
        "/cancelar — cancelar operação em andamento"
    )

async def chatid(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Chat ID: {update.effective_chat.id}")

async def error_handler(update: object, ctx: ContextTypes.DEFAULT_TYPE):
    logger.error("Erro no handler", exc_info=ctx.error)
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(
                "⚠️ Ocorreu um erro inesperado. Use /cancelar e tente novamente."
            )
        except Exception:
            pass

# ── Main ───────────────────────────────────────────────────────────────────────

def build_app():
    from telegram.ext import Application as _App
    app = _App.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("novo", novo)],
        name="novo",
        persistent=False,
        states={
            TIPO:         [MessageHandler(filters.TEXT & ~filters.COMMAND, get_tipo)],
            EMPRESA:      [MessageHandler(filters.TEXT & ~filters.COMMAND, get_empresa)],
            CNPJ:         [MessageHandler(filters.TEXT & ~filters.COMMAND, get_cnpj)],
            REGIME:       [MessageHandler(filters.TEXT & ~filters.COMMAND, get_regime)],
            DATA_INICIO:  [MessageHandler(filters.TEXT & ~filters.COMMAND, get_data)],
            SERVICOS:     [MessageHandler(filters.TEXT & ~filters.COMMAND, get_servicos)],
            FOLHA:        [MessageHandler(filters.TEXT & ~filters.COMMAND, get_folha)],
            HONORARIO:    [MessageHandler(filters.TEXT & ~filters.COMMAND, get_honorario)],
            VENCIMENTO:   [MessageHandler(filters.TEXT & ~filters.COMMAND, get_vencimento)],
            RESPONSAVEL:  [MessageHandler(filters.TEXT & ~filters.COMMAND, get_responsavel)],
            OBS:          [MessageHandler(filters.TEXT & ~filters.COMMAND, get_obs)],
            REPRESENTANTE:[MessageHandler(filters.TEXT & ~filters.COMMAND, get_representante)],
            CPF_REP:      [MessageHandler(filters.TEXT & ~filters.COMMAND, get_cpf_rep)],
            END_EMPRESA:  [MessageHandler(filters.TEXT & ~filters.COMMAND, get_end_empresa)],
            END_REP:      [MessageHandler(filters.TEXT & ~filters.COMMAND, get_end_rep)],
            WHATSAPP:     [MessageHandler(filters.TEXT & ~filters.COMMAND, get_whatsapp)],
            EMAIL:        [MessageHandler(filters.TEXT & ~filters.COMMAND, get_email)],
        },
        fallbacks=[CommandHandler("cancelar", cancelar)],
    )

    conv_boleto = ConversationHandler(
        entry_points=[CommandHandler("boleto", boleto_start)],
        name="boleto",
        persistent=False,
        states={
            BOL_EMPRESA:  [MessageHandler(filters.TEXT & ~filters.COMMAND, bol_empresa)],
            BOL_CNPJ:     [MessageHandler(filters.TEXT & ~filters.COMMAND, bol_cnpj)],
            BOL_DESC:     [MessageHandler(filters.TEXT & ~filters.COMMAND, bol_desc)],
            BOL_VALOR:    [MessageHandler(filters.TEXT & ~filters.COMMAND, bol_valor)],
            BOL_VENC:     [MessageHandler(filters.TEXT & ~filters.COMMAND, bol_venc)],
            BOL_WHATSAPP: [MessageHandler(filters.TEXT & ~filters.COMMAND, bol_whatsapp)],
        },
        fallbacks=[CommandHandler("cancelar", cancelar)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("chatid", chatid))
    app.add_handler(conv)
    app.add_handler(conv_boleto)
    app.add_error_handler(error_handler)
    return app

flask_app = Flask(__name__)
_telegram_app_ref = {"app": None}
_main_event_loop = {"loop": None}


def montar_mensagem_agendor(dados_negocio: dict) -> str:
    org = dados_negocio.get("organization", {}) or {}
    nome_empresa = org.get("name", "Não informado")
    cnpj = org.get("cnpj") or "Não informado"
    valor = dados_negocio.get("value", "Não informado")
    whatsapp = (org.get("contact") or {}).get("whatsapp", "Não informado")
    titulo = dados_negocio.get("title", "")

    return (
        f"🟢 *Novo negócio ganho no Agendor!*\n\n"
        f"*Negócio:* {titulo}\n"
        f"*Empresa:* {nome_empresa}\n"
        f"*CNPJ:* {cnpj}\n"
        f"*Valor:* R$ {valor}\n"
        f"*WhatsApp:* {whatsapp}\n\n"
        f"Use /novo no bot para iniciar o cadastro completo "
        f"(os dados acima já estão prontos para copiar)."
    )


@flask_app.route("/webhook/agendor", methods=["POST"])
def webhook_agendor():
    try:
        payload = request.get_json(force=True)
    except Exception as e:
        return jsonify({"erro": f"JSON invalido: {e}"}), 400

    try:
        dados_lista = payload.get("dados") or payload.get("data") or []
        if isinstance(dados_lista, list) and len(dados_lista) > 0:
            negocio_id = dados_lista[0].get("id")
        elif isinstance(dados_lista, dict):
            negocio_id = dados_lista.get("id")
        else:
            negocio_id = None
    except Exception:
        negocio_id = None

    if not negocio_id:
        return jsonify({"erro": "Nao foi possivel extrair o ID do negocio"}), 400

    import requests as _requests
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Token {os.environ.get('AGENDOR_TOKEN')}",
    }
    resp = _requests.get(
        f"https://api.agendor.com.br/v3/deals/{negocio_id}",
        headers=headers,
    )

    if resp.status_code != 200:
        return jsonify({"erro": f"Agendor retornou {resp.status_code}"}), 502

    dados_negocio = resp.json()

    chat_id = os.environ.get("GRAZIELE_CHAT_ID")
    if not chat_id:
        return jsonify({"erro": "GRAZIELE_CHAT_ID nao configurado"}), 500

    mensagem = montar_mensagem_agendor(dados_negocio)

    app_ref = _telegram_app_ref["app"]
    loop = _main_event_loop["loop"]

    if app_ref and loop:
        asyncio.run_coroutine_threadsafe(
            app_ref.bot.send_message(
                chat_id=int(chat_id), text=mensagem, parse_mode="Markdown"
            ),
            loop,
        )

    return jsonify({"status": "ok"}), 200


def iniciar_flask():
    porta = int(os.environ.get("PORT", 8080))
    flask_app.run(host="0.0.0.0", port=porta, use_reloader=False)

def main():
    try:
        while True:
            try:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                app = build_app()

                _telegram_app_ref["app"] = app
                _main_event_loop["loop"] = loop

                flask_thread = threading.Thread(target=iniciar_flask, daemon=True)
                flask_thread.start()

                print("Canvas Bot rodando -- /novo para cadastrar")
                print("Webhook do Agendor disponivel em /webhook/agendor")
                app.run_polling(drop_pending_updates=True)
                break
            except Conflict:
                print("Conflito: outro bot ativo. Aguardando 20s e tentando novamente...")
                time.sleep(20)
            except Exception as e:
                print(f"Erro inesperado: {e}. Reiniciando em 10s...")
                time.sleep(10)
    finally:
        lock_socket.close()

if __name__ == "__main__":
    main()

import os
import re
import json
from datetime import datetime
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

load_dotenv()

BOT_TOKEN = os.environ["FINA_BOT_TOKEN"]
SHEET_ID  = os.environ["SHEET_ID"]
USER_ID   = os.environ.get("TELEGRAM_USER_ID", "")

SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

def get_sheet(name: str):
    creds_json = json.loads(os.environ["GOOGLE_CREDS_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    gc = gspread.authorize(creds)
    wb = gc.open_by_key(SHEET_ID)
    try:
        return wb.worksheet(name)
    except Exception:
        ws = wb.add_worksheet(name, rows=1000, cols=10)
        return ws

def ensure_headers(ws, headers):
    if ws.row_values(1) != headers:
        ws.insert_row(headers, 1)

def extract_amount(text: str):
    m = re.search(r"(\d{1,6}[.,]\d{2}|\d{1,6})", text)
    return float(m.group().replace(",", ".")) if m else None

def extract_day(text: str):
    m = re.search(r"dia\s+(\d{1,2})", text)
    return int(m.group(1)) if m else None

def extract_ate(text: str):
    m = re.search(r"at[eé]\s+(\d{2})/(\d{4})", text)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None

def parcelas_restantes(mes_fim: int, ano_fim: int) -> int:
    now = datetime.now()
    fim = datetime(ano_fim, mes_fim, 1)
    ini = datetime(now.year, now.month, 1)
    delta = relativedelta(fim, ini)
    return max(0, delta.months + delta.years * 12 + 1)

def remove_amount(text: str):
    return re.sub(r"\d{1,6}[.,]\d{2}|\d{1,6}", "", text).strip()

def is_authorized(update: Update) -> bool:
    return not USER_ID or str(update.effective_user.id) == USER_ID

# ── Comandos ──────────────────────────────────────────────

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Olá! Sou a *FINA*, sua assistente financeira.\n\n"
        "*Como me usar:*\n\n"
        "💳 `paguei C6 209,76`\n"
        "💸 `gastei 80 padaria`\n"
        "💚 `recebi 5000 salário`\n"
        "📌 `nova conta escola 650 dia 5`\n"
        "🔄 `parcela colchão 350 dia 5 até 09/2026`\n\n"
        "📊 /saldo — saldo do mês\n"
        "📋 /contas — contas fixas e parceladas\n"
        "📈 /resumo — resumo completo do mês",
        parse_mode="Markdown"
    )

async def cmd_saldo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return

    month = datetime.now().strftime("%m/%Y")

    try:
        ws_r = get_sheet("Receitas")
        ws_d = get_sheet("Despesas")
        ws_p = get_sheet("Pagamentos")

        total_r = sum(float(r["Valor"]) for r in ws_r.get_all_records() if r.get("Mês") == month)
        total_d = sum(float(r["Valor"]) for r in ws_d.get_all_records() if r.get("Mês") == month)
        total_p = sum(float(r["Valor"]) for r in ws_p.get_all_records() if r.get("Mês") == month)

        total_saidas = total_d + total_p
        saldo = total_r - total_saidas
        emoji = "✅" if saldo >= 0 else "🔴"

        await update.message.reply_text(
            f"📊 *Saldo — {month}*\n\n"
            f"💚 Entradas:     R$ {total_r:,.2f}\n"
            f"💸 Gastos:       R$ {total_d:,.2f}\n"
            f"✅ Pagamentos:   R$ {total_p:,.2f}\n"
            f"──────────────────────\n"
            f"{emoji} *Saldo: R$ {saldo:,.2f}*",
            parse_mode="Markdown"
        )
    except Exception as e:
        await update.message.reply_text(f"Erro ao buscar saldo: {e}")

async def cmd_contas(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return

    try:
        ws_fixas = get_sheet("Contas Fixas")
        ws_parc  = get_sheet("Parcelados")
        fixas    = ws_fixas.get_all_records()
        parc     = ws_parc.get_all_records()

        now = datetime.now()
        parc_ativos = []
        for r in parc:
            try:
                mes_f, ano_f = int(r["Até"][:2]), int(r["Até"][3:])
                if datetime(ano_f, mes_f, 1) >= datetime(now.year, now.month, 1):
                    parc_ativos.append(r)
            except Exception:
                pass

        if not fixas and not parc_ativos:
            await update.message.reply_text(
                "Nenhuma conta cadastrada ainda.\n\n"
                "Use: `nova conta escola 650 dia 5`\n"
                "ou: `parcela colchão 350 dia 5 até 09/2026`",
                parse_mode="Markdown"
            )
            return

        msg   = "📋 *Contas do Mês*\n\n"
        total = 0

        if fixas:
            msg += "🔒 *Fixas:*\n"
            for r in fixas:
                valor = float(r["Valor"])
                msg  += f"• {r['Descrição']} — R$ {valor:,.2f} ({r['Vencimento']})\n"
                total += valor

        if parc_ativos:
            msg += "\n🔄 *Parceladas:*\n"
            for r in parc_ativos:
                valor     = float(r["Valor"])
                mes_f, ano_f = int(r["Até"][:2]), int(r["Até"][3:])
                restantes = parcelas_restantes(mes_f, ano_f)
                msg  += f"• {r['Descrição']} — R$ {valor:,.2f} ({r['Vencimento']}) · {restantes}x até {r['Até']}\n"
                total += valor

        msg += f"\n💰 *Total mensal: R$ {total:,.2f}*"
        await update.message.reply_text(msg, parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"Erro: {e}")

async def cmd_resumo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return

    month = datetime.now().strftime("%m/%Y")

    try:
        ws_r = get_sheet("Receitas")
        ws_d = get_sheet("Despesas")
        ws_p = get_sheet("Pagamentos")

        receitas   = [r for r in ws_r.get_all_records() if r.get("Mês") == month]
        despesas   = [r for r in ws_d.get_all_records() if r.get("Mês") == month]
        pagamentos = [r for r in ws_p.get_all_records() if r.get("Mês") == month]

        total_r = sum(float(r["Valor"]) for r in receitas)
        total_d = sum(float(r["Valor"]) for r in despesas)
        total_p = sum(float(r["Valor"]) for r in pagamentos)
        saldo   = total_r - total_d - total_p

        msg = f"📈 *Resumo — {month}*\n\n"

        if receitas:
            msg += "💚 *Entradas:*\n"
            for r in receitas:
                msg += f"  {r['Data']} · {r['Descrição']} — R$ {float(r['Valor']):,.2f}\n"
            msg += f"  *Total: R$ {total_r:,.2f}*\n\n"

        if pagamentos:
            msg += "✅ *Pagamentos:*\n"
            for r in pagamentos:
                msg += f"  {r['Data']} · {r['Descrição']} — R$ {float(r['Valor']):,.2f}\n"
            msg += f"  *Total: R$ {total_p:,.2f}*\n\n"

        if despesas:
            msg += "💸 *Gastos:*\n"
            for r in despesas:
                msg += f"  {r['Data']} · {r['Descrição']} — R$ {float(r['Valor']):,.2f}\n"
            msg += f"  *Total: R$ {total_d:,.2f}*\n\n"

        emoji = "✅" if saldo >= 0 else "🔴"
        msg += f"──────────────────────\n{emoji} *Saldo: R$ {saldo:,.2f}*"

        await update.message.reply_text(msg, parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"Erro: {e}")

# ── Mensagens livres ──────────────────────────────────────

async def handle_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update):
        return

    text  = update.message.text.strip()
    lower = text.lower()
    now   = datetime.now()
    date_str = now.strftime("%d/%m/%Y")
    month    = now.strftime("%m/%Y")

    # PAGUEI
    if lower.startswith("paguei") or lower.startswith("pag "):
        clean  = re.sub(r"^(paguei|pag)\s*", "", lower).strip()
        amount = extract_amount(clean)
        desc   = remove_amount(clean).title() or "Pagamento"

        if not amount:
            await update.message.reply_text("Não entendi o valor.\nEx: `paguei C6 209,76`", parse_mode="Markdown")
            return

        ws = get_sheet("Pagamentos")
        ensure_headers(ws, ["Data", "Mês", "Descrição", "Valor"])
        ws.append_row([date_str, month, desc, amount])
        await update.message.reply_text(
            f"✅ *Pagamento registrado!*\n📋 {desc}\n💰 R$ {amount:,.2f}\n📅 {date_str}",
            parse_mode="Markdown"
        )

    # GASTEI
    elif lower.startswith("gastei") or lower.startswith("gasto "):
        clean  = re.sub(r"^(gastei|gasto)\s*", "", lower).strip()
        amount = extract_amount(clean)
        desc   = remove_amount(clean).title() or "Gasto"

        if not amount:
            await update.message.reply_text("Não entendi o valor.\nEx: `gastei 50 mercado`", parse_mode="Markdown")
            return

        ws = get_sheet("Despesas")
        ensure_headers(ws, ["Data", "Mês", "Descrição", "Valor"])
        ws.append_row([date_str, month, desc, amount])
        await update.message.reply_text(
            f"💸 *Gasto registrado!*\n📋 {desc}\n💰 R$ {amount:,.2f}\n📅 {date_str}",
            parse_mode="Markdown"
        )

    # RECEBI
    elif lower.startswith("recebi") or lower.startswith("entrada "):
        clean  = re.sub(r"^(recebi|entrada)\s*", "", lower).strip()
        amount = extract_amount(clean)
        desc   = remove_amount(clean).title() or "Entrada"

        if not amount:
            await update.message.reply_text("Não entendi o valor.\nEx: `recebi 5000 salário`", parse_mode="Markdown")
            return

        ws = get_sheet("Receitas")
        ensure_headers(ws, ["Data", "Mês", "Descrição", "Valor"])
        ws.append_row([date_str, month, desc, amount])
        await update.message.reply_text(
            f"💚 *Entrada registrada!*\n📋 {desc}\n💰 R$ {amount:,.2f}\n📅 {date_str}",
            parse_mode="Markdown"
        )

    # PARCELA
    elif lower.startswith("parcela"):
        clean  = re.sub(r"^parcela\s*", "", lower).strip()
        amount = extract_amount(clean)
        day    = extract_day(clean)
        mes_fim, ano_fim = extract_ate(clean)
        desc   = re.sub(r"(dia\s+\d{1,2}|at[eé]\s+\d{2}/\d{4})", "", remove_amount(clean)).strip().title() or "Parcela"

        if not amount or not mes_fim:
            await update.message.reply_text(
                "Não entendi.\nEx: `parcela colchão 350 dia 5 até 09/2026`\n"
                "ou: `parcela C6 209,76 até 09/2026`",
                parse_mode="Markdown"
            )
            return

        restantes = parcelas_restantes(mes_fim, ano_fim)
        venc = f"Dia {day}" if day else "—"
        ate  = f"{mes_fim:02d}/{ano_fim}"

        ws = get_sheet("Parcelados")
        ensure_headers(ws, ["Descrição", "Valor", "Vencimento", "Até", "Adicionado em"])
        ws.append_row([desc, amount, venc, ate, date_str])
        await update.message.reply_text(
            f"🔄 *Parcela cadastrada!*\n"
            f"📋 {desc}\n"
            f"💰 R$ {amount:,.2f}/mês\n"
            f"📅 Vence: {venc} | Até: {ate}\n"
            f"🔢 Parcelas restantes: {restantes}",
            parse_mode="Markdown"
        )

    # NOVA CONTA
    elif lower.startswith("nova conta") or lower.startswith("nova "):
        clean  = re.sub(r"^(nova conta|nova)\s*", "", lower).strip()
        amount = extract_amount(clean)
        day    = extract_day(clean)
        desc   = re.sub(r"dia\s+\d{1,2}", "", remove_amount(clean)).strip().title() or "Conta"

        if not amount:
            await update.message.reply_text(
                "Não entendi.\nEx: `nova conta escola 650 dia 5`", parse_mode="Markdown"
            )
            return

        venc = f"Dia {day}" if day else "—"
        ws   = get_sheet("Contas Fixas")
        ensure_headers(ws, ["Descrição", "Valor", "Vencimento", "Adicionado em"])
        ws.append_row([desc, amount, venc, date_str])
        await update.message.reply_text(
            f"📌 *Conta fixa adicionada!*\n📋 {desc}\n💰 R$ {amount:,.2f}\n📅 Vence: {venc}",
            parse_mode="Markdown"
        )

    else:
        await update.message.reply_text(
            "Não entendi. Tente:\n\n"
            "💳 `paguei C6 209,76`\n"
            "💸 `gastei 80 padaria`\n"
            "💚 `recebi 5000 salário`\n"
            "📌 `nova conta escola 650 dia 5`\n\n"
            "📊 /saldo   📋 /contas   📈 /resumo",
            parse_mode="Markdown"
        )

# ── Main ──────────────────────────────────────────────────

def main():
    import asyncio
    asyncio.set_event_loop(asyncio.new_event_loop())
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start",  start))
    app.add_handler(CommandHandler("saldo",  cmd_saldo))
    app.add_handler(CommandHandler("contas", cmd_contas))
    app.add_handler(CommandHandler("resumo", cmd_resumo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    print("FINA Bot rodando...")
    app.run_polling()

if __name__ == "__main__":
    main()

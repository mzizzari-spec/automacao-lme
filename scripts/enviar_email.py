"""
Script: enviar_email.py
Função: Gera tabela HTML com dados do mês atual e envia por e-mail
Roda: Todo dia às 7h30 via GitHub Actions
"""

import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
DESTINATARIOS = os.environ.get("EMAIL_DESTINATARIOS", "").split(",")
DASHBOARD_URL = "https://mzizzari-spec.github.io/automacao-lme"
MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def conectar_google_sheets():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise ValueError("Secret GOOGLE_CREDENTIALS_JSON não encontrado!")
    creds_dict = json.loads(creds_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def fmt(v, dec=2):
    if v is None or v == "":
        return "-"
    try:
        s = str(v).strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        num = float(s)
        formatted = f"{num:,.{dec}f}"
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "-"


def fmt_var(v):
    """Formata variação percentual com cor."""
    if v is None or v == "":
        return "-"
    try:
        s = str(v).strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        num = float(s) * 100
        sinal = "+" if num > 0 else ""
        cor = "#15803d" if num > 0 else "#dc2626" if num < 0 else "#6b7280"
        return f'<span style="color:{cor};font-size:10px;">{sinal}{num:.2f}%</span>'
    except:
        return "-"


def obter_dados_mes_atual(client):
    hoje = datetime.now()
    nome_aba = f"{MESES_PT[hoje.month-1]}/{hoje.year}"
    planilha = client.open_by_key(GOOGLE_SHEET_ID)
    try:
        aba = planilha.worksheet(nome_aba)
    except gspread.WorksheetNotFound:
        raise ValueError(f"Aba '{nome_aba}' não encontrada!")
    return aba.get_all_values(), nome_aba


def gerar_html_email(dados, nome_mes):
    hoje = datetime.now()

    linhas_dias = [l for l in dados[1:] if len(l) > 2 and l[2] in ("Real", "Projetado")]
    linhas_semana = [l for l in dados[1:] if len(l) > 2 and l[2] == "" and l[0] == "Média Semana"]
    media_real = next((l for l in dados if len(l) > 0 and l[0] == "Média Real"), None)
    media_proj = next((l for l in dados if len(l) > 0 and l[0] == "Média Projetada"), None)
    reais = [l for l in linhas_dias if l[2] == "Real"]
    projetados = [l for l in linhas_dias if l[2] == "Projetado"]

    def card(label, val, var, cor, unidade):
        return f"""
        <td style="padding:0 8px 0 0;" width="33%">
          <div style="background:#fff;border:1px solid #e2e4ea;border-radius:8px;padding:14px;">
            <div style="font-size:10px;color:#6b7280;text-transform:uppercase;margin-bottom:4px;">{label}</div>
            <div style="font-size:16px;font-weight:600;color:{cor};font-family:monospace;">{fmt(val, 2)}</div>
            <div style="font-size:10px;color:#6b7280;margin-top:2px;">{unidade} &nbsp; {fmt_var(var)}</div>
          </div>
        </td>"""

    # Cards médias
    cards_real = ""
    cards_proj = ""
    if media_real and len(media_real) > 12:
        cards_real = f"""
        <tr>
          {card("Cobre Real", media_real[3], media_real[4], "#c45e1a", "US$/t")}
          {card("Alumínio Real", media_real[5], media_real[6], "#2b6cb0", "US$/t")}
          {card("Dólar Real", media_real[7], media_real[8], "#1a7a42", "R$/US$")}
        </tr>"""
    if media_proj and len(media_proj) > 12:
        cards_proj = f"""
        <tr>
          {card("Cobre Proj.", media_proj[3], media_proj[4], "#c45e1a", "US$/t")}
          {card("Alumínio Proj.", media_proj[5], media_proj[6], "#2b6cb0", "US$/t")}
          {card("Dólar Proj.", media_proj[7], media_proj[8], "#1a7a42", "R$/US$")}
        </tr>"""

    # Linhas da tabela
    linhas_html = ""
    semana_idx = 0
    semanas = [l for l in dados[1:] if len(l) > 0 and l[0] == "Média Semana"]

    for l in dados[1:]:
        if not l or len(l) < 3:
            continue

        tipo = l[2] if len(l) > 2 else ""

        if tipo in ("Real", "Projetado"):
            if tipo == "Real":
                badge = '<span style="background:#dcfce7;color:#15803d;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600;">Real</span>'
                bg = "#ffffff"
            else:
                badge = '<span style="background:#dbeafe;color:#1d4ed8;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600;">Proj.</span>'
                bg = "#f8faff"

            def td(col, dec=2):
                val = l[col] if len(l) > col else ""
                return f'<td style="padding:7px 10px;font-size:11px;font-family:monospace;">{fmt(val, dec)}</td>'

            def td_var(col):
                val = l[col] if len(l) > col else ""
                return f'<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(val)}</td>'

            linhas_html += f"""
            <tr style="background:{bg};border-bottom:1px solid #e2e4ea;">
              <td style="padding:7px 10px;font-size:11px;color:#6b7280;">{l[0]}</td>
              <td style="padding:7px 10px;font-size:11px;">{l[1]}</td>
              <td style="padding:7px 10px;">{badge}</td>
              {td(3, 0)}{td_var(4)}
              {td(5, 0)}{td_var(6)}
              {td(7, 4)}{td_var(8)}
              {td(9, 4)}{td_var(10)}
              {td(11, 4)}{td_var(12)}
            </tr>"""

        elif l[0] == "Média Semana":
            def td_med(col, dec=2):
                val = l[col] if len(l) > col else ""
                return f'<td style="padding:6px 10px;font-size:11px;font-family:monospace;color:#6b7280;">{fmt(val, dec)}</td>'

            linhas_html += f"""
            <tr style="background:#f8f9fa;border-bottom:2px solid #e2e4ea;">
              <td colspan="3" style="padding:6px 10px;font-size:10px;color:#9aa0b4;font-style:italic;">Média da semana</td>
              {td_med(3, 0)}<td></td>
              {td_med(5, 0)}<td></td>
              {td_med(7, 4)}<td></td>
              {td_med(9, 4)}<td></td>
              {td_med(11, 4)}<td></td>
            </tr>"""

    # Linha média real e projetada na tabela
    def td_media(linha, col, dec=2):
        val = linha[col] if linha and len(linha) > col else ""
        return f'<td style="padding:7px 10px;font-size:11px;font-weight:600;font-family:monospace;">{fmt(val, dec)}</td>'

    def td_media_var(linha, col):
        val = linha[col] if linha and len(linha) > col else ""
        return f'<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(val)}</td>'

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f5f6f8;font-family:Inter,Arial,sans-serif;">
<div style="max-width:960px;margin:0 auto;padding:28px 16px;">

  <div style="margin-bottom:20px;">
    <h1 style="font-size:20px;font-weight:700;color:#1a1d2e;margin:0;">LME <span style="color:#c45e1a;">Metais</span></h1>
    <p style="font-size:12px;color:#6b7280;margin:4px 0 0;">Cotações {nome_mes} — {hoje.strftime('%d/%m/%Y %H:%M')}</p>
  </div>

  <table width="100%" cellpadding="0" cellspacing="8" style="margin-bottom:20px;">
    <tr><td colspan="3" style="padding:4px 0 8px;font-size:10px;color:#6b7280;text-transform:uppercase;letter-spacing:0.1em;">Média Real</td></tr>
    {cards_real}
    <tr><td colspan="3" style="padding:12px 0 8px;font-size:10px;color:#6b7280;text-transform:uppercase;letter-spacing:0.1em;">Média Projetada</td></tr>
    {cards_proj}
  </table>

  <div style="background:#fff;border:1px solid #e2e4ea;border-radius:10px;overflow:hidden;margin-bottom:20px;">
    <div style="padding:12px 16px;border-bottom:1px solid #e2e4ea;display:flex;justify-content:space-between;">
      <span style="font-size:13px;font-weight:600;">Cotações diárias — {nome_mes}</span>
      <span style="font-size:11px;color:#6b7280;">{len(reais)} reais · {len(projetados)} projetados</span>
    </div>
    <table width="100%" cellpadding="0" cellspacing="0">
      <thead>
        <tr style="background:#f0f1f4;">
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Data</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Dia</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Tipo</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Cobre</th>
          <th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Alumínio</th>
          <th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Dólar</th>
          <th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Cu R$/kg</th>
          <th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Al R$/kg</th>
          <th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>
        </tr>
      </thead>
      <tbody>
        {linhas_html}
        <tr style="background:#f0f1f4;border-top:2px solid #e2e4ea;">
          <td colspan="3" style="padding:7px 10px;font-size:11px;font-weight:600;color:#15803d;">Média Real</td>
          {td_media(media_real, 3, 0)}{td_media_var(media_real, 4)}
          {td_media(media_real, 5, 0)}{td_media_var(media_real, 6)}
          {td_media(media_real, 7, 4)}{td_media_var(media_real, 8)}
          {td_media(media_real, 9, 4)}{td_media_var(media_real, 10)}
          {td_media(media_real, 11, 4)}{td_media_var(media_real, 12)}
        </tr>
        <tr style="background:#f0f1f4;">
          <td colspan="3" style="padding:7px 10px;font-size:11px;font-weight:600;color:#1d4ed8;">Média Projetada</td>
          {td_media(media_proj, 3, 0)}{td_media_var(media_proj, 4)}
          {td_media(media_proj, 5, 0)}{td_media_var(media_proj, 6)}
          {td_media(media_proj, 7, 4)}{td_media_var(media_proj, 8)}
          {td_media(media_proj, 9, 4)}{td_media_var(media_proj, 10)}
          {td_media(media_proj, 11, 4)}{td_media_var(media_proj, 12)}
        </tr>
      </tbody>
    </table>
  </div>

  <div style="text-align:center;padding:16px;">
    <a href="{DASHBOARD_URL}" style="background:#c45e1a;color:#fff;padding:11px 24px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:600;">
      Ver Dashboard Completo
    </a>
  </div>

  <div style="text-align:center;padding:12px 0;font-size:10px;color:#9aa0b4;">
    Enviado automaticamente pelo sistema LME Automação
  </div>
</div>
</body>
</html>"""
    return html


def enviar_email(html, nome_mes):
    assunto = f"LME Metais — {nome_mes} ({datetime.now().strftime('%d/%m/%Y')})"
    msg = MIMEMultipart('alternative')
    msg['Subject'] = assunto
    msg['From'] = GMAIL_USER
    msg['To'] = ", ".join(DESTINATARIOS)
    msg.attach(MIMEText(html, 'html', 'utf-8'))
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.sendmail(GMAIL_USER, DESTINATARIOS, msg.as_string())
    print(f"✅ E-mail enviado para: {', '.join(DESTINATARIOS)}")


def main():
    print("=" * 50)
    print(f"Iniciando envio — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 50)
    client = conectar_google_sheets()
    dados, nome_mes = obter_dados_mes_atual(client)
    html = gerar_html_email(dados, nome_mes)
    enviar_email(html, nome_mes)
    print("✅ Concluído!")


if __name__ == "__main__":
    main()

"""
Script: enviar_email.py
Função: Gera tabela HTML com dados do mês atual e envia por e-mail
Roda: Todo dia às 7h via GitHub Actions
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


def to_float(v):
    if v is None or v == "":
        return None
    try:
        s = str(v).strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)
    except:
        return None


def fmt(v, dec=2):
    n = to_float(v)
    if n is None:
        return "-"
    formatted = f"{n:,.{dec}f}"
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_var(v):
    n = to_float(v)
    if n is None:
        return "-"
    num = n * 100
    sinal = "+" if num > 0 else ""
    cor = "#15803d" if num > 0 else "#dc2626" if num < 0 else "#6b7280"
    return f'<span style="color:{cor};font-size:10px;">{sinal}{num:.2f}%</span>'


def calc_var(atual, anterior):
    a = to_float(atual)
    b = to_float(anterior)
    if a is None or b is None or b == 0:
        return None
    return (a - b) / b


def obter_dados_mes_atual(client):
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year
    nome_aba = f"{MESES_PT[mes_atual-1]}/{ano_atual}"

    mes_ant = mes_atual - 1
    ano_ant = ano_atual
    if mes_ant == 0:
        mes_ant = 12
        ano_ant -= 1
    nome_mes_ant = f"{MESES_PT[mes_ant-1]}/{ano_ant}"

    planilha = client.open_by_key(GOOGLE_SHEET_ID)
    try:
        aba = planilha.worksheet(nome_aba)
    except gspread.WorksheetNotFound:
        raise ValueError(f"Aba '{nome_aba}' não encontrada!")

    dados = aba.get_all_values()

    # Busca médias do mês anterior na aba Historico
    # O Historico tem: Mes | Dolar | Var.Dolar | Cobre | Var.Cobre | Aluminio | Var.Al | CobreKg | Var.CobreKg | AlKg | Var.AlKg
    # Precisamos converter para o formato da aba mensal:
    # col3=Cobre, col5=Aluminio, col7=Dolar, col9=CobreKg, col11=AlKg
    media_real_ant = None
    try:
        hist = planilha.worksheet("Historico")
        dados_hist = hist.get_all_values()
        # Busca linha do mês anterior
        # Pula cabeçalho (linha 0) e busca pelo nome do mês
        linha_ant = next((l for l in dados_hist[1:] if len(l) > 0 and l[0] == nome_mes_ant), None)
        print(f"DEBUG: buscando '{nome_mes_ant}' no Historico, encontrou: {linha_ant}")
        if linha_ant:
            # Historico: [Mes, Dolar, Var.Dolar, Cobre, Var.Cobre, Aluminio, Var.Al, CobreKg, Var.CobreKg, AlKg, Var.AlKg]
            # Monta lista no formato da aba mensal para usar col3,5,7,9,11
            media_real_ant = [''] * 13
            media_real_ant[3] = linha_ant[3]   # Cobre
            media_real_ant[5] = linha_ant[5]   # Aluminio
            media_real_ant[7] = linha_ant[1]   # Dolar
            media_real_ant[9] = linha_ant[7]   # Cobre R$/kg
            media_real_ant[11] = linha_ant[9]  # Aluminio R$/kg
            print(f"✅ Médias de {nome_mes_ant} encontradas no Historico")
    except Exception as e:
        print(f"⚠️  Historico não encontrado: {e}")

    return dados, nome_aba, media_real_ant


def gerar_html_email(dados, nome_mes, media_real_ant=None):
    hoje = datetime.now()

    linhas_dias = [l for l in dados[1:] if len(l) > 2 and l[2] in ("Real", "Projetado")]
    media_real = next((l for l in dados if len(l) > 0 and l[0] == "Média Real"), None)
    media_proj = next((l for l in dados if len(l) > 0 and l[0] == "Média Projetada"), None)
    reais = [l for l in linhas_dias if l[2] == "Real"]
    projetados = [l for l in linhas_dias if l[2] == "Projetado"]

    # Coleta todas as linhas de média semana para calcular variação entre semanas
    semanas = [l for l in dados[1:] if len(l) > 0 and l[0] == "Média Semana"]

    def card(label, val, var, cor, unidade, dec=2, width="20%"):
        return f"""
        <td style="padding:0 4px;" width="{width}">
          <div style="background:#fff;border:1px solid #e2e4ea;border-radius:8px;padding:12px;">
            <div style="font-size:10px;color:#6b7280;text-transform:uppercase;margin-bottom:4px;">{label}</div>
            <div style="font-size:15px;font-weight:600;color:{cor};font-family:monospace;">{fmt(val, dec)}</div>
            <div style="font-size:10px;color:#6b7280;margin-top:2px;">{unidade} &nbsp; {fmt_var(var)}</div>
          </div>
        </td>"""

    cards_real = ""
    cards_proj = ""
    if media_real and len(media_real) > 12:
        def var_r(col):
            atual = to_float(media_real[col]) if len(media_real) > col else None
            ant = to_float(media_real_ant[col]) if media_real_ant and len(media_real_ant) > col else None
            return calc_var(atual, ant)
        cards_real = f"""
        <tr>
          {card("Cobre Real", media_real[3], var_r(3), "#c45e1a", "US$/t", width="20%")}
          {card("Alumínio Real", media_real[5], var_r(5), "#2b6cb0", "US$/t", width="20%")}
          {card("Dólar Real", media_real[7], var_r(7), "#1a7a42", "R$/US$", dec=4, width="20%")}
          {card("Cobre R$/kg Real", media_real[9], var_r(9), "#c45e1a", "R$/kg", width="20%")}
          {card("Alumínio R$/kg Real", media_real[11], var_r(11), "#2b6cb0", "R$/kg", width="20%")}
        </tr>"""
    if media_proj and len(media_proj) > 12:
        def var_p(col):
            proj = to_float(media_proj[col]) if len(media_proj) > col else None
            real = to_float(media_real[col]) if media_real and len(media_real) > col else None
            return calc_var(proj, real)
        <tr>
          {card("Cobre Proj.", media_proj[3], var_p(3), "#c45e1a", "US$/t", width="20%")}
          {card("Alumínio Proj.", media_proj[5], var_p(5), "#2b6cb0", "US$/t", width="20%")}
          {card("Dólar Proj.", media_proj[7], var_p(7), "#1a7a42", "R$/US$", dec=4, width="20%")}
          {card("Cobre R$/kg Proj.", media_proj[9], var_p(9), "#c45e1a", "R$/kg", width="20%")}
          {card("Alumínio R$/kg Proj.", media_proj[11], var_p(11), "#2b6cb0", "R$/kg", width="20%")}
        </tr>"""

    # Gera linhas da tabela
    linhas_html = ""
    semana_idx = 0

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

            def td(col, dec=2, linha=l):
                val = linha[col] if len(linha) > col else ""
                return f'<td style="padding:7px 10px;font-size:11px;font-family:monospace;">{fmt(val, dec)}</td>'

            def td_var(col, linha=l):
                val = linha[col] if len(linha) > col else ""
                return f'<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(val)}</td>'

            linhas_html += f"""
            <tr style="background:{bg};border-bottom:1px solid #e2e4ea;">
              <td style="padding:7px 10px;font-size:11px;color:#6b7280;">{l[0]}</td>
              <td style="padding:7px 10px;font-size:11px;">{l[1]}</td>
              <td style="padding:7px 10px;">{badge}</td>
              {td(3, 0)}{td_var(4)}
              {td(5, 0)}{td_var(6)}
              {td(7, 4)}{td_var(8)}
              {td(9, 2)}{td_var(10)}
              {td(11, 2)}{td_var(12)}
            </tr>"""

        elif l[0] == "Média Semana":
            # Calcula variação em relação à semana anterior
            sem_ant = semanas[semana_idx - 1] if semana_idx > 0 else None

            def td_sem(col, dec=2, linha=l):
                val = linha[col] if len(linha) > col else ""
                return f'<td style="padding:6px 10px;font-size:11px;font-family:monospace;color:#6b7280;">{fmt(val, dec)}</td>'

            def td_sem_var(col, linha=l, ant=sem_ant):
                val = linha[col] if len(linha) > col else None
                val_ant = ant[col] if ant and len(ant) > col else None
                v = calc_var(val, val_ant)
                return f'<td style="padding:6px 10px;font-size:10px;text-align:center;">{fmt_var(v)}</td>'

            linhas_html += f"""
            <tr style="background:#f8f9fa;border-bottom:2px solid #e2e4ea;">
              <td colspan="3" style="padding:6px 10px;font-size:10px;color:#9aa0b4;font-style:italic;">Média da semana</td>
              {td_sem(3, 0)}{td_sem_var(3)}
              {td_sem(5, 0)}{td_sem_var(5)}
              {td_sem(7, 4)}{td_sem_var(7)}
              {td_sem(9, 2)}{td_sem_var(9)}
              {td_sem(11, 2)}{td_sem_var(11)}
            </tr>"""
            semana_idx += 1

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

  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;padding-bottom:16px;border-bottom:1px solid #e2e4ea;">
    <div>
      <h1 style="font-size:20px;font-weight:700;color:#1a1d2e;margin:0;">LME <span style="color:#1a6080;">Metais</span></h1>
      <p style="font-size:12px;color:#6b7280;margin:4px 0 0;">Cotações {nome_mes} - {hoje.strftime('%d/%m/%Y %H:%M')}</p>
    </div>
    <img src="https://mzizzari-spec.github.io/automacao-lme/GMC-logo.png" alt="Grupo Melo Cordeiro" height="48" width="160" style="height:48px;width:160px;object-fit:contain;display:block;">
  </div>

  <table width="100%" cellpadding="0" cellspacing="8" style="margin-bottom:20px;">
    <tr><td colspan="5" style="padding:4px 0 8px;font-size:10px;color:#6b7280;text-transform:uppercase;letter-spacing:0.1em;">Média Real</td></tr>
    {cards_real}
    <tr><td colspan="5" style="padding:12px 0 8px;font-size:10px;color:#6b7280;text-transform:uppercase;letter-spacing:0.1em;">Média Projetada</td></tr>
    {cards_proj}
  </table>

  <div style="background:#fff;border:1px solid #e2e4ea;border-radius:10px;overflow:hidden;margin-bottom:20px;">
    <div style="padding:12px 16px;border-bottom:1px solid #e2e4ea;display:flex;justify-content:space-between;">
      <span style="font-size:13px;font-weight:600;">Cotações diárias - {nome_mes}</span>
      <span style="font-size:11px;color:#6b7280;">{len(reais)} reais - {len(projetados)} projetados</span>
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
          {td_media(media_real, 3, 0)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_real[3]) if media_real else None, to_float(media_real_ant[3]) if media_real_ant else None))}</td>
          {td_media(media_real, 5, 0)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_real[5]) if media_real else None, to_float(media_real_ant[5]) if media_real_ant else None))}</td>
          {td_media(media_real, 7, 4)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_real[7]) if media_real else None, to_float(media_real_ant[7]) if media_real_ant else None))}</td>
          {td_media(media_real, 9, 2)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_real[9]) if media_real else None, to_float(media_real_ant[9]) if media_real_ant else None))}</td>
          {td_media(media_real, 11, 2)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_real[11]) if media_real else None, to_float(media_real_ant[11]) if media_real_ant else None))}</td>
        </tr>
        <tr style="background:#f0f1f4;">
          <td colspan="3" style="padding:7px 10px;font-size:11px;font-weight:600;color:#1d4ed8;">Média Projetada</td>
          {td_media(media_proj, 3, 0)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_proj[3]) if media_proj else None, to_float(media_real[3]) if media_real else None))}</td>
          {td_media(media_proj, 5, 0)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_proj[5]) if media_proj else None, to_float(media_real[5]) if media_real else None))}</td>
          {td_media(media_proj, 7, 4)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_proj[7]) if media_proj else None, to_float(media_real[7]) if media_real else None))}</td>
          {td_media(media_proj, 9, 2)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_proj[9]) if media_proj else None, to_float(media_real[9]) if media_real else None))}</td>
          {td_media(media_proj, 11, 2)}<td style="padding:7px 10px;font-size:10px;text-align:center;">{fmt_var(calc_var(to_float(media_proj[11]) if media_proj else None, to_float(media_real[11]) if media_real else None))}</td>
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
    msg['To'] = GMAIL_USER
    msg['Bcc'] = ", ".join(DESTINATARIOS)
    msg.attach(MIMEText(html, 'html', 'utf-8'))
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.sendmail(GMAIL_USER, DESTINATARIOS, msg.as_string())
    print(f"✅ E-mail enviado!")


def main():
    print("=" * 50)
    print(f"Iniciando envio — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 50)
    client = conectar_google_sheets()
    dados, nome_mes, media_real_ant = obter_dados_mes_atual(client)
    html = gerar_html_email(dados, nome_mes, media_real_ant)
    enviar_email(html, nome_mes)
    print("✅ Concluído!")


if __name__ == "__main__":
    main()

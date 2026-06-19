"""
Script: enviar_email.py
Função: Gera uma tabela HTML com os dados do mês atual e envia por e-mail
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


# ─────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────
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
        num = float(str(v).replace(",", "."))
        return f"{num:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
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

    dados = aba.get_all_values()
    return dados, nome_aba


def calcular_medias(linhas):
    reais = [l for l in linhas if len(l) > 2 and l[2] == "Real"]
    todos = [l for l in linhas if len(l) > 2 and l[2] in ("Real", "Projetado")]

    def media_col(rows, col):
        vals = []
        for r in rows:
            if col < len(r) and r[col] not in ("", None):
                try:
                    vals.append(float(str(r[col]).replace(",", ".")))
                except:
                    pass
        return sum(vals) / len(vals) if vals else None

    return {
        "real": {
            "cobre": media_col(reais, 3),
            "aluminio": media_col(reais, 4),
            "dolar": media_col(reais, 5),
            "cobre_kg": media_col(reais, 6),
            "aluminio_kg": media_col(reais, 7),
        },
        "projetada": {
            "cobre": media_col(todos, 3),
            "aluminio": media_col(todos, 4),
            "dolar": media_col(todos, 5),
            "cobre_kg": media_col(todos, 6),
            "aluminio_kg": media_col(todos, 7),
        }
    }


def gerar_html_email(dados, nome_mes, medias):
    hoje = datetime.now()

    # Filtra só linhas de dados (Real e Projetado)
    linhas_dados = [l for l in dados[1:] if len(l) > 2 and l[2] in ("Real", "Projetado")]
    reais = [l for l in linhas_dados if l[2] == "Real"]
    projetados = [l for l in linhas_dados if l[2] == "Projetado"]

    # Gera linhas da tabela
    linhas_html = ""
    for l in linhas_dados:
        tipo = l[2] if len(l) > 2 else ""
        if tipo == "Real":
            badge = '<span style="background:#dcfce7;color:#15803d;padding:2px 7px;border-radius:4px;font-size:10px;font-weight:600;text-transform:uppercase;">Real</span>'
            bg = "#ffffff"
        else:
            badge = '<span style="background:#dbeafe;color:#1d4ed8;padding:2px 7px;border-radius:4px;font-size:10px;font-weight:600;text-transform:uppercase;">Projetado</span>'
            bg = "#f8faff"

        linhas_html += f"""
        <tr style="background:{bg};border-bottom:1px solid #e2e4ea;">
            <td style="padding:8px 12px;font-size:12px;color:#6b7280;">{l[0] if len(l)>0 else ''}</td>
            <td style="padding:8px 12px;font-size:12px;">{l[1] if len(l)>1 else ''}</td>
            <td style="padding:8px 12px;">{badge}</td>
            <td style="padding:8px 12px;font-size:12px;font-family:monospace;">{fmt(l[3] if len(l)>3 else None, 2)}</td>
            <td style="padding:8px 12px;font-size:12px;font-family:monospace;">{fmt(l[4] if len(l)>4 else None, 2)}</td>
            <td style="padding:8px 12px;font-size:12px;font-family:monospace;">{fmt(l[5] if len(l)>5 else None, 2)}</td>
            <td style="padding:8px 12px;font-size:12px;font-family:monospace;">{fmt(l[6] if len(l)>6 else None, 4)}</td>
            <td style="padding:8px 12px;font-size:12px;font-family:monospace;">{fmt(l[7] if len(l)>7 else None, 4)}</td>
        </tr>"""

    m = medias
    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f5f6f8;font-family:Inter,Arial,sans-serif;">
<div style="max-width:900px;margin:0 auto;padding:32px 16px;">

  <!-- HEADER -->
  <div style="margin-bottom:24px;">
    <h1 style="font-size:22px;font-weight:700;color:#1a1d2e;margin:0;">
      LME <span style="color:#c45e1a;">Metais</span>
    </h1>
    <p style="font-size:13px;color:#6b7280;margin:4px 0 0;">
      Cotações de {nome_mes} — Gerado em {hoje.strftime('%d/%m/%Y às %H:%M')}
    </p>
  </div>

  <!-- CARDS DE MÉDIA -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
    <tr>
      <td style="padding:4px 0 8px;font-size:11px;color:#6b7280;text-transform:uppercase;letter-spacing:0.1em;" colspan="3">
        Média Real
      </td>
    </tr>
    <tr>
      <td style="padding:0 8px 0 0;" width="33%">
        <div style="background:#fff;border:1px solid #e2e4ea;border-radius:8px;padding:14px;">
          <div style="font-size:11px;color:#6b7280;text-transform:uppercase;margin-bottom:6px;">Cobre</div>
          <div style="font-size:18px;font-weight:600;color:#c45e1a;font-family:monospace;">{fmt(m['real']['cobre'], 2)}</div>
          <div style="font-size:11px;color:#6b7280;margin-top:4px;">US$/t</div>
        </div>
      </td>
      <td style="padding:0 8px;" width="33%">
        <div style="background:#fff;border:1px solid #e2e4ea;border-radius:8px;padding:14px;">
          <div style="font-size:11px;color:#6b7280;text-transform:uppercase;margin-bottom:6px;">Alumínio</div>
          <div style="font-size:18px;font-weight:600;color:#2b6cb0;font-family:monospace;">{fmt(m['real']['aluminio'], 2)}</div>
          <div style="font-size:11px;color:#6b7280;margin-top:4px;">US$/t</div>
        </div>
      </td>
      <td style="padding:0 0 0 8px;" width="33%">
        <div style="background:#fff;border:1px solid #e2e4ea;border-radius:8px;padding:14px;">
          <div style="font-size:11px;color:#6b7280;text-transform:uppercase;margin-bottom:6px;">Dólar</div>
          <div style="font-size:18px;font-weight:600;color:#1a7a42;font-family:monospace;">{fmt(m['real']['dolar'], 2)}</div>
          <div style="font-size:11px;color:#6b7280;margin-top:4px;">R$/US$</div>
        </div>
      </td>
    </tr>
    <tr>
      <td style="padding:16px 0 8px;font-size:11px;color:#6b7280;text-transform:uppercase;letter-spacing:0.1em;" colspan="3">
        Média Projetada
      </td>
    </tr>
    <tr>
      <td style="padding:0 8px 0 0;" width="33%">
        <div style="background:#fff;border:1px solid #e2e4ea;border-radius:8px;padding:14px;">
          <div style="font-size:11px;color:#6b7280;text-transform:uppercase;margin-bottom:6px;">Cobre</div>
          <div style="font-size:18px;font-weight:600;color:#c45e1a;font-family:monospace;">{fmt(m['projetada']['cobre'], 2)}</div>
          <div style="font-size:11px;color:#6b7280;margin-top:4px;">US$/t</div>
        </div>
      </td>
      <td style="padding:0 8px;" width="33%">
        <div style="background:#fff;border:1px solid #e2e4ea;border-radius:8px;padding:14px;">
          <div style="font-size:11px;color:#6b7280;text-transform:uppercase;margin-bottom:6px;">Alumínio</div>
          <div style="font-size:18px;font-weight:600;color:#2b6cb0;font-family:monospace;">{fmt(m['projetada']['aluminio'], 2)}</div>
          <div style="font-size:11px;color:#6b7280;margin-top:4px;">US$/t</div>
        </div>
      </td>
      <td style="padding:0 0 0 8px;" width="33%">
        <div style="background:#fff;border:1px solid #e2e4ea;border-radius:8px;padding:14px;">
          <div style="font-size:11px;color:#6b7280;text-transform:uppercase;margin-bottom:6px;">Dólar</div>
          <div style="font-size:18px;font-weight:600;color:#1a7a42;font-family:monospace;">{fmt(m['projetada']['dolar'], 2)}</div>
          <div style="font-size:11px;color:#6b7280;margin-top:4px;">R$/US$</div>
        </div>
      </td>
    </tr>
  </table>

  <!-- TABELA DE DADOS -->
  <div style="background:#fff;border:1px solid #e2e4ea;border-radius:10px;overflow:hidden;margin-bottom:24px;">
    <div style="padding:14px 16px;border-bottom:1px solid #e2e4ea;">
      <span style="font-size:13px;font-weight:600;color:#1a1d2e;">Cotações diárias — {nome_mes}</span>
      <span style="font-size:11px;color:#6b7280;float:right;">{len(reais)} dias reais · {len(projetados)} projetados</span>
    </div>
    <table width="100%" cellpadding="0" cellspacing="0">
      <thead>
        <tr style="background:#f0f1f4;">
          <th style="padding:9px 12px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Data</th>
          <th style="padding:9px 12px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Dia</th>
          <th style="padding:9px 12px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Tipo</th>
          <th style="padding:9px 12px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Cobre US$/t</th>
          <th style="padding:9px 12px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Alumínio US$/t</th>
          <th style="padding:9px 12px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Dólar R$/US$</th>
          <th style="padding:9px 12px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Cobre R$/kg</th>
          <th style="padding:9px 12px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Alumínio R$/kg</th>
        </tr>
      </thead>
      <tbody>
        {linhas_html}
        <!-- MÉDIA REAL -->
        <tr style="background:#f0f1f4;border-top:2px solid #e2e4ea;">
          <td colspan="3" style="padding:9px 12px;font-size:12px;font-weight:600;color:#15803d;">Média Real</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#15803d;font-family:monospace;">{fmt(m['real']['cobre'], 2)}</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#15803d;font-family:monospace;">{fmt(m['real']['aluminio'], 2)}</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#15803d;font-family:monospace;">{fmt(m['real']['dolar'], 2)}</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#15803d;font-family:monospace;">{fmt(m['real']['cobre_kg'], 4)}</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#15803d;font-family:monospace;">{fmt(m['real']['aluminio_kg'], 4)}</td>
        </tr>
        <!-- MÉDIA PROJETADA -->
        <tr style="background:#f0f1f4;">
          <td colspan="3" style="padding:9px 12px;font-size:12px;font-weight:600;color:#1d4ed8;">Média Projetada</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#1d4ed8;font-family:monospace;">{fmt(m['projetada']['cobre'], 2)}</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#1d4ed8;font-family:monospace;">{fmt(m['projetada']['aluminio'], 2)}</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#1d4ed8;font-family:monospace;">{fmt(m['projetada']['dolar'], 2)}</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#1d4ed8;font-family:monospace;">{fmt(m['projetada']['cobre_kg'], 4)}</td>
          <td style="padding:9px 12px;font-size:12px;font-weight:600;color:#1d4ed8;font-family:monospace;">{fmt(m['projetada']['aluminio_kg'], 4)}</td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- LINK DASHBOARD -->
  <div style="text-align:center;padding:20px;">
    <a href="{DASHBOARD_URL}" style="background:#c45e1a;color:#ffffff;padding:12px 28px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:600;">
      Ver Dashboard Completo
    </a>
  </div>

  <!-- RODAPÉ -->
  <div style="text-align:center;padding:16px 0;font-size:11px;color:#9aa0b4;">
    Enviado automaticamente pelo sistema LME Automação
  </div>

</div>
</body>
</html>
"""
    return html


def enviar_email(html, nome_mes):
    assunto = f"LME Metais — Cotações {nome_mes} ({datetime.now().strftime('%d/%m/%Y')})"

    msg = MIMEMultipart('alternative')
    msg['Subject'] = assunto
    msg['From'] = GMAIL_USER
    msg['To'] = ", ".join(DESTINATARIOS)

    parte_html = MIMEText(html, 'html', 'utf-8')
    msg.attach(parte_html)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.sendmail(GMAIL_USER, DESTINATARIOS, msg.as_string())

    print(f"✅ E-mail enviado para: {', '.join(DESTINATARIOS)}")


def main():
    print("=" * 50)
    print(f"Iniciando envio de e-mail — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 50)

    client = conectar_google_sheets()
    dados, nome_mes = obter_dados_mes_atual(client)
    medias = calcular_medias(dados[1:])
    html = gerar_html_email(dados, nome_mes, medias)
    enviar_email(html, nome_mes)

    print("=" * 50)
    print("✅ Concluído!")
    print("=" * 50)


if __name__ == "__main__":
    main()

import os
from datetime import datetime, timedelta

import pandas as pd
import win32com.client

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGENS_DIR = os.path.join(BASE_DIR, "imagens")
BASE_DETENTORAS_PATH = os.path.join(BASE_DIR, "base_detentoras.xlsx")
DETENTORAS_PATH = os.path.join(BASE_DIR, "detentoras.xlsx")
HISTORICO_PATH = os.path.join(BASE_DIR, "historico_envios.xlsx")
ESTADOS_FILTRO = {"BA", "SE", "AL"}
EXTENSOES_IMAGEM = {".png", ".jpg", ".jpeg"}
DIAS_LIMITE_COBRANCA = 10
VALORES_CORRIGIDO = {"SIM", "S", "X", "OK"}


def normalizar(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper()


def esta_corrigido(valor):
    return normalizar(valor) in VALORES_CORRIGIDO


def carregar_base_detentoras():
    df = pd.read_excel(
        BASE_DETENTORAS_PATH,
        sheet_name="Dados",
        usecols=[
            "UFSigla",
            "State / Province",
            "Detentora (Original)",
            "Detentora (Referência)",
            "ID Detentora",
            "Updated",
        ],
    )
    df = df[df["State / Province"].isin(ESTADOS_FILTRO)].copy()

    original = df["Detentora (Original)"]
    referencia = df["Detentora (Referência)"]
    vazio = original.isna() | (original.astype(str).str.strip() == "")
    df["detentora_resolvida"] = original.where(~vazio, referencia)

    df["Updated"] = pd.to_datetime(df["Updated"], dayfirst=True, errors="coerce")
    df["ufsigla_norm"] = df["UFSigla"].map(normalizar)

    df = df.sort_values("Updated").drop_duplicates(subset="ufsigla_norm", keep="last")
    return df.set_index("ufsigla_norm")


def carregar_emails_detentoras():
    df = pd.read_excel(DETENTORAS_PATH, sheet_name="Plan1")
    emails_por_detentora = {}
    for _, row in df.iterrows():
        detentora = normalizar(row["DETENTORA"])
        if not detentora or pd.isna(row["EMAIL"]):
            continue
        emails = [e.strip() for e in str(row["EMAIL"]).split(";") if e.strip()]
        if emails:
            emails_por_detentora[detentora] = emails
    return emails_por_detentora


def listar_grupos_imagens():
    """Cada subpasta de imagens/ representa um UFSigla; retorna {nome_da_pasta: [caminhos das imagens]}."""
    if not os.path.isdir(IMAGENS_DIR):
        return {}
    grupos = {}
    for nome in os.listdir(IMAGENS_DIR):
        caminho_pasta = os.path.join(IMAGENS_DIR, nome)
        if not os.path.isdir(caminho_pasta):
            continue
        imagens = []
        for arquivo in os.listdir(caminho_pasta):
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            _, ext = os.path.splitext(arquivo)
            if os.path.isfile(caminho_arquivo) and ext.lower() in EXTENSOES_IMAGEM:
                imagens.append(caminho_arquivo)
        if imagens:
            grupos[nome] = imagens
    return grupos


HTML_TEMPLATE = """\
<html>
<body style="margin:0; padding:0; background-color:#f2f2f2; font-family:Calibri,Arial,sans-serif;">
<table width="700" align="center" cellpadding="0" cellspacing="0" style="background-color:#ffffff; border-collapse:collapse;">
  <tr>
    <td bgcolor="#4B2E83" style="padding:22px 30px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="color:#ffffff; font-size:22px; font-weight:bold;">Zeladoria</td>
          <td align="right" style="color:#e6e0f0; font-size:11px; letter-spacing:1px;">ALERTA DE VEGETAÇÃO ALTA</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="padding:30px;">
      <p style="color:#4B2E83; font-size:20px; font-weight:bold; margin:0 0 8px 0;">Solicitação de Poda</p>
      <hr style="border:none; border-top:1px solid #d9d9d9; margin:0 0 18px 0;">
      <p style="color:#4B2E83; font-weight:bold; margin:0 0 14px 0;">Prezado(a),</p>
      <p style="color:#333333; font-size:14px; line-height:20px; margin:0 0 16px 0;">
        Por meio desta notificação, servimo-nos do presente para expor e requerer o que segue:
      </p>
      <table cellpadding="0" cellspacing="0" style="margin:0 0 16px 0;">
        <tr>
          <td width="4" bgcolor="#4B2E83"></td>
          <td style="padding:6px 0 6px 14px; color:#4B2E83; font-size:14px; line-height:20px;">
            Identificada <b>vegetação alta (mato alto)</b> no site <b>{ufsigla}</b> (ID Detentora <b>{id_detentora}</b>), presente(s) no(s) anexo(s), evidenciando que o acordo de manutenção e zeladoria está sendo descumprido.
          </td>
        </tr>
      </table>
      <p style="color:#333333; font-size:14px; line-height:20px; margin:0 0 16px 0;">
        A falta de poda no local já apresenta risco iminente de impactar diretamente a rede e interromper os serviços.
      </p>
      <p style="color:#333333; font-size:14px; line-height:20px; margin:0 0 24px 0;">
        Diante do exposto, solicitamos a validação imediata destas ocorrências e resolução das falhas apresentadas. O não cumprimento da poda no prazo estipulado configurará mora contratual definitiva.
      </p>
    </td>
  </tr>
  <tr>
    <td bgcolor="#4B2E83" align="center" style="padding:14px;">
      <span style="color:#ffffff; font-size:16px;">&#10022;</span>
    </td>
  </tr>
  <tr>
    <td style="padding:16px 30px; color:#999999; font-size:10px; line-height:14px;">
      Esta é uma notificação automática referente ao acordo de manutenção e zeladoria. Em caso de dúvidas, entre em contato com a equipe responsável.
    </td>
  </tr>
</table>
</body>
</html>
"""


COBRANCA_HTML_TEMPLATE = """\
<html>
<body style="margin:0; padding:0; background-color:#f2f2f2; font-family:Calibri,Arial,sans-serif;">
<table width="700" align="center" cellpadding="0" cellspacing="0" style="background-color:#ffffff; border-collapse:collapse;">
  <tr>
    <td bgcolor="#4B2E83" style="padding:22px 30px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="color:#ffffff; font-size:22px; font-weight:bold;">Zeladoria</td>
          <td align="right" style="color:#e6e0f0; font-size:11px; letter-spacing:1px;">COBRANÇA DE SOLUÇÃO</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="padding:30px;">
      <p style="color:#4B2E83; font-size:20px; font-weight:bold; margin:0 0 8px 0;">Cobrança - Solicitação de Poda</p>
      <hr style="border:none; border-top:1px solid #d9d9d9; margin:0 0 18px 0;">
      <p style="color:#4B2E83; font-weight:bold; margin:0 0 14px 0;">Prezado(a),</p>
      <p style="color:#333333; font-size:14px; line-height:20px; margin:0 0 16px 0;">
        Reiteramos a notificação enviada em <b>{data_envio}</b> referente à vegetação alta (mato alto) identificada no site <b>{ufsigla}</b> (ID Detentora <b>{id_detentora}</b>).
      </p>
      <table cellpadding="0" cellspacing="0" style="margin:0 0 16px 0;">
        <tr>
          <td width="4" bgcolor="#4B2E83"></td>
          <td style="padding:6px 0 6px 14px; color:#4B2E83; font-size:14px; line-height:20px;">
            Já se passaram <b>{dias} dias</b> sem confirmação da resolução da poda, configurando descumprimento do acordo de manutenção e zeladoria.
          </td>
        </tr>
      </table>
      <p style="color:#333333; font-size:14px; line-height:20px; margin:0 0 24px 0;">
        Diante do exposto, solicitamos providências imediatas para a solução da ocorrência. A manutenção da pendência configura mora contratual definitiva.
      </p>
    </td>
  </tr>
  <tr>
    <td bgcolor="#4B2E83" align="center" style="padding:14px;">
      <span style="color:#ffffff; font-size:16px;">&#10022;</span>
    </td>
  </tr>
  <tr>
    <td style="padding:16px 30px; color:#999999; font-size:10px; line-height:14px;">
      Esta é uma notificação automática referente ao acordo de manutenção e zeladoria. Em caso de dúvidas, entre em contato com a equipe responsável.
    </td>
  </tr>
</table>
</body>
</html>
"""


def registrar_historico(registros):
    novo_df = pd.DataFrame(
        registros,
        columns=["UFSIGLA", "DETENTORA", "ID Detentora", "EMAIL", "DATA_ENVIO", "CORRIGIDO", "DATA_COBRANCA"],
    )
    if os.path.exists(HISTORICO_PATH):
        existente = pd.read_excel(HISTORICO_PATH)
        historico = pd.concat([existente, novo_df], ignore_index=True)
    else:
        historico = novo_df
    historico.to_excel(HISTORICO_PATH, index=False)


def montar_rascunho(outlook, emails, ufsigla, detentora, id_detentora, caminhos_imagens):
    mail = outlook.CreateItem(0)
    mail.CC = 'jobson.silva@telefonica.com'
    mail.To = ";".join(emails)
    mail.Subject = f"Solicitação de Poda - Site {ufsigla}"
    mail.HTMLBody = HTML_TEMPLATE.format(
        ufsigla=ufsigla, detentora=detentora, id_detentora=id_detentora
    )
    for caminho_imagem in caminhos_imagens:
        mail.Attachments.Add(caminho_imagem)
    mail.Display()


def montar_cobranca(outlook, emails, ufsigla, id_detentora, data_envio, dias, caminhos_imagens):
    mail = outlook.CreateItem(0)
    mail.CC = 'jobson.silva@telefonica.com'
    mail.To = ";".join(emails)
    mail.Subject = f"[COBRANÇA] Solicitação de Poda - Site {ufsigla}"
    mail.HTMLBody = COBRANCA_HTML_TEMPLATE.format(
        ufsigla=ufsigla, id_detentora=id_detentora, data_envio=data_envio, dias=dias
    )
    for caminho_imagem in caminhos_imagens:
        mail.Attachments.Add(caminho_imagem)
    mail.Display()


def verificar_cobrancas(outlook, base_detentoras, grupos_imagens):
    """Reabre notificação (modo cobrança) para envios com mais de DIAS_LIMITE_COBRANCA dias e ainda não marcados como CORRIGIDO."""
    if not os.path.exists(HISTORICO_PATH):
        return

    historico = pd.read_excel(HISTORICO_PATH)
    for coluna, padrao in (("CORRIGIDO", "NAO"), ("DATA_COBRANCA", "")):
        if coluna not in historico.columns:
            historico[coluna] = padrao
        historico[coluna] = historico[coluna].astype(object)

    historico["_data_envio_dt"] = pd.to_datetime(historico["DATA_ENVIO"], dayfirst=True, errors="coerce")
    limite = datetime.now() - timedelta(days=DIAS_LIMITE_COBRANCA)

    cobrancas_enviadas = 0
    for indice, linha in historico.iterrows():
        if esta_corrigido(linha["CORRIGIDO"]):
            continue
        if pd.notna(linha["DATA_COBRANCA"]) and str(linha["DATA_COBRANCA"]).strip():
            continue
        data_envio_dt = linha["_data_envio_dt"]
        if pd.isna(data_envio_dt) or data_envio_dt > limite:
            continue

        ufsigla = linha["UFSIGLA"]
        detentora = linha["DETENTORA"]
        emails = [e.strip() for e in str(linha["EMAIL"]).split(";") if e.strip()]
        if not emails:
            continue

        ufsigla_norm = normalizar(ufsigla)
        id_detentora = base_detentoras.loc[ufsigla_norm, "ID Detentora"] if ufsigla_norm in base_detentoras.index else ""
        dias_passados = (datetime.now() - data_envio_dt).days
        caminhos_imagens = grupos_imagens.get(ufsigla, [])

        montar_cobranca(
            outlook, emails, ufsigla, id_detentora,
            data_envio_dt.strftime("%d/%m/%Y"), dias_passados, caminhos_imagens,
        )
        print(f"[COBRANÇA] {ufsigla}: rascunho de cobrança aberto para {detentora} ({', '.join(emails)}) - {dias_passados} dia(s) sem correção")

        historico.at[indice, "DATA_COBRANCA"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        cobrancas_enviadas += 1

    if cobrancas_enviadas:
        historico.drop(columns=["_data_envio_dt"]).to_excel(HISTORICO_PATH, index=False)
        print(f"Histórico atualizado com {cobrancas_enviadas} cobrança(s) registrada(s).")


def main():
    base_detentoras = carregar_base_detentoras()
    emails_detentoras = carregar_emails_detentoras()
    grupos_imagens = listar_grupos_imagens()

    if not grupos_imagens:
        print(f"Nenhuma pasta de imagens encontrada em '{IMAGENS_DIR}'.")
        return

    outlook = win32com.client.Dispatch("Outlook.Application")

    processados = 0
    pulados = 0
    registros_historico = []

    for ufsigla, caminhos_imagens in grupos_imagens.items():
        ufsigla_norm = normalizar(ufsigla)

        if ufsigla_norm not in base_detentoras.index:
            print(f"[SKIP] {ufsigla}: UFSigla não encontrado em BA/SE/AL")
            pulados += 1
            continue

        detentora = base_detentoras.loc[ufsigla_norm, "detentora_resolvida"]
        detentora_norm = normalizar(detentora)
        id_detentora = base_detentoras.loc[ufsigla_norm, "ID Detentora"]

        if not detentora_norm:
            print(f"[SKIP] {ufsigla}: sem detentora associada")
            pulados += 1
            continue

        emails = emails_detentoras.get(detentora_norm)
        if not emails:
            print(f"[SKIP] {ufsigla}: detentora '{detentora}' sem email cadastrado")
            pulados += 1
            continue

        montar_rascunho(outlook, emails, ufsigla, detentora, id_detentora, caminhos_imagens)
        print(f"[OK] {ufsigla}: rascunho aberto para {detentora} ({', '.join(emails)}) com {len(caminhos_imagens)} imagem(ns)")
        processados += 1
        registros_historico.append({
            "UFSIGLA": ufsigla,
            "DETENTORA": detentora,
            "ID Detentora":id_detentora,
            "EMAIL": "; ".join(emails),
            "DATA_ENVIO": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "  GIDO": "NAO",
            "DATA_COBRANCA": "",
        })

    if registros_historico:
        registrar_historico(registros_historico)
        print(f"Histórico atualizado em '{HISTORICO_PATH}' ({len(registros_historico)} registro(s)).")

    print(f"\nResumo: {processados} rascunho(s) aberto(s), {pulados} pulado(s).")

    verificar_cobrancas(outlook, base_detentoras, grupos_imagens)


if __name__ == "__main__":
    main()

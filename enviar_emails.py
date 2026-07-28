import os

import pandas as pd
import win32com.client

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGENS_DIR = os.path.join(BASE_DIR, "imagens")
BASE_DETENTORAS_PATH = os.path.join(BASE_DIR, "base_detentoras.xlsx")
DETENTORAS_PATH = os.path.join(BASE_DIR, "detentoras.xlsx")
ESTADOS_FILTRO = {"BA", "SE", "AL"}
EXTENSOES_IMAGEM = {".png", ".jpg", ".jpeg"}


def normalizar(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper()


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


def montar_rascunho(outlook, emails, ufsigla, detentora, id_detentora, caminhos_imagens):
    mail = outlook.CreateItem(0)
    mail.To = ";".join(emails)
    mail.Subject = f"Documentação - Site {ufsigla}"
    mail.Body = (
        f"Prezados,\n\n"
        f"Segue(m) em anexo a(s) imagem(ns) referente(s) ao site {ufsigla} (detentora {detentora}, ID Detentora {id_detentora}).\n\n"
        f"Atenciosamente."
    )
    for caminho_imagem in caminhos_imagens:
        mail.Attachments.Add(caminho_imagem)
    mail.Display()


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

    print(f"\nResumo: {processados} rascunho(s) aberto(s), {pulados} pulado(s).")


if __name__ == "__main__":
    main()

# Automação de Envio de Emails

Script em Python que monta rascunhos de email no Outlook para notificar as detentoras de sites (torres/estruturas) sobre vegetação alta identificada, anexando as fotos de evidência de cada site.

## Como funciona

1. Lê a planilha `base_detentoras.xlsx` e filtra os sites localizados em **BA**, **SE** e **AL**, identificando a detentora responsável por cada `UFSigla`.
2. Lê a planilha `detentoras.xlsx` para obter os emails de contato de cada detentora.
3. Varre a pasta `imagens/`: cada subpasta deve se chamar exatamente como o `UFSigla` do site e conter as fotos de evidência.
4. Para cada subpasta encontrada, cruza o `UFSigla` com a base de detentoras e monta um rascunho de email (via Outlook), anexando as fotos e preenchendo o corpo do email com o nome da detentora e o `ID Detentora`.
5. Os rascunhos são apenas **exibidos** (`mail.Display()`) — nenhum email é enviado automaticamente, permitindo revisão antes do envio.
6. Ao final, verifica o `historico_envios.xlsx`: para todo envio com mais de 10 dias (`DIAS_LIMITE_COBRANCA`) e ainda não marcado como `CORRIGIDO`, abre um rascunho de **cobrança** (mesmo layout, com aviso de pendência) reforçando a solicitação.

## Estrutura de pastas

```
imagens/
  BASMF/
    foto1.jpg
    foto2.png
  BAJCE/
    evidencia.jpg
```

- O nome da subpasta deve corresponder ao `UFSigla` cadastrado em `base_detentoras.xlsx`.
- Dentro da subpasta, os arquivos podem ter qualquer nome, desde que a extensão seja `.png`, `.jpg` ou `.jpeg`.
- Arquivos soltos diretamente em `imagens/` (fora de subpastas) são ignorados.

## Pré-requisitos

- Windows com **Microsoft Outlook** instalado e configurado (o script usa automação COM via `win32com`).
- Python 3.9+
- Dependências:

```bash
pip install pandas pywin32 openpyxl
```

## Arquivos de entrada

| Arquivo | Aba | Colunas relevantes |
|---|---|---|
| `base_detentoras.xlsx` | `Dados` | `UFSigla`, `State / Province`, `Detentora (Original)`, `Detentora (Referência)`, `ID Detentora`, `Updated` |
| `detentoras.xlsx` | `Plan1` | `DETENTORA`, `EMAIL` (múltiplos emails separados por `;`) |

Quando `Detentora (Original)` está vazia, o script usa `Detentora (Referência)` como detentora resolvida do site. Em caso de duplicidade de `UFSigla`, é mantido o registro com o `Updated` mais recente.

## Uso

```bash
python enviar_emails.py
```

Ao final da execução, o script imprime um resumo com a quantidade de rascunhos abertos e de sites pulados (por falta de detentora, email cadastrado ou por não pertencerem a BA/SE/AL), indicando o motivo de cada ocorrência.

## Histórico de envios (`historico_envios.xlsx`)

| Coluna | Descrição |
|---|---|
| `UFSIGLA`, `DETENTORA`, `EMAIL`, `DATA_ENVIO` | Registro do envio original. |
| `CORRIGIDO` | Preencher manualmente com `SIM` (ou `S`/`X`/`OK`) quando a poda for confirmada. Enquanto vazio ou `NAO`, o site é considerado pendente. |
| `DATA_COBRANCA` | Preenchida automaticamente pelo script quando um email de cobrança é aberto para aquele registro (evita reenvio duplicado). |

## Cobrança automática de pendências

A cada execução, o script verifica linhas do histórico com `DATA_ENVIO` há mais de 10 dias e `CORRIGIDO` ainda não marcado. Para cada uma, abre um novo rascunho (mesmo template visual, agora como notificação de cobrança) reforçando a pendência — reaproveitando as imagens da pasta `imagens/<UFSigla>` se ainda existirem. Após aberto, a coluna `DATA_COBRANCA` é preenchida para não gerar cobrança repetida no mesmo registro.

## Observações

- O corpo do email é HTML (`mail.HTMLBody`), estilizado com cabeçalho e identidade visual próprios — ver `HTML_TEMPLATE` e `COBRANCA_HTML_TEMPLATE` em [enviar_emails.py](enviar_emails.py).
- `base_detentoras.xlsx` e `detentoras.xlsx` contêm dados sensíveis/internos e não devem ser versionados publicamente.

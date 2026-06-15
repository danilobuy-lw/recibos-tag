# Gerador de Recibos de Pagamento Tag — Web App (Streamlit)

App web onde qualquer pessoa faz upload da planilha de histórico de
passagens e baixa os PDFs de "Confirmação de pagamento" preenchidos.

## Como usar (usuário final)

1. Abra a URL do app no navegador.
2. Arraste a planilha `.xlsx` para a área de upload (ou clique em "Browse files").
3. Marque/desmarque "apenas Status Pagamento = Pago" conforme precisar.
4. Clique em **Gerar Recibos**.
5. Baixe o `.zip` com todos os PDFs (ou os PDFs individuais).

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

A conversão para PDF usa o **LibreOffice** (precisa estar instalado).
- macOS: `brew install --cask libreoffice`
- Ubuntu/Debian: `sudo apt install libreoffice`
- Windows: baixe em https://www.libreoffice.org/

Sem o LibreOffice, o app gera apenas arquivos `.docx`.

## Publicar no Streamlit Community Cloud (gratuito)

Pré-requisitos:
- Conta no [GitHub](https://github.com/) (gratuita).
- Conta no [Streamlit Cloud](https://share.streamlit.io/) (gratuita, login com GitHub).

Passos:

1. **Crie um repositório no GitHub** (pode ser público ou privado).
2. **Suba todos os arquivos** deste projeto para o repositório:
   - `streamlit_app.py`
   - `template.docx`
   - `requirements.txt`
   - `packages.txt`  ← importante, instala o LibreOffice no servidor
   - `.streamlit/config.toml`
   - `README.md`
3. **No Streamlit Cloud**, clique em "**New app**" → selecione o repositório,
   o branch (`main`) e o arquivo principal (`streamlit_app.py`).
4. Clique em **Deploy**. Em 2-5 minutos a URL pública estará no ar.

A URL fica no formato:
`https://<seu-usuario>-<nome-do-app>.streamlit.app`

### Atualizar o app

Qualquer commit no GitHub faz o app refazer o deploy automaticamente.

### Restringir acesso (opcional)

Por padrão a URL é pública. Se quiser proteger por senha:

1. No Streamlit Cloud, vá em **Settings → Secrets** do seu app.
2. Adicione:
   ```toml
   APP_PASSWORD = "sua-senha-aqui"
   ```
3. Adicione no início do `streamlit_app.py`:
   ```python
   import streamlit as st
   pwd = st.text_input("Senha", type="password")
   if pwd != st.secrets.get("APP_PASSWORD"):
       st.stop()
   ```

## Estrutura dos arquivos

```
recibos_streamlit/
├── streamlit_app.py        # Código do app
├── template.docx           # Template Word (Confirmação de pagamento Tag)
├── requirements.txt        # Dependências Python
├── packages.txt            # Pacotes apt (LibreOffice) p/ Streamlit Cloud
├── .streamlit/
│   └── config.toml         # Tema e config do servidor
└── README.md               # Este arquivo
```

## Mapeamento de campos

| Coluna da planilha (aceita variações) | Campo no template Word |
|---------------------------------------|------------------------|
| ID_CONCESSIONÁRIA / ID                | ID                     |
| Placa                                 | Placa                  |
| Data da Passagem                      | Data (passagem)        |
| Praça                                 | Concessionária         |
| Rodovia                               | Pórtico                |
| Valor                                 | Valor (passagem)       |
| Data Pagamento / Data de pagamento    | Data (pagamento)       |
| Taxa pedágio                          | Tarifa de pedágio + Valor (pagamento) |

Campos **Multa / Juros** e **Taxa adm** ficam em `R$ 0,00`.

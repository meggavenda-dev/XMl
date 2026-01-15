
# Leitor de XML TISS (Consulta e SP‑SADT)

Extrai **nº do lote**, **quantidade de guias** e **valor total** por arquivo TISS, tanto para **Consulta** (soma `valorProcedimento`) quanto para **SP‑SADT** (usa `valorTotalGeral` e, se faltar, reconstrói somando itens/“outrasDespesas”).

## Rodar Local
```bash
python -m venv .venv
source .venv/bin/activate        # (Windows: .venv\Scripts\activate)
pip install -r requirements.txt
streamlit run app.py

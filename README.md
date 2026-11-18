# Laudos e Conformidade Ambiental · IST Blumenau (CDIA 2026/2) · v4
Protótipo de front-end 100% Python (Streamlit + Pandas + Plotly + OpenPyXL). Dados mocados lidos direto das planilhas UltraLIMS; nada é gravado.

## Rodar
```
python -m venv .venv
.venv\Scripts\activate          # Linux/Mac: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```
## Estrutura
- `app.py`: 5 telas (Dashboard, Importação, Conferência, Parecer, Análise temporal) espelhando o protótipo HTML
- `motor.py`: leitura/avaliação das abas UltraLIMS (despivotamento, isolamento, escopo, vigência, unidades, parecer)
- `estilo.py`: CSS e componentes visuais
- `dados/`: 3 propostas mocadas (`*_UltraLIMS.xlsx`) + `referencia/01_modelo_fisico.sql` (usado só como referência do De/Para)

## Trocar de proposta
Seletor "Proposta ..." na barra superior. Para novas planilhas: tela Importação > Upload (mesmo padrão UltraLIMS).
Overrides, textos e auditoria ficam em `st.session_state` (somem ao recarregar).

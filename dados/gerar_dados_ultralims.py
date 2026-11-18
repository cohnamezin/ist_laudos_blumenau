# -*- coding: utf-8 -*-
"""
Gera 'dados_ensaios_ultralims_teste.xlsx' com dados simulados do UltraLIMS
(IST Blumenau) para testar as 5 telas e as regras RN-01 a RN-12.

Abas:
  1. Dados_Ensaios   -> schema exato de 12 colunas (entrada do Wizard de importação)
  2. Cadastro_VMP    -> regras/limites simulados (inclui duplicidade RN-04 e versões RN-08)
  3. Gabarito        -> resultado esperado de cada caso, para validar o motor de regras

ATENÇÃO: todos os limites (VMP) são SIMULADOS para teste. Os casos 1 a 7 seguem os
valores do prompt; os casos extras (007 a 015) usam limites fictícios.
"""
from datetime import date

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ARQUIVO = "dados_ensaios_ultralims_teste.xlsx"

# ----------------------------------------------------------------------
# Aba 1 - Dados_Ensaios  (12 colunas exatas)
# ----------------------------------------------------------------------
COLUNAS = [
    "ID_Amostra", "Cliente", "Ponto_Coleta", "Matriz_Agua", "Data_Coleta",
    "Cod_Ensaio", "Parametro_Analisado", "Celula_Laboratorial",
    "Resultado_Medido", "Unidade_Medida", "Legislacao_Aplicada",
    "Classe_Legislacao",
]

GMMS = "Portaria GM/MS 888/2021"
SES = "Portaria SES 1468/2022"
CONAMA = "CONAMA 430/2011"

d = date.fromisoformat

dados = [
    # CASO 1 - ATENDE
    ["AMO-2026-001", "Laticínios Serra SC", "Torneira da cozinha industrial", "AT", d("2026-09-03"), 36, "Odor", "FQ", 1.0, "Intensidade", GMMS, "Água Tratada"],
    ["AMO-2026-001", "Laticínios Serra SC", "Torneira da cozinha industrial", "AT", d("2026-09-03"), 62, "Coliformes Totais", "MIC", 0, "UFC/100mL", GMMS, "Água Tratada"],
    # CASO 2 - NÃO ATENDE (DBO)
    ["AMO-2026-002", "Indústria Têxtil Vale do Itajaí", "Saída da ETE sanitária", "Efluente Sanitário", d("2026-09-08"), 92, "DBO 5 dias", "FQ", 145.0, "mg/L", CONAMA, "Sanitário"],
    ["AMO-2026-002", "Indústria Têxtil Vale do Itajaí", "Saída da ETE sanitária", "Efluente Sanitário", d("2026-09-08"), 106, "Nitrogênio Amoniacal", "FQ", 15.0, "mg/L", CONAMA, "Sanitário"],
    # CASO 3 - Heptacloro em ASU (limite zero)
    ["AMO-2026-003", "Piscicultura Vale Verde", "Captação no ribeirão - tanque 2", "ASU", d("2026-09-10"), 397, "Heptacloro", "FQ", 0.05, "µg/L", SES, "Sem classe"],
    # CASO 4 - Duplicidade legada (RN-04)
    ["AMO-2026-004", "Frigorífico Central", "Bebedouro do refeitório", "AT", d("2026-09-12"), 7001, "1,2-Diclorobenzeno", "FQ", 0.0008, "mg/L", GMMS, "Sem classe"],
    # CASO 5 - Divergência de unidades (RN-11)
    ["AMO-2026-005", "Cooperativa Agrícola", "Açude da propriedade 04", "ASU", d("2026-09-17"), 877, "Glifosato+AMPA", "FQ", 0.05, "mg/L", SES, "Sem classe"],
    # CASO 7 - Operador qualitativo (RN-12)
    ["AMO-2026-006", "Mineração Vale Doce", "Poço tubular PT-02", "SUBT", d("2026-09-22"), 6486, "Carbendazim", "FQ", "< 0.20", "µg/L", SES, "Sem classe"],
    # EXTRA - Presente/Ausente conforme (RN-12)
    ["AMO-2026-007", "Laticínios Serra SC", "Torneira da cozinha industrial", "AT", d("2026-10-05"), 63, "Escherichia coli", "MIC", "Ausente", "NMP/100mL", GMMS, "Água Tratada"],
    ["AMO-2026-007", "Laticínios Serra SC", "Torneira da cozinha industrial", "AT", d("2026-10-05"), 70, "Cloro Residual Livre", "FQ", 0.8, "mg/L", GMMS, "Água Tratada"],
    # EXTRA - Efluente industrial: um item reprovado e um aprovado (RN-01/05/07)
    ["AMO-2026-008", "Galvânica Navegantes", "Saída da ETE industrial", "Efluente Industrial", d("2026-10-07"), 150, "Óleos e Graxas Minerais", "FQ", 35.0, "mg/L", CONAMA, "Sem classe"],
    ["AMO-2026-008", "Galvânica Navegantes", "Saída da ETE industrial", "Efluente Industrial", d("2026-10-07"), 200, "Chumbo Total", "ME", 0.3, "mg/L", CONAMA, "Sem classe"],
    # EXTRA - "Presente" reprova (RN-12)
    ["AMO-2026-009", "Frigorífico Central", "Reservatório elevado", "AT", d("2026-10-09"), 62, "Coliformes Totais", "MIC", "Presente", "UFC/100mL", GMMS, "Água Tratada"],
    # CASO 6 - Série temporal (RN-10)
    ["AMO-2026-010", "Metalúrgica Itajaí", "Saída ETE - Ponto 03", "Efluente Industrial", d("2026-09-01"), 106, "Nitrogênio Amoniacal", "FQ", 12.0, "mg/L", CONAMA, "Sem classe"],
    ["AMO-2026-011", "Metalúrgica Itajaí", "Saída ETE - Ponto 03", "Efluente Industrial", d("2026-09-15"), 106, "Nitrogênio Amoniacal", "FQ", 14.5, "mg/L", CONAMA, "Sem classe"],
    ["AMO-2026-012", "Metalúrgica Itajaí", "Saída ETE - Ponto 03", "Efluente Industrial", d("2026-10-01"), 106, "Nitrogênio Amoniacal", "FQ", 16.8, "mg/L", CONAMA, "Sem classe"],
    ["AMO-2026-013", "Metalúrgica Itajaí", "Saída ETE - Ponto 03", "Efluente Industrial", d("2026-10-15"), 106, "Nitrogênio Amoniacal", "FQ", 21.5, "mg/L", CONAMA, "Sem classe"],
    # EXTRA - Mesma amostra, 2 legislações (RN-02)
    ["AMO-2026-014", "Condomínio Residencial Bela Vista", "Reservatório inferior", "AT", d("2026-10-20"), 70, "Cloro Residual Livre", "FQ", 2.4, "mg/L", GMMS, "Água Tratada"],
    ["AMO-2026-014", "Condomínio Residencial Bela Vista", "Reservatório inferior", "AT", d("2026-10-20"), 70, "Cloro Residual Livre", "FQ", 2.4, "mg/L", SES, "Água Tratada"],
    # EXTRA - Vigência histórica (RN-08)
    ["AMO-2026-015", "Indústria Têxtil Vale do Itajaí", "Saída da ETE sanitária", "Efluente Sanitário", d("2025-03-10"), 92, "DBO 5 dias", "FQ", 130.0, "mg/L", CONAMA, "Sanitário"],
]
df_dados = pd.DataFrame(dados, columns=COLUNAS)

# ----------------------------------------------------------------------
# Aba 2 - Cadastro_VMP  (limites simulados, inclui duplicidade e versões)
# ----------------------------------------------------------------------
COL_VMP = [
    "Cod_Ensaio", "Parametro", "Legislacao", "Classe", "VMP_Minimo",
    "VMP_Maximo", "Unidade_VMP", "Inicio_Vigencia", "Fim_Vigencia",
    "Status_Duplicidade_Alerta", "Observacao",
]
v = d("2020-01-01")
vmp = [
    [36, "Odor", GMMS, "Água Tratada", 0, 6, "Intensidade", v, None, "OK", "Caso 1"],
    [62, "Coliformes Totais", GMMS, "Água Tratada", 0, 0, "UFC/100mL", v, None, "OK", "Ausência exigida (RN-12: Presente > 0)"],
    [63, "Escherichia coli", GMMS, "Água Tratada", 0, 0, "NMP/100mL", v, None, "OK", "Ausente = 0 (RN-12)"],
    [70, "Cloro Residual Livre", GMMS, "Água Tratada", 0.2, 5.0, "mg/L", v, None, "OK", "Limite simulado"],
    [70, "Cloro Residual Livre", SES, "Água Tratada", 0.2, 2.0, "mg/L", v, None, "OK", "Limite FICTÍCIO para testar RN-02"],
    [92, "DBO 5 dias", CONAMA, "Sanitário", 0, 150.0, "mg/L", d("2011-05-16"), d("2025-12-31"), "OK", "Versão antiga (FICTÍCIA) - RN-08"],
    [92, "DBO 5 dias", CONAMA, "Sanitário", 0, 120.0, "mg/L", d("2026-01-01"), None, "OK", "Versão vigente - Caso 2"],
    [106, "Nitrogênio Amoniacal", CONAMA, "Sanitário", 0, 20.0, "mg/L", v, None, "OK", "Caso 2"],
    [106, "Nitrogênio Amoniacal", CONAMA, "Sem classe", 0, 20.0, "mg/L", v, None, "OK", "Caso 6 (alerta amarelo a partir de 16.0)"],
    [150, "Óleos e Graxas Minerais", CONAMA, "Sem classe", 0, 20.0, "mg/L", v, None, "OK", "Limite simulado"],
    [200, "Chumbo Total", CONAMA, "Sem classe", 0, 0.5, "mg/L", v, None, "OK", "Limite simulado"],
    [397, "Heptacloro", SES, "Sem classe", 0, 0.000000, "µg/L", v, None, "OK", "Caso 3 - ASU (limite zero); SUBT seria 0.3"],
    [877, "Glifosato+AMPA", SES, "Sem classe", 0, 30.0, "µg/L", v, None, "OK", "Caso 5 - norma em µg/L"],
    [6486, "Carbendazim", SES, "Sem classe", 0, 19.5, "µg/L", v, None, "OK", "Caso 7"],
    [7001, "1,2-Diclorobenzeno", GMMS, "Sem classe", 0, 0.001000, "mg/L", v, None, "DUPLICADO", "Caso 4 - cadastro legado A"],
    [7001, "1,2-Diclorobenzeno", GMMS, "Sem classe", 0, 0.000000, "mg/L", v, None, "DUPLICADO", "Caso 4 - cadastro legado B (mais restritivo)"],
]
df_vmp = pd.DataFrame(vmp, columns=COL_VMP)

# ----------------------------------------------------------------------
# Aba 3 - Gabarito (resultado esperado por amostra)
# ----------------------------------------------------------------------
COL_GAB = ["ID_Amostra", "Parecer_Esperado", "Regras_Testadas", "Resultado_Esperado_Detalhado"]
gab = [
    ["AMO-2026-001", "ATENDE", "RN-01, RN-07", "Odor 1.0 (0-6) e Coliformes 0 (máx 0): todos conformes."],
    ["AMO-2026-002", "NÃO ATENDE", "RN-01, RN-05, RN-07", "DBO 145.0 > 120.0 mg/L (vermelho). Nitrogênio 15.0 <= 20.0 (conforme)."],
    ["AMO-2026-003", "NÃO ATENDE", "RN-01, RN-05", "Heptacloro 0.05 µg/L > VMP 0.000000 µg/L (limite zero violado)."],
    ["AMO-2026-004", "NÃO ATENDE", "RN-04, RN-01", "Duplicidade: adotar VMP 0.000000 (mais restritivo). 0.0008 > 0 -> reprova + alerta de auditoria."],
    ["AMO-2026-005", "NÃO ATENDE", "RN-11, RN-01", "0.05 mg/L = 50.0 µg/L > 30.0 µg/L."],
    ["AMO-2026-006", "ATENDE", "RN-12, RN-01", "'< 0.20' µg/L interpretado como menor que o LQ, abaixo de 19.5 µg/L -> conforme."],
    ["AMO-2026-007", "ATENDE", "RN-12, RN-01", "E. coli 'Ausente' = 0; Cloro 0.8 dentro de 0.2-5.0."],
    ["AMO-2026-008", "NÃO ATENDE", "RN-01, RN-05, RN-07", "Óleos e Graxas 35.0 > 20.0 (vermelho); Chumbo 0.3 <= 0.5 (conforme)."],
    ["AMO-2026-009", "NÃO ATENDE", "RN-12, RN-01", "Coliformes Totais 'Presente' viola ausência exigida."],
    ["AMO-2026-010", "ATENDE", "RN-01, RN-10", "12.0 mg/L = 60% do VMP (sem alerta)."],
    ["AMO-2026-011", "ATENDE", "RN-01, RN-10", "14.5 mg/L = 72.5% do VMP (sem alerta)."],
    ["AMO-2026-012", "ATENDE (com ALERTA AMARELO)", "RN-10", "16.8 mg/L = 84% do VMP (>= 80%) e 3ª medição consecutiva em alta."],
    ["AMO-2026-013", "NÃO ATENDE", "RN-01, RN-05, RN-10", "21.5 mg/L = 107.5% do VMP (estouro vermelho); 4ª elevação consecutiva."],
    ["AMO-2026-014", "NÃO ATENDE (na 2ª legislação)", "RN-02, RN-07", "Cloro 2.4: conforme na GM/MS 888 (0.2-5.0), não conforme na SES 1468 (máx 2.0 fictício) -> amostra reprovada."],
    ["AMO-2026-015", "ATENDE", "RN-08, RN-01", "Coleta em 2025-03-10 usa versão antiga (máx 150.0): 130.0 conforme. Com a regra atual (120.0) reprovaria."],
]
df_gab = pd.DataFrame(gab, columns=COL_GAB)

# ----------------------------------------------------------------------
# Gravação + estilos
# ----------------------------------------------------------------------
with pd.ExcelWriter(ARQUIVO, engine="openpyxl") as xw:
    df_dados.to_excel(xw, sheet_name="Dados_Ensaios", index=False)
    df_vmp.to_excel(xw, sheet_name="Cadastro_VMP", index=False)
    df_gab.to_excel(xw, sheet_name="Gabarito", index=False)

wb = load_workbook(ARQUIVO)
azul = PatternFill("solid", fgColor="1E3A8A")
zebra = PatternFill("solid", fgColor="EFF6FF")
borda = Border(*(Side(style="thin", color="E2E8F0"),) * 4)
f_head = Font(name="Arial", bold=True, color="FFFFFF", size=10)
f_body = Font(name="Arial", size=10)

for ws in wb.worksheets:
    for cell in ws[1]:
        cell.fill, cell.font, cell.border = azul, f_head, borda
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(2, ws.max_row + 1):
        for cell in ws[r]:
            cell.font, cell.border = f_body, borda
            cell.alignment = Alignment(vertical="center", wrap_text=(ws.title == "Gabarito"))
            if r % 2 == 1:  # linhas zebradas
                cell.fill = zebra
    for i, col in enumerate(ws.columns, start=1):
        maior = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(i)].width = min(max(maior + 3, 12), 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 26

# formatos de data
ws = wb["Dados_Ensaios"]
for r in range(2, ws.max_row + 1):
    ws.cell(r, 5).number_format = "yyyy-mm-dd"
    ws.cell(r, 9).alignment = Alignment(horizontal="right", vertical="center")
ws = wb["Cadastro_VMP"]
for r in range(2, ws.max_row + 1):
    ws.cell(r, 8).number_format = "yyyy-mm-dd"
    ws.cell(r, 9).number_format = "yyyy-mm-dd"
    ws.cell(r, 5).number_format = "0.000000"
    ws.cell(r, 6).number_format = "0.000000"
wb["Gabarito"].column_dimensions["D"].width = 80

wb.save(ARQUIVO)
print(f"Arquivo gerado: {ARQUIVO}")

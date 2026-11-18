"""Motor v4 · leitura direta das planilhas UltraLIMS (dados mocados) · IST Blumenau.

Python puro + Pandas. Sem banco, sem fila, sem API: tudo é lido e avaliado em memória.
"""
import datetime as dt
import re

import numpy as np
import pandas as pd

# ------------------------------------------------------------------ constantes
ABAS_OBRIG = ["Worksheet", "Amostra_Metadados", "Proposta_Comercial", "Catalogo_Ensaios", "Limites_VMP"]
ABAS_OPC = ["Gabarito_Itens", "Gabarito_Amostras"]
COLS_OBRIG = {
    "Worksheet": ["Código Amostra", "Data Coleta"],
    "Amostra_Metadados": ["codigo_amostra", "numero_proposta", "descricao_do_ponto", "origem_amostra", "status_amostra"],
    "Proposta_Comercial": ["numero_proposta", "razao_social"],
    "Catalogo_Ensaios": ["codigo_ensaio", "nome_ensaio", "sigla_celula", "unidade_trabalho"],
    "Limites_VMP": ["codigo_ensaio", "nome_legislacao", "nome_classe", "valor_min_permitido", "valor_max_permitido",
                    "sigla_unidade"],
}
AVALIADOS = ["Conforme", "Não Conforme", "Fora de Escopo", "Bloqueado / Incompatível", "Inconclusivo", "Revisão Manual"]
PENDENTES = ["Bloqueado / Incompatível", "Inconclusivo", "Revisão Manual"]
PAD_COL_ENSAIO = re.compile(r"^\s*(\d+)\s*-\s*(.*?)\s*$")
PAD_ADMIN = re.compile(r"controle de qualidade|carta controle|descarte|r\$|\bcq\b", re.I)
PAD_VAZIO = re.compile(r"^[\s_\-–—.]*$")
MASSA = {"g/L": 1e6, "mg/L": 1e3, "µg/L": 1.0, "ug/L": 1.0, "ng/L": 1e-3}   # base: µg/L
INICIO, FIM = pd.Timestamp("1900-01-01"), pd.Timestamp("2200-01-01")

# Contextualização com o modelo físico (01_modelo_fisico.sql): origem na planilha -> destino no modelo
DE_PARA = [
    ("Worksheet", "Código Amostra", "amostra", "codigo_amostra"),
    ("Worksheet", "Código Amostra Externo", "amostra", "— (referência externa)"),
    ("Worksheet", "Data Coleta", "amostra", "data_coleta"),
    ("Worksheet", "Data Conclusão", "amostra", "data_conclusao_amostra"),
    ("Worksheet", "<código> - <nome>", "ensaio_resultado", "valor_numerico / valor_booleano"),
    ("Amostra_Metadados", "numero_proposta", "proposta_comercial", "numero_proposta"),
    ("Amostra_Metadados", "descricao_do_ponto", "amostra", "descricao_do_ponto"),
    ("Amostra_Metadados", "origem_amostra", "amostra", "origem_amostra"),
    ("Amostra_Metadados", "data_recebimento", "amostra", "data_recebimento"),
    ("Amostra_Metadados", "data_liberacao", "amostra", "data_liberacao"),
    ("Amostra_Metadados", "status_amostra", "amostra", "— (filtro de liberação)"),
    ("Proposta_Comercial", "razao_social", "cliente", "razao_social"),
    ("Proposta_Comercial", "cnpj_cpf", "cliente", "cnpj_cpf"),
    ("Proposta_Comercial", "frequencia_atendimento", "proposta_comercial", "frequencia_atendimento"),
    ("Proposta_Comercial", "data_inicio", "proposta_comercial", "data_inicio"),
    ("Proposta_Comercial", "data_fim", "proposta_comercial", "data_fim"),
    ("Catalogo_Ensaios", "codigo_ensaio", "ensaio", "codigo_ensaio"),
    ("Catalogo_Ensaios", "nome_ensaio", "ensaio", "nome_ensaio"),
    ("Catalogo_Ensaios", "sigla_celula", "laboratorio", "sigla_celula"),
    ("Catalogo_Ensaios", "unidade_trabalho", "unidade_medida", "sigla"),
    ("Limites_VMP", "nome_legislacao", "legislacao", "nome_legislacao"),
    ("Limites_VMP", "nome_classe", "classe_legislacao", "nome_classe"),
    ("Limites_VMP", "valor_min_permitido", "item_legislacao", "valor_min_permitido"),
    ("Limites_VMP", "valor_max_permitido", "item_legislacao", "valor_max_permitido"),
    ("Limites_VMP", "sigla_unidade", "unidade_medida", "sigla"),
    ("Limites_VMP", "data_inicio_vigencia", "legislacao", "data_inicio_vigencia"),
    ("Limites_VMP", "data_fim_vigencia", "legislacao", "data_fim_vigencia"),
    ("Limites_VMP", "no_escopo_proposta", "item_proposta_comercial", "vínculo item × proposta"),
]


# ------------------------------------------------------------------ leitura e conversões básicas
def ler_planilha(fonte):
    """Lê todas as abas preservando os tipos originais das células (sem inferência de datas/números)."""
    x = pd.read_excel(fonte, sheet_name=None, dtype=object)
    for df in x.values():
        df.columns = [str(c).strip() for c in df.columns]
    return x


def _vazio(v):
    if v is None:
        return True
    if isinstance(v, float) and np.isnan(v):
        return True
    try:
        return bool(pd.isna(v)) if not isinstance(v, str) else v.strip() == ""
    except (TypeError, ValueError):
        return False


def _dt(v):
    if _vazio(v):
        return pd.NaT
    if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return pd.Timestamp(v)
    s = str(v).strip()
    for f in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return pd.to_datetime(s, format=f)
        except (ValueError, TypeError):
            continue
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def _num_br(s):
    """'1.100,00' -> 1100.0 ; '0,001' -> 0.001 ; '14.5' -> 14.5"""
    s = str(s).strip().replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return float(s)


def fmt_br(v, casas=6):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    s = f"{float(v):.{casas}g}"
    if "e" in s:
        s = f"{float(v):.8f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def interpretar(x):
    """Devolve (tipo, valor, operador, qualitativo). tipo: vazio | ok | invalido."""
    if _vazio(x):
        return "vazio", np.nan, "", ""
    if isinstance(x, (int, float, np.number)) and not isinstance(x, bool):
        return "ok", float(x), "", ""
    s = str(x).strip()
    if PAD_VAZIO.match(s):
        return "vazio", np.nan, "", ""
    if s.lower() == "ausente":
        return "ok", 0.0, "", "Ausente"
    if s.lower() == "presente":
        return "ok", 1.0, "", "Presente"
    mt = re.match(r"^([<>]=?)\s*(.+)$", s)
    try:
        if mt:
            return "ok", _num_br(mt.group(2)), mt.group(1)[0], ""
        return "ok", _num_br(s), "", ""
    except ValueError:
        return "invalido", np.nan, "", ""


def _un(u):
    return str(u).strip().replace("μ", "µ")


def converter(v, de, para):
    """Converte v de 'de' para 'para'. Retorna (valor|None, obs). Unidades de natureza diferente = incompatíveis."""
    de, para = _un(de), _un(para)
    if de == para:
        return v, ""
    if de in MASSA and para in MASSA:
        return v * MASSA[de] / MASSA[para], f"Convertido de {de} para {para}"
    return None, f"Unidade do ensaio ({de}) incompatível com a do limite ({para})"


def _prep_vmp(v):
    v = v.copy()
    v["cod"] = pd.to_numeric(v["codigo_ensaio"], errors="coerce")
    v = v.dropna(subset=["cod"]).copy()
    v["cod"] = v["cod"].astype(int)
    v["vmin"] = pd.to_numeric(v["valor_min_permitido"], errors="coerce")
    v["vmax"] = pd.to_numeric(v["valor_max_permitido"], errors="coerce")
    ini = pd.Series([_dt(a) for a in v.get("data_inicio_vigencia", [None] * len(v))], index=v.index, dtype="datetime64[ns]")
    fim = pd.Series([_dt(a) for a in v.get("data_fim_vigencia", [None] * len(v))], index=v.index, dtype="datetime64[ns]")
    v["ini"], v["fim"] = ini.fillna(INICIO), fim.fillna(FIM)
    esc = v["no_escopo_proposta"] if "no_escopo_proposta" in v else pd.Series("Sim", index=v.index)
    v["escopo"] = ~esc.astype(str).str.strip().str.lower().str.startswith("n")
    v["origem"] = (v["origem_aplicavel"] if "origem_aplicavel" in v else pd.Series("", index=v.index)).fillna("").astype(str).str.strip()
    v["unid"] = v["sigla_unidade"].map(_un)
    return v


# ------------------------------------------------------------------ validação
def validar(x):
    erros, av = [], []
    falt = [a for a in ABAS_OBRIG if a not in x]
    if falt:
        return [f"Abas ausentes: {', '.join(falt)}."], av
    for aba, cols in COLS_OBRIG.items():
        f = [c for c in cols if c not in x[aba].columns]
        if f:
            erros.append(f"Aba {aba}: colunas ausentes {f}.")
    if erros:
        return erros, av
    if not any(PAD_COL_ENSAIO.match(c) for c in x["Worksheet"].columns):
        erros.append("Aba Worksheet: nenhuma coluna de ensaio no padrão '<código> - <nome>'.")
    if len(x["Proposta_Comercial"]) == 0:
        erros.append("Aba Proposta_Comercial vazia.")
    elif len(x["Proposta_Comercial"]) > 1:
        av.append("Proposta_Comercial tem mais de uma linha; foi usada a primeira.")
    return erros, av


# ------------------------------------------------------------------ avaliação
def _avaliar(L, base):
    """L: DataFrame de limites candidatos (mesmo ensaio/origem/vigência/escopo). Retorna linhas de resultado."""
    out = []
    un_t = base["Unidade"]
    for (leg, cls), g in L.groupby(["nome_legislacao", "nome_classe"], sort=False):
        r = dict(base, Legislacao=leg, Classe=cls, Duplicidade=len(g) > 1)
        comp = [converter(1.0, u, un_t) for u in g["unid"]]
        if any(c[0] is None for c in comp):
            r.update(Status="Bloqueado / Incompatível", Obs=comp[[c[0] is None for c in comp].index(True)][1],
                     VMP_Min=np.nan, VMP_Max=np.nan)
            out.append(r)
            continue
        vmax = np.nanmin([a * c[0] for a, c in zip(g["vmax"], comp)]) if g["vmax"].notna().any() else np.nan
        vmin = np.nanmax([a * c[0] for a, c in zip(g["vmin"], comp)]) if g["vmin"].notna().any() else np.nan
        obs = []
        if any(c[0] != 1.0 for c in comp):
            obs.append(f"Limite convertido de {g['unid'].iloc[0]} para {un_t}.")
        if len(g) > 1:
            obs.append(f"{len(g)} limites cadastrados para o mesmo critério; adotado o mais restritivo.")
        v, op, tol = base["Valor"], base["Operador"], 1e-12
        if op == "<":
            if pd.notna(vmax) and v > vmax + tol:
                st_, o = "Inconclusivo", f"Limite de quantificação (< {fmt_br(v)}) acima do limite permitido: não é possível atestar conformidade."
            else:
                st_, o = "Conforme", "Resultado abaixo do limite de quantificação."
        elif op == ">":
            if pd.notna(vmax) and v >= vmax - tol:
                st_, o = "Não Conforme", "Resultado acima do limite permitido."
            else:
                st_, o = "Inconclusivo", "Resultado com operador '>' sem limite de comparação conclusivo."
        else:
            ok = (pd.isna(vmax) or v <= vmax + tol) and (pd.isna(vmin) or v >= vmin - tol)
            st_, o = ("Conforme" if ok else "Não Conforme"), ""
        if o:
            obs.append(o)
        pct = v / vmax if (op == "" and pd.notna(vmax) and vmax > 0 and (pd.isna(vmin) or vmin == 0)) else np.nan
        r.update(VMP_Min=vmin, VMP_Max=vmax, Status=st_, Obs=" ".join(obs), Pct_VMP=pct)
        out.append(r)
    return out


def processar(x):
    """Lê as abas UltraLIMS e devolve (itens, aud, amostras, info)."""
    ws = x["Worksheet"]
    meta = x["Amostra_Metadados"].copy()
    meta["codigo_amostra"] = meta["codigo_amostra"].astype(str).str.strip()
    meta = meta.drop_duplicates("codigo_amostra").set_index("codigo_amostra")
    info = x["Proposta_Comercial"].iloc[0].to_dict()
    num = str(info["numero_proposta"]).strip()
    cat = x["Catalogo_Ensaios"].copy()
    cat["cod"] = pd.to_numeric(cat["codigo_ensaio"], errors="coerce")
    cat = cat.dropna(subset=["cod"]).drop_duplicates("cod")
    cat.index = cat["cod"].astype(int)
    vmp = _prep_vmp(x["Limites_VMP"])
    colunas = [(c, int(m.group(1)), m.group(2)) for c in ws.columns if (m := PAD_COL_ENSAIO.match(c))]

    itens, aud, amostras = [], [], []

    def log(linha, am, problema, verif, status):
        aud.append(dict(Linha=linha, Amostra=am, Problema=problema, Verificação=verif, Status=status))

    for i, row in ws.reset_index(drop=True).iterrows():
        linha, am = i + 2, str(row["Código Amostra"]).strip()
        coleta, concl = _dt(row["Data Coleta"]), _dt(row.get("Data Conclusão"))
        m = meta.loc[am] if am in meta.index else None
        cliente = info.get("razao_social", "")
        base0 = dict(Proposta=num, ID_Amostra=am, Linha=linha, Cliente=cliente, Data_Coleta=coleta)
        celulas = [(c, cod, nome, row[c]) for c, cod, nome in colunas if not _vazio(row[c])]
        # --- nível amostra: isolamento da proposta e descarte de ruído
        motivo, tipo = None, None
        if m is None:
            motivo, tipo = "Amostra sem metadados em Amostra_Metadados", "Descartado"
        elif str(m["numero_proposta"]).strip() != num:
            motivo, tipo = f"Amostra pertence a outra proposta ({str(m['numero_proposta']).strip()})", "Isolado"
        elif str(m["status_amostra"]).strip() != "Liberada":
            motivo, tipo = f"Amostra não liberada (status: {str(m['status_amostra']).strip()})", "Descartado"
        elif pd.to_numeric(m.get("possui_amostra_cq", 1), errors="coerce") > 1:
            motivo, tipo = "Amostra de controle de qualidade (duplicata de CQ)", "Descartado"
        if motivo:
            log(linha, am, motivo, "Isolamento da proposta" if tipo == "Isolado" else "Situação da amostra",
                "Bloqueada" if tipo == "Isolado" else "Descartada")
            for c, cod, nome, val in celulas:
                itens.append(dict(base0, Cod_Ensaio=cod, Parametro=cat.loc[cod, "nome_ensaio"] if cod in cat.index else nome,
                                  Celula=cat.loc[cod, "sigla_celula"] if cod in cat.index else "", Resultado=str(val),
                                  Legislacao="-", Classe="", Status=tipo, Obs=motivo, Duplicidade=False, Valor=np.nan,
                                  Operador="", Unidade="", Pct_VMP=np.nan, VMP_Min=np.nan, VMP_Max=np.nan,
                                  Ponto_Coleta=str(m["descricao_do_ponto"]) if m is not None else "",
                                  Matriz_Agua=str(m["origem_amostra"]) if m is not None else ""))
            continue
        origem, ponto = str(m["origem_amostra"]).strip(), str(m["descricao_do_ponto"]).strip()
        amostras.append(dict(ID_Amostra=am, Cliente=cliente, Ponto_Coleta=ponto, Matriz_Agua=origem, Data_Coleta=coleta,
                             Data_Conclusao=concl, Data_Recebimento=_dt(m.get("data_recebimento")),
                             Data_Liberacao=_dt(m.get("data_liberacao")), Linha=linha))
        base0.update(Ponto_Coleta=ponto, Matriz_Agua=origem)
        # --- nível item
        for c, cod, nome, val in celulas:
            tp, v, op, qual = interpretar(val)
            no_cat = cod in cat.index
            B = dict(base0, Cod_Ensaio=cod, Parametro=cat.loc[cod, "nome_ensaio"] if no_cat else nome,
                     Celula=cat.loc[cod, "sigla_celula"] if no_cat else "", Resultado=str(val), Valor=v, Operador=op,
                     Unidade=_un(cat.loc[cod, "unidade_trabalho"]) if no_cat else "", Legislacao="-", Classe="",
                     VMP_Min=np.nan, VMP_Max=np.nan, Pct_VMP=np.nan, Duplicidade=False, Obs="")
            if tp != "invalido":
                B["Resultado"] = qual or (f"{op} {fmt_br(v)}" if op else fmt_br(v))
            if tp == "vazio":
                itens.append(dict(B, Status="Descartado", Obs="Resultado em branco (placeholder)"))
                log(linha, am, f"Resultado em branco em '{cod} - {nome}'", "Campo obrigatório", "Erro")
                continue
            if not no_cat:
                if PAD_ADMIN.search(nome) or PAD_ADMIN.search(str(val)):
                    itens.append(dict(B, Status="Descartado", Obs="Coluna administrativa / controle de qualidade"))
                    log(linha, am, f"Coluna administrativa '{cod} - {nome}' desconsiderada", "Catálogo de ensaios", "Descartada")
                else:
                    itens.append(dict(B, Status="Fora de Escopo", Obs="Parâmetro não cadastrado no catálogo de ensaios"))
                    log(linha, am, f"Parâmetro '{nome}' ({cod}) fora do escopo", "Catálogo de ensaios", "Informativo")
                continue
            if tp == "invalido":
                itens.append(dict(B, Status="Revisão Manual", Obs="Resultado não interpretável"))
                log(linha, am, f"Resultado não interpretável em {B['Parametro']}: '{val}'", "Leitura do resultado", "Revisar")
                continue
            L = vmp[vmp.cod == cod]
            Lo = L[L.origem == origem]
            if L.empty or Lo.empty or not Lo.escopo.any():
                why = ("Parâmetro não contratado na proposta" if (not Lo.empty) else
                       "Sem limite cadastrado para o parâmetro" if L.empty else f"Sem limite aplicável à origem '{origem}'")
                itens.append(dict(B, Status="Fora de Escopo", Obs=why))
                log(linha, am, f"{B['Parametro']}: {why.lower()}", "Escopo contratado", "Informativo")
                continue
            Lv = Lo[Lo.escopo & (Lo.ini <= coleta.normalize()) & (Lo.fim >= coleta.normalize())]
            if Lv.empty:
                itens.append(dict(B, Status="Revisão Manual", Obs="Sem limite vigente na data da coleta"))
                log(linha, am, f"{B['Parametro']}: sem limite vigente em {coleta:%d/%m/%Y}", "Vigência do limite", "Revisar")
                continue
            for r in _avaliar(Lv, B):
                itens.append(r)
                if r["Status"] == "Bloqueado / Incompatível":
                    log(linha, am, f"{r['Parametro']}: {r['Obs']}", "Compatibilidade de unidades", "Revisar")
                elif r["Status"] == "Inconclusivo":
                    log(linha, am, f"{r['Parametro']}: {r['Obs']}", "Limite de quantificação", "Revisar")
    it = pd.DataFrame(itens)
    it["Alertas"] = ""
    it = _alertas_item(it)
    # observações informativas (duplicidade de limites e conversões) uma vez por critério
    for r in it[it.Duplicidade].drop_duplicates(["Cod_Ensaio", "Legislacao", "Classe"]).itertuples():
        log("—", "Vários", f"{r.Parametro}: limites duplicados em {r.Legislacao}; adotado o mais restritivo ({fmt_br(r.VMP_Max)} {r.Unidade})", "Limites cadastrados", "Revisar")
    for r in it[it.Obs.str.contains("convertido de", case=False)].drop_duplicates(["Cod_Ensaio"]).itertuples():
        log("—", "Vários", f"{r.Parametro}: {r.Obs.split('. ')[0].rstrip('.')}", "Conversão de unidade", "Informativo")
    n_ok = int(it.Status.isin(["Conforme", "Não Conforme"]).sum())
    log("—", "Todas", f"{n_ok} registros válidos avaliados contra os limites", "Resumo", "OK")
    a = pd.DataFrame(aud)
    a["_o"] = pd.to_numeric(a.Linha, errors="coerce").fillna(1e9)
    a = a.sort_values("_o", kind="stable").drop(columns="_o").reset_index(drop=True)
    return it, a, pd.DataFrame(amostras), info


def _alertas_item(it):
    """Marca, por item, ≥80% do VMP (conformes) e 3 altas consecutivas na série do ponto."""
    it = it.copy()
    flags = {i: [] for i in it.index}
    for i, r in it.iterrows():
        if r.Status == "Conforme" and pd.notna(r.Pct_VMP) and r.Pct_VMP >= 0.8:
            flags[i].append("LIMIAR_80PCT")
    ok = it[it.Pct_VMP.notna() & it.Status.isin(["Conforme", "Não Conforme"])]
    for _, g in ok.groupby(["Ponto_Coleta", "Cod_Ensaio", "Legislacao"]):
        g = g.sort_values("Data_Coleta")
        v, idx = g.Valor.values, list(g.index)
        for k in range(2, len(v)):
            if v[k] > v[k - 1] > v[k - 2]:
                flags[idx[k]].append("TENDENCIA_3_ALTA")
    it["Alertas"] = [", ".join(flags[i]) for i in it.index]
    return it


# ------------------------------------------------------------------ montagem do pacote
def montar(x, origem=""):
    erros, av = validar(x)
    if erros:
        raise ValueError(" | ".join(erros))
    itens, aud, amostras, info = processar(x)
    res = itens[itens.Status.isin(AVALIADOS)].reset_index(drop=True)
    return dict(info=info, bruto=x, itens=itens, res=res, aud=aud, amostras=amostras, avisos=av, origem=origem,
                pareceres=pareceres(itens, amostras), overrides={}, textos={},
                gab_itens=x.get("Gabarito_Itens"), gab_amostras=x.get("Gabarito_Amostras"))


def pareceres(itens, amostras):
    out = []
    for a in amostras.itertuples():
        g = itens[(itens.ID_Amostra == a.ID_Amostra) & itens.Status.isin(AVALIADOS)]
        av = g[g.Status.isin(["Conforme", "Não Conforme"]) | g.Status.isin(PENDENTES)]
        fe = g[g.Status == "Fora de Escopo"]
        nc, pend = g[g.Status == "Não Conforme"], g[g.Status.isin(PENDENTES)]
        if len(nc):
            p = "NÃO ATENDE"
        elif len(pend) or av.empty:
            p = "REVISÃO MANUAL"
        else:
            p = "ATENDE"
        legs = g[g.Legislacao != "-"][["Legislacao", "Classe"]].drop_duplicates().values
        cl = {}
        for l, c in legs:
            cl.setdefault(l, [])
            if c and c != "Sem classe" and c not in cl[l]:
                cl[l].append(c)
        base = "; ".join(f"{l} (Classe {', '.join(c)})" if c else l for l, c in cl.items()) or "—"
        txt = f"Análise realizada com base em: {base}. "
        if p == "NÃO ATENDE":
            det = "; ".join(f"{r.Parametro} (Medido: {r.Resultado} {r.Unidade} | VMP máx.: {fmt_br(r.VMP_Max)} {r.Unidade} | {r.Legislacao})"
                            for r in nc.itertuples())
            txt += f"A amostra NÃO ATENDE às especificações regulatórias devido à não conformidade no(s) parâmetro(s): {det}."
            if len(pend):
                txt += " Itens adicionais em revisão: " + "; ".join(f"{r.Parametro} ({r.Status})" for r in pend.itertuples()) + "."
        elif p == "REVISÃO MANUAL":
            txt += "Parecer em revisão manual: " + "; ".join(f"{r.Parametro} ({r.Status})" for r in pend.itertuples()) + "." \
                if len(pend) else "Parecer em revisão manual: nenhum parâmetro avaliável."
        else:
            txt += f"A amostra ATENDE: os {len(av)} parâmetros avaliados estão dentro dos limites aplicáveis."
        if len(fe):
            txt += " Parâmetros fora do escopo contratado (informativo): " + ", ".join(sorted(set(fe.Parametro))) + "."
        out.append(dict(ID_Amostra=a.ID_Amostra, Cliente=a.Cliente, Ponto_Coleta=a.Ponto_Coleta, Matriz_Agua=a.Matriz_Agua,
                        Data_Coleta=a.Data_Coleta, Parametros=len(av), Nao_Conformes=len(nc), Pendentes=len(pend),
                        Fora_Escopo=len(fe), Duplicidades=int(g.Duplicidade.sum()), Parecer_Auto=p, Texto=txt,
                        Legislacoes=", ".join(sorted({l for l, _ in legs})), Desvios=", ".join(nc.Parametro.unique())))
    return pd.DataFrame(out)


# ------------------------------------------------------------------ análises
def estatisticas(res):
    n = res[(res.Operador == "") & res.Valor.notna() & res.Status.isin(["Conforme", "Não Conforme"])]
    n = n[~n.Resultado.isin(["Presente", "Ausente"])].drop_duplicates(["ID_Amostra", "Cod_Ensaio"])
    if n.empty:
        return pd.DataFrame(columns=["Parametro", "Unidade", "N", "Media", "Minimo", "Maximo", "Desvio", "Amplitude", "CV_pct", "Max_pct_VMP"])
    s = n.groupby(["Parametro", "Unidade"]).Valor.agg(N="count", Media="mean", Minimo="min", Maximo="max", Desvio="std").reset_index()
    s["Amplitude"] = s.Maximo - s.Minimo
    s["CV_pct"] = np.where(s.Media > 0, s.Desvio / s.Media * 100, np.nan)
    mx = res.groupby("Parametro").Pct_VMP.max().rename("Max_pct_VMP").reset_index()
    return s.merge(mx, on="Parametro", how="left")


def tendencias(res):
    """Alerta preventivo (≥80% do VMP ou 3 altas consecutivas) / estouro (≥100%) + projeção linear."""
    out = []
    ok = res[res.Pct_VMP.notna() & res.Status.isin(["Conforme", "Não Conforme"])]
    for (pto, par, leg), g in ok.groupby(["Ponto_Coleta", "Parametro", "Legislacao"]):
        g = g.sort_values("Data_Coleta")
        v, d, vmax = g.Valor.values, g.Data_Coleta.values, g.VMP_Max.iloc[-1]
        dias = (d - d[0]).astype("timedelta64[D]").astype(float)
        slope = np.polyfit(dias, v, 1)[0] if len(g) >= 3 and dias[-1] > 0 else 0.0
        alta3 = len(v) >= 3 and v[-1] > v[-2] > v[-3]
        pct = v[-1] / vmax
        cruz = pd.Timestamp(d[-1]) + pd.Timedelta(days=float((vmax - v[-1]) / slope)) if slope > 0 and v[-1] < vmax else pd.NaT
        nivel = "Estouro do VMP" if pct >= 1 else "Alerta preventivo" if (pct >= 0.8 or alta3) else "Normal"
        tend = "Ascendente Crítica" if (nivel != "Normal" and slope > 0) else "Ascendente" if slope > 0 else "Descendente" if slope < 0 else "Estável"
        out.append(dict(Ponto_Coleta=pto, Parametro=par, Legislacao=leg, Medicoes=len(g), Ultimo=v[-1], Ultima_Data=pd.Timestamp(d[-1]),
                        VMP_Max=vmax, Pct_VMP=pct, Alta_3=alta3, Inclinacao_dia=slope, Cruzamento_Previsto=cruz,
                        Nivel=nivel, Tendencia=tend, Unidade=g.Unidade.iloc[-1]))
    return pd.DataFrame(out, columns=["Ponto_Coleta", "Parametro", "Legislacao", "Medicoes", "Ultimo", "Ultima_Data", "VMP_Max", "Pct_VMP",
                                      "Alta_3", "Inclinacao_dia", "Cruzamento_Previsto", "Nivel", "Tendencia", "Unidade"])


_GAB_ITEM = {"Conforme": "CONFORME", "Não Conforme": "NÃO CONFORME", "Fora de Escopo": "FORA DE ESCOPO",
             "Bloqueado / Incompatível": "BLOQUEADO", "Inconclusivo": "INCONCLUSIVO",
             "Descartado": "DESCARTADO (RUÍDO)", "Isolado": "BLOQUEADO (ISOLAMENTO)", "Revisão Manual": "REVISÃO MANUAL"}


def _limpa(s):
    return re.sub(r"\s*\(RN-\d+\)", "", str(s)).strip().upper()


def conferir_gabarito(pk):
    """Compara o resultado do motor com as abas Gabarito_* da planilha (quando existirem)."""
    r = {}
    if pk.get("gab_itens") is not None:
        gi = pk["gab_itens"].copy()
        gi["k"] = gi.codigo_amostra.astype(str).str.strip() + "|" + gi.codigo_ensaio.astype(str) + "|" + gi.nome_legislacao.astype(str).str.strip()
        it = pk["itens"].copy()
        it["k"] = it.ID_Amostra + "|" + it.Cod_Ensaio.astype(str) + "|" + it.Legislacao
        it["Obtido"] = it.Status.map(_GAB_ITEM)
        t = gi.merge(it[["k", "Obtido", "Alertas"]], on="k", how="left")
        t["Esperado"] = t.status_esperado.map(_limpa)
        ae = t.alertas_esperados.fillna("").astype(str).map(lambda s: ", ".join(sorted(p.strip() for p in s.split(",") if p.strip())))
        ao = t.Alertas.fillna("").map(lambda s: ", ".join(sorted(p.strip() for p in s.split(",") if p.strip())))
        t["Confere"] = (t.Esperado == t.Obtido.fillna("")) & (ae == ao)
        r["itens"] = pd.DataFrame(dict(Amostra=t.codigo_amostra, Ensaio=t.nome_ensaio, Esperado=t.Esperado,
                                       Obtido=t.Obtido.fillna("(ausente)"), Alertas_esperados=ae, Alertas_obtidos=ao, Confere=t.Confere))
    if pk.get("gab_amostras") is not None:
        ga = pk["gab_amostras"].copy()
        p = pk["pareceres"][["ID_Amostra", "Parecer_Auto"]]
        t = ga.merge(p, left_on="codigo_amostra", right_on="ID_Amostra", how="left")
        t["Confere"] = t.parecer_esperado.map(_limpa) == t.Parecer_Auto.fillna("")
        r["amostras"] = pd.DataFrame(dict(Amostra=t.codigo_amostra, Esperado=t.parecer_esperado, Obtido=t.Parecer_Auto.fillna("(ausente)"), Confere=t.Confere))
    return r

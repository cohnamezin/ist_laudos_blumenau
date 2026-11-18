"""Automação de Laudos e Conformidade Legal Ambiental · IST Blumenau (CDIA 2026/2) · v4
Protótipo de front-end em Python (Streamlit). Dados mocados lidos das planilhas UltraLIMS; tudo em memória.
"""
import datetime as dt
import itertools
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import estilo as e
import motor as m
from estilo import AMB, B, D, G, MUT, R, Y

PAGES = ["Dashboard", "Importação", "Conferência", "Parecer", "Análise temporal"]
FLUXO = ["Conferência de dados", "Comparação legal", "Análise histórica", "Emissão de parecer"]
ETAPAS = ["Upload", "Mapeamento", "Consistência", "Pré-visualização"]
PERFIS = ["Analista", "Gestor Técnico", "Administrador"]
PARECERES = ["Atende", "Não Atende", "Revisão Manual"]
UP = {"Atende": "ATENDE", "Não Atende": "NÃO ATENDE", "Revisão Manual": "REVISÃO MANUAL"}
DADOS = Path(__file__).parent / "dados"
COR_STATUS = {"Conforme": G, "Não Conforme": R, "Fora de Escopo": "#9AA7B8", "Bloqueado / Incompatível": AMB,
              "Inconclusivo": "#E0B04A", "Revisão Manual": "#E8C877"}

st.set_page_config(page_title="IST Blumenau · Laudos e Conformidade", page_icon="🧪", layout="wide", initial_sidebar_state="expanded")
st.markdown(e.CSS, unsafe_allow_html=True)
md = lambda h: st.markdown(h, unsafe_allow_html=True)
_cn = itertools.count()
card = lambda: st.container(key=f"card_{next(_cn)}")
data_br = lambda t: "—" if pd.isna(t) else f"{pd.Timestamp(t):%d/%m/%Y}"


def titulo(t, sub=""):
    md(f"<h1>{t}</h1><div class='sub'>{sub}</div>")


def ch(t):
    md(f"<h3 class='ch'>{t}</h3>")


def esp(px=16):
    md(f"<div style='height:{px}px'></div>")


# ------------------------------------------------------------------ estado
def registrar(x, origem):
    try:
        pk = m.montar(x, origem)
    except ValueError as ex:
        return False, str(ex)
    num = str(pk["info"]["numero_proposta"])
    st.session_state.props[num] = pk
    st.session_state.ativa = num
    return True, num


def go_to(i):
    st.session_state.page = i


def passo(d):
    st.session_state.etapa = max(0, min(3, st.session_state.etapa + d))


def ir_parecer(am):
    st.session_state.sel_parecer, st.session_state.page = am, 3


if "props" not in st.session_state:
    st.session_state.update(props={}, audit=[], page=0, etapa=0, stage=None, ovr_msg=None, perfil="Gestor Técnico", busca="")
    for f in sorted(DADOS.glob("*_UltraLIMS*.xlsx")):
        registrar(m.ler_planilha(f), f"Planilha mocada: {f.name}")
    st.session_state.ativa = sorted(st.session_state.props)[-1]
PROPS = st.session_state.props

# ------------------------------------------------------------------ barra superior: busca · proposta · perfil
with st.container(key="topbar"):
    c1, c2, c3 = st.columns([3, 3.6, 2])
    c1.text_input("Buscar amostra ou cliente", key="busca", placeholder="Buscar amostra ou cliente", label_visibility="collapsed")
    c2.selectbox("Proposta comercial", list(PROPS), key="ativa", label_visibility="collapsed",
                 format_func=lambda n: f"Proposta {n} · {PROPS[n]['info']['razao_social']}")
    c3.selectbox("Perfil", PERFIS, key="perfil", label_visibility="collapsed", format_func=lambda p: f"{p} · Perfil autorizado")
PK = PROPS[st.session_state.ativa]
NUM, I = st.session_state.ativa, PK["info"]
perfil, PODE = st.session_state.perfil, st.session_state.perfil != "Analista"
busca = st.session_state.busca.strip().lower()

with st.sidebar:
    md("<div class='brand'>IST Blumenau<small>Laudos e Conformidade</small></div>")
    for i, p in enumerate(PAGES):
        st.button(p, key=f"{'navon' if i == st.session_state.page else 'nav'}_{i}", on_click=go_to, args=(i,), use_container_width=True)
    vig = f"{data_br(I.get('data_inicio'))} a {data_br(I.get('data_fim'))}" if pd.notna(I.get("data_inicio")) else "—"
    md(f"<div class='side-ctx'>Proposta ativa<b>{NUM}</b>{I['razao_social']}<br>{I.get('cnpj_cpf', '')}<br>"
       f"Vigência {vig}<br>Atendimento {I.get('frequencia_atendimento', '—')}</div>"
       "<div class='side-proto'>Protótipo · dados mocados das planilhas UltraLIMS. Nada é gravado.</div>")
page = st.session_state.page

# ------------------------------------------------------------------ dados derivados da proposta ativa
res, itens = PK["res"], PK["itens"]
P = PK["pareceres"].copy()
P["Parecer"] = [PK["overrides"].get(i, a) for i, a in zip(P.ID_Amostra, P.Parecer_Auto)]
P["Override"] = P.ID_Amostra.isin(PK["overrides"])
TEND = m.tendencias(res)
AUDIT = [a for a in st.session_state.audit if a["Proposta"] == NUM]
nome = lambda p: e.NOME_PARECER.get(p, p)


def ocultar_busca(df, cols):
    if not busca:
        return df
    msk = pd.Series(False, index=df.index)
    for c in cols:
        msk |= df[c].astype(str).str.lower().str.contains(busca, regex=False)
    return df[msk]


def empilhado(col):
    t = res.assign(**{col: res[col].replace("", "—")})
    ct = pd.crosstab(t[col], t.Status)
    f = go.Figure()
    for k, c in COR_STATUS.items():
        if k in ct:
            f.add_bar(x=ct.index, y=ct[k], name=k, marker_color=c)
    f.update_layout(barmode="stack")
    return e.plotly_estilo(f, 280)


def grafico(f):
    st.plotly_chart(f, use_container_width=True, config=e.PLOTLY_CFG)


# ================================================================== 1. DASHBOARD
if page == 0:
    titulo("Dashboard", f"Visão geral dos laudos · {I['razao_social']}")
    esp(8)
    nconf = int((P.Parecer == "ATENDE").sum())
    k = st.columns(4)
    alertas_n = int((TEND.Nivel != "Normal").sum()) if len(TEND) else 0
    k[0].markdown(e.kpi("Laudos processados", len(P), B, f"{len(itens)} ensaios lidos"), unsafe_allow_html=True)
    k[1].markdown(e.kpi("Taxa de conformidade", f"{nconf / len(P) * 100:.0f}%" if len(P) else "—", G, f"{nconf} de {len(P)} amostras atendem"), unsafe_allow_html=True)
    k[2].markdown(e.kpi("Pendentes de revisão", int((P.Parecer == "REVISÃO MANUAL").sum()), Y, "Revisão manual"), unsafe_allow_html=True)
    k[3].markdown(e.kpi("Alertas de tendência", alertas_n, R, "Pontos e parâmetros monitorados"), unsafe_allow_html=True)
    esp()
    a, b = st.columns([2, 1])
    with a:
        with card():
            ch("Conformidade por semana")
            semana = P.Data_Coleta.dt.normalize() - pd.to_timedelta(P.Data_Coleta.dt.weekday, unit="D")
            w = pd.crosstab(semana, P.Parecer).sort_index()
            f = go.Figure()
            for k_, cor in (("ATENDE", G), ("NÃO ATENDE", R), ("REVISÃO MANUAL", AMB)):
                if k_ in w:
                    f.add_bar(x=[f"{d:%d/%m}" for d in w.index], y=w[k_], name=nome(k_), marker_color=cor)
            f.update_layout(barmode="stack", xaxis_title="Semana (início)", yaxis_title="Amostras")
            f.update_xaxes(type="category"); f.update_yaxes(dtick=1)
            grafico(e.plotly_estilo(f, 260))
    with b:
        with card():
            ch("Alertas recentes")
            linhas = []
            for r in TEND[TEND.Nivel != "Normal"].sort_values("Ultima_Data", ascending=False).itertuples():
                sel_ = e.badge("Acima do VMP", "r") if r.Pct_VMP >= 1 else e.badge(f"{r.Pct_VMP * 100:.0f}% do VMP", "y") if r.Pct_VMP >= .8 else e.badge("Tendência de alta", "y")
                linhas.append(f"<div class='alrow'><span>{r.Parametro} · {r.Ponto_Coleta}</span>{sel_}</div>")
            for r in res[res.Status.isin(m.PENDENTES)].itertuples():
                linhas.append(f"<div class='alrow'><span>{r.Parametro} · {r.ID_Amostra}</span>{e.badge(r.Status)}</div>")
            md("".join(linhas[:8]) if linhas else "<div class='legend'>Nenhum alerta para esta proposta.</div>")
    esp()
    with card():
        ch("Últimas amostras")
        U = ocultar_busca(P, ["ID_Amostra", "Cliente", "Ponto_Coleta", "Legislacoes"]).sort_values("Data_Coleta", ascending=False).head(8)
        W = [1.3, 2.3, 2.6, 1.5, 1.2]
        hd = st.columns(W)
        for c, t in zip(hd, ["Amostra", "Cliente", "Legislação", "Parecer", "Ação"]):
            c.markdown(f"<div style='font-size:12px;color:{MUT};padding:4px 12px'>{t}</div>", unsafe_allow_html=True)
        for r in U.itertuples():
            cs = st.columns(W)
            cs[0].markdown(f"<div style='padding-top:10px'>{r.ID_Amostra}</div>", unsafe_allow_html=True)
            cs[1].markdown(f"<div style='padding-top:10px'>{r.Cliente}</div>", unsafe_allow_html=True)
            cs[2].markdown(f"<div style='padding-top:10px'>{r.Legislacoes or '—'}</div>", unsafe_allow_html=True)
            cs[3].markdown(f"<div style='padding-top:8px'>{e.badge(nome(r.Parecer))}</div>", unsafe_allow_html=True)
            cs[4].button("Revisar", key=f"rev_{NUM}_{r.ID_Amostra}", on_click=ir_parecer, args=(r.ID_Amostra,))
        if U.empty:
            st.caption("Nenhuma amostra encontrada para a busca.")
    # ---- análises complementares
    esp(8)
    md(f"<h3 class='ch' style='margin-top:8px'>Análises complementares</h3>")
    g1, g2 = st.columns(2)
    with g1:
        with card():
            ch("Distribuição por matriz / origem"); grafico(empilhado("Matriz_Agua"))
    with g2:
        with card():
            ch("Distribuição por célula laboratorial"); grafico(empilhado("Celula"))
    esp()
    with card():
        ch("Estatísticas descritivas dos ensaios")
        E = m.estatisticas(res)
        if E.empty:
            st.caption("Sem ensaios numéricos suficientes.")
        else:
            E = E.rename(columns={"N": "Nº", "Media": "Média", "Minimo": "Mínimo", "Maximo": "Máximo", "Desvio": "Desvio-padrão", "CV_pct": "CV (%)", "Max_pct_VMP": "Máx. % do VMP"})
            f4 = {c: "{:.4g}" for c in ["Média", "Mínimo", "Máximo", "Amplitude", "Desvio-padrão"]}
            md(e.tabela(E, fmt={**f4, "CV (%)": "{:.1f}", "Máx. % do VMP": "{:.0%}"}))
    esp()
    v1, v2 = st.columns(2)
    with v1:
        with card():
            ch("Maior variabilidade (coeficiente de variação)")
            cv = E.dropna(subset=["CV (%)"]).nlargest(5, "CV (%)") if not E.empty else E
            if len(cv):
                f = go.Figure(go.Bar(x=cv["CV (%)"], y=cv.Parametro, orientation="h", marker_color=B))
                f.update_layout(yaxis=dict(autorange="reversed"), xaxis_title="CV (%)")
                grafico(e.plotly_estilo(f, 250))
            else:
                st.caption("São necessárias ao menos 2 medições por parâmetro.")
    with v2:
        with card():
            ch("Proximidade do limite legal (% do VMP máximo)")
            px_ = res.dropna(subset=["Pct_VMP"]).groupby("Parametro").Pct_VMP.max().nlargest(6).reset_index()
            if len(px_):
                f = go.Figure(go.Bar(x=px_.Pct_VMP * 100, y=px_.Parametro, orientation="h",
                                     marker_color=[R if v >= 1 else AMB if v >= .8 else G for v in px_.Pct_VMP]))
                f.add_vline(x=80, line=dict(color=AMB, dash="dash")); f.add_vline(x=100, line=dict(color=R, dash="dash"))
                f.update_layout(yaxis=dict(autorange="reversed"), xaxis_title="% do VMP")
                grafico(e.plotly_estilo(f, 250))
    cg = m.conferir_gabarito(PK)
    if cg:
        esp()
        with card():
            ch("Conferência com a planilha de referência")
            for chave, rot in (("amostras", "pareceres"), ("itens", "itens avaliados")):
                if chave in cg:
                    t = cg[chave]
                    md(e.alerta(f"{int(t.Confere.sum())} de {len(t)} {rot} coincidem com o gabarito da planilha.", "g" if t.Confere.all() else "y"))
            if "itens" in cg and not cg["itens"].Confere.all():
                md(e.tabela(cg["itens"][~cg["itens"].Confere]))

# ================================================================== 2. IMPORTAÇÃO
elif page == 1:
    et = st.session_state.etapa
    stg = st.session_state.stage
    if stg is None or not stg.get("upload"):
        stg = st.session_state.stage = dict(x=PK["bruto"], origem=PK["origem"], upload=False, pk=PK)
    titulo("Importação UltraLIMS", f"Etapa {et + 1} de 4")
    esp(8)
    with card():
        md(e.stepper(ETAPAS, et))
    esp()
    X = stg["x"]
    erros, avisos = m.validar(X)
    if not erros and stg.get("pk") is None:
        try:
            stg["pk"] = m.montar(X, stg["origem"])
        except Exception as ex:  # planilha fora do padrão
            erros = [f"Não foi possível processar a planilha: {ex}"]
    SP = stg.get("pk")
    with card():
        if et == 0:
            ch("Upload do arquivo")
            up = st.file_uploader("Planilha UltraLIMS (.xlsx)", type=["xlsx"], key="up_file")
            if up is not None and st.session_state.get("up_visto") != (up.name, up.size):
                st.session_state.up_visto = (up.name, up.size)
                st.session_state.stage = dict(x=m.ler_planilha(up), origem=f"Upload: {up.name}", upload=True, pk=None)
                st.rerun()
            md(e.alerta(f"Arquivo em análise: <b>{stg['origem']}</b>", "b"))
            ok = [a for a in m.ABAS_OBRIG + m.ABAS_OPC if a in X]
            md("".join(e.chip(a) for a in ok) or "—")
            for t in erros:
                md(e.alerta(t, "r"))
            for t in avisos:
                md(e.alerta(t, "y"))
            if stg["upload"]:
                st.button("Voltar a usar a proposta mocada selecionada no topo", on_click=lambda: st.session_state.update(stage=None))
            else:
                st.caption("Dados mocados: troque a proposta no seletor do topo ou envie outra planilha no mesmo padrão.")
        elif erros:
            for t in erros:
                md(e.alerta(t, "r"))
        elif et == 1:
            ch("Mapeamento de campos (De/Para)")
            cols_ens = [c for c in X["Worksheet"].columns if m.PAD_COL_ENSAIO.match(c)]
            lin = []
            for aba, campo, ent, dest in m.DE_PARA:
                det = bool(cols_ens) if campo.startswith("<") else campo in X.get(aba, pd.DataFrame()).columns
                lin.append({"Aba da planilha": aba, "Campo": campo, "Entidade": ent, "Campo de destino": dest, "Situação": "OK" if det else "Erro"})
            md(e.tabela(pd.DataFrame(lin), badges=["Situação"], linha="Situação", altura=380))
            esp()
            ch("Colunas de ensaio detectadas (despivotamento)")
            cat = X["Catalogo_Ensaios"].assign(c=lambda t: pd.to_numeric(t.codigo_ensaio, errors="coerce")).set_index("c")
            ln = []
            for c in cols_ens:
                cod = int(m.PAD_COL_ENSAIO.match(c).group(1))
                no = cod in cat.index
                ln.append({"Coluna": c, "Código": cod, "Ensaio (catálogo)": cat.loc[cod, "nome_ensaio"] if no else "—",
                           "Célula": cat.loc[cod, "sigla_celula"] if no else "—", "Unidade": cat.loc[cod, "unidade_trabalho"] if no else "—",
                           "Situação": "OK" if no else ("Descartada" if m.PAD_ADMIN.search(c) else "Informativo")})
            md(e.tabela(pd.DataFrame(ln), badges=["Situação"]))
        elif et == 2:
            ch("Checagem de consistência")
            A = SP["aud"]
            chips = [f"{len(SP['amostras'])} amostras válidas", f"{int((A.Status.isin(['Bloqueada', 'Descartada'])).sum())} bloqueadas/descartadas",
                     f"{int((SP['itens'].Status == 'Fora de Escopo').sum())} fora de escopo",
                     f"{int((SP['itens'].Status == 'Bloqueado / Incompatível').sum())} unidades incompatíveis"]
            md("".join(e.chip(c) for c in chips)); esp(8)
            md(e.tabela(A, badges=["Status"], linha="Status", altura=460))
        else:
            ch(f"Pré-visualização · proposta {SP['info']['numero_proposta']}")
            k = st.columns(4)
            k[0].markdown(e.kpi("Registros lidos", len(SP["itens"])), unsafe_allow_html=True)
            k[1].markdown(e.kpi("Amostras válidas", len(SP["amostras"])), unsafe_allow_html=True)
            k[2].markdown(e.kpi("Legislações", SP["res"].query("Legislacao != '-'").Legislacao.nunique()), unsafe_allow_html=True)
            k[3].markdown(e.kpi("Cliente", f"<span style='font-size:18px'>{SP['info']['razao_social']}</span>"), unsafe_allow_html=True)
            esp(8)
            v = SP["itens"][["ID_Amostra", "Ponto_Coleta", "Data_Coleta", "Parametro", "Resultado", "Unidade", "Status"]].copy()
            v["Data_Coleta"] = v.Data_Coleta.map(data_br)
            v.columns = ["Amostra", "Ponto de coleta", "Coleta", "Ensaio", "Resultado", "Unid.", "Status"]
            md(e.tabela(v, badges=["Status"], linha="Status", altura=380))
            if stg["upload"] and str(SP["info"]["numero_proposta"]) in PROPS:
                st.caption("Esta proposta já está carregada e será substituída.")
    esp(4)
    b = st.columns([1, 3, 7])
    b[0].button("Voltar", on_click=passo, args=(-1,), disabled=et == 0, use_container_width=True)
    if et < 3:
        rot = ["Continuar para mapeamento", "Continuar para consistência", "Continuar para pré-visualização"][et]
        b[1].button(rot, type="primary", on_click=passo, args=(1,), disabled=bool(erros), use_container_width=True)
    else:
        def fim():
            if stg["upload"]:
                ok_, num_ = registrar(stg["x"], stg["origem"])
                if not ok_:
                    return
            st.session_state.update(page=0, etapa=0, stage=None)
        b[1].button("Registrar proposta", type="primary", on_click=fim, disabled=bool(erros), use_container_width=True)

# ================================================================== 3. CONFERÊNCIA (MATRIZ DE VMP)
elif page == 2:
    ids = [i for i in P.ID_Amostra if not busca or busca in i.lower() or busca in P.set_index("ID_Amostra").loc[i, "Ponto_Coleta"].lower()]
    ph = st.container()
    esp(8)
    with card():
        md(e.stepper(FLUXO, 1))
    esp()
    with card():
        c = st.columns([1.6, 2, 1.6, 1.8])
        am = c[0].selectbox("Amostra", ["Todas"] + ids, key=f"mz_am_{NUM}")
        base = res if am == "Todas" else res[res.ID_Amostra == am]
        base = base[base.ID_Amostra.isin(ids)]
        leg = c[1].selectbox("Legislação", ["Todas"] + sorted(set(base.Legislacao) - {"-"}), key=f"mz_leg_{NUM}")
        base = base if leg == "Todas" else base[base.Legislacao == leg]
        cls = c[2].selectbox("Classe", ["Todas"] + sorted(set(base.Classe) - {""}), key=f"mz_cls_{NUM}")
        base = base if cls == "Todas" else base[base.Classe == cls]
        mtz = c[3].selectbox("Matriz / origem", ["Todas"] + sorted(set(base.Matriz_Agua)), key=f"mz_mtz_{NUM}")
        base = base if mtz == "Todas" else base[base.Matriz_Agua == mtz]
    with ph:
        titulo("Matriz de conferência de VMP", f"{'Amostra ' + am if am != 'Todas' else 'Todas as amostras'} · dados mocados da proposta {NUM}")
    esp()
    with card():
        chips = [f"{l} · Classe {c_}" if c_ and c_ != "Sem classe" else l for l, c_ in base[base.Legislacao != "-"][["Legislacao", "Classe"]].drop_duplicates().values]
        chips += sorted(set(base.Matriz_Agua))
        md(f"<span style='font-size:13px;color:{MUT};margin-right:8px'>Filtros:</span>" + "".join(e.chip(x) for x in chips))
    esp()
    with card():
        def vmp_txt(r, col):
            if r.Unidade == "P/A" and col == "VMP_Max" and r.VMP_Max == 0:
                return "Ausente"
            return "—" if pd.isna(getattr(r, col)) or (col == "VMP_Min" and r.Unidade == "P/A") else m.fmt_br(getattr(r, col))
        t = pd.DataFrame([{"Amostra": r.ID_Amostra, "Ensaio": r.Parametro, "Resultado": r.Resultado, "Unid.": r.Unidade,
                           "VMP mín.": vmp_txt(r, "VMP_Min"), "VMP máx.": vmp_txt(r, "VMP_Max"),
                           "Legislação": "—" if r.Legislacao == "-" else r.Legislacao,
                           "Status": r.Status + ("|Duplicidade" if r.Duplicidade else ""),
                           "Observação": r.Obs} for r in base.itertuples()])
        if t.empty:
            st.caption("Nenhum ensaio para os filtros selecionados.")
        else:
            show = t if am == "Todas" else t.drop(columns="Amostra")
            md(e.tabela(show, badges=["Status"], linha="Status", altura=460))
    esp()
    n_nc = int((base.Status == "Não Conforme").sum())
    z = st.columns([3, 2.4, 2])
    if am != "Todas":
        pa = P.set_index("ID_Amostra").loc[am, "Parecer"]
        z[0].markdown(e.badge(f"Amostra: {nome(pa)}"), unsafe_allow_html=True)
    else:
        z[0].markdown(e.badge(f"{len(base)} ensaios · {n_nc} não conformes", "r" if n_nc else "g"), unsafe_allow_html=True)
    z[1].button("Avançar para análise histórica", type="primary", on_click=go_to, args=(4,), use_container_width=True)
    if not t.empty:
        z[2].download_button("Baixar matriz (CSV)", t.to_csv(index=False).encode("utf-8-sig"), f"matriz_vmp_{NUM}.csv", use_container_width=True)

# ================================================================== 4. PARECER CONCLUSIVO
elif page == 3:
    ids = list(P.ID_Amostra)
    if st.session_state.get("sel_parecer") not in ids:
        st.session_state.sel_parecer = ids[0]
    h = st.columns([4, 1.6])
    with h[0]:
        titulo("Parecer conclusivo", "Detalhe da amostra")
    sel = h[1].selectbox("Amostra", ids, key="sel_parecer", label_visibility="collapsed")
    p = P[P.ID_Amostra == sel].iloc[0]
    g = res[res.ID_Amostra == sel]
    esp(8)
    cor = G if p.Parecer == "ATENDE" else R if p.Parecer == "NÃO ATENDE" else Y
    pend = ", ".join(g[g.Status.isin(m.PENDENTES)].Parametro.unique())
    desv = f"Parâmetro em desvio: {p.Desvios}" if p.Desvios else (f"Itens em revisão: {pend}" if pend else "Nenhum parâmetro em desvio")
    a, b = st.columns([1, 1.4])
    with a:
        with card():
            md(f"<div class='kl'>Parecer {'(override manual)' if p.Override else 'automático'}</div>"
               f"<div class='status' style='color:{cor}'>{nome(p.Parecer)}</div><div style='font-size:14px;margin-top:8px'>{desv}</div>")
    with b:
        with card():
            ch("Resumo")
            md(f"<div style='font-size:14px;line-height:1.7'>Amostra {sel} · {p.Cliente}<br>Ponto: {p.Ponto_Coleta} · Matriz: {p.Matriz_Agua}<br>"
               f"Coleta: {data_br(p.Data_Coleta)}<br>Legislações: {p.Legislacoes or '—'}</div>")
    esp()
    key = f"txt_{NUM}_{sel}"
    if key not in st.session_state:
        st.session_state[key] = PK["textos"].get(sel, p.Texto)
    kj, kn = f"ovr_just_{NUM}", f"ovr_novo_{NUM}"

    def override(anterior):
        just, novo = st.session_state.get(kj, "").strip(), UP[st.session_state.get(kn, "Atende")]
        if len(just) < 10:
            st.session_state.ovr_msg = ("r", "Justificativa obrigatória (mínimo de 10 caracteres).")
        elif novo == anterior:
            st.session_state.ovr_msg = ("y", "Selecione um parecer diferente do atual.")
        else:
            st.session_state.audit.append({"Proposta": NUM, "Usuário": perfil, "Data e hora": dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                                           "Amostra": sel, "Valor anterior": nome(anterior), "Valor atual": nome(novo), "Justificativa": just})
            PK["overrides"][sel] = novo
            st.session_state[kj] = ""
            st.session_state.ovr_msg = ("g", f"Parecer da amostra {sel} alterado para {nome(novo)} e registrado na auditoria.")

    with card():
        ch("Texto do parecer e análise crítica")
        st.text_area("Texto do parecer", key=key, height=150, label_visibility="collapsed")
        z = st.columns([1.4, 2.2, 1.6, 2])
        if z[0].button("Salvar alterações", type="primary", use_container_width=True):
            PK["textos"][sel] = st.session_state[key]
            st.toast("Texto do parecer salvo (em memória).")
        z[1].button("Alterar parecer (override)", on_click=override, args=(p.Parecer,), disabled=not PODE, use_container_width=True)
        z[2].button("Restaurar sugestão", on_click=lambda: st.session_state.update({key: p.Texto}), use_container_width=True)
        z[3].download_button("Baixar parecer (TXT)", f"{sel} - {nome(p.Parecer)}\n\n{st.session_state[key]}".encode("utf-8"), f"parecer_{sel.replace('/', '-')}.txt", use_container_width=True)
    esp()
    with card():
        ch("Justificativa do override (obrigatória)")
        if not PODE:
            md(e.alerta("O perfil Analista não pode alterar pareceres. Selecione Gestor Técnico ou Administrador no topo da página.", "y"))
        if st.session_state.ovr_msg:
            md(e.alerta(st.session_state.ovr_msg[1], st.session_state.ovr_msg[0]))
            st.session_state.ovr_msg = None
        j = st.columns([1, 3])
        j[0].selectbox("Novo parecer", PARECERES, key=kn, disabled=not PODE)
        j[1].text_input("Motivo da alteração", key=kj, disabled=not PODE)
    esp()
    with card():
        ch("Histórico de auditoria")
        if AUDIT:
            md(e.tabela(pd.DataFrame(AUDIT).drop(columns="Proposta")[["Usuário", "Data e hora", "Amostra", "Valor anterior", "Valor atual", "Justificativa"]],
                        badges=["Valor anterior", "Valor atual"]))
        else:
            st.caption("Nenhuma alteração manual registrada nesta proposta nesta sessão.")
    esp()
    with card():
        ch("Parâmetros avaliados")
        t = g[["Parametro", "Resultado", "Unidade", "VMP_Min", "VMP_Max", "Legislacao", "Status", "Obs"]].copy()
        t.columns = ["Parâmetro", "Resultado", "Unid.", "VMP mín.", "VMP máx.", "Legislação", "Status", "Observação"]
        md(e.tabela(t, badges=["Status"], linha="Status", fmt={"VMP mín.": "{:g}", "VMP máx.": "{:g}"}))

# ================================================================== 5. ANÁLISE TEMPORAL
else:
    titulo("Análise temporal e alertas", "Tendências e desvios críticos")
    esp(8)
    if TEND.empty:
        md(e.alerta("Esta proposta ainda não tem parâmetros com percentual de VMP calculável.", "b"))
        st.stop()
    a, b = st.columns([2, 1])
    with b:
        with card():
            ch("Filtros")
            md(f"<div style='font-size:14px;margin-bottom:8px'>Cliente: {I['razao_social']}</div>")
            per = st.selectbox("Período", ["Todo o período", "Últimos 60 dias", "Últimos 30 dias"], key=f"tp_per_{NUM}")
            ptos = sorted(TEND.Ponto_Coleta.unique())
            pto = st.selectbox("Ponto de coleta", ptos, key=f"tp_pto_{NUM}")
            t2 = TEND[TEND.Ponto_Coleta == pto]
            par = st.selectbox("Parâmetro", sorted(t2.Parametro.unique()), key=f"tp_par_{NUM}_{pto}")
            t3 = t2[t2.Parametro == par]
            leg = st.selectbox("Legislação", sorted(t3.Legislacao.unique()), key=f"tp_leg_{NUM}_{pto}_{par}") if len(t3) > 1 else t3.Legislacao.iloc[0]
    tr = t3[t3.Legislacao == leg].iloc[0]
    g = res[(res.Ponto_Coleta == pto) & (res.Parametro == par) & (res.Legislacao == leg) & res.Pct_VMP.notna()].sort_values("Data_Coleta")
    if per != "Todo o período":
        g = g[g.Data_Coleta >= res.Data_Coleta.max() - pd.Timedelta(days=int(per.split()[1]))]
    vmax = tr.VMP_Max
    with a:
        with card():
            ch(f"Parâmetro: {par} · Ponto de coleta {pto}")
            f = go.Figure(go.Scatter(x=g.Data_Coleta, y=g.Valor, mode="lines+markers", name="Resultado", line=dict(color=B, width=3)))
            if tr.Inclinacao_dia > 0 and len(g):
                fim = tr.Cruzamento_Previsto if pd.notna(tr.Cruzamento_Previsto) else g.Data_Coleta.max() + pd.Timedelta(days=30)
                f.add_scatter(x=[g.Data_Coleta.max(), fim], y=[g.Valor.iloc[-1], g.Valor.iloc[-1] + tr.Inclinacao_dia * (fim - g.Data_Coleta.max()).days],
                              mode="lines", name="Projeção", line=dict(color=AMB, dash="dot", width=3))
            f.add_hline(y=vmax, line=dict(color=R, dash="dash"), annotation_text="VMP máximo", annotation_font_color=R)
            f.add_hline(y=vmax * .8, line=dict(color=AMB, dash="dash"), annotation_text="Alerta 80%", annotation_font_color=Y)
            f.update_layout(yaxis_title=tr.Unidade, xaxis_title="Data da coleta")
            grafico(e.plotly_estilo(f, 360))
            md("<div class='legend'>Linha azul: histórico. Tracejado vermelho: limite legal. Amarelo: alerta de 80% e projeção.</div>")
    esp()
    k = st.columns(4)
    k[0].markdown(e.kpi("Tendência", f"<span style='font-size:22px'>{tr.Tendencia}</span>", R if "Crítica" in tr.Tendencia else B), unsafe_allow_html=True)
    k[1].markdown(e.kpi("% do VMP (última)", f"{tr.Pct_VMP * 100:.0f}%", R if tr.Pct_VMP >= 1 else Y if tr.Pct_VMP >= .8 else G), unsafe_allow_html=True)
    k[2].markdown(e.kpi("Medições", int(tr.Medicoes)), unsafe_allow_html=True)
    k[3].markdown(e.kpi("Cruzamento previsto", data_br(tr.Cruzamento_Previsto)), unsafe_allow_html=True)
    esp()
    if tr.Nivel == "Estouro do VMP":
        md(e.alerta("Limite legal excedido. Ação: reter o laudo e investigar a causa no ponto de coleta.", "r"))
    elif tr.Nivel == "Alerta preventivo":
        md(e.alerta("Alerta preventivo: parâmetro com 80% ou mais do VMP, ou três altas consecutivas. Ação preventiva recomendada.", "y"))
    else:
        md(e.alerta("Parâmetro dentro da faixa segura.", "g"))
    with card():
        ch("Alertas")
        al = TEND[TEND.Nivel != "Normal"].copy()
        if al.empty:
            st.caption("Nenhum alerta ativo.")
        else:
            al["Pct"] = (al.Pct_VMP * 100).round().astype(int).astype(str) + "%"
            al["Proj"] = al.Cruzamento_Previsto.map(data_br)
            al = al[["Parametro", "Ponto_Coleta", "Nivel", "Pct", "Tendencia", "Proj"]]
            al.columns = ["Parâmetro", "Ponto", "Situação", "% do VMP", "Tendência", "Projeção de cruzamento"]
            md(e.tabela(al, badges=["Situação"], linha="Situação"))
md(f"<div class='legend' style='margin-top:20px'>Protótipo de front-end · dados mocados da planilha {PK['origem'].replace('Planilha mocada: ', '')} · nada é gravado fora da sessão.</div>")

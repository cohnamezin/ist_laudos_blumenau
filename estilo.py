"""Tokens visuais, CSS e componentes HTML que reproduzem as telas de referência (Python puro)."""
from html import escape

import pandas as pd

D, B, G, R, Y = "#0B2545", "#1F5FBF", "#16794C", "#B42323", "#8A5A00"   # navy, azul, verde, vermelho, âmbar (texto)
AMB, MUT, LINE = "#C98A12", "#4A5A70", "#DDE3EC"
TINT = {"g": ("#DDF3E7", G), "r": ("#FBE0E0", R), "y": ("#FFF1CC", Y), "b": ("#DCE8FA", B), "n": ("#EEF1F6", MUT)}
KIND = {
    "Conforme": "g", "Atende": "g", "ATENDE": "g", "OK": "g",
    "Não Conforme": "r", "Não Atende": "r", "NÃO ATENDE": "r", "Erro": "r", "Estouro do VMP": "r",
    "Bloqueado / Incompatível": "y", "Inconclusivo": "y", "Revisão Manual": "y", "REVISÃO MANUAL": "y", "Alerta preventivo": "y",
    "Duplicidade": "y", "Revisar": "y", "Bloqueada": "y", "Pendente": "y", "Tendência de alta": "y",
    "Override": "b", "Informativo": "b",
    "Fora de Escopo": "n", "Descartada": "n", "Descartado": "n", "Isolado": "n",
}
ROWBG = {"Não Conforme": "#FBE0E0", "Não Atende": "#FBE0E0", "NÃO ATENDE": "#FBE0E0", "Erro": "#FBE0E0", "Estouro do VMP": "#FBE0E0",
         "Bloqueado / Incompatível": "#FFF1CC", "Inconclusivo": "#FFF1CC", "Revisão Manual": "#FFF1CC", "REVISÃO MANUAL": "#FFF1CC",
         "Revisar": "#FFF1CC", "Bloqueada": "#FFF1CC", "Alerta preventivo": "#FFF1CC"}
NOME_PARECER = {"ATENDE": "Atende", "NÃO ATENDE": "Não Atende", "REVISÃO MANUAL": "Revisão Manual"}
PLOTLY_CFG = {"displayModeBar": False, "scrollZoom": False, "doubleClick": False, "displaylogo": False}

CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
:root{{color-scheme:light}}
html,body,.stApp,[class*="css"]{{font-family:Inter,Arial,sans-serif}}
.stApp{{background:#F4F7FA;color:{D}}}
header[data-testid=stHeader],[data-testid=stToolbar],[data-testid=stDecoration],footer,#MainMenu{{display:none!important}}
[data-testid=stMainBlockContainer],.block-container{{padding:0 28px 40px!important;max-width:none!important}}
/* textos: sempre escuros no conteúdo principal */
.stMarkdown,.stMarkdown p,[data-testid=stMarkdownContainer],[data-testid=stMarkdownContainer] p,[data-testid=stCaptionContainer]{{color:{D}}}
[data-testid=stCaptionContainer],[data-testid=stCaptionContainer] p{{color:{MUT}}}
label,[data-testid=stWidgetLabel],[data-testid=stWidgetLabel] *{{color:{MUT}!important;font-size:13px}}
h1{{font-size:24px!important;font-weight:700!important;color:{D}!important;margin:0!important;padding:0!important}}
h3.ch{{margin:0 0 12px;font-size:16px;font-weight:600;color:{D}}}
.sub{{font-size:14px;color:{MUT};margin:4px 0 4px}}
/* barra lateral: somente elementos próprios (sem widgets) */
section[data-testid=stSidebar]{{background:{D}!important;width:220px!important;min-width:220px!important}}
section[data-testid=stSidebar]>div{{width:220px!important;padding:0 12px}}
[data-testid=stSidebarHeader],[data-testid=stSidebarCollapseButton],[data-testid=stSidebarCollapsedControl]{{display:none!important}}
section[data-testid=stSidebar] [data-testid=stSidebarUserContent]{{padding-top:20px}}
.brand{{color:#fff;font-weight:700;font-size:16px;padding:8px 16px 20px;line-height:1.35}}
.brand small{{display:block;font-weight:400;font-size:12px;opacity:.8}}
.side-ctx{{color:#fff;font-size:12px;line-height:1.5;padding:16px 16px 0;opacity:.85;border-top:1px solid #2a4a7a;margin-top:18px}}
.side-ctx b{{display:block;font-size:13px}}
.side-proto{{color:#fff;font-size:11px;opacity:.6;padding:14px 16px}}
section[data-testid=stSidebar] [class*="st-key-nav"] button{{width:100%;justify-content:flex-start;background:transparent!important;
  border:0!important;border-radius:8px;min-height:44px;padding:12px 16px;opacity:.85;box-shadow:none}}
section[data-testid=stSidebar] [class*="st-key-nav"] button:hover{{background:rgba(255,255,255,.08)!important}}
section[data-testid=stSidebar] [class*="st-key-navon"] button{{background:{B}!important;opacity:1}}
section[data-testid=stSidebar] button,section[data-testid=stSidebar] button *{{color:#fff!important;font-weight:400;font-size:14px;text-align:left}}
section[data-testid=stSidebar] [data-testid=stVerticalBlock]{{gap:6px}}
/* barra superior (busca · proposta · perfil) */
.st-key-topbar{{background:#fff;border-bottom:1px solid {LINE};margin:0 -28px 24px;padding:12px 28px;min-height:64px;align-items:center}}
.st-key-topbar [data-testid=stHorizontalBlock]{{align-items:center}}
/* cartões */
[class*="st-key-card"]{{background:#fff;border:1px solid {LINE};border-radius:12px;padding:20px}}
.kpi{{background:#fff;border:1px solid {LINE};border-radius:12px;padding:20px;cursor:default}}
.kl{{font-size:13px;color:{MUT}}}.kv{{font-size:32px;font-weight:700;margin-top:6px;line-height:1.1}}.ks{{font-size:12px;color:{MUT};margin-top:4px}}
.badge{{display:inline-block;font-size:12px;font-weight:600;padding:4px 10px;border-radius:999px;white-space:nowrap;margin-right:4px}}
.chip{{display:inline-block;font-size:12px;font-weight:600;padding:4px 10px;border-radius:999px;background:#DCE8FA;color:{B};margin:0 6px 4px 0}}
.tbw{{overflow:auto}}.tb{{width:100%;border-collapse:collapse}}
.tb th{{text-align:left;font-size:12px;color:{MUT};padding:10px 12px;border-bottom:1px solid {LINE};position:sticky;top:0;background:#fff;white-space:nowrap;font-weight:400}}
.tb td{{font-size:14px;color:{D};padding:10px 12px;border-bottom:1px solid #EEF1F6;vertical-align:top}}
.alert{{padding:10px 14px;border-radius:8px;margin-bottom:8px;font-size:14px;border-left:5px solid;color:{D}}}
.steps{{display:flex;gap:12px;flex-wrap:wrap}}.st{{display:flex;align-items:center;gap:8px;flex:1;font-size:14px;color:{D}}}
.st i{{width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;font-style:normal}}
.status{{font-size:36px;font-weight:700;line-height:1.1;margin-top:6px}}
.alrow{{font-size:14px;color:{D};padding:6px 0;display:flex;justify-content:space-between;gap:8px;align-items:center}}
.legend{{font-size:13px;color:{MUT};margin-top:6px}}
.spc{{height:16px}}
/* itens apenas informativos: sem cursor de clique e sem reação ao clique */
.kpi,.kl,.kv,.ks,.tb,.alert,.alrow,.legend,.brand,.side-ctx,.side-proto,h3.ch,.sub,[class*="st-key-card"] h1{{cursor:default}}
.badge,.chip,.steps,.st,.status{{cursor:default;pointer-events:none;user-select:none}}
.js-plotly-plot .nsewdrag,.js-plotly-plot .cursor-pointer,.js-plotly-plot .cursor-crosshair,.js-plotly-plot .legend *{{cursor:default!important}}
.modebar{{display:none!important}}
/* botões (principal: azul/branco · secundário: branco/azul · desabilitado: cinza) */
[data-testid=stMain] .stButton>button,[data-testid=stMain] .stDownloadButton>button,[data-testid=stMain] [data-testid^=stBaseButton]{{
  min-height:44px;border-radius:8px;font-weight:600;font-size:14px;padding:0 20px}}
[data-testid=stMain] [data-testid=stBaseButton-secondary],[data-testid=stMain] .stDownloadButton button{{background:#fff;color:{B};border:1px solid {B}}}
[data-testid=stMain] [data-testid=stBaseButton-primary]{{background:{B};color:#fff;border:1px solid {B}}}
[data-testid=stMain] button p,[data-testid=stMain] button *{{color:inherit!important}}
[data-testid=stMain] button:disabled{{background:#EEF1F6!important;color:#8A97A8!important;border-color:{LINE}!important;cursor:not-allowed}}
/* campos: fundo claro e texto escuro (evita texto branco sobre branco) */
input,textarea{{background:#fff!important;color:{D}!important;-webkit-text-fill-color:{D}!important;border-radius:8px!important}}
input::placeholder,textarea::placeholder{{color:#7A889C!important;-webkit-text-fill-color:#7A889C!important}}
[data-baseweb=input],[data-baseweb=textarea],[data-baseweb=base-input]{{background:#fff!important;border-color:{LINE}!important;border-radius:8px!important}}
[data-baseweb=select]>div{{background:#fff!important;border:1px solid {LINE}!important;border-radius:8px!important;min-height:40px}}
[data-baseweb=select] *{{color:{D}!important}}
[data-baseweb=select] svg{{fill:{MUT}!important}}
[data-baseweb=popover],[data-baseweb=popover] *,[data-baseweb=menu],[data-baseweb=menu] *{{background-color:#fff;color:{D}!important}}
[data-baseweb=menu] li:hover,[data-baseweb=menu] [aria-selected=true]{{background-color:#EEF4FF!important}}
[data-testid=stFileUploaderDropzone]{{background:#F4F7FA!important;border:1px dashed {B}!important;border-radius:8px}}
[data-testid=stFileUploaderDropzone] *,[data-testid=stFileUploader] small,[data-testid=stFileUploaderFile] *{{color:{D}!important}}
[data-testid=stFileUploaderDropzone] button{{background:#fff!important;border:1px solid {B}!important}}
[data-testid=stFileUploaderDropzone] button *{{color:{B}!important}}
[data-testid=stToast],[data-testid=stToast] *{{color:{D}!important}}
</style>
"""


def badge(t, kind=None):
    bg, fg = TINT[kind or KIND.get(t, "b")]
    return f'<span class="badge" style="background:{bg};color:{fg}">{escape(str(t))}</span>'


def chip(t):
    return f'<span class="chip">{escape(str(t))}</span>'


def kpi(label, valor, cor=D, sub=""):
    return f'<div class="kpi"><div class="kl">{label}</div><div class="kv" style="color:{cor}">{valor}</div><div class="ks">{sub}</div></div>'


def alerta(txt, kind):
    bg, fg = TINT[kind]
    return f'<div class="alert" style="background:{bg};border-color:{fg}">{txt}</div>'


def stepper(itens, atual):
    o = ""
    for i, t in enumerate(itens):
        feito, ativo = i < atual, i == atual
        cor = G if feito else B if ativo else LINE
        fg = "#fff" if (feito or ativo) else MUT
        o += f'<div class="st"><i style="background:{cor};color:{fg}">{i + 1}</i><span>{t}</span></div>'
    return f'<div class="steps">{o}</div>'


def tabela(df, badges=(), fmt=None, linha=None, altura=None):
    """Tabela HTML como no protótipo: cabeçalho cinza, linhas finas, selos e linha em destaque."""
    fmt = fmt or {}
    th = "".join(f"<th>{escape(str(c))}</th>" for c in df.columns)
    rows = ""
    for r in df.to_dict("records"):
        bg = ROWBG.get(str(r.get(linha, "")).split("|")[0], "#fff") if linha else "#fff"
        tds = ""
        for c in df.columns:
            v = r[c]
            vazio = v is None or (not isinstance(v, (str, list)) and pd.isna(v)) or v == ""
            if c in badges:
                tds += "<td>" + ("—" if vazio else "".join(badge(p) for p in str(v).split("|"))) + "</td>"
            elif vazio:
                tds += "<td>—</td>"
            elif c in fmt and not isinstance(v, str):
                tds += f"<td>{escape(fmt[c].format(v))}</td>"
            else:
                tds += f"<td>{escape(str(v))}</td>"
        rows += f'<tr style="background:{bg}">{tds}</tr>'
    h = f"max-height:{altura}px;" if altura else ""
    return f'<div class="tbw" style="{h}"><table class="tb"><tr>{th}</tr>{rows}</table></div>'


def plotly_estilo(fig, altura=300):
    """Gráfico sem modebar, sem zoom/arraste e sem ocultar séries ao clicar na legenda."""
    fig.update_layout(height=altura, margin=dict(t=20, b=10, l=10, r=10), paper_bgcolor="rgba(0,0,0,0)",
                      plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter, Arial", color=D, size=13),
                      legend=dict(orientation="h", y=-0.2, itemclick=False, itemdoubleclick=False), dragmode=False,
                      hoverlabel=dict(bgcolor="#fff", font=dict(color=D)))
    fig.update_xaxes(gridcolor="#EEF1F6", fixedrange=True, tickfont=dict(color=D), title_font=dict(color=MUT))
    fig.update_yaxes(gridcolor="#EEF1F6", fixedrange=True, tickfont=dict(color=D), title_font=dict(color=MUT))
    return fig

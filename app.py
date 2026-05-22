# app.py
# BabeliUM — Editor Unificado (Interface Contínua & Parser Estável)

import streamlit as st
import pandas as pd
import copy
import sys
import os
import re
import time
import string
import docx

# --- GARANTIR QUE OS IMPORTS FUNCIONAM ---
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

try:
    from models import TA, Question, ChoiceOption, Blank, MatchPair
    from utils import new_id, count_gaps
    from validators import validate_ficha, update_ficha_status
    from export import build_moodle_xml_stub
except ImportError as e:
    st.error(f"Erro crítico: {e}. Verifique se models.py, utils.py, validators.py e export.py estão na pasta.")
    st.stop()

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="BabeliUM Editor", page_icon="📝", layout="wide")

st.markdown("""
    <style>
    div[data-testid="stExpander"], div[data-testid="stMetric"], div[data-testid="stForm"] {
        background-color: var(--secondary-background-color);
        border-radius: 8px; border: 1px solid rgba(128, 128, 128, 0.2); padding: 5px;
    }
    .stButton>button { border-radius: 6px; font-weight: 500; }
    img { border-radius: 0px !important; background-color: transparent !important; }
    div[data-testid="stDataFrameResizable"] { border-radius: 8px; overflow: hidden; }
    </style>
    """, unsafe_allow_html=True)

# --- CONSTANTES ---
UI_TYPES = {
    "Texto com lacunas (Escrever)": "cloze",
    "Texto com lacunas (Menu/Seleção)": "cloze_mc",
    "Escolha múltipla (1 correta)": "multichoice_single",
    "Escolha múltipla (várias corretas)": "multichoice_multi",
    "Verdadeiro/Falso": "truefalse",
    "Associação (Matching)": "matching",
    "Resposta Curta": "shortanswer",
    "Ensaio (Texto livre)": "essay",
    "Texto de Apoio (Instrução)": "description"
}

NIVEL_OPTIONS  = ["A1", "A2", "B1", "B2", "C1", "C2"]

if "ta" not in st.session_state: st.session_state.ta = TA(ta_id=new_id("ta"))
if "active_view" not in st.session_state: st.session_state.active_view = "Dashboard"
if "active_qid" not in st.session_state: st.session_state.active_qid = None
if "nivel_global" not in st.session_state: st.session_state.nivel_global = "A2"
ta = st.session_state.ta

def delete_question(idx):
    ta.questions.pop(idx)

# ==============================================================================
# MOTOR DE IMPORTAÇÃO (PARSER ESTÁVEL E SEGURO DO APP_V1.2)
# ==============================================================================
def _bold_text(paragraph) -> str:
    return " ".join(r.text.strip() for r in paragraph.runs if r.bold and r.text.strip())

def _iter_block_items(doc):
    import docx.text.paragraph as _para
    import docx.table as _tbl
    for child in doc.element.body:
        tag = child.tag.split('}')[-1]
        if tag == 'p': yield ('paragraph', _para.Paragraph(child, doc))
        elif tag == 'tbl': yield ('table', _tbl.Table(child, doc))

def _normalizar_marcador_correto(paragraph) -> str:
    _MARCADORES = {"✓", "✔", "☑"}
    partes = []
    for run in paragraph.runs:
        texto_run = run.text
        for m in _MARCADORES: texto_run = texto_run.replace(m, "*")
        partes.append(texto_run)
    return "".join(partes).strip()

def _extrair_pares_tabela(table) -> list:
    pares = []
    for i, row in enumerate(table.rows):
        cells = [c.text.strip() for c in row.cells]
        if not cells or not cells[0]: continue
        if i == 0 and any(h in cells[0].lower() for h in ["coluna", "column"]): continue
        if len(cells) >= 2:
            esq = re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', cells[0]).strip()
            dir_ = re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', cells[1]).strip()
            if esq: pares.append({"esquerda": esq, "direita": dir_})
    return pares

def _detetar_tipo(texto_bold: str, texto: str) -> str:
    t = (texto_bold + " " + texto).lower()
    if any(k in t for k in ["complete", "preencha", "lacunas"]): return "Lacunas"
    if any(k in t for k in ["escolha", "assinale", "selecione", "opção"]): return "Escolha Múltipla"
    if any(k in t for k in ["corresponder", "associe", "coluna"]): return "Associação"
    if any(k in t for k in ["identifique e corrija", "erros", "ensaio"]): return "Ensaio"
    return "Desconhecido"

def _e_instrucao(txt: str) -> bool:
    return any(txt.startswith(i) for i in ["⚙", "[[", "🚀", "🔵", "🟡", "🟢", "📝", "[EXEMPLO", "instrução"])

def extrair_perguntas_docx(uploaded_file):
    doc = docx.Document(uploaded_file)
    estado, buffer_enunciado, buffer_pares, secao_atual = None, "", [], "Geral"
    novas_perguntas = 0

    def guardar():
        nonlocal buffer_enunciado, buffer_pares, estado, secao_atual, novas_perguntas
        if not estado: return
        enunciado = buffer_enunciado.strip()
        if not enunciado and not buffer_pares: return
        
        q = Question(qid=new_id("q"), ui_type="A definir", moodle_type="unknown", prompt=enunciado, section=secao_atual)
        
        if estado == "Associação":
            q.ui_type = "Associação (Matching)"
            q.moodle_type = "matching"
            q.pairs = [MatchPair(new_id("p"), p["esquerda"], p["direita"]) for p in buffer_pares]
            q.prompt = "Faça a correspondência correta:" if not q.prompt else q.prompt
            
        elif estado == "Lacunas":
            q.ui_type = "Texto com lacunas (Escrever)"
            q.moodle_type = "cloze"
            q.prompt = re.sub(r'_{3,}', '[ ]', q.prompt) 
            
        elif estado == "Ensaio":
            q.ui_type = "Ensaio (Texto livre)"
            q.moodle_type = "essay"
            
        elif estado == "Escolha Múltipla":
            q.ui_type = "Escolha múltipla (1 correta)"
            q.moodle_type = "multichoice_single"
            
            linhas = buffer_enunciado.strip().split('\n')
            enunciado_limpo, opcoes_extraidas = [], []
            for linha in linhas:
                linha_limpa = linha.strip()
                match = re.match(r'^([a-eA-E][\.\)])\s*(.*)', linha_limpa)
                if match:
                    texto_opcao = match.group(2)
                    is_correct = '*' in texto_opcao
                    texto_limpo = texto_opcao.replace('*', '').strip()
                    opcoes_extraidas.append(ChoiceOption(new_id("o"), texto_limpo, is_correct))
                else:
                    if linha_limpa: enunciado_limpo.append(linha_limpa)
            
            q.prompt = "\n".join(enunciado_limpo).strip()
            if opcoes_extraidas: q.options = opcoes_extraidas
            else: q.options = [ChoiceOption(new_id("o"), ""), ChoiceOption(new_id("o"), "")]
                
        else:
            q.ui_type = "Texto de Apoio (Instrução)"
            q.moodle_type = "description"

        ta.questions.append(q)
        novas_perguntas += 1
        buffer_enunciado, buffer_pares = "", []

    for kind, obj in _iter_block_items(doc):
        if kind == 'paragraph':
            txt = _normalizar_marcador_correto(obj)
            if not txt or _e_instrucao(txt.lower()): continue
            if "chave de respostas" in txt.lower(): guardar(); break

            bold = _bold_text(obj)
            if bold and txt == bold and len(txt) < 50 and not any(c.isdigit() for c in txt):
                guardar()
                estado = None
                secao_atual = txt
                continue

            tipo_det = _detetar_tipo(bold, txt)
            if tipo_det != "Desconhecido":
                guardar()
                estado = tipo_det
                buffer_enunciado = txt + "\n"
            elif estado:
                buffer_enunciado += txt + "\n"

        elif kind == 'table':
            texto_tab = " ".join(c.text.lower() for row in obj.rows for c in row.cells)
            if "instrução:" in texto_tab or "⚙" in texto_tab: continue
            
            if estado == "Associação" or estado is None:
                if estado != "Associação":
                    guardar()
                    estado = "Associação"
                buffer_pares.extend(_extrair_pares_tabela(obj))
            elif estado:
                for row in obj.rows:
                    linha_tab = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
                    if linha_tab: buffer_enunciado += linha_tab + "\n"

    guardar()
    return novas_perguntas

# ==============================================================================
# BARRA LATERAL UNIFICADA (SIDEBAR)
# ==============================================================================
with st.sidebar:
    pasta_atual = os.path.dirname(os.path.abspath(__file__))
    cam_uminho = os.path.join(pasta_atual, "img", "Logo_uminho.png")
    cam_babelium = os.path.join(pasta_atual, "img", "Logo_babeliUM.jpg")

    if os.path.exists(cam_uminho): st.image(cam_uminho, width=180) 
    else: st.header("UMinho")
    
    st.write("") 
    c_txt, c_img = st.columns([2, 1])
    with c_txt: st.markdown("<h3 style='margin:0; padding-top: 15px; text-align: right; font-size: 1.2rem;'>BabeliUM Editor</h3>", unsafe_allow_html=True)
    with c_img:
        if os.path.exists(cam_babelium): st.image(cam_babelium, width=70)
        else: st.write("📝")

    st.markdown("---")

    def _nav_btn(label, view):
        tipo = "primary" if st.session_state.active_view == view else "secondary"
        if st.button(label, type=tipo, use_container_width=True):
            st.session_state.active_view = view
            st.rerun()

    st.markdown("**1. CRIAR CONTEÚDO**")
    _nav_btn("📥 Importar Ficha Word", "Importar")
    _nav_btn("✨ Gerador IA (Brevemente)", "IA")

    st.markdown("---")
    st.markdown("**2. GESTÃO E EXPORTAÇÃO**")
    _nav_btn("🛒 Painel da Ficha", "Dashboard")
    _nav_btn("📦 Exportar Moodle XML", "Exportar")

    st.markdown("---")
    c1, c2 = st.columns(2)
    c1.metric("Questões", len(ta.questions))
    pts = sum(q.meta.points for q in ta.questions if q.meta.points)
    c2.metric("Pontos", f"{pts:g}")

# ==============================================================================
# VIEW 1: IMPORTAR WORD
# ==============================================================================
def render_importar():
    st.title("📥 Importar Documento Word")
    st.info("Utilize o template oficial. Garanta que os Títulos e Enunciados estão em **Negrito**.")

    with st.container(border=True):
        st.session_state.nivel_global = st.selectbox("Nível QECR a aplicar às perguntas:", NIVEL_OPTIONS, index=1)
        uploaded_file = st.file_uploader("Arraste o ficheiro .docx para aqui", type=["docx"], label_visibility="collapsed")
        
        if uploaded_file is not None:
            if st.button("🔄 Processar e Adicionar à Ficha", type="primary", use_container_width=True):
                with st.spinner("A analisar o documento..."):
                    novas = extrair_perguntas_docx(uploaded_file)
                if novas > 0:
                    st.success(f"✅ {novas} perguntas adicionadas com sucesso à base de dados!")
                    time.sleep(1)
                    st.session_state.active_view = "Dashboard"
                    st.rerun()
                else: st.error("Nenhuma pergunta encontrada.")

# ==============================================================================
# VIEW 2: DASHBOARD (PAINEL DE EDIÇÃO)
# ==============================================================================
def render_dashboard():
    st.title("🛒 Painel da Ficha")
    
    with st.container(border=True):
        c1, c2 = st.columns([3, 2])
        ta.ta_name = c1.text_input("Nome da Ficha", value=ta.ta_name, placeholder="Ex: Ficha de Avaliação 1")
        ta.course = c2.text_input("Curso / Nível", value=ta.course, placeholder="Ex: PLE A2")

    if not ta.questions:
        st.write("")
        st.info("A ficha está vazia. Importe um documento Word no menu lateral ou crie uma pergunta manual.")
        if st.button("➕ Criar Pergunta Manual", type="primary"):
            st.session_state.active_qid = None
            st.session_state.active_view = "Editor"
            st.rerun()
        return

    st.subheader(f"Questões Preparadas ({len(ta.questions)})")
    st.write("Verifique a tabela. Pode apagar linhas indesejadas diretamente.")

    # Construir Tabela
    tabela_dados = []
    for i, q in enumerate(ta.questions):
        tabela_dados.append({
            "_idx": i,
            "ID": f"Q{i+1:02d}",
            "Tipo": q.ui_type,
            "Secção": q.section,
            "Enunciado": q.prompt
        })
    
    df = pd.DataFrame(tabela_dados)
    
    edited_df = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "_idx": None,
            "ID": st.column_config.TextColumn("ID", disabled=True, width="small"),
            "Tipo": st.column_config.TextColumn("Tipo", disabled=True, width="medium"),
            "Secção": st.column_config.TextColumn("Secção", width="small"),
            "Enunciado": st.column_config.TextColumn("Enunciado", width="large")
        }
    )

    # Sincronizar Edições e Remoções
    if len(edited_df) < len(df):
        indices_restantes = edited_df["_idx"].tolist()
        ta.questions = [ta.questions[i] for i in indices_restantes]
        st.rerun()
    else:
        for _, row in edited_df.iterrows():
            idx = int(row["_idx"])
            ta.questions[idx].section = row["Secção"]
            ta.questions[idx].prompt = row["Enunciado"]

    st.markdown("---")
    c1, c2 = st.columns([3, 1])
    opcoes_dropdown = {f"Q{i+1:02d} - {q.ui_type}": i for i, q in enumerate(ta.questions)}
    pergunta_selecionada = c1.selectbox("Selecione uma pergunta para definir respostas:", list(opcoes_dropdown.keys()))
    
    if c2.button("✏️ Abrir Editor Avançado", use_container_width=True):
        idx_real = opcoes_dropdown[pergunta_selecionada]
        st.session_state.active_qid = ta.questions[idx_real].qid
        st.session_state.active_view = "Editor"
        st.rerun()

# ==============================================================================
# VIEW 3: EDITOR AVANÇADO DE QUESTÃO
# ==============================================================================
def render_editor():
    if "draft_q" not in st.session_state:
        if st.session_state.active_qid:
            original = get_question_by_id(st.session_state.active_qid)
            st.session_state.draft_q = copy.deepcopy(original)
        else:
            st.session_state.draft_q = Question(qid=new_id("q"), ui_type="Escolha múltipla (1 correta)", moodle_type="multichoice_single", prompt="")
    q = st.session_state.draft_q

    c_back, c_tit = st.columns([1, 6])
    if c_back.button("🔙 Voltar ao Painel"):
        del st.session_state.draft_q
        st.session_state.active_view = "Dashboard"
        st.rerun()
    c_tit.subheader("✏️ Editor de Questão")

    with st.container(border=True):
        c1, c2, c3 = st.columns([2, 1, 1])
        new_type = c1.selectbox("Tipo", list(UI_TYPES.keys()), index=list(UI_TYPES.keys()).index(q.ui_type) if q.ui_type in UI_TYPES else 0)
        
        if new_type != q.ui_type:
            q.ui_type = new_type
            q.moodle_type = UI_TYPES[new_type]
            if q.moodle_type == "truefalse": q.options = [ChoiceOption(new_id("o"), "Verdadeiro", True), ChoiceOption(new_id("o"), "Falso", False)]
            elif "multichoice" in q.moodle_type: q.options = [ChoiceOption(new_id("o"), ""), ChoiceOption(new_id("o"), "")]
            elif q.moodle_type == "shortanswer": q.options = [ChoiceOption(new_id("o"), "", True)]
            st.rerun()

        if q.moodle_type == "description": c2.text_input("Pontos", "0.0", disabled=True)
        else: q.meta.points = c2.number_input("Pontos", value=float(q.meta.points), min_value=0.1, step=0.5, key=f"pts_{q.qid}")
        q.section = c3.text_input("Secção", q.section)

    mt = q.moodle_type
    
    st.markdown("##### 1. Enunciado")
    if mt in ["cloze", "cloze_mc"]:
        cb, ch = st.columns([1, 3])
        if cb.button("➕ Inserir [ ]", help="Adiciona uma lacuna", use_container_width=True): 
            q.prompt += " [ ] "
            st.session_state[f"prompt_{q.qid}"] = q.prompt
            st.rerun()

    q.prompt = st.text_area("Texto", value=q.prompt, height=150, label_visibility="collapsed", key=f"prompt_{q.qid}")

    st.markdown("##### 2. Respostas e Alíneas")
    if mt in ["cloze", "cloze_mc"]:
        n = count_gaps(q.prompt)
        if n == 0: st.warning("⚠️ Insira `[ ]` no texto.")
        else:
            while len(q.blanks) < n: q.blanks.append(Blank(new_id("b"), f"L{len(q.blanks)+1}", [""], []))
            q.blanks = q.blanks[:n]
            cols = st.columns(2 if mt == "cloze_mc" else 3)
            for i, b in enumerate(q.blanks):
                with cols[i%len(cols)]:
                    with st.container(border=True):
                        st.markdown(f"**Lacuna {i+1}**")
                        b.answers[0] = st.text_input("Correta", b.answers[0], key=f"a_{b.bid}")
                        if mt == "cloze_mc":
                            ds = st.text_input("Erradas (sep. ;)", "; ".join(b.distractors), key=f"d_{b.bid}")
                            b.distractors = [x.strip() for x in ds.split(";") if x.strip()]

    elif mt.startswith("multichoice"):
        for i, o in enumerate(q.options):
            c1, c2, c3 = st.columns([0.5, 4, 1])
            if c1.button("🗑️", key=f"del_{o.oid}"): q.options.pop(i); st.rerun()
            o.text = c2.text_input(f"Op {i+1}", o.text, label_visibility="collapsed", key=f"txt_{o.oid}")
            chk = c3.checkbox("Correta", o.is_correct, key=f"chk_{o.oid}")
            if mt == "multichoice_single":
                if chk and not o.is_correct:
                    o.is_correct = True
                    for other in q.options: 
                        if other.oid != o.oid: other.is_correct = False
                    st.rerun()
                elif not chk: o.is_correct = False
            else: o.is_correct = chk
        if st.button("➕ Opção"): q.options.append(ChoiceOption(new_id("o"), "")); st.rerun()

    elif mt == "matching":
        for i, p in enumerate(q.pairs):
            c1, c2, c3 = st.columns([0.5, 2.5, 2.5])
            if c1.button("🗑️", key=f"dmat_{p.pid}"): q.pairs.pop(i); st.rerun()
            p.left = c2.text_input("A", p.left, label_visibility="collapsed", key=f"pl_{p.pid}")
            p.right = c3.text_input("B", p.right, label_visibility="collapsed", key=f"pr_{p.pid}")
        if st.button("➕ Par"): q.pairs.append(MatchPair(new_id("p"), "", "")); st.rerun()

    st.write("")
    if st.button("💾 Guardar e Voltar", type="primary", use_container_width=True):
        if st.session_state.active_qid:
            for i, ex in enumerate(ta.questions):
                if ex.qid == st.session_state.active_qid: ta.questions[i] = copy.deepcopy(q); break
        else: ta.questions.append(copy.deepcopy(q))
        del st.session_state.draft_q
        st.session_state.active_view = "Dashboard"
        st.session_state.active_qid = None
        st.rerun()

# ==============================================================================
# VIEW 4: EXPORTAR E IA
# ==============================================================================
def render_export_view():
    st.header("📦 Exportar para o Moodle")
    issues = validate_ficha(ta)
    update_ficha_status(ta, issues)
    
    ok = not any(i.level == "ERRO" for i in issues)
    if ok: st.success("A sua ficha está validada e pronta para exportar!")
    else: st.error("Tem de resolver os erros abaixo antes de exportar.")

    for i in issues:
        c = "red" if i.level == "ERRO" else "orange"
        st.markdown(f":{c}[**{i.level}**] {i.message} ({i.where})")

    if ok:
        xml = build_moodle_xml_stub(ta)
        st.download_button("📥 Descarregar Moodle XML", xml, "ficha.xml", "application/xml", type="primary")

def render_ia():
    st.title("✨ Assistente IA (Brevemente)")
    st.info("A ligação direta à OpenAI para geração automática de exercícios está em desenvolvimento.")

# ==============================================================================
# ROTAS
# ==============================================================================
VIEW = st.session_state.active_view
if VIEW == "Dashboard": render_dashboard()
elif VIEW == "Importar": render_importar()
elif VIEW == "Editor": render_editor()
elif VIEW == "Exportar": render_export_view()
elif VIEW == "IA": render_ia()
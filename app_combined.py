# app_combined.py
# BabeliUM — Aplicação Unificada
# Combina o Editor de Fichas (app.py) com o Importador & Painel de Triagem (app_v1.py)

import streamlit as st
import pandas as pd
import copy
import time
import re
import os
import sys
import string
import io

# --- GARANTIR QUE OS IMPORTS FUNCIONAM ---
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

try:
    import docx
except ImportError:
    docx = None

try:
    from models import TA, Question, ChoiceOption, Blank, MatchPair, new_id_default
    from utils import new_id, count_gaps
    from validators import validate_ficha, update_ficha_status
    from export import build_moodle_xml_stub
except ImportError as e:
    st.error(f"Erro crítico: {e}. Verifique se models.py, utils.py, validators.py e export.py estão na pasta.")
    st.stop()


# ══════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ══════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="BabeliUM Editor",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ══════════════════════════════════════════════════════════════════
# ESTILOS CSS (COMPATÍVEL COM LIGHT/DARK MODE)
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
div[data-testid="stExpander"],
div[data-testid="stMetric"],
div[data-testid="stForm"] {
    background-color: var(--secondary-background-color);
    border-radius: 8px;
    border: 1px solid rgba(128, 128, 128, 0.2);
    padding: 5px;
}
.stButton>button {
    border-radius: 6px;
    font-weight: 500;
}
img {
    border-radius: 0px !important;
    background-color: transparent !important;
}
div[data-testid="stDataFrameResizable"] {
    border-radius: 8px;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# CONSTANTES — EDITOR
# ══════════════════════════════════════════════════════════════════
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

# ══════════════════════════════════════════════════════════════════
# CONSTANTES — IMPORTADOR / TRIAGEM
# ══════════════════════════════════════════════════════════════════
TIPO_ICONS = {
    "Lacunas (Cloze)":      "📝",
    "Lacunas (Escrita)":    "✍️",
    "Associação":           "🔗",
    "Ensaio (Texto Livre)": "📄",
    "Escolha Múltipla":     "🔘",
    "Desconhecido":         "❓",
}

NIVEL_OPTIONS  = ["A1", "A2", "B1", "B2", "C1", "C2"]
TOPICO_OPTIONS = ["Tema 1", "Tema 2", "Tema 3", "Tema 4", "Tema 5", "Outro"]
LINGUAS        = ["Português (PLE)", "Inglês"]

# ══════════════════════════════════════════════════════════════════
# GESTÃO DE ESTADO
# ══════════════════════════════════════════════════════════════════

# --- Módulo activo: "editor" ou "importador" ---
if "modulo_ativo" not in st.session_state:
    st.session_state.modulo_ativo = "editor"

# --- Estado do Editor ---
if "ta" not in st.session_state:
    st.session_state.ta = TA(ta_id=new_id("ta"))
if "active_view" not in st.session_state:
    st.session_state.active_view = "Dashboard"
if "active_qid" not in st.session_state:
    st.session_state.active_qid = None

# --- Estado do Importador ---
defaults_imp = {
    "imp_view":        "Dashboard",
    "view_details_id": None,
    "perguntas_df":    pd.DataFrame(),
    "nivel_global":    "A2",
    "curso_global":    "PLE",
}
for k, v in defaults_imp.items():
    if k not in st.session_state:
        st.session_state[k] = v

ta = st.session_state.ta


# ══════════════════════════════════════════════════════════════════
# HELPERS — EDITOR
# ══════════════════════════════════════════════════════════════════
def get_question_by_id(qid):
    for q in ta.questions:
        if q.qid == qid:
            return q
    return None

def delete_question(idx):
    ta.questions.pop(idx)


# ══════════════════════════════════════════════════════════════════
# HELPERS — IMPORTADOR
# ══════════════════════════════════════════════════════════════════
def obter_proximo_id() -> int:
    df = st.session_state.perguntas_df
    if df.empty or "ID" not in df.columns:
        return 1
    try:
        nums = df["ID"].str.replace("Q", "", regex=False).astype(int)
        return int(nums.max()) + 1
    except (ValueError, AttributeError):
        return len(df) + 1


# ══════════════════════════════════════════════════════════════════
# PARSER ESTRUTURAL WORD
# ══════════════════════════════════════════════════════════════════
def _bold_text(paragraph) -> str:
    return " ".join(r.text.strip() for r in paragraph.runs if r.bold and r.text.strip())

def _iter_block_items(doc):
    import docx.text.paragraph as _para
    import docx.table as _tbl
    for child in doc.element.body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            yield ('paragraph', _para.Paragraph(child, doc))
        elif tag == 'tbl':
            yield ('table', _tbl.Table(child, doc))

def _extrair_pares_tabela(table) -> list:
    pares = []
    for i, row in enumerate(table.rows):
        cells = [c.text.strip() for c in row.cells]
        if not cells or not cells[0]:
            continue
        if i == 0 and any(h in cells[0].lower() for h in ["coluna", "column"]):
            continue
        if len(cells) >= 2:
            esq_raw = cells[0]
            dir_raw = cells[1] if len(cells) > 1 else ""
            esq_linhas = [l.strip() for l in esq_raw.split("\n") if l.strip()]
            dir_linhas = [l.strip() for l in dir_raw.split("\n") if l.strip()]
            if len(esq_linhas) > 1 or len(dir_linhas) > 1:
                esq_limpa = [re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', l).strip() for l in esq_linhas]
                dir_limpa = [re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', l).strip() for l in dir_linhas]
                for j, esq in enumerate(esq_limpa):
                    if esq:
                        pares.append({"esquerda": esq, "direita": dir_limpa[j] if j < len(dir_limpa) else ""})
                return pares
            esq  = re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', esq_raw).strip()
            dir_ = re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', dir_raw).strip()
            if esq:
                pares.append({"esquerda": esq, "direita": dir_})
    return pares

def _normalizar_marcador_correto(paragraph) -> str:
    _MARCADORES = {"✓", "✔", "☑"}
    partes = []
    for run in paragraph.runs:
        texto_run = run.text
        for m in _MARCADORES:
            texto_run = texto_run.replace(m, "*")
        partes.append(texto_run)
    return "".join(partes).strip()

def _detetar_tipo(texto_bold: str, texto: str) -> str:
    t = (texto_bold + " " + texto).lower()
    if any(k in t for k in ["selecione a opção", "escolha a opção mais adequada"]):
        return "Lacunas (Cloze)"
    if any(k in t for k in ["complete o texto", "complete as frases com palavras"]):
        return "Lacunas (Escrita)"
    if any(k in t for k in ["complete as frases da coluna", "faça corresponder", "associe"]):
        return "Associação"
    if any(k in t for k in ["identifique e corrija", "identifique os erros"]):
        return "Ensaio (Texto Livre)"
    return "Desconhecido"

_MARCADORES_INSTRUCAO = [
    "instrução:", "apague este bloco", "[[", "⚙", "[exemplo", "[template",
    "guia rápido", "🚀", "🔵", "🟡", "🟢", "📝",
]

def _tabela_e_instrucao(table) -> bool:
    texto = " ".join(
        c.text.strip().lower()
        for row in table.rows
        for c in row.cells
    )
    return any(m in texto for m in _MARCADORES_INSTRUCAO)

def _e_titulo_seccao(bold: str, txt: str) -> bool:
    if not bold or txt != bold:
        return False
    if len(txt) > 50:
        return False
    if any(c.isdigit() for c in txt):
        return False
    if re.match(r'^[IVX]+\.', txt.strip()):
        return False
    return True

def extrair_perguntas_docx(uploaded_file, start_id: int = 1) -> pd.DataFrame:
    doc = docx.Document(uploaded_file)
    perguntas = []
    contador = start_id
    estado, buffer_enunciado, buffer_pares, secao_atual = None, "", [], ""

    def guardar():
        nonlocal contador, buffer_enunciado, buffer_pares, estado
        if not estado:
            return
        enunciado = buffer_enunciado.strip()
        if not enunciado and not buffer_pares:
            return
        perguntas.append({
            "Exportar": True,
            "ID":       f"Q{contador:02d}",
            "Secção":   secao_atual,
            "Enunciado": enunciado,
            "Pares":    copy.deepcopy(buffer_pares),
            "Tipo":     estado,
            "Nível":    st.session_state.nivel_global,
            "Tópico":   "Tema 1",
        })
        contador += 1
        buffer_enunciado, buffer_pares = "", []

    for kind, obj in _iter_block_items(doc):
        if kind == 'paragraph':
            txt = _normalizar_marcador_correto(obj)
            if not txt:
                continue
            if "chave de respostas" in txt.lower():
                guardar()
                break
            bold = _bold_text(obj)
            if _e_titulo_seccao(bold, txt):
                guardar()
                estado = None
                secao_atual = txt
                continue
            tipo_det = _detetar_tipo(bold, txt)
            if tipo_det != "Desconhecido":
                primeira_linha_buffer = buffer_enunciado.strip().split("\n")[0].strip().lower()
                if txt.strip().lower() == primeira_linha_buffer:
                    continue
                guardar()
                estado = tipo_det
                buffer_enunciado = txt + "\n"
            elif estado:
                buffer_enunciado += txt + "\n"
        elif kind == 'table':
            if _tabela_e_instrucao(obj):
                continue
            if estado == "Associação" or estado is None:
                if estado != "Associação":
                    guardar()
                    estado = "Associação"
                    buffer_enunciado = f"(Tabela de associação — {secao_atual})\n"
                buffer_pares.extend(_extrair_pares_tabela(obj))
            elif estado:
                for row in obj.rows:
                    linha_tab = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
                    if linha_tab:
                        buffer_enunciado += linha_tab + "\n"

    guardar()
    return pd.DataFrame(perguntas) if perguntas else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════
# EXEMPLOS IA
# ══════════════════════════════════════════════════════════════════
EXEMPLOS_IA = {
    "Português (PLE)": {
        "A1": [
            {"tipo": "Escolha Múltipla", "enunciado": "Qual é a forma correta do verbo 'ser' na 3.ª pessoa do singular?\na) são\nb) somos\nc) é\nd) és"},
            {"tipo": "Lacunas (Escrita)", "enunciado": "Complete as frases com o artigo definido correto (o, a, os, as):\n1. _____ livro está na mesa.\n2. _____ professora chama-se Ana.\n3. _____ alunos são simpáticos."},
        ],
        "A2": [
            {"tipo": "Lacunas (Cloze)", "enunciado": "Escolha a opção correta para completar cada frase:\n1. Ontem eu _______ ao cinema. (fui / vou / irei)\n2. Nós _______ juntos todos os dias. (almoçamos / almoçam / almoçar)\n3. Eles já _______ em Lisboa há cinco anos. (vivem / vivo / vivemos)"},
            {"tipo": "Associação", "enunciado": "Associe cada profissão com a sua descrição.", "pares": [{"esquerda": "Médico", "direita": "Trata de doentes"}, {"esquerda": "Professor", "direita": "Ensina alunos"}, {"esquerda": "Engenheiro", "direita": "Constrói edifícios"}]},
        ],
        "B1": [
            {"tipo": "Ensaio (Texto Livre)", "enunciado": "Descreva um dia típico da sua semana. Inclua as suas rotinas matinais, o seu trabalho ou estudos, e as atividades que faz ao final do dia. (80–100 palavras)"},
        ],
    },
    "Inglês": {
        "A1": [
            {"tipo": "Escolha Múltipla", "enunciado": "Choose the correct form of the verb 'to be':\n1. She _______ a student. (is / are / am)\n2. They _______ from Brazil. (is / are / am)\n3. I _______ happy today. (is / are / am)"},
        ],
        "A2": [
            {"tipo": "Lacunas (Cloze)", "enunciado": "Choose the best option to complete each sentence:\n1. Yesterday, I _______ to the supermarket. (went / go / goes)\n2. She _______ English very well. (speaks / speak / speaking)\n3. We _______ dinner at 7 pm every night. (have / has / having)"},
            {"tipo": "Associação", "enunciado": "Match each word with its correct definition.", "pares": [{"esquerda": "Library", "direita": "A place to borrow books"}, {"esquerda": "Pharmacy", "direita": "A place to buy medicine"}, {"esquerda": "Bakery", "direita": "A place to buy bread"}]},
        ],
        "B1": [
            {"tipo": "Ensaio (Texto Livre)", "enunciado": "Write a short paragraph about your last holiday. Where did you go? Who did you go with? What did you do? (80–100 words)"},
        ],
    },
}

def obter_exemplo_ia(lingua, nivel, topico, tipo_pergunta):
    exemplos_lingua = EXEMPLOS_IA.get(lingua, {})
    exemplos_nivel  = exemplos_lingua.get(nivel, [])
    candidatos = [e for e in exemplos_nivel if e.get("tipo") == tipo_pergunta]
    if not candidatos:
        candidatos = exemplos_nivel
    if candidatos:
        exemplo = copy.deepcopy(candidatos[0])
        if topico and topico.lower() not in exemplo["enunciado"].lower():
            exemplo["enunciado"] = f"[Tópico: {topico}]\n\n" + exemplo["enunciado"]
        return exemplo
    lang_label = "em Português" if "Português" in lingua else "in English"
    return {
        "tipo": tipo_pergunta,
        "enunciado": (
            f"Exemplo de pergunta {lang_label} — Nível {nivel}."
            + (f"\nTópico: {topico}." if topico else "")
            + f"\n\n[Este é um exemplo ilustrativo. A integração com a API gerará conteúdo real.]"
        ),
        "pares": [],
    }


# ══════════════════════════════════════════════════════════════════
# SIDEBAR UNIFICADA
# ══════════════════════════════════════════════════════════════════
with st.sidebar:
    pasta_atual = os.path.dirname(os.path.abspath(__file__))
    cam_uminho  = os.path.join(pasta_atual, "img", "Logo_uminho.png")
    cam_babel   = os.path.join(pasta_atual, "img", "Logo_babeliUM.jpg")

    if os.path.exists(cam_uminho):
        st.image(cam_uminho, width=180)
    else:
        st.header("UMinho")

    st.write("")
    c_txt, c_img = st.columns([2, 1])
    with c_txt:
        st.markdown("<h3 style='margin:0; padding-top:15px; text-align:right; font-size:1.2rem;'>BabeliUM Editor</h3>", unsafe_allow_html=True)
    with c_img:
        if os.path.exists(cam_babel): st.image(cam_babel, width=70)
        else: st.write("📝")

    st.markdown("---")

    # Função inteligente para navegar sem esconder botões
    def _nav_global(label, view, module):
        # Verifica se estamos na view e no módulo certos para pintar de azul
        is_active = (st.session_state.modulo_ativo == module) and (
            (module == "editor" and st.session_state.active_view == view) or 
            (module == "importador" and st.session_state.imp_view == view)
        )
        tipo = "primary" if is_active else "secondary"
        
        if st.button(label, type=tipo, use_container_width=True):
            st.session_state.modulo_ativo = module
            if module == "editor": st.session_state.active_view = view
            else: st.session_state.imp_view = view
            st.rerun()

    st.markdown("**1. CRIAR & IMPORTAR**")
    _nav_global("📝 Criar Ficha Manual", "Dashboard", "editor")
    _nav_global("📥 Importar Word", "Importar_Word", "importador")
    _nav_global("✨ Gerador IA", "Gerador_IA", "importador")

    st.markdown("---")
    st.markdown("**2. GESTÃO & EXPORTAÇÃO**")
    _nav_global("🛒 Painel de Triagem", "Dashboard", "importador")
    _nav_global("📦 Exportar Moodle XML", "Exportar", "editor")

    st.markdown("**3. FERRAMENTAS ADMIN**")
    _nav_global("🔄 Conversor PEAKIT", "Conversor_CSV", "importador")

    st.markdown("---")
    st.caption("Métricas Moodle (Ficha Manual)")
    c1, c2 = st.columns(2)
    c1.metric("Questões", len(ta.questions))
    pts = sum(q.meta.points for q in ta.questions if q.meta.points)
    c2.metric("Pontos", f"{pts:g}")

# ══════════════════════════════════════════════════════════════════
# MÓDULO EDITOR — VIEWS
# ══════════════════════════════════════════════════════════════════

def render_editor_dashboard():
    st.title("Painel de Edição")

    with st.container(border=True):
        c1, c2, c3 = st.columns([3, 2, 2])
        ta.ta_name = c1.text_input("Nome da Ficha", value=ta.ta_name, placeholder="Ex: Ficha 1")
        ta.course  = c2.text_input("Curso / Nível", value=ta.course, placeholder="Ex: PLE A2")
        with c3:
            st.write("")
            st.write("")
            if st.button("➕ Nova Questão", type="primary", use_container_width=True):
                st.session_state.active_qid  = None
                st.session_state.active_view = "Editor"
                st.rerun()

    st.write("")

    if not ta.questions:
        with st.container(border=True):
            st.write("")
            col_center = st.columns([1, 2, 1])[1]
            with col_center:
                st.markdown("<div style='text-align: center; color: #888;'><h3>A ficha está vazia</h3><p>Adicione a sua primeira questão acima.</p></div>", unsafe_allow_html=True)
            st.write("")
    else:
        c_tot, c_view = st.columns([4, 1.5])
        c_tot.subheader(f"Questões ({len(ta.questions)})")
        view_mode = c_view.radio("Ver como:", ["Lista Compacta", "Cartões"], horizontal=True, label_visibility="collapsed")
        st.divider()

        for idx, q in enumerate(ta.questions):
            icon = "❓"
            if "cloze" in q.moodle_type:      icon = "📝"
            elif "multichoice" in q.moodle_type: icon = "🔘"
            elif "truefalse" in q.moodle_type:   icon = "⚖️"
            elif "matching" in q.moodle_type:    icon = "🔗"

            titulo = q.title if q.title else f"Questão {idx+1}"

            if view_mode == "Lista Compacta":
                with st.expander(f"{idx+1}. {icon} {titulo}  |  {q.meta.points} pts"):
                    c_info, c_acts = st.columns([5, 1])
                    with c_info:
                        st.caption(f"{q.ui_type} | {q.section}")
                        st.markdown(f"_{q.prompt[:100]}..._")
                    with c_acts:
                        if st.button("✏️", key=f"le_{q.qid}"):
                            st.session_state.active_qid  = q.qid
                            st.session_state.active_view = "Editor"
                            st.rerun()
                        if st.button("🗑️", key=f"ld_{q.qid}"):
                            delete_question(idx)
                            st.rerun()
            else:
                with st.container(border=True):
                    c_head, c_body, c_acts = st.columns([0.5, 4.5, 1])
                    c_head.markdown(f"### {idx+1}")
                    with c_body:
                        st.markdown(f"**{titulo}** ({q.meta.points} pts)")
                        st.caption(q.ui_type)
                        st.markdown(q.prompt)
                    with c_acts:
                        if st.button("✏️ Editar", key=f"ce_{q.qid}", use_container_width=True):
                            st.session_state.active_qid  = q.qid
                            st.session_state.active_view = "Editor"
                            st.rerun()
                        if st.button("🗑️ Apagar", key=f"cd_{q.qid}", use_container_width=True):
                            delete_question(idx)
                            st.rerun()
# 👇 NOVO CÓDIGO AQUI NO FUNDO DA FUNÇÃO 👇
    st.markdown("---")
    st.markdown("### 📦 Finalizar Ficha Manual")
    if st.button("🚀 Validar e Exportar para Moodle XML", type="primary"):
        st.session_state.modulo_ativo = "editor"
        st.session_state.active_view = "Exportar"
        st.rerun()

def render_editor_questao():
    if "draft_q" not in st.session_state:
        if st.session_state.active_qid:
            original = get_question_by_id(st.session_state.active_qid)
            st.session_state.draft_q = copy.deepcopy(original)
        else:
            st.session_state.draft_q = Question(
                qid=new_id("q"),
                ui_type="Escolha múltipla (1 correta)",
                moodle_type="multichoice_single",
                prompt=""
            )
    q = st.session_state.draft_q

    c_back, c_tit = st.columns([1, 6])
    if c_back.button("🔙 Voltar"):
        del st.session_state.draft_q
        st.session_state.active_view = "Dashboard"
        st.rerun()
    c_tit.subheader("✏️ Editor de Questão")

    with st.container(border=True):
        c1, c2, c3 = st.columns([2, 1, 1])
        new_type = c1.selectbox(
            "Tipo", list(UI_TYPES.keys()),
            index=list(UI_TYPES.keys()).index(q.ui_type) if q.ui_type in UI_TYPES else 0
        )
        if new_type != q.ui_type:
            q.ui_type    = new_type
            q.moodle_type = UI_TYPES[new_type]
            if q.moodle_type == "truefalse":
                q.options = [ChoiceOption(new_id("o"), "Verdadeiro", True), ChoiceOption(new_id("o"), "Falso", False)]
            elif "multichoice" in q.moodle_type:
                q.options = [ChoiceOption(new_id("o"), ""), ChoiceOption(new_id("o"), "")]
            elif q.moodle_type == "shortanswer":
                q.options = [ChoiceOption(new_id("o"), "", True)]
            st.rerun()

        if q.moodle_type == "description":
            c2.text_input("Pontos", "0.0", disabled=True)
        else:
            q.meta.points = c2.number_input(
                "Pontos", value=float(q.meta.points),
                min_value=0.1, step=0.5, key=f"pts_{q.qid}"
            )
        q.title = c3.text_input("Título Interno", q.title)

    mt = q.moodle_type
    st.markdown("##### 1. Enunciado")

    if mt in ["cloze", "cloze_mc"]:
        cb, ch = st.columns([1, 3])
        if cb.button("➕ Inserir [ ]", help="Adiciona uma lacuna ao texto", use_container_width=True):
            q.prompt += " [ ] "
            st.session_state[f"prompt_{q.qid}"] = q.prompt
            st.rerun()
        with ch.expander("💡 Como funciona este tipo de pergunta?"):
            if mt == "cloze":
                st.markdown("**Lacunas de Escrever (Teclado)**\nO aluno terá de escrever a palavra no espaço vazio.\n\n**Exemplo:**\n> Ontem, eu [ ] (ir) ao cinema.\n\n*Nas 'Respostas' em baixo, indique apenas a palavra correta (ex: fui).*")
                st.info("⚠️ Se quer dar opções aos alunos (ex: a, b, c), mude o Tipo de Pergunta para **Escolha Múltipla** ou **Lacunas (Menu)**.")
            elif mt == "cloze_mc":
                st.markdown("**Lacunas de Menu (Dropdown)**\nO aluno clica no espaço vazio e abre-se uma lista de opções.\n\n**Exemplo:**\n> O céu é [ ] e a erva é verde.\n\n*Nas 'Respostas' em baixo, coloque a opção correta (ex: azul) e as erradas separadas por ponto e vírgula (ex: vermelho; amarelo).*")
                st.info("⚠️ Se prefere o formato clássico de teste com as alíneas listadas debaixo do texto, mude o Tipo de Pergunta no topo para **Escolha Múltipla**.")

    q.prompt = st.text_area(
        "Texto", value=q.prompt, height=150,
        label_visibility="collapsed", key=f"prompt_{q.qid}"
    )

    st.markdown("##### 2. Respostas")

    if mt in ["cloze", "cloze_mc"]:
        n = count_gaps(q.prompt)
        if n == 0:
            st.warning("⚠️ Insira `[ ]` no texto.")
        else:
            while len(q.blanks) < n:
                q.blanks.append(Blank(new_id("b"), f"L{len(q.blanks)+1}", [""], []))
            q.blanks = q.blanks[:n]
            cols = st.columns(2 if mt == "cloze_mc" else 3)
            for i, b in enumerate(q.blanks):
                with cols[i % len(cols)]:
                    with st.container(border=True):
                        st.markdown(f"**Lacuna {i+1}**")
                        b.answers[0] = st.text_input("Correta", b.answers[0], key=f"a_{b.bid}")
                        if mt == "cloze_mc":
                            ds = st.text_input("Erradas (sep. ;)", "; ".join(b.distractors), key=f"d_{b.bid}")
                            b.distractors = [x.strip() for x in ds.split(";") if x.strip()]

    elif mt.startswith("multichoice"):
        for i, o in enumerate(q.options):
            c1, c2, c3 = st.columns([0.5, 4, 1])
            if c1.button("🗑️", key=f"del_{o.oid}"):
                q.options.pop(i); st.rerun()
            o.text = c2.text_input(f"Op {i+1}", o.text, label_visibility="collapsed", key=f"txt_{o.oid}")
            chk = c3.checkbox("Correta", o.is_correct, key=f"chk_{o.oid}")
            if mt == "multichoice_single":
                if chk and not o.is_correct:
                    o.is_correct = True
                    for other in q.options:
                        if other.oid != o.oid: other.is_correct = False
                    st.rerun()
                elif not chk:
                    o.is_correct = False
            else:
                o.is_correct = chk
        if st.button("➕ Opção"):
            q.options.append(ChoiceOption(new_id("o"), "")); st.rerun()

    elif mt == "truefalse":
        q.tf_require_correction = st.toggle("Pedir correção?", q.tf_require_correction)
        for i, o in enumerate(q.options):
            with st.container(border=True):
                c1, c2, c3 = st.columns([0.5, 4, 2])
                if c1.button("🗑️", key=f"dvf_{o.oid}"):
                    q.options.pop(i); st.rerun()
                o.text = c2.text_input("Frase", o.text, label_visibility="collapsed", key=f"tvf_{o.oid}")
                sel = c3.radio("R", ["V", "F"], 0 if o.is_correct else 1, horizontal=True, label_visibility="collapsed", key=f"rvf_{o.oid}")
                o.is_correct = (sel == "V")
        if st.button("➕ Frase"):
            q.options.append(ChoiceOption(new_id("o"), "", True)); st.rerun()

    elif mt == "matching":
        for i, p in enumerate(q.pairs):
            c1, c2, c3 = st.columns([0.5, 2.5, 2.5])
            if c1.button("🗑️", key=f"dmat_{p.pid}"):
                q.pairs.pop(i); st.rerun()
            p.left  = c2.text_input("A", p.left,  label_visibility="collapsed", key=f"pl_{p.pid}")
            p.right = c3.text_input("B", p.right, label_visibility="collapsed", key=f"pr_{p.pid}")
        if st.button("➕ Par"):
            q.pairs.append(MatchPair(new_id("p"), "", "")); st.rerun()

    elif mt == "shortanswer":
        st.info("O aluno terá de escrever uma resposta curta. O Moodle fará a correção automática, por isso deve prever a(s) resposta(s) exata(s).")
        st.markdown("**Respostas Aceites como Corretas (100%):**")
        for i, o in enumerate(q.options):
            c1, c2 = st.columns([0.5, 5])
            if c1.button("🗑️", key=f"dsa_{o.oid}"):
                q.options.pop(i); st.rerun()
            o.text = c2.text_input(f"Resposta {i+1}", o.text, label_visibility="collapsed", key=f"tsa_{o.oid}", placeholder="Ex: Lisboa")
            o.is_correct = True
        if st.button("➕ Adicionar outra variante aceite"):
            q.options.append(ChoiceOption(new_id("o"), "", True)); st.rerun()

    # Pré-visualização
    st.divider()
    with st.expander("👁️ Pré-visualização (Clique para expandir/recolher)", expanded=False):
        if mt == "cloze":
            st.markdown(q.prompt.replace("[ ]", " `[ ____ ]` "))
        elif mt == "cloze_mc":
            st.markdown(q.prompt.replace("[ ]", " `[ ▼ ]` "))
            if q.blanks:
                st.markdown("---")
                for i, b in enumerate(q.blanks):
                    st.markdown(f"**Opções da Lacuna {i+1}:**")
                    opcoes = []
                    if b.answers and b.answers[0].strip():
                        opcoes.append(b.answers[0].strip())
                    opcoes.extend([d.strip() for d in b.distractors if d.strip()])
                    opcoes = sorted(list(set(opcoes)))
                    for j, opt in enumerate(opcoes):
                        letra = string.ascii_lowercase[j % 26]
                        st.markdown(f"&nbsp;&nbsp;&nbsp; {letra}) {opt}")
        elif "multichoice" in mt:
            st.write(q.prompt)
            for o in q.options: st.write(f"⚪ {o.text}")
        elif mt == "truefalse":
            st.write(q.prompt)
            st.write("---")
            for o in q.options: st.write(f"- {o.text} (V/F)")
        elif mt == "matching":
            st.write(q.prompt)
            st.write("---")
            ca, cb = st.columns(2)
            ca.markdown("**Coluna A (Perguntas)**")
            cb.markdown("**Coluna B (Opções)**")
            opcoes_b = []
            for p in q.pairs:
                ca.markdown(f"- {p.left}")
                if p.right.strip():
                    opcoes_b.append(p.right.strip())
            for opt in sorted(list(set(opcoes_b))):
                cb.markdown(f"🔹 {opt}")
        elif mt == "shortanswer":
            st.write(q.prompt)
            st.markdown("---")
            st.markdown("**O aluno verá uma caixa de texto vazia. O sistema vai considerar CERTO se ele escrever:**")
            for o in q.options:
                if o.text.strip():
                    st.markdown(f"✔️ `{o.text}`")
        else:
            st.write(q.prompt)

    st.write("")
    cs, cn = st.columns(2)
    if cs.button("💾 Guardar e Sair", type="primary", use_container_width=True):
        _editor_save(q)
        del st.session_state.draft_q
        st.session_state.active_view = "Dashboard"
        st.session_state.active_qid  = None
        st.rerun()
    if cn.button("⏩ Guardar e Seguir", use_container_width=True):
        _editor_save(q)
        next_q = Question(
            qid=new_id("q"), ui_type=q.ui_type, moodle_type=q.moodle_type,
            prompt="", section=q.section, meta=copy.deepcopy(q.meta)
        )
        if "multichoice" in q.moodle_type:
            next_q.options = [ChoiceOption(new_id("o"), ""), ChoiceOption(new_id("o"), "")]
        elif q.moodle_type == "truefalse":
            next_q.options = [ChoiceOption(new_id("o"), "Verdadeiro", True), ChoiceOption(new_id("o"), "Falso", False)]
        elif q.moodle_type == "matching":
            next_q.pairs = [MatchPair(new_id("p"), "", "")]
        st.session_state.draft_q     = next_q
        st.session_state.active_qid  = None
        st.rerun()


def _editor_save(q):
    if st.session_state.active_qid:
        for i, ex in enumerate(ta.questions):
            if ex.qid == st.session_state.active_qid:
                ta.questions[i] = copy.deepcopy(q)
                break
    else:
        ta.questions.append(copy.deepcopy(q))


def render_editor_exportar():
    # 👇 NOVO CÓDIGO: Botão de voltar 👇
    c_back, c_tit = st.columns([1, 6])
    if c_back.button("🔙 Voltar ao Painel"):
        st.session_state.active_view = "Dashboard"
        st.rerun()

    c_tit.header("📦 Exportar Ficha")

    issues = validate_ficha(ta)
    # ... (o resto da função mantém-se intocado)
    issues = validate_ficha(ta)
    update_ficha_status(ta, issues)

    ok = not any(i.level == "ERRO" for i in issues)
    if ok:
        st.success("Ficha pronta!")
    else:
        st.error("Resolva os erros abaixo.")

    for i in issues:
        c = "red" if i.level == "ERRO" else "orange"
        st.markdown(f":{c}[**{i.level}**] {i.message} ({i.where})")

    if ok:
        xml = build_moodle_xml_stub(ta)
        st.download_button("📥 Descarregar XML", xml, "ficha.xml", "application/xml", type="primary")


# ══════════════════════════════════════════════════════════════════
# MÓDULO IMPORTADOR — VIEWS
# ══════════════════════════════════════════════════════════════════

def render_imp_dashboard():
    st.title("🛒 Painel de Triagem")
    df_atual = st.session_state.perguntas_df

    if df_atual.empty:
        with st.container(border=True):
            st.write("")
            col_center = st.columns([1, 2, 1])[1]
            with col_center:
                st.markdown(
                    "<div style='text-align:center; color:#888;'>"
                    "<h3>O carrinho está vazio</h3>"
                    "<p>Use o menu à esquerda para <b>Importar Word</b> ou usar o <b>Assistente IA</b>.</p>"
                    "</div>",
                    unsafe_allow_html=True
                )
            st.write("")
        return

    c_tot, c_view = st.columns([4, 1.5])
    c_tot.subheader(f"Questões no carrinho ({len(df_atual)})")
    view_mode = c_view.radio("Modo:", ["Tabela", "Cartões"], horizontal=True, label_visibility="collapsed")

    st.markdown(
        "Edite **Tipo**, **Nível** e **Tópico** diretamente na tabela. "
        "Selecione linhas e prima `Delete` para eliminar. "
        "Use **Edição Profunda** para editar o enunciado completo."
    )
    st.divider()

    if view_mode == "Tabela":
        cols_visiveis = [c for c in ["Exportar", "ID", "Secção", "Tipo", "Nível", "Tópico", "Enunciado"] if c in df_atual.columns]
        edited = st.data_editor(
            df_atual[cols_visiveis],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Exportar":  st.column_config.CheckboxColumn("Exportar?", default=True, width="small"),
                "ID":        st.column_config.TextColumn("ID", disabled=True, width="small"),
                "Secção":    st.column_config.TextColumn("Secção", disabled=True, width="medium"),
                "Tipo":      st.column_config.SelectboxColumn("Tipo", options=list(TIPO_ICONS.keys()), width="medium"),
                "Nível":     st.column_config.SelectboxColumn("Nível", options=NIVEL_OPTIONS, width="small"),
                "Tópico":    st.column_config.SelectboxColumn("Tópico", options=TOPICO_OPTIONS, width="small"),
                "Enunciado": st.column_config.TextColumn("Enunciado", width="large"),
            },
            key="tabela_carrinho"
        )
        if len(edited) < len(df_atual):
            st.session_state.perguntas_df = st.session_state.perguntas_df.iloc[:len(edited)].copy()
            for col in cols_visiveis:
                if col in edited.columns:
                    st.session_state.perguntas_df[col] = edited[col].values
            st.rerun()
        else:
            for col in cols_visiveis:
                if col in edited.columns:
                    st.session_state.perguntas_df[col] = edited[col].values
    else:
        for idx, row in df_atual.iterrows():
            tipo  = row.get("Tipo", "Desconhecido")
            icon  = TIPO_ICONS.get(tipo, "❓")
            id_q  = row.get("ID", f"Q{idx+1:02d}")
            nivel = row.get("Nível", "—")
            seccao = row.get("Secção", "—")

            with st.container(border=True):
                c_num, c_body, c_acts = st.columns([0.5, 4.5, 1])
                c_num.markdown(f"### {icon}")
                with c_body:
                    st.markdown(f"**{id_q}** &nbsp;|&nbsp; {tipo} &nbsp;|&nbsp; Nível {nivel}")
                    st.caption(f"Secção: {seccao}")
                    enunciado_preview = row.get("Enunciado", "")
                    st.markdown(f"_{enunciado_preview[:120]}{'…' if len(enunciado_preview) > 120 else ''}_")
                with c_acts:
                    if st.button("✏️ Editar", key=f"card_edit_{id_q}", use_container_width=True):
                        st.session_state.view_details_id = id_q
                        st.session_state.imp_view = "Detalhes"
                        st.rerun()
                    if st.button("🗑️ Apagar", key=f"card_del_{id_q}", use_container_width=True):
                        st.session_state.perguntas_df = df_atual.drop(index=idx).reset_index(drop=True)
                        st.rerun()

    # Substitui toda a zona inferior da função por isto:
    st.divider()
    c_edit, c_exp = st.columns([1.5, 1])

    with c_edit:
        st.markdown("#### ✏️ Edição Profunda")
        df_ref = st.session_state.perguntas_df
        if not df_ref.empty:
            opcoes = (df_ref["ID"] + " — " + df_ref["Enunciado"].str[:55] + "…").tolist()
            sel = st.selectbox("Pergunta:", opcoes, label_visibility="collapsed")
            if st.button("✏️ Editar Selecionada", type="primary", use_container_width=True):
                st.session_state.view_details_id = sel.split(" — ")[0]
                st.session_state.imp_view = "Detalhes"
                st.rerun()

    with c_exp:
        st.markdown("#### 💾 Exportação de Segurança")
        if not st.session_state.perguntas_df.empty:
            csv = st.session_state.perguntas_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Descarregar Tabela (CSV)",
                data=csv,
                file_name="triagem_babelium.csv",
                mime="text/csv",
                use_container_width=True,
                type="primary"
            )
        else:
            st.warning("O carrinho está vazio.")

def render_imp_importar_word():
    st.title("📥 Importar Documento Word")
    st.caption("Extraia perguntas de um ficheiro .docx e adicione-as ao Painel de Triagem.")

    if docx is None:
        st.error("A biblioteca `python-docx` não está instalada. Execute `pip install python-docx`.")
        return

    with st.container(border=True):
        st.markdown("#### Configuração de Importação")
        c_nivel, c_curso = st.columns(2)
        st.session_state.nivel_global = c_nivel.selectbox(
            "Nível QECR das perguntas", NIVEL_OPTIONS,
            index=NIVEL_OPTIONS.index(st.session_state.nivel_global)
        )
        st.session_state.curso_global = c_curso.text_input(
            "Curso", value=st.session_state.curso_global, placeholder="Ex: PLE, Inglês B1"
        )

    with st.container(border=True):
        st.markdown("#### Ficheiro Word")
        uploaded_file = st.file_uploader(
            "Arraste ou selecione um ficheiro .docx",
            type=["docx"],
            label_visibility="collapsed"
        )
        if uploaded_file is not None:
            st.success(f"📄 **{uploaded_file.name}** carregado.")
            if st.button("🔄 Processar e Adicionar ao Carrinho", type="primary", use_container_width=True):
                with st.spinner("A extrair perguntas do documento..."):
                    prox_id = obter_proximo_id()
                    novo_df = extrair_perguntas_docx(uploaded_file, start_id=prox_id)

                if not novo_df.empty:
                    st.session_state.perguntas_df = pd.concat(
                        [st.session_state.perguntas_df, novo_df],
                        ignore_index=True
                    )
                    st.success(f"✅ **{len(novo_df)} perguntas** adicionadas ao carrinho!")
                    time.sleep(1)
                    st.session_state.imp_view = "Dashboard"
                    st.rerun()
                else:
                    st.error(
                        "Nenhuma pergunta encontrada neste ficheiro.\n\n"
                        "Verifique se os enunciados estão em **negrito** e se o documento "
                        "segue a estrutura BabeliUM padrão."
                    )


def render_imp_gerador_ia():
    st.title("✨ Assistente IA")
    st.caption(
        "Modo de pré-visualização — simula a futura integração com a API OpenAI. "
        "As perguntas geradas usam exemplos pedagógicos reais alinhados com os descritores QECR."
    )

    with st.container(border=True):
        st.markdown("#### Parâmetros de Geração")
        c1, c2 = st.columns(2)
        lingua = c1.selectbox("Língua", LINGUAS)
        nivel  = c2.selectbox("Nível QECR", NIVEL_OPTIONS, index=1)

        c3, c4 = st.columns(2)
        topico    = c3.text_input("Tópico / Tema", placeholder="Ex: Rotinas diárias, Viagens…")
        tipo_perg = c4.selectbox("Tipo de Pergunta", [t for t in TIPO_ICONS.keys() if t != "Desconhecido"])

        st.markdown("---")

        with st.expander("ℹ️ Como funcionará a integração OpenAI?", expanded=False):
            st.markdown("""
            **Fase 2 — Motor RAG + OpenAI**

            O sistema enviará à API um *system prompt* contendo:
            - Os **descritores QECR** do nível selecionado
            - O **contexto do BabeliUM** (língua, tópico, tipo de exercício)
            - A **estrutura JSON esperada** para integração direta no DataFrame

            *Por agora, o botão abaixo injeta um exemplo pedagógico real.*
            """)

        if st.button("🪄 Gerar e Adicionar ao Carrinho", type="primary", use_container_width=True):
            with st.spinner("A simular geração IA…"):
                time.sleep(1.2)

            exemplo  = obter_exemplo_ia(lingua, nivel, topico, tipo_perg)
            prox_id  = obter_proximo_id()
            nova_linha = {
                "Exportar":  True,
                "ID":        f"Q{prox_id:02d}",
                "Secção":    f"IA — {lingua}",
                "Enunciado": exemplo["enunciado"],
                "Pares":     exemplo.get("pares", []),
                "Tipo":      exemplo["tipo"],
                "Nível":     nivel,
                "Tópico":    topico if topico else "Tema 1",
            }
            st.session_state.perguntas_df = pd.concat(
                [st.session_state.perguntas_df, pd.DataFrame([nova_linha])],
                ignore_index=True
            )
            st.success(f"✅ Pergunta de **{lingua} — {nivel}** adicionada ao carrinho!")
            time.sleep(1)
            st.session_state.imp_view = "Dashboard"
            st.rerun()


def render_imp_detalhes():
    id_atual = st.session_state.view_details_id
    df       = st.session_state.perguntas_df
    idx_list = df.index[df["ID"] == id_atual].tolist()

    if not idx_list:
        st.session_state.imp_view = "Dashboard"
        st.rerun()
        return

    idx   = idx_list[0]
    linha = df.iloc[idx].to_dict()

    c_back, c_tit = st.columns([1, 6])
    if c_back.button("🔙 Voltar"):
        st.session_state.imp_view = "Dashboard"
        st.rerun()
    c_tit.subheader(f"✏️ Edição Profunda — {id_atual}")
    st.divider()

    with st.container(border=True):
        st.markdown("**Metadados**")
        cm1, cm2, cm3 = st.columns(3)
        tipos_list   = [t for t in TIPO_ICONS.keys()]
        tipo_atual   = linha.get("Tipo", "Desconhecido")
        novo_tipo    = cm1.selectbox("Tipo", tipos_list, index=tipos_list.index(tipo_atual) if tipo_atual in tipos_list else 0)
        nivel_atual  = linha.get("Nível", "A2")
        novo_nivel   = cm2.selectbox("Nível QECR", NIVEL_OPTIONS, index=NIVEL_OPTIONS.index(nivel_atual) if nivel_atual in NIVEL_OPTIONS else 1)
        topico_atual = linha.get("Tópico", "Tema 1")
        novo_topico  = cm3.selectbox("Tópico", TOPICO_OPTIONS, index=TOPICO_OPTIONS.index(topico_atual) if topico_atual in TOPICO_OPTIONS else 0)

    with st.container(border=True):
        st.markdown("**Enunciado**")
        st.caption("Corrija ou reformule o texto. As alterações guardam-se ao clicar em 'Guardar'.")
        novo_enunciado = st.text_area(
            "Enunciado", value=linha.get("Enunciado", ""),
            height=220, label_visibility="collapsed"
        )

    pares_atuais = linha.get("Pares", [])
    novos_pares  = []

    if novo_tipo == "Associação":
        with st.container(border=True):
            st.markdown("**Pares de Associação**")
            st.caption("Coluna A (pergunta) ↔ Coluna B (resposta). Use 🗑️ para eliminar ou ➕ para adicionar.")
            key_pares = f"pares_edit_{id_atual}"
            if key_pares not in st.session_state:
                init = pares_atuais if isinstance(pares_atuais, list) and pares_atuais else [{"esquerda": "", "direita": ""}]
                st.session_state[key_pares] = copy.deepcopy(init)

            temp = []
            for i, par in enumerate(st.session_state[key_pares]):
                c_del, c_esq, c_dir = st.columns([0.3, 2, 2])
                if c_del.button("🗑️", key=f"del_par_{id_atual}_{i}"):
                    st.session_state[key_pares].pop(i); st.rerun()
                esq  = c_esq.text_input(f"A #{i+1}", par.get("esquerda", ""), label_visibility="collapsed", placeholder="Coluna A", key=f"esq_{id_atual}_{i}")
                dir_ = c_dir.text_input(f"B #{i+1}", par.get("direita",   ""), label_visibility="collapsed", placeholder="Coluna B", key=f"dir_{id_atual}_{i}")
                temp.append({"esquerda": esq, "direita": dir_})
            novos_pares = temp

            if st.button("➕ Adicionar par"):
                st.session_state[key_pares].append({"esquerda": "", "direita": ""}); st.rerun()

    with st.expander("👁️ Pré-visualização (clique para expandir)", expanded=False):
        icon = TIPO_ICONS.get(novo_tipo, "❓")
        st.markdown(f"**{icon} {novo_tipo}** — Nível {novo_nivel}")
        st.markdown("---")
        st.markdown(novo_enunciado)
        if novo_tipo == "Associação" and novos_pares:
            st.markdown("---")
            ca, cb = st.columns(2)
            ca.markdown("**Coluna A**")
            cb.markdown("**Coluna B**")
            for p in novos_pares:
                if p.get("esquerda"): ca.markdown(f"- {p['esquerda']}")
                if p.get("direita"):  cb.markdown(f"🔹 {p['direita']}")

    st.write("")
    cs, cn, cc = st.columns(3)

    def _guardar():
        st.session_state.perguntas_df.at[idx, "Enunciado"] = novo_enunciado
        st.session_state.perguntas_df.at[idx, "Tipo"]      = novo_tipo
        st.session_state.perguntas_df.at[idx, "Nível"]     = novo_nivel
        st.session_state.perguntas_df.at[idx, "Tópico"]    = novo_topico
        if novo_tipo == "Associação":
            st.session_state.perguntas_df.at[idx, "Pares"] = novos_pares

    if cs.button("💾 Guardar e Voltar", type="primary", use_container_width=True):
        _guardar()
        st.success("✅ Alterações guardadas.")
        time.sleep(0.7)
        st.session_state.imp_view = "Dashboard"
        st.rerun()

    if cn.button("⏭️ Guardar e Seguir", use_container_width=True):
        _guardar()
        df_nav = st.session_state.perguntas_df
        pos    = df_nav.index.get_loc(idx)
        if pos + 1 < len(df_nav):
            st.session_state.view_details_id = df_nav.iloc[pos + 1]["ID"]
            st.rerun()
        else:
            st.info("Chegou à última pergunta.")
            st.session_state.imp_view = "Dashboard"
            st.rerun()

    if cc.button("❌ Cancelar", use_container_width=True):
        st.session_state.imp_view = "Dashboard"
        st.rerun()

# ==============================================================================
# VIEW: CONVERSOR PEAKIT -> MOODLE
# ==============================================================================

def render_conversor_CSV():
    st.title("🔄 Conversor PEAKIT -> Moodle")
    
    # --- MANUAL DE UTILIZADOR INTEGRADO ---
    with st.expander("📖 GUIA DE UTILIZAÇÃO: Como converter e importar"):
        st.markdown("""
        ### 1. No Portal PEAKIT
        * Exporte a listagem de inscritos em formato **CSV**.
        * Garanta que o ficheiro inclui as colunas: *ID*, *Nome*, *Email* e *Serviço Educativo*.

        ### 2. No BabeliUM Editor (Aqui)
        * Carregue o ficheiro no campo abaixo.
        * A aplicação separa automaticamente nomes próprio/apelido e identifica o curso.
        * **Verifique a tabela:** Se o nome do curso na coluna `course1` não coincidir com o *Shortname* do Moodle, corrija-o diretamente na célula da tabela.
        * Clique em **Descarregar CSV para Moodle**.

        ### 3. No MOODLE
        Vá a *Administração do Site > Utilizadores > Contas > Carregar utilizadores*:
        1. **Ficheiro:** Submeta o CSV que descarregou desta app.
        2. **Delimitador:** Selecione obrigatoriamente **Vírgula (,)**.
        3. **Posteriormente à importação - Senha de novo utilizador:** Selecione `Criar senha, se necessário, e enviar por e-mail`.
        4. **Forçar mudança de senha:** Selecione `Todos`.
        """)

    st.divider()

    uploaded_file = st.file_uploader("Selecione o ficheiro CSV do Peakit", type=["csv"])
    
    if uploaded_file:
        try:
            raw_data = uploaded_file.read()
            try:
                content = raw_data.decode('utf-8')
            except UnicodeDecodeError:
                content = raw_data.decode('latin-1')
            
            sep = ';' if ';' in content.split('\n')[0] else ','
            df = pd.read_csv(io.StringIO(content), sep=sep, header=None)
            df = df.dropna(how='all')

            if not str(df.iloc[0, 0]).isdigit():
                df = df.iloc[1:].reset_index(drop=True)
            
            def process_names(full_name):
                name = str(full_name).strip().replace('?', ' ') 
                parts = name.split(' ', 1)
                first = parts[0]
                last = parts[1] if len(parts) > 1 else ' '
                return first, last

            new_data = []
            for _, row in df.iterrows():
                if pd.isna(row[0]) or str(row[0]).strip() == "":
                    continue
                
                num_cols = len(row)
                idx_email = 2 
                if num_cols > 3 and "@" in str(row[3]):
                    idx_email = 3
                
                username = str(row[0]).strip().lower()
                
                if idx_email == 3:
                    fname = str(row[1]).strip()
                    lname = str(row[2]).strip()
                    email = str(row[3]).strip()
                    course = str(row[4]).strip() if num_cols > 4 else ""
                else:
                    fname, lname = process_names(row[1])
                    email = str(row[2]).strip()
                    course = str(row[3]).strip() if num_cols > 3 else ""

                new_data.append({
                    "username": username,
                    "firstname": fname,
                    "lastname": lname,
                    "email": email,
                    "course1": course,
                    "role1": "student"
                })
            
            df_moodle = pd.DataFrame(new_data).dropna(subset=['username', 'firstname'])
            
            st.success("✅ Ficheiro processado com sucesso!")
            
            # Interface de edição
            df_final = st.data_editor(df_moodle, use_container_width=True, hide_index=True)
            
            csv_ready = df_final.to_csv(
                index=False, sep=',', encoding='utf-8', quoting=1, lineterminator='\n'
            )
            
            st.download_button(
                label="📥 Descarregar CSV para Moodle",
                data=csv_ready,
                file_name=f"moodle_inscricao_{time.strftime('%Y%m%d')}.csv",
                mime="text/csv",
                type="primary",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"Erro ao processar o ficheiro: {e}")

# ══════════════════════════════════════════════════════════════════
# ROUTER PRINCIPAL
# ══════════════════════════════════════════════════════════════════
if st.session_state.modulo_ativo == "editor":
    view = st.session_state.active_view
    if   view == "Dashboard": render_editor_dashboard()
    elif view == "Editor":    render_editor_questao()
    elif view == "Exportar":  render_editor_exportar()

else:  # importador
    view = st.session_state.imp_view
    if   view == "Dashboard":    render_imp_dashboard()
    elif view == "Importar_Word": render_imp_importar_word()
    elif view == "Gerador_IA":   render_imp_gerador_ia()
    elif view == "Detalhes":     render_imp_detalhes()
    elif view == "Conversor_CSV": render_conversor_CSV()
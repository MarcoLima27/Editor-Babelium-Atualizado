# app_v1.py
# BabeliUM — Importador & Painel de Triagem (V1.1)

import streamlit as st
import pandas as pd
import copy
import time
import re
import docx
import os

# ══════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ══════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="BabeliUM — Importador",
    page_icon="📥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ══════════════════════════════════════════════════════════════════
# ESTILOS CSS — espelho do app.py original (Light / Dark Mode)
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
/* Containers secundários — igual ao app.py */
div[data-testid="stExpander"],
div[data-testid="stMetric"],
div[data-testid="stForm"] {
    background-color: var(--secondary-background-color);
    border-radius: 8px;
    border: 1px solid rgba(128, 128, 128, 0.2);
    padding: 5px;
}

/* Botões */
.stButton>button {
    border-radius: 6px;
    font-weight: 500;
}

/* Logos sem fundo nem arredondamento */
img {
    border-radius: 0px !important;
    background-color: transparent !important;
}

/* Tabela com cantos arredondados */
div[data-testid="stDataFrameResizable"] {
    border-radius: 8px;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# CONSTANTES
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

# Línguas suportadas — apenas PT e EN (sem alemão / francês por agora)
LINGUAS = ["Português (PLE)", "Inglês"]

# ══════════════════════════════════════════════════════════════════
# GESTÃO DE ESTADO
# ══════════════════════════════════════════════════════════════════
defaults = {
    "active_view":    "Dashboard",  # Dashboard | Importar_Word | Gerador_IA | Detalhes
    "view_details_id": None,
    "perguntas_df":   pd.DataFrame(),
    "nivel_global":   "A2",
    "curso_global":   "PLE",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def obter_proximo_id() -> int:
    """Calcula o próximo número de ID sequencial a partir do DataFrame atual."""
    df = st.session_state.perguntas_df
    if df.empty or "ID" not in df.columns:
        return 1
    try:
        nums = df["ID"].str.replace("Q", "", regex=False).astype(int)
        return int(nums.max()) + 1
    except (ValueError, AttributeError):
        return len(df) + 1


# ══════════════════════════════════════════════════════════════════
# PARSER ESTRUTURAL WORD (mantido e estável)
# ══════════════════════════════════════════════════════════════════
def _bold_text(paragraph) -> str:
    return " ".join(r.text.strip() for r in paragraph.runs if r.bold and r.text.strip())

def _iter_block_items(doc):
    """Itera parágrafos e tabelas na ordem real do documento."""
    import docx.text.paragraph as _para
    import docx.table as _tbl
    for child in doc.element.body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            yield ('paragraph', _para.Paragraph(child, doc))
        elif tag == 'tbl':
            yield ('table', _tbl.Table(child, doc))

def _extrair_pares_tabela(table) -> list:
    """
    Extrai pares (esquerda, direita) de tabelas de associação.
    Suporta células simples e células multi-linha (Word mescla colunas em bloco).
    """
    pares = []
    for i, row in enumerate(table.rows):
        cells = [c.text.strip() for c in row.cells]
        if not cells or not cells[0]:
            continue
        if i == 0 and any(h in cells[0].lower() for h in ["coluna", "column"]):
            continue  # saltar cabeçalho
        if len(cells) >= 2:
            esq_raw = cells[0]
            dir_raw = cells[1] if len(cells) > 1 else ""
            esq_linhas = [l.strip() for l in esq_raw.split("\n") if l.strip()]
            dir_linhas = [l.strip() for l in dir_raw.split("\n") if l.strip()]
            # Célula multi-linha → processar como bloco
            if len(esq_linhas) > 1 or len(dir_linhas) > 1:
                esq_limpa = [re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', l).strip() for l in esq_linhas]
                dir_limpa = [re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', l).strip() for l in dir_linhas]
                for j, esq in enumerate(esq_limpa):
                    if esq:
                        pares.append({"esquerda": esq, "direita": dir_limpa[j] if j < len(dir_limpa) else ""})
                return pares
            # Linha simples
            esq  = re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', esq_raw).strip()
            dir_ = re.sub(r'^[A-Za-z0-9]+[.)]\s*', '', dir_raw).strip()
            if esq:
                pares.append({"esquerda": esq, "direita": dir_})
    return pares

def _normalizar_marcador_correto(paragraph) -> str:
    """
    Reconstrói o texto do parágrafo substituindo qualquer marcador de
    'resposta correta' pelo asterisco (*), que é o standard do BabeliUM Editor.

    Marcadores normalizados:
      ✓  (U+2713) — visto simples, o mais comum no template
      ✔  (U+2714) — visto pesado
      ☑  (U+2611) — caixa com visto
    """
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

# Marcadores que identificam tabelas de instrução/template — nunca conteúdo
_MARCADORES_INSTRUCAO = [
    "instrução:", "apague este bloco", "[[", "⚙", "[exemplo", "[template",
    "guia rápido", "🚀", "🔵", "🟡", "🟢", "📝",
]

def _tabela_e_instrucao(table) -> bool:
    """Devolve True se a tabela for um bloco de instrução/exemplo do template."""
    texto = " ".join(
        c.text.strip().lower()
        for row in table.rows
        for c in row.cells
    )
    return any(m in texto for m in _MARCADORES_INSTRUCAO)

def _e_titulo_seccao(bold: str, txt: str) -> bool:
    """
    Devolve True se o parágrafo for um título de secção (ex: 'Gramática').
    Critérios: totalmente em bold, curto, sem dígitos, não é um enunciado numerado.
    """
    if not bold or txt != bold:
        return False
    if len(txt) > 50:
        return False
    if any(c.isdigit() for c in txt):
        return False
    # Enunciados reais começam com algarismo romano (I., II., III.) — excluir
    if re.match(r'^[IVX]+\.', txt.strip()):
        return False
    return True

def extrair_perguntas_docx(uploaded_file, start_id: int = 1) -> pd.DataFrame:
    """
    Extrai perguntas de um .docx na ordem real do documento.

    Correções aplicadas:
    - FIX 1: Mudança de secção força guardar() imediatamente, evitando que
             parágrafos de texto livre (ex: texto de apoio) sejam absorvidos
             pelo buffer da pergunta anterior.
    - FIX 2: Tabelas de instrução/template são ignoradas (não contaminam buffers).
    - FIX 3: Parágrafos que repetem o enunciado já em buffer são ignorados
             (evita perguntas duplicadas/vazias quando o professor cola o título duas vezes).
    - FIX 4: Parágrafos sem bold fora de um estado ativo são ignorados
             (texto solto como 'Leia o texto seguinte:' não entra em nenhum buffer).
    """
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
            # Usar o reconstrutor normalizado — substitui ✓ / ✔ / ☑ por *
            txt = _normalizar_marcador_correto(obj)
            if not txt:
                continue

            # Parar na Chave de Respostas
            if "chave de respostas" in txt.lower():
                guardar()
                break

            bold = _bold_text(obj)

            # FIX 1 — Título de secção: forçar guardar() antes de mudar secção
            if _e_titulo_seccao(bold, txt):
                guardar()          # fecha a pergunta anterior imediatamente
                estado = None      # reset explícito — texto livre pós-secção não entra em nenhum buffer
                secao_atual = txt
                continue

            tipo_det = _detetar_tipo(bold, txt)
            if tipo_det != "Desconhecido":
                # FIX 3 — ignorar se este parágrafo é idêntico ao início do buffer atual
                # (professor duplicou o enunciado sem bold dentro do exercício)
                primeira_linha_buffer = buffer_enunciado.strip().split("\n")[0].strip().lower()
                if txt.strip().lower() == primeira_linha_buffer:
                    continue  # duplicado — ignorar silenciosamente

                guardar()
                estado = tipo_det
                buffer_enunciado = txt + "\n"
            elif estado:
                # FIX 4 — só acumular no buffer se já temos um estado activo
                buffer_enunciado += txt + "\n"
            # Se estado is None e não é título nem enunciado → ignorar (texto solto)

        elif kind == 'table':
            # FIX 2 — ignorar tabelas de instrução/template
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
# EXEMPLOS ILUSTRATIVOS PARA O GERADOR IA
# ══════════════════════════════════════════════════════════════════
EXEMPLOS_IA = {
    "Português (PLE)": {
        "A1": [
            {
                "tipo": "Escolha Múltipla",
                "enunciado": (
                    "Qual é a forma correta do verbo 'ser' na 3.ª pessoa do singular?\n"
                    "a) são\nb) somos\nc) é\nd) és"
                ),
            },
            {
                "tipo": "Lacunas (Escrita)",
                "enunciado": (
                    "Complete as frases com o artigo definido correto (o, a, os, as):\n"
                    "1. _____ livro está na mesa.\n"
                    "2. _____ professora chama-se Ana.\n"
                    "3. _____ alunos são simpáticos."
                ),
            },
        ],
        "A2": [
            {
                "tipo": "Lacunas (Cloze)",
                "enunciado": (
                    "Escolha a opção correta para completar cada frase:\n"
                    "1. Ontem eu _______ ao cinema. (fui / vou / irei)\n"
                    "2. Nós _______ juntos todos os dias. (almoçamos / almoçam / almoçar)\n"
                    "3. Eles já _______ em Lisboa há cinco anos. (vivem / vivo / vivemos)"
                ),
            },
            {
                "tipo": "Associação",
                "enunciado": "Associe cada profissão com a sua descrição.",
                "pares": [
                    {"esquerda": "Médico",      "direita": "Trata de doentes"},
                    {"esquerda": "Professor",   "direita": "Ensina alunos"},
                    {"esquerda": "Engenheiro",  "direita": "Constrói edifícios"},
                ],
            },
        ],
        "B1": [
            {
                "tipo": "Ensaio (Texto Livre)",
                "enunciado": (
                    "Descreva um dia típico da sua semana. "
                    "Inclua as suas rotinas matinais, o seu trabalho ou estudos, "
                    "e as atividades que faz ao final do dia. (80–100 palavras)"
                ),
            },
        ],
    },
    "Inglês": {
        "A1": [
            {
                "tipo": "Escolha Múltipla",
                "enunciado": (
                    "Choose the correct form of the verb 'to be':\n"
                    "1. She _______ a student. (is / are / am)\n"
                    "2. They _______ from Brazil. (is / are / am)\n"
                    "3. I _______ happy today. (is / are / am)"
                ),
            },
        ],
        "A2": [
            {
                "tipo": "Lacunas (Cloze)",
                "enunciado": (
                    "Choose the best option to complete each sentence:\n"
                    "1. Yesterday, I _______ to the supermarket. (went / go / goes)\n"
                    "2. She _______ English very well. (speaks / speak / speaking)\n"
                    "3. We _______ dinner at 7 pm every night. (have / has / having)"
                ),
            },
            {
                "tipo": "Associação",
                "enunciado": "Match each word with its correct definition.",
                "pares": [
                    {"esquerda": "Library",    "direita": "A place to borrow books"},
                    {"esquerda": "Pharmacy",   "direita": "A place to buy medicine"},
                    {"esquerda": "Bakery",     "direita": "A place to buy bread"},
                ],
            },
        ],
        "B1": [
            {
                "tipo": "Ensaio (Texto Livre)",
                "enunciado": (
                    "Write a short paragraph about your last holiday. "
                    "Where did you go? Who did you go with? "
                    "What did you do? (80–100 words)"
                ),
            },
        ],
    },
}

def obter_exemplo_ia(lingua: str, nivel: str, topico: str, tipo_pergunta: str) -> dict:
    """
    Devolve um exemplo ilustrativo realista com base nas configurações escolhidas.
    Usa os exemplos pré-definidos quando disponíveis, ou gera um genérico adaptado.
    """
    exemplos_lingua = EXEMPLOS_IA.get(lingua, {})
    exemplos_nivel  = exemplos_lingua.get(nivel, [])

    # Filtrar por tipo se existir correspondência
    candidatos = [e for e in exemplos_nivel if e.get("tipo") == tipo_pergunta]
    if not candidatos:
        candidatos = exemplos_nivel

    if candidatos:
        exemplo = copy.deepcopy(candidatos[0])
        # Injetar tópico se fornecido e o enunciado não o mencionar
        if topico and topico.lower() not in exemplo["enunciado"].lower():
            exemplo["enunciado"] = f"[Tópico: {topico}]\n\n" + exemplo["enunciado"]
        return exemplo

    # Fallback genérico (quando não há exemplo pré-definido para aquele nível)
    lang_label = "em Português" if "Português" in lingua else "in English"
    return {
        "tipo": tipo_pergunta,
        "enunciado": (
            f"Exemplo de pergunta {lang_label} — Nível {nivel}."
            + (f"\nTópico: {topico}." if topico else "")
            + f"\n\n[Este é um exemplo ilustrativo. A integração com a API OpenAI "
              f"gerará conteúdo real baseado nos descritores QECR {nivel}.]"
        ),
        "pares": [],
    }


# ══════════════════════════════════════════════════════════════════
# SIDEBAR — grafismo do app.py original
# ══════════════════════════════════════════════════════════════════
with st.sidebar:
    pasta_atual = os.path.dirname(os.path.abspath(__file__))
    cam_uminho  = os.path.join(pasta_atual, "img", "Logo_uminho.png")
    cam_babel   = os.path.join(pasta_atual, "img", "Logo_babeliUM.jpg")

    # Logo UMinho
    if os.path.exists(cam_uminho):
        st.image(cam_uminho, width=180)
    else:
        st.header("UMinho")

    st.write("")

    # Título + Logo BabeliUM (colunas [2,1] igual ao app.py)
    c_txt, c_img = st.columns([2, 1])
    with c_txt:
        st.markdown(
            "<h3 style='margin:0; padding-top:15px; text-align:right; font-size:1.2rem;'>"
            "BabeliUM Editor</h3>",
            unsafe_allow_html=True
        )
    with c_img:
        if os.path.exists(cam_babel):
            st.image(cam_babel, width=70)
        else:
            st.write("📥")

    st.markdown("---")

    # ── Navegação principal ──
    def _nav_btn(label: str, view: str):
        tipo = "primary" if st.session_state.active_view == view else "secondary"
        if st.button(label, type=tipo, use_container_width=True):
            st.session_state.active_view = view
            st.rerun()

    _nav_btn("🛒 Painel de Triagem",       "Dashboard")

    st.markdown("---")
    st.markdown("**➕ Adicionar ao Carrinho**")

    _nav_btn("📥 Importar Word",           "Importar_Word")
    _nav_btn("✨ Assistente IA (Mockup)",  "Gerador_IA")

    st.markdown("---")

    # ── Estatísticas (métricas estilo app.py) ──
    df_side = st.session_state.perguntas_df
    c1, c2 = st.columns(2)
    c1.metric("No Carrinho", len(df_side))
    if not df_side.empty and "Exportar" in df_side.columns:
        c2.metric("✅ A Exportar", int(df_side["Exportar"].sum()))

        # Distribuição por tipo
        st.markdown("**Distribuição**")
        for tipo, n in df_side["Tipo"].value_counts().items():
            icon = TIPO_ICONS.get(tipo, "❓")
            st.markdown(f"{icon} **{tipo}** — {n}")

        st.markdown("---")
        if st.button("🗑️ Esvaziar Carrinho", use_container_width=True):
            st.session_state.perguntas_df = pd.DataFrame()
            st.session_state.active_view = "Dashboard"
            st.rerun()
    else:
        c2.metric("✅ A Exportar", 0)
        st.info("Use o menu acima para adicionar perguntas.")


# ══════════════════════════════════════════════════════════════════
# VIEW: IMPORTAR WORD
# ══════════════════════════════════════════════════════════════════
def render_importar_word():
    st.title("📥 Importar Documento Word")
    st.caption("Extraia perguntas de um ficheiro .docx e adicione-as ao Painel de Triagem.")

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
                    st.session_state.active_view = "Dashboard"
                    st.rerun()
                else:
                    st.error(
                        "Nenhuma pergunta encontrada neste ficheiro.\n\n"
                        "Verifique se os enunciados estão em **negrito** e se o documento "
                        "segue a estrutura BabeliUM padrão."
                    )


# ══════════════════════════════════════════════════════════════════
# VIEW: GERADOR IA (MOCKUP)
# ══════════════════════════════════════════════════════════════════
def render_gerador_ia():
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
        topico = c3.text_input("Tópico / Tema", placeholder="Ex: Rotinas diárias, Viagens…")
        tipo_perg = c4.selectbox(
            "Tipo de Pergunta",
            [t for t in TIPO_ICONS.keys() if t != "Desconhecido"]
        )

        st.markdown("---")

        # Aviso de mockup com contexto técnico
        with st.expander("ℹ️ Como funcionará a integração OpenAI?", expanded=False):
            st.markdown("""
            **Fase 2 — Motor RAG + OpenAI**

            O sistema enviará à API um *system prompt* contendo:
            - Os **descritores QECR** do nível selecionado (ex: A2 → vocabulário quotidiano, frases simples)
            - O **contexto do BabeliUM** (língua, tópico, tipo de exercício)
            - A **estrutura JSON esperada** para integração direta no DataFrame

            O modelo devolverá um objeto estruturado que é automaticamente convertido
            numa linha do Painel de Triagem — pronta para revisão e exportação para o Moodle.

            *Por agora, o botão abaixo injeta um exemplo pedagógico real correspondente
            às configurações escolhidas.*
            """)

        if st.button("🪄 Gerar e Adicionar ao Carrinho", type="primary", use_container_width=True):
            with st.spinner("A simular geração OpenAI…"):
                time.sleep(1.2)  # simular latência da API

            exemplo = obter_exemplo_ia(lingua, nivel, topico, tipo_perg)
            prox_id = obter_proximo_id()

            nova_linha = {
                "Exportar": True,
                "ID":       f"Q{prox_id:02d}",
                "Secção":   f"IA — {lingua}",
                "Enunciado": exemplo["enunciado"],
                "Pares":    exemplo.get("pares", []),
                "Tipo":     exemplo["tipo"],
                "Nível":    nivel,
                "Tópico":   topico if topico else "Tema 1",
            }

            st.session_state.perguntas_df = pd.concat(
                [st.session_state.perguntas_df, pd.DataFrame([nova_linha])],
                ignore_index=True
            )
            st.success(f"✅ Pergunta de **{lingua} — {nivel}** adicionada ao carrinho!")
            time.sleep(1)
            st.session_state.active_view = "Dashboard"
            st.rerun()


# ══════════════════════════════════════════════════════════════════
# VIEW: DASHBOARD — PAINEL DE TRIAGEM
# ══════════════════════════════════════════════════════════════════
def render_dashboard():
    st.title("🛒 Painel de Triagem")

    df_atual = st.session_state.perguntas_df

    # ── Carrinho vazio ──
    if df_atual.empty:
        with st.container(border=True):
            st.write("")
            col_center = st.columns([1, 2, 1])[1]
            with col_center:
                st.markdown(
                    "<div style='text-align:center; color:#888;'>"
                    "<h3>O carrinho está vazio</h3>"
                    "<p>Use o menu à esquerda para <b>Importar Word</b> "
                    "ou usar o <b>Assistente IA</b>.</p>"
                    "</div>",
                    unsafe_allow_html=True
                )
            st.write("")
        return

    # ── Cabeçalho com contagem ──
    c_tot, c_view = st.columns([4, 1.5])
    c_tot.subheader(f"Questões no carrinho ({len(df_atual)})")
    view_mode = c_view.radio(
        "Modo:", ["Tabela", "Cartões"],
        horizontal=True,
        label_visibility="collapsed"
    )

    st.markdown(
        "Edite **Tipo**, **Nível** e **Tópico** diretamente na tabela. "
        "Selecione linhas e prima `Delete` para eliminar. "
        "Use **Edição Profunda** para editar o enunciado completo."
    )
    st.divider()

    # ── Modo Tabela ──
    if view_mode == "Tabela":
        cols_visiveis = [c for c in ["Exportar", "ID", "Secção", "Tipo", "Nível", "Tópico", "Enunciado"]
                         if c in df_atual.columns]

        edited = st.data_editor(
            df_atual[cols_visiveis],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Exportar": st.column_config.CheckboxColumn("Exportar?", default=True, width="small"),
                "ID":       st.column_config.TextColumn("ID", disabled=True, width="small"),
                "Secção":   st.column_config.TextColumn("Secção", disabled=True, width="medium"),
                "Tipo":     st.column_config.SelectboxColumn("Tipo", options=list(TIPO_ICONS.keys()), width="medium"),
                "Nível":    st.column_config.SelectboxColumn("Nível", options=NIVEL_OPTIONS, width="small"),
                "Tópico":   st.column_config.SelectboxColumn("Tópico", options=TOPICO_OPTIONS, width="small"),
                "Enunciado": st.column_config.TextColumn("Enunciado", width="large"),
            },
            key="tabela_carrinho"
        )

        # Sincronização robusta (eliminar linhas e edições inline)
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

    # ── Modo Cartões (estilo app.py) ──
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
                        st.session_state.active_view = "Detalhes"
                        st.rerun()
                    if st.button("🗑️ Apagar", key=f"card_del_{id_q}", use_container_width=True):
                        st.session_state.perguntas_df = df_atual.drop(index=idx).reset_index(drop=True)
                        st.rerun()

    # ── Zona inferior: Edição Profunda + Exportação ──
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
                st.session_state.active_view = "Detalhes"
                st.rerun()

    with c_exp:
        st.markdown("#### 🚀 Exportação")
        df_ref = st.session_state.perguntas_df
        n_exp  = int(df_ref["Exportar"].sum()) if not df_ref.empty and "Exportar" in df_ref.columns else 0
        if n_exp > 0:
            st.success(f"**{n_exp}** pergunta(s) marcada(s) para exportar.")
        else:
            st.warning("Nenhuma pergunta selecionada.")
        st.button(
            "📦 Enviar para Moodle (Em breve)",
            type="primary",
            use_container_width=True,
            disabled=True,
            help="Integração Moodle REST API em desenvolvimento."
        )


# ══════════════════════════════════════════════════════════════════
# VIEW: EDIÇÃO PROFUNDA (DETALHES)
# ══════════════════════════════════════════════════════════════════
def render_detalhes():
    id_atual = st.session_state.view_details_id
    df       = st.session_state.perguntas_df
    idx_list = df.index[df["ID"] == id_atual].tolist()

    # Segurança: se o ID não existir, voltar ao Dashboard
    if not idx_list:
        st.session_state.active_view = "Dashboard"
        st.rerun()
        return

    idx   = idx_list[0]
    linha = df.iloc[idx].to_dict()

    # ── Cabeçalho com botão Voltar (estilo app.py) ──
    c_back, c_tit = st.columns([1, 6])
    if c_back.button("🔙 Voltar"):
        st.session_state.active_view = "Dashboard"
        st.rerun()
    c_tit.subheader(f"✏️ Edição Profunda — {id_atual}")

    st.divider()

    # ── Metadados ──
    with st.container(border=True):
        st.markdown("**Metadados**")
        cm1, cm2, cm3 = st.columns(3)

        tipos_list  = [t for t in TIPO_ICONS.keys()]
        tipo_atual  = linha.get("Tipo", "Desconhecido")
        novo_tipo   = cm1.selectbox(
            "Tipo", tipos_list,
            index=tipos_list.index(tipo_atual) if tipo_atual in tipos_list else 0
        )
        nivel_atual = linha.get("Nível", "A2")
        novo_nivel  = cm2.selectbox(
            "Nível QECR", NIVEL_OPTIONS,
            index=NIVEL_OPTIONS.index(nivel_atual) if nivel_atual in NIVEL_OPTIONS else 1
        )
        topico_atual = linha.get("Tópico", "Tema 1")
        novo_topico  = cm3.selectbox(
            "Tópico", TOPICO_OPTIONS,
            index=TOPICO_OPTIONS.index(topico_atual) if topico_atual in TOPICO_OPTIONS else 0
        )

    # ── Enunciado ──
    with st.container(border=True):
        st.markdown("**Enunciado**")
        st.caption("Corrija ou reformule o texto. As alterações guardam-se ao clicar em 'Guardar'.")
        novo_enunciado = st.text_area(
            "Enunciado",
            value=linha.get("Enunciado", ""),
            height=220,
            label_visibility="collapsed"
        )

    # ── Pares de Associação (só se o tipo for Associação) ──
    pares_atuais = linha.get("Pares", [])
    novos_pares  = []

    if novo_tipo == "Associação":
        with st.container(border=True):
            st.markdown("**Pares de Associação**")
            st.caption("Coluna A (pergunta) ↔ Coluna B (resposta). Use 🗑️ para eliminar ou ➕ para adicionar.")

            key_pares = f"pares_edit_{id_atual}"
            if key_pares not in st.session_state:
                init = pares_atuais if isinstance(pares_atuais, list) and pares_atuais \
                       else [{"esquerda": "", "direita": ""}]
                st.session_state[key_pares] = copy.deepcopy(init)

            temp = []
            for i, par in enumerate(st.session_state[key_pares]):
                c_del, c_esq, c_dir = st.columns([0.3, 2, 2])
                if c_del.button("🗑️", key=f"del_par_{id_atual}_{i}"):
                    st.session_state[key_pares].pop(i)
                    st.rerun()
                esq  = c_esq.text_input(f"A #{i+1}", par.get("esquerda", ""),
                                         label_visibility="collapsed", placeholder="Coluna A",
                                         key=f"esq_{id_atual}_{i}")
                dir_ = c_dir.text_input(f"B #{i+1}", par.get("direita", ""),
                                         label_visibility="collapsed", placeholder="Coluna B",
                                         key=f"dir_{id_atual}_{i}")
                temp.append({"esquerda": esq, "direita": dir_})
            novos_pares = temp

            if st.button("➕ Adicionar par"):
                st.session_state[key_pares].append({"esquerda": "", "direita": ""})
                st.rerun()

    # ── Pré-visualização ──
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
                if p.get("esquerda"):
                    ca.markdown(f"- {p['esquerda']}")
                if p.get("direita"):
                    cb.markdown(f"🔹 {p['direita']}")

    # ── Botões de Ação ──
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
        st.session_state.active_view = "Dashboard"
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
            st.session_state.active_view = "Dashboard"
            st.rerun()

    if cc.button("❌ Cancelar", use_container_width=True):
        st.session_state.active_view = "Dashboard"
        st.rerun()


# ══════════════════════════════════════════════════════════════════
# ROUTER
# ══════════════════════════════════════════════════════════════════
VIEW = st.session_state.active_view
if   VIEW == "Dashboard":    render_dashboard()
elif VIEW == "Importar_Word": render_importar_word()
elif VIEW == "Gerador_IA":   render_gerador_ia()
elif VIEW == "Detalhes":     render_detalhes()

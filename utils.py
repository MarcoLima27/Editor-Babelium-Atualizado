import uuid
import re
import xml.etree.ElementTree as ET
from xml.dom import minidom
import xml.etree.ElementTree as ET
from xml.dom import minidom
import docx  # Garante que tens 'python-docx' instalado

# Gera um ID único (ex: "q_a1b2c3d4")
def new_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:10]}"

# Conta quantos espaços [ ] existem no texto
def count_gaps(text: str) -> int:
    return text.count("[ ]")

# Prepara texto para XML (substitui & < > " ')
def escape_xml(s: str) -> str:
    if s is None:
        return ""
    return (str(s).replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;"))

def parse_ficha_word(uploaded_file):
    """
    Lê o ficheiro Word e extrai as perguntas de escolha múltipla 
    corrigindo a falsa identificação de Cloze.
    """
    doc = docx.Document(uploaded_file)
    texto_completo = "\n".join([p.text for p in doc.paragraphs])
    
    # Regex para capturar cada subpergunta numerada (ex: 1. Eu _______ em Braga...)
    # Captura o número, o enunciado com a lacuna e o bloco de opções seguinte
    padrao_bloco = re.compile(
        r'(\d+)\.\s*(.*?_______.*?)\s*((?:[a-d]\).*?\n?)+)(?=\n\d+\.|\Z)', 
        re.DOTALL | re.IGNORECASE
    )
    
    blocos_questoes = padrao_bloco.findall(texto_completo)
    perguntas_processadas = []
    
    for num, enunciado, bloco_opcoes in blocos_questoes:
        # Extrair alíneas individuais (a, b, c, d) dentro do bloco de opções
        opcoes = re.findall(r'([a-d])\)\s*(.*?)(?=\s*[a-d]\)|\Z)', bloco_opcoes, re.DOTALL)
        
        choices_dict = {}
        resposta_correta = "a" # Fallback por defeito
        
        for alinea, texto_opcao in opcoes:
            texto_limpo = texto_opcao.strip()
            
            # SE no teu Word a resposta correta vier marcada (ex: com um asterisco ou algo semelhante)
            # podes mapear aqui. Caso contrário, o professor edita na tabela do Streamlit.
            if "*" in texto_limpo:
                texto_limpo = texto_limpo.replace("*", "").strip()
                resposta_correta = alinea.lower().strip()
                
            choices_dict[alinea.lower().strip()] = texto_limpo
            
        # Garante o preenchimento de opções vazias se o regex falhar em alguma alínea
        for alinea in ['a', 'b', 'c', 'd']:
            if alinea not in choices_dict:
                choices_dict[alinea] = "Opção em falta"

        perguntas_processadas.append({
            "name": f"Questão {num}",
            "questiontext": f"{enunciado.strip()}",
            "choice_a": choices_dict.get("a"),
            "choice_b": choices_dict.get("b"),
            "choice_c": choices_dict.get("c"),
            "choice_d": choices_dict.get("d"),
            "answer": resposta_correta
        })
        
    return perguntas_processadas

def gerar_moodle_xml(lista_perguntas):
    """
    Gera a estrutura XML válida do Moodle com base nas perguntas da tabela.
    """
    quiz = ET.Element("quiz")
    
    for q in lista_perguntas:
        question = ET.SubElement(quiz, "question", type="multichoice")
        
        # Nome
        name = ET.SubElement(question, "name")
        ET.SubElement(name, "text").text = str(q.get("name", "Pergunta"))
        
        # Enunciado em HTML
        qtext = ET.SubElement(question, "questiontext", format="html")
        text_content = ET.SubElement(qtext, "text")
        enunciado_html = f"<p>{str(q.get('questiontext', '')).replace('\n', '<br>')}</p>"
        text_content.text = f"<![CDATA[{enunciado_html}]]>"
        
        # Configurações padrão do Moodle para Escolha Múltipla
        ET.SubElement(question, "defaultgrade").text = "1.0000000"
        ET.SubElement(question, "penalty").text = "0.3333333"
        ET.SubElement(question, "hidden").text = "0"
        ET.SubElement(question, "single").text = "true" 
        ET.SubElement(question, "shuffleanswers").text = "true"
        ET.SubElement(question, "answernumbering").text = "abc"
        
        # Mapeamento das 4 opções
        opcoes_mapeadas = {
            "a": q.get("choice_a", ""),
            "b": q.get("choice_b", ""),
            "c": q.get("choice_c", ""),
            "d": q.get("choice_d", "")
        }
        
        resposta_correta = str(q.get("answer", "")).lower().strip()
        
        for alinea, texto_opcao in opcoes_mapeadas.items():
            fraction = "100" if alinea == resposta_correta else "0"
            answer = ET.SubElement(question, "answer", fraction=fraction, format="html")
            ans_text = ET.SubElement(answer, "text")
            ans_text.text = f"<![CDATA[{texto_opcao}]]>"
            ET.SubElement(answer, "feedback", format="html").text = "<text></text>"

    xml_str = ET.tostring(quiz, encoding="utf-8")
    reparsed = minidom.parseString(xml_str)
    
    # Limpeza para garantir que as tags CDATA do Moodle não sofrem escape de caracteres
    return reparsed.toprettyxml(indent="  ").replace("&lt;", "<").replace("&gt;", ">")

def identificar_e_transformar_questoes(texto_bloco):
    # Expressão regular para apanhar cada subpergunta numerada com as suas opções
    # Procura por: Número + Texto com lacuna + Alíneas a, b, c, d
    padrao_questao = re.compile(
        r'(\d+)\.\s*(.*?_______.*?)\s*(a\).*?)(?=\n\d+\.|\Z)', 
        re.DOTALL | re.IGNORECASE
    )
    
    questoes_detetadas = padrao_questao.findall(texto_bloco)
    parsed_questions = []
    
    if questoes_detetadas:
        for num, enunciado, bloco_opcoes in questoes_detetadas:
            # Extrair as opções individuais (a, b, c, d)
            opcoes = re.findall(r'([a-d])\)\s*(.*?)(?=\s*[a-d]\)|\Z)', bloco_opcoes, re.DOTALL)
            
            opcoes_dict = {}
            resposta_correta = ""
            
            for alinea, texto_opcao in opcoes:
                texto_limpo = texto_opcao.strip()
                # Lógica para detetar a resposta correta (ex: se tiver um asterisco * ou se estiver a negrito no Word)
                # Se o teu modelo assume a primeira ou uma marcação específica:
                is_correct = False
                if "*" in texto_limpo: # Exemplo comum de marcação
                    texto_limpo = texto_limpo.replace("*", "").strip()
                    is_correct = True
                
                opcoes_dict[alinea] = texto_limpo
                if is_correct:
                    resposta_correta = alinea
            
            # Montar o dicionário no formato que o teu 'models.py' espera para multichoice
            parsed_questions.append({
                "type": "multichoice",
                "name": f"Pergunta {num}",
                "questiontext": f"{num}. {enunciado.strip()}",
                "choices": opcoes_dict,
                "answer": resposta_correta if resposta_correta else "a" # Fallback se não detetar
            })
            
    return parsed_questions

def gerar_moodle_xml_da_tabela(lista_perguntas):
    """
    Recebe os dados da tabela editável do Streamlit e gera o XML padrão do Moodle.
    """
    quiz = ET.Element("quiz")
    
    for q in lista_perguntas:
        question = ET.SubElement(quiz, "question", type="multichoice")
        
        # Nome da Pergunta
        name = ET.SubElement(question, "name")
        text_name = ET.SubElement(name, "text")
        text_name.text = q.get("name", "Pergunta de Escolha Múltipla")
        
        # Enunciado (Question Text)
        qtext = ET.SubElement(question, "questiontext", format="html")
        text_content = ET.SubElement(qtext, "text")
        # Transforma quebras de linha em <br> para o Moodle ler bem o HTML
        text_content.text = f"<![CDATA[<p>{q.get('questiontext', '').replace('\n', '<br>')}</p>]]>"
        
        # Definições padrão de feedback e notas
        ET.SubElement(question, "generalfeedback", format="html").text = "<text></text>"
        ET.SubElement(question, "defaultgrade").text = "1.0000000"
        ET.SubElement(question, "penalty").text = "0.3333333"
        ET.SubElement(question, "hidden").text = "0"
        ET.SubElement(question, "idnumber").text = ""
        ET.SubElement(question, "single").text = "true" # Apenas uma resposta correta
        ET.SubElement(question, "shuffleanswers").text = "true" # Baralhar opções
        ET.SubElement(question, "answernumbering").text = "abc" # Formato a), b), c)
        
        # Adicionar as Opções (Opções a, b, c, d)
        choices = q.get("choices", {})
        resposta_correta = q.get("answer", "").lower().strip()
        
        for alinea, texto_opcao in choices.items():
            # No Moodle, a resposta correta leva a nota (fraction) 100, as erradas levam 0
            fraction = "100" if alinea.lower().strip() == resposta_correta else "0"
            
            answer = ET.SubElement(question, "answer", fraction=fraction, format="html")
            ans_text = ET.SubElement(answer, "text")
            ans_text.text = f"<![CDATA[{texto_opcao}]]>"
            
            feedback = ET.SubElement(answer, "feedback", format="html")
            ET.SubElement(feedback, "text").text = ""

    # Formatar o XML de forma bonita (pretty print)
    xml_str = ET.tostring(quiz, encoding="utf-8")
    reparsed = minidom.parseString(xml_str)
    
    # O Moodle precisa das CDATA limpas (o minidom tende a escapar os caracteres, ajustamos aqui)
    return reparsed.toprettyxml(indent="  ").replace("&lt;", "<").replace("&gt;", ">")

def gerar_moodle_xml_da_triagem(lista_perguntas):
    """
    Gera a estrutura XML válida do Moodle com base nas perguntas mapeadas no Painel de Triagem.
    """
    quiz = ET.Element("quiz")
    
    for q in lista_perguntas:
        if not q.get("Exportar", True):
            continue
            
        tipo = q.get("Tipo", "Desconhecido")
        
        # Converte as perguntas que identificamos como Escolha Múltipla
        if tipo == "Escolha Múltipla":
            question = ET.SubElement(quiz, "question", type="multichoice")
            
            name = ET.SubElement(question, "name")
            ET.SubElement(name, "text").text = str(q.get("ID", "Pergunta"))
            
            qtext = ET.SubElement(question, "questiontext", format="html")
            text_content = ET.SubElement(qtext, "text")
            enunciado_html = f"<p>{str(q.get('Enunciado', '')).replace('\n', '<br>')}</p>"
            text_content.text = f"<![CDATA[{enunciado_html}]]>"
            
            ET.SubElement(question, "defaultgrade").text = "1.0000000"
            ET.SubElement(question, "penalty").text = "0.3333333"
            ET.SubElement(question, "hidden").text = "0"
            ET.SubElement(question, "single").text = "true" 
            ET.SubElement(question, "shuffleanswers").text = "true"
            ET.SubElement(question, "answernumbering").text = "abc"
            
            opcoes_mapeadas = {
                "a": q.get("Opção A", ""),
                "b": q.get("Opção B", ""),
                "c": q.get("Opção C", ""),
                "d": q.get("Opção D", "")
            }
            
            resposta_correta = str(q.get("Resposta Correta", "")).lower().strip()
            
            for alinea, texto_opcao in opcoes_mapeadas.items():
                if not str(texto_opcao).strip():
                    continue
                fraction = "100" if alinea == resposta_correta else "0"
                answer = ET.SubElement(question, "answer", fraction=fraction, format="html")
                ans_text = ET.SubElement(answer, "text")
                ans_text.text = f"<![CDATA[{texto_opcao}]]>"
                ET.SubElement(answer, "feedback", format="html").text = "<text></text>"
                
        # Fallback para os restantes formatos estruturados em Cloze ou texto descritivo
        else:
            tipo_moodle = "cloze" if "Lacunas" in tipo else "description"
            question = ET.SubElement(quiz, "question", type=tipo_moodle)
            
            name = ET.SubElement(question, "name")
            ET.SubElement(name, "text").text = str(q.get("ID", "Pergunta"))
            
            qtext = ET.SubElement(question, "questiontext", format="html")
            text_content = ET.SubElement(qtext, "text")
            text_content.text = f"<![CDATA[<p>{str(q.get('Enunciado', '')).replace('\n', '<br>')}</p>]]>"
            
            ET.SubElement(question, "defaultgrade").text = "1.0000000" if tipo_moodle == "cloze" else "0.0000000"
            ET.SubElement(question, "penalty").text = "0.3333333" if tipo_moodle == "cloze" else "0.0000000"
            ET.SubElement(question, "hidden").text = "0"

    xml_str = ET.tostring(quiz, encoding="utf-8")
    reparsed = minidom.parseString(xml_str)
    return reparsed.toprettyxml(indent="  ").replace("&lt;", "<").replace("&gt;", ">")
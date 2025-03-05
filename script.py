import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import time
from datetime import datetime
import difflib
import re
import numpy as np
from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Inches
import matplotlib.pyplot as plt
import xml.etree.ElementTree as ET
from xml.dom import minidom
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans

# =============================================================================
# Funções Auxiliares
# =============================================================================

def print_status(message):
    print(f"[STATUS] {message}")
    time.sleep(1)

def get_column_index(ws, column_name):
    for cell in ws[1]:
        if cell.value == column_name:
            return cell.column
    return None

def adjust_column_width(ws):
    for col in ws.columns:
        max_length = 0
        column = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width

def apply_header_style(ws):
    header_fill = PatternFill(start_color="000066", end_color="000066", fill_type="solid")
    header_font = Font(color="FFFFFF", size=12, bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

def apply_content_style(ws):
    content_font = Font(size=10)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = content_font

def apply_heatmap(ws, column_idx, values):
    values = pd.Series(values).dropna().sort_values()
    if values.empty:
        return
    
    total = len(values)
    red_threshold = int(total * 0.25)
    orange_threshold = int(total * 0.55)
    yellow_threshold = int(total * 0.85)
    
    red_max = values.iloc[red_threshold - 1] if red_threshold > 0 else values.min()
    orange_max = values.iloc[orange_threshold - 1] if orange_threshold > 0 else values.min()
    yellow_max = values.iloc[yellow_threshold - 1] if yellow_threshold > 0 else values.min()

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    white_font = Font(color="FFFFFF", size=10)
    black_font = Font(color="000000", size=10)

    for row in ws.iter_rows(min_row=2, min_col=column_idx, max_col=column_idx):
        for cell in row:
            if cell.value is not None and not pd.isna(cell.value):
                val = float(cell.value)
                if val <= red_max:
                    cell.fill = red_fill
                    cell.font = white_font
                elif val <= orange_max:
                    cell.fill = orange_fill
                    cell.font = white_font
                elif val <= yellow_max:
                    cell.fill = yellow_fill
                    cell.font = black_font
                else:
                    cell.fill = green_fill
                    cell.font = black_font

def fuzzy_match(str1, str2, threshold=0.8):
    return difflib.SequenceMatcher(None, str1.lower(), str2.lower()).ratio() >= threshold

def carregar_cidades_brasil():
    try:
        cidades_df = pd.read_excel('cidades_brasil.xlsx')
        return cidades_df['CIDADE'].str.lower().str.strip().tolist()
    except FileNotFoundError:
        print_status("Erro: Arquivo 'cidades_brasil.xlsx' não encontrado na pasta do script!")
        raise
    except KeyError:
        print_status("Erro: Coluna 'CIDADE' não encontrada na planilha 'cidades_brasil.xlsx'!")
        raise
    except Exception as e:
        print_status(f"Erro ao carregar cidades: {str(e)}")
        raise

def mapear_objetivo(intent):
    intent = str(intent).lower()
    if "informational" in intent:
        return "Mais Acessos"
    elif "transactional" in intent or "transacional" in intent:
        return "Vendas no E-commerce"
    elif "commercial" in intent:
        return "Captura de Leads"
    elif "navegacional" in intent:
        return "Branding/Autoridade"
    else:
        return "Outro"

# Funções para o Relatório
def add_title(doc, text):
    title = doc.add_heading(text, level=1)
    title.style.font.name = 'Arial'
    title.style.font.size = Pt(16)
    title.style.font.color.rgb = RGBColor(0, 0, 102)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

def add_subtitle(doc, text):
    subtitle = doc.add_heading(text, level=2)
    subtitle.style.font.name = 'Arial'
    subtitle.style.font.size = Pt(14)
    subtitle.style.font.color.rgb = RGBColor(0, 102, 204)

def add_paragraph(doc, text):
    p = doc.add_paragraph(text)
    p.style.font.name = 'Arial'
    p.style.font.size = Pt(12)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

def add_image(doc, image_path, width=Inches(5)):
    doc.add_picture(image_path, width=width)
    last_paragraph = doc.paragraphs[-1]
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

# =============================================================================
# Mapeamento da Jornada e Tipologia
# =============================================================================

def get_etapa_da_jornada(intent):
    intent = str(intent).strip().lower()
    if intent == "informational":
        return "Conscientização"
    elif intent in ["transactional", "transacional"]:
        return "Decisão"
    elif intent == "commercial":
        return "Consideração"
    elif intent == "navegacional":
        return "Fidelização"
    else:
        return "Sem Jornada Definida"

def get_tipologia_sugerida(row):
    intent = str(row.get('Intent', '')).strip().lower()
    serp = str(row.get('SERP Features', '')).strip().lower()
    
    if intent == "informational":
        if "featured snippets" in serp:
            return "Artigo de Blog"
        elif "instant answer" in serp:
            return "Artigo de Blog (respostas rápidas)"
        elif any(term in serp for term in ["video", "featured video", "video carousel"]):
            return "Guia"
        elif any(term in serp for term in ["image", "image pack"]):
            return "Infográfico"
        elif "people also ask" in serp:
            return "FAQs"
        elif "knowledge panel" in serp:
            return "Artigo de Blog (definições amplas)"
        elif any(term in serp for term in ["news", "top stories"]):
            return "Notícias do Setor"
        else:
            return "Análise Manual"
    elif intent in ["transactional", "transacional"]:
        if any(term in serp for term in ["shopping ads", "ads top", "ads bottom", "ads middle"]):
            return "Página de Produto/Serviço (otimizadas para conversão)"
        elif any(term in serp for term in ["hotel pack", "flights", "recipes", "jobs"]):
            return "Página de Produto/Serviço"
        elif "buying guide" in serp:
            return "Página de Produto/Serviço"
        elif any(term in serp for term in ["popular products", "related products", "organic carousel"]):
            return "Página de Produto/Serviço"
        elif any(term in serp for term in ["address pack", "twitter carousel"]):
            return "Página de Produto/Serviço"
        else:
            return "Análise Manual"
    elif intent == "commercial":
        if any(term in serp for term in ["featured reviews", "video carousel"]):
            return "Comparativo"
        elif "buying guide" in serp:
            return "Comparativo"
        elif "discussions and forums" in serp:
            return "Comparativo"
        elif any(term in serp for term in ["brands", "explore", "related searches", "related products"]):
            return "Comparativo"
        elif "questions and answers" in serp:
            return "FAQs"
        else:
            return "Análise Manual"
    elif intent == "navegacional":
        if "sitelinks" in serp:
            return "Documentação"
        elif "knowledge panel" in serp:
            return "Blog de Atualizações"
        elif any(term in serp for term in ["twitter", "twitter carousel"]):
            return "Blog de Atualizações"
        elif any(term in serp for term in ["find results on", "address pack"]):
            return "Página de Suporte"
        else:
            return "Análise Manual"
    else:
        return "Análise Manual"

# =============================================================================
# Função para Estratégia com Palavras-Chave
# =============================================================================

def criar_planilha_palavras_por_estrategia(folder_name, objective, combined_df):
    print_status("Criando a planilha 'Palavras por Estratégia.xlsx'...")
    
    dados_estrategia = [
        ["Captura de Leads", "Local SEO", "Melhor [serviço] em [cidade]", "Local Pack, Snippets", "Landing Page, Blog Local", "SEO Local + Conversão"],
        ["Vendas no E-commerce", "Transactional", "Comprar [produto] com desconto", "Shopping Ads, Reviews", "Páginas de Produto, Comparativos", "CRO + SEO para Produtos"],
        ["Mais Acessos", "Informational", "Como funciona [assunto]", "Featured Snippets, PAA", "Blog, Guia Completo", "SEO para Topo de Funil"],
        ["Monetização com Adsense", "High CPC", "Melhor seguro de saúde nos EUA", "Featured Snippets, PAA", "Lista Comparativa, Blog", "SEO para Alto CPC"],
        ["Branding/Autoridade", "Institucional", "[Empresa] é confiável?", "Knowledge Panel, Twitter Carousel", "Página Institucional, Blog", "SEO para Reputação"],
        ["Outro", "Custom", "Personalizado conforme análise", "Variável", "Variável", "Definido pelo usuário"]
    ]
    
    estrategia_base_df = pd.DataFrame(dados_estrategia, columns=["Objetivo", "Keyword Type", "Exemplo de Palavra-chave", "SERP Features", "Tipologia de Conteúdo", "Estratégia"])
    
    objetivo_map = {
        "1": "Captura de Leads",
        "2": "Vendas no E-commerce",
        "3": "Mais Acessos",
        "4": "Monetização com Adsense",
        "5": "Branding/Autoridade",
        "6": "Outro"
    }
    objetivo_selecionado = objetivo_map.get(objective, "Outro") if objective in "123456" else objective.capitalize()

    estrategia_df = combined_df.copy()
    estrategia_df["Objetivo"] = estrategia_df["Intent"].apply(mapear_objetivo)
    estrategia_df = estrategia_df.merge(estrategia_base_df.drop(columns=["Exemplo de Palavra-chave"]), 
                                        on="Objetivo", how="left", suffixes=('', '_base'))

    estrategia_df = estrategia_df.rename(columns={"Keyword": "Palavra-chave"})
    colunas_finais = ["Objetivo", "Palavra-chave", "Volume", "Intent", "SERP Features", 
                      "Keyword Type", "Tipologia de Conteúdo", "Estratégia"]
    estrategia_df = estrategia_df[colunas_finais]

    if objetivo_selecionado != "Outro":
        estrategia_df = estrategia_df[estrategia_df["Objetivo"] == objetivo_selecionado]

    wb = Workbook()
    ws = wb.active
    ws.title = "Palavras por Estratégia"
    for r in dataframe_to_rows(estrategia_df, index=False, header=True):
        ws.append(r)
    
    volume_col_index = get_column_index(ws, "Volume")
    if volume_col_index and not estrategia_df['Volume'].dropna().empty:
        apply_heatmap(ws, volume_col_index, estrategia_df['Volume'])

    apply_header_style(ws)
    apply_content_style(ws)
    adjust_column_width(ws)
    wb.save(os.path.join(folder_name, "Palavras por Estratégia.xlsx"))
    print_status("Planilha 'Palavras por Estratégia.xlsx' criada com sucesso!")
    return estrategia_df, objetivo_selecionado

# =============================================================================
# Função para Planejamento de Crescimento
# =============================================================================

def criar_planilha_planejamento_crescimento(folder_name, combined_df, volume_atual, crescimento_mensal, meses_planejamento, palavras_por_mes, objective, cidades_brasil):
    print_status("Criando a planilha 'Planejamento de Crescimento.xlsx'...")
    output_path = os.path.join(folder_name, "Planejamento de Crescimento.xlsx")

    ctr_rates = {
        1: (0.25, 0.35), 2: (0.15, 0.20), 3: (0.10, 0.15), 4: (0.07, 0.10), 5: (0.05, 0.07),
        6: (0.04, 0.06), 7: (0.03, 0.05), 8: (0.02, 0.04), 9: (0.02, 0.03), 10: (0.01, 0.02)
    }

    crescimento_absoluto = volume_atual * (crescimento_mensal / 100)
    acessos_alvo = volume_atual + crescimento_absoluto
    volume_min_por_palavra = crescimento_absoluto / (palavras_por_mes * ctr_rates[10][0])
    volume_max_por_palavra = crescimento_absoluto / (palavras_por_mes * ctr_rates[1][1])

    dados_calculo = [
        ["Volume Atual (mensal)", volume_atual],
        ["Crescimento Desejado (%)", crescimento_mensal],
        ["Crescimento Absoluto (acessos)", crescimento_absoluto],
        ["Acessos Alvo (mensal)", acessos_alvo],
        ["Meses de Planejamento", meses_planejamento],
        ["Palavras por Mês", palavras_por_mes],
        ["Volume Mínimo por Palavra (pior cenário - Posição 10)", round(volume_min_por_palavra)],
        ["Volume Máximo por Palavra (melhor cenário - Posição 1)", round(volume_max_por_palavra)]
    ]
    calculo_df = pd.DataFrame(dados_calculo, columns=["Métrica", "Valor"])

    objetivo_map = {
        "1": "Captura de Leads",
        "2": "Vendas no E-commerce",
        "3": "Mais Acessos",
        "4": "Monetização com Adsense",
        "5": "Branding/Autoridade",
        "6": "Outro"
    }
    objetivo_selecionado = objetivo_map.get(objective, "Outro") if objective in "123456" else objective.capitalize()

    palavras_df = combined_df.copy()
    palavras_df["Objetivo"] = palavras_df["Intent"].apply(mapear_objetivo)
    if objetivo_selecionado != "Outro":
        palavras_df = palavras_df[palavras_df["Objetivo"] == objetivo_selecionado]

    palavras_df["Keyword"] = palavras_df["Keyword"].fillna("")

    palavras_df = palavras_df[~palavras_df["Keyword"].str.lower().apply(
        lambda x: any(re.search(r'\b' + re.escape(cidade) + r'\b', x) for cidade in cidades_brasil) or 
                  bool(re.search(r'\b(em|no|na|de)\s+[a-záéíóúâêîôûãõç]+$', x, re.IGNORECASE)) or
                  bool(re.search(r'\b[a-z]{2,3}\b', x, re.IGNORECASE))
    )]

    print_status(f"Linhas após filtro de cidades: {len(palavras_df)}")
    if palavras_df.empty:
        print_status("Aviso: O DataFrame está vazio após o filtro de cidades!")
    else:
        print_status(f"Colunas disponíveis: {list(palavras_df.columns)}")

    volume_col = None
    for col in palavras_df.columns:
        if "volume" in col.lower():
            volume_col = col
            break
    if not volume_col:
        print_status("Erro: Nenhuma coluna de volume encontrada no DataFrame!")
        return None, None, None, None, None, None, None, None

    sort_columns = [volume_col]
    if "Competitive Density" in palavras_df.columns:
        sort_columns.append("Competitive Density")
        palavras_df = palavras_df.sort_values(by=sort_columns, ascending=[False, True])
    else:
        palavras_df = palavras_df.sort_values(by=volume_col, ascending=False)

    total_palavras = min(meses_planejamento * palavras_por_mes, len(palavras_df))
    palavras_selecionadas = palavras_df.head(total_palavras)
    colunas_selecao = ["Keyword", volume_col, "Intent", "SERP Features"] + \
                     (["Competitive Density"] if "Competitive Density" in palavras_df.columns else [])

    palavras_por_mes_ajustado = max(1, total_palavras // meses_planejamento)
    meses = []
    for i in range(meses_planejamento):
        inicio = i * palavras_por_mes_ajustado
        fim = min((i + 1) * palavras_por_mes_ajustado, total_palavras)
        mes_df = palavras_selecionadas.iloc[inicio:fim][colunas_selecao]
        meses.append(mes_df)

    palavras_df["Comprimento"] = palavras_df["Keyword"].apply(lambda x: len(str(x).split()))
    cauda_curta = palavras_df[palavras_df["Comprimento"] <= 2][colunas_selecao]
    cauda_media = palavras_df[palavras_df["Comprimento"] == 3][colunas_selecao]
    cauda_longa = palavras_df[palavras_df["Comprimento"] >= 4][colunas_selecao]

    total_grupos = palavras_por_mes // 2
    total_palavras_semantico = total_grupos * 10
    palavras_semantico = palavras_df.head(total_palavras_semantico)
    if len(palavras_semantico) >= 10:
        vectorizer = TfidfVectorizer()
        X = vectorizer.fit_transform(palavras_semantico["Keyword"])
        kmeans = KMeans(n_clusters=total_grupos, random_state=42)
        palavras_semantico["Grupo Semântico"] = kmeans.fit_predict(X).astype(str)
        palavras_semantico["Grupo Semântico"] = "Grupo " + palavras_semantico["Grupo Semântico"]
    else:
        palavras_semantico["Grupo Semântico"] = "Sem Grupo (poucas palavras)"
    colunas_semantico = ["Grupo Semântico"] + colunas_selecao

    palavras_blog = palavras_df[palavras_df["Intent"].str.lower().str.contains("informational", na=False)][colunas_selecao]

    wb = Workbook()

    ws_calculo = wb.active
    ws_calculo.title = "Calculo de Crescimento"
    for r in dataframe_to_rows(calculo_df, index=False, header=True):
        ws_calculo.append(r)
    apply_header_style(ws_calculo)
    apply_content_style(ws_calculo)
    adjust_column_width(ws_calculo)

    ws_selecao = wb.create_sheet("Selecao de Palavras")
    headers = ["Mês"] + colunas_selecao
    ws_selecao.append(headers)
    for i, mes_df in enumerate(meses, 1):
        for idx, row in mes_df.iterrows():
            ws_selecao.append([f"Mês {i}"] + row.tolist())
    volume_col_index = get_column_index(ws_selecao, volume_col)
    if volume_col_index and not palavras_selecionadas[volume_col].dropna().empty:
        apply_heatmap(ws_selecao, volume_col_index, palavras_selecionadas[volume_col])
    apply_header_style(ws_selecao)
    apply_content_style(ws_selecao)
    adjust_column_width(ws_selecao)

    ws_curta = wb.create_sheet("Cauda Curta")
    for r in dataframe_to_rows(cauda_curta, index=False, header=True):
        ws_curta.append(r)
    volume_col_index = get_column_index(ws_curta, volume_col)
    if volume_col_index and not cauda_curta[volume_col].dropna().empty:
        apply_heatmap(ws_curta, volume_col_index, cauda_curta[volume_col])
    apply_header_style(ws_curta)
    apply_content_style(ws_curta)
    adjust_column_width(ws_curta)

    ws_media = wb.create_sheet("Cauda Media")
    for r in dataframe_to_rows(cauda_media, index=False, header=True):
        ws_media.append(r)
    volume_col_index = get_column_index(ws_media, volume_col)
    if volume_col_index and not cauda_media[volume_col].dropna().empty:
        apply_heatmap(ws_media, volume_col_index, cauda_media[volume_col])
    apply_header_style(ws_media)
    apply_content_style(ws_media)
    adjust_column_width(ws_media)

    ws_longa = wb.create_sheet("Cauda Longa")
    for r in dataframe_to_rows(cauda_longa, index=False, header=True):
        ws_longa.append(r)
    volume_col_index = get_column_index(ws_longa, volume_col)
    if volume_col_index and not cauda_longa[volume_col].dropna().empty:
        apply_heatmap(ws_longa, volume_col_index, cauda_longa[volume_col])
    apply_header_style(ws_longa)
    apply_content_style(ws_longa)
    adjust_column_width(ws_longa)

    ws_semantico = wb.create_sheet("Grupos Semanticos")
    for r in dataframe_to_rows(palavras_semantico[colunas_semantico], index=False, header=True):
        ws_semantico.append(r)
    volume_col_index = get_column_index(ws_semantico, volume_col)
    if volume_col_index and not palavras_semantico[volume_col].dropna().empty:
        apply_heatmap(ws_semantico, volume_col_index, palavras_semantico[volume_col])
    apply_header_style(ws_semantico)
    apply_content_style(ws_semantico)
    adjust_column_width(ws_semantico)

    ws_blog = wb.create_sheet("Palavras para Blog")
    for r in dataframe_to_rows(palavras_blog, index=False, header=True):
        ws_blog.append(r)
    volume_col_index = get_column_index(ws_blog, volume_col)
    if volume_col_index and not palavras_blog[volume_col].dropna().empty:
        apply_heatmap(ws_blog, volume_col_index, palavras_blog[volume_col])
    apply_header_style(ws_blog)
    apply_content_style(ws_blog)
    adjust_column_width(ws_blog)

    wb.save(output_path)
    print_status("Planilha 'Planejamento de Crescimento.xlsx' criada com sucesso!")
    return calculo_df, palavras_selecionadas, cauda_curta, cauda_media, cauda_longa, palavras_semantico, palavras_blog, meses, colunas_selecao

# =============================================================================
# Configuração Inicial e Criação da Pasta de Saída
# =============================================================================

now = datetime.now()
folder_name = f"Analise {now.strftime('%d-%m-%Y')} {now.strftime('%H')} horas {now.strftime('%M')} minutos {now.strftime('%S')} segundos"
os.makedirs(folder_name, exist_ok=True)
print_status(f"Pasta de saída criada: {folder_name}")

print_status("Bem-vindo ao Script de Análise de Palavras-Chave para SEO!")
use_gpt = input("[PERGUNTA] Deseja conectar à API do ChatGPT para assistência? (s/n): ").lower()
if use_gpt == 's':
    print_status("A opção de API foi escolhida, mas este código usará o mapeamento interno para tipologia.")
objective = input(
    "[PERGUNTA] Qual o objetivo estratégico da análise?\n"
    "Opções: 1) Captura de leads, 2) Vendas no e-commerce, 3) Mais acessos, 4) Monetização com Adsense, 5) Branding/Autoridade, 6) Outro\n"
    "Digite o número ou descreva: "
)
volume_atual = int(input("[PERGUNTA] Qual o volume de acessos mensal atual do site? "))
crescimento_mensal = float(input("[PERGUNTA] Qual o percentual de crescimento desejado por mês? (ex: 10 para 10%): "))
meses_planejamento = int(input("[PERGUNTA] Quantos meses será o planejamento? "))
palavras_por_mes = int(input("[PERGUNTA] Quantas palavras-chave serão trabalhadas por mês? "))

folder_path = os.getcwd()
print_status(f"Usando a pasta atual como fonte das planilhas: {folder_path}")

print_status("Carregando lista de cidades do Brasil do arquivo 'cidades_brasil.xlsx'...")
cidades_brasil = carregar_cidades_brasil()
print_status(f"{len(cidades_brasil)} cidades carregadas para exclusão.")

# =============================================================================
# Fase 1 – Aglutinar as Planilhas
# =============================================================================

print_status("Iniciando Fase 1: Aglutinando planilhas...")
all_data = []
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') and not filename.startswith("Analise") and filename != "cidades_brasil.xlsx":
        print_status(f"Lendo arquivo: {filename}")
        try:
            df = pd.read_excel(os.path.join(folder_path, filename))
            all_data.append(df)
        except Exception as e:
            print_status(f"Erro ao ler {filename}: {str(e)}")
            raise
if not all_data:
    print_status("Erro: Nenhuma planilha .xlsx encontrada na pasta (exceto 'cidades_brasil.xlsx')!")
    raise ValueError("Nenhum arquivo válido encontrado")

try:
    combined_df = pd.concat(all_data, ignore_index=True)
    combined_df = combined_df.sort_values(by=['Keyword'], ascending=True)
except Exception as e:
    print_status(f"Erro ao concatenar planilhas: {str(e)}")
    raise

volume_col = None
for col in combined_df.columns:
    if "volume" in col.lower():
        volume_col = col
        break
if volume_col:
    combined_df = combined_df.sort_values(by=[volume_col], ascending=False)

wb = Workbook()
ws = wb.active
ws.title = "Visao Geral de Palavras"
for r in dataframe_to_rows(combined_df, index=False, header=True):
    ws.append(r)
volume_col_index = get_column_index(ws, volume_col) if volume_col else None
if volume_col_index and not combined_df[volume_col].dropna().empty:
    apply_heatmap(ws, volume_col_index, combined_df[volume_col])
apply_header_style(ws)
apply_content_style(ws)
adjust_column_width(ws)
visao_geral_filename = os.path.join(folder_name, "Visao Geral de Palavras.xlsx")
wb.save(visao_geral_filename)
print_status("Fase 1 concluída: Planilha 'Visao Geral de Palavras.xlsx' gerada!")

combined_df.to_excel(os.path.join(folder_name, "combined_df_temp.xlsx"), index=False)

plt.figure(figsize=(8, 4))
top_10 = combined_df.head(10)
bars = plt.bar(top_10['Keyword'], top_10[volume_col])
plt.title("Top 10 Palavras por Volume de Busca")
plt.xlabel("Palavras-chave")
plt.ylabel("Volume de Busca")
plt.xticks(rotation=45)
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, yval, int(yval), ha='center', va='bottom')
plt.tight_layout()
plt.savefig(os.path.join(folder_name, "visao_geral.png"))
plt.close()

# =============================================================================
# Fase 2 – Separação por Intent
# =============================================================================

print_status("Iniciando Fase 2: Separando por intenção de busca...")
intents = ['Informational', 'Transactional', 'Commercial', 'Navigational']
wb_intent = Workbook()
ws_overview = wb_intent.active
ws_overview.title = "Visao Geral"
for r in dataframe_to_rows(combined_df, index=False, header=True):
    ws_overview.append(r)
volume_col_index = get_column_index(ws_overview, volume_col) if volume_col else None
if volume_col_index and not combined_df[volume_col].dropna().empty:
    apply_heatmap(ws_overview, volume_col_index, combined_df[volume_col])
apply_header_style(ws_overview)
apply_content_style(ws_overview)
adjust_column_width(ws_overview)

intent_counts = {}
for intent in intents:
    print_status(f"Criando aba para Intent: {intent}")
    try:
        intent_df = combined_df[combined_df['Intent'].str.contains(intent, case=False, na=False)].sort_values(by=['Keyword'], ascending=True)
        if volume_col:
            intent_df = intent_df.sort_values(by=[volume_col], ascending=False)
        intent_counts[intent] = len(intent_df)
        ws_intent = wb_intent.create_sheet(intent)
        for r in dataframe_to_rows(intent_df, index=False, header=True):
            ws_intent.append(r)
        volume_col_index = get_column_index(ws_intent, volume_col) if volume_col else None
        if volume_col_index and not intent_df[volume_col].dropna().empty:
            apply_heatmap(ws_intent, volume_col_index, intent_df[volume_col])
        apply_header_style(ws_intent)
        apply_content_style(ws_intent)
        adjust_column_width(ws_intent)
    except Exception as e:
        print_status(f"Erro ao processar Intent '{intent}': {str(e)}")
        raise

no_intent_df = combined_df[combined_df['Intent'].isna()].sort_values(by=['Keyword'], ascending=True)
if volume_col:
    no_intent_df = no_intent_df.sort_values(by=[volume_col], ascending=False)
ws_no_intent = wb_intent.create_sheet("Sem Intent")
for r in dataframe_to_rows(no_intent_df, index=False, header=True):
    ws_no_intent.append(r)
volume_col_index = get_column_index(ws_no_intent, volume_col) if volume_col else None
if volume_col_index and not no_intent_df[volume_col].dropna().empty:
    apply_heatmap(ws_no_intent, volume_col_index, no_intent_df[volume_col])
apply_header_style(ws_no_intent)
apply_content_style(ws_no_intent)
adjust_column_width(ws_no_intent)

intents_filename = os.path.join(folder_name, "Intents.xlsx")
wb_intent.save(intents_filename)
print_status("Fase 2 concluída: Planilha 'Intents.xlsx' gerada!")

plt.figure(figsize=(6, 6))
intent_values = [intent_counts.get(i, 0) for i in intents]
plt.pie(intent_values, labels=intents, autopct='%1.1f%%', colors=['#FF9999', '#66B2FF', '#99FF99', '#FFCC99'])
plt.title("Distribuição por Intenção de Busca")
for i, (value, label) in enumerate(zip(intent_values, intents)):
    angle = sum(intent_values[:i]) + value / 2
    angle_rad = angle * 2 * np.pi / sum(intent_values)
    x = 0.5 * np.cos(angle_rad)
    y = 0.5 * np.sin(angle_rad)
    plt.text(x, y, str(value), ha='center', va='center')
plt.savefig(os.path.join(folder_name, "intents.png"))
plt.close()

# =============================================================================
# Fase 3 – Separação por SERP Features
# =============================================================================

print_status("Iniciando Fase 3: Separando por SERP Features...")
wb_serp = Workbook()
ws_serp_overview = wb_serp.active
ws_serp_overview.title = "Visao Geral"
for r in dataframe_to_rows(combined_df, index=False, header=True):
    ws_serp_overview.append(r)
volume_col_index = get_column_index(ws_serp_overview, volume_col) if volume_col else None
if volume_col_index and not combined_df[volume_col].dropna().empty:
    apply_heatmap(ws_serp_overview, volume_col_index, combined_df[volume_col])
apply_header_style(ws_serp_overview)
apply_content_style(ws_serp_overview)
adjust_column_width(ws_serp_overview)

serp_col = None
for col in combined_df.columns:
    if col.strip().lower() == "serp features":
        serp_col = col
        break

serp_counts = {}
if serp_col:
    features_set = set()
    for val in combined_df[serp_col].dropna():
        for token in str(val).split(','):
            token = token.strip()
            if token:
                features_set.add(token)
    for feature in features_set:
        print_status(f"Criando aba para SERP Feature: {feature}")
        feature_df = combined_df[combined_df[serp_col].str.contains(feature, case=False, na=False)]
        if volume_col:
            feature_df = feature_df.sort_values(by=[volume_col], ascending=False)
        serp_counts[feature] = len(feature_df)
        if not feature_df.empty:
            ws_feature = wb_serp.create_sheet(feature)
            for r in dataframe_to_rows(feature_df, index=False, header=True):
                ws_feature.append(r)
            volume_col_index = get_column_index(ws_feature, volume_col) if volume_col else None
            if volume_col_index and not feature_df[volume_col].dropna().empty:
                apply_heatmap(ws_feature, volume_col_index, feature_df[volume_col])
            apply_header_style(ws_feature)
            apply_content_style(ws_feature)
            adjust_column_width(ws_feature)
else:
    print_status("Aviso: Coluna 'SERP Features' não encontrada. Pulando separação por SERP Features.")
serp_features_filename = os.path.join(folder_name, "SERP Features.xlsx")
wb_serp.save(serp_features_filename)
print_status("Fase 3 concluída: Planilha 'SERP Features.xlsx' gerada!")

top_serp = sorted(serp_counts.items(), key=lambda x: x[1], reverse=True)[:4]
plt.figure(figsize=(8, 4))
bars = plt.bar([x[0] for x in top_serp], [x[1] for x in top_serp])
plt.title("Principais Recursos de SERP")
plt.xlabel("Recurso")
plt.ylabel("Quantidade")
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, yval, int(yval), ha='center', va='bottom')
plt.savefig(os.path.join(folder_name, "serp_features.png"))
plt.close()

# =============================================================================
# Fase 4 – Mapeamento por Jornada e Tipologia
# =============================================================================

print_status("Iniciando Fase 4: Mapeando por Jornada e Tipologia...")
jornada_list = []
tipologia_list = []
for idx, row in combined_df.iterrows():
    jornada_list.append(get_etapa_da_jornada(row.get('Intent', '')))
    tipologia_list.append(get_tipologia_sugerida(row))
combined_df['Etapa da Jornada'] = jornada_list
combined_df['Tipologia Sugerida'] = tipologia_list

cols_to_drop = ["CPC (USD)", "Competitive Density", "Number of Results"]
combined_df = combined_df.drop(columns=cols_to_drop, errors='ignore')

if volume_col:
    combined_df = combined_df.sort_values(by=[volume_col], ascending=False)
else:
    combined_df = combined_df.sort_values(by=['Keyword'], ascending=True)

wb_journey = Workbook()
ws_journey_overview = wb_journey.active
ws_journey_overview.title = "Visao Geral da Jornada"
for r in dataframe_to_rows(combined_df, index=False, header=True):
    ws_journey_overview.append(r)
volume_col_index = get_column_index(ws_journey_overview, volume_col) if volume_col else None
if volume_col_index and not combined_df[volume_col].dropna().empty:
    apply_heatmap(ws_journey_overview, volume_col_index, combined_df[volume_col])
apply_header_style(ws_journey_overview)
apply_content_style(ws_journey_overview)
adjust_column_width(ws_journey_overview)

etapas = ["Conscientização", "Consideração", "Decisão", "Fidelização", "Sem Jornada Definida"]
jornada_counts = {}
for etapa in etapas:
    print_status(f"Criando aba para Jornada: {etapa}")
    etapa_df = combined_df[combined_df['Etapa da Jornada'] == etapa]
    if volume_col:
        etapa_df = etapa_df.sort_values(by=[volume_col], ascending=False)
    jornada_counts[etapa] = len(etapa_df)
    ws_etapa = wb_journey.create_sheet(etapa)
    for r in dataframe_to_rows(etapa_df, index=False, header=True):
        ws_etapa.append(r)
    volume_col_index = get_column_index(ws_etapa, volume_col) if volume_col else None
    if volume_col_index and not etapa_df[volume_col].dropna().empty:
        apply_heatmap(ws_etapa, volume_col_index, etapa_df[volume_col])
    apply_header_style(ws_etapa)
    apply_content_style(ws_etapa)
    adjust_column_width(ws_etapa)

jornada_filename = os.path.join(folder_name, "Jornada e Tipologias.xlsx")
wb_journey.save(jornada_filename)
print_status("Fase 4 concluída: Planilha 'Jornada e Tipologias.xlsx' gerada!")

plt.figure(figsize=(6, 6))
jornada_values = [jornada_counts[e] for e in etapas if e in jornada_counts]
plt.pie(jornada_values, labels=[e for e in etapas if e in jornada_counts], autopct='%1.1f%%', colors=['#FF6666', '#FFCC66', '#66CCFF', '#66FF66', '#999999'])
plt.title("Distribuição por Etapa da Jornada")
for i, (value, label) in enumerate(zip(jornada_values, [e for e in etapas if e in jornada_counts])):
    angle = sum(jornada_values[:i]) + value / 2
    angle_rad = angle * 2 * np.pi / sum(jornada_values)
    x = 0.5 * np.cos(angle_rad)
    y = 0.5 * np.sin(angle_rad)
    plt.text(x, y, str(value), ha='center', va='center')
plt.savefig(os.path.join(folder_name, "jornada.png"))
plt.close()

# =============================================================================
# Fase 5 – CTR por Posição
# =============================================================================

print_status("Iniciando Fase 5: Calculando CTR por posição...")
ctr_rates = {
    1: (0.25, 0.35), 2: (0.15, 0.20), 3: (0.10, 0.15), 4: (0.07, 0.10), 5: (0.05, 0.07),
    6: (0.04, 0.06), 7: (0.03, 0.05), 8: (0.02, 0.04), 9: (0.02, 0.03), 10: (0.01, 0.02)
}
if volume_col:
    ctr_df = combined_df[(combined_df[volume_col] > 0) & (combined_df[volume_col].notna())].copy()
    for pos, (min_rate, max_rate) in ctr_rates.items():
        ctr_df[f'Posicao {pos}'] = ctr_df[volume_col].apply(lambda x: f"{int(x * min_rate)} - {int(x * max_rate)}")
    selected_columns = ['Keyword', volume_col, 'Intent', 'Trend'] + [f'Posicao {i}' for i in range(1, 11)]
    ctr_export_df = ctr_df[selected_columns].copy()
    for pos, (min_rate, max_rate) in ctr_rates.items():
        old_col = f'Posicao {pos}'
        new_col = f'Posicao {pos} ({int(min_rate*100)}%-{int(max_rate*100)}%)'
        ctr_export_df.rename(columns={old_col: new_col}, inplace=True)
    wb_ctr = Workbook()
    ws_ctr = wb_ctr.active
    ws_ctr.title = "CTR por Posicao"
    for r in dataframe_to_rows(ctr_export_df, index=False, header=True):
        ws_ctr.append(r)
    volume_col_index = get_column_index(ws_ctr, volume_col)
    if volume_col_index and not ctr_export_df[volume_col].dropna().empty:
        apply_heatmap(ws_ctr, volume_col_index, ctr_export_df[volume_col])
    apply_header_style(ws_ctr)
    apply_content_style(ws_ctr)
    adjust_column_width(ws_ctr)
    ctr_filename = os.path.join(folder_name, "CTR por Posicao.xlsx")
    wb_ctr.save(ctr_filename)
    print_status("Fase 5 concluída: Planilha 'CTR por Posicao.xlsx' gerada!")
else:
    print_status("Aviso: Nenhuma coluna de volume encontrada. Pulando Fase 5.")
    ctr_export_df = pd.DataFrame()

if not ctr_export_df.empty:
    exemplo_ctr = ctr_export_df.iloc[0]
    plt.figure(figsize=(8, 4))
    positions = range(1, 11)
    ctr_min = [float(exemplo_ctr[f'Posicao {i} ({int(min_rate*100)}%-{int(max_rate*100)}%)'].split('-')[0]) for i, (min_rate, max_rate) in ctr_rates.items()]
    ctr_max = [float(exemplo_ctr[f'Posicao {i} ({int(min_rate*100)}%-{int(max_rate*100)}%)'].split('-')[1]) for i, (min_rate, max_rate) in ctr_rates.items()]
    plt.plot(positions, ctr_min, label="Cenário Pessimista", marker='o')
    plt.plot(positions, ctr_max, label="Cenário Otimista", marker='o')
    plt.title(f"Estimativa de CTR por Posição ({exemplo_ctr['Keyword']})")
    plt.xlabel("Posição")
    plt.ylabel("Cliques")
    for i, (min_val, max_val) in enumerate(zip(ctr_min, ctr_max)):
        plt.text(positions[i], min_val, int(min_val), ha='center', va='bottom')
        plt.text(positions[i], max_val, int(max_val), ha='center', va='bottom')
    plt.legend()
    plt.savefig(os.path.join(folder_name, "ctr_posicao.png"))
    plt.close()

# =============================================================================
# Fase 6 – Estratégia por Objetivo com Palavras-Chave
# =============================================================================

print_status("Iniciando Fase 6: Gerando estratégias por objetivo com palavras-chave...")
estrategia_df, objetivo_selecionado = criar_planilha_palavras_por_estrategia(folder_name, objective, combined_df)
print_status("Fase 6 concluída: Planilha 'Palavras por Estratégia.xlsx' gerada!")

# =============================================================================
# Fase 7 – Planejamento de Crescimento
# =============================================================================

print_status("Iniciando Fase 7: Gerando planejamento de crescimento...")
result = criar_planilha_planejamento_crescimento(folder_name, combined_df, volume_atual, crescimento_mensal, meses_planejamento, palavras_por_mes, objective, cidades_brasil)
if result is None:
    print_status("Erro na Fase 7. Abortando execução.")
    raise ValueError("Fase 7 falhou devido à ausência de coluna de volume ou outro erro.")
calculo_df, palavras_selecionadas, cauda_curta, cauda_media, cauda_longa, palavras_semantico, palavras_blog, meses, colunas_selecao = result
print_status("Fase 7 concluída: Planilha 'Planejamento de Crescimento.xlsx' gerada!")

plt.figure(figsize=(8, 4))
meses_grafico = [f'Mês {i+1}' for i in range(meses_planejamento)]
acessos = [volume_atual * (1 + crescimento_mensal / 100) ** i for i in range(meses_planejamento)]
plt.plot(meses_grafico, acessos, marker='o')
plt.title("Projeção de Crescimento")
plt.xlabel("Meses")
plt.ylabel("Acessos Mensais")
for i, val in enumerate(acessos):
    plt.text(i, val, int(val), ha='center', va='bottom')
plt.savefig(os.path.join(folder_name, "crescimento.png"))
plt.close()

# =============================================================================
# Geração do Relatório
# =============================================================================

print_status("Gerando Relatório.docx...")
doc = Document()

add_title(doc, "Relatório de Análise de Palavras-Chave para SEO")

add_subtitle(doc, "Introdução")
add_paragraph(doc, "Este relatório apresenta uma análise detalhada das palavras-chave fornecidas, com o objetivo de otimizar a estratégia de SEO do seu site. Foram realizadas sete fases analíticas, cada uma com um propósito específico para entender o comportamento de busca, o potencial de tráfego e as oportunidades de conteúdo. Abaixo, detalhamos cada fase, os motivos de sua execução e os resultados obtidos.")

add_subtitle(doc, "Fase 1: Visão Geral de Palavras")
add_paragraph(doc, "Objetivo: Agregar todas as palavras-chave de diferentes fontes em uma única planilha para fornecer uma visão consolidada do volume de busca, intenção e características de SERP.")
add_paragraph(doc, "Por que fizemos isso: A consolidação permite identificar padrões gerais, como palavras-chave de alto volume ou com baixa concorrência, servindo como base para as análises subsequentes.")
add_paragraph(doc, f"Resultado: Foram analisadas {len(combined_df)} palavras-chave. Abaixo, o gráfico mostra as 10 principais por volume de busca.")
add_image(doc, os.path.join(folder_name, "visao_geral.png"))

add_subtitle(doc, "Fase 2: Separação por Intenção de Busca")
add_paragraph(doc, "Objetivo: Classificar as palavras-chave em categorias de intenção (Informacional, Transacional, Comercial, Navegacional) para alinhar o conteúdo às expectativas dos usuários.")
add_paragraph(doc, "Por que fizemos isso: Diferentes intenções requerem abordagens distintas de conteúdo, otimizando a conversão.")
add_paragraph(doc, "Resultado: Distribuição das intenções encontrada nas palavras-chave.")
add_image(doc, os.path.join(folder_name, "intents.png"))

add_subtitle(doc, "Fase 3: Separação por Recursos de SERP")
add_paragraph(doc, "Objetivo: Identificar palavras-chave associadas a recursos específicos da SERP (ex.: Snippets, Local Pack) para explorar oportunidades de destaque.")
add_paragraph(doc, "Por que fizemos isso: Recursos de SERP aumentam a visibilidade e a taxa de cliques (CTR).")
add_paragraph(doc, "Resultado: Quantidade de palavras por recurso principal.")
add_image(doc, os.path.join(folder_name, "serp_features.png"))

add_subtitle(doc, "Fase 4: Mapeamento por Jornada e Tipologia")
add_paragraph(doc, "Objetivo: Mapear as palavras-chave às etapas da jornada do cliente e sugerir tipologias de conteúdo.")
add_paragraph(doc, "Por que fizemos isso: Alinhar o conteúdo à jornada melhora a experiência do usuário e a eficácia do SEO.")
add_paragraph(doc, "Resultado: Distribuição por etapa da jornada.")
add_image(doc, os.path.join(folder_name, "jornada.png"))

add_subtitle(doc, "Fase 5: CTR por Posição")
add_paragraph(doc, "Objetivo: Calcular a estimativa de cliques (CTR) por posição no ranking para cada palavra-chave.")
add_paragraph(doc, "Por que fizemos isso: Entender o impacto do posicionamento ajuda a definir metas de tráfego orgânico.")
if not ctr_export_df.empty:
    add_paragraph(doc, f"Resultado: Exemplo de CTR para '{exemplo_ctr['Keyword']}' com volume {exemplo_ctr[volume_col]}.")
    add_image(doc, os.path.join(folder_name, "ctr_posicao.png"))
else:
    add_paragraph(doc, "Resultado: Nenhuma análise de CTR devido à ausência de coluna de volume.")

add_subtitle(doc, "Fase 6: Estratégia por Objetivo com Palavras-Chave")
add_paragraph(doc, "Objetivo: Associar palavras-chave a estratégias específicas com tipologias de conteúdo e ações recomendadas.")
add_paragraph(doc, "Por que fizemos isso: Personalizar a abordagem por objetivo maximiza os resultados desejados.")
add_paragraph(doc, f"Resultado: Exemplo para '{estrategia_df.iloc[0]['Palavra-chave']}' com volume {estrategia_df.iloc[0]['Volume']}: Estratégia '{estrategia_df.iloc[0]['Estratégia']}', Tipologia '{estrategia_df.iloc[0]['Tipologia de Conteúdo']}'.")

add_subtitle(doc, "Fase 7: Planejamento de Crescimento")
add_paragraph(doc, "Objetivo: Projetar o crescimento de tráfego orgânico com base no volume atual, meta de crescimento, e palavras-chave trabalhadas por mês, excluindo termos geográficos.")
add_paragraph(doc, "Por que fizemos isso: Fornece um plano acionável com metas mensais e priorização de palavras-chave.")
add_paragraph(doc, f"Resultado: Projeção para {meses_planejamento} meses com volume inicial de {volume_atual} e crescimento de {crescimento_mensal}% ao mês.")
add_image(doc, os.path.join(folder_name, "crescimento.png"))

add_subtitle(doc, "Conclusão")
add_paragraph(doc, "A análise realizada oferece uma visão estratégica das palavras-chave, permitindo otimizar o conteúdo e planejar o crescimento do tráfego orgânico. Recomendamos implementar as tipologias sugeridas e priorizar as palavras-chave do planejamento de crescimento.")

doc.save(os.path.join(folder_name, "Relatório.docx"))
print_status("Relatório.docx gerado com sucesso na pasta " + folder_name)

# =============================================================================
# Geração do XML
# =============================================================================

print_status("Gerando resultados_finais.xml...")

def dict_to_xml(tag, d):
    elem = ET.Element(tag)
    for key, val in d.items():
        if isinstance(val, dict):
            child = dict_to_xml(key, val)
            elem.append(child)
        elif isinstance(val, list):
            for i, item in enumerate(val):
                if isinstance(item, dict):
                    child = dict_to_xml(key, item)
                else:
                    child = ET.Element(key)
                    child.text = str(item)
                elem.append(child)
        else:
            child = ET.Element(key)
            child.text = str(val)
            elem.append(child)
    return elem

resultados = {
    "Visao_Geral": {
        "total_palavras": len(combined_df),
        "top_palavras": combined_df.head(10)[['Keyword', volume_col, 'Intent']].to_dict(orient='records') if volume_col else []
    },
    "Intents": {
        intent: {
            "count": intent_counts.get(intent, 0),
            "exemplo": combined_df[combined_df['Intent'].str.contains(intent, case=False, na=False)]['Keyword'].iloc[0] if intent_counts.get(intent, 0) > 0 and not combined_df[combined_df['Intent'].str.contains(intent, case=False, na=False)].empty else ""
        } for intent in intents
    },
    "SERP_Features": {feature: {"count": count} for feature, count in serp_counts.items()},
    "Jornada": {
        etapa: {
            "count": jornada_counts.get(etapa, 0),
            "tipologia": combined_df[combined_df['Etapa da Jornada'] == etapa]['Tipologia Sugerida'].iloc[0] if jornada_counts.get(etapa, 0) > 0 and not combined_df[combined_df['Etapa da Jornada'] == etapa].empty else ""
        } for etapa in etapas
    },
    "CTR_por_Posicao": {
        "exemplo_palavra": ctr_export_df.iloc[0].to_dict() if not ctr_export_df.empty else {}
    },
    "Palavras_por_Estrategia": {
        objetivo_selecionado: estrategia_df.head(5).to_dict(orient='records')
    },
    "Planejamento_Crescimento": {
        "calculo": calculo_df.to_dict(orient='records'),
        "selecao_palavras": {f"Mes_{i+1}": mes.to_dict(orient='records') for i, mes in enumerate(meses)},
        "cauda_curta": {"count": len(cauda_curta), "exemplo": cauda_curta.head(1).to_dict(orient='records') if not cauda_curta.empty else []},
        "cauda_media": {"count": len(cauda_media), "exemplo": cauda_media.head(1).to_dict(orient='records') if not cauda_media.empty else []},
        "cauda_longa": {"count": len(cauda_longa), "exemplo": cauda_longa.head(1).to_dict(orient='records') if not cauda_longa.empty else []},
        "grupos_semanticos": palavras_semantico.groupby("Grupo Semântico").apply(lambda x: x[colunas_selecao].head(5).to_dict(orient='records')).to_dict() if not palavras_semantico.empty else {},
        "palavras_blog": palavras_blog.head(5).to_dict(orient='records') if not palavras_blog.empty else []
    }
}

root = dict_to_xml("Resultados", resultados)
xml_str = ET.tostring(root, encoding='utf-8', method='xml')
pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="  ")

with open(os.path.join(folder_name, "resultados_finais.xml"), "w", encoding="utf-8") as xml_file:
    xml_file.write(pretty_xml)
print_status("resultados_finais.xml gerado com sucesso na pasta " + folder_name)

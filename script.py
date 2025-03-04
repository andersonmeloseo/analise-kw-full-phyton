import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import time
from datetime import datetime
import difflib

# =============================================================================
# Funções Auxiliares
# =============================================================================

def print_status(message):
    """Exibe mensagem de status com uma pequena pausa."""
    print(f"[STATUS] {message}")
    time.sleep(1)

def get_column_index(ws, column_name):
    """Retorna o índice da coluna a partir do cabeçalho da planilha."""
    for cell in ws[1]:
        if cell.value == column_name:
            return cell.column
    return None

def apply_heatmap(ws, column_idx, min_val, max_val):
    """Aplica um mapa de calor na coluna indicada: laranja para valores menores e verde para maiores."""
    for row in ws.iter_rows(min_row=2, min_col=column_idx, max_col=column_idx):
        for cell in row:
            if cell.value is not None and not pd.isna(cell.value):
                ratio = (cell.value - min_val) / (max_val - min_val) if max_val != min_val else 0
                green = int(255 * ratio)
                red = int(255 * (1 - ratio))
                fill = PatternFill(start_color=f'FF{red:02x}{green:02x}00',
                                   end_color=f'FF{red:02x}{green:02x}00',
                                   fill_type='solid')
                cell.fill = fill

def fuzzy_match(str1, str2, threshold=0.8):
    """Retorna True se a similaridade entre as strings for maior ou igual ao limiar."""
    return difflib.SequenceMatcher(None, str1.lower(), str2.lower()).ratio() >= threshold

# =============================================================================
# Mapeamento da Jornada e Tipologia – utilizando as regras fornecidas
# =============================================================================

def get_etapa_da_jornada(intent):
    """Define a etapa da jornada com base na Intent (case insensitive)."""
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
    """
    Retorna a tipologia recomendada com base na Intent e nos SERP features,
    conforme as regras fornecidas.
    """
    intent = str(row.get('Intent', '')).strip().lower()
    serp = str(row.get('SERP Features', '')).strip().lower()  # Note que aqui usamos 'SERP Features'
    
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
# Configuração Inicial e Criação da Pasta de Saída
# =============================================================================

now = datetime.now()
folder_name = f"Analise {now.strftime('%d-%m-%Y')} {now.strftime('%H')} horas {now.strftime('%M')} minutos {now.strftime('%S')} segundos"
os.makedirs(folder_name, exist_ok=True)
print_status(f"Pasta de saída criada: {folder_name}")

print_status("Bem-vindo ao Script de Análise de Palavras-Chave para SEO!")
use_gpt = input("[PERGUNTA] Deseja conectar à API do ChatGPT para assistência? (s/n): ").lower()
# Nesta versão usaremos o mapeamento interno para tipologia.
if use_gpt == 's':
    print_status("A opção de API foi escolhida, mas este código usará o mapeamento interno para tipologia.")
objective = input(
    "[PERGUNTA] Qual o objetivo estratégico da análise?\n"
    "Opções: 1) Captura de leads, 2) Vendas no e-commerce, 3) Mais acessos, 4) Monetização com Adsense, 5) Branding/Autoridade, 6) Outro\n"
    "Digite o número ou descreva: "
)

folder_path = os.getcwd()
print_status(f"Usando a pasta atual como fonte das planilhas: {folder_path}")

# =============================================================================
# Passo 1 – Aglutinar as Planilhas
# =============================================================================

print_status("Iniciando Passo 1: Aglutinando planilhas...")
all_data = []
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') and not filename.startswith("Analise"):
        print_status(f"Lendo arquivo: {filename}")
        df = pd.read_excel(os.path.join(folder_path, filename))
        all_data.append(df)
if not all_data:
    print_status("Erro: Nenhuma planilha .xlsx encontrada na pasta!")
    exit()
combined_df = pd.concat(all_data, ignore_index=True)
combined_df = combined_df.sort_values(by=['Keyword', 'Volume'], ascending=[True, False])

wb = Workbook()
ws = wb.active
ws.title = "Visao Geral de Palavras"
for r in dataframe_to_rows(combined_df, index=False, header=True):
    ws.append(r)
volume_col_index = get_column_index(ws, "Volume")
volumes = combined_df['Volume'].dropna()
if volume_col_index and not volumes.empty:
    apply_heatmap(ws, volume_col_index, volumes.min(), volumes.max())
else:
    print_status("Aviso: Nenhum valor válido na coluna Volume para aplicar mapa de calor.")
visao_geral_filename = os.path.join(folder_name, "Visao Geral de Palavras.xlsx")
wb.save(visao_geral_filename)
print_status("Passo 1 concluído: Planilha 'Visao Geral de Palavras.xlsx' gerada!")

# =============================================================================
# Passo 2 – Separação por Intent
# =============================================================================

print_status("Iniciando Passo 2: Separando por intenção de busca...")
intents = ['Informational', 'Transactional', 'Commercial', 'Navigational']
wb_intent = Workbook()
ws_overview = wb_intent.active
ws_overview.title = "Visao Geral"
for r in dataframe_to_rows(combined_df, index=False, header=True):
    ws_overview.append(r)
for intent in intents:
    print_status(f"Criando aba para Intent: {intent}")
    intent_df = combined_df[combined_df['Intent'].str.contains(intent, case=False, na=False)]\
                .sort_values(by=['Keyword', 'Volume'], ascending=[True, False])
    ws_intent = wb_intent.create_sheet(intent)
    for r in dataframe_to_rows(intent_df, index=False, header=True):
        ws_intent.append(r)
no_intent_df = combined_df[combined_df['Intent'].isna()]\
                .sort_values(by=['Keyword', 'Volume'], ascending=[True, False])
ws_no_intent = wb_intent.create_sheet("Sem Intent")
for r in dataframe_to_rows(no_intent_df, index=False, header=True):
    ws_no_intent.append(r)
intents_filename = os.path.join(folder_name, "Intents.xlsx")
wb_intent.save(intents_filename)
print_status("Passo 2 concluído: Planilha 'Intents.xlsx' gerada!")

# =============================================================================
# Passo 3 – Separação por SERP Features
# =============================================================================

print_status("Iniciando Passo 3: Separando por SERP Features...")
wb_serp = Workbook()
ws_serp_overview = wb_serp.active
ws_serp_overview.title = "Visao Geral"
for r in dataframe_to_rows(combined_df, index=False, header=True):
    ws_serp_overview.append(r)
# Procura pela coluna SERP Features ignorando caixa (case insensitive)
serp_col = None
for col in combined_df.columns:
    if col.strip().lower() == "serp features":
        serp_col = col
        break

if serp_col:
    # Cria um conjunto com todos os tokens encontrados na coluna
    features_set = set()
    for val in combined_df[serp_col].dropna():
        for token in str(val).split(','):
            token = token.strip()
            if token:
                features_set.add(token)
    # Para cada feature encontrada, cria uma aba com os registros que contenham essa feature
    for feature in features_set:
        print_status(f"Criando aba para SERP Feature: {feature}")
        feature_df = combined_df[combined_df[serp_col].str.contains(feature, case=False, na=False)]
        if not feature_df.empty:
            ws_feature = wb_serp.create_sheet(feature)
            for r in dataframe_to_rows(feature_df, index=False, header=True):
                ws_feature.append(r)
else:
    print_status("Aviso: Coluna 'SERP Features' não encontrada. Pulando separação por SERP Features.")
serp_features_filename = os.path.join(folder_name, "SERP Features.xlsx")
wb_serp.save(serp_features_filename)
print_status("Passo 3 concluído: Planilha 'SERP Features.xlsx' gerada!")

# =============================================================================
# Passo 4 – Mapeamento por Jornada e Tipologia
# =============================================================================

print_status("Iniciando Passo 4: Mapeando por Jornada e Tipologia...")

jornada_list = []
tipologia_list = []
for idx, row in combined_df.iterrows():
    jornada_list.append(get_etapa_da_jornada(row.get('Intent', '')))
    tipologia_list.append(get_tipologia_sugerida(row))
combined_df['Etapa da Jornada'] = jornada_list
combined_df['Tipologia Sugerida'] = tipologia_list

# Remove as colunas indesejadas, se existirem
cols_to_drop = ["CPC (USD)", "Competitive Density", "Number of Results"]
combined_df = combined_df.drop(columns=cols_to_drop, errors='ignore')

combined_df = combined_df.sort_values(by=['Keyword', 'Volume'], ascending=[True, False])

wb_journey = Workbook()
ws_journey_overview = wb_journey.active
ws_journey_overview.title = "Visao Geral da Jornada"
for r in dataframe_to_rows(combined_df, index=False, header=True):
    ws_journey_overview.append(r)

etapas = ["Conscientização", "Consideração", "Decisão", "Fidelização", "Sem Jornada Definida"]
for etapa in etapas:
    print_status(f"Criando aba para Jornada: {etapa}")
    etapa_df = combined_df[combined_df['Etapa da Jornada'] == etapa]
    ws_etapa = wb_journey.create_sheet(etapa)
    for r in dataframe_to_rows(etapa_df, index=False, header=True):
        ws_etapa.append(r)
    if not etapa_df['Volume'].dropna().empty:
        volume_idx = None
        for cell in ws_etapa[1]:
            if cell.value == "Volume":
                volume_idx = cell.column
                break
        if volume_idx:
            apply_heatmap(ws_etapa, volume_idx, etapa_df['Volume'].dropna().min(), etapa_df['Volume'].dropna().max())

jornada_filename = os.path.join(folder_name, "Jornada e Tipologias.xlsx")
wb_journey.save(jornada_filename)
print_status("Passo 4 concluído: Planilha 'Jornada e Tipologias.xlsx' gerada!")

# =============================================================================
# Passo 5 – CTR por Posição
# =============================================================================

print_status("Iniciando Passo 5: Calculando CTR por posição...")
ctr_rates = {
    1: (0.25, 0.35), 2: (0.15, 0.20), 3: (0.10, 0.15), 4: (0.07, 0.10), 5: (0.05, 0.07),
    6: (0.04, 0.06), 7: (0.03, 0.05), 8: (0.02, 0.04), 9: (0.02, 0.03), 10: (0.01, 0.02)
}
ctr_df = combined_df[(combined_df['Volume'] > 0) & (combined_df['Volume'].notna())].copy()
for pos, (min_rate, max_rate) in ctr_rates.items():
    ctr_df[f'Posicao {pos}'] = ctr_df['Volume'].apply(lambda x: f"{int(x * min_rate)} - {int(x * max_rate)}")
selected_columns = ['Keyword', 'Volume', 'Intent', 'Trend'] + [f'Posicao {i}' for i in range(1, 11)]
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
ctr_filename = os.path.join(folder_name, "CTR por Posicao.xlsx")
wb_ctr.save(ctr_filename)
print_status("Passo 5 concluído: Planilha 'CTR por Posicao.xlsx' gerada!")

print_status("Análise concluída com sucesso! Verifique as planilhas na pasta " + folder_name)

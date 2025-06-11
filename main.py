from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel
import pandas as pd
import re
import os
import time
from io import BytesIO
import tempfile
import logging
import numpy as np

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

# Configurar CORS para permitir acesso de origens diferentes
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Permite todas as origens em ambiente de desenvolvimento
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Modelo para entrada de códigos individuais
class CodeInput(BaseModel):
    codes: list[str]

# Função para normalizar o código
def normalize_code(code):
    if not isinstance(code, str):
        code = str(code)
    return code.replace(".", "")

# Função para verificar a qual descrição o código (NCM ou NBS) pertence
def classify_code(code):
    code_normalized = normalize_code(code)

    # Item 1: Biofertilizantes (3101.00.00)
    if code_normalized == "31010000":
        return "Biofertilizantes, em conformidade com as definições e demais requisitos da legislação específica"

    # Item 4: Inoculantes (3002.49, 3002.90.00, 3821.00.00)
    if code_normalized in ["300249", "30029000", "38210000"]:
        return "Inoculantes, meios de cultura e outros microorganismos para uso agrícola; em conformidade com as definições e demais requisitos da legislação específica"

    # Item 5: Bioestimulantes e bioinsumos (38.24, 3807.00.00, 12.11, 38.08)
    if re.match(r"^3824\d{4}$", code_normalized) and code_normalized not in ["38249977", "38249979", "38249989"] or code_normalized == "38070000" or re.match(r"^1211\d{4}$", code_normalized) or re.match(r"^3808\d{4}$", code_normalized):
        return "Bioestimulantes e bioinsumos para controle fitossanitário, em conformidade com as definições e demais requisitos da legislação específica"

    # Item 6: Defensivos agrícolas (38.08, 3824.99.89)
    if re.match(r"^3808\d{4}$", code_normalized) or code_normalized == "38249989":
        return "Inseticidas, fungicidas, formicidas, herbicidas, parasiticidas, germicidas, acaricidas, nematicidas, raticidas, desfolhantes, dessecantes, espalhantes adesivos, estimuladores e inibidores de crescimento (reguladores); todos destinados diretamente ao uso agropecuário ou destinados diretamente à fabricação de defensivo agropecuário; em conformidade com as definições e demais requisitos da legislação específica"

    # Item 7: Matérias-primas para insumos (vários NCMs)
    if code_normalized in ["0506", "12011000", "12130000", "13019090", "1302199", "14019000", "14049090", "21022000", "2302", "2303", "230400", "23050000", "2306", "23080000", "27030000", "28399010", "28399050", "29224", "293040", "3301", "38029040", "380400", "38249971", "44013900", "44014", "44029000", "47010000", "53050090", "68062000"]:
        return "Calcário, casca de coco triturada, turfa; tortas, bagaços e demais resíduos e desperdícios vegetais das indústrias alimentares; cascas, serragens e demais resíduos e desperdícios de madeira; resíduos da indústria de celulose (dregs e grits), ossos, borra de carnaúba, cinzas, resíduos agroindustriais orgânicos, DL-Metionina e seus análogos, vermiculita e argilas expandidas, palhas e cascas de produtos vegetais, fibra de coco e outras fibras vegetais, silicatos de potássio ou de magnésio, resinas e oleorresinas naturais, sucos e extratos vegetais, aminoácidos e microrganismos mortos, óleos essenciais, argilas e terras, carvão vegetal e pastas mecânicas de madeira; todos destinados diretamente à fabricação de biofertilizantes, fertilizantes, corretivos de solo (inclusive condicionadores), remineralizadores, substratos para plantas, bioestimulantes ou biodefensivos para controle fitossanitário ou utilizados diretamente como biofertilizantes, fertilizantes, corretivos de solo (inclusive condicionadores), remineralizadores, substratos para plantas, bioestimulantes ou biodefensivos para controle fitossanitário; em conformidade com as definições e demais requisitos da legislação específica"

    # Item 8: Ácidos para fertilizantes (vários NCMs)
    if code_normalized in ["25030010", "25030090", "25101010", "25101090", "25102010", "25102090", "28020000", "28061020", "28070010", "28080010", "28092011", "28092019", "28111920", "28151100", "28151200", "28362010", "28362090", "29152100"]:
        return "Ácido nítrico, ácido sulfúrico, ácido fosfórico, fosfatos de cálcio naturais, enxofre, ácido clorídrico, ácido fosforoso, ácido acético, hidróxido de sódio e carbonato dissódico; todos destinados diretamente à fabricação de fertilizantes"

    # Item 9: Enzimas (3507.90.4X)
    if re.match(r"^3507904\d$", code_normalized):
        return "Enzimas preparadas para decomposição de matéria orgânica animal e vegetal"

    # Item 11: Mudas (06.01, 06.02)
    if re.match(r"^060[12]\d{4}$", code_normalized):
        return "Mudas de plantas e demais materiais propagativos de plantas e fungos, inclusive plantas e fungos nativos de espécies florestais; em conformidade com as definições e demais requisitos da legislação específica"

    # Item 12: Vacinas veterinárias (3002.12, 3002.15, 3002.42, 3002.90.00, 30.04)
    if code_normalized in ["300212", "300215", "300242", "30029000"] or re.match(r"^3004\d{4}$", code_normalized):
        return "Vacinas, soros e medicamentos, de uso veterinário, exceto de animais domésticos"

    # Item 13: Aves de um dia (0105.1)
    if re.match(r"^01051\d{3}$", code_normalized):
        return "Aves de um dia, exceto as ornamentais"

    # Item 14: Embriões e sêmen (0511.10.00, 0511.9)
    if code_normalized == "05111000" or re.match(r"^05119\d{3}$", code_normalized):
        return "Embriões e sêmen, congelado ou resfriado"

    # Item 15: Reprodutores (01.02, 01.03, 01.04)
    if re.match(r"^010[234]\d{4}$", code_normalized):
        return "Reprodutores de raça pura, inclusive matrizes de animais puros de origem com registro genealógico; em conformidade com as definições e demais requisitos da legislação específica"

    # Item 16: Ovos fertilizados (0407.1)
    if re.match(r"^04071\d{3}$", code_normalized):
        return "Ovos fertilizados"

    # Item 17: Girinos e alevinos (0106.90.00)
    if code_normalized == "01069000":
        return "Girinos e alevinos"

    # Item 18: Rações (2309.90)
    if re.match(r"^230990\d{2}$", code_normalized):
        return "Rações para animais, concentrados, suplementos, aditivos, premix ou núcleo, exceto para animais domésticos"

    # Item 20: Farelos para ração (23.01 a 23.06, 2308.00.00)
    if re.match(r"^230[1-6]\d{4}$", code_normalized) or code_normalized == "23080000":
        return "Farelos e tortas de produtos vegetais e demais resíduos e desperdícios das indústrias alimentares; todos destinados diretamente à fabricação de ração para animais ou diretamente à alimentação animal, exceto de animais domésticos"

    # Item 21: Matérias-primas para ração (vários NCMs)
    if code_normalized in ["0210", "0309", "07129010", "250100", "25210000", "293040"] or re.match(r"^15\d{6}$", code_normalized):
        return "Alho em pó, sal mineralizado, farinhas de peixe, de ostra, de carne, de osso, de pena, de sangue e de víscera, calcário calcítico, gorduras e óleos animais, resíduos de óleo e de gordura de origem animal ou vegetal descartados por empresas do ramo alimentício, e DL-Metionina e seus análogos; todos destinados diretamente à fabricação de ração para animais ou diretamente à alimentação animal, exceto de animais domésticos"

    # Item 35: Vinhaça (2303.30.00, 2303.20.00)
    if code_normalized in ["23033000", "23032000"]:
        return "Vinhaça"

    # Item 2: Fertilizantes (Capítulo 31, exceto 3101.00.00 já tratado no Item 1)
    if re.match(r"^31\d{6}$", code_normalized) and code_normalized != "31010000" or code_normalized in ["38249977", "38249979", "38249989"]:
        return "Fertilizantes (adubos), em conformidade com as definições e demais requisitos da legislação específica"

    # Item 3: Corretivos de solo (Capítulo 25, exceto NCMs do Item 8 e Item 21)
    if (re.match(r"^25\d{6}$", code_normalized) and
        code_normalized not in ["25030010", "25030090", "25101010", "25101090", "25102010", "25102090", "250100", "25210000"]):
        return "Corretivos de solo (inclusive condicionadores), remineralizadores e substratos para plantas; em conformidade com as definições e demais requisitos da legislação específica"

    # Item 10: Sementes (Capítulos 7, 10, 12, exceto NCMs do Item 5, Item 7 e Item 21)
    if (re.match(r"^(07|10|12)\d{6}$", code_normalized) and
        code_normalized not in ["07129010", "12011000", "12130000"] and
        not re.match(r"^1211\d{4}$", code_normalized)):
        return "Semente genética, semente básica, semente nativa in natura, semente certificada de primeira geração (C1), semente certificada de segunda geração (C2), semente não certificada de primeira geração (S1), semente não certificada de segunda geração (S2) e sementes de cultivar local, tradicional ou crioula; em conformidade com as definições e demais requisitos da legislação específica"

    # Item 19: Sementes/cereais para ração (Capítulos 10, 11, 12, exceto NCMs do Item 5, Item 7, Item 10 e Item 21)
    if (re.match(r"^(10|11|12)\d{6}$", code_normalized) and
        code_normalized not in ["07129010", "12011000", "12130000"] and
        not re.match(r"^1211\d{4}$", code_normalized)):
        if re.match(r"^11\d{6}$", code_normalized):
            return "Sementes e cereais, mesmo triturados, em grãos esmagados ou trabalhados de outro modo; todos destinados diretamente à fabricação de ração para animais ou diretamente à alimentação animal, exceto de animais domésticos"

    # Item 22: Serviços agronômicos (1.1410.90.00)
    if code_normalized == "114109000":
        return "Serviços agronômicos"

    # Item 23: Serviços de técnico agrícola, agropecuário ou em agroecologia (1.1410.90.00)
    if code_normalized == "114109000":
        return "Serviços de técnico agrícola, agropecuário ou em agroecologia"

    # Item 24: Serviços veterinários para produção animal (1.1405.21.00, 1.1405.22.00, 1.1405.90.00)
    if code_normalized in ["114052100", "114052200", "114059000"]:
        return "Serviços veterinários para produção animal"

    # Item 25: Serviços de zootecnistas (1.1410.90.00)
    if code_normalized == "114109000":
        return "Serviços de zootecnistas"

    # Item 26: Serviços de inseminação e fertilização de animais de criação (1.1405.22.00)
    if code_normalized == "114052200":
        return "Serviços de inseminação e fertilização de animais de criação"

    # Item 27: Serviços de engenharia florestal (1.1403.10.00)
    if code_normalized == "114031000":
        return "Serviços de engenharia florestal"

    # Item 28: Serviços de pulverização e controle de pragas (1.1901.10.00)
    if code_normalized == "119011000":
        return "Serviços de pulverização e controle de pragas"

    # Item 29: Serviços de semeadura, adubação, reparação de solo, plantio e colheita (1.1901.10.00)
    if code_normalized == "119011000":
        return "Serviços de semeadura, adubação, inclusive mistura de adubos, reparação de solo, plantio e colheita"

    # Item 30: Serviços de projetos para irrigação e fertirrigação (1.1403.29.00)
    if code_normalized == "114032900":
        return "Serviços de projetos para irrigação e fertirrigação"

    # Item 31: Serviços de análise laboratorial (1.1404.41.00)
    if code_normalized == "114044100":
        return "Serviços de análise laboratorial de solos, sementes e outros materiais propagativos, fitossanitários, água de produção, bromatologia e sanidade animal"

    # Item 32: Licenciamento de direitos sobre cultivares (1.1105.10.00)
    if code_normalized == "111051000":
        return "Licenciamento de direitos sobre cultivares"

    # Item 33: Cessão definitiva de direitos sobre cultivares (1.1109.10.00)
    if code_normalized == "111091000":
        return "Cessão definitiva de direitos sobre cultivares"

    return f"Código {code} (normalizado: {code_normalized}) não encontrado na tabela fornecida."

# Função para verificar se uma descrição indica que o código foi encontrado
def is_code_matched(description):
    return "não encontrado" not in description

# Função para obter o código do item com base no NCM
def get_item_code(ncm_code):
    code_normalized = normalize_code(ncm_code)
    
    # Item 1: Biofertilizantes (3101.00.00)
    if code_normalized == "31010000":
        return "1"
    
    # Item 2: Fertilizantes (Capítulo 31, exceto 3101.00.00)
    if re.match(r"^31\d{6}$", code_normalized) and code_normalized != "31010000" or code_normalized in ["38249977", "38249979", "38249989"]:
        return "2"
    
    # Item 3: Corretivos de solo (Capítulo 25, com exceções)
    if (re.match(r"^25\d{6}$", code_normalized) and 
        code_normalized not in ["25030010", "25030090", "25101010", "25101090", "25102010", "25102090", "250100", "25210000"]):
        return "3"
    
    # Item 4: Inoculantes (3002.49, 3002.90.00, 3821.00.00)
    if code_normalized in ["300249", "30029000", "38210000"]:
        return "4"
    
    # Item 5: Bioestimulantes e bioinsumos
    if re.match(r"^3824\d{4}$", code_normalized) and code_normalized not in ["38249977", "38249979", "38249989"] or code_normalized == "38070000" or re.match(r"^1211\d{4}$", code_normalized) or re.match(r"^3808\d{4}$", code_normalized):
        return "5"
    
    # Item 6: Defensivos agrícolas
    if re.match(r"^3808\d{4}$", code_normalized) or code_normalized == "38249989":
        return "6"
    
    # Item 7: Matérias-primas para insumos
    if code_normalized in ["0506", "12011000", "12130000", "13019090", "1302199", "14019000", "14049090", "21022000", "2302", "2303", "230400", "23050000", "2306", "23080000", "27030000", "28399010", "28399050", "29224", "293040", "3301", "38029040", "380400", "38249971", "44013900", "44014", "44029000", "47010000", "53050090", "68062000"]:
        return "7"
    
    # Item 8: Ácidos para fertilizantes
    if code_normalized in ["25030010", "25030090", "25101010", "25101090", "25102010", "25102090", "28020000", "28061020", "28070010", "28080010", "28092011", "28092019", "28111920", "28151100", "28151200", "28362010", "28362090", "29152100"]:
        return "8"
    
    # Item 9: Enzimas
    if re.match(r"^3507904\d$", code_normalized):
        return "9"
    
    # Item 10: Sementes
    if (re.match(r"^(07|10|12)\d{6}$", code_normalized) and 
        code_normalized not in ["07129010", "12011000", "12130000"] and 
        not re.match(r"^1211\d{4}$", code_normalized)):
        return "10"
    
    # Item 11: Mudas
    if re.match(r"^060[12]\d{4}$", code_normalized):
        return "11"
    
    # Item 12: Vacinas veterinárias
    if code_normalized in ["300212", "300215", "300242", "30029000"] or re.match(r"^3004\d{4}$", code_normalized):
        return "12"
    
    # Item 13: Aves de um dia
    if re.match(r"^01051\d{3}$", code_normalized):
        return "13"
    
    # Item 14: Embriões e sêmen
    if code_normalized == "05111000" or re.match(r"^05119\d{3}$", code_normalized):
        return "14"
    
    # Item 15: Reprodutores
    if re.match(r"^010[234]\d{4}$", code_normalized):
        return "15"
    
    # Item 16: Ovos fertilizados
    if re.match(r"^04071\d{3}$", code_normalized):
        return "16"
    
    # Item 17: Girinos e alevinos
    if code_normalized == "01069000":
        return "17"
    
    # Item 18: Rações
    if re.match(r"^230990\d{2}$", code_normalized):
        return "18"
    
    # Item 19: Sementes/cereais para ração
    if (re.match(r"^(10|11|12)\d{6}$", code_normalized) and 
        code_normalized not in ["07129010", "12011000", "12130000"] and 
        not re.match(r"^1211\d{4}$", code_normalized) and
        re.match(r"^11\d{6}$", code_normalized)):
        return "19"
    
    # Item 20: Farelos para ração
    if re.match(r"^230[1-6]\d{4}$", code_normalized) or code_normalized == "23080000":
        return "20"
    
    # Item 21: Matérias-primas para ração
    if code_normalized in ["0210", "0309", "07129010", "250100", "25210000", "293040"] or re.match(r"^15\d{6}$", code_normalized):
        return "21"
    
    # Item 22: Serviços agronômicos
    if code_normalized == "114109000":
        return "22"
    
    # Item 23: Serviços de técnico agrícola
    if code_normalized == "114109000":
        return "23"
    
    # Item 24: Serviços veterinários
    if code_normalized in ["114052100", "114052200", "114059000"]:
        return "24"
    
    # Item 25: Serviços de zootecnistas
    if code_normalized == "114109000":
        return "25"
    
    # Item 26: Serviços de inseminação
    if code_normalized == "114052200":
        return "26"
    
    # Item 27: Serviços de engenharia florestal
    if code_normalized == "114031000":
        return "27"
    
    # Item 28: Serviços de pulverização
    if code_normalized == "119011000":
        return "28"
    
    # Item 29: Serviços de semeadura
    if code_normalized == "119011000":
        return "29"
    
    # Item 30: Serviços de projetos para irrigação
    if code_normalized == "114032900":
        return "30"
    
    # Item 31: Serviços de análise laboratorial
    if code_normalized == "114044100":
        return "31"
    
    # Item 32: Licenciamento de direitos sobre cultivares
    if code_normalized == "111051000":
        return "32"
    
    # Item 33: Cessão definitiva de direitos sobre cultivares
    if code_normalized == "111091000":
        return "33"
    
    # Item 35: Vinhaça
    if code_normalized in ["23033000", "23032000"]:
        return "35"
    
    return "N/A"

# Rota para servir o frontend
@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    try:
        with open("index.html", "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    except FileNotFoundError:
        return HTMLResponse(content="<h1>Erro: index.html não encontrado</h1>", status_code=500)

# Rota para classificar códigos individuais
@app.post("/classify")
async def classify_codes(input: CodeInput):
    results = []
    for code in input.codes:
        description = classify_code(code)
        item_code = get_item_code(code)
        is_matched = is_code_matched(description)
        results.append({
            "code": code,
            "description": description,
            "item_code": item_code,
            "classification": "Enquadrado" if is_matched else "Não enquadrado"
        })
    return {"results": results}

# Rota para processar a planilha Excel
@app.post("/classify-excel")
async def classify_excel(file: UploadFile = File(...)):
    try:
        start_time = time.time()
        
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Apenas arquivos Excel (.xlsx ou .xls) são aceitos.")
        
        content = await file.read()
        
        # Log para debug
        logger.info(f"Recebido arquivo: {file.filename}, tamanho: {len(content)} bytes")
        
        try:
            # Tente ler o arquivo Excel com diferentes engines
            try:
                df = pd.read_excel(BytesIO(content), engine='openpyxl')
                logger.info("Arquivo lido com sucesso usando engine 'openpyxl'")
            except Exception as e:
                logger.warning(f"Erro ao ler com openpyxl: {str(e)}. Tentando com xlrd...")
                df = pd.read_excel(BytesIO(content), engine='xlrd')
                logger.info("Arquivo lido com sucesso usando engine 'xlrd'")
        except Exception as e:
            logger.error(f"Falha ao ler o arquivo Excel: {str(e)}")
            raise HTTPException(status_code=400, detail=f"Não foi possível ler o arquivo Excel: {str(e)}")
        
        # Log das colunas encontradas
        logger.info(f"Colunas no arquivo: {df.columns.tolist()}")
        
        # Verificar coluna NCM
        if "NCM" not in df.columns:
            # Tentar encontrar uma coluna que possa conter NCMs
            potential_columns = [col for col in df.columns if col.upper() in ["NCM", "CODIGO", "CÓDIGO", "COD", "CÓD"]]
            
            if potential_columns:
                # Usar a primeira coluna potencial encontrada
                df = df.rename(columns={potential_columns[0]: "NCM"})
                logger.info(f"Coluna renomeada: {potential_columns[0]} -> NCM")
            else:
                # Verificar se a primeira coluna contém dados que parecem NCMs
                first_col = df.columns[0]
                sample_values = df[first_col].iloc[:5].astype(str)
                
                # Verificar se os valores da primeira coluna parecem NCMs
                if all(re.match(r'^\d+(\.\d+)*$', val) for val in sample_values if val != 'nan'):
                    df = df.rename(columns={first_col: "NCM"})
                    logger.info(f"Primeira coluna renomeada para NCM: {first_col}")
                else:
                    logger.error("Coluna NCM não encontrada e não foi possível identificar uma coluna adequada")
                    raise HTTPException(status_code=400, detail="A planilha deve conter uma coluna chamada 'NCM' ou similar. Colunas encontradas: " + ", ".join(df.columns.tolist()))
        
        # Garantir que valores da coluna NCM sejam strings
        df['NCM'] = df['NCM'].astype(str)
        
        # Verificar se há coluna de faturamento
        has_revenue_data = False
        revenue_column = None
        
        # Procurar por colunas de faturamento potenciais
        potential_revenue_columns = [col for col in df.columns if any(term in col.upper() for term in ["FATUR", "RECEITA", "VALOR", "REVENUE", "VENDAS", "SALES"])]
        
        if potential_revenue_columns:
            revenue_column = potential_revenue_columns[0]
            has_revenue_data = True
            logger.info(f"Coluna de faturamento encontrada: {revenue_column}")
            
            # Converter valores de faturamento para numérico
            try:
                df[revenue_column] = pd.to_numeric(df[revenue_column], errors='coerce')
                # Substituir NaN por 0
                df[revenue_column].fillna(0, inplace=True)
            except Exception as e:
                logger.warning(f"Erro ao converter coluna de faturamento para numérico: {str(e)}")
                has_revenue_data = False
        
        # Criando uma coluna para "Descrição do Produto" se não existir
        if "Descrição do Produto" not in df.columns:
            # Verificar se existe alguma coluna que pareça conter descrições
            desc_columns = [col for col in df.columns if any(s in col.upper() for s in ["DESCR", "PRODUTO", "ITEM", "MERCAD"])]
            
            if desc_columns:
                # Usar a coluna existente
                df = df.rename(columns={desc_columns[0]: "Descrição do Produto"})
                logger.info(f"Coluna de descrição encontrada e renomeada: {desc_columns[0]} -> 'Descrição do Produto'")
            else:
                # Criar uma coluna vazia para descrição do produto
                df["Descrição do Produto"] = "Não informado"
                logger.info("Coluna 'Descrição do Produto' criada com valor padrão")
        
        # Classificar os NCMs
        df['Classificação NCM'] = df['NCM'].apply(classify_code)
        df['Código do item (índice)'] = df['NCM'].apply(get_item_code)
        df['Classificação (enquadra ou não)'] = df['Classificação NCM'].apply(lambda x: "Enquadrado" if is_code_matched(x) else "Não enquadrado")
        df['Enquadrado'] = df['Classificação (enquadra ou não)'].apply(lambda x: True if x == "Enquadrado" else False)

        # Calcular métricas básicas
        total_ncms = len(df)
        matched_ncms = df['Enquadrado'].sum()
        percentage = round((matched_ncms / total_ncms) * 100, 2) if total_ncms > 0 else 0
        
        # Métricas financeiras (apenas se houver dados de faturamento)
        financial_metrics = {}
        if has_revenue_data:
            # Total de faturamento
            total_revenue = df[revenue_column].sum()
            
            # Faturamento dos NCMs enquadrados
            matched_revenue = df[df['Enquadrado']][revenue_column].sum()
            
            # Percentual do faturamento enquadrado
            revenue_percentage = round((matched_revenue / total_revenue) * 100, 2) if total_revenue > 0 else 0
            
            # Média de faturamento por NCM
            avg_revenue_per_ncm = round(total_revenue / total_ncms, 2) if total_ncms > 0 else 0
            
            # Média de faturamento por NCM enquadrado
            avg_revenue_per_matched_ncm = round(matched_revenue / matched_ncms, 2) if matched_ncms > 0 else 0
            
            # Top 5 NCMs por faturamento
            top_ncms = df.nlargest(5, revenue_column)[['NCM', revenue_column, 'Enquadrado']].to_dict(orient='records')
            
            # Faturamento médio dos NCMs enquadrados vs não enquadrados
            avg_revenue_matched = df[df['Enquadrado']][revenue_column].mean() if matched_ncms > 0 else 0
            avg_revenue_not_matched = df[~df['Enquadrado']][revenue_column].mean() if (total_ncms - matched_ncms) > 0 else 0
            
            # Adicionar porcentagem do total para cada NCM no top 5
            for ncm in top_ncms:
                ncm['percentage_of_total'] = round((ncm[revenue_column] / total_revenue) * 100, 2)
            
            # Montar o dicionário de métricas financeiras
            financial_metrics = {
                "total_revenue": float(total_revenue),
                "matched_revenue": float(matched_revenue),
                "revenue_percentage": float(revenue_percentage),
                "avg_revenue_per_ncm": float(avg_revenue_per_ncm),
                "avg_revenue_per_matched_ncm": float(avg_revenue_per_matched_ncm),
                "top_ncms": top_ncms,
                "avg_revenue_matched": float(avg_revenue_matched),
                "avg_revenue_not_matched": float(avg_revenue_not_matched),
                "potential_tax_impact": float(matched_revenue * 0.08),  # 8% de economia fiscal estimada
                "annual_savings": float(matched_revenue * 0.08)  # Mesma economia, mas para referência anual
            }
            
            # Adicionar coluna de faturamento percentual para todos os NCMs
            df['Percentual do Faturamento'] = df[revenue_column] / total_revenue * 100
        
        # Calcular tempo de processamento
        processing_time = round(time.time() - start_time, 2)
        logger.info(f"Processamento concluído em {processing_time} segundos. Total: {total_ncms}, Classificados: {matched_ncms}")
        
        # Reordenar colunas no formato solicitado
        if has_revenue_data:
            output_columns = [
                'Código do item (índice)', 
                'NCM', 
                'Descrição do Produto', 
                revenue_column, 
                'Percentual do Faturamento',
                'Classificação (enquadra ou não)', 
                'Classificação NCM'
            ]
        else:
            output_columns = [
                'Código do item (índice)', 
                'NCM', 
                'Descrição do Produto', 
                'Classificação (enquadra ou não)', 
                'Classificação NCM'
            ]
        
        # Verificar quais colunas existem antes de reorganizar
        available_columns = [col for col in output_columns if col in df.columns]
        other_columns = [col for col in df.columns if col not in output_columns and col != 'Enquadrado']
        final_columns = available_columns + other_columns
        
        # Gerar arquivo de saída
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            output_path = tmp.name
            
            # Criar um escritor do Excel com formatação
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df_export = df[final_columns].copy()
                
                # Formatar valores percentuais
                if has_revenue_data and 'Percentual do Faturamento' in df_export.columns:
                    df_export['Percentual do Faturamento'] = df_export['Percentual do Faturamento'].round(2).astype(str) + '%'
                
                df_export.to_excel(writer, index=False, sheet_name='Classificação NCM')
                
                # Adicionar planilha de resumo
                summary_data = {
                    'Métrica': [
                        'Total de NCMs', 
                        'NCMs enquadrados na redução', 
                        'Percentual de Cobertura',
                        'Tempo de Processamento'
                    ],
                    'Valor': [
                        total_ncms,
                        matched_ncms,
                        f"{percentage}%",
                        f"{processing_time}s"
                    ]
                }
                
                if has_revenue_data:
                    summary_data['Métrica'].extend([
                        'Faturamento Total',
                        'Faturamento Enquadrado',
                        'Percentual do Faturamento Enquadrado',
                        'Faturamento Médio por NCM',
                        'Faturamento Médio por NCM Enquadrado',
                        'Faturamento Médio por NCM Não Enquadrado',
                        'Economia Fiscal Estimada (8%)'
                    ])
                    
                    summary_data['Valor'].extend([
                        f"R$ {total_revenue:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'),
                        f"R$ {matched_revenue:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'),
                        f"{revenue_percentage}%",
                        f"R$ {avg_revenue_per_ncm:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'),
                        f"R$ {avg_revenue_matched:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'),
                        f"R$ {avg_revenue_not_matched:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'),
                        f"R$ {matched_revenue * 0.08:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
                    ])
                
                # Criar DataFrame para o resumo
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, index=False, sheet_name='Resumo')
            
            logger.info(f"Arquivo de saída criado: {output_path}")
        
        # Preparar métricas para resposta
        metrics = {
            "total_ncms": int(total_ncms),
            "matched_ncms": int(matched_ncms),
            "percentage": float(percentage),
            "processing_time": float(processing_time),
            "has_revenue_data": has_revenue_data
        }
        
        # Adicionar métricas financeiras se disponíveis
        if has_revenue_data:
            metrics.update(financial_metrics)
        
        # Preparar resposta simplificada para a interface
        display_columns = ['NCM', 'Classificação NCM']
        if has_revenue_data:
            display_columns.append(revenue_column)
        
        display_results = df[display_columns].rename(columns={'Classificação NCM': 'Descrição'})
        
        # Preparar resposta com resultados e métricas
        response_data = {
            "results": display_results.to_dict(orient='records'), 
            "output_file": os.path.basename(output_path),
            "metrics": metrics,
            "planilha_tipo": 2 if has_revenue_data else 1
        }
        
        logger.info("Enviando resposta com métricas: " + str(metrics))
        return response_data
        
    except Exception as e:
        logger.error(f"Erro ao processar a planilha: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Erro ao processar a planilha: {str(e)}")

# Rota para download da planilha processada
@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(tempfile.gettempdir(), filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename="classificacao_ncm.xlsx")
    raise HTTPException(status_code=404, detail="Arquivo não encontrado.")

# Iniciar o servidor com o comando: uvicorn main:app --reload
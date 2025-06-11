# ProjetoNCM
O Classificador NCM/NBS é uma aplicação web para classificação automatizada de códigos NCM (Nomenclatura Comum do Mercosul) e NBS (Nomenclatura Brasileira de Serviços), com suporte a análise de faturamento e geração de relatórios em Excel.

## Descrição
O **ProjetoNCM** é uma solução desenvolvida para facilitar a classificação de códigos NCM e NBS conforme categorias predefinidas, atendendo às necessidades de empresas e consultores que buscam conformidade fiscal e benefícios tributários. A aplicação permite classificar códigos individualmente ou em lote via upload de planilhas Excel, oferecendo análises detalhadas, incluindo métricas financeiras quando dados de faturamento são fornecidos.

## Funcionalidades
- **Classificação Individual**: Insira códigos NCM/NBS manualmente para obter descrição, código do item e status de enquadramento.
- **Classificação em Lote**: Procesamento de planilhas Excel nos formatos:
  - **Tipo 1**: Apenas códigos NCM/NBS.
  - **Tipo 2**: Códigos NCM/NBS com dados de faturamento.
- **Análise Financeira** (para planilhas Tipo 2):
  - Faturamento total e enquadrado.
  - Percentual de cobertura e economia fiscal estimada (8%).
  - Top 5 NCMs por faturamento e médias por NCM.
- **Relatórios**: Gera planilhas Excel com resultados organizados e uma aba de resumo com métricas.
- **Interface Web**: Frontend intuitivo em HTML para upload de planilhas, teste 1 a 1 e visualização de resultados.

## Tecnologias
- **Backend**: FastAPI (Python)
- **Frontend**: HTML com JavaScript
- **Bibliotecas Python**: `pandas`, `openpyxl`, `xlrd`, `numpy`, `uvicorn`, `gunicorn`

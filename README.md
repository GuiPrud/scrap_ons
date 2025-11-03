# 📊 Web Scraping ONS - Power BI Dashboard

Scripts para extração automatizada de dados dos dashboards Power BI da ONS (Operador Nacional do Sistema Elétrico).

## 🎯 Objetivo

Este repositório contém ferramentas para fazer web scraping dos gráficos e tabelas interativas disponibilizados pela ONS em seus dashboards Power BI, permitindo análise e processamento dos dados de forma automatizada.

## 🚀 Scripts Disponíveis

### `scrape_ons_powerbi_direct.py`
Script principal para extração completa de múltiplas páginas do dashboard Power BI.

**Funcionalidades:**
- ✅ Navegação automática entre páginas
- ✅ Extração de tabelas, cards/KPIs e gráficos
- ✅ Exportação em CSV, Excel e Pickle
- ✅ Screenshots de cada página
- ✅ Organização automática em pastas com timestamp

### `scrape_powerbi.py`
Script alternativo com foco em captura de requisições de rede e dados visuais.

**Funcionalidades:**
- ✅ Interceptação de requisições HTTP
- ✅ Extração de elementos visuais
- ✅ Análise de estrutura do dashboard

### `scrape_ons.py`
Script para extração via página ONS com detecção automática de iframes Power BI.

## 📦 Instalação

```bash
pip install selenium pandas beautifulsoup4 lxml openpyxl
```

## 💻 Uso

```bash
# Extração completa com múltiplas páginas
python scrape_ons_powerbi_direct.py

# Extração alternativa
python scrape_powerbi.py
```

## 📁 Estrutura de Saída

Os dados são salvos automaticamente em pastas organizadas por timestamp:

```
extracao_powerbi_20251103_143052/
├── ons_powerbi_ALL_TABLES_CONSOLIDATED.csv    # Todas as tabelas consolidadas
├── ons_powerbi_dataframe.pkl                  # DataFrame Pandas
├── ons_powerbi_dataframe.xlsx                 # Arquivo Excel
├── ons_powerbi_data_complete.json             # Dados completos JSON
├── ons_powerbi_ALL_cards_kpis.txt            # Cards e KPIs
├── powerbi_screenshot_page1.png               # Screenshots
└── ...
```

## 📊 Dados Extraídos

- **Tabelas:** Dados tabulares em CSV/Excel
- **Cards/KPIs:** Métricas individuais
- **Gráficos:** Labels, valores e dados SVG
- **Screenshots:** Capturas visuais de cada página

## 🔧 Requisitos

- Python 3.7+
- Google Chrome ou Firefox
- ChromeDriver/GeckoDriver
- Conexão com internet

## 📖 Documentação Adicional

- `README_POWERBI_MULTIPLAS_PAGINAS.md` - Guia detalhado de extração
- `FORMATOS_DATAFRAME.md` - Formatos de exportação
- `README_PASTAS_ORGANIZADAS.md` - Organização de arquivos

## ⚠️ Avisos

- Respeite os termos de uso da ONS
- Use os dados de forma responsável
- Verifique a frequência de requisições

## 📝 Licença

Este projeto é disponibilizado para fins educacionais e de pesquisa.

---

**Desenvolvido para análise de dados do setor elétrico brasileiro** ⚡


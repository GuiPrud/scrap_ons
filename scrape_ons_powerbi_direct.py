"""
Script otimizado para extrair dados diretamente do Power BI Dashboard da ONS
Foca especificamente no iframe Power BI identificado
"""

import time
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
import sys
from datetime import datetime

# URL da página ONS
PAGE_URL = "https://www.ons.org.br/Paginas/faq_curtailment.aspx"

# URL direta do Power BI (extraída do iframe)
POWERBI_DIRECT_URL = "https://app.powerbi.com/view?r=eyJrIjoiYmU0ODUxNGMtNWU2MS00YTM5LThkMGYtNWFkYWQzYmU3ZWY2IiwidCI6IjNhZGVlNWZjLTkzM2UtNDkxMS1hZTFiLTljMmZlN2I4NDQ0OCIsImMiOjR9"


def create_output_folder():
    """Cria pasta para salvar os arquivos gerados"""
    folder_name = f"extracao_powerbi"
    
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        print(f"📁 Pasta criada: {folder_name}")
    
    return folder_name


def setup_driver():
    """Configura Chrome com opções otimizadas para Power BI"""
    options = Options()
    
    # Opções para melhor desempenho
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--start-maximized')
    
    # Ativa logs de performance (útil para debugar Power BI)
    options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
    
    print("Inicializando Chrome driver...")
    try:
        driver = webdriver.Chrome(options=options)
        return driver
    except Exception as e:
        print(f"❌ Erro ao inicializar Chrome: {e}")
        print("\nTente instalar webdriver-manager:")
        print("  pip install webdriver-manager")
        return None


def wait_for_powerbi_load(driver, timeout=60):
    """
    Aguarda Power BI carregar completamente
    Power BI usa renderização assíncrona complexa
    """
    print(f"\n⏳ Aguardando Power BI carregar (timeout: {timeout}s)...")
    
    start_time = time.time()
    
    # Estratégias progressivas
    strategies = [
        ("Body presente", By.TAG_NAME, "body", 5),
        ("Elemento embed", By.CSS_SELECTOR, "[class*='embed'], [class*='iframe']", 10),
        ("Containers visuais", By.CSS_SELECTOR, "[class*='visual'], [class*='Visual']", 15),
        ("Elementos SVG (gráficos)", By.TAG_NAME, "svg", 10),
        ("Elementos de dados", By.CSS_SELECTOR, "[class*='label'], [class*='value']", 10),
    ]
    
    for description, by_type, selector, wait_time in strategies:
        try:
            elapsed = time.time() - start_time
            remaining = max(1, timeout - elapsed)
            
            print(f"  • {description}...", end=" ")
            WebDriverWait(driver, min(wait_time, remaining)).until(
                EC.presence_of_element_located((by_type, selector))
            )
            print("✓")
            
        except TimeoutException:
            print("✗")
            
        except Exception as e:
            print(f"⚠️  {e}")
    
    # Aguarda adicional para JavaScript finalizar
    print("  • Aguardando JavaScript finalizar...", end=" ")
    time.sleep(10)
    print("✓")
    
    total_time = time.time() - start_time
    print(f"\n✓ Carregamento concluído em {total_time:.1f}s")


def navigate_powerbi_pages(driver, max_pages=10):
    """
    Navega pelas páginas do Power BI clicando no botão 'Próxima Página'
    Retorna o número de páginas navegadas
    """
    print("\n📄 Navegando pelas páginas do Power BI...")
    
    page_count = 1
    print(f"  • Página {page_count} (atual)")
    
    for page_num in range(2, max_pages + 1):
        try:
            # Procura pelo botão de próxima página
            next_button_selectors = [
                # Seletor específico fornecido
                "//button[@aria-label='Próxima Página']",
                "//button[contains(@aria-label, 'Próxima')]",
                # Alternativas comuns
                "//button[contains(@class, 'glyphicon')]//i[contains(@class, 'chevronrightmedium')]/..",
                "//i[contains(@class, 'chevronrightmedium')]/..",
                "//button[contains(@class, 'navigation')]//i[contains(@class, 'chevron')]/..",
                "//button[@aria-label='Next Page']",
            ]
            
            next_button = None
            for selector in next_button_selectors:
                try:
                    buttons = driver.find_elements(By.XPATH, selector)
                    # Procura por botão ativo (não desabilitado)
                    for btn in buttons:
                        aria_disabled = btn.get_attribute('aria-disabled')
                        if aria_disabled != 'true':
                            next_button = btn
                            break
                    if next_button:
                        break
                except:
                    continue
            
            if not next_button:
                print(f"  ✓ Última página alcançada (página {page_count})")
                break
            
            # Verifica se o botão está desabilitado
            if next_button.get_attribute('aria-disabled') == 'true':
                print(f"  ✓ Botão desabilitado - última página (página {page_count})")
                break
            
            # Scroll até o botão
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
            time.sleep(0.5)
            
            # Tenta clicar
            try:
                next_button.click()
            except:
                # Fallback: JavaScript click
                driver.execute_script("arguments[0].click();", next_button)
            
            print(f"  • Navegando para página {page_num}...", end=" ")
            
            # Aguarda a página carregar
            time.sleep(3)
            
            # Aguarda novos elementos carregarem
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            print("✓")
            page_count = page_num
            
        except TimeoutException:
            print(f"\n  ⚠️  Timeout ao navegar para página {page_num}")
            break
        except Exception as e:
            print(f"\n  ⚠️  Erro ao navegar para página {page_num}: {e}")
            break
    
    print(f"\n✓ Total de páginas navegadas: {page_count}")
    return page_count


def extract_powerbi_visuals(driver):
    """
    Extrai dados dos visuais do Power BI usando JavaScript
    """
    print("\n📊 Extraindo visuais do Power BI...")
    
    # Script JavaScript para extrair dados estruturados
    js_extraction = """
    function extractPowerBIData() {
        let results = {
            tables: [],
            cards: [],
            charts: [],
            raw_text: []
        };
        
        // 1. TABELAS
        console.log('Procurando tabelas...');
        const tables = document.querySelectorAll('table');
        tables.forEach((table, idx) => {
            let tableData = {
                index: idx,
                headers: [],
                rows: []
            };
            
            // Headers
            table.querySelectorAll('th').forEach(th => {
                const text = th.innerText || th.textContent || '';
                tableData.headers.push(text.trim());
            });
            
            // Rows
            table.querySelectorAll('tr').forEach(tr => {
                let cells = [];
                tr.querySelectorAll('td').forEach(td => {
                    const text = td.innerText || td.textContent || '';
                    cells.push(text.trim());
                });
                if (cells.length > 0) {
                    tableData.rows.push(cells);
                }
            });
            
            if (tableData.headers.length > 0 || tableData.rows.length > 0) {
                results.tables.push(tableData);
            }
        });
        
        // 2. CARDS (KPIs)
        console.log('Procurando cards/KPIs...');
        const cardSelectors = [
            '[class*="card"]',
            '[class*="Card"]',
            '[class*="kpi"]',
            '[class*="KPI"]',
            '[role="figure"]'
        ];
        
        cardSelectors.forEach(selector => {
            document.querySelectorAll(selector).forEach((card, idx) => {
                try {
                    const text = (card.innerText || card.textContent || '').trim();
                    if (text && text.length > 0 && text.length < 500) {
                        results.cards.push({
                            selector: selector,
                            index: idx,
                            text: text
                        });
                    }
                } catch (e) {
                    console.error('Erro ao processar card:', e);
                }
            });
        });
        
        // 3. VISUAL CONTAINERS (estrutura Power BI)
        console.log('Procurando visual containers...');
        const visualContainers = document.querySelectorAll(
            '[class*="visual"], [class*="Visual"]'
        );
        
        visualContainers.forEach((container, idx) => {
            try {
                // Procura por labels e valores
                let labels = [];
                let values = [];
                
                container.querySelectorAll('[class*="label"], [class*="axisLabel"]').forEach(label => {
                    const text = (label.textContent || label.innerText || '').trim();
                    if (text) labels.push(text);
                });
                
                container.querySelectorAll('[class*="value"], [class*="data"]').forEach(val => {
                    const text = (val.textContent || val.innerText || '').trim();
                    if (text) values.push(text);
                });
                
                if (labels.length > 0 || values.length > 0) {
                    const combinedText = (container.innerText || container.textContent || '').trim();
                    results.charts.push({
                        index: idx,
                        labels: labels,
                        values: values,
                        combined: combinedText.substring(0, 1000)
                    });
                }
            } catch (e) {
                console.error('Erro ao processar visual container:', e);
            }
        });
        
        // 4. SVG (gráficos vetoriais)
        console.log('Procurando elementos SVG...');
        const svgs = document.querySelectorAll('svg');
        svgs.forEach((svg, idx) => {
            try {
                let textElements = [];
                svg.querySelectorAll('text').forEach(text => {
                    const content = (text.textContent || text.innerText || '').trim();
                    if (content) textElements.push(content);
                });
                
                if (textElements.length > 0) {
                    results.charts.push({
                        index: `svg_${idx}`,
                        type: 'svg',
                        texts: textElements
                    });
                }
            } catch (e) {
                console.error('Erro ao processar SVG:', e);
            }
        });
        
        // 5. TEXTO RAW (fallback)
        try {
            const allText = document.body.innerText || document.body.textContent || '';
            results.raw_text = allText.split('\\n')
                .map(line => line.trim())
                .filter(line => line.length > 0);
        } catch (e) {
            console.error('Erro ao processar texto raw:', e);
            results.raw_text = [];
        }
        
        return results;
    }
    
    return extractPowerBIData();
    """
    
    try:
        data = driver.execute_script(js_extraction)
        
        print(f"✓ Encontrado:")
        print(f"  • {len(data.get('tables', []))} tabela(s)")
        print(f"  • {len(data.get('cards', []))} card(s)/KPI(s)")
        print(f"  • {len(data.get('charts', []))} gráfico(s)")
        print(f"  • {len(data.get('raw_text', []))} linha(s) de texto")
        
        return data
        
    except Exception as e:
        print(f"❌ Erro ao executar JavaScript: {e}")
        return None


def get_user_page_selection():
    """
    Solicita ao usuário qual(is) página(s) extrair
    Retorna: (modo, paginas_especificas)
        modo: 'all', 'specific', 'range'
        paginas_especificas: lista de números de páginas ou None
    """
    print("\n" + "="*70)
    print("  SELEÇÃO DE PÁGINAS PARA EXTRAÇÃO")
    print("="*70)
    print("\nEscolha qual(is) página(s) você deseja extrair:")
    print("  1. Extrair TODAS as páginas")
    print("  2. Extrair página(s) específica(s)")
    print("  3. Extrair um intervalo de páginas")
    print("="*70)
    
    while True:
        try:
            choice = input("\nDigite sua escolha (1, 2 ou 3): ").strip()
            
            if choice == '1':
                print("✓ Modo selecionado: TODAS as páginas")
                return ('all', None)
            
            elif choice == '2':
                pages_input = input("\nDigite o(s) número(s) da(s) página(s) separados por vírgula (ex: 1,3,5): ").strip()
                pages = [int(p.strip()) for p in pages_input.split(',') if p.strip().isdigit()]
                
                if not pages:
                    print("❌ Nenhuma página válida informada. Tente novamente.")
                    continue
                
                pages = sorted(list(set(pages)))  # Remove duplicatas e ordena
                print(f"✓ Páginas selecionadas: {', '.join(map(str, pages))}")
                return ('specific', pages)
            
            elif choice == '3':
                start = input("\nDigite a página inicial: ").strip()
                end = input("Digite a página final: ").strip()
                
                if not (start.isdigit() and end.isdigit()):
                    print("❌ Valores inválidos. Tente novamente.")
                    continue
                
                start_page = int(start)
                end_page = int(end)
                
                if start_page < 1 or end_page < start_page:
                    print("❌ Intervalo inválido. Tente novamente.")
                    continue
                
                pages = list(range(start_page, end_page + 1))
                print(f"✓ Intervalo selecionado: páginas {start_page} a {end_page}")
                return ('range', pages)
            
            else:
                print("❌ Opção inválida. Digite 1, 2 ou 3.")
        
        except ValueError:
            print("❌ Entrada inválida. Tente novamente.")
        except KeyboardInterrupt:
            print("\n\n❌ Operação cancelada pelo usuário.")
            return (None, None)


def extract_all_pages_data(driver, max_pages=10, mode='all', target_pages=None):
    """
    Extrai dados de todas as páginas do Power BI ou páginas específicas
    
    Args:
        driver: Selenium WebDriver
        max_pages: Número máximo de páginas a navegar (para modo 'all')
        mode: 'all' (todas), 'specific' (específicas), 'range' (intervalo)
        target_pages: Lista de páginas a extrair (para modes 'specific' e 'range')
    """
    print("\n" + "="*70)
    if mode == 'all':
        print("  EXTRAÇÃO DE TODAS AS PÁGINAS")
    elif mode == 'specific' and target_pages:
        print(f"  EXTRAÇÃO DE PÁGINAS ESPECÍFICAS: {', '.join(map(str, target_pages))}")
    elif mode == 'range' and target_pages:
        print(f"  EXTRAÇÃO DE INTERVALO: páginas {min(target_pages)} a {max(target_pages)}")
    print("="*70)
    
    all_data = {
        'pages': [],
        'total_tables': 0,
        'total_cards': 0,
        'total_charts': 0,
        'mode': mode,
        'target_pages': target_pages
    }
    
    page_count = 1
    
    while page_count <= max_pages:
        print(f"\n{'='*70}")
        print(f"  PÁGINA {page_count}")
        print(f"{'='*70}")
        
        # Verifica se deve extrair esta página
        should_extract = False
        
        if mode == 'all':
            should_extract = True
        elif mode in ['specific', 'range'] and target_pages:
            should_extract = page_count in target_pages
        
        if should_extract:
            print("  ✓ Extraindo dados desta página...")
            # Extrai dados da página atual
            page_data = extract_powerbi_visuals(driver)
            
            if page_data:
                page_data['page_number'] = page_count
                all_data['pages'].append(page_data)
                
                all_data['total_tables'] += len(page_data.get('tables', []))
                all_data['total_cards'] += len(page_data.get('cards', []))
                all_data['total_charts'] += len(page_data.get('charts', []))
        else:
            print("  ⊘ Pulando esta página (não selecionada)")
        
        # Verifica se deve continuar navegando
        should_continue = False
        
        if mode == 'all' and page_count < max_pages:
            should_continue = True
        elif mode in ['specific', 'range'] and target_pages:
            # Continua se ainda há páginas a extrair
            remaining_pages = [p for p in target_pages if p > page_count]
            if remaining_pages:
                should_continue = True
                next_target = min(remaining_pages)
                print(f"\n  ℹ️  Próxima página alvo: {next_target}")
        
        # Tenta ir para próxima página
        if should_continue:
            print(f"\n➡️  Tentando navegar para página {page_count + 1}...")
            
            try:
                # Procura pelo botão de próxima página
                next_button_selectors = [
                    "//button[@aria-label='Próxima Página' and @aria-disabled='false']",
                    "//button[contains(@aria-label, 'Próxima') and @aria-disabled='false']",
                    "//button[@aria-label='Next Page' and @aria-disabled='false']",
                ]
                
                next_button = None
                for selector in next_button_selectors:
                    try:
                        next_button = WebDriverWait(driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        if next_button:
                            break
                    except:
                        continue
                
                if not next_button:
                    print("  ✓ Última página alcançada (botão não encontrado)")
                    break
                
                # Scroll até o botão
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                time.sleep(0.5)
                
                # Clica no botão
                try:
                    next_button.click()
                except:
                    driver.execute_script("arguments[0].click();", next_button)
                
                print(f"  ✓ Navegado para página {page_count + 1}")
                
                # Aguarda a nova página carregar
                print("  ⏳ Aguardando nova página carregar...")
                time.sleep(5)  # Aguarda 5 segundos para conteúdo carregar
                
                # Aguarda elementos visuais carregarem
                try:
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "[class*='visual'], svg, table"))
                    )
                except:
                    print("  ⚠️  Timeout aguardando elementos visuais")
                
                page_count += 1
                
            except Exception as e:
                print(f"  ✗ Não foi possível navegar: {e}")
                print("  ✓ Finalizando extração")
                break
        else:
            print(f"\n  ⚠️  Limite de {max_pages} páginas alcançado")
            break
    
    print(f"\n{'='*70}")
    print(f"  RESUMO DA EXTRAÇÃO")
    print(f"{'='*70}")
    
    if mode == 'all':
        print(f"  • Modo: TODAS as páginas")
    elif mode == 'specific' and target_pages:
        print(f"  • Modo: Páginas específicas ({', '.join(map(str, target_pages))})")
    elif mode == 'range' and target_pages:
        print(f"  • Modo: Intervalo (páginas {min(target_pages)} a {max(target_pages)})")
    
    print(f"  • Total de páginas extraídas: {len(all_data['pages'])}")
    if all_data['pages']:
        extracted_pages = [p['page_number'] for p in all_data['pages']]
        print(f"  • Páginas extraídas: {', '.join(map(str, extracted_pages))}")
    print(f"  • Total de tabelas: {all_data['total_tables']}")
    print(f"  • Total de cards/KPIs: {all_data['total_cards']}")
    print(f"  • Total de gráficos: {all_data['total_charts']}")
    
    return all_data


def save_data(data, prefix="powerbi", output_folder="."):
    """Salva os dados extraídos em diferentes formatos - suporta múltiplas páginas"""
    print(f"\n💾 Salvando dados...")
    
    saved_files = []
    main_dataframe = None  # DataFrame principal consolidado
    
    # 1. JSON completo
    json_file = os.path.join(output_folder, f"{prefix}_data_complete.json")
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"✓ {json_file}")
    saved_files.append(json_file)
    
    # Se for estrutura de múltiplas páginas
    if 'pages' in data:
        print(f"\n📚 Processando {len(data['pages'])} página(s)...")
        
        all_tables = []
        all_cards = []
        all_charts = []
        
        for page in data['pages']:
            page_num = page.get('page_number', 'unknown')
            
            # 2. Tabelas de cada página
            if page.get('tables'):
                for i, table in enumerate(page['tables']):
                    try:
                        if table['headers']:
                            df = pd.DataFrame(table['rows'], columns=table['headers'])
                        else:
                            df = pd.DataFrame(table['rows'])
                        
                        # Adiciona coluna de página
                        df.insert(0, 'Página', page_num)
                        
                        csv_file = os.path.join(output_folder, f"{prefix}_page{page_num}_table_{i+1}.csv")
                        df.to_csv(csv_file, index=False, encoding='utf-8-sig')
                        print(f"✓ {csv_file} - {df.shape[0]} linhas × {df.shape[1]} colunas")
                        saved_files.append(csv_file)
                        
                        all_tables.append(df)
                        
                    except Exception as e:
                        print(f"⚠️  Erro ao salvar tabela {i+1} da página {page_num}: {e}")
            
            # Coleta cards e gráficos
            if page.get('cards'):
                all_cards.extend(page['cards'])
            if page.get('charts'):
                all_charts.extend(page['charts'])
        
        # 3. Consolida todas as tabelas em um único CSV
        if all_tables:
            try:
                consolidated_df = pd.concat(all_tables, ignore_index=True)
                main_dataframe = consolidated_df  # Salva para exportação pickle
                
                consolidated_file = os.path.join(output_folder, f"{prefix}_ALL_TABLES_CONSOLIDATED.csv")
                consolidated_df.to_csv(consolidated_file, index=False, encoding='utf-8-sig')
                print(f"\n✓ {consolidated_file} - {consolidated_df.shape[0]} linhas × {consolidated_df.shape[1]} colunas")
                print("  (Todas as tabelas de todas as páginas)")
                saved_files.append(consolidated_file)
                
                # Mostra preview consolidado
                print("\n" + "="*60)
                print("Preview dos Dados Consolidados:")
                print("="*60)
                print(consolidated_df.head(15).to_string())
                print("="*60 + "\n")
                
            except Exception as e:
                print(f"⚠️  Erro ao consolidar tabelas: {e}")
        
        # 4. Cards/KPIs consolidados
        if all_cards:
            cards_file = os.path.join(output_folder, f"{prefix}_ALL_cards_kpis.txt")
            with open(cards_file, 'w', encoding='utf-8') as f:
                f.write("=== CARDS E KPIs EXTRAÍDOS (TODAS AS PÁGINAS) ===\n\n")
                for i, card in enumerate(all_cards, 1):
                    f.write(f"Card {i}:\n")
                    f.write(f"{card['text']}\n")
                    f.write("-" * 40 + "\n")
            print(f"✓ {cards_file}")
            saved_files.append(cards_file)
        
        # 5. Gráficos consolidados
        if all_charts:
            charts_file = os.path.join(output_folder, f"{prefix}_ALL_charts_data.json")
            with open(charts_file, 'w', encoding='utf-8') as f:
                json.dump(all_charts, f, indent=2, ensure_ascii=False)
            print(f"✓ {charts_file}")
            saved_files.append(charts_file)
        
        # 6. DataFrame consolidado em Pickle (para uso em Python/Pandas)
        if main_dataframe is not None:
            try:
                pickle_file = os.path.join(output_folder, f"{prefix}_dataframe.pkl")
                main_dataframe.to_pickle(pickle_file)
                print(f"\n✓ {pickle_file} - DataFrame salvo em formato Pickle")
                print(f"  💡 Para carregar: df = pd.read_pickle('{pickle_file}')")
                saved_files.append(pickle_file)
            except Exception as e:
                print(f"⚠️  Erro ao salvar DataFrame pickle: {e}")
        
        # 7. DataFrame consolidado em Excel (se não for muito grande)
        if main_dataframe is not None:
            try:
                if len(main_dataframe) <= 1000000:  # Limite do Excel
                    excel_file = os.path.join(output_folder, f"{prefix}_dataframe.xlsx")
                    main_dataframe.to_excel(excel_file, index=False, engine='openpyxl')
                    print(f"✓ {excel_file} - DataFrame salvo em formato Excel")
                    saved_files.append(excel_file)
                else:
                    print(f"⚠️  DataFrame muito grande ({len(main_dataframe)} linhas) para Excel - use CSV ou Pickle")
            except Exception as e:
                print(f"⚠️  Erro ao salvar DataFrame Excel: {e}")
                print(f"   (Instale openpyxl: pip install openpyxl)")
    
    else:
        # Estrutura de página única (compatibilidade)
        # 2. Tabelas em CSV
        if data.get('tables'):
            all_single_tables = []
            for i, table in enumerate(data['tables']):
                try:
                    if table['headers']:
                        df = pd.DataFrame(table['rows'], columns=table['headers'])
                    else:
                        df = pd.DataFrame(table['rows'])
                    
                    csv_file = os.path.join(output_folder, f"{prefix}_table_{i+1}.csv")
                    df.to_csv(csv_file, index=False, encoding='utf-8-sig')
                    print(f"✓ {csv_file} - {df.shape[0]} linhas × {df.shape[1]} colunas")
                    saved_files.append(csv_file)
                    
                    all_single_tables.append(df)
                    
                    # Mostra preview
                    print("\n" + "="*60)
                    print(f"Preview da Tabela {i+1}:")
                    print("="*60)
                    print(df.head(10).to_string())
                    print("="*60 + "\n")
                    
                except Exception as e:
                    print(f"⚠️  Erro ao salvar tabela {i+1}: {e}")
            
            # Salva DataFrame consolidado (mesmo para página única)
            if all_single_tables:
                try:
                    main_dataframe = pd.concat(all_single_tables, ignore_index=True)
                    
                    # Pickle
                    pickle_file = os.path.join(output_folder, f"{prefix}_dataframe.pkl")
                    main_dataframe.to_pickle(pickle_file)
                    print(f"\n✓ {pickle_file} - DataFrame salvo em formato Pickle")
                    print(f"  💡 Para carregar: df = pd.read_pickle('{pickle_file}')")
                    saved_files.append(pickle_file)
                    
                    # Excel
                    try:
                        excel_file = os.path.join(output_folder, f"{prefix}_dataframe.xlsx")
                        main_dataframe.to_excel(excel_file, index=False, engine='openpyxl')
                        print(f"✓ {excel_file} - DataFrame salvo em formato Excel")
                        saved_files.append(excel_file)
                    except:
                        pass
                        
                except Exception as e:
                    print(f"⚠️  Erro ao salvar DataFrame: {e}")
        
        # 3. Cards/KPIs em arquivo texto
        if data.get('cards'):
            cards_file = os.path.join(output_folder, f"{prefix}_cards_kpis.txt")
            with open(cards_file, 'w', encoding='utf-8') as f:
                f.write("=== CARDS E KPIs EXTRAÍDOS ===\n\n")
                for i, card in enumerate(data['cards'], 1):
                    f.write(f"Card {i}:\n")
                    f.write(f"{card['text']}\n")
                    f.write("-" * 40 + "\n")
            print(f"✓ {cards_file}")
            saved_files.append(cards_file)
        
        # 4. Dados de gráficos
        if data.get('charts'):
            charts_file = os.path.join(output_folder, f"{prefix}_charts_data.json")
            with open(charts_file, 'w', encoding='utf-8') as f:
                json.dump(data['charts'], f, indent=2, ensure_ascii=False)
            print(f"✓ {charts_file}")
            saved_files.append(charts_file)
        
        # 5. Texto completo
        if data.get('raw_text'):
            text_file = os.path.join(output_folder, f"{prefix}_full_text.txt")
            with open(text_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(data['raw_text']))
            print(f"✓ {text_file}")
            saved_files.append(text_file)
    
    return saved_files


def main():
    """Função principal"""
    print("="*70)
    print("  EXTRATOR DE DADOS - POWER BI ONS (MÚLTIPLAS PÁGINAS)")
    print("="*70)
    
    # Cria pasta para salvar os arquivos
    output_folder = create_output_folder()
    
    # Setup
    driver = setup_driver()
    if not driver:
        return
    
    try:
        # Acessa diretamente o Power BI
        print(f"\n🌐 Acessando Power BI diretamente...")
        print(f"URL: {POWERBI_DIRECT_URL[:60]}...")
        
        driver.get(POWERBI_DIRECT_URL)
        
        # Aguarda carregar
        wait_for_powerbi_load(driver, timeout=60)
        
        # Salva screenshot da primeira página
        screenshot_file = os.path.join(output_folder, "powerbi_screenshot_page1.png")
        driver.save_screenshot(screenshot_file)
        print(f"\n📸 Screenshot salvo: {screenshot_file}")
        
        # Salva HTML completo
        html_file = os.path.join(output_folder, "powerbi_page_source.html")
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        print(f"📄 HTML salvo: {html_file}")
        
        # Solicita seleção de páginas ao usuário
        mode, target_pages = get_user_page_selection()
        
        if mode is None:
            print("\n❌ Extração cancelada.")
            return
        
        # Extrai dados das páginas selecionadas
        print("\n" + "="*70)
        print("  INICIANDO EXTRAÇÃO")
        print("="*70)
        print("  ℹ️  O script irá:")
        print("  1. Extrair dados da(s) página(s) selecionada(s)")
        print("  2. Navegar entre páginas conforme necessário")
        print("  3. Salvar os dados extraídos")
        print("="*70)
        
        data = extract_all_pages_data(driver, max_pages=20, mode=mode, target_pages=target_pages)
        
        if data and data.get('pages'):
            # Salva screenshot da última página
            last_page = len(data['pages'])
            screenshot_file = os.path.join(output_folder, f"powerbi_screenshot_page{last_page}.png")
            driver.save_screenshot(screenshot_file)
            print(f"\n📸 Screenshot da última página salvo: {screenshot_file}")
            
            # Salva resultados
            saved_files = save_data(data, prefix="ons_powerbi", output_folder=output_folder)
            
            print("\n" + "="*70)
            print("✅ EXTRAÇÃO CONCLUÍDA COM SUCESSO!")
            print("="*70)
            print(f"\n📊 Estatísticas:")
            print(f"  • Páginas processadas: {len(data['pages'])}")
            print(f"  • Total de tabelas: {data['total_tables']}")
            print(f"  • Total de cards/KPIs: {data['total_cards']}")
            print(f"  • Total de gráficos: {data['total_charts']}")
            print(f"\n📁 Pasta de saída: {os.path.abspath(output_folder)}")
            print(f"\n📁 Arquivos gerados ({len(saved_files)}):")
            for f in saved_files:
                file_size = os.path.getsize(f) / 1024  # KB
                filename = os.path.basename(f)
                print(f"  📄 {filename} ({file_size:.1f} KB)")
            
        else:
            print("\n❌ Nenhum dado foi extraído")
            print("Verifique o screenshot e HTML para diagnóstico")
        
    except Exception as e:
        print(f"\n❌ Erro durante execução: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        print("\n🔒 Fechando navegador...")
        driver.quit()
        print("✓ Concluído!")


if __name__ == "__main__":
    main()

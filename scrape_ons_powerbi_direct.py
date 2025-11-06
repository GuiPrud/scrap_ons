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

def extract_specific_class_data(driver, target_class=None, additional_selectors=None):
    """
    Extrai dados organizados por elementos 'series' e seus respectivos 'column setFocusRing'
    Cada série tem um aria-label específico e contém elementos filhos
    
    Args:
        driver: Selenium WebDriver
        target_class: String com a classe CSS a ser buscada dentro de cada série
        additional_selectors: Lista de seletores CSS adicionais para buscar
    """
    # Define classe padrão se não especificada
    if target_class is None:
        target_class = 'column setFocusRing'
    
    # Define seletores adicionais se não especificados
    if additional_selectors is None:
        additional_selectors = []
    
    print(f"\n🎯 Extraindo dados organizados por SERIES > '{target_class}'")
    if additional_selectors:
        print(f"   + Seletores adicionais: {additional_selectors}")
    
    # Script JavaScript para extrair dados organizados por série
    js_extraction = f"""
    function extractSeriesData() {{
        let results = {{
            target_class: '{target_class}',
            additional_selectors: {additional_selectors},
            series: [],
            summary: {{
                total_series: 0,
                total_elements_across_all_series: 0,
                series_with_elements: 0
            }}
        }};
        
        console.log('Procurando elementos com class="series"...');
        
        // Busca todos os elementos com class="series"
        const seriesElements = document.querySelectorAll('[class*="series"]');
        console.log(`Encontradas ${{seriesElements.length}} séries`);
        
        seriesElements.forEach((seriesElement, seriesIndex) => {{
            try {{
                let seriesData = {{
                    series_index: seriesIndex,
                    aria_label: '',
                    series_attributes: {{}},
                    elements: [],
                    series_summary: {{
                        total_elements: 0,
                        elements_with_text: 0,
                        unique_texts: new Set()
                    }}
                }};
                
                // Extrai informações da série
                seriesData.aria_label = seriesElement.getAttribute('aria-label') || 'Sem aria-label';
                
                // Extrai outros atributos importantes da série
                const importantSeriesAttrs = ['class', 'id', 'data-testid', 'role', 'title'];
                importantSeriesAttrs.forEach(attr => {{
                    const value = seriesElement.getAttribute(attr);
                    if (value) {{
                        seriesData.series_attributes[attr] = value;
                    }}
                }});
                
                // Adiciona atributos data-*
                Array.from(seriesElement.attributes).forEach(attribute => {{
                    if (attribute.name.startsWith('data-')) {{
                        seriesData.series_attributes[attribute.name] = attribute.value;
                    }}
                }});
                
                console.log(`Série ${{seriesIndex}}: "${{seriesData.aria_label}}"`);
                
                // Constrói seletores para buscar dentro desta série específica
                let selectors = [];
                
                // Seletor principal - converte espaços em pontos para CSS
                const mainClass = '{target_class}'.replace(/\\s+/g, '.');
                selectors.push('.' + mainClass);
                
                // Seletor alternativo com [class*=]
                selectors.push('[class*="{target_class}"]');
                
                // Adiciona seletores extras se fornecidos
                const additionalSelectors = {additional_selectors};
                if (Array.isArray(additionalSelectors)) {{
                    selectors = selectors.concat(additionalSelectors);
                }}
                
                let allElementsInSeries = new Set(); // Para evitar duplicatas
                
                // Busca elementos dentro desta série específica
                selectors.forEach(selector => {{
                    try {{
                        const elementsInSeries = seriesElement.querySelectorAll(selector);
                        console.log(`  Série ${{seriesIndex}} - Seletor "${{selector}}": ${{elementsInSeries.length}} elementos`);
                        
                        elementsInSeries.forEach(element => {{
                            allElementsInSeries.add(element);
                        }});
                    }} catch (e) {{
                        console.warn(`Erro com seletor "${{selector}}" na série ${{seriesIndex}}:`, e);
                    }}
                }});
                
                const uniqueElementsInSeries = Array.from(allElementsInSeries);
                console.log(`  Série ${{seriesIndex}} - Total elementos únicos: ${{uniqueElementsInSeries.length}}`);
                
                // Processa cada elemento encontrado nesta série
                uniqueElementsInSeries.forEach((element, elementIndex) => {{
                    try {{
                        let elementData = {{
                            element_index: elementIndex,
                            text_content: '',
                            inner_text: '',
                            aria_label: ''
                        }};
                        
                        // Extrai texto principal
                        elementData.text_content = (element.textContent || '').trim();
                        elementData.inner_text = (element.innerText || '').trim();
                        
                        // Extrai aria-label do elemento
                        elementData.aria_label = element.getAttribute('aria-label') || '';
                        
                        // Adiciona aos resultados da série
                        seriesData.elements.push(elementData);
                        
                        // Atualiza sumário da série
                        if (elementData.text_content) {{
                            seriesData.series_summary.elements_with_text++;
                            seriesData.series_summary.unique_texts.add(elementData.text_content);
                        }}
                        
                    }} catch (e) {{
                        console.error(`Erro ao processar elemento ${{elementIndex}} da série ${{seriesIndex}}:`, e);
                    }}
                }});
                
                // Finaliza sumário da série
                seriesData.series_summary.total_elements = seriesData.elements.length;
                seriesData.series_summary.unique_texts = Array.from(seriesData.series_summary.unique_texts);
                
                // Adiciona série aos resultados
                results.series.push(seriesData);
                
                // Atualiza sumário geral
                results.summary.total_elements_across_all_series += seriesData.series_summary.total_elements;
                if (seriesData.series_summary.total_elements > 0) {{
                    results.summary.series_with_elements++;
                }}
                
            }} catch (e) {{
                console.error(`Erro ao processar série ${{seriesIndex}}:`, e);
            }}
        }});
        
        results.summary.total_series = results.series.length;
        
        return results;
    }}
    
    return extractSeriesData();
    """
    
    try:
        data = driver.execute_script(js_extraction)
        
        print(f"✓ Processamento concluído:")
        print(f"  • Classe alvo: '{target_class}'")
        print(f"  • Total de séries encontradas: {data['summary']['total_series']}")
        print(f"  • Séries com elementos: {data['summary']['series_with_elements']}")
        print(f"  • Total de elementos em todas as séries: {data['summary']['total_elements_across_all_series']}")
        
        # Mostra preview de cada série
        if data['series']:
            print(f"\n📋 Preview das séries:")
            for i, series in enumerate(data['series']):
                print(f"\n  📊 Série {i+1}: \"{series['aria_label']}\"")
                print(f"     • Elementos encontrados: {series['series_summary']['total_elements']}")
                print(f"     • Elementos com texto: {series['series_summary']['elements_with_text']}")
                
                # Mostra primeiros elementos de cada série
                if series['elements']:
                    print(f"     • Preview dos primeiros elementos:")
                    for j, element in enumerate(series['elements'][:3]):
                        text_preview = element['text_content'][:60] + '...' if len(element['text_content']) > 60 else element['text_content']
                        print(f"       {j+1}. {text_preview}")
                else:
                    print(f"     • Nenhum elemento '{target_class}' encontrado nesta série")
        
        return data
        
    except Exception as e:
        print(f"❌ Erro ao executar extração: {e}")
        return None

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
            select_date_in_powerbi_calendar(driver, target_date="01/10/2021", date_type="início")

            # Extrai dados da página atual
            page_data = extract_specific_class_data(driver, target_class='column setFocusRing')
            
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
    """Salva os dados extraídos em diferentes formatos - suporta múltiplas páginas e estrutura por séries"""
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
        
        all_series_data = []
        consolidated_elements = []
        
        for page in data['pages']:
            page_num = page.get('page_number', 'unknown')
            
            # Verifica se a página tem estrutura por séries
            if 'series' in page:
                print(f"\n📊 Página {page_num} - Estrutura por séries:")
                
                # Processa cada série
                for series_idx, series in enumerate(page['series']):
                    series_label = series.get('aria_label', f'Serie_{series_idx}')
                    element_count = series.get('series_summary', {}).get('total_elements', 0)
                    
                    print(f"  📈 Série {series_idx + 1}: \"{series_label}\" ({element_count} elementos)")
                    
                    # Salva dados de cada série em arquivo separado
                    if series.get('elements'):
                        # Cria DataFrame para esta série
                        series_data = []
                        for element in series['elements']:
                            row = {
                                'Página': page_num,
                                'Serie_Index': series_idx,
                                'Serie_Label': series_label,
                                'Serie_Aria_Label': series.get('aria_label', ''),
                                'Element_Index': element.get('element_index', ''),
                                'Element_Aria_Label': element.get('aria_label', ''),
                                'Text_Content': element.get('text_content', ''),
                                'Inner_Text': element.get('inner_text', '')
                            }
                            series_data.append(row)
                            consolidated_elements.append(row)
                        
                        if series_data:
                            df_series = pd.DataFrame(series_data)
                            
                            # Limpa o nome da série para usar no nome do arquivo
                            safe_series_name = "".join(c for c in series_label if c.isalnum() or c in (' ', '-', '_')).rstrip()
                            safe_series_name = safe_series_name.replace(' ', '_')[:50]  # Limita tamanho
                            
                            csv_file = os.path.join(output_folder, f"{prefix}_page{page_num}_serie_{series_idx}_{safe_series_name}.csv")
                            df_series.to_csv(csv_file, index=False, encoding='utf-8-sig')
                            print(f"    ✓ {os.path.basename(csv_file)} - {df_series.shape[0]} elementos")
                            saved_files.append(csv_file)
                            
                            all_series_data.append(df_series)
            
            # Compatibilidade com estrutura antiga (tabelas, cards, charts)
            elif any(key in page for key in ['tables', 'cards', 'charts']):
                # Processa estrutura antiga...
                if page.get('tables'):
                    for i, table in enumerate(page['tables']):
                        try:
                            if table['headers']:
                                df = pd.DataFrame(table['rows'], columns=table['headers'])
                            else:
                                df = pd.DataFrame(table['rows'])
                            
                            df.insert(0, 'Página', page_num)
                            csv_file = os.path.join(output_folder, f"{prefix}_page{page_num}_table_{i+1}.csv")
                            df.to_csv(csv_file, index=False, encoding='utf-8-sig')
                            print(f"✓ {csv_file} - {df.shape[0]} linhas × {df.shape[1]} colunas")
                            saved_files.append(csv_file)
                        except Exception as e:
                            print(f"⚠️  Erro ao salvar tabela {i+1} da página {page_num}: {e}")
        
        # 3. Consolida dados de todas as séries
        if consolidated_elements:
            try:
                consolidated_df = pd.DataFrame(consolidated_elements)
                main_dataframe = consolidated_df
                
                consolidated_file = os.path.join(output_folder, f"{prefix}_ALL_SERIES_CONSOLIDATED.csv")
                consolidated_df.to_csv(consolidated_file, index=False, encoding='utf-8-sig')
                print(f"\n✓ {consolidated_file} - {consolidated_df.shape[0]} elementos × {consolidated_df.shape[1]} colunas")
                print("  (Todos os elementos de todas as séries de todas as páginas)")
                saved_files.append(consolidated_file)
                
                # Mostra preview consolidado
                print("\n" + "="*80)
                print("Preview dos Dados Consolidados por Série:")
                print("="*80)
                # Mostra apenas algumas colunas essenciais para caber na tela
                display_columns = ['Página', 'Serie_Label', 'Element_Aria_Label', 'Text_Content']
                available_columns = [col for col in display_columns if col in consolidated_df.columns]
                print(consolidated_df[available_columns].head(10).to_string())
                print("="*80 + "\n")
                
                # Estatísticas por série
                print("📊 Estatísticas por Série:")
                series_stats = consolidated_df.groupby(['Serie_Label', 'Serie_Aria_Label']).agg({
                    'Text_Content': 'count',
                    'Element_Aria_Label': lambda x: sum(1 for val in x if val.strip())
                }).rename(columns={
                    'Text_Content': 'Total_Elementos',
                    'Element_Aria_Label': 'Elementos_com_Aria_Label'
                })
                print(series_stats.to_string())
                
            except Exception as e:
                print(f"⚠️  Erro ao consolidar dados das séries: {e}")
        
        # 4. DataFrame consolidado em Pickle (para uso em Python/Pandas)
        if main_dataframe is not None:
            try:
                pickle_file = os.path.join(output_folder, f"{prefix}_dataframe.pkl")
                main_dataframe.to_pickle(pickle_file)
                print(f"\n✓ {pickle_file} - DataFrame salvo em formato Pickle")
                print(f"  💡 Para carregar: df = pd.read_pickle('{pickle_file}')")
                saved_files.append(pickle_file)
            except Exception as e:
                print(f"⚠️  Erro ao salvar DataFrame pickle: {e}")
        
        # 5. DataFrame consolidado em Excel
        if main_dataframe is not None:
            try:
                if len(main_dataframe) <= 1000000:  # Limite do Excel
                    excel_file = os.path.join(output_folder, f"{prefix}_dataframe.xlsx")
                    
                    # Cria Excel com múltiplas abas
                    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                        # Aba principal com todos os dados
                        main_dataframe.to_excel(writer, sheet_name='Todos_Dados', index=False)
                        
                        # Aba por série (máximo 10 séries para não sobrecarregar)
                        unique_series = main_dataframe['Serie_Label'].unique()[:10]
                        for series_label in unique_series:
                            series_df = main_dataframe[main_dataframe['Serie_Label'] == series_label]
                            safe_sheet_name = "".join(c for c in series_label if c.isalnum() or c in (' ', '-', '_'))[:31]
                            series_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                    
                    print(f"✓ {excel_file} - DataFrame salvo em formato Excel com múltiplas abas")
                    saved_files.append(excel_file)
                else:
                    print(f"⚠️  DataFrame muito grande ({len(main_dataframe)} linhas) para Excel - use CSV ou Pickle")
            except Exception as e:
                print(f"⚠️  Erro ao salvar DataFrame Excel: {e}")
                print(f"   (Instale openpyxl: pip install openpyxl)")
    
    else:
        # Estrutura de página única - mantém compatibilidade
        print("\n📄 Processando página única...")
        # [Código existente para estrutura antiga permanece o mesmo]
        pass
    
    return saved_files

def select_date_in_powerbi_calendar(driver, target_date="01/10/2021", date_type="início"):
    """
    Seleciona uma data específica no calendário do Power BI
    
    Args:
        driver: Selenium WebDriver
        target_date: Data no formato DD/MM/AAAA
        date_type: Tipo de data ('início' ou 'fim') para identificar o slicer correto
    """
    print(f"\n📅 Selecionando data {target_date} ({date_type})...")
    
    try:
        # Parse da data
        from datetime import datetime
        date_obj = datetime.strptime(target_date, "%d/%m/%Y")
        day = date_obj.day
        month = date_obj.month
        year = date_obj.year
        
        print(f"  • Data parseada: {day:02d}/{month:02d}/{year}")
        
        # 1. Encontra o input de data específico
        wait = WebDriverWait(driver, 10)
        
        # Seletores para encontrar o input de data correto
        date_input_selectors = [
            f"//input[contains(@aria-label, 'Data de {date_type}')]",
            f"//input[contains(@aria-label, '{date_type}') and contains(@class, 'date-slicer-datepicker')]",
            "//input[contains(@class, 'date-slicer-datepicker') and contains(@aria-label, 'Data de início')]",
            "//input[contains(@class, 'item-fill ng-valid date-slicer-datepicker')]"
        ]
        
        date_input = None
        for selector in date_input_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                if elements:
                    date_input = elements[0]
                    aria_label = date_input.get_attribute('aria-label') or ''
                    print(f"  ✓ Date input encontrado: {aria_label[:50]}...")
                    break
            except:
                continue
        
        if not date_input:
            print("  ❌ Date input não encontrado")
            return False
        
        # 2. Limpa o campo e insere a nova data
        print(f"  • Definindo data para {target_date}...")
        
        # Scroll até o elemento
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", date_input)
        time.sleep(0.5)
        
        # Limpa o campo atual
        date_input.clear()
        time.sleep(0.3)
        
        # Insere a nova data
        date_input.send_keys(target_date)
        time.sleep(0.5)
        
        # Dispara evento de mudança para garantir que o Power BI processe
        driver.execute_script("""
            var event = new Event('input', { bubbles: true });
            arguments[0].dispatchEvent(event);
            
            var changeEvent = new Event('change', { bubbles: true });
            arguments[0].dispatchEvent(changeEvent);
            
            var blurEvent = new Event('blur', { bubbles: true });
            arguments[0].dispatchEvent(blurEvent);
        """, date_input)
        
        time.sleep(1)
        
        # Verifica se a data foi definida corretamente
        current_value = date_input.get_attribute('value')
        print(f"  • Valor atual do campo: '{current_value}'")
        
        if current_value == target_date:
            print(f"  ✅ Data {target_date} definida com sucesso!")
            
            # Aguarda o Power BI processar a mudança
            print("  ⏳ Aguardando Power BI processar a mudança...")
            time.sleep(3)
            
            return True
        else:
            print(f"  ⚠️  Data definida mas valor diferente: '{current_value}' != '{target_date}'")
            
            # Tenta abordagem alternativa com JavaScript direto
            print("  • Tentando abordagem alternativa com JavaScript...")
            
            driver.execute_script(f"""
                arguments[0].value = '{target_date}';
                arguments[0].setAttribute('value', '{target_date}');
                
                // Dispara múltiplos eventos para garantir detecção
                ['input', 'change', 'blur', 'keyup'].forEach(eventType => {{
                    var event = new Event(eventType, {{ bubbles: true, cancelable: true }});
                    arguments[0].dispatchEvent(event);
                }});
                
                // Força atualização Angular/React se existir
                if (window.angular) {{
                    var scope = window.angular.element(arguments[0]).scope();
                    if (scope) {{
                        scope.$apply();
                    }}
                }}
            """, date_input)
            
            time.sleep(2)
            
            # Verifica novamente
            new_value = date_input.get_attribute('value')
            if new_value == target_date:
                print(f"  ✅ Data {target_date} definida com JavaScript!")
                time.sleep(3)
                return True
            else:
                print(f"  ❌ Falha ao definir data. Valor final: '{new_value}'")
                return False
        
    except Exception as e:
        print(f"  ❌ Erro ao selecionar data: {e}")
        import traceback
        print(f"  Detalhes: {traceback.format_exc()}")
        return False


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
        
        # NOVO: Seleciona data no calendário ANTES de extrair dados
        print("\n" + "="*70)
        print("  CONFIGURANDO FILTROS DE DATA")
        print("="*70)
        
        # Seleciona data de início
        if select_date_in_powerbi_calendar(driver, target_date="01/10/2021", date_type="início"):
            print("✅ Data de início configurada!")
            
            # Aguarda o dashboard atualizar após mudança de filtro
            print("⏳ Aguardando dashboard atualizar...")
            time.sleep(5)
            
            # Aguarda novamente o carregamento após filtro
            wait_for_powerbi_load(driver, timeout=30)
        else:
            print("⚠️  Falha ao configurar data de início, continuando mesmo assim...")
        
        # Salva screenshot da primeira página (APÓS configurar filtros)
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

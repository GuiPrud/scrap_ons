import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import os
import sys

url = "https://www.ons.org.br/Paginas/faq_curtailment.aspx"

options = Options()
# Removendo headless para debug - você pode adicionar de volta depois
# options.add_argument('--headless')

# Adiciona argumentos para contornar políticas de segurança
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-extensions')
options.add_argument('--disable-gpu')
options.add_argument('--remote-debugging-port=9222')

def get_chrome_driver():
    """Tenta diferentes formas de inicializar o Chrome driver"""
    
    # Método 1: Tentar com ChromeDriver baixado manualmente
    possible_paths = [
        r"C:\chromedriver\chromedriver.exe",
        r"C:\Program Files\chromedriver\chromedriver.exe",
        r"C:\Users\%USERNAME%\Downloads\chromedriver.exe",
        "./chromedriver.exe",
        "chromedriver.exe"
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            print(f"Tentando usar ChromeDriver em: {path}")
            try:
                service = Service(path)
                return webdriver.Chrome(service=service, options=options)
            except Exception as e:
                print(f"Falhou com {path}: {e}")
                continue
    
    # Método 2: Tentar sem especificar service (padrão)
    print("Tentando inicializar Chrome sem service específico...")
    try:
        return webdriver.Chrome(options=options)
    except Exception as e:
        print(f"Falhou método padrão: {e}")
    
    # Método 3: Tentar com Firefox como alternativa
    print("Tentando Firefox como alternativa...")
    try:
        from selenium.webdriver.firefox.options import Options as FirefoxOptions
        firefox_options = FirefoxOptions()
        if '--headless' in [arg for arg in options.arguments]:
            firefox_options.add_argument('--headless')
        return webdriver.Firefox(options=firefox_options)
    except Exception as e:
        print(f"Firefox também falhou: {e}")
    
    return None


def close_cookie_banner(driver):
    """Tenta fechar banner de cookies/política de privacidade"""
    print("Tentando fechar banner de cookies...")
    try:
        # Procura por botões comuns de aceitar cookies
        cookie_selectors = [
            "//button[contains(text(), 'Accept')]",
            "//button[contains(text(), 'Aceitar')]", 
            "//button[@id='onetrust-accept-btn-handler']",
            "//button[contains(@class, 'accept')]",
            "//*[@id='onetrust-close-btn-container']//button",
            "//button[contains(@class, 'cookie')]",
            "//*[contains(@id, 'onetrust')]//button",
            "//button[contains(@class, 'onetrust')]",
            "//*[@id='onetrust-policy']//button",
            "//button[contains(@aria-label, 'close')]",
            "//button[contains(@title, 'close')]"
        ]
        
        for selector in cookie_selectors:
            try:
                cookie_btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                cookie_btn.click()
                print("Banner de cookies fechado!")
                time.sleep(1)
                return True
            except:
                continue
                
        # Se não encontrou botão, tenta remover o elemento diretamente
        print("Tentando remover banner via JavaScript...")
        driver.execute_script("""
            var banner = document.getElementById('onetrust-policy');
            if (banner) {
                banner.remove();
                console.log('Banner onetrust-policy removido');
            }
            var consent = document.getElementById('onetrust-consent-sdk');
            if (consent) {
                consent.remove();
                console.log('SDK onetrust removido');
            }
            var group = document.getElementById('onetrust-pc-sdk');
            if (group) {
                group.remove();
                console.log('PC SDK removido');
            }
            // Remove qualquer overlay que possa estar bloqueando
            var overlays = document.querySelectorAll('[id*="onetrust"], [class*="onetrust"]');
            overlays.forEach(function(el) {
                el.remove();
            });
        """)
        print("Script de remoção executado")
        return True
        
    except Exception as e:
        print(f"Erro ao tentar fechar banner: {e}")
    
    print("Nenhum banner de cookies encontrado ou já foi fechado")
    return False

print("=== Extração de dados da ONS ===")

# def extract_data():
session = requests.Session()
response = session.get(url)
if response.status_code != 200:
    print(f"Erro ao acessar a página: {response.status_code}")
print("Página acessada com sucesso!")

try:
    print("Iniciando o driver do Chrome...")
    driver = get_chrome_driver()

    if not driver:
        print("\n" + "="*50)
        print("ERRO: Não foi possível inicializar nenhum driver!")
        print("="*50)
        print("Soluções possíveis:")
        print("1. Baixe o ChromeDriver manualmente:")
        print("   - Vá para: https://chromedriver.chromium.org/")
        print("   - Baixe a versão compatível com seu Chrome")
        print("   - Extraia para C:\\chromedriver\\chromedriver.exe")
        print("")
        print("2. Ou instale o webdriver-manager:")
        print("   pip install webdriver-manager")
        print("")
        print("3. Ou use Firefox:")
        print("   - Instale o Firefox")
        print("   - pip install selenium")
        print("="*50)
        sys.exit(1)

    print(f"Navegando para: {url}")
    driver.get(url)
    
    print("Aguardando a página carregar...")
    # Aguarda alguns segundos para a página carregar completamente
    time.sleep(5)

    close_cookie_banner(driver)

    print("Aguardando página carregar completamente...")
    wait = WebDriverWait(driver, 30)  # Aumentado para 30 segundos
    
    # Aguarda o body estar carregado
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print("✓ Body carregado")
    
    # Aguarda um pouco mais para JavaScript executar
    time.sleep(5)
 
     # Salva o HTML completo para análise
    page_source = driver.page_source
    with open('ons_page_source.html', 'w', encoding='utf-8') as f:
        f.write(page_source)
    print("✓ HTML da página salvo em 'ons_page_source.html'")
    
    # Procura por diferentes seletores possíveis
    print("\nProcurando elementos na página...")
    possible_selectors = [
        ("//div[@class='displayAreaViewport']", "displayAreaViewport original"),
        ("//div[contains(@class, 'displayArea')]", "displayArea (contains)"),
        ("//div[contains(@class, 'viewport')]", "viewport (contains)"),
        ("//table", "Qualquer tabela"),
        ("//div[@id='content']", "Content div"),
        ("//div[@class='ms-rtestate-field']", "SharePoint RTE field"),
        ("//div[contains(@class, 'PublishingPageContent')]", "Publishing content"),
        ("//iframe", "Iframe"),
    ]
    
    found_elements = []
    for selector, description in possible_selectors:
        try:
            elements = driver.find_elements(By.XPATH, selector)
            if elements:
                print(f"✓ Encontrado: {description} - {len(elements)} elemento(s)")
                found_elements.append((selector, description, elements))
            else:
                print(f"✗ Não encontrado: {description}")
        except Exception as e:
            print(f"✗ Erro ao buscar {description}: {e}")
    
    # Tenta encontrar o elemento principal
    info = None
    table_html = None
    
    # Primeiro tenta o seletor original
    try:
        print("\nTentando seletor original...")
        info = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='displayAreaViewport']")))
        print("✓ displayAreaViewport encontrado!")
        table_html = info.get_attribute('outerHTML')
        
    except Exception as e:
        print(f"⚠️  Seletor original falhou: {e}")
        
        # Tenta seletores alternativos
        for selector, description, elements in found_elements:
            if elements and selector != "//iframe":  # Pula iframe por enquanto
                try:
                    print(f"\nTentando extrair de: {description}")
                    info = elements[0]
                    table_html = info.get_attribute('outerHTML')
                    
                    # Salva para análise
                    with open(f'ons_element_{description.replace(" ", "_")}.html', 'w', encoding='utf-8') as f:
                        f.write(table_html or '')
                    print(f"✓ HTML extraído e salvo em 'ons_element_{description.replace(' ', '_')}.html'")
                    
                    if table_html and len(table_html) > 100:  # Verifica se tem conteúdo substancial
                        print(f"✓ Conteúdo encontrado ({len(table_html)} caracteres)")
                        break
                except Exception as ex:
                    print(f"✗ Falha ao extrair de {description}: {ex}")
                    continue
        
    # Se ainda não encontrou, tenta iframe
        if not table_html or len(table_html) < 100:
            print("\nVerificando iframes...")
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            print(f"Encontrados {len(iframes)} iframes")
            
            for i, iframe in enumerate(iframes):
                try:
                    print(f"\nMudando para iframe {i+1}...")
                    iframe_src = iframe.get_attribute('src')
                    print(f"URL do iframe: {iframe_src}")
                    
                    # Detecta se é Power BI
                    if 'powerbi.com' in (iframe_src or ''):
                        print("⚡ Power BI detectado!")
                        print("Aplicando estratégia específica para Power BI...")
                        
                        driver.switch_to.frame(iframe)
                        
                        # Aguarda mais tempo para Power BI carregar
                        print("Aguardando Power BI carregar (30 segundos)...")
                        time.sleep(30)
                        
                        # Procura por elementos específicos do Power BI
                        powerbi_selectors = [
                            "//div[contains(@class, 'visualContainer')]",
                            "//div[contains(@class, 'visual')]",
                            "//div[@role='figure']",
                            "//div[contains(@class, 'card')]",
                            "//svg",
                            "//text[contains(@class, 'label')]",
                            "//div[contains(@class, 'tableEx')]",
                            "//div[contains(@class, 'pivotTable')]",
                        ]
                        
                        powerbi_data = {}
                        
                        for selector in powerbi_selectors:
                            try:
                                elements = driver.find_elements(By.XPATH, selector)
                                if elements:
                                    print(f"✓ Encontrados {len(elements)} elementos: {selector}")
                                    
                                    # Extrai texto de cada elemento
                                    for idx, elem in enumerate(elements[:10]):  # Limita a 10 primeiros
                                        try:
                                            text = elem.text
                                            if text and len(text.strip()) > 0:
                                                powerbi_data[f"{selector}_{idx}"] = text
                                                print(f"  Elemento {idx}: {text[:100]}...")
                                        except:
                                            pass
                            except:
                                pass
                        
                        # Salva dados extraídos do Power BI
                        if powerbi_data:
                            with open(f'ons_powerbi_data.json', 'w', encoding='utf-8') as f:
                                json.dump(powerbi_data, f, indent=2, ensure_ascii=False)
                            print(f"✓ Dados do Power BI salvos em 'ons_powerbi_data.json'")
                        
                        # Tenta obter HTML completo do iframe
                        iframe_body = driver.find_element(By.TAG_NAME, "body")
                        table_html = iframe_body.get_attribute('outerHTML')
                        
                        # Tenta executar JavaScript para extrair dados de tabelas
                        try:
                            print("Tentando extrair dados via JavaScript...")
                            js_data = driver.execute_script("""
                                var data = [];
                                
                                // Procura por tabelas visíveis
                                var tables = document.querySelectorAll('table');
                                tables.forEach((table, index) => {
                                    var rows = [];
                                    table.querySelectorAll('tr').forEach(tr => {
                                        var cells = [];
                                        tr.querySelectorAll('td, th').forEach(cell => {
                                            cells.push(cell.innerText);
                                        });
                                        if (cells.length > 0) rows.push(cells);
                                    });
                                    if (rows.length > 0) {
                                        data.push({table: index, data: rows});
                                    }
                                });
                                
                                // Procura por divs com dados tabulares
                                var visualContainers = document.querySelectorAll('[class*="visual"]');
                                visualContainers.forEach((container, index) => {
                                    var text = container.innerText;
                                    if (text && text.length > 10) {
                                        data.push({visual: index, text: text});
                                    }
                                });
                                
                                return data;
                            """)
                            
                            if js_data:
                                print(f"✓ Dados extraídos via JavaScript: {len(js_data)} elementos")
                                with open('ons_powerbi_js_data.json', 'w', encoding='utf-8') as f:
                                    json.dump(js_data, f, indent=2, ensure_ascii=False)
                                print("✓ Dados salvos em 'ons_powerbi_js_data.json'")
                                
                                # Processa tabelas extraídas
                                for item in js_data:
                                    if 'table' in item and 'data' in item:
                                        try:
                                            df = pd.DataFrame(item['data'][1:], columns=item['data'][0])
                                            csv_name = f"ons_powerbi_table_{item['table']}.csv"
                                            df.to_csv(csv_name, index=False, encoding='utf-8')
                                            print(f"✓ Tabela {item['table']} salva em '{csv_name}'")
                                            print(f"  Shape: {df.shape}")
                                            print(df.head())
                                        except Exception as e:
                                            print(f"⚠️  Erro ao processar tabela {item.get('table')}: {e}")
                        
                        except Exception as js_error:
                            print(f"⚠️  Erro ao executar JavaScript: {js_error}")
                        
                        # Volta para o contexto principal
                        driver.switch_to.default_content()
                        
                        if table_html and len(table_html) > 100:
                            print(f"✓ Conteúdo Power BI extraído ({len(table_html)} caracteres)")
                            break
                    
                    else:
                        # Iframe normal (não Power BI)
                        driver.switch_to.frame(iframe)
                        time.sleep(2)
                        
                        # Tenta encontrar o elemento dentro do iframe
                        iframe_body = driver.find_element(By.TAG_NAME, "body")
                        table_html = iframe_body.get_attribute('outerHTML')
                        
                        with open(f'ons_iframe_{i+1}.html', 'w', encoding='utf-8') as f:
                            f.write(table_html or '')
                        print(f"✓ Conteúdo do iframe {i+1} salvo")
                        
                        # Volta para o contexto principal
                        driver.switch_to.default_content()
                        
                        if table_html and len(table_html) > 100:
                            print(f"✓ Conteúdo substancial encontrado no iframe {i+1}")
                            break
                        
                except Exception as ex:
                    print(f"✗ Erro no iframe {i+1}: {ex}")
                    driver.switch_to.default_content()
                    continue
    
    if table_html:
        print(f"\n{'='*60}")
        print("✓ DADOS EXTRAÍDOS COM SUCESSO!")
        print(f"{'='*60}")
        print(f"Tamanho do HTML: {len(table_html)} caracteres")
        
        # Salva o resultado final
        with open('ons_data_final.html', 'w', encoding='utf-8') as f:
            f.write(table_html)
        print("✓ Dados salvos em 'ons_data_final.html'")
        
        # Tenta parsear com BeautifulSoup
        soup = BeautifulSoup(table_html, 'html.parser')
        
        # Procura por tabelas
        tables = soup.find_all('table')
        print(f"\n✓ Encontradas {len(tables)} tabela(s) no conteúdo")
        
        if tables:
            for i, table in enumerate(tables):
                print(f"\nProcessando tabela {i+1}...")
                try:
                    df = pd.read_html(str(table))[0]
                    print(f"Shape: {df.shape}")
                    print(df.head())
                    
                    # Salva em CSV
                    df.to_csv(f'ons_table_{i+1}.csv', index=False, encoding='utf-8')
                    print(f"✓ Tabela {i+1} salva em 'ons_table_{i+1}.csv'")
                except Exception as e:
                    print(f"⚠️  Erro ao processar tabela {i+1}: {e}")
    else:
        print("\n❌ Nenhum dado foi extraído")
        print("Verifique os arquivos HTML salvos para diagnóstico")
            
except Exception as e:
    print(f"Erro durante a execução: {type(e).__name__}: {e}")
    print("Detalhes do erro:")
    import traceback
    traceback.print_exc()
    
finally:
    if driver:
        print("Fechando o driver...")
        driver.quit()
        print("Driver fechado com sucesso!")
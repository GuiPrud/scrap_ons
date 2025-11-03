import time
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import os

# URL do Power BI embedado
POWERBI_URL = "https://app.powerbi.com/view?r=eyJrIjoiYmU0ODUxNGMtNWU2MS00YTM5LThkMGYtNWFkYWQzYmU3ZWY2IiwidCI6IjNhZGVlNWZjLTkzM2UtNDkxMS1hZTFiLTljMmZlN2I4NDQ0OCIsImMiOjR9"


def create_output_folder():
    """Cria pasta para salvar os arquivos gerados"""
    folder_name = f"extracao_powerbi"
    
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        print(f"📁 Pasta criada: {folder_name}")
    
    return folder_name


def setup_driver():
    """Configura o Chrome com opções para capturar requisições de rede"""
    options = Options()
    # options.add_argument('--headless')  # Descomente para executar sem interface
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-blink-features=AutomationControlled')
    
    # Habilita logging de performance para capturar requisições
    options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
    
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver

def wait_for_powerbi_load(driver, timeout=60):
    """Aguarda o Power BI carregar completamente"""
    print("Aguardando Power BI carregar...")
    
    # Aguarda elementos específicos do Power BI
    wait = WebDriverWait(driver, timeout)
    
    try:
        # Aguarda iframe principal carregar
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe")))
        print("✓ Iframe carregado")
        
        # Aguarda visualizações carregarem
        time.sleep(15)  # Power BI precisa de tempo para renderizar
        
        # Tenta encontrar elementos visuais
        visual_loaded = wait.until(EC.presence_of_element_located((
            By.XPATH, 
            "//div[contains(@class, 'visual') or contains(@class, 'card')]"
        )))
        print("✓ Visualizações carregadas")
        
        return True
        
    except Exception as e:
        print(f"⚠️  Timeout ao aguardar Power BI: {e}")
        driver.switch_to.default_content()
        return False

def extract_network_requests(driver, output_folder="."):
    """Extrai requisições de rede que podem conter dados"""
    print("\nExtraindo requisições de rede...")
    
    logs = driver.get_log('performance')
    network_data = []
    
    for entry in logs:
        try:
            log = json.loads(entry['message'])['message']
            
            # Procura por requisições de resposta
            if log['method'] == 'Network.responseReceived':
                response = log['params']['response']
                url = response.get('url', '')
                
                # Filtra requisições do Power BI que podem conter dados
                if any(keyword in url.lower() for keyword in ['query', 'data', 'api', 'execute']):
                    network_data.append({
                        'url': url,
                        'status': response.get('status'),
                        'mimeType': response.get('mimeType'),
                        'requestId': log['params']['requestId']
                    })
                    print(f"✓ Requisição encontrada: {url[:100]}...")
        
        except Exception as e:
            continue
    
    # Salva requisições encontradas
    if network_data:
        filepath = os.path.join(output_folder, 'powerbi_network_requests.json')
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(network_data, f, indent=2)
        print(f"✓ {len(network_data)} requisições salvas em '{filepath}'")
    
    return network_data

def extract_visual_data(driver, output_folder="."):
    """Extrai dados visíveis dos elementos do Power BI"""
    print("\nExtraindo dados visuais...")
    
    all_data = {
        'tables': [],
        'cards': [],
        'charts': [],
        'text_elements': []
    }
    
    try:
        # 1. Procura por tabelas
        tables = driver.find_elements(By.XPATH, "//table | //div[contains(@class, 'tableEx')]")
        print(f"Encontradas {len(tables)} tabelas")
        
        for idx, table in enumerate(tables):
            try:
                # Tenta extrair via pandas
                table_html = table.get_attribute('outerHTML')
                df = pd.read_html(table_html)[0]
                
                all_data['tables'].append({
                    'index': idx,
                    'shape': df.shape,
                    'data': df.to_dict('records')
                })
                
                # Salva CSV
                csv_file = os.path.join(output_folder, f'powerbi_table_{idx}.csv')
                df.to_csv(csv_file, index=False, encoding='utf-8')
                print(f"✓ Tabela {idx} salva: {csv_file} - Shape: {df.shape}")
                
            except Exception as e:
                print(f"⚠️  Erro ao processar tabela {idx}: {e}")
        
        # 2. Procura por cards/KPIs
        cards = driver.find_elements(By.XPATH, "//div[contains(@class, 'card') or contains(@class, 'kpi')]")
        print(f"Encontrados {len(cards)} cards/KPIs")
        
        for idx, card in enumerate(cards):
            try:
                text = card.text.strip()
                if text:
                    all_data['cards'].append({
                        'index': idx,
                        'text': text
                    })
                    print(f"  Card {idx}: {text[:100]}")
            except:
                pass
        
        # 3. Procura por elementos de texto (labels, títulos)
        text_elements = driver.find_elements(By.XPATH, 
            "//text | //span[contains(@class, 'label')] | //div[@role='heading']")
        
        unique_texts = set()
        for elem in text_elements:
            try:
                text = elem.text.strip()
                if text and len(text) > 2 and text not in unique_texts:
                    unique_texts.add(text)
                    all_data['text_elements'].append(text)
            except:
                pass
        
        print(f"Encontrados {len(unique_texts)} elementos de texto únicos")
        
        # 4. Executa JavaScript para extrair dados estruturados
        print("\nExecutando JavaScript para extrair dados...")
        js_data = driver.execute_script("""
            var data = {
                visualData: [],
                tableData: [],
                textData: []
            };
            
            // Extrai de visualizações
            document.querySelectorAll('[class*="visual"]').forEach((visual, idx) => {
                var text = visual.innerText;
                if (text && text.length > 5) {
                    data.visualData.push({index: idx, content: text});
                }
            });
            
            // Extrai tabelas
            document.querySelectorAll('table').forEach((table, idx) => {
                var rows = [];
                table.querySelectorAll('tr').forEach(tr => {
                    var cells = Array.from(tr.querySelectorAll('td, th')).map(cell => cell.innerText);
                    if (cells.length > 0) rows.push(cells);
                });
                if (rows.length > 0) data.tableData.push({index: idx, rows: rows});
            });
            
            // Extrai todos os textos visíveis
            var allText = document.body.innerText;
            data.textData = allText.split('\\n').filter(line => line.trim().length > 0);
            
            return data;
        """)
        
        if js_data:
            print(f"✓ Dados JavaScript extraídos:")
            print(f"  - {len(js_data.get('visualData', []))} visualizações")
            print(f"  - {len(js_data.get('tableData', []))} tabelas")
            print(f"  - {len(js_data.get('textData', []))} linhas de texto")
            
            # Processa tabelas do JavaScript
            for table_data in js_data.get('tableData', []):
                try:
                    rows = table_data['rows']
                    if rows:
                        df = pd.DataFrame(rows[1:], columns=rows[0])
                        csv_file = os.path.join(output_folder, f"powerbi_js_table_{table_data['index']}.csv")
                        df.to_csv(csv_file, index=False, encoding='utf-8')
                        print(f"✓ Tabela JS {table_data['index']} salva: {csv_file}")
                except Exception as e:
                    print(f"⚠️  Erro ao processar tabela JS: {e}")
            
            all_data['javascript'] = js_data
    
    except Exception as e:
        print(f"⚠️  Erro durante extração visual: {e}")
    
    # Salva todos os dados
    filepath = os.path.join(output_folder, 'powerbi_all_data.json')
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, indent=2, ensure_ascii=False)
    print(f"\n✓ Todos os dados salvos em '{filepath}'")
    
    return all_data

def main():
    print("="*60)
    print("EXTRATOR DE DADOS DO POWER BI")
    print("="*60)
    
    # Cria pasta para salvar os arquivos
    output_folder = create_output_folder()
    
    driver = setup_driver()
    
    try:
        print(f"\nAcessando Power BI: {POWERBI_URL}")
        driver.get(POWERBI_URL)
        
        # Aguarda Power BI carregar
        if wait_for_powerbi_load(driver):
            
            # Extrai dados visuais
            visual_data = extract_visual_data(driver, output_folder=output_folder)
            
            # Volta para contexto padrão
            driver.switch_to.default_content()
            
            # Extrai requisições de rede
            network_data = extract_network_requests(driver, output_folder=output_folder)
            
            print("\n" + "="*60)
            print("EXTRAÇÃO CONCLUÍDA!")
            print("="*60)
            print(f"\n📁 Pasta de saída: {os.path.abspath(output_folder)}")
            print("\nArquivos gerados:")
            print("  - powerbi_all_data.json")
            print("  - powerbi_network_requests.json")
            print("  - powerbi_table_*.csv (se encontradas)")
            print("  - powerbi_js_table_*.csv (se encontradas)")
            
        else:
            print("❌ Não foi possível carregar o Power BI completamente")
    
    except Exception as e:
        print(f"\n❌ Erro: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        print("\nPressione Enter para fechar o navegador...")
        input()
        driver.quit()

if __name__ == "__main__":
    main()

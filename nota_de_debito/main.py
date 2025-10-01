import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# --- CONFIGURAÇÃO ---
TITULO_DA_PAGINA_ALVO = "Fiscal | Extranet"

# --- 1. CONECTAR AO NAVEGADOR ---
print("Conectando ao navegador...")
try:
    edge_options = Options()
    edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Edge(options=edge_options)
    print(f"Conectado. A aba inicial é: '{driver.title}'")
except Exception as e:
    print(f"\n--- ERRO NA CONEXÃO ---\nDetalhes: {e}")
    exit()

# --- 2. ENCONTRAR E MUDAR PARA A ABA CORRETA ---
aba_correta_encontrada = False
for handle in driver.window_handles:
    driver.switch_to.window(handle)
    if TITULO_DA_PAGINA_ALVO.lower() in driver.title.lower():
        print(f"Aba correta encontrada e selecionada: '{driver.title}'")
        aba_correta_encontrada = True
        break
if not aba_correta_encontrada:
    print(f"\nAVISO: Nenhuma aba com o título '{TITULO_DA_PAGINA_ALVO}' foi encontrada.")

# --- 3. LÓGICA PRINCIPAL COM LOOP DE PAGINAÇÃO ---
dados_processados = []
pagina_atual = 1

# --- NOVO LOOP DE PAGINAÇÃO ---
while True:
    try:
        if not aba_correta_encontrada:
            break

        print(f"\n--- PROCESSANDO PÁGINA {pagina_atual} ---")
        
        wait = WebDriverWait(driver, 20)
        
        seletor_linhas_dados = "//div[contains(@class, 'flora--c-gqwkJN-ihfFBCg-css') and .//input[starts-with(@data-testid, 'select-debit-note-')]]"
        
        # Espera as linhas da página ATUAL aparecerem
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, seletor_linhas_dados)))
        except TimeoutException:
            print("Tempo esgotado. Nenhuma linha de dados encontrada nesta página. Encerrando.")
            break

        # Pega uma referência a um elemento da página atual para saber quando ela recarregar
        elemento_referencia = driver.find_element(By.XPATH, seletor_linhas_dados)
        
        total_linhas = len(driver.find_elements(By.XPATH, seletor_linhas_dados))
        print(f"Encontradas {total_linhas} linhas de dados para processar.")
        
        # SEU CÓDIGO DE SUCESSO PARA PROCESSAR UMA PÁGINA
        for i in range(total_linhas):
            try:
                elementos_linha = driver.find_elements(By.XPATH, seletor_linhas_dados)
                linha_atual_element = elementos_linha[i]
                
                celulas_texto = [span.text for span in linha_atual_element.find_elements(By.XPATH, ".//span[contains(@class, 'flora--c-LLdDZ')]")]
                
                if not celulas_texto or len(celulas_texto) < 6:
                    print(f"  -> Linha {i+1} ignorada (sem dados ou cabeçalho).")
                    continue

                print(f"Processando linha {i+1}/{total_linhas} | Loja: {celulas_texto[0]}")
                
                status_arquivo = "Erro ao clicar"
                try:
                    botao_pdf = linha_atual_element.find_element(By.XPATH, ".//button[text()='PDF']")
                    
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", botao_pdf)
                    time.sleep(0.5)
                    
                    actions = ActionChains(driver)
                    actions.move_to_element(botao_pdf).double_click().perform()
                    
                    print(f"  -> Botão PDF clicado (duplo). Aguardando 2 segundos...")
                    time.sleep(2)
                    status_arquivo = "Download iniciado"

                except Exception as click_error:
                    print(f"  -> AVISO: Falha ao clicar no botão da linha {i+1}: {click_error}")
                
                dados_para_df = celulas_texto + [status_arquivo]
                dados_processados.append(dados_para_df)
            
            except Exception as loop_error:
                print(f"  -> ERRO inesperado no loop na linha {i+1}: {loop_error}")

        # --- NOVA LÓGICA PARA NAVEGAR PARA A PRÓXIMA PÁGINA ---
        print("\nProcessamento da página concluído. Verificando próxima página...")
        try:
            # Encontra o botão da próxima página (o irmão do botão que está ativo)
            proxima_pagina_btn = driver.find_element(By.XPATH, "//button[@aria-current='true']/following-sibling::button[1]")
            
            # Garante que o botão não é o "..."
            if "..." in proxima_pagina_btn.text:
                print("Fim das páginas sequenciais. Encerrando.")
                break

            print("Próxima página encontrada. Clicando...")
            proxima_pagina_btn.click()
            
            # Espera a página recarregar (esperando o elemento antigo desaparecer)
            wait.until(EC.staleness_of(elemento_referencia))
            
            pagina_atual += 1

        except NoSuchElementException:
            print("Não há mais botões de próxima página. Automação concluída.")
            break # Encerra o loop principal 'while True'

    except Exception as e:
        print(f"\n--- ERRO NO PROCESSAMENTO GERAL DA PÁGINA {pagina_atual} ---")
        print(f"Ocorreu um erro inesperado: {e}")
        break

# --- 4. ORGANIZAR OS DADOS COLETADOS DE TODAS AS PÁGINAS ---
if dados_processados:
    print(f"\n--- RELATÓRIO FINAL ---")
    print(f"Sucesso! {len(dados_processados)} linhas foram processadas no total.")
    colunas = ['LOJA', 'REFERÊNCIA SAP', 'NÚMERO DUPLIC.', 'DATA DE EMISSÃO', 'VENCIMENTO', 'VALOR', 'ARQUIVO']
    df = pd.DataFrame(dados_processados)
    df.columns = colunas[:len(df.columns)]
    print(df)

print("\nScript finalizado. O navegador continua aberto.")
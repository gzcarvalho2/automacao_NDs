import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException

class FiscalBot:
    """
    Classe que encapsula toda a lógica para a automação do portal Fiscal.
    """
    def __init__(self, driver):
        """
        O construtor da classe. Recebe a instância do driver do Selenium.
        """
        self.driver = driver
        self.wait = WebDriverWait(self.driver, 20)
        self.dados_processados = []
        # Agrupa todos os seletores em um único lugar para fácil manutenção
        self.seletores = {
            "linhas_dados": "//div[contains(@class, 'flora--c-gqwkJN-ihfFBCg-css') and .//input[starts-with(@data-testid, 'select-debit-note-')]]",
            "botao_pdf": ".//button[text()='PDF']",
            "celulas_texto": ".//span[contains(@class, 'flora--c-LLdDZ')]",
            "proxima_pagina": "//button[@aria-current='true']/following-sibling::button[1]"
        }

    def navegar_para_aba_correta(self, titulo_alvo):
        """
        Encontra e muda para a aba do navegador com o título especificado.
        """
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            if titulo_alvo.lower() in self.driver.title.lower():
                print(f"Aba correta encontrada e selecionada: '{self.driver.title}'")
                return True
        print(f"\nAVISO: Nenhuma aba com o título '{titulo_alvo}' foi encontrada.")
        return False

    def _processar_pagina_atual(self):
        """
        Método 'privado' que contém a lógica para processar todas as linhas da página visível.
        """
        # Pega a contagem inicial de linhas
        total_linhas = len(self.driver.find_elements(By.XPATH, self.seletores["linhas_dados"]))
        print(f"Encontradas {total_linhas} linhas de dados para processar.")
        
        # Loop por índice, a abordagem mais segura
        for i in range(total_linhas):
            try:
                elementos_linha = self.driver.find_elements(By.XPATH, self.seletores["linhas_dados"])
                linha_atual_element = elementos_linha[i]
                
                celulas_texto = [span.text for span in linha_atual_element.find_elements(By.XPATH, self.seletores["celulas_texto"])]
                
                if not celulas_texto or len(celulas_texto) < 6:
                    print(f"  -> Linha {i+1} ignorada (sem dados ou cabeçalho).")
                    continue

                print(f"Processando linha {i+1}/{total_linhas} | Loja: {celulas_texto[0]}")
                
                status_arquivo = "Erro ao clicar"
                try:
                    botao_pdf = linha_atual_element.find_element(By.XPATH, self.seletores["botao_pdf"])
                    
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", botao_pdf)
                    time.sleep(0.5)
                    
                    actions = ActionChains(self.driver)
                    actions.move_to_element(botao_pdf).double_click().perform()
                    
                    print(f"  -> Botão PDF clicado (duplo). Aguardando 2 segundos...")
                    time.sleep(2)
                    status_arquivo = "Download iniciado"

                except Exception as click_error:
                    print(f"  -> AVISO: Falha ao clicar no botão da linha {i+1}: {click_error}")
                
                dados_para_df = celulas_texto + [status_arquivo]
                self.dados_processados.append(dados_para_df)
            
            except Exception as loop_error:
                print(f"  -> ERRO inesperado no loop na linha {i+1}: {loop_error}")

    def _ir_para_proxima_pagina(self):
        """
        Método 'privado' que clica no botão de próxima página e espera o recarregamento.
        """
        print("\nProcessamento da página concluído. Verificando próxima página...")
        try:
            elemento_referencia = self.driver.find_element(By.XPATH, self.seletores["linhas_dados"])
            proxima_pagina_btn = self.driver.find_element(By.XPATH, self.seletores["proxima_pagina"])

            if "..." in proxima_pagina_btn.text:
                print("Fim das páginas sequenciais. Encerrando.")
                return False

            print("Próxima página encontrada. Clicando...")
            proxima_pagina_btn.click()
            
            self.wait.until(EC.staleness_of(elemento_referencia))
            return True

        except NoSuchElementException:
            print("Não há mais botões de próxima página. Automação concluída.")
            return False

    def gerar_relatorio_final(self):
        """
        Cria e exibe o DataFrame do Pandas com todos os dados coletados.
        """
        if self.dados_processados:
            print(f"\n--- RELATÓRIO FINAL ---")
            print(f"Sucesso! {len(self.dados_processados)} linhas foram processadas no total.")
            colunas = ['LOJA', 'REFERÊNCIA SAP', 'NÚMERO DUPLIC.', 'DATA DE EMISSÃO', 'VENCIMENTO', 'VALOR', 'ARQUIVO']
            df = pd.DataFrame(self.dados_processados)
            df.columns = colunas[:len(df.columns)]
            print(df)

    def executar(self):
        """
        Método principal que orquestra toda a automação.
        """
        if not self.navegar_para_aba_correta(TITULO_DA_PAGINA_ALVO):
            return

        pagina_atual = 1
        while True:
            print(f"\n--- PROCESSANDO PÁGINA {pagina_atual} ---")
            
            try:
                self.wait.until(EC.presence_of_element_located((By.XPATH, self.seletores["linhas_dados"])))
            except TimeoutException:
                print("Tempo esgotado. Nenhuma linha de dados encontrada nesta página. Encerrando.")
                break

            self._processar_pagina_atual()
            
            if not self._ir_para_proxima_pagina():
                break
            
            pagina_atual += 1
        
        self.gerar_relatorio_final()

# --- BLOCO PRINCIPAL DE EXECUÇÃO ---
if __name__ == "__main__":
    TITULO_DA_PAGINA_ALVO = "Fiscal | Extranet"

    print("Iniciando automação...")
    try:
        edge_options = Options()
        edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        driver = webdriver.Edge(options=edge_options)
        
        # Cria a instância do robô
        bot = FiscalBot(driver)
        # Executa a automação
        bot.executar()

    except Exception as e:
        print(f"\n--- ERRO NO PROCESSAMENTO GERAL ---")
        print("Não foi possível iniciar o robô. Verifique se o navegador está aberto e foi iniciado com o comando correto.")
        print(f"Ocorreu um erro inesperado: {e}")
    finally:
        print("\nScript finalizado. O navegador continua aberto.")
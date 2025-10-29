# -*- coding: utf-8 -*-

"""
AUTOMAÇÃO FISCAL EXTRANET
===========================================

**Descrição:**
Este script automatiza o portal "Fiscal | Extranet".
O bot navega por um site que requer login (feito manualmente via modo debug), 
percorre tabelas de dados paginadas, baixa um arquivo PDF para cada linha e, 
em seguida, organiza esses arquivos em pastas com base em seu conteúdo.

As pastas de destino são criadas dinamicamente com a data do dia da execução
para manter um histórico organizado (ex: "Notas_de_Debito_Organizadas_2025-10-24").

Arquivos que não se encaixam em nenhuma regra são movidos para uma
pasta "arquivos gerais" dentro do diretório do dia.

A lógica de classificação foi melhorada para ignorar espaços no texto do PDF
(ex: "E C A D" será lido como "ECAD"), garantindo a classificação correta.

**Tecnologias e Habilidades Demonstradas:**
- **Automação Web com Selenium:**
  - Conexão a uma sessão de navegador já aberta (remote debugging).
  - Navegação entre abas e manipulação de elementos da página.
  - Uso de esperas explícitas (WebDriverWait) para lidar com conteúdo dinâmico.
  - Paginação automática.
- **Processamento de Arquivos PDF:**
  - Extração de texto de documentos PDF usando a biblioteca PyMuPDF (fitz).
- **Manipulação de Arquivos e Diretórios:**
  - Criação de pastas, renomeação e movimentação de arquivos com 'os' e 'shutil'.
  - Monitoramento de uma pasta de downloads para identificar novos arquivos.
- **Boas Práticas de Código:**
  - Código orientado a objetos (OOP).
  - Funções claras e bem documentadas.
  - Normalização de texto (remoção de acentos) para comparações robustas.
  - Geração dinâmica de nomes de pastas com data.

"""

import time
import os
import shutil
import fitz  # PyMuPDF
import pandas as pd
import unicodedata
from datetime import date  # Importado para usar a data atual
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException

class FiscalBot:
    """
    Classe completa para automação do portal Fiscal, incluindo download, análise e organização de PDFs.
    """
    def __init__(self, driver, pasta_download, pasta_destino, regras_classificacao):
        self.driver = driver
        self.wait = WebDriverWait(self.driver, 20)
        self.dados_processados = []
        
        self.seletores = {
            "linhas_dados": "//div[contains(@class, 'flora--c-gqwkJN-ihfFBCg-css') and .//input[starts-with(@data-testid, 'select-debit-note-')]]",
            "botao_pdf": ".//button[text()='PDF']",
            "celulas_texto": ".//span[contains(@class, 'flora--c-LLdDZ')]",
            "proxima_pagina": "//button[@aria-current='true']/following-sibling::button[1]"
        }
        
        self.pasta_download_temp = pasta_download
        self.pasta_destino_final = pasta_destino
        self.regras = regras_classificacao

    def _normalizar_texto(self, texto):
        """Remove acentos, caracteres especiais de uma string e a converte para maiúsculas."""
        if not texto: return ""
        texto = texto.upper()
        nfkd_form = unicodedata.normalize('NFKD', texto)
        return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

    def navegar_para_aba_correta(self, titulo_alvo):
        """Encontra e muda para a aba do navegador com o título especificado."""
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            if titulo_alvo.lower() in self.driver.title.lower():
                print(f"Aba correta encontrada e selecionada: '{self.driver.title}'")
                return True
        print(f"\nAVISO: Nenhuma aba com o título '{titulo_alvo}' foi encontrada.")
        return False

    def _processar_pagina_atual(self):
        """Contém a lógica para processar todas as linhas da página visível."""
        total_linhas = len(self.driver.find_elements(By.XPATH, self.seletores["linhas_dados"]))
        print(f"Encontradas {total_linhas} linhas de dados para processar.")
        
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
                    
                    print(f"  -> Botão PDF clicado (duplo). Aguardando download concluir...")
                    status_arquivo = "Download iniciado"

                    self._organizar_ultimo_arquivo_baixado()

                except Exception as click_error:
                    print(f"  -> AVISO: Falha ao clicar no botão da linha {i+1}: {click_error}")
                
                dados_para_df = celulas_texto + [status_arquivo]
                self.dados_processados.append(dados_para_df)
            
            except Exception as loop_error:
                print(f"  -> ERRO inesperado no loop na linha {i+1}: {loop_error}")

    def _organizar_ultimo_arquivo_baixado(self):
        """Orquestra o processo de análise e organização do arquivo mais recente."""
        print("    -> Iniciando rotina de organização de arquivo...")
        
        arquivo_recente = self._esperar_e_encontrar_novo_download()
        if not arquivo_recente:
            print("    -> AVISO: Download não concluído ou arquivo não encontrado no tempo limite.")
            return

        texto_pdf = self._extrair_texto_do_pdf(arquivo_recente)
        if not texto_pdf:
            print(f"    -> AVISO: Não foi possível ler o conteúdo do arquivo {os.path.basename(arquivo_recente)}. Movendo para destino.")
            # Move para "arquivos gerais" mesmo se não puder ler
            self._mover_para_arquivos_gerais(arquivo_recente)
            return

        self._classificar_renomear_e_mover(arquivo_recente, texto_pdf)

    def _esperar_e_encontrar_novo_download(self, timeout=30):
        """Espera ativamente por um novo arquivo .pdf e verifica se o download terminou."""
        segundos_passados = 0
        arquivos_antes = set(os.listdir(self.pasta_download_temp))
        while segundos_passados < timeout:
            time.sleep(1)
            segundos_passados += 1
            arquivos_depois = set(os.listdir(self.pasta_download_temp))
            novos_arquivos = arquivos_depois - arquivos_antes
            
            for arquivo in novos_arquivos:
                if arquivo.endswith(".pdf") and not arquivo.endswith(".crdownload"):
                    caminho_completo = os.path.join(self.pasta_download_temp, arquivo)
                    try:
                        tamanho_inicial = os.path.getsize(caminho_completo)
                        time.sleep(1.5)
                        tamanho_final = os.path.getsize(caminho_completo)
                        if tamanho_inicial == tamanho_final and tamanho_final > 0:
                            print(f"    -> Novo arquivo '{arquivo}' encontrado e download concluído.")
                            return caminho_completo
                    except (OSError, FileNotFoundError):
                        continue
        return None

    def _extrair_texto_do_pdf(self, caminho_arquivo):
        """Extrai o texto de um arquivo PDF."""
        try:
            with fitz.open(caminho_arquivo) as doc:
                texto_completo = ""
                for page in doc:
                    texto_completo += page.get_text()
                print(f"    -> Texto extraído de {os.path.basename(caminho_arquivo)}.")
                return texto_completo
        except Exception as e:
            print(f"      -> Erro ao ler o PDF: {e}")
            return ""
            
    def _mover_para_arquivos_gerais(self, caminho_arquivo):
        """Função auxiliar para mover arquivos para a pasta 'arquivos gerais'."""
        try:
            # Cria o caminho para a pasta "arquivos gerais" dentro da pasta de destino final (que já tem a data)
            pasta_arquivos_gerais = os.path.join(self.pasta_destino_final, "arquivos gerais")
            
            # Garante que essa pasta exista
            os.makedirs(pasta_arquivos_gerais, exist_ok=True)
            
            # Move o arquivo para lá
            shutil.move(caminho_arquivo, pasta_arquivos_gerais)
            print(f"    -> Arquivo movido para: {pasta_arquivos_gerais}")
        except Exception as e:
            print(f"      -> ERRO ao mover arquivo para 'arquivos gerais': {e}")

    def _classificar_renomear_e_mover(self, caminho_arquivo, texto_pdf):
        """Compara o texto normalizado com as regras e move o arquivo."""
        texto_pdf_normalizado = self._normalizar_texto(texto_pdf)
        
        # --- INÍCIO DA CORREÇÃO (ECAD BUG) ---
        # Cria uma versão do texto do PDF sem espaços para uma correspondência mais robusta
        # Isso corrige casos como "E C A D" ou "Mídia Regional"
        texto_pdf_sem_espacos = texto_pdf_normalizado.replace(" ", "")
        # --- FIM DA CORREÇÃO (ECAD BUG) ---

        print("    -> Verificando regras de classificação...")
        
        categoria_final_para_nome = None
        caminho_relativo_pasta = None

        for categoria, regra in self.regras.items():
            if isinstance(regra, dict):
                sub_categoria_encontrada = None
                # PRIMEIRO, procura pelas subcategorias que são mais específicas
                for sub_cat, sub_palavra_chave in regra["subcategorias"].items():
                    # Remove espaços da regra para a comparação
                    sub_chave_normalizada = self._normalizar_texto(sub_palavra_chave).replace(" ", "")
                    # print(f"      - Testando sub-regra '{sub_cat}'. Procurando por: '{sub_chave_normalizada}'...")
                    
                    # Compara a regra (sem espaços) com o PDF (sem espaços)
                    if sub_chave_normalizada in texto_pdf_sem_espacos: 
                        print(f"      -> Subcategoria encontrada: '{sub_cat}'")
                        sub_categoria_encontrada = sub_cat
                        break
                
                if sub_categoria_encontrada:
                    caminho_relativo_pasta = os.path.join(categoria, sub_categoria_encontrada)
                    categoria_final_para_nome = sub_categoria_encontrada
                    break
                
                # SE NÃO ACHOU subcategoria, procura pelo gatilho geral
                # Remove espaços da regra para a comparação
                gatilho_normalizado = self._normalizar_texto(regra.get("gatilho", "")).replace(" ", "")
                # print(f"      - Testando gatilho '{categoria}'. Procurando por: '{gatilho_normalizado}'...")
                
                if gatilho_normalizado and gatilho_normalizado in texto_pdf_sem_espacos: # Compara com PDF sem espaços
                    print(f"      -> GATILHO ENCONTRADO para '{categoria}'.")
                    caminho_relativo_pasta = categoria
                    categoria_final_para_nome = categoria
                    break
            
            elif isinstance(regra, str):
                # Remove espaços da regra para a comparação
                regra_normalizada = self._normalizar_texto(regra).replace(" ", "")
                # print(f"      - Testando regra simples '{categoria}'. Procurando por: '{regra_normalizada}'...")
                
                if regra_normalizada and regra_normalizada in texto_pdf_sem_espacos: # Compara com PDF sem espaços
                    print(f"      -> REGRA ENCONTRADA para '{categoria}'")
                    caminho_relativo_pasta = categoria
                    categoria_final_para_nome = categoria
                    break

        if caminho_relativo_pasta:
            try:
                pasta_destino_final_abs = os.path.join(self.pasta_destino_final, caminho_relativo_pasta)
                os.makedirs(pasta_destino_final_abs, exist_ok=True)
                nome_base, extensao = os.path.splitext(os.path.basename(caminho_arquivo))
                novo_nome = f"{nome_base}_{categoria_final_para_nome}{extensao}"
                caminho_final_arquivo = os.path.join(pasta_destino_final_abs, novo_nome)
                shutil.move(caminho_arquivo, caminho_final_arquivo)
                print(f"    -> Arquivo classificado como '{categoria_final_para_nome}' e movido para: {pasta_destino_final_abs}")
            except Exception as e:
                print(f"      -> ERRO ao mover/renomear arquivo: {e}")
        else:
            # --- LÓGICA 'arquivos gerais' ---
            print("    -> Nenhuma regra correspondeu. Movendo para a pasta 'arquivos gerais'.")
            self._mover_para_arquivos_gerais(caminho_arquivo)
            # --- FIM DA LÓGICA ---

    def _ir_para_proxima_pagina(self):
        """Clica no botão de próxima página e espera o recarregamento."""
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
        """Cria e exibe o DataFrame do Pandas com todos os dados coletados."""
        if self.dados_processados:
            print(f"\n--- RELATÓRIO FINAL ---")
            print(f"Sucesso! {len(self.dados_processados)} linhas foram processadas no total.")
            colunas = ['LOJA', 'REFERÊNCIA SAP', 'NÚMERO DUPLIC.', 'DATA DE EMISSÃO', 'VENCIMENTO', 'VALOR', 'ARQUIVO']
            df = pd.DataFrame(self.dados_processados)
            df.columns = colunas[:len(df.columns)]
            print(df.to_string()) # .to_string() garante que todo o df seja impresso

    def executar(self):
        """Método principal que orquestra toda a automação."""
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
    
    # ========================== ÁREA DE CONFIGURAÇÃO ==========================
    
    # 1. Defina a pasta TEMPORÁRIA para onde o navegador vai baixar os arquivos.
    PASTA_DOWNLOAD_TEMP = "C:\\Users\\gabri\\Downloads\\Notas_Temporarias"

    # 2. Defina a pasta FINAL onde os arquivos serão organizados.
    # --- MODIFICAÇÃO PARA ADICIONAR DATA ---
    
    # Define o nome base da pasta
    pasta_base_destino = "C:\\Users\\gabri\\Downloads\\Notas_de_Debito_Organizadas"
    # Pega a data de hoje e formata (ex: 2025-10-24)
    hoje_str = date.today().strftime("%d-%m-%Y")
    # Junta o nome base com a data de hoje
    PASTA_DESTINO_FINAL = f"{pasta_base_destino}_{hoje_str}"
    
    # --- FIM DA MODIFICAÇÃO ---

    # 3. Defina suas regras de classificação.
    #    Não se preocupe com acentos ou maiúsculas/minúsculas.
    REGRAS_DE_CLASSIFICACAO = {
        "marketing_institucional": "despesas de propaganda e esforços de marketing",
        "seguro": "Seguro",
        "outras_despesas_administrativas": "ECAD",
        "telecom": "Remuneração Esforços Tech",
        "MKT-REG": {
            "gatilho": "Mídia Regional",
            "subcategorias": {
                "MKT-REG_1": "Gestão Franqueador",
                "MKT-REG_5": "Gestão Individual",
                "MKT-REG_9": "REEMB ESF BOTIEXPERT"
            }
        }
    }
    # ========================================================================
    
    print("Iniciando automação...")
    print(f"Pasta de destino final será: {PASTA_DESTINO_FINAL}")
    
    try:
        os.makedirs(PASTA_DOWNLOAD_TEMP, exist_ok=True)
        os.makedirs(PASTA_DESTINO_FINAL, exist_ok=True) # Cria a pasta com data

        edge_options = Options()
        prefs = {"download.default_directory": PASTA_DOWNLOAD_TEMP}
        edge_options.add_experimental_option("prefs", prefs)
        edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        
        driver = webdriver.Edge(options=edge_options)
        
        bot = FiscalBot(driver, PASTA_DOWNLOAD_TEMP, PASTA_DESTINO_FINAL, REGRAS_DE_CLASSIFICACAO)
        
        bot.executar()

    except Exception as e:
        print(f"\n--- ERRO NO PROCESSAMENTO GERAL ---")
        print(f"Ocorreu um erro inesperado: {e}")
        print("Dicas:")
        print("- Verifique se o navegador está aberto no modo de depuração (debug).")
        print("- Confirme se os caminhos das pastas estão corretos.")
        print("- Verifique se os seletores (selectors) no código correspondem ao site.")
    finally:
        print("\nScript finalizado. O navegador continua aberto.")
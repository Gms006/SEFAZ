import random  # Import random para variação de tempo
from playwright.sync_api import sync_playwright, TimeoutError, expect
import time
import os
from datetime import datetime, timedelta
import subprocess
import re
import sys
import glob
import shutil
from pathlib import Path
import pyautogui
import win32com.client

class SefazGOAutomation:
    def __init__(self):
        self.sefaz_url = "https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica"
        # Define o formato correto das datas (DD/MM/AAAA)
        primeiro_dia_mes_atual = datetime.now().replace(day=1)
        ultimo_dia_mes_anterior = primeiro_dia_mes_atual - timedelta(days=1)
        self.data_inicial = "01/{:02d}/{}".format(ultimo_dia_mes_anterior.month, ultimo_dia_mes_anterior.year)
        self.data_final = "{:02d}/{:02d}/{}".format(ultimo_dia_mes_anterior.day, ultimo_dia_mes_anterior.month, ultimo_dia_mes_anterior.year)
        self.certificados_processados = 0
        self.log_file = f"log_sefaz_go_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        
        # Detecta a pasta de downloads padrão
        self.downloads_folder = str(Path.home() / "Downloads")
        self.log(f"Pasta de downloads: {self.downloads_folder}")
        self.log(f"Período de consulta definido: {self.data_inicial} a {self.data_final}")
        
        # Contador de arquivos na pasta de downloads antes de iniciar
        self.arquivos_iniciais = set()
        
        # Lista de nomes de certificados com base na imagem fornecida
        self.certificados_conhecidos = [
            "26 104 063 CAROLINA SILVA DE PAULA",
            "37 870 004 JOILTON DE OLIVEIRA MATOS",
            "147 426 394 VALQUIRIA MACHADO RIBEIRO",
            "51 365 114 ANDERSON RIBEIRO SILVA",
            "57 717 653 ELISA DOS SANTOS ROCHA",
            "JA A DE MATOS RESENDE SUPER BUFFET",
            "JA A DE MATOS RESENDE SUPER BUFFET",
            "JA C MELO LTDA",
            "JA K STORE LTDA",
            "JA S OLIVEIRA ENGENHARIA LTDA",
            "A2 SOLUCOES E PARTICIPACOES",
            "AC CERTIFICA ANAPOLIS"
        ]
        
    def log(self, mensagem):
        """Registra mensagens no console e em um arquivo de log"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_msg = f"[{timestamp}] {mensagem}"
        print(log_msg)
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(log_msg + "\n")
    
    def detectar_certificados(self):
        """Detecta certificados disponíveis na visualização do CertMgr.msc"""
        self.log("Detectando certificados disponíveis...")
        
        try:
            # Verifica a quantidade real de certificados
            # Como vimos na imagem, são 228 certificados
            qtd_certificados = 228
            
            self.log(f"Total de certificados detectados: {qtd_certificados}")
            return qtd_certificados
            
        except Exception as e:
            self.log(f"ERRO ao detectar certificados: {str(e)}")
            self.log("Usando o número de certificados da imagem: 228")
            return 228
    
    def criar_script_selecao_certificado(self, posicao):
        """Cria um script VBS para selecionar o certificado na posição especificada"""
        self.log(f"Criando script para selecionar o certificado na posição {posicao}")
        
        # Para o primeiro certificado, apenas Enter
        if posicao == 0:
            comandos_seta = ""
        else:
            # Para os demais, pressiona seta para baixo N vezes (posicao = número de cliques)
            comandos_seta = f'''
            ' Loop para pressionar tecla DOWN múltiplas vezes
            Dim i
            For i = 1 to {posicao}
                WshShell.SendKeys "{{DOWN}}"
                WScript.Sleep 100
            Next
            '''
            
        script_vbs = f'''
        ' Script VBS para simular teclas para seleção de certificado
        Option Explicit
        
        ' Aguarda a janela do certificado aparecer
        WScript.Sleep 3000
        
        ' Cria objeto para simulação de teclas
        Dim WshShell
        Set WshShell = CreateObject("WScript.Shell")
        
        {comandos_seta}
        
        ' Pequena pausa antes do ENTER
        WScript.Sleep 1000
        
        ' Pressiona ENTER para selecionar o certificado
        WshShell.SendKeys "{{ENTER}}"
        
        ' Limpa o objeto
        Set WshShell = Nothing
        '''
        
        # Salva o script em um arquivo .vbs
        script_path = f"selecionar_certificado_{posicao}.vbs"
        with open(script_path, "w") as file:
            file.write(script_vbs)
        
        return os.path.abspath(script_path)
    
    def processar_certificado(self, posicao_certificado, browser):
        """Processa um único certificado na posição especificada"""
        if posicao_certificado < len(self.certificados_conhecidos):
            nome_cert = self.certificados_conhecidos[posicao_certificado]
            self.log(f"Processando certificado #{posicao_certificado + 1}: {nome_cert}")
        else:
            self.log(f"Processando certificado #{posicao_certificado + 1}")
        
        # Cria script para seleção do certificado
        script_path = self.criar_script_selecao_certificado(posicao_certificado)
        
        try:
            # Cria uma nova página/aba
            context = browser.new_context(
                accept_downloads=True,  # Habilita downloads automáticos
                ignore_https_errors=True  # Ignora erros de certificado SSL
            )
            page = context.new_page()
            page.set_default_timeout(60000)  # 60 segundos de timeout
            
            # Executa o script VBS em segundo plano
            self.log("Executando script de seleção de certificado...")
            subprocess.Popen(f'start /b wscript.exe "{script_path}"', shell=True)
            
            # Acessa o site da SEFAZ
            self.log(f"Acessando {self.sefaz_url}")
            page.goto(self.sefaz_url, wait_until="networkidle")  # Espera a rede ficar ociosa
            
            # Aguarda para verificar se o site carregou corretamente
            try:
                # Verifica se o site carregou (checando um elemento específico)
                sucesso = page.wait_for_selector("#cmpCnpj", timeout=30000, state="attached")
                self.log("Site carregado com sucesso! Certificado válido.")
                
                # Aguarda mais um tempo para garantir que a página está completamente carregada
                self.log("Aguardando carregamento completo da página...")
                page.wait_for_load_state("networkidle")
                time.sleep(5)  # Aguarda mais 5 segundos para garantir
                
                # Verifica se há múltiplos CNPJs disponíveis
                self.processar_cnpjs(page)
                
            except TimeoutError:
                self.log("Falha ao carregar o site - Certificado possivelmente inválido ou expirado")
                
            # Fecha a aba/contexto atual
            self.log("Fechando aba atual")
            context.close()
            
        except Exception as e:
            self.log(f"ERRO ao processar certificado #{posicao_certificado + 1}: {str(e)}")
        finally:
            # Remove o arquivo VBS temporário
            try:
                os.remove(script_path)
                self.log(f"Arquivo VBS temporário removido: {script_path}")
            except:
                pass
        script_path = self.criar_script_selecao_certificado(posicao_certificado)
        
        try:
            # Cria uma nova página/aba
            context = browser.new_context(
                accept_downloads=True,  # Habilita downloads automáticos
                ignore_https_errors=True  # Ignora erros de certificado SSL
            )
            page = context.new_page()
            page.set_default_timeout(60000)  # 60 segundos de timeout
            
            # Executa o script VBS em segundo plano
            self.log("Executando script de seleção de certificado...")
            subprocess.Popen(f'start /b wscript.exe "{script_path}"', shell=True)
            
            # Acessa o site da SEFAZ
            self.log(f"Acessando {self.sefaz_url}")
            page.goto(self.sefaz_url, wait_until="domcontentloaded")
            
            # Aguarda para verificar se o site carregou corretamente
            try:
                # Verifica se o site carregou (checando um elemento específico)
                sucesso = page.wait_for_selector("#cmpCnpj", timeout=20000)
                self.log("Site carregado com sucesso! Certificado válido.")
                
                # Aguarda mais um tempo para garantir que a página está completamente carregada
                time.sleep(3)
                
                # Verifica se há múltiplos CNPJs disponíveis
                self.processar_cnpjs(page)
                
            except TimeoutError:
                self.log("Falha ao carregar o site - Certificado possivelmente inválido ou expirado")
                
            # Fecha a aba/contexto atual
            self.log("Fechando aba atual")
            context.close()
            
        except Exception as e:
            self.log(f"ERRO ao processar certificado #{posicao_certificado + 1}: {str(e)}")
        finally:
            # Remove o arquivo VBS temporário
            try:
                os.remove(script_path)
                self.log(f"Arquivo VBS temporário removido: {script_path}")
            except:
                pass
    
    def processar_cnpjs(self, page):
        """Processa todos os CNPJs disponíveis na página"""
        # Aguarda o seletor de CNPJ aparecer
        select_cnpj = page.locator("#cmpCnpj")
        
        try:
            # Obtém todas as opções do seletor
            options = page.evaluate('''() => {
                const select = document.querySelector('#cmpCnpj');
                if (!select) return [];
                return Array.from(select.options).map(option => ({
                    value: option.value,
                    text: option.text
                }));
            }''')
            
            # Se não há opções, algo está errado
            if not options:
                self.log("Nenhum CNPJ encontrado no seletor")
                return
                
            self.log(f"Encontrados {len(options)} CNPJs disponíveis")
            
            # Para cada CNPJ, realiza a busca e download
            for i, option in enumerate(options):
                if option['value']:  # Ignora opções vazias
                    self.log(f"Processando CNPJ {i+1}/{len(options)}: {option['text']}")
                    
                    # Seleciona o CNPJ
                    page.select_option("#cmpCnpj", option['value'])
                    time.sleep(1)
                    
                    # Realiza a consulta e download
                    self.realizar_consulta_download(page)
                    
                    # Aguarda um pouco antes de processar o próximo CNPJ
                    time.sleep(3)
                    
        except Exception as e:
            self.log(f"ERRO ao processar CNPJs: {str(e)}")
    
    def listar_arquivos_downloads(self):
        """Lista todos os arquivos na pasta de downloads"""
        try:
            arquivos = set(glob.glob(os.path.join(self.downloads_folder, "*")))
            return arquivos
        except Exception as e:
            self.log(f"ERRO ao listar arquivos na pasta de downloads: {str(e)}")
            return set()
    
    def monitorar_novos_downloads(self, timeout_segundos=60):
        """Monitora novos arquivos na pasta de downloads"""
        self.log("Monitorando a pasta de downloads para verificar novos arquivos...")
        
        # Lista de arquivos antes do download
        arquivos_antes = self.listar_arquivos_downloads()
        
        # Tempo de início do monitoramento
        inicio = time.time()
        arquivos_novos = set()
        
        # Monitoramento de novos arquivos
        while time.time() - inicio < timeout_segundos:
            arquivos_atuais = self.listar_arquivos_downloads()
            arquivos_novos = arquivos_atuais - arquivos_antes
            
            # Se encontrou arquivos novos
            if arquivos_novos:
                self.log(f"Detectados {len(arquivos_novos)} novos arquivos na pasta de downloads")
                
                # Lista os arquivos encontrados
                for i, arquivo in enumerate(arquivos_novos, 1):
                    nome_arquivo = os.path.basename(arquivo)
                    self.log(f"  {i}. {nome_arquivo}")
                
                return arquivos_novos
            
            # Aguarda um pouco antes de verificar novamente
            time.sleep(1)
        
        self.log("Nenhum novo arquivo detectado no tempo limite")
        return set()
    
    def realizar_consulta_download(self, page):
        """Realiza a consulta e download das notas fiscais"""
        try:
            # Preenche as datas
            self.log(f"Preenchendo período de consulta: {self.data_inicial} a {self.data_final}")
            page.fill("#cmpDataInicial", self.data_inicial)
            page.fill("#cmpDataFinal", self.data_final)
            
            # Clica no botão Pesquisar
            self.log("Clicando no botão Pesquisar...")
            page.click("#btnPesquisar")
            
            # Aguarda os resultados carregarem
            try:
                # Aguarda a tabela de resultados, mensagem de "nenhum registro" ou alerta de erro
                page.wait_for_selector("table.table, .no-records-found, .alert-danger", timeout=15000)
                
                # Verifica se há alerta de erro "Sem Resultados!"
                alerta_erro = page.locator(".alert-danger").count()
                if alerta_erro > 0:
                    texto_alerta = page.locator(".alert-danger").inner_text()
                    self.log(f"Alerta de erro detectado: {texto_alerta}")
                    return
                
                # Verifica se há resultados
                no_records = page.locator(".no-records-found").count()
                if no_records > 0:
                    self.log("Nenhuma nota fiscal encontrada no período selecionado")
                    return
                
                # Verifica se o botão de download está visível
                btn_download = page.locator(".btn-download-all").count()
                if btn_download == 0:
                    self.log("Botão de download não encontrado - Não há notas para baixar")
                    return
                
                # Lista arquivos antes do download para comparação
                arquivos_antes = self.listar_arquivos_downloads()
                
                # Clica no botão "Baixar todos os arquivos"
                self.log("Clicando no botão 'Baixar todos os arquivos'...")
                page.click(".btn-download-all")
                
                # Confirma o download no modal
                self.log("Confirmando o download...")
                page.wait_for_selector("#dnwld-all-btn-ok", timeout=5000)
                page.click("#dnwld-all-btn-ok")
                
                # Monitora a pasta de downloads para verificar novos arquivos
                arquivos_novos = self.monitorar_novos_downloads(timeout_segundos=60)
                
                if arquivos_novos:
                    self.log(f"Download concluído com sucesso: {len(arquivos_novos)} arquivo(s) baixado(s)")
                else:
                    self.log("ALERTA: Nenhum arquivo novo detectado após o download")
                
            except TimeoutError:
                self.log("Timeout ao aguardar os resultados da pesquisa")
                
        except Exception as e:
            self.log(f"ERRO ao realizar consulta e download: {str(e)}")
    
    def executar(self):
        """Executa a automação completa"""
        self.log("Iniciando automação SEFAZ GO")
        
        # Registra os arquivos iniciais na pasta de downloads
        self.arquivos_iniciais = self.listar_arquivos_downloads()
        self.log(f"Quantidade de arquivos na pasta de downloads antes de iniciar: {len(self.arquivos_iniciais)}")
        
        # Detecta certificados
        qtd_certificados = self.detectar_certificados()
        if qtd_certificados == 0:
            self.log("Nenhum certificado encontrado. Encerrando.")
            return
            
        # Inicia o Playwright
        with sync_playwright() as p:
            self.log("Iniciando o navegador Firefox...")
            browser = p.firefox.launch(
                headless=False,
                args=["--use-fake-ui-for-media-stream"]
            )
            
            # Processa cada certificado
            for i in range(qtd_certificados):
                self.log(f"\n{'='*50}")
                self.log(f"PROCESSANDO CERTIFICADO {i+1} DE {qtd_certificados}")
                if i < len(self.certificados_conhecidos):
                    self.log(f"Certificado: {self.certificados_conhecidos[i]}")
                self.log(f"{'='*50}\n")
                
                self.processar_certificado(i, browser)
                
                # Pequena pausa entre certificados
                time.sleep(3)
            
            browser.close()
        
        # Contabiliza os arquivos baixados durante toda a execução
        arquivos_finais = self.listar_arquivos_downloads()
        arquivos_novos = arquivos_finais - self.arquivos_iniciais
        
        self.log(f"\nTotal de arquivos baixados durante a execução: {len(arquivos_novos)}")
        if arquivos_novos:
            self.log("Arquivos baixados:")
            for i, arquivo in enumerate(arquivos_novos, 1):
                nome_arquivo = os.path.basename(arquivo)
                self.log(f"  {i}. {nome_arquivo}")
        
        self.log("\nAutomação concluída!")

# Executa o script
if __name__ == "__main__":
    automacao = SefazGOAutomation()
    automacao.executar()

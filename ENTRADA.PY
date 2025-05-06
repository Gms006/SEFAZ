from playwright.sync_api import sync_playwright, expect
import time
import os
import pandas as pd
from datetime import datetime
import subprocess

# Configurações
CERTIFICADO_ALVO = "MARCO ANTONIO CARVALHO NETO:81843909120"
URL_SEFAZ = "https://portal.sefaz.go.gov.br/portalsefaz-apps/auth/login-form"
SENHA_SEFAZ = "A2N5OPS8CG"
NUMERO_SETAS = 195  # Ajuste conforme a posição do certificado
PLANILHA_IE = r"C:\Notas de Entrada\Inscrições Estaduais\Notas_de_Entradas.xlsx"
DATA_INICIAL = "01/01/2024"  # Ajuste conforme necessário
DATA_FINAL = "31/01/2024"    # Ajuste conforme necessário
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")

# Função para criar script VBS que simula as setas e ENTER
def criar_script_vbs():
    script_vbs = f'''
    ' Script VBS para simular setas e ENTER
    Option Explicit
    
    ' Aguarda a janela do certificado aparecer
    WScript.Sleep 3000
    
    ' Cria objeto para simulação de teclas
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")
    
    ' Loop para pressionar tecla DOWN {NUMERO_SETAS} vezes
    Dim i
    For i = 1 to {NUMERO_SETAS}
        WshShell.SendKeys "{{DOWN}}"
        WScript.Sleep 50
    Next
    
    ' Pequena pausa antes do ENTER
    WScript.Sleep 1000
    
    ' Pressiona ENTER
    WshShell.SendKeys "{{ENTER}}"
    
    ' Limpa o objeto
    Set WshShell = Nothing
    '''
    
    # Salva o script em um arquivo .vbs
    with open("simular_teclas.vbs", "w") as file:
        file.write(script_vbs)
    
    return os.path.abspath("simular_teclas.vbs")

# Função para esperar download
def esperar_download(pasta, timeout=120):
    start_time = time.time()
    while time.time() - start_time < timeout:
        arquivos = [f for f in os.listdir(pasta) if f.endswith('.zip') and not f.endswith('.part')]
        if arquivos:
            # Verifica se o arquivo está completo (não está sendo baixado)
            time.sleep(2)  # Pequena pausa para garantir que o download terminou
            return True
        time.sleep(1)
    return False

# Função principal
def baixar_notas_fiscais():
    print("Iniciando processo de download de notas fiscais...")
    
    # Carrega as inscrições estaduais da planilha
    try:
        df = pd.read_excel(PLANILHA_IE)
        print(f"Planilha carregada com sucesso. Encontradas {len(df)} empresas.")
    except Exception as e:
        print(f"Erro ao carregar planilha: {str(e)}")
        return
    
    # Cria o script VBS
    script_path = criar_script_vbs()
    print(f"Script VBS criado em: {script_path}")
    
    with sync_playwright() as p:
        # Inicia o navegador
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            accept_downloads=True,  # Aceita downloads automaticamente
        )
        page = context.new_page()
        page.set_default_navigation_timeout(120000)  # 2 minutos
        
        # ===== ETAPA 1: LOGIN COM CERTIFICADO DIGITAL =====
        
        # Executa o script VBS em segundo plano
        os.system(f'start /b wscript.exe "{script_path}"')
        print(f"Script para simular {NUMERO_SETAS} teclas para baixo + ENTER iniciado.")
        
        # Acessa a página da SEFAZ-GO
        print("Acessando a SEFAZ-GO...")
        try:
            page.goto(URL_SEFAZ, wait_until="domcontentloaded")
            print("Página inicial acessada!")
        except Exception as e:
            print(f"Erro ao acessar página inicial (esperado se certificado foi selecionado): {str(e)}")
        
        # Aguarda a página carregar após selecionar o certificado
        print("Aguardando carregamento após seleção de certificado...")
        time.sleep(10)
        
        # ===== ETAPA 2: NAVEGAR ATÉ A PÁGINA DE DOWNLOAD =====
        
        # Clica no botão Autenticar com Certificado Digital
        print("Clicando em 'Autenticar com Certificado Digital'...")
        try:
            page.wait_for_selector("button#btnAuthenticate", timeout=30000)
            page.click("button#btnAuthenticate")
            page.wait_for_load_state("domcontentloaded")
            print("Autenticação com certificado iniciada.")
        except Exception as e:
            print(f"Erro ao clicar em autenticar certificado: {str(e)}")
        
        # Aguarda a página carregar
        time.sleep(5)
        
        # IMPORTANTE: Registra quantas abas temos antes de clicar em Acesso Restrito
        num_abas_antes = len(context.pages)
        print(f"Número de abas antes de clicar em Acesso Restrito: {num_abas_antes}")
        
        # Clica no item Acesso Restrito - configurado para abrir em nova aba
        print("Selecionando 'Acesso Restrito'...")
        try:
            # Aguarda o elemento ficar visível
            page.wait_for_selector("a[title='Acessar o Sistema']:has-text('Acesso Restrito')", timeout=30000)
            
            # Adiciona um event listener para detectar novas abas
            with context.expect_page() as new_page_info:
                # Clica no elemento que abrirá uma nova aba
                page.click("a[title='Acessar o Sistema']:has-text('Acesso Restrito')")
                print("Clique em 'Acesso Restrito' realizado.")
            
            # Captura a nova página que foi aberta
            nova_aba = new_page_info.value
            print("Nova aba capturada através do event listener!")
            
            # Muda o foco para a nova aba
            nova_aba.wait_for_load_state("domcontentloaded")
            page = nova_aba
            print(f"Foco alterado para a nova aba. URL: {page.url}")
            
        except Exception as e:
            print(f"Erro no método principal de captura da nova aba: {str(e)}")
            
            # Método alternativo: verifica manualmente se uma nova aba foi aberta
            print("Tentando método alternativo para capturar a nova aba...")
            time.sleep(5)  # Aguarda um pouco para a nova aba abrir
            
            if len(context.pages) > num_abas_antes:
                # Uma nova aba foi aberta
                page = context.pages[-1]  # Muda para a última aba
                print(f"Método alternativo: Nova aba encontrada! URL: {page.url}")
            else:
                print("ALERTA: Nenhuma nova aba foi detectada!")
        
        # Verifica todas as abas abertas (para diagnóstico)
        all_pages = context.pages
        print(f"Total de abas abertas agora: {len(all_pages)}")
        for i, pg in enumerate(all_pages):
            try:
                print(f"Aba {i+1}: Título = '{pg.title()}'")
                print(f"Aba {i+1}: URL = '{pg.url}'")
            except:
                print(f"Aba {i+1}: Não foi possível obter título/URL")
        
        # Aguarda a página carregar completamente
        page.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(5)
        
        # Captura screenshot para diagnóstico
        screenshot_path = "pagina_acesso_restrito.png"
        page.screenshot(path=screenshot_path)
        print(f"Screenshot salva em: {os.path.abspath(screenshot_path)}")
        
        # ===== ETAPA 3: USAR JAVASCRIPT PARA CLICAR EM BAIXAR XML NFE =====
        print("Buscando e clicando no link 'Baixar XML NFE'...")
        
        # ATENÇÃO: O link "Baixar XML NFE" na página de Acesso Restrito
        baixar_xml_executado = False
        
        try:
            # Vamos tentar executar o JavaScript para abrir o link
            resultado = page.evaluate('''
            () => {
                try {
                    // Procura por links com onclick contendo OpenUrl
                    const links = Array.from(document.querySelectorAll('a[onclick*="OpenUrl"]'));
                    for (const link of links) {
                        console.log("Link encontrado:", link.outerHTML);
                        
                        // Verifica se o texto contém "Baixar XML" ou similar
                        if (link.innerText.includes("Baixar XML") || 
                            link.innerText.includes("XML NFE") || 
                            link.innerText.includes("NFE")) {
                            console.log("Link com texto relevante encontrado!");
                            link.click();
                            return {sucesso: true, método: "link encontrado por texto", html: link.outerHTML};
                        }
                        
                        // Verifica onclick para ver se contém URL relevante
                        const onclick = link.getAttribute("onclick");
                        if (onclick && (onclick.includes("consulta-notas") || 
                                        onclick.includes("xml"))) {
                            console.log("Link com onclick relevante encontrado!");
                            link.click();
                            return {sucesso: true, método: "link encontrado por onclick", html: link.outerHTML};
                        }
                    }
                    
                    // Se não encontrou links específicos, procura pela função OpenUrl
                    if (typeof OpenUrl === 'function') {
                        console.log("Função OpenUrl encontrada, tentando executar diretamente");
                        OpenUrl('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-notas-recebidas', false, '', 'False', 'true');
                        return {sucesso: true, método: "função OpenUrl"};
                    }
                    
                    return {sucesso: false, erro: "Link não encontrado e função OpenUrl não disponível"};
                } catch (e) {
                    return {sucesso: false, erro: e.toString()};
                }
            }
            ''')
            
            print(f"Resultado da execução JavaScript: {resultado}")
            
            if resultado.get('sucesso', False):
                baixar_xml_executado = True
                print("Execução JavaScript bem-sucedida para o link Baixar XML NFE")
            else:
                print("Falha na execução JavaScript para o link Baixar XML NFE")
            
            # Aguarda a navegação acontecer
            time.sleep(5)
            
        except Exception as e:
            print(f"Erro ao executar JavaScript para clicar no link: {str(e)}")
        
        # Se não conseguiu executar o JavaScript, tenta navegar diretamente
        if not baixar_xml_executado:
            try:
                print("Tentando navegar diretamente para a URL de consulta...")
                page.goto("https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-notas-recebidas", timeout=60000)
                print("Navegação direta para página de consulta realizada.")
                baixar_xml_executado = True
            except Exception as e:
                print(f"Erro na navegação direta: {str(e)}")
        
        # ===== ETAPA 4: LIDAR COM A PÁGINA DE SENHA =====
        
        # Captura screenshot da página atual
        page.screenshot(path="pagina_apos_link_xml.png")
        print("Screenshot após link XML salva.")
        
        # Verifica se estamos na página de senha
        print("Verificando se estamos na página de senha...")
        
        # Várias tentativas para encontrar o campo de senha
        senha_element = None
        try:
            senha_element = page.query_selector("input[name='NetAccess.Password']")
            if not senha_element:
                senha_element = page.query_selector("#NetAccess\\.Password")
            if not senha_element:
                senha_element = page.query_selector("input[type='password']")
            
            if senha_element:
                print("Campo de senha encontrado!")
                
                # Aguarda um pouco para garantir que a página carregou completamente
                time.sleep(2)
                
                # Tenta clicar no campo de senha para garantir o foco
                print("Clicando no campo de senha...")
                senha_element.click()
                time.sleep(1)
                
                # Preenche a senha lentamente
                print(f"Preenchendo a senha: {SENHA_SEFAZ}")
                senha_element.fill("")  # Limpa o campo primeiro
                time.sleep(0.5)
                senha_element.type(SENHA_SEFAZ, delay=100)  # Digite lentamente
                
                # Screenshot após preencher a senha
                page.screenshot(path="senha_preenchida.png")
                print("Screenshot após preencher senha salva.")
                
                # Clica no botão Autenticar
                print("Clicando no botão Autenticar...")
                
                # Tenta encontrar o botão de autenticação
                botao_autenticar = page.query_selector("button[id='btnAuthenticate']")
                if not botao_autenticar:
                    botao_autenticar = page.query_selector("button[onclick*='NetAccess.internetFormSessionLogon']")
                if not botao_autenticar:
                    botao_autenticar = page.query_selector("button[type='submit']")
                
                if botao_autenticar:
                    print("Botão de autenticação encontrado. Clicando...")
                    botao_autenticar.click()
                    print("Clique no botão Autenticar realizado.")
                else:
                    print("Botão Autenticar não encontrado pelos seletores. Tentando via JavaScript...")
                    # Tenta usar JavaScript para clicar no botão
                    resultado = page.evaluate('''
                    () => {
                        try {
                            // Procura pelo botão de autenticação
                            const botao = document.querySelector("button[id='btnAuthenticate']") || 
                                         document.querySelector("button[onclick*='NetAccess.internetFormSessionLogon']") || 
                                         document.querySelector("button[type='submit']") ||
                                         document.querySelector("button");
                            
                            if (botao) {
                                console.log("Botão encontrado via JS:", botao.outerHTML);
                                botao.click();
                                return {sucesso: true, html: botao.outerHTML};
                            }
                            
                            // Se não encontrou o botão, tenta executar a função diretamente
                            if (typeof NetAccess !== 'undefined' && 
                                typeof NetAccess.internetFormSessionLogon === 'function') {
                                NetAccess.internetFormSessionLogon();
                                return {sucesso: true, método: "função NetAccess.internetFormSessionLogon"};
                            }
                            
                            return {sucesso: false, erro: "Botão não encontrado"};
                        } catch (e) {
                            return {sucesso: false, erro: e.toString()};
                        }
                    }
                    ''')
                    
                    print(f"Resultado da tentativa de clicar no botão via JavaScript: {resultado}")
                
                # Aguarda a página processar a autenticação
                print("Aguardando processamento da autenticação...")
                page.wait_for_load_state("networkidle", timeout=30000)
                time.sleep(5)
                
                # Screenshot após autenticação
                page.screenshot(path="apos_autenticacao.png")
                print("Screenshot após autenticação salva.")
                
                # ===== ETAPA 5: CLICAR NOVAMENTE EM BAIXAR XML NFE =====
                
                print("Tentando clicar novamente em 'Baixar XML NFE' após autenticação...")
                
                # Este segundo clique em "Baixar XML NFE" é crucial
                segundo_clique_sucesso = False
                
                try:
                    # Vamos tentar executar o JavaScript novamente para o segundo clique
                    resultado = page.evaluate('''
                    () => {
                        try {
                            // Procura por links com onclick contendo OpenUrl
                            const links = Array.from(document.querySelectorAll('a[onclick*="OpenUrl"]'));
                            for (const link of links) {
                                console.log("Link encontrado para segundo clique:", link.outerHTML);
                                
                                // Verifica se o texto contém "Baixar XML" ou similar
                                if (link.innerText.includes("Baixar XML") || 
                                    link.innerText.includes("XML NFE") || 
                                    link.innerText.includes("NFE")) {
                                    console.log("Link com texto relevante encontrado para segundo clique!");
                                    link.click();
                                    return {sucesso: true, método: "link encontrado por texto", html: link.outerHTML};
                                }
                                
                                // Verifica onclick para ver se contém URL relevante
                                const onclick = link.getAttribute("onclick");
                                if (onclick && (onclick.includes("consulta-notas") || 
                                                onclick.includes("xml"))) {
                                    console.log("Link com onclick relevante encontrado para segundo clique!");
                                    link.click();
                                    return {sucesso: true, método: "link encontrado por onclick", html: link.outerHTML};
                                }
                            }
                            
                            // Se não encontrou links específicos, procura pela função OpenUrl
                            if (typeof OpenUrl === 'function') {
                                console.log("Função OpenUrl encontrada para segundo clique, tentando executar diretamente");
                                OpenUrl('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-notas-recebidas', false, '', 'False', 'true');
                                return {sucesso: true, método: "função OpenUrl"};
                            }
                            
                            return {sucesso: false, erro: "Link não encontrado e função OpenUrl não disponível"};
                        } catch (e) {
                            return {sucesso: false, erro: e.toString()};
                        }
                    }
                    ''')
                    
                    print(f"Resultado do segundo clique: {resultado}")
                    
                    if resultado.get('sucesso', False):
                        segundo_clique_sucesso = True
                        print("Segundo clique em Baixar XML NFE realizado com sucesso")
                    else:
                        print("Falha no segundo clique em Baixar XML NFE")
                    
                    # Aguarda a navegação acontecer
                    page.wait_for_load_state("networkidle", timeout=30000)
                    time.sleep(5)
                    
                except Exception as e:
                    print(f"Erro ao executar segundo clique: {str(e)}")
                
                # Se o segundo clique falhou, tenta navegar diretamente
                if not segundo_clique_sucesso:
                    try:
                        print("Tentando navegar diretamente para a URL após autenticação...")
                        page.goto("https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-notas-recebidas", timeout=60000)
                        print("Navegação direta após autenticação realizada.")
                    except Exception as e:
                        print(f"Erro na navegação direta após autenticação: {str(e)}")
            else:
                print("Campo de senha não encontrado!")
                
                # Se não encontrou o campo de senha, pode ser que já estejamos na página após autenticação
                # Tenta clicar no link Baixar XML NFE diretamente
                try:
                    print("Tentando clicar diretamente no link (assumindo que já estamos autenticados)...")
                    resultado = page.evaluate('''
                    () => {
                        try {
                            // Procura por links com onclick contendo OpenUrl
                            const links = Array.from(document.querySelectorAll('a[onclick*="OpenUrl"]'));
                            for (const link of links) {
                                console.log("Link encontrado:", link.outerHTML);
                                
                                // Verifica se o texto contém "Baixar XML" ou similar
                                if (link.innerText.includes("Baixar XML") || 
                                    link.innerText.includes("XML NFE") || 
                                    link.innerText.includes("NFE")) {
                                    link.click();
                                    return {sucesso: true, método: "link encontrado por texto", html: link.outerHTML};
                                }
                                
                                // Verifica onclick para ver se contém URL relevante
                                const onclick = link.getAttribute("onclick");
                                if (onclick && (onclick.includes("consulta-notas") || 
                                               onclick.includes("xml"))) {
                                    link.click();
                                    return {sucesso: true, método: "link encontrado por onclick", html: link.outerHTML};
                                }
                            }
                            
                            // Se não encontrou links específicos, procura pela função OpenUrl
                            if (typeof OpenUrl === 'function') {
                                OpenUrl('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-notas-recebidas', false, '', 'False', 'true');
                                return {sucesso: true, método: "função OpenUrl"};
                            }
                            
                            return {sucesso: false, erro: "Link não encontrado e função OpenUrl não disponível"};
                        } catch (e) {
                            return {sucesso: false, erro: e.toString()};
                        }
                    }
                    ''')
                    
                    print(f"Resultado do clique direto: {resultado}")
                    
                    # Aguarda a navegação acontecer
                    page.wait_for_load_state("networkidle", timeout=30000)
                    time.sleep(5)
                    
                except Exception as e:
                    print(f"Erro ao tentar clique direto: {str(e)}")
                    
                    # Tenta navegar diretamente como último recurso
                    try:
                        print("Navegando diretamente para a URL como último recurso...")
                        page.goto("https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-notas-recebidas", timeout=60000)
                        print("Navegação direta como último recurso realizada.")
                    except Exception as e2:
                        print(f"Erro na navegação direta como último recurso: {str(e2)}")
                
        except Exception as e:
            print(f"Erro ao verificar página de senha: {str(e)}")
            
            # Tenta recuperação navegando diretamente
            try:
                print("Navegando diretamente para a URL de consulta como recuperação...")
                page.goto("https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-notas-recebidas", timeout=60000)
                print("Navegação direta de recuperação realizada.")
            except Exception as e2:
                print(f"Erro na navegação direta de recuperação: {str(e2)}")
        
        # Aguarda a página carregar
        page.wait_for_load_state("networkidle", timeout=60000)
        time.sleep(5)
        
        # ===== ETAPA 6: VERIFICAR SE CHEGAMOS NA PÁGINA DE BUSCA =====
        
        # Captura screenshot da página atual
        page.screenshot(path="pagina_final.png")
        print("Screenshot da página final salva.")
        
        # Verifica se estamos na página correta de busca
        data_inicial_element = page.query_selector("input#cmpDataInicial")
        if not data_inicial_element:
            print("ERRO: Não conseguimos chegar à página de busca de notas!")
            print("Encerramento prematuro do script.")
            browser.close()
            return
        
        print("SUCESSO: Página de busca de notas acessada!")
        
        # ===== ETAPA 7: LOOP PARA BAIXAR NOTAS DE CADA EMPRESA =====
        
        for index, row in df.iterrows():
            empresa = str(row[0])  # Nome da empresa (primeira coluna)
            inscricao = str(row[1])  # Inscrição estadual (segunda coluna)
            
            print(f"\n==== Processando empresa: {empresa} - IE: {inscricao} ====")
            
            # Preenche a data inicial
            print(f"Preenchendo data inicial: {DATA_INICIAL}")
            try:
                page.fill("input#cmpDataInicial", DATA_INICIAL)
            except Exception as e:
                print(f"Erro ao preencher data inicial: {str(e)}")
                continue
            
            # Preenche a data final
            print(f"Preenchendo data final: {DATA_FINAL}")
            try:
                page.fill("input#cmpDataFinal", DATA_FINAL)
            except Exception as e:
                print(f"Erro ao preencher data final: {str(e)}")
                continue
            
            # Preenche a inscrição estadual
            print(f"Preenchendo inscrição estadual: {inscricao}")
            try:
                page.fill("input#cmpNumIeDest", inscricao)
            except Exception as e:
                print(f"Erro ao preencher inscrição estadual: {str(e)}")
                continue
            
            # Seleciona o tipo de nota (entrada)
            print("Selecionando tipo de nota: ENTRADA")
            try:
                page.check("input#cmpTipoNota[value='0']")
            except Exception as e:
                print(f"Erro ao selecionar tipo de nota: {str(e)}")
            
            # Seleciona "Exibir notas canceladas"
            print("Selecionando exibir notas canceladas...")
            try:
                page.check("input#cmpExbNotasCanceladas")
            except Exception as e:
                print(f"Erro ao selecionar exibir notas canceladas: {str(e)}")
            
            # Clica em Pesquisar
            print("Clicando em Pesquisar...")
            try:
                page.click("button#btnPesquisar")
                page.wait_for_load_state("networkidle")
                print("Pesquisa realizada.")
            except Exception as e:
                print(f"Erro ao clicar em Pesquisar: {str(e)}")
                continue
            
            # Verifica se há resultados
            print("Verificando resultados...")
            time.sleep(3)
            
            sem_resultados = page.query_selector("div.alert-danger:has-text('Sem Resultados!')")
            if sem_resultados:
                print(f"Sem resultados para a empresa {empresa} no período especificado.")
                continue
            
            # Clica em "Baixar todos os arquivos"
            print("Clicando em 'Baixar todos os arquivos'...")
            try:
                page.click("button.btn-download-all")
                time.sleep(2)
                print("Modal de download aberto.")
            except Exception as e:
                print(f"Erro ao clicar em Baixar todos os arquivos: {str(e)}")
                continue
            
            # Clica em "Baixar" no modal
            print("Confirmando download...")
            try:
                page.click("button#dnwld-all-btn-ok")
                print("Download iniciado.")
            except Exception as e:
                print(f"Erro ao confirmar download: {str(e)}")
                continue
            
            # Aguarda o download
            print("Aguardando download completar...")
            if esperar_download(PASTA_DOWNLOADS):
                print(f"Download das notas da empresa {empresa} concluído com sucesso!")
            else:
                print(f"Timeout ao aguardar download para a empresa {empresa}.")
            
            # Pequena pausa entre empresas
            time.sleep(5)
        
        # ===== FINALIZAÇÃO =====
        print("\nProcesso de download de notas fiscais concluído!")
        
        # Remove o arquivo VBS temporário
        try:
            os.remove("simular_teclas.vbs")
            print("Arquivo VBS temporário removido.")
        except:
            pass
        
        # Fecha o navegador
        browser.close()

if __name__ == "__main__":
    print("=== AUTOMAÇÃO DE DOWNLOAD DE NOTAS FISCAIS DE ENTRADA ===")
    print(f"Data de processamento: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"Certificado alvo: {CERTIFICADO_ALVO}")
    print(f"Período: {DATA_INICIAL} a {DATA_FINAL}")
    
    baixar_notas_fiscais()
    
    print("\nProcesso finalizado. Pressione ENTER para encerrar...")
    input()

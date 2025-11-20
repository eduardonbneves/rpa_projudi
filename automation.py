import os
import pandas as pd
from playwright.sync_api import sync_playwright, expect, TimeoutError
import time
import datetime

def ler_processos_excel(caminho_arquivo):
    """Lê uma lista de números de processo de um arquivo Excel."""
    try:
        df = pd.read_excel(caminho_arquivo)
        if 'PROCESSOS' not in df.columns:
            return None, "Coluna 'PROCESSOS' não encontrada no arquivo."
        
        processos = df['PROCESSOS'].dropna().astype(str).str.strip().tolist()
        return processos, None
    except Exception as e:
        return None, f"Erro ao ler arquivo: {e}"

def automacao_projudi(usuario, senha, oab, lista_processos, update_callback=None, check_stop=None, log_callback=None, lista_sucesso=None):
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    if not usuario or not senha:
        if update_callback:
            update_callback("Erro: Usuário ou senha não fornecidos.", 0)
        return

    total_processos = len(lista_processos)
    
    with sync_playwright() as p:
        browser = None
        try:
            browser = p.chromium.launch(headless=True, slow_mo=100)
            page = browser.new_page()

            if update_callback:
                update_callback("Iniciando login...", 0)

            # Login
            try:
                page.goto("https://projudi.tjam.jus.br/projudi/usuario/logon.do?actionType=inicio")
                page.locator("#login").fill(usuario)
                page.locator("#senha").fill(senha)
                page.locator("#btEntrar").click()
                expect(page.locator("#barraMenu")).to_be_visible(timeout=30000)
            except Exception as e:
                if update_callback:
                    update_callback(f"Erro no login: {e}", 0)
                return

            contador = 0
            for processo in lista_processos:
                if check_stop and check_stop():
                    if update_callback:
                        update_callback("Execução interrompida pelo usuário.", (contador / total_processos) * 100)
                    break

                contador += 1
                progresso = (contador / total_processos) * 100
                
                if update_callback:
                    update_callback(f"Processando {contador}/{total_processos}: {processo}", progresso)

                log(f"{contador} - Iniciando automação para o processo: {processo}")

                try:
                    frame_busca = page.frame_locator('[name="userMainFrame"]')

                    try:
                        page.locator("#Stm0p0i8e").hover()
                        page.locator("#Stm0p8i0eTX").click()
                    except Exception as e:
                        if update_callback:
                            update_callback(f"Erro na navegação do menu: {e}", 0)
                            return
                    
                    frame_busca.locator("#numeroProcesso").fill(processo)
                    frame_busca.locator("#pesquisar").click()

                    try:
                        frame_busca.get_by_text(processo, exact=True).click(timeout=5000)
                    except TimeoutError:
                        log(f"Processo {processo} não encontrado na busca.")
                        continue

                    try:
                        # Tenta clicar, mas espera no MÁXIMO 3 segundos (3000 ms)
                        frame_busca.locator("#habilitacaoButton").click(timeout=3000)

                        # ... Continua o código de sucesso ...
                        frame_busca.locator("#addButton").click()

                    except TimeoutError:
                        # Se passar 3 segundos e não achar o botão:
                        log(f"⚠️ Botão não apareceu (timeout) para o processo {processo}.")
                        continue  # Pula para o próximo

                    frame_incluir = frame_busca.frame_locator("iframe").first
                    frame_incluir.locator("#oab").fill(oab)
                    frame_incluir.locator("#searchButton").click()
                    frame_incluir.locator("#selectButton").click()
                    frame_busca.locator('input[type="checkbox"][value="0"]').check()
                    frame_busca.locator("#saveButton").click()

                    log(f"✅ Habilitação concluída: {processo}")

                    if lista_sucesso is not None:
                        lista_sucesso.append({
                            "Processo": processo,
                            "Status": "Sucesso",
                            "Horario": datetime.datetime.now().strftime("%H:%M:%S")
                        })

                except TimeoutError:
                    log(f"\n❌ ERRO DE TIMEOUT no processo {processo}")
                    nome_arquivo = f"erro_{processo.replace('.', '').replace('-', '')}.png"
                    page.screenshot(path=nome_arquivo)
                    log(f"   📸 Print do erro salvo como: {nome_arquivo}")

                except Exception as e:
                    log(f"\n❌ Erro inesperado no processo {processo}: {e}")
                
            if update_callback:
                update_callback("Finalizado!", 100)

        except Exception as e:
             if update_callback:
                update_callback(f"Erro geral: {e}", 0)
        finally:
            if browser:
                browser.close()

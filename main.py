from playwright.sync_api import sync_playwright, expect, TimeoutError
import time
import csv
import os
from dotenv import load_dotenv

def ler_processos(nome_arquivo="processos.txt"):
    """Lê uma lista de números de processo de um arquivo de texto."""
    try:
        with open(nome_arquivo, 'r') as f:
            # list comprehension para remover espaços em branco e linhas vazias
            processos = [linha.strip() for linha in f if linha.strip()]
        return processos
    except FileNotFoundError:
        print(f"Erro: Arquivo '{nome_arquivo}' não encontrado.")
        return []

def automacao_projudi(lista_processos):
    """Função principal que orquestra a automação."""
    resultados_finais = []

    load_dotenv()

    usuario = os.getenv("LOGIN_USUARIO")
    senha = os.getenv("LOGIN_SENHA")
    oab = os.getenv("OAB")

    if not usuario or not senha:
        print("Erro: As credenciais LOGIN_USUARIO e LOGIN_SENHA não foram encontradas no arquivo .env")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=100)
        page = browser.new_page()

        try:
            # --- LOGIN MANUAL ---
            # !!! Substitua pela URL correta do Projudi do seu estado !!!
            page.goto("https://projudi.tjam.jus.br/projudi/usuario/logon.do?actionType=inicio")
            print("="*50)
            print("Preenchendo credenciais...")
            print("="*50)

            page.locator("#login").fill(usuario)
            page.locator("#senha").fill(senha)

            # 3. CLICAR NO BOTÃO DE ENTRAR
            page.locator("#btEntrar").click()

            expect(page.locator("#barraMenu")).to_be_visible(timeout=30000)
            print("\n✅ LOGIN REALIZADO COM SUCESSO!\n")

            page.locator("#Stm0p0i8e").hover()
            page.locator("#Stm0p8i0eTX").click()

            #for processo in lista_processos:
            frame_busca = page.frame_locator('[name="userMainFrame"]')
            frame_busca.locator("#numeroProcesso").fill(lista_processos[0])
            frame_busca.locator("#pesquisar").click()

            frame_busca.get_by_text(lista_processos[0], exact=True).click()

            frame_busca.locator("#habilitacaoButton").click()

            frame_busca.locator("#addButton").click()


            frame_incluir = frame_busca.frame_locator("iframe").first

            frame_incluir.locator("#oab").fill(oab)

            frame_incluir.locator("#searchButton").click()

            #frame_incluir.locator('input[type="radio"][value="&oab={oab}&complemento=A&uf=AM&idTipoAdvogado=2"]').check()

            frame_incluir.locator("#selectButton").click()


            frame_busca.locator('input[type="checkbox"][value="0"]').check()

            frame_busca.locator("#saveButton").click()

            time.sleep(5)

        except TimeoutError as e:
            print("\n❌ ERRO DE TIMEOUT: Um elemento demorou demais para aparecer.")
            print("   - Pode haver um pop-up ou um 'loading spinner' bloqueando a tela.")
        except Exception as e:
            print(f"\n❌ Ocorreu um erro inesperado: {e}")
        finally:
            print("Finalizando e fechando o navegador.")
            browser.close()

        return resultados_finais

numeros_de_processo = ler_processos()
print(f"Processos a serem pesquisados: {numeros_de_processo}")
if numeros_de_processo:
    dados_coletados = automacao_projudi(numeros_de_processo)
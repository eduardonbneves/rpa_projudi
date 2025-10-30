from playwright.sync_api import sync_playwright, expect, TimeoutError
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

        contador = 1

        for processo in lista_processos:

            print(contador, " - Iniciando automação para o processo:", processo)

            try:
                page.goto("https://projudi.tjam.jus.br/projudi/usuario/logon.do?actionType=inicio")

                page.locator("#login").fill(usuario)
                page.locator("#senha").fill(senha)

                # 3. CLICAR NO BOTÃO DE ENTRAR
                page.locator("#btEntrar").click()

                expect(page.locator("#barraMenu")).to_be_visible(timeout=30000)

                page.locator("#Stm0p0i8e").hover()
                page.locator("#Stm0p8i0eTX").click()

                frame_busca = page.frame_locator('[name="userMainFrame"]')
                frame_busca.locator("#numeroProcesso").fill(processo)
                frame_busca.locator("#pesquisar").click()

                frame_busca.get_by_text(processo, exact=True).click()

                frame_busca.locator("#habilitacaoButton").click()

                frame_busca.locator("#addButton").click()


                frame_incluir = frame_busca.frame_locator("iframe").first

                frame_incluir.locator("#oab").fill(oab)

                frame_incluir.locator("#searchButton").click()

                #frame_incluir.locator('input[type="radio"][value="&oab={oab}&complemento=A&uf=AM&idTipoAdvogado=2"]').check()

                frame_incluir.locator("#selectButton").click()


                frame_busca.locator('input[type="checkbox"][value="0"]').check()

                frame_busca.locator("#saveButton").click()

            except TimeoutError as e:
                print("\n❌ ERRO DE TIMEOUT: Um elemento demorou demais para aparecer.")
                print("   - Pode haver um pop-up ou um 'loading spinner' bloqueando a tela.")
            except Exception as e:
                print(f"\n❌ Ocorreu um erro inesperado: {e}")
            finally:
                print(f"Finalizada automação para o processo: {processo}\n")

            contador += 1

        print("Finalizando e fechando o navegador.")
        browser.close()

numeros_de_processo = ler_processos()
if numeros_de_processo:
    dados_coletados = automacao_projudi(numeros_de_processo)
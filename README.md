# Automação PROJUDI AM

Este é um aplicativo de desktop RPA (Robotic Process Automation) desenvolvido para automatizar a habilitação em processos no sistema PROJUDI do Amazonas.

## Pré-requisitos

- Python 3.8 ou superior instalado.

## Instalação

1.  **Crie um ambiente virtual (recomendado):**

    ```bash
    python -m venv .venv
    ```

2.  **Ative o ambiente virtual:**

    -   No Windows (PowerShell):
        ```bash
        .venv\Scripts\Activate
        ```
    -   No Windows (CMD):
        ```cmd
        .venv\Scripts\activate.bat
        ```

3.  **Instale as dependências:**

    ```bash
    pip install -r requirements.txt
    ```

4.  **Instale os navegadores do Playwright:**

    ```bash
    playwright install
    ```

## Como Usar

1.  **Execute o aplicativo:**

    ```bash
    python main.py
    ```

2.  **Configure o Acesso:**
    -   Insira seu **Usuário** e **Senha** do PROJUDI.
    -   Insira o número da **OAB**.
    -   Clique em "Salvar Configurações" para memorizar seus dados para a próxima vez.

3.  **Selecione os Processos:**
    -   Clique em "Selecionar" no campo "Arquivo Excel".
    -   Escolha um arquivo `.xlsx` ou `.xls` que contenha uma coluna chamada **`PROCESSOS`** com os números dos processos a serem trabalhados.

4.  **Inicie a Automação:**
    -   Clique no botão **"Iniciar Habilitação"**.
    -   Acompanhe o progresso na barra de progresso e os detalhes na janela de logs abaixo.

## Funcionalidades

-   **Interface Gráfica Amigável:** Fácil de configurar e usar.
-   **Logs em Tempo Real:** Visualize o que o robô está fazendo.
-   **Tratamento de Erros:** O robô tenta lidar com falhas de carregamento e timeouts, salvando prints de erros quando necessário.
-   **Configurações Salvas:** Não precisa digitar seus dados toda vez.
### Automação de E-mails com Python e Outlook

Este projeto é um script em Python para automatizar o envio de e-mails personalizados para colaboradores. Ele utiliza dados de uma planilha Excel (`.xlsx`) e interage diretamente com a interface do Microsoft Outlook para otimizar a comunicação interna, permitindo o envio de convites ou comunicados em massa de forma individualizada e eficiente.

-----

### ⚙️ Funcionalidades

  * **Leitura de Dados**: Importa informações de colaboradores (nome, e-mail, etc.) de um arquivo Excel.
  * **Personalização**: Gera um corpo de e-mail único para cada destinatário, usando *placeholders* como `{nome}`.
  * **Integração com Outlook**: Utiliza a biblioteca `pywin32` para interagir diretamente com o Outlook, criando e enviando os e-mails.
  * **Controle de Envio**: Evita o envio de e-mails duplicados, marcando os colaboradores já processados na planilha original.
  * **Sistema de Log**: Mantém um registro detalhado de todas as operações, incluindo sucessos e falhas, em um arquivo de log.

-----

### 🚀 Tecnologias e Pré-requisitos

Para rodar este script, você precisa ter:

  * **Python 3.x**
  * **Microsoft Outlook** instalado e configurado em sua máquina.
  * As seguintes bibliotecas Python instaladas:
    ```
    pip install pandas
    pip install pywin32
    ```

-----

### 🛠️ Como Usar

1.  **Estruture a Planilha:**
    Crie uma planilha Excel (`.xlsx`) chamada `dados_colaboradores.xlsx` dentro de uma pasta chamada `dados_envio`. A planilha deve ter, no mínimo, as seguintes colunas: `nome` e `email`. Você pode adicionar outras colunas personalizadas se desejar.

2.  **Ajuste as Configurações:**
    Abra o script `aplicacao.py` e configure as variáveis na seção **"1. Configurações e Variáveis"** de acordo com as suas necessidades (nome do projeto, nome da empresa, URL, etc.).

3.  **Execute o Script:**
    Abra o terminal ou prompt de comando, navegue até a pasta do projeto e execute o script com o seguinte comando:

    ```
    python aplicacao.py


    # Nome do Projeto 🤖

![Status](https://img.shields.io/badge/Status-Concluído-success?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.10+-blue?style=for-the-badge&logo=python)

## 📝 Sobre
Uma breve descrição do problema que este projeto resolve.
*Exemplo: Automação desenvolvida para extrair dados do site X, processar com Pandas e salvar em um banco PostgreSQL diariamente.*

## 🛠️ Tecnologias Utilizadas
- **Linguagem:** Python
- **Bibliotecas:** Pandas, Selenium, SQLAlchemy
- **Banco de Dados:** PostgreSQL
- **Infra:** AWS Lambda / Docker

## ⚙️ Funcionalidades
- [x] Login automático no portal administrativo
- [x] Extração de relatórios em PDF
- [x] Tratamento de dados (remoção de duplicatas)
- [x] Envio de alerta via Email/Slack em caso de erro

## 🚀 Como Rodar o Projeto

### Pré-requisitos
* Python 3.10+
* Docker (Opcional)

### Instalação
```bash
# Clone o repositório
git clone [https://github.com/lucassantoss4/nome-do-projeto.git](https://github.com/lucassantoss4/nome-do-projeto.git)

# Entre na pasta
cd nome-do-projeto

# Instale as dependências
pip install -r requirements.txt

# Configure as variáveis de ambiente
cp .env.example .env
# (Edite o arquivo .env com suas credenciais)

# Execute
python src/main.py
    ```


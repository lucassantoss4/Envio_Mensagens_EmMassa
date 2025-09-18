# 📨 Automação de E-mails com Python e Outlook

Este projeto é um script em Python para automatizar o envio de e-mails personalizados para colaboradores, utilizando dados de uma planilha Excel (`.xlsx`) e a interface do Microsoft Outlook.

O objetivo é otimizar a comunicação interna, permitindo o envio de convites ou comunicados em massa de forma individualizada e eficiente.

---

## ⚙️ Funcionalidades

* **Leitura de Dados:** Importa informações de colaboradores (nome, e-mail, etc.) de um arquivo Excel.
* **Personalização:** Gera um corpo de e-mail único para cada destinatário, usando `placeholders` como `{nome}`.
* **Integração com Outlook:** Utiliza a biblioteca `pywin32` para interagir diretamente com o Outlook, criando e enviando os e-mails.
* **Controle de Envio:** Evita o envio de e-mails duplicados, marcando os colaboradores já processados na planilha original.
* **Sistema de Log:** Mantém um registro detalhado de todas as operações, incluindo sucessos e falhas, em um arquivo de log.

---

## 🚀 Como Usar

### Pré-requisitos
Certifique-se de que você tem o Python instalado em seu sistema e o Microsoft Outlook configurado.

1.  **Instalar as Bibliotecas:**
    Abra o terminal ou prompt de comando e execute:
    ```bash
    pip install pandas pywin32
    ```

2.  **Preparar a Planilha:**
    Crie um arquivo Excel chamado `dados_colaboradores.xlsx` no mesmo diretório do script, com as seguintes colunas:
    - `nome`
    - `email`
    - `enviado` (opcional; o script criará se não existir)

    Exemplo:
    | nome          | email                     | enviado |
    |---------------|---------------------------|---------|
    | Ana Silva     | ana.silva@empresa.com.br  |         |
    | Bruno Souza   | bruno.souza@empresa.com.br|         |

3.  **Executar o Script:**
    - Feche a planilha Excel.
    - Abra o Microsoft Outlook.
    - Execute o script Python no terminal:
    ```bash
    python [nome_do_seu_script].py
    ```

---

## 📝 Personalização do E-mail

O corpo do e-mail está configurado no script como uma **f-string** em formato HTML, o que permite formatação rica e inclusão dinâmica de dados.

Para adaptar o corpo do e-mail, edite a variável `corpo_email_html` no arquivo `[nome_do_seu_script].py`.

---

## 📄 Licença

Este projeto é de código aberto e está licenciado sob a [Nome da Licença - Ex: MIT License].

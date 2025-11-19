# Automação de Relatórios Diários (Excel + E-mail)

Este projeto é uma automação em Python desenvolvida para processar planilhas Excel, gerar relatórios em PDF/Imagem e enviar e-mails personalizados para parceiros/clientes.

## 🚀 Funcionalidades

- **Processamento de Excel:** Abre planilhas complexas (`.xlsm`), atualiza conexões de dados e executa macros VBA automaticamente.
- **Geração de PDF/Imagem:** Converte abas específicas do Excel em PDF e posteriormente em imagens (PNG) usando a biblioteca `pdf2image`.
- **Envio de E-mail:** Envia e-mails autenticados via SMTP (Outlook/Office365) com anexos e corpo em HTML personalizado.
- **Interface Gráfica:** Possui uma interface simples em `tkinter` para facilitar a execução das tarefas pelo usuário.

## 🛠️ Pré-requisitos

Para rodar este projeto, você precisará de:

1.  **Python 3.12+**
2.  **Poppler:** Ferramenta necessária para manipulação de PDFs.
    - Baixe a versão para Windows e adicione a pasta `bin` ao PATH do sistema.
3.  **Microsoft Excel:** Instalado na máquina (para automação via `win32com`).

## 📦 Instalação

1.  Clone o repositório:
    ```bash
    git clone [https://github.com/vmontesNT/relatorio-diario.git](https://github.com/vmontesNT/relatorio-diario.git)
    cd relatorio-diario
    ```

2.  Crie e ative um ambiente virtual:
    ```bash
    python -m venv venv
    # Windows:
    .\venv\Scripts\activate
    ```

3.  Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```

## ⚙️ Configuração (.env)

Este projeto utiliza variáveis de ambiente para segurança. Crie um arquivo `.env` na raiz do projeto e configure suas credenciais e caminhos:

```ini
# Credenciais de E-mail
EMAIL_REMETENTE=seu_email@dominio.com.br
USUARIO_SMTP=seu_usuario
SENHA_SMTP=sua_senha
SERVIDOR_SMTP=smtp.office365.com
PORTA_SMTP=587

# Caminhos Locais (Ajuste conforme sua máquina)
CAMINHO_PASTA_PARCEIROS=C:\Caminho\Para\Arquivos
CAMINHO_ARQUIVO_EXCEL=C:\Caminho\Para\Planilha.xlsm
CAMINHO_POPPLER=C:\Caminho\Para\poppler\Library\bin
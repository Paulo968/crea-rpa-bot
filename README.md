# CREA-BOT 🚀  
Automação real para problemas reais.

> **Bot RPA em Python + Selenium que automatiza o envio de contratos no portal do CREA-MG.  
> Resultado real: +400% de produtividade e padronização do processo em 7 filiais.**

---

## 📌 Principais Funcionalidades

- 🎨 **Interface moderna (CustomTkinter)** — tema claro/escuro  
- 📂 **Carrega automaticamente a última planilha usada**  
- 🔄 **Retomada inteligente** — continua do ponto exato onde parou  
- ➕ **Agrupa contratos duplicados**  
- 🔐 **Validação avançada** com indicação de célula com erro  
- 🧾 **Datas e valores editáveis pela planilha** (DATA_INICIO, DATA_FIM, VALOR_RECEITA)  
- 🚫 **Botão “Parar” seguro** — encerra o atual e retoma depois  
- 💾 **Backup automático** da planilha original  
- 🛠 **Empacotável em .exe** com PyInstaller  
- 🌙 **Modo escuro/claro**  
- 🧠 **Seleção automática da fazenda** (pausa para cadastro e retoma sozinho)  
- 📊 **Logs detalhados + barra de progresso**

---

## 🧠 Diferenciais Técnicos

- Arquitetura **modular profissional**
- Automação **Selenium WebDriver** com WebDriverManager
- Suporte a **retomada de execução**  
- Validação do Excel com apontamento direto de células  
- Mecanismo de agrupamento para evitar duplicidade  
- Persistência de dados via `config.json`  
- Protocolo seguro de interrupção (“Soft Stop”)  
- 100% compatível com empacotamento em `.exe`

---

## 🗂 Estrutura do Projeto

```
crea-bot/
├── automation/       # Selenium e lógica de automação
├── core/             # Leitura e validação das planilhas
├── interface/        # GUI moderna (CustomTkinter)
├── utils/            # Funções auxiliares, logs, backup
├── config.json       # Persistência de estado
├── main.py           # Inicializador principal
└── README.md
```

---

## ▶️ Como Rodar

### 1. Clone o repositório
```bash
git clone https://github.com/Paulo968/crea-bot.git
cd crea-bot
```

### 2. Crie o ambiente virtual
```bash
python -m venv venv
venv\Scripts\activate   # Windows
```

### 3. Instale as dependências
```bash
pip install -r requirements.txt
```

### 4. Inicie o bot
```bash
python main.py
```

---

## 📋 Requisitos da Planilha

Campos obrigatórios:

- `NUMERO DO CONTRATO`  
- `CPF_CNPJ`  
- `DATA DO REGISTRO`  
- `FAZENDA`  
- `CPF_LOGIN`  
- `SENHA_LOGIN`  
- `ARTCREA`  

Campos opcionais/editáveis:

- `DATA_INICIO`  
- `DATA_FIM`  
- `VALOR_RECEITA`  

---

## 💡 Evoluções Futuras

- Exportação automática em PDF  
- Histórico em banco de dados  
- Modo multiusuário  
- Atualização automática  
- Notificação por e-mail ao finalizar

---

## ⚠ Avisos Importantes

- Use a última versão do Google Chrome  
- Evite usar durante manutenção do sistema CREA  
- Sempre mantenha backup das planilhas  
- Edite datas e valores direto no Excel

---

## 👨‍💻 Autor

Feito por **Paulo Zaqueu de Oliveira Junior**  
_“Automação sem enrolação para o CREA-MG.”_

🔗 GitHub: https://github.com/Paulo968  
🔗 LinkedIn: https://www.linkedin.com/in/paulo-zaqueu-762459187  
📧 Email: paulozaqueu3@gmail.com

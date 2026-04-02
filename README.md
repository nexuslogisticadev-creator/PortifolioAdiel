# 🤖 BOT CONTROL — Enterprise Edition V8.0

Sistema de gerenciamento operacional para delivery, integrando automação do **Zé Delivery**, controle de motoboys, impressão térmica, WhatsApp e controle remoto via **Telegram**.

---

## 📁 Estrutura dos Componentes

| Arquivo | Função |
|---|---|
| `painel.py` | Interface gráfica (GUI) de controle — centro de comando |
| `robo.py` | Robô principal de automação — monitora pedidos e aciona impressoras |
| `telegram_bot.py` | Bot Telegram — controle remoto do sistema via smartphone |
| `utils.py` | Utilitários compartilhados entre os três sistemas |
| `config.json` | Arquivo central de configuração |
| `estoque.json` | Base de dados do estoque |
| `alertas_atraso.json` | Registro de alertas de atraso de pedidos |
| `memoria_fechamento.json` | Cache de horários do fechamento diário |
| `robo.log` | Log em tempo real do robô |

---

## 🖥️ painel.py — Painel de Controle

Interface gráfica desktop desenvolvida com **CustomTkinter** (tema dark). Funciona como o centro de comando do sistema.

### Abas disponíveis

| Aba | Descrição |
|---|---|
| 📊 **Dashboard (Monitor)** | Tabela de pedidos do dia, busca e filtros |
| 💰 **Fechamento** | Cálculo automático de fechamento por motoboy, exporta para Excel |
| 💸 **Vales & Descontos** | Controle de vales e alertas de atraso |
| 📦 **Estoque** | Gestão de produtos por categoria, alertas de estoque baixo, lista de compras |
| 🛵 **Equipe** | Cadastro de motoboys e taxas por bairro |
| 📍 **Zonas** | Mapeamento de bairros e valores de entrega |
| 🔑 **PIX** | Chaves PIX de cada motoboy |
| 💻 **Terminal** | Visualização do `robo.log` em tempo real |
| ⚙️ **Configurações** | Edição do `config.json` pelo painel |

### Como iniciar o painel

```bash
# Ativar ambiente virtual e rodar
.\venv\Scripts\python.exe painel.py
# ou usar o executável compilado
painel.exe
```

O painel detecta automaticamente se o robô está rodando e permite iniciar/parar com o botão **⚡ INICIAR SISTEMA**.

---

## 🤖 robo.py — Robô Principal

Processo que roda em segundo plano (headless ou com Chrome), responsável pela automação do Zé Delivery.

### O que o robô faz

- **Monitora pedidos** via API do Zé Delivery (GraphQL)
- **Distribui pedidos** para motoboys com base em bairros e disponibilidade
- **Imprime tickets** automáticos na impressora térmica (ESC/POS)
- **Imprime extratos de fechamento** por motoboy
- **Envia notificações** no grupo WhatsApp via Selenium + Chrome
- **Alerta atrasos** de retirada automaticamente
- **Gera planilhas Excel** de controle financeiro diário
- **Sincroniza** com o Google Sheets
- **Responde comandos** do painel e do Telegram via arquivos de texto

### Comunicação com outros sistemas

```
painel.py  ──→  comando_imprimir.txt  ──→  robo.py
telegram_bot.py ──→  telegram_command.txt  ──→  robo.py
robo.py  ──→  Telegram (notificações automáticas)
robo.py  ──→  robo.log (log em tempo real)
```

### Como iniciar o robô separado

```bash
.\venv\Scripts\python.exe robo.py
```

> ⚠️ O robô requer o **Google Chrome** instalado. O ChromeDriver é gerenciado automaticamente pelo `webdriver-manager`.

---

## 📱 telegram_bot.py — Bot Telegram

Bot autônomo que permite controlar o sistema remotamente pelo Telegram, sem precisar estar no computador.

### Comandos disponíveis

| Comando | Descrição |
|---|---|
| `/start` | Inicia o robô |
| `/stop` | Para o robô |
| `/restart` | Reinicia o robô |
| `/status` | Status atual (online/offline) |
| `/logs` | Últimas 25 linhas do `robo.log` |
| `/imprimir <Nome>` | Imprime extrato de fechamento de um motoboy |
| `/garantia <Nome> <Início> <Fim>` | Gera recibo de garantia |
| `/resumo` | Total de taxas e total do dia |
| `/canceladas` | Relatório de pedidos cancelados/perdidos |
| `/fechamento_manual` | Força geração do relatório |
| `/atualizar_estoque` | Atualiza estoque pelo histórico |
| `/motos` | Ver entregadores na rua |
| `/pendentes` | Pedidos na fila |
| `/estoque` | Itens com estoque baixo |
| `/ocultar` | Oculta janela do Chrome |
| `/mostrar` | Mostra janela do Chrome |
| `/excel` | Gera relatório no Google Sheets |
| `/consultar_vale [Motoboy]` | Lista vales do dia |
| `/lancar_vale` | Lança novo vale |
| `/excluir_vale <ID>` | Exclui um vale |
| `/alerta_auto` | Ativa/desativa alertas automáticos |
| `/mencao` | Ativa/desativa menção no WhatsApp |
| `/recarregar_config` | Recarrega `config.json` |
| `/help` ou `/menu` | Lista todos os comandos |

### Como iniciar o bot Telegram

```bash
.\venv\Scripts\python.exe telegram_bot.py
# ou via o .bat incluso
start_telegram_bot.bat
```

---

## ⚙️ Configuração (config.json)

Todas as configurações ficam em `config.json`. Os campos principais:

```json
{
  "grupo_whatsapp": "Nome do grupo no WhatsApp",
  "endereco_loja": "Rua..., Cidade",
  "email": "login@zedelivery.com.br",
  "senha": "SuaSenha",
  "telegram_token": "TOKEN_DO_BOT",
  "telegram_chat_id": "SEU_CHAT_ID",
  "path_backup": "caminho/para/backup",
  "motoboys": { "NomeMotoboy": "chave_api" },
  "bairros": { "NomeBairro": 8.00 },
  "pix_motoboys": { "NomeMotoboy": "chave_pix" },
  "google_sheets": { ... },
  "url_api": "https://api.zedelivery.com.br/graphql",
  "headers_api": { ... },
  "alerta_retirada_auto": false,
  "whatsapp_mencao_ativa": false
}
```

---

## 📦 Instalação das Dependências

```bash
# Criar e ativar o ambiente virtual (recomendado)
python -m venv venv
.\venv\Scripts\activate

# Instalar todas as dependências
pip install -r requirements.txt
```

### Bibliotecas utilizadas

| Biblioteca | Uso |
|---|---|
| `customtkinter` | Interface gráfica dark moderna |
| `tkinter` / `tkcalendar` | Widgets e calendário |
| `selenium` + `webdriver-manager` | Automação do Chrome/WhatsApp |
| `openpyxl` + `pandas` | Leitura e escrita de planilhas Excel |
| `gspread` + `google-auth` | Integração com Google Sheets |
| `python-telegram-bot` | Bot Telegram assíncrono |
| `pywin32` | Impressão térmica e controle de janelas (Windows) |
| `psutil` | Gerenciamento de processos |
| `requests` + `curl_cffi` | Requisições HTTP à API |
| `pdfplumber` | Leitura de PDFs |
| `matplotlib` + `folium` | Gráficos e mapas de calor |
| `flask` | Servidor HTTP interno |
| `pyperclip` | Copiar para área de transferência |
| `geocoder` | Geolocalização de bairros |
| `numpy` | Cálculos numéricos |

---

## 🚀 Scripts de Inicialização

| Arquivo | Função |
|---|---|
| `iniciar_bot.bat` | Inicia o robô principal |
| `start_telegram_bot.bat` | Inicia o bot do Telegram |
| `parar_tudo.bat` | Para todos os processos |
| `run_executor_foreground.bat` | Executa o executor de fila em primeiro plano |
| `install_as_service.bat` | Instala o robô como serviço Windows (via NSSM) |

---

## 🔗 Fluxo de Integração

```
┌──────────────────────────────────────────┐
│              ZDELIVERY API               │
│              (GraphQL)                   │
└───────────────────┬──────────────────────┘
                    │ pedidos
                    ▼
┌──────────────────────────────────────────┐
│              robo.py                     │
│  • Processa pedidos                      │
│  • Aciona WhatsApp (Selenium)            │
│  • Imprime tickets (ESC/POS)             │
│  • Gera Excel / Google Sheets            │
│  • Envia alertas Telegram                │
└────────┬────────────────────┬────────────┘
         │ robo.log           │ Arquivos txt
         │ Excel              │ de comando
         ▼                    ▼
┌──────────────┐    ┌─────────────────────┐
│  painel.py   │    │  telegram_bot.py    │
│  (GUI local) │    │  (controle remoto)  │
└──────────────┘    └─────────────────────┘
```

---

## 📋 Requisitos do Sistema

- **OS:** Windows 10/11 (64-bit)
- **Python:** 3.10 ou superior
- **Google Chrome:** Última versão instalada
- **Impressora:** Térmica com suporte ESC/POS (opcional)
- **RAM:** Mínimo 4GB recomendado (Chrome + robô)

---

## 📝 Logs e Monitoramento

- **`robo.log`** — Log completo do robô, visualizável na aba Terminal do painel ou via `/logs` no Telegram
- **`alertas_atraso.json`** — Registro de alertas de pedidos demorados
- O sistema usa **auto-refresh inteligente** de 10 segundos: só recarrega dados quando o arquivo Excel muda, economizando CPU.

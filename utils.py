# ==================================================================================
# MÓDULO UTILITÁRIO CENTRALIZADO
# ==================================================================================
# Funções, constantes e utilitários compartilhados entre:
# - robo.py (robô principal)
# - telegram_bot.py (bot Telegram)
# - painel.py (painel de controle)
# ==================================================================================
import os
import sys
import json
import time
from datetime import datetime, timedelta


def _obter_hora_virada_operacional(default=10):
    """Retorna a hora de virada operacional (0-23) a partir do config.json."""
    caminho_config = os.path.join(get_caminho_base(), ARQUIVO_CONFIG)
    try:
        with open(caminho_config, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
        hora = int(cfg.get("hora_virada_operacional", default))
        return max(0, min(23, hora))
    except Exception:
        return default

# ==================================================================================
# CAMINHOS E DIRETÓRIOS
# ==================================================================================

def get_caminho_base():
    """Retorna o caminho base do executável ou do script."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_caminho_excel(data_str=None):
    """Retorna o caminho para o arquivo Excel do dia."""
    if data_str is None:
        data_str = get_data_operacional()
    nome_arquivo = f"Controle_Financeiro_{data_str}.xlsx"
    return os.path.join(get_caminho_base(), nome_arquivo)

def get_log_path():
    """Retorna o caminho para o arquivo de log."""
    return os.path.join(get_caminho_base(), "robo.log")

# ==================================================================================
# CONFIGURAÇÃO
# ==================================================================================

ARQUIVO_CONFIG = 'config.json'
ARQUIVO_ROBO = 'robo.py'
ARQUIVO_COMANDO = 'comando_imprimir.txt'
ARQUIVO_COMANDO_TELEGRAM = 'telegram_command.txt'
ARQUIVO_ESTOQUE = 'estoque.json'
ARQUIVO_ESTOQUE_BAIXAS = 'estoque_baixas.json'
ARQUIVO_FECHAMENTO_STATUS = 'fechamento_status.json'
ARQUIVO_ALERTAS = 'alertas_atraso.json'
ARQUIVO_MEMORIA_FECH = 'memoria_fechamento.json'

def carregar_config():
    """Carrega configurações do config.json."""
    caminho_config = os.path.join(get_caminho_base(), ARQUIVO_CONFIG)
    if not os.path.exists(caminho_config):
        print(f"❌ Erro: {ARQUIVO_CONFIG} não encontrado")
        return None
    
    try:
        with open(caminho_config, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        print(f"❌ Erro ao ler {ARQUIVO_CONFIG}: {e}")
        return None

def salvar_config(config):
    """Salva configurações no config.json."""
    caminho_config = os.path.join(get_caminho_base(), ARQUIVO_CONFIG)
    try:
        with open(caminho_config, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        return True
    except IOError as e:
        print(f"❌ Erro ao salvar {ARQUIVO_CONFIG}: {e}")
        return False

# ==================================================================================
# DATA E HORA
# ==================================================================================

def get_data_operacional():
    """Retorna a data operacional (muda às 10h da manhã)."""
    agora = datetime.now()
    hora_virada = _obter_hora_virada_operacional(default=10)
    if agora.hour < hora_virada:
        agora -= timedelta(days=1)
    return agora.strftime("%d-%m-%Y")

def get_data_iso():
    """Retorna a data em formato ISO YYYY-MM-DD."""
    return datetime.now().strftime("%Y-%m-%d")

# ==================================================================================
# NORMALIZAÇÃO DE TEXTO
# ==================================================================================

import unicodedata

def normalizar_texto(texto):
    """Normaliza texto removendo acentos e caracteres de combinação."""
    if not texto:
        return ""
    try:
        texto_nfc = unicodedata.normalize('NFD', str(texto))
        return ''.join(c for c in texto_nfc if unicodedata.category(c) != 'Mn').lower()
    except Exception:
        return str(texto).lower()

# ==================================================================================
# PROCESSAMENTO DE ARQUIVOS
# ==================================================================================

def arquivo_existe(nome_arquivo):
    """Verifica se um arquivo existe no diretório base."""
    caminho = os.path.join(get_caminho_base(), nome_arquivo)
    return os.path.exists(caminho)

def ler_arquivo(nome_arquivo, encoding='utf-8'):
    """Lê arquivo de texto."""
    caminho = os.path.join(get_caminho_base(), nome_arquivo)
    try:
        with open(caminho, 'r', encoding=encoding) as f:
            return f.read().strip()
    except FileNotFoundError:
        return ""
    except Exception as e:
        print(f"❌ Erro ao ler {nome_arquivo}: {e}")
        return ""

def escrever_arquivo(nome_arquivo, conteudo, encoding='utf-8'):
    """Escreve conteúdo em arquivo."""
    caminho = os.path.join(get_caminho_base(), nome_arquivo)
    try:
        with open(caminho, 'w', encoding=encoding) as f:
            f.write(conteudo)
        return True
    except Exception as e:
        print(f"❌ Erro ao escrever {nome_arquivo}: {e}")
        return False

def deletar_arquivo(nome_arquivo):
    """Deleta um arquivo."""
    caminho = os.path.join(get_caminho_base(), nome_arquivo)
    try:
        if os.path.exists(caminho):
            os.remove(caminho)
        return True
    except Exception as e:
        print(f"❌ Erro ao deletar {nome_arquivo}: {e}")
        return False

# ==================================================================================
# UTILITÁRIOS DE PROCESSO
# ==================================================================================

import psutil

def processo_existe(nome_script):
    """Verifica se um processo Python com o script está rodando."""
    nome_script_lower = nome_script.lower()
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            cmdline = proc.info.get('cmdline')
            if not cmdline:
                continue
            is_python = 'python' in proc.info.get('name', '').lower()
            if is_python and any(arg.lower().endswith(nome_script_lower) for arg in cmdline):
                return proc
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return None

def matar_processo(pid):
    """Mata um processo pelo PID."""
    try:
        proc = psutil.Process(pid)
        proc.terminate()
        proc.wait(timeout=5)
        return True
    except (psutil.NoSuchProcess, psutil.TimeoutExpired):
        try:
            proc.kill()
            return True
        except:
            return False

# ==================================================================================
# UTILITÁRIOS DE PARSING
# ==================================================================================

def parse_float(texto, default=0.0):
    """Converte texto em float com segurança."""
    if not texto:
        return default
    try:
        return float(str(texto).replace(',', '.'))
    except ValueError:
        return default

def parse_int(texto, default=0):
    """Converte texto em int com segurança."""
    if not texto:
        return default
    try:
        return int(str(texto).replace(',', '.'))
    except ValueError:
        return default

# ==================================================================================
# REPOUSO E DELAYS
# ==================================================================================

import random

def esperar_humano(min_s=2, max_s=4):
    """Aguarda um tempo aleatório para simular comportamento humano."""
    tempo = random.uniform(min_s, max_s)
    time.sleep(tempo)

def fazer_delay(segundos, com_jitter=True):
    """Faz delay com opção de jitter."""
    if com_jitter:
        segundos = random.uniform(segundos * 0.8, segundos * 1.2)
    time.sleep(segundos)

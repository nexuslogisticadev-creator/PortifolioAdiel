def gravar_subtabela_pagamentos(data_str, pagamentos):
    """Adiciona uma sub-tabela de pagamentos dos motoboys ao final da aba do dia no Google Sheets."""
    aba = _nome_aba(data_str)
    ws = _aba(aba)
    header = ["PAGAMENTOS MOTOBOYS", "QTD TOTAL", "QTD R$ 8,00", "QTD R$ 11,00", "TOTAL A PAGAR (R$)"]
    # Remover sub-tabela antiga se existir
    all_vals = ws.get_all_values()
    linhas_remover = []
    for i, row in enumerate(all_vals):
        if row and row[0] == "PAGAMENTOS MOTOBOYS":
            # Encontrou cabeçalho, remove até próxima linha em branco ou fim
            j = i
            while j < len(all_vals) and any(all_vals[j]):
                linhas_remover.append(j+1)  # Sheet é 1-indexed
                j += 1
            break
    for linha in reversed(linhas_remover):
        try:
            ws.delete_rows(linha)
        except Exception:
            pass
    # Grava nova sub-tabela
    ws.append_row(["" for _ in range(len(HEADERS))], value_input_option="RAW")  # Linha em branco
    ws.append_row(header, value_input_option="USER_ENTERED")
    for nome, dados in pagamentos.items():
        row = [nome, dados["qtd"], dados["qtd8"], dados["qtd11"], f"R$ {dados['total']:.2f}"]
        ws.append_row(row, value_input_option="USER_ENTERED")
    print(f"✅ Sub-tabela de pagamentos gravada no Sheets ({data_str})")
"""
db_sheets.py — Banco de dados Google Sheets
Cada dia = uma aba (DD-MM-YYYY). Pedidos e vales na mesma aba,
diferenciados pela coluna "Tipo" (ENTREGA ou VALE).
"""

import os
import time
import threading
from datetime import datetime, timedelta

try:
    import gspread
    from google.oauth2.service_account import Credentials
    from gspread.exceptions import WorksheetNotFound, APIError
    TEM_SHEETS = True
except ImportError:
    TEM_SHEETS = False

# ─── Config ──────────────────────────────────────────────────────────────────────
SPREADSHEET_ID = "1BNNntndWkEKGUn5SfodKF3E93RUuamUg0_ccVy19O1w"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Detecta ambiente PyInstaller
def _get_service_account_path():
    import sys
    if hasattr(sys, '_MEIPASS'):
        # Executável PyInstaller
        return os.path.join(sys._MEIPASS, "gen-lang-client-0592009269-3d0b6d104f80.json")
    else:
        # Execução normal
        _BASE = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(_BASE, "gen-lang-client-0592009269-3d0b6d104f80.json")

SERVICE_ACCOUNT_JSON = _get_service_account_path()

HEADERS = [
    "Data", "Hora", "Numero", "Cliente", "Bairro", "Status",
    "Motoboy", "Combo", "Valor", "Itens", "Tipo", "Vale_Motivo",
]

# ─── Estado interno ──────────────────────────────────────────────────────────────
_client = None
_spreadsheet = None
_lock = threading.Lock()
_cache = {}        # {nome_aba: {"rows": list[dict], "ts": float}}
CACHE_TTL = 15     # segundos


# ─── Helpers ─────────────────────────────────────────────────────────────────────

def _nome_aba(data_str=None):
    """Retorna nome da aba no formato DD-MM-YYYY."""
    if data_str:
        # Remove caminho, extensão e caracteres inválidos
        nome = os.path.basename(str(data_str))
        nome = nome.replace('.xlsx', '')
        nome = nome.replace('/', '-')
        nome = nome.replace('\\', '-')
        # Remove espaços e caracteres especiais
        import re
        nome = re.sub(r'[^\w\-]', '', nome)
        # Garante prefixo padrão
        if not nome.startswith('Controle_Financeiro_'):
            nome = f'Controle_Financeiro_{nome}'
        return nome
    agora = datetime.now()
    if agora.hour < 10:
        agora -= timedelta(days=1)
    return agora.strftime("%d-%m-%Y")


def conectar():
    """Autentica e retorna o spreadsheet (reutiliza conexão)."""
    global _client, _spreadsheet
    with _lock:
        if _client is not None and _spreadsheet is not None:
            return _spreadsheet
    if not TEM_SHEETS:
        raise RuntimeError("gspread / google-auth não instalados.")
    if not os.path.exists(SERVICE_ACCOUNT_JSON):
        raise FileNotFoundError(f"Credenciais não encontradas: {SERVICE_ACCOUNT_JSON}")
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_JSON, scopes=SCOPES)
    cli = gspread.authorize(creds)
    sh = cli.open_by_key(SPREADSHEET_ID)
    with _lock:
        _client = cli
        _spreadsheet = sh
    return sh


def reconectar():
    """Força reconexão (útil após erro de auth expirado)."""
    global _client, _spreadsheet
    with _lock:
        _client = None
        _spreadsheet = None
    return conectar()


def _aba(nome):
    """Obtém ou cria worksheet para o dia."""
    sh = conectar()
    try:
        return sh.worksheet(nome)
    except WorksheetNotFound:
        ws = sh.add_worksheet(title=nome, rows=500, cols=len(HEADERS))
        ws.append_row(HEADERS, value_input_option="RAW")
        try:
            ws.format("A1:L1", {"textFormat": {"bold": True}})
        except Exception:
            pass
        return ws


# ─── Cache ───────────────────────────────────────────────────────────────────────

def _cache_get(aba):
    with _lock:
        e = _cache.get(aba)
        if e and time.time() - e["ts"] < CACHE_TTL:
            return e["rows"]
    return None


def _cache_set(aba, rows):
    with _lock:
        _cache[aba] = {"rows": rows, "ts": time.time()}


def invalidar_cache(data_str=None):
    """Invalida cache de uma aba ou de todas."""
    with _lock:
        if data_str:
            _cache.pop(_nome_aba(data_str), None)
        else:
            _cache.clear()


# ─── LEITURA ─────────────────────────────────────────────────────────────────────

def ler_todos_registros(data_str=None):
    """Retorna list[dict] com TODOS os registros do dia (pedidos + vales)."""
    aba = _nome_aba(data_str)
    cached = _cache_get(aba)
    if cached is not None:
        return cached
    try:
        ws = _aba(aba)
        registros = ws.get_all_records()
        _cache_set(aba, registros)
        return registros
    except APIError as e:
        if "401" in str(e) or "UNAUTHENTICATED" in str(e):
            reconectar()
            try:
                ws = _aba(aba)
                registros = ws.get_all_records()
                _cache_set(aba, registros)
                return registros
            except Exception:
                pass
        print(f"❌ Sheets API error: {e}")
        return []
    except Exception as e:
        print(f"❌ Erro leitura Sheets: {e}")
        return []


def ler_pedidos(data_str=None):
    """Retorna apenas pedidos (Tipo != VALE)."""
    return [r for r in ler_todos_registros(data_str) if r.get("Tipo", "ENTREGA") != "VALE"]


def ler_vales(data_str=None):
    """Retorna apenas vales (Tipo == VALE)."""
    return [r for r in ler_todos_registros(data_str) if r.get("Tipo") == "VALE"]


def ler_vales_com_linhas(data_str=None):
    """Retorna [(sheet_row, hora, moto, valor, motivo), ...] com número real da linha."""
    aba = _nome_aba(data_str)
    ws = _aba(aba)
    all_vals = ws.get_all_values()
    result = []
    for i, row in enumerate(all_vals):
        if i == 0:
            continue  # header
        if len(row) > 10 and row[10] == "VALE":
            try:
                val = float(row[8]) if row[8] else 0.0
            except (ValueError, TypeError):
                val = 0.0
            motivo = row[11] if len(row) > 11 else ""
            result.append((i + 1, row[1], row[6], val, motivo))  # i+1 = sheet row (1-indexed)
    return result


def ler_pedidos_como_tuplas(data_str=None):
    """Retorna pedidos no mesmo formato de ws.iter_rows(values_only=True).
    Cada tupla: (Data, Hora, Numero, Cliente, Bairro, Status, Motoboy, Combo, Valor, Itens)
    """
    registros = ler_pedidos(data_str)
    result = []
    for r in registros:
        try:
            val = float(r.get("Valor", 0))
        except (ValueError, TypeError):
            val = 0.0
        result.append((
            r.get("Data", ""), r.get("Hora", ""), str(r.get("Numero", "")),
            r.get("Cliente", ""), r.get("Bairro", ""), str(r.get("Status", "")),
            r.get("Motoboy", ""), r.get("Combo", ""), val,
            r.get("Itens", "")
        ))
    return result


def ler_vales_como_tuplas(data_str=None):
    """Retorna vales no mesmo formato da aba VALES antiga.
    Cada tupla: (Hora, Motoboy, Valor, Motivo)
    """
    registros = ler_vales(data_str)
    result = []
    for r in registros:
        try:
            val = float(r.get("Valor", 0))
        except (ValueError, TypeError):
            val = 0.0
        result.append((
            r.get("Hora", ""), r.get("Motoboy", ""), val, r.get("Vale_Motivo", "")
        ))
    return result


# ─── ESCRITA ─────────────────────────────────────────────────────────────────────

def inicializar_aba_dia(data_str=None):
    """Cria aba se não existir. Retorna registros existentes."""
    aba = _nome_aba(data_str)
    _aba(aba)
    return ler_todos_registros(data_str)


def registrar_pedido(dados_pedido, data_str=None):
    """Registra ou atualiza pedido. Retorna True/False."""
    numero = str(dados_pedido.get("numero", "")).strip()
    if not numero:
        return False
    motoboy = str(dados_pedido.get("motoboy", "")).strip()
    if motoboy in ("Desconhecido", "Aguardando..."):
        return False

    aba = _nome_aba(data_str)
    ws = _aba(aba)

    dt = dados_pedido.get("_dt") or datetime.now()
    status = str(dados_pedido.get("status", "")).upper()

    try:
        valor = float(dados_pedido.get("valor", 0))
    except (ValueError, TypeError):
        valor = 0.0

    row = [
        dt.strftime("%d/%m/%Y") if isinstance(dt, datetime) else str(dt),
        dt.strftime("%H:%M") if isinstance(dt, datetime) else "",
        numero,
        dados_pedido.get("cliente", ""),
        dados_pedido.get("bairro", ""),
        status,
        motoboy,
        dados_pedido.get("combo", ""),
        valor,
        dados_pedido.get("itens", ""),
        "ENTREGA",
        "",
    ]

    try:
        # Busca linha existente pelo numero (coluna C)
        col_c = ws.col_values(3)
        linha = None
        for i, v in enumerate(col_c):
            if str(v).strip() == numero:
                linha = i + 1
                break

        if linha:
            ws.update(f"A{linha}:L{linha}", [row], value_input_option="USER_ENTERED")
        else:
            ws.append_row(row, value_input_option="USER_ENTERED")

        invalidar_cache(data_str)
        return True
    except Exception as e:
        print(f"❌ Erro registrar pedido Sheets: {e}")
        return False


def registrar_vale(nome_moto, valor, motivo="Desconto/Vale", data_str=None):
    """Registra um vale na aba do dia."""
    aba = _nome_aba(data_str)
    ws = _aba(aba)
    agora = datetime.now()

    row = [
        agora.strftime("%d/%m/%Y"),
        agora.strftime("%H:%M"),
        "", "", "", "",  # numero, cliente, bairro, status vazios
        nome_moto,
        "",  # combo
        float(valor),
        "",  # itens
        "VALE",
        motivo,
    ]

    try:
        ws.append_row(row, value_input_option="USER_ENTERED")
        invalidar_cache(data_str)
        print(f"💾 Vale registrado: {nome_moto} R$ {valor}")
        return True
    except Exception as e:
        print(f"❌ Erro registrar vale: {e}")
        return False


def excluir_linha(data_str, sheet_row):
    """Remove uma linha pelo número da linha no Sheet (1-indexed)."""
    aba = _nome_aba(data_str)
    ws = _aba(aba)
    try:
        ws.delete_rows(sheet_row)
        invalidar_cache(data_str)
        return True
    except Exception as e:
        print(f"❌ Erro excluir linha: {e}")
        return False


def editar_celula(data_str, sheet_row, col_idx, valor):
    """Edita uma célula. col_idx = 1-indexed."""
    aba = _nome_aba(data_str)
    ws = _aba(aba)
    try:
        ws.update_cell(sheet_row, col_idx, valor)
        invalidar_cache(data_str)
        return True
    except Exception as e:
        print(f"❌ Erro editar célula: {e}")
        return False


def atualizar_pedido(data_str, numero, novos_dados):
    """Atualiza campos de um pedido existente."""
    aba = _nome_aba(data_str)
    ws = _aba(aba)

    try:
        col_c = ws.col_values(3)
        linha = None
        for i, v in enumerate(col_c):
            if str(v).strip() == str(numero).strip():
                linha = i + 1
                break

        if not linha:
            return False

        current = ws.row_values(linha)
        while len(current) < len(HEADERS):
            current.append("")

        # Mapa campo → índice 0-based na lista
        mapa = {
            "Bairro": 4,
            "Status": 5,
            "Motoboy": 6,
            "Valor": 8,
            "Valor (R$)": 8,
        }

        for campo, col in mapa.items():
            if campo in novos_dados:
                current[col] = novos_dados[campo]

        ws.update(f"A{linha}:L{linha}", [current], value_input_option="USER_ENTERED")
        invalidar_cache(data_str)
        return True
    except Exception as e:
        print(f"❌ Erro atualizar pedido: {e}")
        return False

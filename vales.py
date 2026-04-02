# ===============================
# Módulo de Vales e Descontos
# ===============================
import json
import os
from utils import get_caminho_base


def carregar_vales():
    """Carrega vales do arquivo JSON."""
    caminho = os.path.join(get_caminho_base(), 'vales.json')
    try:
        with open(caminho, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def salvar_vales(vales_data):
    """Salva vales no disco."""
    caminho = os.path.join(get_caminho_base(), 'vales.json')
    try:
        with open(caminho, 'w', encoding='utf-8') as f:
            json.dump(vales_data, f, indent=4, ensure_ascii=False)
        return True
    except Exception:
        return False


def adicionar_vale(nome, valor, descricao):
    """Adiciona um novo vale."""
    vales_data = carregar_vales()
    vales_data.append({
        "nome": nome,
        "valor": valor,
        "descricao": descricao
    })
    salvar_vales(vales_data)
    return vales_data


def remover_vale(nome):
    """Remove vale pelo nome."""
    vales_data = carregar_vales()
    vales_data = [v for v in vales_data if v.get("nome", "") != nome]
    salvar_vales(vales_data)
    return vales_data


def buscar_vale(nome):
    """Busca vale pelo nome."""
    vales_data = carregar_vales()
    return next((v for v in vales_data if v.get("nome", "") == nome), None)


# Outras funções de vales podem ser adicionadas aqui

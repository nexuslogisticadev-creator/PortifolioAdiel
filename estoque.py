# ===============================
# Módulo de Estoque
# ===============================
import json
import os
from utils import ARQUIVO_ESTOQUE, ARQUIVO_ESTOQUE_BAIXAS, get_caminho_base


def carregar_estoque(estoque_categorias=None, identificar_categoria_func=None):
    """Carrega o estoque do arquivo JSON, garantindo retorno de lista e compatibilidade com legado.
    estoque_categorias: lista de categorias válidas (opcional)
    identificar_categoria_func: função para identificar categoria (opcional)
    """
    caminho = os.path.join(get_caminho_base(), ARQUIVO_ESTOQUE)
    if not os.path.exists(caminho):
        return []

    try:
        if os.path.exists(caminho):
            with open(caminho, "r", encoding="utf-8") as f:

                dados = json.load(f)
    except Exception as e:
        print(f"Erro ao carregar o estoque: {e}")
        return []

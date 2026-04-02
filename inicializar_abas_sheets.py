from db_sheets import inicializar_aba_dia, conectar
from datetime import datetime

def inicializar_abas_principais():
    # Inicializa aba de estoque
    sh = conectar()
    try:
        ws_estoque = sh.worksheet("Estoque")
    except Exception:
        ws_estoque = sh.add_worksheet(title="Estoque", rows=200, cols=10)
        ws_estoque.append_row(["Nome", "Quantidade", "Categoria", "Preco", "Fornecedor"], value_input_option="RAW")
    print("✅ Aba 'Estoque' pronta.")

    # Inicializa aba de memoria de fechamento
    try:
        ws_memoria = sh.worksheet("MemoriaFechamento")
    except Exception:
        ws_memoria = sh.add_worksheet(title="MemoriaFechamento", rows=100, cols=10)
        ws_memoria.append_row(["Data", "HoraIni", "HoraFim", "Motoboy", "Observacao"], value_input_option="RAW")
    print("✅ Aba 'MemoriaFechamento' pronta.")

    # Inicializa aba do dia atual
    hoje = datetime.now().strftime("%d-%m-%Y")
    inicializar_aba_dia(hoje)
    print(f"✅ Aba do dia '{hoje}' pronta.")

if __name__ == "__main__":
    inicializar_abas_principais()
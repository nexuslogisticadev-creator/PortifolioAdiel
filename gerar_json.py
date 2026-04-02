import pandas as pd
import json
import os

# Caminhos absolutos dos dois arquivos
caminho_fornecedores = r"C:\Users\Usuario\Desktop\githubnew\Conferencia_24_23-02-2026.xlsx"
caminho_quantidades = r"C:\Users\Usuario\Desktop\githubnew\RelatorioCompleto_Inventario_2026-02-24 (1).xlsx"
caminho_json = "estoque.json"

print("Iniciando o cruzamento dos arquivos (Extraindo Fornecedores, Categorias e Preços)...\n")

# ==========================================
# PASSO 1: EXTRAIR FORNECEDORES, CATEGORIAS E PREÇOS
# ==========================================
print(f"-> Lendo o arquivo Base:\n{caminho_fornecedores}")
try:
    df_forn_bruto = pd.read_excel(caminho_fornecedores, header=None)
except Exception as e:
    print(f"❌ ERRO ao ler o arquivo Base: {e}")
    exit()

linha_cabecalho_forn = -1
for indice, linha in df_forn_bruto.head(15).iterrows():
    valores_linha = [str(celula).lower().strip() for celula in linha.values]
    if ('nome' in valores_linha or 'produto' in valores_linha) and any('estoque' in v for v in valores_linha):
        linha_cabecalho_forn = indice
        break

if linha_cabecalho_forn == -1:
    print("❌ ERRO: Não achei o cabeçalho no arquivo Base.")
    exit()

df_forn = df_forn_bruto.copy()
df_forn.columns = [str(c).strip() for c in df_forn.iloc[linha_cabecalho_forn]]
df_forn = df_forn[(linha_cabecalho_forn + 1):].reset_index(drop=True)

colunas_forn_lower = [str(c).lower() for c in df_forn.columns]
idx_nome_forn = next(i for i, col in enumerate(colunas_forn_lower) if col == 'nome' or col == 'produto')

# Busca a coluna do Fornecedor
try:
    idx_marca_forn = next(i for i, col in enumerate(colunas_forn_lower) if 'fornecedor' in col or 'fabricante' in col or 'marca' in col)
    col_marca_forn = df_forn.columns[idx_marca_forn]
except StopIteration:
    col_marca_forn = None

# Busca a coluna da Categoria
try:
    idx_categoria_forn = next(i for i, col in enumerate(colunas_forn_lower) if 'categoria' in col)
    col_categoria_forn = df_forn.columns[idx_categoria_forn]
except StopIteration:
    col_categoria_forn = None

# ADICIONADO: Busca a coluna do Preço (Coluna L - Preço de Venda Atual)
try:
    idx_preco_forn = next(i for i, col in enumerate(colunas_forn_lower) if 'preço de venda' in col or 'preco de venda' in col)
    col_preco_forn = df_forn.columns[idx_preco_forn]
except StopIteration:
    col_preco_forn = None

col_nome_forn = df_forn.columns[idx_nome_forn]

# Criar um "Mapa" que guarda Fornecedor, Categoria E Preço para cada Produto
mapa_info = {}
for index, row in df_forn.iterrows():
    nome_prod = str(row[col_nome_forn]).strip().lower()
    
    fornecedor = "Diversos"
    if col_marca_forn:
        forn_str = str(row[col_marca_forn]).strip()
        if forn_str.lower() not in ['nan', 'none', '']:
            fornecedor = forn_str
            
    categoria = "GERAL"
    if col_categoria_forn:
        cat_str = str(row[col_categoria_forn]).strip()
        if cat_str.lower() not in ['nan', 'none', '']:
            categoria = cat_str
            
    # ADICIONADO: Extração e limpeza do preço
    preco_venda = 0.0
    if col_preco_forn:
        preco_str = str(row[col_preco_forn]).strip()
        if preco_str.lower() not in ['nan', 'none', '']:
            # Limpa o texto "R$ ", tira pontos de milhar e troca vírgula por ponto
            preco_str = preco_str.replace('R$', '').replace('r$', '').strip()
            preco_str = preco_str.replace('.', '').replace(',', '.')
            try:
                preco_venda = float(preco_str)
            except ValueError:
                preco_venda = 0.0

    mapa_info[nome_prod] = {
        "fornecedor": fornecedor,
        "categoria": categoria,
        "preco_venda": preco_venda
    }

print(f"✅ Sucesso! Memorizei informações de {len(mapa_info)} produtos do primeiro arquivo.\n")

# ==========================================
# PASSO 2: EXTRAIR PRODUTOS/QTD DO ARQUIVO 2
# ==========================================
print(f"-> Lendo o arquivo de Quantidades:\n{caminho_quantidades}")
try:
    try:
        df_qtd_bruto = pd.read_excel(caminho_quantidades, sheet_name="Relatório Resumido", header=None)
    except Exception:
        df_qtd_bruto = pd.read_excel(caminho_quantidades, header=None)
except Exception as e:
    print(f"❌ ERRO ao ler o arquivo de Quantidades: {e}")
    exit()

linha_cabecalho_qtd = -1
for indice, linha in df_qtd_bruto.head(15).iterrows():
    valores_linha = [str(celula).lower().strip() for celula in linha.values]
    if ('nome' in valores_linha or 'produto' in valores_linha) and any('estoque' in v or 'qtd' in v or 'quantidade' in v for v in valores_linha):
        linha_cabecalho_qtd = indice
        break

if linha_cabecalho_qtd == -1:
    print("❌ ERRO: Não achei o cabeçalho no arquivo de quantidades.")
    exit()

df_qtd = df_qtd_bruto.copy()
df_qtd.columns = [str(c).strip() for c in df_qtd.iloc[linha_cabecalho_qtd]]
df_qtd = df_qtd[(linha_cabecalho_qtd + 1):].reset_index(drop=True)

colunas_qtd_lower = [str(c).lower() for c in df_qtd.columns]
idx_nome_qtd = next(i for i, col in enumerate(colunas_qtd_lower) if col == 'nome' or col == 'produto' or col == 'descrição')

try:
    idx_estoque_qtd = next(i for i, col in enumerate(colunas_qtd_lower) if 'físico' in col or 'conferido' in col)
except StopIteration:
    idx_estoque_qtd = next(i for i, col in enumerate(colunas_qtd_lower) if 'estoque' in col or 'qtd' in col or 'quantidade' in col)

col_nome_qtd = df_qtd.columns[idx_nome_qtd]
col_estoque_qtd = df_qtd.columns[idx_estoque_qtd]

estoque_final = []

for index, row in df_qtd.iterrows():
    nome_produto = str(row[col_nome_qtd]).strip()
    estoque_bruto = str(row[col_estoque_qtd]).strip()
    
    try:
        if estoque_bruto == '' or estoque_bruto.lower() in ['nan', 'none']:
            qtd = 0.0
        else:
            qtd = float(estoque_bruto.replace(',', '.'))
    except ValueError:
        qtd = 0.0

    if nome_produto and nome_produto.lower() not in ['nan', 'none', 'total', 'subtotal']:
        
        # Cruzamento: Procura as infos no mapa gerado no Passo 1
        nome_busca = nome_produto.lower()
        fornecedor_encontrado = "Diversos"
        categoria_encontrada = "GERAL"
        preco_encontrado = 0.0  # ADICIONADO: Variável para o preço
        
        # Tenta achar o nome exato
        if nome_busca in mapa_info:
            fornecedor_encontrado = mapa_info[nome_busca]["fornecedor"]
            categoria_encontrada = mapa_info[nome_busca]["categoria"]
            preco_encontrado = mapa_info[nome_busca]["preco_venda"]
        else:
            # Se não for exato, tenta ver se um nome está contido dentro do outro
            for nome_key, info in mapa_info.items():
                if nome_key in nome_busca or nome_busca in nome_key:
                    fornecedor_encontrado = info["fornecedor"]
                    categoria_encontrada = info["categoria"]
                    preco_encontrado = info["preco_venda"]
                    break

        estoque_final.append({
            "nome": nome_produto,
            "estoque_fisico": qtd,
            "categoria": categoria_encontrada,
            "preco_venda": preco_encontrado, # ADICIONADO: Agora passa o preço real extraído
            "fornecedor": fornecedor_encontrado
        })

# ==========================================
# PASSO 3: SALVAR ARQUIVO JSON FINAL
# ==========================================
with open(caminho_json, "w", encoding="utf-8") as f:
    json.dump(estoque_final, f, ensure_ascii=False, indent=2)

print(f"✅ SUCESSO ABSOLUTO! O cruzamento foi concluído.")
print(f"Foram exportados {len(estoque_final)} produtos (com Categorias, Fornecedores e Preços) para o arquivo '{caminho_json}'.")
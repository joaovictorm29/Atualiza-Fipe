import openpyxl
import unicodedata
import difflib
import requests

CAMINHO_PLANILHA = "Planilha Veículos c- depreciação.xlsx"

wb = openpyxl.load_workbook(CAMINHO_PLANILHA)
ws = wb["VEICULOS"]

LINHA_CABECALHO = 2
LINHA_INICIO_DADOS = 3  # dados começam na linha 3 (linha 2 é cabeçalho)

colunas = {}
for cell in ws[LINHA_CABECALHO]:
    if cell.value is not None:
        nome_normalizado = cell.value.strip().upper()
        colunas[nome_normalizado] = cell.column  # número da coluna (int)

col_cod = colunas["CÓD ATIVO"]
col_tipo = colunas["TIPO"]
col_modelo = colunas["MODELO"]
col_marca = colunas["MARCA"]
col_categoria = colunas["CATEGORIA FIPE"]  # nova coluna: "cars" ou "trucks"
col_ano = colunas["ANO"]
col_fipe = colunas["VALOR FIPE CHASSI"]  # ainda não usamos, mas vamos precisar pra escrever o resultado

veiculos = []

for num_linha in range(LINHA_INICIO_DADOS, ws.max_row + 1):
    cod = ws.cell(row=num_linha, column=col_cod).value
    tipo = ws.cell(row=num_linha, column=col_tipo).value
    modelo = ws.cell(row=num_linha, column=col_modelo).value
    marca = ws.cell(row=num_linha, column=col_marca).value
    categoria_fipe = ws.cell(row=num_linha, column=col_categoria).value
    ano = ws.cell(row=num_linha, column=col_ano).value

    # pula linha totalmente vazia (fim da tabela)
    if cod is None:
        continue

    # pula linha de totalização (não é um veículo)
    if "TOTAL" in str(cod).upper():
        continue

    # pula veículo com dado essencial faltando
    if marca is None or modelo is None or ano is None or categoria_fipe is None:
        print(f"[AVISO] Linha {num_linha} ({cod}) sem MARCA/MODELO/ANO/CATEGORIA — pulando.")
        continue

    veiculos.append({
        "linha": num_linha,
        "codigo": cod,
        "tipo": tipo,
        "modelo": modelo,
        "marca": marca,
        "categoria_fipe": categoria_fipe.strip().lower(),  # garante "cars"/"trucks" mesmo com maiúscula/espaço
        "ano": ano,
    })

print(f"Total de veículos válidos encontrados: {len(veiculos)}\n")

# aqui eu vou tentar normalizar os nomes para evitar erros na busca
def normalizar(texto: str) -> str:
    """Remove acentos, deixa maiúsculo e tira espaços das pontas."""
    # NFKD separa a letra do acento (ex: 'ç' -> 'c' + '¸')
    sem_acento = unicodedata.normalize("NFKD", texto)
    # mantém só os caracteres que NÃO são marca de acentuação (Mn = "Mark, nonspacing")
    sem_acento = "".join(c for c in sem_acento if unicodedata.category(c) != "Mn")
    return sem_acento.strip().upper()


def encontrar_marca(texto_marca: str, marcas_api: list[dict]) -> dict | None:
    """Entrada já é só o nome da marca (coluna MARCA), comparação direta."""
    alvo = normalizar(texto_marca)
    nomes_api_normalizados = {normalizar(m["name"]): m for m in marcas_api}

    candidatos = difflib.get_close_matches(
        alvo, nomes_api_normalizados.keys(), n=1, cutoff=0.75
        # funciona tipo como uma porcentagem de acerto, se for menor que 0.75 não considera como candidato
    )
    if candidatos:
        return nomes_api_normalizados[candidatos[0]]
    return None


marcas_carros = requests.get("https://fipe.parallelum.com.br/api/v2/cars/brands").json()
marcas_caminhoes = requests.get("https://fipe.parallelum.com.br/api/v2/trucks/brands").json()

# antes juntávamos tudo em "todas_as_marcas" -- não faz mais sentido, porque
# agora sabemos a categoria certa de cada veículo pela coluna CATEGORIA FIPE,
# então buscamos direto na lista certa (evita marca de carro colidir com marca de caminhão)

for v in veiculos:
    marcas_da_categoria_bruta = marcas_carros if v["categoria_fipe"] == "cars" else marcas_caminhoes
    marcas_da_categoria = [{"categoria": v["categoria_fipe"], **m} for m in marcas_da_categoria_bruta]

    marca_api = encontrar_marca(v["marca"], marcas_da_categoria)
    if marca_api:
        v["marca_code"] = marca_api["code"]
        print(f"{v['codigo']:8} | {v['categoria_fipe']:6} | MARCA planilha={v['marca']:15} -> {marca_api['name']} (code={marca_api['code']})")
    else:
        print(f"{v['codigo']:8} | {v['categoria_fipe']:6} | MARCA planilha={v['marca']:15} -> ⚠️  NÃO ENCONTRADO")
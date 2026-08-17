import openpyxl
import unicodedata
import difflib
import requests
import re

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


for v in veiculos:
    marcas_da_categoria_bruta = marcas_carros if v["categoria_fipe"] == "cars" else marcas_caminhoes
    marcas_da_categoria = [{"categoria": v["categoria_fipe"], **m} for m in marcas_da_categoria_bruta]

    marca_api = encontrar_marca(v["marca"], marcas_da_categoria)
    if marca_api:
        v["marca_code"] = marca_api["code"]
        print(f"{v['codigo']:8} | {v['categoria_fipe']:6} | MARCA planilha={v['marca']:15} -> {marca_api['name']} (code={marca_api['code']})")
    else:
        print(f"{v['codigo']:8} | {v['categoria_fipe']:6} | MARCA planilha={v['marca']:15} -> ⚠️  NÃO ENCONTRADO")


def buscar_modelos(categoria: str, marca_code: str) -> list[dict]:
    url = f"https://fipe.parallelum.com.br/api/v2/{categoria}/brands/{marca_code}/models"
    resposta = requests.get(url)
    resposta.raise_for_status()
    return resposta.json()


# teste com o primeiro caminhão Volkswagen (CBV-01)
modelos_vw_caminhao = buscar_modelos("trucks", "115")
print(f"\nTotal de modelos VW caminhão: {len(modelos_vw_caminhao)}")
for m in modelos_vw_caminhao[:15]:
    print(m)


def extrair_codigo_numerico(texto_modelo: str) -> str | None:
    """Procura um padrão tipo '31.320' ou '11-180' dentro do texto e
    normaliza para o formato da API, que usa hífen (ex: '31-320')."""
    match = re.search(r"(\d{1,2})[.\-](\d{3})", texto_modelo)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
    return None


def encontrar_modelo_por_codigo(codigo: str, modelos_api: list[dict]) -> dict | None:
    """Acha o primeiro modelo da API cujo nome começa com o código extraído."""
    for m in modelos_api:
        if m["name"].startswith(codigo):
            return m
    return None


# --- teste com CBV-01 ---
texto_teste = "VW - 31.320 6X4/ 14M³ (THR-6E50)"
codigo = extrair_codigo_numerico(texto_teste)
print(f"\nCódigo extraído: {codigo}")

modelo_encontrado = encontrar_modelo_por_codigo(codigo, modelos_vw_caminhao)
print(f"Modelo encontrado na API: {modelo_encontrado}")


def encontrar_todos_modelos_por_codigo(codigo: str, modelos_api: list[dict]) -> list[dict]:
    """Em vez de devolver só o primeiro, devolve TODOS os que batem —
    assim a gente enxerga quando há mais de uma opção."""
    return [m for m in modelos_api if m["name"].startswith(codigo)]


# roda pra todo mundo que é VW + trucks, sem gravar nada ainda, só pra inspecionar
print()
for v in veiculos:
    if v["categoria_fipe"] != "trucks" or v.get("marca_code") != "115":
        continue  # por enquanto só olhando os caminhões VW, que é onde o padrão numérico funciona

    codigo = extrair_codigo_numerico(v["modelo"])
    if codigo is None:
        print(f"{v['codigo']:8} | MODELO={v['modelo']:45} -> ⚠️  não achei código numérico")
        continue

    candidatos = encontrar_todos_modelos_por_codigo(codigo, modelos_vw_caminhao)
    print(f"{v['codigo']:8} | código={codigo} | {len(candidatos)} candidato(s):")
    for c in candidatos:
        print(f"      {c}")

PISTAS_DESEMPATE = ["4X2", "4X4", "6X2", "6X4", "8X4", "DELIVERY", "CONSTELLATION", "TRACTOR", "3-EIXOS", "4-EIXOS"]

def desambiguar_modelo(texto_modelo_original: str, candidatos: list[dict]):
    """Recebe os candidatos empatados e tenta desempatar usando palavras-chave
    que já aparecem no texto original da planilha (ex: '4X2', 'Delivery')."""
    texto_normalizado = normalizar(texto_modelo_original)
    pistas_presentes = [p for p in PISTAS_DESEMPATE if p in texto_normalizado]

    pontuacoes = []
    for c in candidatos:
        nome_normalizado = normalizar(c["name"])
        pontuacao = sum(1 for pista in pistas_presentes if pista in nome_normalizado)
        pontuacoes.append((pontuacao, c))

    maior_pontuacao = max(p for p, _ in pontuacoes)
    vencedores = [c for p, c in pontuacoes if p == maior_pontuacao]

    if len(vencedores) == 1:
        return vencedores[0]
    return None  # continua ambíguo mesmo com as pistas


# --- aplica em todas as marcas e categorias ---
# Mantém a etapa específica da Volkswagen acima como teste/inspeção, mas a
# resolução efetiva abaixo usa a marca e a categoria de cada veículo.
MODELOS_EM_CACHE = {}
PISTAS_TECNICAS = ["4X2", "4X4", "6X2", "6X4", "8X4", "DELIVERY", "CONSTELLATION", "TRACTOR"]


def normalizar_modelo_para_comparacao(texto: str) -> str:
    """Remove a placa entre parênteses, que não faz parte do nome FIPE."""
    texto = normalizar(texto)
    return re.sub(r"\([A-Z]{3}-[A-Z0-9]{4}\)", " ", texto)


def tokens_do_modelo(texto: str) -> set[str]:
    """Gera termos comparáveis, preservando códigos como 31.320 e 4X2."""
    return set(re.findall(r"[A-Z0-9]+(?:[.\-][A-Z0-9]+)?", texto))


def encontrar_modelo_generico(texto_modelo: str, modelos_api: list[dict]):
    """Escolhe um modelo apenas quando a melhor correspondência é confiável.

    Caminhões costumam ter um código (ex.: 31-320), enquanto carros em geral
    dependem da semelhança entre os termos do nome. Casos próximos continuam
    como ambíguos para revisão humana, evitando escrever um modelo errado.
    """
    texto = normalizar_modelo_para_comparacao(texto_modelo)
    termos_texto = tokens_do_modelo(texto)
    codigo_numerico = extrair_codigo_numerico(texto_modelo)
    pontuacoes = []

    for modelo_api in modelos_api:
        nome_api = normalizar_modelo_para_comparacao(modelo_api["name"])
        termos_api = tokens_do_modelo(nome_api)
        termos_em_comum = termos_texto & termos_api
        cobertura = len(termos_em_comum) / max(len(termos_texto), 1)
        semelhanca = difflib.SequenceMatcher(None, texto, nome_api).ratio()
        pontuacao = cobertura * 60 + semelhanca * 40

        # Código do veículo é o indício mais forte para caminhões.
        if codigo_numerico and nome_api.startswith(codigo_numerico):
            pontuacao += 100

        pontuacao += sum(
            10 for pista in PISTAS_TECNICAS
            if pista in texto and pista in nome_api
        )
        pontuacoes.append((pontuacao, modelo_api))

    if not pontuacoes:
        return None, []

    pontuacoes.sort(key=lambda item: item[0], reverse=True)
    melhor_pontuacao, melhor_modelo = pontuacoes[0]
    segunda_pontuacao = pontuacoes[1][0] if len(pontuacoes) > 1 else -1
    candidatos_com_mesmo_codigo = [
        item for item in pontuacoes
        if codigo_numerico and normalizar_modelo_para_comparacao(item[1]["name"]).startswith(codigo_numerico)
    ]
    tem_codigo_confiavel = (
        candidatos_com_mesmo_codigo
        and (
            len(candidatos_com_mesmo_codigo) == 1
            or melhor_pontuacao - candidatos_com_mesmo_codigo[1][0] >= 5
        )
    )

    # Sem código, exige uma boa nota e diferença clara para o segundo colocado.
    if tem_codigo_confiavel or (melhor_pontuacao >= 55 and melhor_pontuacao - segunda_pontuacao >= 5):
        return melhor_modelo, [m for _, m in pontuacoes[:5]]
    return None, [m for _, m in pontuacoes[:5]]


def obter_modelos_da_marca(categoria: str, marca_code: str) -> list[dict]:
    """Consulta uma vez cada combinação de categoria e marca."""
    chave = (categoria, marca_code)
    if chave not in MODELOS_EM_CACHE:
        MODELOS_EM_CACHE[chave] = buscar_modelos(categoria, marca_code)
    return MODELOS_EM_CACHE[chave]


print()
resolvidos, ambiguos, sem_candidato = [], [], []

for v in veiculos:
    if not v.get("marca_code"):
        sem_candidato.append(v)
        continue

    try:
        modelos_api = obter_modelos_da_marca(v["categoria_fipe"], v["marca_code"])
    except requests.RequestException as erro:
        print(f"{v['codigo']:8} | ⚠️ erro ao consultar modelos: {erro}")
        sem_candidato.append(v)
        continue

    escolhido, candidatos = encontrar_modelo_generico(v["modelo"], modelos_api)
    if escolhido:
        v["modelo_code"] = escolhido["code"]
        resolvidos.append(v)
    elif candidatos:
        ambiguos.append((v, candidatos))
    else:
        sem_candidato.append(v)

print(f"✅ Resolvidos automaticamente: {len(resolvidos)}")
for v in resolvidos:
    print(f"   {v['codigo']:8} -> modelo_code={v['modelo_code']}")

print(f"\n⚠️  Ainda ambíguos (precisam de revisão manual): {len(ambiguos)}")
for v, candidatos in ambiguos:
    print(f"   {v['codigo']:8} | MODELO={v['modelo']}")
    for c in candidatos:
        print(f"      {c}")

print(f"\n❌ Sem candidato nenhum: {len(sem_candidato)}")
for v in sem_candidato:
    print(f"   {v['codigo']:8} | MODELO={v['modelo']}")

import difflib

from .cliente_api import marcas
from .models import Veiculo
from .normalizacao import (
    chave_marca,
    codigos_tecnicos,
    normalizar_modelo,
    termos_relevantes,
    tracoes,
)


def encontrar_marca_exata(categoria: str, texto_marca: str) -> tuple[str, str] | None:
    alvo = chave_marca(texto_marca)
    encontrados = [marca for marca in marcas(categoria) if chave_marca(marca[1]) == alvo]
    return encontrados[0] if len(encontrados) == 1 else None


def pontuar_modelos(veiculo: Veiculo, modelos_fipe: tuple[tuple[str, str], ...]) -> list[tuple[float, str, str]]:
    """Retorna somente modelos que não contradizem o CRLV.

    Código técnico e tração são filtros eliminatórios. Similaridade textual é
    usada apenas para ordenar os candidatos restantes, nunca para "adivinhar"
    um modelo incompatível.
    """
    texto_origem = normalizar_modelo(f"{veiculo.tipo} {veiculo.modelo_crlv}")
    codigos_origem = codigos_tecnicos(texto_origem)
    tracoes_origem = tracoes(texto_origem)
    termos_origem = termos_relevantes(texto_origem)
    pontuados: list[tuple[float, str, str]] = []

    for modelo_code, nome_modelo in modelos_fipe:
        texto_fipe = normalizar_modelo(nome_modelo)
        codigos_fipe = codigos_tecnicos(texto_fipe)
        tracoes_fipe = tracoes(texto_fipe)

        if codigos_origem and not (codigos_origem & codigos_fipe):
            continue
        if tracoes_origem and (not tracoes_fipe or not (tracoes_origem & tracoes_fipe)):
            continue

        termos_fipe = termos_relevantes(texto_fipe)
        comuns = termos_origem & termos_fipe
        if not codigos_origem and len(comuns) < 2:
            continue

        pontuacao = len(comuns) * 18
        pontuacao += len(codigos_origem & codigos_fipe) * 80
        pontuacao += len(tracoes_origem & tracoes_fipe) * 30
        pontuacao += 20 * difflib.SequenceMatcher(
            None, texto_origem, texto_fipe
        ).ratio()
        pontuados.append((pontuacao, modelo_code, nome_modelo))

    return sorted(pontuados, key=lambda item: item[0], reverse=True)

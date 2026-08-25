from typing import Any

from .cliente_api import (
    anos_do_modelo,
    anos_por_codigo_fipe,
    detalhe_por_codigo_fipe,
    detalhe_por_modelo,
    modelos,
)
from .matching import encontrar_marca_exata, pontuar_modelos
from .models import Veiculo
from .normalizacao import ano_desejado


def anos_compativeis(
    categoria: str, marca_code: str, modelo_code: str, ano: int
) -> list[tuple[str, str]]:
    return [
        item for item in anos_do_modelo(categoria, marca_code, modelo_code)
        if ano_desejado(item[1]) == ano
    ]


def resolver_automaticamente(
    veiculo: Veiculo, marca_code: str
) -> tuple[dict[str, Any] | None, list[str], str]:
    ano = ano_desejado(veiculo.ano)
    if ano is None:
        return None, [], "ANO inválido; informe o ano/modelo com quatro dígitos."

    candidatos_modelo = pontuar_modelos(
        veiculo, modelos(veiculo.categoria, marca_code)
    )
    if not candidatos_modelo:
        return None, [], "Nenhum modelo FIPE compatível com os dados do CRLV."

    melhor_pontuacao = candidatos_modelo[0][0]
    melhores = [item for item in candidatos_modelo if item[0] >= melhor_pontuacao - 18][:8]

    opcoes: list[tuple[str, str, str, str]] = []
    for _, modelo_code, nome_modelo in melhores:
        for ano_code, nome_ano in anos_compativeis(
            veiculo.categoria, marca_code, modelo_code, ano
        ):
            opcoes.append((modelo_code, nome_modelo, ano_code, nome_ano))

    texto_candidatos = [
        f"{modelo_code} | {nome_modelo} | {nome_ano} ({ano_code})"
        for modelo_code, nome_modelo, ano_code, nome_ano in opcoes
    ]
    if len(opcoes) != 1:
        motivo = (
            "Mais de uma combinação modelo/ano compatível; confirme o código FIPE."
            if opcoes else "Nenhuma versão desse modelo está disponível para o ano informado."
        )
        return None, texto_candidatos, motivo

    modelo_code, _, ano_code, _ = opcoes[0]
    detalhe = {
        **detalhe_por_modelo(veiculo.categoria, marca_code, modelo_code, ano_code),
        "_modelo_code": modelo_code,
        "_ano_code": ano_code,
    }
    return detalhe, texto_candidatos, "Correspondência única por marca, modelo e ano."


def confirmar_por_codigo(
    veiculo: Veiculo, codigo_fipe: str, ano_fipe: str
) -> tuple[dict[str, Any] | None, str]:
    ano = ano_desejado(veiculo.ano)
    anos = anos_por_codigo_fipe(veiculo.categoria, codigo_fipe)
    if ano_fipe:
        encontrados = [item for item in anos if item[0] == ano_fipe]
    else:
        encontrados = [item for item in anos if ano_desejado(item[1]) == ano]

    if len(encontrados) != 1:
        return None, "CÓDIGO FIPE informado, mas ANO FIPE não é único/compatível."
    ano_code = encontrados[0][0]
    detalhe = {
        **detalhe_por_codigo_fipe(veiculo.categoria, codigo_fipe, ano_code),
        "_ano_code": ano_code,
    }
    return detalhe, "Atualizado pelo código FIPE confirmado."


def confirmar_por_modelo(
    veiculo: Veiculo, marca_code: str, modelo_code: str, ano_fipe: str
) -> tuple[dict[str, Any] | None, str]:
    """Usa a seleção humana de MODELO CODE FIPE + ANO FIPE com segurança."""
    ano = ano_desejado(veiculo.ano)
    anos = anos_do_modelo(veiculo.categoria, marca_code, modelo_code)
    if ano_fipe:
        encontrados = [item for item in anos if item[0] == ano_fipe]
    else:
        encontrados = [item for item in anos if ano_desejado(item[1]) == ano]

    if len(encontrados) != 1:
        return None, "MODELO CODE FIPE informado, mas ANO FIPE não é único/compatível."

    ano_code = encontrados[0][0]
    detalhe = {
        **detalhe_por_modelo(veiculo.categoria, marca_code, modelo_code, ano_code),
        "_modelo_code": modelo_code,
        "_ano_code": ano_code,
    }
    return detalhe, "Atualizado pela combinação modelo/ano confirmada."

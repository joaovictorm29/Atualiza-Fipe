import re
import unicodedata
from typing import Any

from .config import ALIASES_MARCA, PALAVRAS_NEUTRAS


def normalizar(texto: Any) -> str:
    texto = "" if texto is None else str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto.upper().strip()


def chave_marca(texto: Any) -> str:
    chave = normalizar(texto).replace("-", " ").replace("/", " ")
    chave = re.sub(r"\bVW\b", "", chave)
    chave = re.sub(r"\s+", " ", chave).strip()
    return ALIASES_MARCA.get(chave, chave)


def normalizar_modelo(texto: Any) -> str:
    """Remove placa e uniformiza grafias que não alteram o modelo."""
    texto = normalizar(texto)
    texto = re.sub(r"\([A-Z]{3}-[A-Z0-9]{4}\)", " ", texto)
    texto = re.sub(r"\bCONSTEL\.?\b", "CONSTELLATION", texto)
    texto = re.sub(r"\bROBUSTER\b", "ROBUST", texto)
    texto = re.sub(r"\bCDSRV\b", "CD SRV", texto)
    texto = re.sub(r"\b(\d{1,2})[.]\s*(\d{3})\b", r"\1-\2", texto)
    return re.sub(r"\s+", " ", texto).strip()


def tokens(texto: str) -> set[str]:
    return set(re.findall(r"[A-Z0-9]+(?:[.-][A-Z0-9]+)?", texto))


def codigos_tecnicos(texto: str) -> set[str]:
    return {f"{a}-{b}" for a, b in re.findall(r"\b(\d{1,2})[.-](\d{3})\b", texto)}


def tracoes(texto: str) -> set[str]:
    return set(re.findall(r"\b[468]X[246]\b", texto))


def termos_relevantes(texto: str) -> set[str]:
    resultado = set()
    for termo in tokens(texto):
        if termo in PALAVRAS_NEUTRAS or re.fullmatch(r"\d+(?:[.-]\d+)?", termo):
            continue
        if re.fullmatch(r"\d+P", termo):
            continue
        resultado.add(termo)
    return resultado


def ano_desejado(valor: Any) -> int | None:
    """Em 2025/2026, usa 2026: o ano/modelo que a FIPE normalmente lista."""
    encontrados = re.findall(r"(?:19|20)\d{2}", str(valor))
    return int(encontrados[-1]) if encontrados else None


def preco_para_numero(preco: str) -> float:
    return float(preco.replace("R$", "").replace(".", "").replace(",", ".").strip())

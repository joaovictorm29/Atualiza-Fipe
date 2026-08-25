import os
import time
from functools import lru_cache
from typing import Any

import requests

from .config import API_BASE, TENTATIVAS_API, TIMEOUT_SEGUNDOS
from .models import FipeErro


class ClienteFipe:
    def __init__(self) -> None:
        self.session = requests.Session()
        token = os.getenv("FIPE_API_TOKEN")
        if token:
            self.session.headers["X-Subscription-Token"] = token

    def get(self, caminho: str) -> Any:
        ultimo_erro: Exception | None = None
        for tentativa in range(TENTATIVAS_API):
            try:
                resposta = self.session.get(
                    f"{API_BASE}{caminho}", timeout=TIMEOUT_SEGUNDOS
                )
                resposta.raise_for_status()
                return resposta.json()
            except requests.RequestException as erro:
                ultimo_erro = erro
                if tentativa + 1 < TENTATIVAS_API:
                    time.sleep(2 ** tentativa)
        raise FipeErro(str(ultimo_erro))


cliente = ClienteFipe()


@lru_cache(maxsize=None)
def marcas(categoria: str) -> tuple[tuple[str, str], ...]:
    resposta = cliente.get(f"/{categoria}/brands")
    return tuple((str(item["code"]), item["name"]) for item in resposta)


@lru_cache(maxsize=None)
def modelos(categoria: str, marca_code: str) -> tuple[tuple[str, str], ...]:
    resposta = cliente.get(f"/{categoria}/brands/{marca_code}/models")
    return tuple((str(item["code"]), item["name"]) for item in resposta)


@lru_cache(maxsize=None)
def anos_do_modelo(
    categoria: str, marca_code: str, modelo_code: str
) -> tuple[tuple[str, str], ...]:
    resposta = cliente.get(
        f"/{categoria}/brands/{marca_code}/models/{modelo_code}/years"
    )
    return tuple((str(item["code"]), item["name"]) for item in resposta)


@lru_cache(maxsize=None)
def detalhe_por_modelo(
    categoria: str, marca_code: str, modelo_code: str, ano_code: str
) -> dict[str, Any]:
    return cliente.get(
        f"/{categoria}/brands/{marca_code}/models/{modelo_code}/years/{ano_code}"
    )


@lru_cache(maxsize=None)
def anos_por_codigo_fipe(categoria: str, codigo_fipe: str) -> tuple[tuple[str, str], ...]:
    resposta = cliente.get(f"/{categoria}/{codigo_fipe}/years")
    return tuple((str(item["code"]), item["name"]) for item in resposta)


@lru_cache(maxsize=None)
def detalhe_por_codigo_fipe(
    categoria: str, codigo_fipe: str, ano_code: str
) -> dict[str, Any]:
    return cliente.get(f"/{categoria}/{codigo_fipe}/years/{ano_code}")

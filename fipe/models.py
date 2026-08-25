from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class Veiculo:
    linha: int
    codigo_ativo: str
    tipo: str
    modelo_crlv: str
    marca: str
    categoria: str
    ano: Any


class FipeErro(RuntimeError):
    pass

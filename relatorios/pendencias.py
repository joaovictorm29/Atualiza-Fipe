import csv
from pathlib import Path
from typing import Any


def salvar_pendencias(
    caminho: Path, pendencias: list[dict[str, Any]]
) -> None:
    with caminho.open("w", newline="", encoding="utf-8-sig") as arquivo:
        campos = ["linha", "codigo_ativo", "marca", "modelo_crlv", "ano", "motivo", "candidatos"]
        escritor = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
        escritor.writeheader()
        escritor.writerows(pendencias)

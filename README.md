# Atualiza FIPE

Automação em Python para reduzir uma atividade repetitiva de trabalho: analisar os veículos de uma planilha, consultar a API da FIPE usando marca, modelo e ano e preencher o valor da tabela FIPE correspondente.

O programa tenta encontrar uma correspondência segura para cada veículo. Quando existe mais de uma possibilidade, quando os dados estão incompletos ou quando a API não encontra uma combinação compatível, o registro não é escolhido automaticamente: ele é marcado para revisão e incluído em um relatório de pendências.

## Como funciona

Para cada linha válida da planilha, a automação:

1. Lê marca, modelo, tipo, categoria e ano do veículo.
2. Consulta marcas, modelos e anos disponíveis na API FIPE.
3. Compara os dados informados com os modelos retornados, considerando códigos técnicos, tração e termos relevantes.
4. Preenche o valor FIPE e os dados de controle quando encontra uma única correspondência.
5. Registra pendências, candidatos e motivos para conferência manual quando não há segurança suficiente.

As consultas são feitas pela [API FIPE Parallelum](https://fipe.parallelum.com.br/).

## Requisitos

- Windows, macOS ou Linux;
- Python 3.10 ou superior;
- acesso à internet para consultar a API;
- uma planilha Excel `.xlsx` no formato esperado.

As dependências externas do projeto são:

- `openpyxl`: leitura e atualização de arquivos Excel;
- `requests`: requisições HTTP para a API FIPE.

Os módulos `os`, `time`, `csv`, `pathlib`, `typing`, `dataclasses`, `functools`, `difflib`, `re` e `unicodedata` pertencem à biblioteca padrão do Python e não precisam ser instalados separadamente.

## Clonar o repositório

No terminal, execute:

```bash
git clone <URL_DO_REPOSITORIO>
cd Atualiza-Fipe
```

Substitua `<URL_DO_REPOSITORIO>` pela URL deste repositório.

## Instalar as dependências

É recomendado criar um ambiente virtual para manter as dependências isoladas:

```bash
python -m venv .venv
```

No Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

No Windows pelo Prompt de Comando:

```bat
.venv\Scripts\activate
```

No macOS ou Linux:

```bash
source .venv/bin/activate
```

Depois, instale os pacotes listados em `requirements.txt`:

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Preparar a planilha

A planilha deve conter:

- uma aba chamada `VEICULOS`;
- os cabeçalhos na linha 2;
- os dados dos veículos a partir da linha 3;
- as colunas obrigatórias `CÓD ATIVO`, `TIPO`, `MODELO`, `MARCA`, `CATEGORIA FIPE`, `ANO` e `VALOR FIPE CHASSI`.

Na coluna `CATEGORIA FIPE`, utilize exatamente `cars` para carros ou `trucks` para caminhões. A coluna `CÓD ATIVO` identifica o veículo; linhas sem esse código e linhas que contenham `TOTAL` são ignoradas.

Os dados de `MARCA`, `MODELO` e `ANO` precisam ser explícitos e estar de acordo com a nomenclatura apresentada no site da Tabela FIPE. Informe o modelo completo, incluindo versão, motorização, cabine, tração e demais características relevantes quando existirem. Quanto mais fiel for o preenchimento ao cadastro FIPE e ao documento do veículo, maior será a precisão da correspondência, podendo chegar a quase 100% em dados completos e consistentes. A automação não substitui a conferência humana nos casos ambíguos.

Antes de executar, faça uma cópia de segurança da planilha: o arquivo original é salvo novamente no mesmo caminho após a atualização.

## Usar outra planilha

Abra `fipe/config.py` e altere a constante `CAMINHO_PLANILHA` para o caminho completo do arquivo que deseja processar:

```python
CAMINHO_PLANILHA = Path(r"C:\caminho\para\minha_planilha.xlsx")
```

O prefixo `r` evita problemas com as barras invertidas do Windows. Também é possível usar barras normais:

```python
CAMINHO_PLANILHA = Path("C:/caminho/para/minha_planilha.xlsx")
```

Confira também, no mesmo arquivo, se estes valores correspondem à sua planilha:

```python
CAMINHO_PENDENCIAS = Path(r"C:\caminho\para\pendencias_fipe.csv")
ABA_VEICULOS = "VEICULOS"
LINHA_CABECALHO = 2
LINHA_INICIO_DADOS = 3
```

Se a aba, a linha dos cabeçalhos ou a primeira linha de dados forem diferentes, ajuste as constantes correspondentes. Os nomes das colunas obrigatórias continuam sendo necessários.

## Executar

Com o ambiente virtual ativado e estando na pasta do projeto, execute:

```bash
python main.py
```

Ao final, o terminal informa quantos veículos foram atualizados e quantos ficaram para revisão. A planilha configurada recebe os valores e as colunas de controle:

- `VALOR FIPE CHASSI`;
- `CÓDIGO FIPE`;
- `ANO FIPE`;
- `MODELO FIPE`;
- `REFERÊNCIA FIPE`;
- `STATUS FIPE`;
- `MOTIVO FIPE`;
- `CANDIDATOS FIPE`.

As pendências são salvas em `CAMINHO_PENDENCIAS`, por padrão em `pendencias_fipe.csv`, usando `;` como separador. Registros com `REVISÃO NECESSÁRIA`, `DADO INCOMPLETO` ou `ERRO DE API` devem ser verificados manualmente.

## Token da API

O projeto funciona sem token quando a API permite a consulta. Caso você possua um token de assinatura, defina a variável de ambiente `FIPE_API_TOKEN` antes da execução.

No Windows PowerShell:

```powershell
$env:FIPE_API_TOKEN = "SEU_TOKEN"
python main.py
```

No macOS ou Linux:

```bash
export FIPE_API_TOKEN="SEU_TOKEN"
python main.py
```

Não coloque tokens diretamente no código nem os versione no repositório.

## Estrutura principal

```text
main.py                    # ponto de entrada
fipe/config.py             # caminhos e parâmetros da planilha/API
fipe/cliente_api.py        # cliente e consultas à API FIPE
fipe/matching.py           # comparação de modelos
fipe/processamento.py      # fluxo de leitura e atualização
fipe/resolver.py           # confirmação de modelo e ano
relatorios/pendencias.py   # geração do CSV de pendências
```

## Autor e uso de IA

Projeto desenvolvido por **João Victor Maciel Chaves**.

A Inteligência Artificial foi utilizada como ferramenta de otimização de tempo durante a conclusão do projeto, auxiliando no desenvolvimento, na organização e na revisão da automação. A responsabilidade pela configuração, validação dos dados e uso dos resultados permanece com o usuário.

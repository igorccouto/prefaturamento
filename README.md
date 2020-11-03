# Pre-Faturamento

Cria a Planilha de Pré-Faturamento mensal de forma automática. O arquivo de entrada deve ser o Relatório 5.1 do Remedy.

## Instalação

Para instalar, basta baixar o repositório, instalar as dependências e criar o executável.

```console
$ git clone <REPOSITORIO>
$ cd <REPOSITORIO>
$ pip install -r requirements.txt
$ pyinstaller --onefile --clean --windowed --name CTR_017-2016_PreFaturamento --log-level WARNING prefaturamento.py
```

O executável estará dentro da pasta *dist*.
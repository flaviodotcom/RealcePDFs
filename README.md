<p align="center">
  <img height="150" src="resources/images/vigarista.ico">
</p>

<h1 align="center">Destacar PDFs por Matrícula</h1>
<br>
Este é um programa executável que permite destacar matrículas de funcionários em um arquivo PDF com base em informações de uma planilha do Excel. Ele é especialmente útil para realçar benefícios como Seguro de Vida, Plano Odontológico, Vale Transporte, Vale Alimentação e Vale Refeição em documentos PDF.

## Requerimentos

### `Python 3.11.0`

### Configuração de Virtual Environment

Para executar o programa pelo Python, Você pode configurar um ambiente virtual usando `virtualenv`. Siga os passos
abaixo:

```bash
pip install virtualenv
python3.11 -m venv <virtual-environment-name>
source env/bin/activate
```

### Instalação de Dependências

Instale as dependências necessárias usando o `pip`:

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

## Executável

O projeto oferece um arquivo executável, possibilitando a utilização do programa em computadores que não possuem o
interpretador Python instalado.

## Como Usar

**Selecione os Arquivos**: Clique nos botões "Selecionar" para escolher os arquivos Excel e PDF correspondentes.

**Formato do Arquivo Excel**: Certifique-se de que as matrículas estejam na coluna B da planilha. O programa analisará essa
coluna em busca das matrículas. Não deve haver outras informações nessa coluna.

**Nomes dos Colaboradores**: Para identificar as matrículas, o programa usa os nomes dos colaboradores na coluna C da
planilha. Mantenha os nomes dos colaboradores nessa coluna.

**Arquivo de Texto**: O programa cria um arquivo de texto que lista as matrículas não encontradas, juntamente com os nomes
dos colaboradores. Esse arquivo será salvo no mesmo diretório do PDF editado.

**Salvar**: Clique no botão "Salvar" para salvar o PDF editado na Área de Trabalho (Desktop), dentro da pasta "BENEFÍCIOS
DESTACADOS". Você também pode clicar em "Salvar Como" para escolher um local de armazenamento personalizado.

## Funcionalidades Adicionais

### Separar PDFs por Matrícula

Este programa também permite separar o PDF em vários arquivos, um para cada matrícula encontrada. Os arquivos serão
salvos em uma pasta de sua escolha.

> Requer a seleção de um arquivo Excel (.xlsx, .csv ou .xls) e um arquivo PDF para funcionar corretamente.

## Como Funciona?

O programa utiliza bibliotecas como PyPDF2, PyPDF4 e openpyxl para realizar as seguintes tarefas:

Abre o arquivo PDF e a planilha do Excel selecionados.
Percorre as páginas do PDF em busca das matrículas e realça os números encontrados.
Cria um arquivo de texto com as matrículas não encontradas.
Salva o PDF editado na Área de Trabalho (Desktop) em uma pasta chamada "BENEFÍCIOS DESTACADOS".
Oferece a opção de salvar o PDF em um local personalizado.
Permite separar o PDF em arquivos individuais para cada matrícula encontrada.

## Aviso

- Certifique-se de seguir as orientações acima para garantir o funcionamento adequado do programa.


- Em alguns casos, a biblioteca `PymuPDF` pode causar alguns problemas de importação. Nesse caso, execute no terminal o
  comando `pip install --upgrade --force-reinstall pymupdf`


<p align="center">
  <img height="150" src="resources/images/vigarista.ico">
</p>

<h1 align="center">Destacar PDFs por Matrícula</h1>
<br>
Este é um programa executável que permite destacar matrículas de funcionários em um arquivo PDF com base em informações de uma planilha do Excel. Ele é especialmente útil para realçar benefícios como Seguro de Vida, Plano Odontológico, Vale Transporte, Vale Alimentação e Vale Refeição em documentos PDF. Suas vantagens são a velocidade de processamento e interface amigável.

## Executável

O projeto disponibiliza um arquivo executável, viabilizando a utilização do programa em computadores sem o
interpretador Python instalado. Se você fez adaptações no código e deseja criar um executável a partir dele, basta
executar o script [`build.py`][build.py].

**Download**:

- [`RealcePDFs.exe`][latest-release] (64-bit)

[latest-release]: https://github.com/flaviodotcom/RealcePDFs/tree/main/exe

[build.py]: https://github.com/flaviodotcom/RealcePDFs/blob/main/build.py

## Como Usar

**Selecione os Arquivos**: Clique nos botões "Selecionar" para escolher os arquivos Excel e PDF correspondentes.

**Formato do Arquivo Excel**: Certifique-se de que as matrículas estejam na coluna B da planilha. O programa analisará
essa
coluna em busca das matrículas. Não deve haver outras informações nessa coluna.

**Nomes dos Colaboradores**: Para identificar as matrículas, o programa usa os nomes dos colaboradores na coluna C da
planilha. Mantenha os nomes dos colaboradores nessa coluna.

**Arquivo de Texto**: O programa cria um arquivo de texto que lista as matrículas não encontradas, juntamente com os
nomes
dos colaboradores. Esse arquivo será salvo no mesmo diretório do PDF editado.

**Salvar**: Clique no botão "Salvar" para salvar o PDF editado na Área de Trabalho (Desktop), dentro da pasta "
BENEFÍCIOS
DESTACADOS". Você também pode clicar em "Salvar Como" para escolher um local de armazenamento personalizado.

**Separar PDFs por Matrícula**: Clique no botão 'Separar PDFs' para dividir o PDF em vários arquivos, um para cada
matrícula encontrada. Esta função também destaca as matrículas nos PDFs. Os arquivos serão salvos na pasta de sua
escolha.

> Requer a seleção de um arquivo Excel (.xlsx) e um arquivo PDF para funcionar corretamente.

## Como Funciona?

O programa utiliza bibliotecas como PyMuPDF, pypdf e openpyxl para realizar as seguintes tarefas:

Abre o arquivo PDF e a planilha do Excel selecionados.
Percorre as páginas do PDF em busca das matrículas e realça os números encontrados.
Cria um arquivo de texto com as matrículas não encontradas.
Salva o PDF editado na Área de Trabalho (Desktop) em uma pasta chamada "BENEFÍCIOS DESTACADOS".
Oferece a opção de salvar o PDF em um local personalizado.
Permite separar o PDF em arquivos individuais para cada matrícula encontrada.

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

### Testes

Para executar os testes, execute o comando abaixo:

```bash
python -m unittest discover -s ./tests -p 'test_*.py'
```

ou

```bash
pip install pytest
python -m pytest
```

### Tags de Commits

As seguintes tags são usadas como padrão de projeto no prefixo dos commits:

- `[FIX]` para correções de bugs: principalmente usado em versões estáveis, mas também válido se você estiver corrigindo
  um bug recente na versão de desenvolvimento;
- `[REF]` para refatoração: quando uma funcionalidade é totalmente reescrita;
- `[ADD]` para adicionar novos módulos;
- `[REM]` para remover recursos: remover código morto, vistas, módulos, ...;
- `[REV]` para reverter commits: se um commit causa problemas ou não é desejado, ele é revertido usando esta tag;
- `[MOV]` para mover arquivos: use `git move` e não altere o conteúdo do arquivo movido, caso contrário, o Git pode
  perder o controle e o histórico do arquivo; também usado ao mover código de um arquivo para outro;
- `[REL]` para commits de versão: novas versões estáveis principais ou menores;
- `[IMP]` para melhorias: a maioria das alterações feitas na versão de desenvolvimento são melhorias incrementais não
  relacionadas a outra tag;
- `[MERGE]` para commits de merge: usado na portabilidade direta de correções de bugs, mas também como commit principal
  para recursos envolvendo vários commits separados;
- `[CLA]` para assinar a Licença de Contribuidor Individual;
- `[I18N]` para alterações em arquivos de tradução.

## Aviso

- Certifique-se de seguir as orientações acima para garantir o funcionamento adequado do programa.


- Em alguns casos, a biblioteca `PymuPDF` pode causar alguns problemas de importação. Nesse caso, execute no terminal o
  comando `pip install --upgrade --force-reinstall pymupdf`


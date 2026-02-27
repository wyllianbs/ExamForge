# 🔨 ExamForge — Sorteio de Questões $\LaTeX$ via Planilha ODS

Código em Python para sortear questões de provas em formato $\LaTeX$, com configuração centralizada em planilha `.ods`, controle de repetição entre etapas e suporte a múltiplos grupos e tipos de prova.


## 📋 Descrição

O **ExamForge** lê uma planilha `.ods` que descreve a estrutura da prova (grupos de questões, quantidade a sortear por grupo, etc.) e extrai questões de um banco de arquivos `.tex`. O sorteio prioriza questões **inéditas** — nunca sorteadas em etapas anteriores — consultando um histórico de logs. Quando o banco de inéditas se esgota, o sistema oferece reutilizar as questões mais antigas, com confirmação do usuário.

O sistema foi projetado para funcionar em conjunto com o _template_ LaTeX Para Provas Com Gabarito, disponível em [https://github.com/wyllianbs/carderno_prova](https://github.com/wyllianbs/carderno_prova).


## ✨ Características

- ✅ **Configuração via planilha ODS** — grupos, arquivos e quantidades definidos em `.ods`.
- ✅ **Prioridade a questões inéditas** — questões já sorteadas são usadas apenas quando necessário.
- ✅ **Controle de histórico** — logs `.log` por etapa registram os IDs sorteados e evitam repetição.
- ✅ **Round-robin entre arquivos** — sorteio distribuído entre os arquivos de um mesmo grupo.
- ✅ **Distribuição automática de `k`** — quando a coluna `k` não está preenchida na planilha, distribui proporcionalmente o total solicitado pelo usuário entre os grupos.
- ✅ **Verificação de contagens** — confronta o número de questões declarado no ODS com o encontrado nos `.tex`, interrompendo a execução em caso de divergência.
- ✅ **Suporte a múltiplos tipos** — diferentes categorias de questões (`P1`, `P2`, etc.) podem ser mapeadas para diretórios distintos.
- ✅ **Não modifica** os arquivos originais (apenas os lê).
- ✅ **Arquitetura POO** (Orientação a Objetos).


## 📁 Estrutura do Projeto

```
.
├── ExamForge.py               # Script principal
├── db.ods                     # Planilha ODS de configuração
├── P1/                        # Diretório com banco de questões tipo P1 (.tex)
│   ├── P1_py__misc.tex
│   ├── P1_programming_misc.tex
│   ├── P1_py_arrays_data_structure.tex
│   ├── P1_py_arrays_list_1D.tex
│   └── ...
├── P2/                        # Diretório com banco de questões tipo P2 (.tex)
│   └── ...
├── logs/                      # Logs de sorteios anteriores (gerados automaticamente)
│   └── P1/
│       ├── P1-2025-2.log
│       └── P1-2026-1.log
├── P1.tex                     # Arquivo de saída — questões sorteadas (gerado)
└── README.md                  # Este arquivo
```


## 🚀 Instalação

### Requisitos

- **Python 3.8+**
- **Linux** (testado no SO Linux, distro Debian Trixie).
- Biblioteca [`odfpy`](https://pypi.org/project/odfpy/) para leitura de arquivos `.ods`:

```bash
pip install odfpy
```

### Clone o repositório

```bash
git clone https://github.com/wyllianbs/ExamForge.git
cd ExamForge
```


## 📖 Como Usar

### Execução Básica

```bash
python3 ExamForge.py
```

### Fluxo de Uso

1. **Arquivo ODS**: Informe o caminho da planilha de configuração (default: `db.ods`).
2. **Planilha**: Escolha a configuração desejada — `P1`, `P1C2`, `P2`, `P2C2` ou `Rec` (default: `P1`).
3. **Diretórios**: Informe os diretórios onde estão os arquivos `.tex` correspondentes a cada tipo.
4. **Modo aleatório (opcional)**: Se a coluna `k` não estiver preenchida, o sistema pergunta quantas questões sortear ao todo.
5. **Arquivo de saída**: Informe o caminho do `.tex` de saída (default: `<planilha>.tex`).
6. **Arquivo de log**: Informe o caminho do log a ser gerado (default: `logs/<planilha>/<planilha>-<ano>-<semestre>.log`).

### Exemplo de Execução

```
python3 ExamForge.py

Informe o caminho do arquivo ODS
[Enter = db.ods]:

Qual configuração deseja ler?
Opções: P1, P1C2, P2, P2C2, Rec
[Enter = P1]:

── Lendo Configuração ─────────────────────────────────────────────────────
  Planilha: P1   Arquivo: db.ods

Types encontrados em 'P1': P1
Informe o(s) diretório(s) dos arquivos (separados por espaço)
[Enter = P1]:

── Mapeamento Type → Diretório ────────────────────────────────────────────
    •  P1  →  P1

── Verificação de Arquivos ────────────────────────────────────────────────
  Diretórios: P1
  ✔  Todos os arquivos foram encontrados.

── Verificação de Contagens (ODS vs .tex) ─────────────────────────────────
  ✔  Todas as contagens coincidem.

════════════════════════════════════════════════════════════
  Estrutura de questões extraídas
════════════════════════════════════════════════════════════
  Total de arquivos : 14
  Total de questões : 14
════════════════════════════════════════════════════════════

  P1 │ P1/P1_py__misc.tex ........................ : 1
  P1 │ P1/P1_programming_misc.tex ................ : 1
  P1 │ P1/P1_py_arrays_data_structure.tex ........ : 1
  P1 │ P1/P1_py_arrays_data_structure_Claude.tex . : 1
  P1 │ P1/P1_py_arrays_list_1D.tex ............... : 1
  P1 │ P1/P1_py_arrays_list_slicing.tex .......... : 1
  P1 │ P1/P1_py_arrays_dict.tex .................. : 1
  P1 │ P1/P1_py_arrays_tuple.tex ................. : 1
  P1 │ P1/P1_py_arrays_set.tex ................... : 1
  P1 │ P1/P1_py_strings.tex ...................... : 1
  P1 │ P1/P1_py_arrays_dict_comprehensions.tex ... : 1
  P1 │ P1/P1_py_arrays_list_comprehension.tex .... : 1
  P1 │ P1/P1_py_arrays_generator_comprehension.tex : 1
  P1 │ P1/P1_py_arrays_methods.tex ............... : 1
  ──────────────────────────────────────────────────────
  Total .......................................... : 14

── Histórico de Sorteios ──────────────────────────────────────────────────
  Diretório: logs/P1
  ℹ  Diretório 'logs/P1' não encontrado.
     Deseja criá-lo recursivamente? [S/n]:
  ✔  Diretório 'logs/P1' criado.
  Nenhum histórico de sorteios anteriores será usado.

Informe o caminho do arquivo de saída (.tex)
[Enter = P1.tex]:

Informe o caminho do arquivo de log
[Enter = logs/P1/P1-2026-1.log]:

── Sorteando Questões ─────────────────────────────────────────────────────

  ✔  Arquivo salvo: P1.tex  (6 questões)
  ✔  Log salvo: logs/P1/P1-2026-1.log  (6 IDs)

═══════════════════════════════════════════════════════════════════════════
  Resumo do Sorteio
───────────────────────────────────────────────────────────────────────────
  Questões sorteadas  →  6
  Arquivo .tex        →  P1.tex
  Arquivo .log        →  logs/P1/P1-2026-1.log
═══════════════════════════════════════════════════════════════════════════
```


## 📊 Formato da Planilha ODS

A planilha possui uma aba por configuração de prova (`P1`, `P2`, etc.) e as seguintes colunas:

| Coluna | Descrição |
|--------|-----------|
| `Type` | Categoria das questões (ex.: `P1`, `P2`). Define o diretório de busca. |
| `Group` | ID numérico do grupo (tema/tópico). Questões do mesmo grupo são sorteadas em round-robin. |
| `File` | Nome do arquivo `.tex` (sem o caminho do diretório). |
| `n` | Número de questões naquele arquivo (validado contra o `.tex`). |
| `g` | Número total de questões do grupo (soma dos `n` dos arquivos). Preenchido apenas na primeira linha do grupo. |
| `k` | Quantidade de questões a sortear daquele grupo. Preenchido apenas na primeira linha. `0` ou vazio ativa distribuição automática. |
| `Sum` | Soma acumulada de questões sorteadas até aquele grupo (auxiliar). |

A **primeira linha** de cada grupo contém `g`, `k` e `Sum`. As linhas seguintes (arquivos adicionais do mesmo grupo) contêm apenas `Type`, `Group`, `File` e `n`. Abaixo, o conteúdo real do arquivo `db.ods` (aba `P1`):

```
Type  Group  File                                      n  g  k  Sum
P1        0  P1_py__misc.tex                           1  2  1    1
P1        0  P1_programming_misc.tex                   1
P1        1  P1_py_arrays_data_structure.tex           1  2  1    2
P1        1  P1_py_arrays_data_structure_Claude.tex    1
P1        2  P1_py_arrays_list_1D.tex                  1  5  1    3
P1        4  P1_py_arrays_list_slicing.tex             1
P1        5  P1_py_arrays_dict.tex                     1
P1        6  P1_py_arrays_tuple.tex                    1
P1        7  P1_py_arrays_set.tex                      1
P1       11  P1_py_strings.tex                         1  1  1    4
P1       13  P1_py_arrays_dict_comprehensions.tex      1  3  1    5
P1       14  P1_py_arrays_list_comprehension.tex       1
P1       17  P1_py_arrays_generator_comprehension.tex  1
P1       18  P1_py_arrays_methods.tex                  1  1  1    6
```

> **Nota sobre os IDs de grupo**: Os valores da coluna `Group` são identificadores de tópico/tema do banco de questões, não índices sequenciais. Grupos com múltiplos arquivos têm `g` e `k` definidos apenas na sua primeira linha; as demais linhas do grupo são linhas de continuação.


## 📝 Formato das Questões `.tex`

Cada questão é delimitada por um bloco `{% ID ... }` que contenha o marcador `\rtask`. O ID no comentário de abertura é usado pelo sistema de logs para rastrear cada questão individualmente entre os sorteios.

### Questão de Múltipla Escolha

```latex
{% Q1383489[49D]
\needspace{8\baselineskip}
\item \rtask \ponto{\pt}
Assinale a opção abaixo que contém SOMENTE informações CORRETAS.

\begin{answerlist}[label={\texttt{\Alph*}.},leftmargin=*]
    \ti Python 3 possui retrocompatibilidade total com Python 2.
    \ti Python 3 não é compatível com strings Unicode.
    \ti \lstinline[style=Python]|count(d)| retorna o número de elementos do dict \lstinline[style=Python]|d|.
    \di Dicionários em Python 3 preservam a ordem de inserção.
    \ti Utiliza-se \lstinline[style=Python]|array.add(x)| para adicionar x a array.
\end{answerlist}
}
```

- `\ti` — alternativa **incorreta**.
- `\di` — alternativa **correta** (gabarito).

### Questão Verdadeiro/Falso

```latex
{% Q3258082
\needspace{7\baselineskip}
\item \rtask \ponto{\pt}
Julgue o próximo item.

Em Python, listas podem ser preenchidas por qualquer tipo de objeto, porém
a quantidade de objetos só poderá ser alterada durante a criação delas.

% F
{\setlength{\columnsep}{0pt}\renewcommand{\columnseprule}{0pt}
\begin{multicols}{2}
\begin{answerlist}[label={\texttt{\Alph*}.},leftmargin=*]
    \ti[V.]
    \ifnum\gabarito=1\doneitem[F.]\else\ti[F.]\fi % gabarito
\end{answerlist}
\end{multicols}
}
}
```

O gabarito (`V.` ou `F.`) é controlado pelo condicional `\ifnum\gabarito=1` do _template_ LaTeX, que exibe a alternativa correta destacada ao compilar a versão com gabarito.

### Formato do ID

Os IDs seguem o padrão `Q<número>` com sufixos opcionais que indicam origem ou metadados da questão, como `Q1383489[49D]` ou `Q535634[24B]`. O ExamForge trata o ID como uma string opaca — apenas o utiliza para rastrear questões nos logs.


## 🎯 Lógica de Sorteio

### Prioridade a questões inéditas
O sistema consulta todos os logs anteriores para identificar quais questões já foram sorteadas. As questões **nunca sorteadas** têm prioridade absoluta no sorteio.

### Round-robin entre arquivos de um grupo
Quando um grupo possui múltiplos arquivos `.tex` (como o Grupo 0, com `P1_py__misc.tex` e `P1_programming_misc.tex`), as questões são distribuídas em ordem circular entre os arquivos, garantindo equilíbrio na seleção.

### Reutilização com confirmação
Se um grupo não tiver questões inéditas suficientes para atingir o `k` configurado, o sistema identifica as questões mais antigas (dos logs mais antigos) e pergunta ao usuário antes de reutilizá-las.

### Distribuição automática de `k`
Se a coluna `k` não estiver preenchida na planilha, o sistema solicita o total desejado e distribui o valor proporcionalmente entre os grupos, ponderando pelo número de questões inéditas disponíveis em cada um.


## 🏗️ Arquitetura (POO)

O projeto utiliza Programação Orientada a Objetos com as seguintes classes:

| Classe | Responsabilidade |
|--------|------------------|
| `Question` | Representa uma questão extraída de um arquivo `.tex` |
| `GroupMeta` | Metadados de posição de um arquivo dentro de seu grupo |
| `FileEntry` | Questões extraídas de um único arquivo `.tex` |
| `GroupConfig` | Configuração de um grupo lida da planilha ODS |
| `DrawRecord` | Registro de um sorteio anterior (lido de arquivo `.log`) |
| `DrawnQuestion` | Uma questão efetivamente sorteada |
| `Printer` | Utilitário de formatação padronizada no terminal |
| `OdsReader` | Lê e interpreta arquivos `.ods` |
| `GroupParser` | Converte linhas brutas do ODS em `GroupConfig` |
| `FileManager` | Gerencia caminhos e verificação de arquivos `.tex` |
| `TexParser` | Extrai questões de arquivos `.tex` via rastreamento de chaves |
| `EntryBuilder` | Constrói `FileEntry` a partir dos dados brutos de extração |
| `LogManager` | Lê e grava logs de sorteios (`.log`) |
| `QuestionDrawer` | Sorteia questões respeitando prioridades e round-robin |
| `TexWriter` | Grava questões sorteadas em arquivo `.tex` |
| `UserInterface` | Interface com o usuário (entrada/saída no terminal) |
| `ExamForge` | Classe principal (Facade) que orquestra todo o fluxo |


## 🔧 Integração com $\LaTeX$

Para compilar o arquivo gerado e criar a prova final:

1. Clone o _template_:

```bash
git clone https://github.com/wyllianbs/caderno_prova.git
```

2. Copie o `.tex` gerado para o diretório do _template_ e inclua-o no arquivo principal:

```latex
\input{P1.tex}
```

3. Compile com `pdflatex` ou `xelatex`:

```bash
pdflatex main.tex
```


## 🐛 Tratamento de Erros

O programa valida:

- ✅ Existência e leitura do arquivo `.ods`.
- ✅ Nome da planilha (aba) informado pelo usuário.
- ✅ Existência de todos os arquivos `.tex` declarados na planilha.
- ✅ Contagem de questões: ODS (`n`) vs. número real encontrado nos `.tex`.
- ✅ Disponibilidade de questões inéditas — avisa e pede confirmação antes de reutilizar.
- ✅ Erros de I/O na leitura de logs e na gravação dos arquivos de saída.


## 📜 Licença

Este projeto está licenciado sob a Licença [GNU General Public License v3.0](LICENSE).


## 👤 Autor

**Prof. Wyllian B. da Silva**  
Universidade Federal de Santa Catarina (UFSC)  
Departamento de Informática e Estatística (INE)


---

**Nota**: Este projeto foi desenvolvido especificamente para uso na UFSC, mas pode ser facilmente adaptado para outras instituições de ensino e outros formatos de prova.

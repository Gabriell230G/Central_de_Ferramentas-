# Auto Encerramento Plus (Auto_Encerramento_Plus_novo)

Robô que encerra processos no GCPJ e lança pagamentos complementares, a partir de uma base Access com os processos que devem ser tratados.

| Item | Valor |
|---|---|
| Código no inventário | **C@KJ** (Auto Encerramento) |
| Área | 4040 Dejur / COLDA - Collab Dev's & Analytics |
| Linguagem | C# (Windows Forms) |
| Framework | .NET Framework 4.8 |
| Sistema alvo | GCPJ (Internet Explorer, via SHDocVw / mshtml) |
| Base local | `BD_DADOS.accdb` (mesma pasta do executável) |
| Controle S5 | Implementado (limite de **5.000** processos por execução) |

> Este repositório tem o mesmo código do `Auto_Encerramento_Plus_novo_Parecer` (C@KM). O Controle S5 foi aplicado igual nos dois.

---

## Sumário

1. [O que o sistema faz](#1-o-que-o-sistema-faz)
2. [Como rodar](#2-como-rodar)
3. [Estrutura do projeto](#3-estrutura-do-projeto)
4. [Controle S5: Limitador de Operações](#4-controle-s5-limitador-de-operações)
5. [Guia rápido: quero mudar X, onde vou?](#5-guia-rápido-quero-mudar-x-onde-vou)
6. [Pontos de atenção](#6-pontos-de-atenção)
7. [Histórico](#7-histórico)

---

## 1. O que o sistema faz

Na tela, o usuário escolhe o **tipo de processamento**, informa a **máquina** e clica em **Processar**.

| Opção | Rotina | Tabela | O que faz no GCPJ |
|---|---|---|---|
| **ENCERRAMENTO** | `encerramento()` | `TB_ENCERRAMENTO` | Encerra o processo (motivo, classificação, envolvidos, depósitos e IR) |
| **ACORDOS** | `encerramento()` | `TB_ACORDOS` | Mesmo fluxo, para processos de acordo |
| **COMPLEMENTAR** | `complementar()` | `TB_COMPLEMENTAR` | Lança o pagamento avulso complementar (depósito, finalidade, forma de pagamento e composição) |

```
Escolher a opção ➜ informar a máquina ➜ Processar
   ➜ lê da base os processos pendentes daquela máquina
   ➜ para cada processo: abre a tela do GCPJ, preenche e salva
   ➜ marca o processo como feito na base (ENCERRADO / PROCESSADO = TRUE)
   ➜ termina quando acabam os processos OU quando atinge o limite do Controle S5
```

Os botões **Distribuir** dividem os processos pendentes entre várias máquinas (coluna `MAQUINA`).

## 2. Como rodar

1. Abra `Auto_Encerramento_Plus.sln` no Visual Studio.
2. Instale o **Microsoft Access Database Engine** (ACE OLEDB 12.0).
3. Coloque o `BD_DADOS.accdb` na mesma pasta do executável.
4. Rode, conecte no GCPJ, escolha a opção, informe a máquina e clique em **Processar**.

## 3. Estrutura do projeto

| Arquivo | Para que serve |
|---|---|
| `Form1.cs` | Tela e as rotinas de automação (`encerramento`, `complementar`, `solicitarEncerramento`, `atualizaEvento`) |
| `Form1.Designer.cs` | Layout da tela |
| `Automacao.cs` | Funções que controlam o Internet Explorer (navegar, preencher, clicar, ler mensagens) |
| `Conexao.cs` | Conexão com o `BD_DADOS.accdb` |
| `ClsVersion.cs` | Versão exibida na barra de status |
| `LimitadorS5.cs` | **Controle S5**: limite de operações por execução |

## 4. Controle S5: Limitador de Operações

Exigência do normativo **05.1073** para aplicativos com C10 = Sim (automação que captura ou atualiza dados em sistema estruturado).

### Como funciona

```
Clique em Processar ➜ contador = 0
   Para cada processo da base:
      contador já chegou no limite?  ➜ SIM: avisa o usuário e para
                                      ➜ NÃO: soma 1 e processa normalmente
```

Ao atingir o limite aparece a mensagem **"Processamento Limitado - Controle S5"** e a grade registra *"Controle S5: limite de 5.000 registros atingido. Processamento interrompido."*. Os processos que sobraram continuam pendentes na base e entram na próxima execução.

### Onde mexer

| O que | Arquivo | Onde |
|---|---|---|
| **Valor do limite** | `LimitadorS5.cs` | constante `LIMITE_OPERACOES` (linha 27) |
| Mensagem exibida ao usuário | `LimitadorS5.cs` | método `LimiteAtingido` |
| ENCERRAMENTO e ACORDOS | `Form1.cs` | `encerramento()`, bloco `[S5]` na linha 1124 |
| COMPLEMENTAR | `Form1.cs` | `complementar()`, bloco `[S5]` na linha 1490 |

> Dica: **Ctrl+Shift+F** e pesquise `[S5]`.

Detalhes que importam:

- Em `encerramento()` o bloco fica **antes** do rótulo `reinicia:`. Quando o robô tenta de novo o mesmo processo (`goto reinicia`), não conta duas vezes.
- Em `complementar()` o bloco fica **depois** do trecho que pula linhas repetidas (mesmo GCPJ e valor). Linha repetida não conta no limite.

### Documento de evidências

O Word com as evidências do Controle S5 (trechos do código alterado e resultado dos testes) fica em [`docs/Evidencias_S5_AutoEncerramento.docx`](docs/Evidencias_S5_AutoEncerramento.docx).

### Como testar

1. Em `LimitadorS5.cs`, troque temporariamente `LIMITE_OPERACOES = 5000` para `2`.
2. Rode a opção desejada com pelo menos 3 processos pendentes.
3. Confira: 2 processos tratados e o aviso **Processamento Limitado - Controle S5**.
4. **Volte o valor para 5000** antes de gerar a versão final.

## 5. Guia rápido: quero mudar X, onde vou?

| Quero... | Vá em |
|---|---|
| Mudar o limite do Controle S5 | `LimitadorS5.cs` ➜ `LIMITE_OPERACOES` |
| Mudar as opções de processamento | `Form1.cs` ➜ `Form1_Load` (`CBSelecionaProcessamento`) |
| Mudar os códigos de depósito, finalidade e pagamento | `Form1.cs` ➜ `tipoDeposito`, `tipoFinalidade`, `tipoPagamento`, `depositoComplementar` |
| Mudar a data de depósito (dia útil e feriados) | `Form1.cs` ➜ `solicitarEncerramento` e `retornaDataPagamento` (tabela de feriados) |
| Mudar a URL de uma tela do GCPJ | `Form1.cs` ➜ chamadas `Automacao.navegar(...)` |

## 6. Pontos de atenção

- Processos com **mais de 5 envolvidos** não são encerrados pelo robô (ficam com a observação na base).
- Em `complementar()`, quando a pesquisa ou um campo falha, a observação é gravada na `TB_INCLUSAO`, não na `TB_COMPLEMENTAR`. Esses processos continuam pendentes e voltam na próxima execução.
- O limite S5 vale **por execução e por máquina**.

## 7. Histórico

| Data | Autor | Alteração |
|---|---|---|
| 06/2023 | Pedro Cosmo | Pagamento complementar e fluxo LIQUIDACAO - CIVEL |
| 10/2026 | Gabriel Vasconcellos Gomes | Implementação do **Controle S5** (`LimitadorS5.cs`) e novo README |

---

<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="docs/legalops-logo-dark.png">
    <img src="docs/legalops-logo.png" alt="Legal Ops Bradesco" width="220">
  </picture>
  <br>
  <sub>Legal Operations • Banco Bradesco S.A. • Classificação: Interna</sub>
</p>

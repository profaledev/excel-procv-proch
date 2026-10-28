# PROCV e PROCH no Excel: Conceitos, Uso e Importância no Mercado

---

## 1. O que são PROCV e PROCH?

### PROCV (VLOOKUP - Vertical Lookup)
Função que **busca um valor na primeira coluna de uma tabela** e retorna um valor correspondente de outra coluna na mesma linha.

**Sintaxe:**
```excel
=PROCV(valor_procurado; matriz_tabela; núm_índice_coluna; [procurar_intervalo])
```

| Parâmetro | Descrição | Obrigatório? |
|-----------|-----------|--------------|
| `valor_procurado` | O que você quer encontrar (ex: ID, nome, código, célula) | Sim |
| `matriz_tabela` | Intervalo onde está a tabela **incluindo a coluna de busca** (ex: A1:D100) | Sim |
| `núm_índice_coluna` | Número da coluna que contém o resultado **contando da primeira coluna da matriz_tabela** | Sim |
| `[procurar_intervalo]` | **FALSO** (busca exata) ou **VERDADEIRO** (busca aproximada) — *padrão: VERDADEIRO* | Não |

---

### PROCH (HLOOKUP - Horizontal Lookup)
Funciona igual ao PROCV, mas **busca na primeira LINHA de uma tabela** e retorna valores de linhas abaixo.

**Sintaxe:**
```excel
=PROCH(valor_procurado; matriz_tabela; núm_índice_linha; [procurar_intervalo])
```

| Parâmetro | Descrição | Obrigatório? |
|-----------|-----------|--------------|
| `valor_procurado` | Valor a buscar na primeira linha da tabela | Sim |
| `matriz_tabela` | Intervalo da tabela horizontal **incluindo a linha de busca** | Sim |
| `núm_índice_linha` | Número da linha do resultado **contando da primeira linha da matriz_tabela** | Sim |
| `[procurar_intervalo]` | FALSO (exato) ou VERDADEIRO (aproximado) — *padrão: VERDADEIRO* | Não |

---

## 2. Diferença Visual Fundamental

```
PROCV (Vertical - Busca por COLUNA)          PROCH (Horizontal - Busca por LINHA)
┌─────┬────────┬────────┐                    ┌─────┬─────┬─────┬─────┐
│ ID  │ Nome   │ Preço  │                    │     │ Jan │ Fev │ Mar │
├─────┼────────┼────────┤                    ├─────┼─────┼─────┼─────┤
│ 1   │ Mouse  │ 80     │                    │ Vendas│1000│1500│ 800 │
│ 2   │Teclado │ 150    │                    │ Meta  │1200│1200│1200 │
│ 3   │Monitor │ 900    │                    └─────┴─────┴─────┴─────┘
└─────┴────────┴────────┘
Busca na COLUNA 1 (ID)                       Busca na LINHA 1 (Mês)
Retorna da COLUNA N                          Retorna da LINHA N
```

**Regra de Ouro:**  
- **PROCV** = Dados organizados em **colunas** (cada registro = uma linha) → **99% dos casos**  
- **PROCH** = Dados organizados em **linhas** (cada registro = uma coluna) → Raro, tabelas transpostas

---

## 3. Por que Usar? (Benefícios Práticos)

### ✅ Automatização de Buscas
- Elimina buscas manuais (`Ctrl+F`) em bases com centenas/milhares de linhas
- Reduz erros humanos (digitação, olho cansado, copia/cola errado)

### ✅ Integração de Dados (O "Join" do Excel)
- Une informações de **abas diferentes** ou **arquivos diferentes**
- Ex: Tabela de Vendas (ID_Produto) + Tabela de Produtos (ID, Nome, Preço, Estoque)

### ✅ Dashboards e Relatórios Dinâmicos
- Atualizam automaticamente quando a fonte muda
- Base para `TABELAS DINÂMICAS`, `GRÁFICOS`, `SE`, `SOMASE`

### ✅ Validação e Consistência
- Garante que códigos/IDs existem na tabela mestre
- Evita "fantasmas" (vendedor inexistente na comissão, produto fantasma no estoque)

### ✅ Produtividade Massiva
- O que levaria **horas** manualmente vira **segundos**
- Fórmula arrasta para 10.000 linhas instantaneamente

---

## 4. Exemplo Didático Completo Passo a Passo

### Cenário Real
**Aba "Vendas" (dados brutos):**
| A (ID_Produto) | B (Qtd) | C (Vendedor) |
|----------------|---------|--------------|
| 1              | 5       | Ana          |
| 2              | 3       | Bruno        |
| 1              | 2       | Carlos       |
| 3              | 1       | Ana          |

**Aba "Produtos" (tabela mestre/cadastro):**
| A (ID) | B (Código) | C (Nome) | D (Preço) | E (Categoria) | F (Estoque) |
|--------|------------|----------|-----------|---------------|-------------|
| 1      | SKU-001    | Mouse    | 80        | Periférico    | 50          |
| 2      | SKU-002    | Teclado  | 150       | Periférico    | 30          |
| 3      | SKU-003    | Monitor  | 900       | Vídeo         | 15          |
| 4      | SKU-004    | Headset  | 250       | Áudio         | 20          |

### Objetivo
Na aba **Vendas**, trazer **Nome**, **Preço** e **Categoria** usando o `ID_Produto` (coluna A).

---

### Passo 1: Entender a "Chave Primária"
- A **coluna de busca** deve ser a **PRIMEIRA coluna** da `matriz_tabela`
- Na aba Produtos, coluna A (ID) é a primeira → **OK para PROCV**

---

### Passo 2: Contar o Índice da Coluna de Retorno
Contagem **sempre começa na coluna A da matriz_tabela**:

| Coluna da Tabela | Letra | Índice (núm_índice_coluna) |
|------------------|-------|----------------------------|
| ID               | A     | 1 (chave de busca)         |
| Código           | B     | 2                          |
| **Nome**         | **C** | **3** ← **Queremos esta**  |
| **Preço**        | **D** | **4** ← **Queremos esta**  |
| **Categoria**    | **E** | **5** ← **Queremos esta**  |
| Estoque          | F     | 6                          |

---

### Passo 3: Montar as Fórmulas (Na aba Vendas, linha 2)

#### Para trazer **Nome** (coluna C da tabela Produtos = índice 3):
```excel
=PROCV(A2; Produtos!A:F; 3; FALSO)
```

#### Para trazer **Preço** (coluna D = índice 4):
```excel
=PROCV(A2; Produtos!A:F; 4; FALSO)
```

#### Para trazer **Categoria** (coluna E = índice 5):
```excel
=PROCV(A2; Produtos!A:F; 5; FALSO)
```

---

### Passo 4: Entender Cada Parte da Fórmula

| Parte | Exemplo | Explicação |
|-------|---------|------------|
| `A2` | `valor_procurado` | O ID da venda na linha atual (1, 2, 1, 3...) |
| `Produtos!A:F` | `matriz_tabela` | Intervalos A até F da aba Produtos. **Inclui a coluna de busca (A) e todas de retorno** |
| `3` / `4` / `5` | `núm_índice_coluna` | **Contado a partir da coluna A da matriz** (A=1, B=2, C=3...) |
| `FALSO` | `[procurar_intervalo]` | **OBRIGATÓRIO na maioria dos casos**. Busca **exata**. Se não achar → `#N/D` |

---

### Passo 5: Resultado Final (Aba Vendas)

| A (ID) | B (Qtd) | C (Vendedor) | D (Nome) | E (Preço) | F (Categoria) |
|--------|---------|--------------|----------|-----------|---------------|
| 1      | 5       | Ana          | Mouse    | 80        | Periférico    |
| 2      | 3       | Bruno        | Teclado  | 150       | Periférico    |
| 1      | 2       | Carlos       | Mouse    | 80        | Periférico    |
| 3      | 1       | Ana          | Monitor  | 900       | Vídeo         |

**Fórmulas nas colunas D, E, F arrastadas para baixo.**

---

## 5. O Parâmetro `[procurar_intervalo]`: FALSO vs VERDADEIRO

### ❌ **NUNCA use VERDADEIRO (ou omita) para buscas exatas** — 99% dos casos

| Valor | Comportamento | Perigo |
|-------|---------------|--------|
| `FALSO` (ou `0`) | Busca **exata**. Se não encontra → `#N/D` | **Seguro. Use sempre.** |
| `VERDADEIRO` (ou `1`/omitido) | Busca **aproximada**. **Exige coluna de busca ORDENADA crescente**. Retorna "maior valor menor ou igual" | **Perigoso.** Retorna resultado errado silenciosamente se não ordenado. |

### Exemplo do Perigo do VERDADEIRO:
Tabela **NÃO ordenada**:
| A (ID) | B (Nome) |
|--------|----------|
| 10     | Dez      |
| 1      | Um       |
| 5      | Cinco    |

Fórmula: `=PROCV(5; A:B; 2; VERDADEIRO)`  
**Resultado:** `Dez` (ERRADO! Achou 10 como "maior menor ou igual" porque não está ordenado)

**Mesma tabela ORDENADA:**
| A (ID) | B (Nome) |
|--------|----------|
| 1      | Um       |
| 5      | Cinco    |
| 10     | Dez      |

`=PROCV(5; A:B; 2; VERDADEIRO)` → `Cinco` (Correto)  
`=PROCV(7; A:B; 2; VERDADEIRO)` → `Cinco` (Aproximado: maior ≤ 7)

### 🎯 **Regra Prática:**  
> **SEMPRE use `FALSO` (ou `0`)** para buscas por ID, código, nome, CPF, etc.  
> Só use `VERDADEIRO` para **tabelas de faixas** (imposto, comissão, desconto por volume) **e a tabela estiver ordenada**.

---

## 6. Boas Práticas (Regras de Ouro)

| Prática | Fórmula Certa | Fórmula Errada | Por quê? |
|---------|---------------|----------------|----------|
| **Sempre use `FALSO`** | `=PROCV(A2; B:C; 2; FALSO)` | `=PROCV(A2; B:C; 2)` | Evita resultados errados silenciosos |
| **Trave referências com `$`** | `=PROCV(A2; Produtos!$A$2:$F$100; 3; FALSO)` | `=PROCV(A2; Produtos!A2:F100; 3; FALSO)` | Ao arrastar, a tabela não "anda" |
| **Use Tabelas Nomeadas (Ctrl+T)** | `=PROCV([@ID]; TabelaProdutos; 3; FALSO)` | Referências A1 | Legível, expande auto, não quebra ao inserir linha |
| **Verifique duplicatas na chave** | — | — | PROCV retorna **apenas a PRIMEIRA** ocorrência |
| **Trate `#N/D` com `SEERRO`** | `=SEERRO(PROCV(...); "Não encontrado")` | Deixa `#N/D` | Visual limpo, evita erros em cálculos downstream |

---

## 7. Erros Comuns, Causas e Soluções

| Erro | Causa Principal | Solução |
|------|-----------------|---------|
| `#N/D` | Valor **não existe** na coluna de busca | 1. Verifique espaços: `ARRUMAR()`<br>2. Verifique tipos: Texto vs Número (`VALOR()`, `TEXTO()`)<br>3. Use `SEERRO()` para mensagem amigável |
| `#REF!` | `núm_índice_coluna` > total de colunas da `matriz_tabela` | Confira contagem: A=1, B=2... Aumente a matriz ou diminua o índice |
| `#VALOR!` | `valor_procurado` ou `núm_índice_coluna` inválido | Índice deve ser ≥ 1. Verifique se célula de busca não tem erro |
| `#NOME?` | Nome da função errado (ex: `VLOOKUP` em PT-BR) | Use `PROCV` (PT-BR) ou `VLOOKUP` (EN-US) |
| Retorna valor errado | Usou `VERDADEIRO` sem ordenar / Duplicatas / Coluna errada | Use `FALSO`. Confira índice. Remova duplicatas da chave. |

---

## 8. PROCH (HLOOKUP) - Exemplos Simples e Diretos

### 8.1 Quando Usar PROCH (Regra Simples)
**Use PROCH quando seus dados estão "deitados": cabeçalhos na LINHA 1, dados nas linhas abaixo.**

| Formato dos Dados | Função |
|-------------------|--------|
| ID na **coluna A**, registros descendo (normal) | **PROCV** |
| Meses/Períodos na **linha 1**, indicadores descendo | **PROCH** |

---

### 8.2 Exemplo Básico 1: Consulta de Vendas por Mês

**Aba "Vendas_Mensais" (Horizontal):**
| A (Indicador) | B (Jan) | C (Fev) | D (Mar) | E (Abr) | F (Mai) | G (Jun) |
|---------------|---------|---------|---------|---------|---------|---------|
| **Vendas**    | 10000   | 12000   | 9500    | 14000   | 13000   | 15000   |
| **Meta**      | 11000   | 11000   | 11000   | 11000   | 11000   | 11000   |
| **Custo**     | 6000    | 7200    | 5700    | 8400    | 7800    | 9000    |

**Objetivo:** Na célula `B1` digito o mês (ex: "Mar"), e abaixo aparecem os valores automaticamente.

**Fórmulas (travadas com $ para copiar):**
```excel
B2 (Vendas): =PROCH($B$1; Vendas_Mensais!$A$1:$G$3; 2; FALSO)
B3 (Meta):   =PROCH($B$1; Vendas_Mensais!$A$1:$G$3; 3; FALSO)
B4 (Custo):  =PROCH($B$1; Vendas_Mensais!$A$1:$G$3; 4; FALSO)
B5 (Lucro):  =B2-B4
```

**Resultado se B1 = "Mar":**
| | B |
|---|---|
| **Mês** | Mar |
| **Vendas** | 9500 |
| **Meta** | 11000 |
| **Custo** | 5700 |
| **Lucro** | 3800 |

**Explicação rápida:**
- `$B$1` = mês buscado (travado para não mudar ao copiar)
- `Vendas_Mensais!$A$1:$G$3` = matriz **incluindo linha 1 (cabeçalhos)**
- `2` = linha 2 da matriz = **Vendas** (linha 1=1, linha 2=2, linha 3=3, linha 4=4)
- `FALSO` = busca exata do mês

---

### 8.3 Exemplo Básico 2: Notas de Alunos (Bimestres nas Colunas)

**Aba "Notas_Turma" (Horizontal):**
| A (Aluno/Disciplina) | B (1º Bim) | C (2º Bim) | D (3º Bim) | E (4º Bim) |
|----------------------|------------|------------|------------|------------|
| **João - Matemática** | 7.5        | 8.0        | 6.5        | 9.0        |
| **João - Português**  | 8.0        | 7.5        | 8.5        | 8.0        |
| **Maria - Matemática**| 9.0        | 9.5        | 8.0        | 9.5        |
| **Maria - Português** | 8.5        | 9.0        | 9.0        | 9.0        |

**Objetivo:** Escolher aluno+disciplina (coluna A) e bimestre (linha 1), trazer a nota.

**Fórmula:**
```excel
=PROCH("2º Bim"; Notas_Turma!$A$1:$E$5; CORRESP("João - Matemática"; Notas_Turma!$A$1:$A$5; 0); FALSO)
```
*Simplificado sem CORRESP (fixo para João - Matemática = linha 2):*
```excel
=PROCH("2º Bim"; Notas_Turma!$A$1:$E$5; 2; FALSO)
```
**Resultado:** `8.0`

---

### 8.4 Exemplo Básico 3: Preços por Faixa de Quantidade (Horizontal)

**Aba "Tabela_Preco" (Horizontal):**
| A (Faixa) | B (1-10) | C (11-50) | D (51-100) | E (101+) |
|-----------|----------|-----------|------------|----------|
| **Mouse** | 80       | 75        | 70         | 65       |
| **Teclado**| 150     | 140       | 130        | 120      |
| **Monitor**| 900     | 850       | 800        | 750      |

**Objetivo:** Produto na célula `A2` (ex: "Teclado"), faixa na `B1` (ex: "11-50"), retornar preço.

**Fórmula simples (linha fixa = 2 para Teclado):**
```excel
=PROCH(B1; Tabela_Preco!$A$1:$E$4; 2; FALSO)
```
**Resultado:** `140`

> **Dica:** Se o produto muda, use `CORRESP` na linha (ver seção avançada). Para fixo, número da linha direto.

---

### 8.5 Exemplo Básico 4: Transpor Dados Horizontal → Vertical

**Fonte Horizontal (Aba "Fonte_H"):**
| A | B (Jan) | C (Fev) | D (Mar) |
|---|---------|---------|---------|
| **Produto** | Mouse | Teclado | Monitor |
| **Vendas**  | 100   | 50      | 20      |

**Destino Vertical (Aba "Destino_V"):**
| A (Mês) | B (Produto) | C (Vendas) |
|---------|-------------|------------|
| Jan     | `=PROCH(A2; Fonte_H!$A$1:$D$2; 2; FALSO)` | `=PROCH(A2; Fonte_H!$A$1:$D$2; 3; FALSO)` |
| Fev     | `=PROCH(A3; Fonte_H!$A$1:$D$2; 2; FALSO)` | `=PROCH(A3; Fonte_H!$A$1:$D$2; 3; FALSO)` |
| Mar     | `=PROCH(A4; Fonte_H!$A$1:$D$2; 2; FALSO)` | `=PROCH(A4; Fonte_H!$A$1:$D$2; 3; FALSO)` |

**Resultado:**
| Mês | Produto | Vendas |
|-----|---------|--------|
| Jan | Mouse   | 100    |
| Fev | Teclado | 50     |
| Mar | Monitor | 20     |

---

### 8.6 Erros Comuns do PROCH (Resumo Rápido)

| Erro | Causa | Correção |
|------|-------|----------|
| `#N/D` | Mês/nome não está na **linha 1** da matriz | Confira grafia, espaços, use `SEERRO` |
| `#REF!` | Número da linha > total de linhas da matriz | Matriz A1:E4 tem 4 linhas → máx índice = 4 |
| Valor errado | Matriz **não começa na linha do cabeçalho** | Matriz deve ser `A1:E4`, não `A2:E4` |
| Não copia certo | Referências não travadas com `$` | Use `$A$1:$E$4` |

---

## 9. Exercícios Práticos - PROCV e PROCH

### 9.1 Exercícios PROCV (Vertical - Dados em Colunas)

**Base de Dados (crie aba "Produtos"):**
| ID | Código | Nome | Preço | Categoria | Estoque |
|----|--------|------|-------|-----------|---------|
| 1  | SKU-001| Mouse | 80    | Periférico| 50      |
| 2  | SKU-002| Teclado| 150  | Periférico| 30      |
| 3  | SKU-003| Monitor| 900  | Vídeo     | 15      |
| 4  | SKU-004| Headset| 250  | Áudio     | 20      |
| 5  | SKU-005| Webcam | 300  | Vídeo     | 10      |

**Crie aba "Vendas" com:**
| ID_Produto | Qtd | Vendedor |
|------------|-----|----------|
| 1          | 5   | Ana      |
| 3          | 1   | Bruno    |
| 2          | 2   | Ana      |
| 5          | 3   | Carlos   |
| 1          | 1   | Bruno    |

#### Exercício 1 - Básico (Nome e Preço)
Na aba Vendas, colunas D e E, traga **Nome** e **Preço** do produto usando PROCV.
- D2: `=PROCV(A2; Produtos!A:F; 3; FALSO)`
- E2: `=PROCV(A2; Produtos!A:F; 4; FALSO)`
- Arraste para baixo.

#### Exercício 2 - Com SEERRO (Tratamento)
Modifique as fórmulas para mostrar **"Não cadastrado"** se o ID não existir.
- `=SEERRO(PROCV(A2; Produtos!A:F; 3; FALSO); "Não cadastrado")`

#### Exercício 3 - Categoria e Estoque
Colunas F e G: traga **Categoria** (índice 5) e **Estoque** (índice 6).

#### Exercício 4 - Validação de Estoque
Coluna H: `=SE(E2<=F2; "OK"; "Estoque Baixo")` (comparar Qtd vendida vs Estoque).

#### Exercício 5 - Busca por Código (Não é a primeira coluna!)
**Desafio:** Na aba Vendas, coluna B tem **Código** (SKU-001), não ID.
Crie coluna auxiliar em Produtos!G: `=B2` (cópia do Código) → PROCV na coluna G.
Ou use: `=ÍNDICE(Produtos!C:C; CORRESP(B2; Produtos!B:B; 0))` para Nome.

---

### 9.2 Exercícios PROCH (Horizontal - Dados em Linhas)

**Base de Dados (crie aba "Metas_2024"):**
| A (Indicador) | B (Jan) | C (Fev) | D (Mar) | E (Abr) | F (Mai) | G (Jun) |
|---------------|---------|---------|---------|---------|---------|---------|
| **Receita**   | 120000  | 135000  | 110000  | 145000  | 160000  | 155000  |
| **Custo**     | 72000   | 81000   | 66000   | 87000   | 96000   | 93000   |
| **Despesa**   | 24000   | 27000   | 22000   | 29000   | 32000   | 31000   |
| **Meta**      | 115000  | 115000  | 115000  | 115000  | 115000  | 115000  |

**Crie aba "Dashboard" com:**
- Célula `B1`: **Validação de Dados** → Lista → Fonte: `Metas_2024!B1:G1` (meses)
- Linhas 3 a 6: fórmulas PROCH

#### Exercício 1 - Dashboard Básico
B3 (Receita): `=PROCH($B$1; Metas_2024!$A$1:$G$4; 2; FALSO)`
B4 (Custo):   `=PROCH($B$1; Metas_2024!$A$1:$G$4; 3; FALSO)`
B5 (Despesa): `=PROCH($B$1; Metas_2024!$A$1:$G$4; 4; FALSO)`
B6 (Meta):    `=PROCH($B$1; Metas_2024!$A$1:$G$4; 5; FALSO)`
B7 (Lucro):   `=B3-B4-B5`
B8 (vs Meta): `=B3-B6`

Teste trocando mês em B1.

#### Exercício 2 - Com SEERRO
Envolva todas em `SEERRO(...; "Mês inválido")`.

#### Exercício 3 - Transposição Horizontal → Vertical
Crie aba "Metas_Vertical":
| A (Mês) | B (Receita) | C (Custo) | D (Despesa) | E (Meta) | F (Lucro) |
|---------|-------------|-----------|-------------|----------|-----------|
| Jan     | PROCH...    | PROCH...  | PROCH...    | PROCH... | =B2-C2-D2 |
| Fev     | ...         | ...       | ...         | ...      | ...       |
| Mar     | ...         | ...       | ...         | ...      | ...       |
| ...     | ...         | ...       | ...         | ...      | ...       |

Preencha A2:A7 com Jan a Jun. Arraste fórmulas.

#### Exercício 4 - Busca Aproximada (Faixas) - OPCIONAL
Crie aba "Comissao_Faixas" horizontal:
| A (Faixa) | B (Até 5k) | C (5k-15k) | D (15k-30k) | E (30k+) |
|-----------|------------|------------|-------------|----------|
| **%**     | 5%         | 7%         | 10%         | 15%      |

Fórmula (VERDADEIRO exige cabeçalhos numéricos ordenados):
`=PROCH(18000; Comissao_Faixas!$A$1:$E$2; 2; VERDADEIRO)` → deve retornar `10%`
*Nota: Cabeçalhos devem ser números: 5000, 15000, 30000, 99999999*

#### Exercício 5 - Tabela de Preços Simples
Crie aba "Precos_H" horizontal:
| A (Produto) | B (Pequeno) | C (Médio) | D (Grande) |
|-------------|-------------|-----------|------------|
| **Camiseta**| 30          | 35        | 40         |
| **Calça**   | 80          | 90        | 100        |

Em outra aba, célula A1 = "Camiseta", B1 = "Médio".
Fórmula: `=PROCH(B1; Precos_H!$A$1:$D$3; 2; FALSO)` → `35`
(Troque 2 por 3 para Calça).

---

### 9.3 Exercícios Mistos (PROCV + PROCH Juntos)

**Cenário:** Você tem:
- Aba "Cadastro_Clientes" (Vertical: ID, Nome, Cidade, Estado)
- Aba "Vendas_Mensais" (Horizontal: Meses na linha 1, IDs dos clientes nas linhas, valores nas células)

**Desafio 1:** No Dashboard, escolha Mês (PROCH na Vendas_Mensais) → pega ID do cliente → PROCV no Cadastro_Clientes → traz Nome e Cidade.

**Desafio 2:** Crie relatório vertical final: Mês | Cliente | Cidade | Venda
Use PROCH para pegar venda do mês + PROCV para dados do cliente.

---

### 9.4 Gabarito Rápido (Conferência)

| Exercício | Fórmula Chave | Resultado Esperado |
|-----------|---------------|-------------------|
| PROCV 1 | `PROCV(A2; Produtos!A:F; 3; FALSO)` | "Mouse", "Monitor", "Teclado"... |
| PROCV 2 | `SEERRO(PROCV(...); "Não cadastrado")` | Sem #N/D |
| PROCH 1 | `PROCH($B$1; Metas_2024!$A$1:$G$4; 2; FALSO)` | Receita do mês escolhido |
| PROCH 3 | `PROCH(A2; Metas_2024!$A$1:$G$4; 2; FALSO)` arrastado | Tabela vertical completa |
| Misto | PROCH(p/ venda) + PROCV(p/ cliente) | Relatório unificado |

---

### 9.5 Checklist Final (Marque ao Concluir)

**PROCV:**
- [ ] Todos usam `FALSO`
- [ ] Matriz travada com `$` ou Tabela Nomeada
- [ ] Índice conferido (A=1, B=2, C=3...)
- [ ] `SEERRO` nas visíveis
- [ ] Sem duplicatas na chave (ou ciente que pega 1º)

**PROCH:**
- [ ] Matriz **inclui linha 1 (cabeçalhos)**
- [ ] Índice de linha conferido (linha 1 da matriz = 1)
- [ ] `FALSO` para textos (meses, nomes)
- [ ] Referências travadas com `$`
- [ ] Linha 1 sem células mescladas

---

> **Dica de Ouro:** Faça os exercícios **na ordem**, sem pular. A memória muscular do `FALSO`, `$` e contagem de índice é o que evita erros no trabalho real.

---

## 9. PROCV Avançado: Truques Essenciais

### 9.1 Buscar em Outra Planilha/Arquivo
```excel
=PROCV(A2; '[Cadastro_Produtos.xlsx]Produtos'!$A$2:$F$500; 3; FALSO)
```
- Arquivo **aberto**: Referência normal
- Arquivo **fechado**: Caminho completo `'C:\Pasta\[Arquivo.xlsx]Aba'!$A$2:$F$500`

### 9.2 PROCV com Múltiplos Critérios (Chave Composta)
**Problema:** Buscar por **ID + Filial** (duas colunas)
**Solução:** Criar **Coluna Auxiliar** concatenada nas duas tabelas.

**Na tabela fonte (Produtos):**
| A (ID) | B (Filial) | C (Chave_Aux) | D (Nome) |
|--------|------------|---------------|----------|
| 1      | SP         | `=A2&B2` → "1SP" | Mouse    |

**Na tabela busca (Vendas):**
| A (ID) | B (Filial) | C (Chave_Aux) | D (Nome - PROCV) |
|--------|------------|---------------|------------------|
| 1      | SP         | `=A2&B2` → "1SP" | `=PROCV(C2; Produtos!C:D; 2; FALSO)` |

### 9.3 PROCV "À Esquerda" (Gambiarra com ÍNDICE + CORRESP)
**Problema:** Chave de busca **NÃO é a primeira coluna** (ex: buscar Nome, retornar ID)
**Solução:** `ÍNDICE` + `CORRESP` (não usa PROCV)

```excel
=ÍNDICE(Produtos!A:A; CORRESP("Mouse"; Produtos!C:C; 0))
```
| Função | Papel |
|--------|-------|
| `CORRESP("Mouse"; C:C; 0)` | Acha **número da linha** onde "Mouse" está na coluna C → retorna `2` |
| `ÍNDICE(A:A; 2)` | Retorna valor da **coluna A, linha 2** → `1` (o ID) |

**Vantagem:** Busca em **qualquer coluna**, retorna de **qualquer coluna**.

### 9.4 PROCV com Validação de Lista (Data Validation)
1. `Dados` → `Validação de Dados` → `Lista`
2. Fonte: `=PROCV(A2; Produtos!A:F; 3; FALSO)` → **Não funciona direto**
3. **Correto:** Use `ÍNDICE` + `CORRESP` ou intervalo nomeado dinâmico, ou Tabela (Ctrl+T) + `INDIRETO`

### 9.5 PROCV em Coluna Não-Única (Duplicatas) - Pega a Primeira
Se a tabela tem IDs repetidos (ex: histórico de preços), PROCV **sempre pega o primeiro**.
**Solução para pegar o último:** Ordene por Data decrescente + use `ÍNDICE` + `CORRESP` com `1` (maior menor ou igual) ou Power Query.

---

## 10. Importância no Mercado de Trabalho

### 🎯 Habilidade **Obrigatória** para Cargos:

| Cargo | Uso Típico do PROCV/PROCH |
|-------|---------------------------|
| **Analista de Dados / BI** | ETL leve, validação cruzada, staging tables, enriquecimento |
| **Analista Financeiro / Controladoria** | Conciliação bancária, DRE, fluxo de caixa, rateios, comissões |
| **Gestor de Projetos / PMO** | Acompanhamento de KPIs, relatórios de status, alocação |
| **Administrativo / Comercial / Atendimento** | Cadastros, precificação, lookup de cliente/produto, comissões |
| **Contador / Auditor** | Conferência SPED, importação contábil, cruzamento de bases |
| **Logística / Supply Chain** | Endereçamento, estoque, rastreamento, fretes |
| **RH / DP** | Folha de ponto, benefícios, cadastro funcionários, férias |

### 📈 Diferencial Competitivo Real
| Habilidade | Impacto |
|------------|---------|
| **Fazer PROCV certo (com FALSO, $, SEERRO)** | Evita retrabalho, auditoria, erros financeiros |
| **Saber diagnosticar `#N/D` rápido** | Autonomia, não paralisa equipe |
| **Integrar abas/arquivos sem VBA/Power Query** | Entrega imediata em ambientes restritos |
| **Explicar a lógica para não-técnicos** | Comunicação, liderança técnica |
| **Saber quando NÃO usar (Power Query, SQL, PROCX)** | Maturidade, escalabilidade |

### 💬 Perguntas Comuns em Entrevistas Técnicas
> 1. "Como você traria o nome do cliente para a planilha de vendas?"
> 2. "Qual a diferença entre `FALSO` e `VERDADEIRO` no último argumento?"
> 3. "O que significa `#N/D` e como você resolve?"
> 4. "PROCV busca à esquerda? Como contornar?"
> 5. "Como fazer PROCV com dois critérios (ex: ID + Data)?"
> 6. "Quando o PROCV fica lento? O que fazer?"
> 7. "Diferença entre PROCV e PROCH?"

---

## 11. Checklist de Qualidade (Antes de Entregar Planilha)

- [ ] Todos os PROCV/POCH usam `FALSO` (ou `0`)?
- [ ] Referências da `matriz_tabela` travadas com `$` (ou Tabela Nomeada)?
- [ ] `núm_índice_coluna/linha` conferido (contagem começa na 1 da matriz)?
- [ ] `SEERRO` envolvendo fórmulas visíveis ao usuário?
- [ ] Coluna de busca **não tem duplicatas** (ou assumido pegar a primeira)?
- [ ] Tipos de dados compatíveis (texto com texto, número com número)?
- [ ] Sem espaços invisíveis (`ARRUMAR` aplicado se veio de sistema externo)?
- [ ] Documentado (comentário na célula ou aba "Legenda") o que cada PROCV faz?

---

## 12. Resumo Conceitual Para Fixar (Mental Model)

| Conceito | Analogia | No Excel |
|----------|----------|----------|
| **Chave Primária** | CPF, RG, Placa do carro | Coluna única que identifica cada linha (ex: ID) |
| **Tabela de Referência (Lookup Table)** | Lista telefônica, Catálogo | Fonte mestra onde a busca acontece (`matriz_tabela`) |
| **Correspondência Exata** | "Ache **igualzinho**" | `FALSO` / `0` — **Padrão obrigatório** |
| **Índice de Retorno** | "Pule N colunas da chave" | `núm_índice_coluna` (PROCV) ou `núm_índice_linha` (PROCH) |
| **Join (Junção)** | Unir duas planilhas pelo que têm em comum | O que PROCV faz na prática (LEFT JOIN) |
| **Cardinalidade** | 1:1 ou 1:N | PROCV assume 1:1 (chave única). Se 1:N → pega primeiro |

---

## 13. Referências Oficiais Microsoft

- [Função PROCV (VLOOKUP)](https://support.microsoft.com/pt-br/office/fun%C3%A7%C3%A3o-procv-0bbc8083-26fe-4963-8ab8-93a18ad188a1)
- [Função PROCH (HLOOKUP)](https://support.microsoft.com/pt-br/office/fun%C3%A7%C3%A3o-proch-446d94af-663b-451d-8c1f-2315e8b9e5f0)
- [Função ÍNDICE](https://support.microsoft.com/pt-br/office/fun%C3%A7%C3%A3o-indice-a549f90f-5b1e-4d0b-8762-6d8f8d8a0d8e)
- [Função CORRESP (MATCH)](https://support.microsoft.com/pt-br/office/fun%C3%A7%C3%A3o-corresp-e8d1d4d2-4d7f-4f8d-9b5b-5b5b5b5b5b5b)
- [Função SEERRO (IFERROR)](https://support.microsoft.com/pt-br/office/fun%C3%A7%C3%A3o-seerro-66c9c281-64b6-49e8-9781-5b5b5b5b5b5b)

---

> **"Domine PROCV com `FALSO`, `$` e `SEERRO`. É a habilidade que separa 'usuário de Excel' de 'analista que entrega'."** 🚀
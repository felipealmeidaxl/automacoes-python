# 04 · Triagem de publicações jurídicas

**Problema:** todos os dias chegam milhares de publicações dos diários oficiais. Uma pessoa precisava ler cada uma e decidir se ela é do escritório (manter) ou não (descartar), cruzando com a base de processos cadastrados.

**Solução:** um classificador baseado em regras que reproduz o raciocínio de quem faz a triagem manual e grava o resultado numa cópia da planilha.

## Como funciona

1. Lê a planilha de tratamento (aba com o teor das publicações) e a base de processos cadastrados
2. Normaliza o nome do diário (`DJEN-TJRJ` → `TJRJ`, `DJE-STF` → `STF`...)
3. Classifica cada publicação por categoria (cliente localizado, STJ/STF, agravo, outros escritórios, trabalhista...)
4. Para recursos em tribunais superiores, procura no teor o **número CNJ do processo originário** e confere se ele está na base
5. Procura advogados do escritório e clientes da casa citados no teor
6. Grava numa **cópia** da planilha (via Excel COM, preservando fórmulas e links): decisão, categoria, motivo e um resumo

## Destaques

- Casos de dúvida **ficam em branco** para revisão humana — a automação não adivinha
- Cada linha recebe o **motivo** da decisão, então a conferência é rápida
- Relatório no console comparando a decisão automática com a marcação atual
- Python para as regras + PowerShell/Excel para gravar sem quebrar a planilha original

> Todos os nomes de clientes, escritórios e advogados no código são fictícios.

## Rodar

Coloque o script na pasta com a `PLANILHA DE TRATAMENTO ... .xlsx` e a `BASE PROCESSOS ... .xlsx` e dê dois cliques, ou:

```bash
python triagem_publicacoes.py "PLANILHA DE TRATAMENTO - 06.11.2026.xlsx" "BASE PROCESSOS 06.11.2026.xlsx"
```

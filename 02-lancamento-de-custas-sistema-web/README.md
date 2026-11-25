# 02 · Lançamento de custas em sistema web

**Problema:** as custas judiciais pagas no dia chegam numa planilha e precisavam ser lançadas manualmente no módulo financeiro de cada processo, escolhendo o tipo de movimento certo para cada natureza de despesa.

**Solução:** um robô que lê a planilha mais recente da pasta e faz todos os lançamentos no sistema.

## Como funciona

1. Calcula o **próximo dia útil** e pergunta se é feriado (pula feriados seguidos, como o Carnaval)
2. Abre a planilha `.xlsx` mais recente da pasta configurada
3. Para cada linha: pesquisa o processo → abre o **Financeiro** → preenche data, valor (formato brasileiro) e histórico
4. Escolhe o tipo de movimento por uma tabela **DE/PARA** (natureza da planilha → opção do sistema)
5. Natureza desconhecida? A automação **pausa** e pede para o usuário escolher à mão, depois continua

## Destaques

- **Modo teste**: preenche tudo, mas não clica em "Incluir" — para conferir antes de rodar valendo
- Fecha alertas JavaScript inesperados e trabalha com pop-ups
- Movimentos genéricos recebem a natureza da despesa no histórico automaticamente

## Rodar

```bash
python lancar_custas.py
```

Colunas usadas: **D** processo · **J** valor · **K** natureza · **L** histórico.

# 05 · Enumerar PDFs do malote

**Problema:** o malote de guias judiciais precisa sair com os PDFs numerados na mesma ordem da planilha de controle, com os documentos do mesmo processo agrupados e os de valor alto destacados no topo. Fazer isso à mão, renomeando arquivo por arquivo, tomava tempo e gerava erros.

**Solução:** um script que lê a planilha (inclusive protegida por senha), lê os valores dentro dos PDFs e renomeia tudo na ordem certa.

## Como funciona

1. Abre a planilha de guias mais recente da pasta (descriptografa se tiver senha)
2. Numera os PDFs na ordem dos nomes da planilha: `01- `, `02- `...
3. PDFs do mesmo processo recebem ` - 1`, ` - 2`...
4. Lê o "valor a ser atualizado" da memória de cálculo (ou o valor da guia) e, acima de R$ 20.000,00, coloca o arquivo no topo com o sufixo ` - AC. MED.`; guias de honorários nunca entram nessa regra
5. Grava um **log CSV** (de → para) com o critério usado em cada arquivo

## Destaques

- `--simular`: mostra o que faria sem renomear nada
- `--desfazer`: volta todos os nomes a partir do log
- Renomeia em duas etapas (nomes temporários) para não colidir arquivos
- Instala sozinho as bibliotecas que faltarem

## Rodar

```bash
python enumerar_malote.py                 # abre uma janela para escolher a pasta
python enumerar_malote.py "C:\malote" --simular
python enumerar_malote.py "C:\malote" --desfazer
```

Senha da planilha (se houver) no `.env`: `SENHA_PLANILHA_MALOTE`.

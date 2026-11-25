# 03 · Parear comprovantes por código de barras

**Problema:** comprovantes de pagamento e guias chegam misturados numa pasta. Antes de anexá-los no sistema, era preciso abrir cada PDF, achar a guia correspondente e renomear o comprovante com o nome dela.

**Solução:** um script que lê todos os PDFs, extrai a **linha digitável** do comprovante e encontra a guia que contém esse código.

## Como funciona

1. Separa comprovantes (palavra "COMPROVANTE" na 1ª página) e guias
2. Extrai o código de 47/48 dígitos com uma busca tolerante a texto "quebrado" pelo leitor de PDF
3. Valida o par mesmo quando a guia tem números intrometidos no meio do código (compara em 3 blocos)
4. Renomeia o comprovante com o nome da guia e move para `RESULTADO/`

## Destaques

- **Nunca sobrescreve**: nomes repetidos viram `nome_2.pdf`, `nome_3.pdf`
- Grava um **relatório .txt** com pareados, pendentes, renomeados e erros
- Pop-up final de sucesso ou de "verificar pendências"

## Rodar

Coloque o script na pasta dos PDFs e execute:

```bash
python parear_comprovantes.py
```

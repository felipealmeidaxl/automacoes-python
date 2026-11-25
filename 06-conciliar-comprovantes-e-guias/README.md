# 06 · Conciliar comprovantes e guias do dia

**Problema:** os pagamentos do dia geram comprovantes de boleto e de PIX que precisam ser ligados às guias pagas. PIX não tem código de barras, então o pareamento manual dependia de comparar valor, data e nome do beneficiário.

**Solução:** uma evolução da automação 03, empacotada como `.exe` para a equipe usar com dois cliques, com pareamento em várias etapas.

## Ordem de pareamento

1. **Código de barras** (comprovante × texto da guia)
2. **Valor + data** — boleto: vencimento igual · PIX: pago de 15 dias antes a 5 dias depois do vencimento
3. **Valor + data + nome do beneficiário** (pontuação por palavras em comum) quando sobra mais de uma guia
4. **Eliminação**: cada guia só é usada uma vez, e o processo repete enquanto resolver novos casos
5. **Empate ou dúvida**: não move — avisa o motivo e pede conferência manual

## Destaques

- Lê valor, vencimento, data de pagamento e beneficiário de diferentes layouts de comprovante
- Explica **por que** cada comprovante não foi pareado (PDF escaneado, protegido por senha, valor sem guia, data não bate...)
- Nunca apaga nada em `RESULTADO/` (numera `(2)`, `(3)`)
- Pop-ups nativos do Windows de início e fim, com totais

## Rodar

```bash
python conciliar_comprovantes.py
# ou gerar o executável:
python -m PyInstaller --onefile --console --name "Automação de Comprovantes" conciliar_comprovantes.py
```

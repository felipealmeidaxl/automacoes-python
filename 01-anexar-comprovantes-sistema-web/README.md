# 01 · Anexar comprovantes em sistema web

**Problema:** todo dia, dezenas de comprovantes de pagamento precisavam ser anexados, um a um, no processo correspondente de um sistema jurídico web: pesquisar o processo, abrir a tela de documentos, preencher descrição e tipo, escolher o arquivo e enviar.

**Solução:** um robô com Selenium que faz login, lê o número do processo no nome de cada PDF e executa todo o envio sozinho.

## Como funciona

1. Login automático (credenciais no `.env`)
2. Para cada PDF da pasta, pega o número do processo no início do nome (`123456 - comprovante.pdf`)
3. Pesquisa o processo → abre **Documentos** (nova janela) → preenche descrição e tipo → envia o arquivo
4. Volta para a página inicial e segue para o próximo

## Destaques

- Procura os campos dentro de **frames e iframes** (o sistema usa vários)
- **ESC** interrompe com segurança entre um arquivo e outro
- Se o sistema cair ou travar, a automação **para e informa exatamente onde parou** e o que já foi enviado
- Relatório final em pop-up com os processos atualizados

## Rodar

```bash
python anexar_comprovantes.py
```

Configuração no `.env`: `SISTEMA_URL`, `SISTEMA_USUARIO`, `SISTEMA_SENHA`, `PASTA_COMPROVANTES`.

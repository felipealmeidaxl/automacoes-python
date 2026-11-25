# 🤖 Automações em Python para rotinas financeiras e jurídicas

Coleção de automações que desenvolvi para eliminar trabalho manual e repetitivo em rotinas de **faturamento, conciliação de pagamentos e controle de processos**. Todas rodam hoje no dia a dia de uma equipe.

> **Observação:** esta é uma versão de portfólio. Credenciais, endereços de sistemas, nomes de pessoas, clientes e empresas foram removidos ou trocados por dados **fictícios**. Configurações sensíveis ficam em um arquivo `.env` que não é versionado.

## Projetos

| # | Automação | O que resolve | Tecnologias |
|---|-----------|---------------|-------------|
| 01 | [Anexar comprovantes em sistema web](01-anexar-comprovantes-sistema-web) | Envia dezenas de comprovantes em PDF para os processos certos de um sistema web, sem digitação | Selenium, Tkinter |
| 02 | [Lançamento de custas](02-lancamento-de-custas-sistema-web) | Lê uma planilha e lança cada despesa no financeiro do sistema, com regra de dia útil/feriado e tabela DE/PARA | Selenium, openpyxl |
| 03 | [Parear comprovantes por código de barras](03-parear-comprovantes-por-codigo-de-barras) | Encontra a guia de cada comprovante pela linha digitável e renomeia/organiza os arquivos | PyPDF2, Regex |
| 04 | [Triagem de publicações jurídicas](04-triagem-de-publicacoes-juridicas) | Classifica milhares de publicações (manter/descartar) com regras de negócio e grava o resultado no Excel | openpyxl, PowerShell + Excel COM |
| 05 | [Enumerar PDFs do malote](05-enumerar-pdfs-do-malote) | Numera e ordena os PDFs conforme a planilha, destaca valores altos e permite desfazer | pypdf, msoffcrypto, openpyxl |
| 06 | [Conciliar comprovantes e guias](06-conciliar-comprovantes-e-guias) | Pareia comprovantes de boleto e PIX com as guias por código de barras, valor, data e beneficiário | PyPDF2, Regex, PyInstaller |

## Destaques técnicos

- **Automação web resiliente:** busca de elementos dentro de frames/iframes, tratamento de pop-ups e alertas, parada segura com ESC e relatório de onde parou se o sistema cair.
- **Leitura de PDFs:** extração de linha digitável (47/48 dígitos), valores, datas e beneficiários com expressões regulares.
- **Regras de negócio claras:** tabelas DE/PARA, critérios de pareamento em etapas e casos de dúvida que **não** são adivinhados — vão para revisão manual.
- **Segurança dos dados:** nunca sobrescreve arquivos, sempre grava cópias, mantém log que permite desfazer e tem modo de simulação/teste.
- **Feito para quem não programa:** pop-ups de início e fim, mensagens de erro em português e execução com dois cliques (ou `.exe`).

## Como rodar

```bash
pip install -r requirements.txt
cp .env.example .env   # preencha com os seus dados
python 03-parear-comprovantes-por-codigo-de-barras/parear_comprovantes.py
```

Cada pasta tem um README com o passo a passo da automação.

---

Feito por **Felipe de Almeida Bastos** · [LinkedIn](https://www.linkedin.com/in/felipe-d-62204629a/)

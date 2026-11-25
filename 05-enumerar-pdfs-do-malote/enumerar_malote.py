# -*- coding: utf-8 -*-
"""
ROBÔ - ENUMERAR PDFs DO MALOTE

O que faz:
  1. Lê a planilha (.xlsx) da pasta (pode ser protegida por senha, definida no .env)
     -> coluna H (NOME) e G (Nº PROCESSO).
  2. Enumera os PDFs na ordem em que os nomes aparecem na planilha: "01- ", "02- ", ...
  3. PDFs do mesmo nome/processo recebem " - 1", " - 2", " - 3"... no final.
  4. AC. MED. (vai para o topo "01- " e recebe " - AC. MED." no final):
       - guia de HONORÁRIOS (coluna A da planilha) -> NUNCA é AC. MED.;
       - PDF com memória de cálculo -> "Valor a ser atualizado" acima de R$ 20.000,00;
       - PDF sem memória de cálculo (ex.: execução) -> valor da guia acima de R$ 20.000,00.
     A guia é ligada à linha da planilha pelo processo + valor ("importe total" = coluna E).
  5. Usa a planilha "GUIAS..." mais recente da pasta (o nome muda a cada malote).
  6. Grava um LOG na pasta (de -> para) que permite desfazer.

Uso:
  python enumerar_malote.py                 -> abre janela para escolher a pasta
  python enumerar_malote.py "C:\\caminho\\do\\MALOTE"  -> usa a pasta informada
  opções:  --simular   (só mostra o que faria, não renomeia)
           --desfazer  (volta os nomes do último LOG)
"""
import csv
import datetime
import glob
import io
import logging
import os
import re
import sys
import unicodedata

import warnings

logging.disable(logging.CRITICAL)  # silencia avisos de fontes do pypdf
warnings.filterwarnings("ignore")   # silencia aviso de validação de dados do openpyxl

# Senha da planilha protegida: vem da variável de ambiente / arquivo .env (nunca no código)
try:
    from dotenv import find_dotenv, load_dotenv
    load_dotenv(find_dotenv())
except ImportError:
    pass
SENHA_PLANILHA = os.getenv("SENHA_PLANILHA_MALOTE", "")
LIMITE_AC_MED = 20000.00
SUFIXO_AC_MED = " - AC. MED."
NOME_LOG = "_LOG_ENUMERACAO.csv"

RE_PROCESSO = re.compile(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}")


# ----------------------------------------------------------------- utilidades
def normalizar(txt):
    """Maiúsculo, sem acento, espaços simples."""
    txt = unicodedata.normalize("NFKD", str(txt or ""))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", txt).strip().upper()


def valor_br(txt):
    return float(txt.replace(".", "").replace(",", "."))


def dependencias():
    faltando = []
    for mod, pacote in (("openpyxl", "openpyxl"), ("msoffcrypto", "msoffcrypto-tool"), ("pypdf", "pypdf")):
        try:
            __import__(mod)
        except ImportError:
            faltando.append(pacote)
    if faltando:
        print("Instalando bibliotecas:", ", ".join(faltando))
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", *faltando])


# ------------------------------------------------------------------ planilha
def _linhas_da_planilha(caminho):
    """Devolve as linhas abaixo do cabeçalho (coluna H = 'NOME') ou None se não for a planilha de guias."""
    import msoffcrypto
    import openpyxl

    with open(caminho, "rb") as f:
        dados = f.read()
    buf = io.BytesIO(dados)
    try:
        office = msoffcrypto.OfficeFile(io.BytesIO(dados))
        if office.is_encrypted():
            if not SENHA_PLANILHA:
                return None  # planilha protegida, mas nenhuma senha configurada no .env
            office.load_key(password=SENHA_PLANILHA)
            buf = io.BytesIO()
            office.decrypt(buf)
    except Exception:
        return None  # outra senha / arquivo que não é a planilha de guias
    buf.seek(0)
    try:
        wb = openpyxl.load_workbook(buf, data_only=True, read_only=True)
    except Exception:
        return None
    linhas = []
    achou_cabecalho = False
    for linha in wb.worksheets[0].iter_rows(values_only=True):
        linha = tuple(linha) + (None,) * 8
        if not achou_cabecalho:
            if normalizar(linha[7]) == "NOME":
                achou_cabecalho = True
            continue
        if linha[7]:
            linhas.append(linha)
    wb.close()
    return linhas if achou_cabecalho else None


def ler_planilha(pasta):
    planilhas = [p for p in glob.glob(os.path.join(pasta, "*.xls*"))
                 if not os.path.basename(p).startswith("~$")]
    if not planilhas:
        raise SystemExit("ERRO: nenhuma planilha .xlsx encontrada na pasta.")
    # primeiro as que começam com "GUIAS", da mais recente para a mais antiga
    planilhas.sort(key=lambda p: (not normalizar(os.path.basename(p)).startswith("GUIAS"),
                                  -os.path.getmtime(p)))
    for caminho in planilhas:
        linhas = _linhas_da_planilha(caminho)
        if linhas:
            print("Planilha:", os.path.basename(caminho))
            break
    else:
        raise SystemExit("ERRO: não achei a planilha de guias (coluna H = NOME) na pasta.")

    ordem = []   # (nome_normalizado, processo) na ordem da planilha
    guias = []   # uma entrada por linha: processo, nome, natureza (col. A), valor (col. E)
    vistos = set()
    for linha in linhas:
        m = RE_PROCESSO.search(str(linha[6] or ""))
        proc = m.group(0) if m else ""
        nome = normalizar(linha[7])
        try:
            valor = float(linha[4])
        except (TypeError, ValueError):
            valor = None
        guias.append({"proc": proc, "nome": nome, "natureza": normalizar(linha[0]), "valor": valor})
        if (nome, proc) not in vistos:
            vistos.add((nome, proc))
            ordem.append((nome, proc))
    return ordem, guias


# ---------------------------------------------------------------------- PDFs
def nome_base(arquivo):
    """Tira numeração, ' - 1', ' (2)', ' - AC. MED.', ' - Cópia' etc. Devolve (base, nº dup ou None)."""
    s = os.path.splitext(arquivo)[0].strip()
    s = re.sub(r"^\d+\s*-\s*", "", s)                       # "01- "
    dup = None
    mudou = True
    while mudou:
        antes = s
        s = re.sub(r"\s*-\s*AC\.?\s*MED\.?\s*\.?$", "", s, flags=re.I)
        s = re.sub(r"\s*-\s*C[óo]pia(\s*\(\d+\))?$", "", s, flags=re.I)
        m = re.search(r"\s*(?:-\s*(\d{1,2})|\((\d{1,2})\))$", s)
        if m:
            dup = int(m.group(1) or m.group(2))
            s = s[:m.start()]
        mudou = s != antes
    return s.strip(), dup


def ler_pdf(caminho):
    """Devolve (maior 'Valor a ser atualizado' das memórias de cálculo, valor da guia)."""
    from pypdf import PdfReader
    iniciais = []
    valor_guia = None
    try:
        leitor = PdfReader(caminho)
        for i, pagina in enumerate(leitor.pages):
            texto = pagina.extract_text() or ""
            if valor_guia is None and i < 3:
                m = re.search(r"importe\s+total\s+de\s+R\$\s*([\d\.]+,\d{2})", texto, re.I)
                if m:
                    valor_guia = valor_br(m.group(1))
            if "atualizado" not in texto:
                continue
            for m in re.finditer(r"R\$\s*([\d\.]+,\d{2})\s*Valor a ser atualizado", texto, re.I):
                iniciais.append(valor_br(m.group(1)))
            for m in re.finditer(r"Valor a ser atualizado:?\s*R\$\s*([\d\.]+,\d{2})", texto, re.I):
                iniciais.append(valor_br(m.group(1)))
    except Exception as e:
        print("  ATENÇÃO: não consegui ler", os.path.basename(caminho), "-", e)
    return (max(iniciais) if iniciais else None), valor_guia


def reais(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def avaliar_ac_med(proc, nome, valor_inicial, valor_guia, guias):
    """Aplica a regra do AC. MED. Devolve (é_ac_med, valor_usado, natureza, motivo)."""
    do_processo = [g for g in guias if (proc and g["proc"] == proc) or (not proc and g["nome"] == nome)]
    linha = None
    if valor_guia is not None:
        linha = next((g for g in do_processo
                      if g["valor"] is not None and abs(g["valor"] - valor_guia) < 0.01), None)
    if linha is None and len(do_processo) == 1:
        linha = do_processo[0]
    natureza = linha["natureza"] if linha else ""

    eh_honorario = ("HONORARIO" in natureza) if linha else \
        (bool(do_processo) and all("HONORARIO" in g["natureza"] for g in do_processo))
    if eh_honorario:
        return False, valor_guia, natureza or "HONORARIOS", "guia de honorários"

    if valor_inicial is not None:
        return valor_inicial > LIMITE_AC_MED, valor_inicial, natureza, "valor inicial do cálculo"

    if valor_guia is None and linha:
        valor_guia = linha["valor"]
    if valor_guia is not None:
        return valor_guia > LIMITE_AC_MED, valor_guia, natureza, "valor da guia (sem memória de cálculo)"
    return False, None, natureza, "sem valor encontrado - conferir"


# --------------------------------------------------------------------- robô
def montar_plano(pasta):
    ordem, guias = ler_planilha(pasta)
    pos_nome = {}
    pos_proc = {}
    for i, (nome, proc) in enumerate(ordem):
        pos_nome.setdefault(nome, i)
        if proc:
            pos_proc.setdefault(proc, i)

    pdfs = sorted(f for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
    if not pdfs:
        raise SystemExit("ERRO: nenhum PDF na pasta.")

    itens = []
    for arq in pdfs:
        base, dup = nome_base(arq)
        m = RE_PROCESSO.search(base)
        proc = m.group(0) if m else ""
        nome_pdf = normalizar(base.split(" - ")[0])
        posicao = pos_nome.get(nome_pdf)
        if posicao is None and proc:
            posicao = pos_proc.get(proc)
        valor_inicial, valor_guia = ler_pdf(os.path.join(pasta, arq))
        ac_med, valor, natureza, motivo = avaliar_ac_med(proc, nome_pdf, valor_inicial, valor_guia, guias)
        itens.append({
            "arquivo": arq, "base": base, "dup": dup, "posicao": posicao,
            "chave": proc or normalizar(base), "valor": valor, "ac_med": ac_med,
            "natureza": natureza, "motivo": motivo,
        })

    # numeração " - 1", " - 2" para PDFs do mesmo processo
    grupos = {}
    for it in itens:
        grupos.setdefault(it["chave"], []).append(it)
    for grupo in grupos.values():
        grupo.sort(key=lambda x: (x["dup"] or 0, x["arquivo"]))
        for n, it in enumerate(grupo, 1):
            it["sufixo_dup"] = f" - {n}" if len(grupo) > 1 else ""

    # ordem final: AC. MED. primeiro, depois ordem da planilha, depois os que não estão na planilha
    def ordenacao(it):
        fora = it["posicao"] is None
        return (0 if it["ac_med"] else 1, fora, it["posicao"] if not fora else 0,
                it["base"], it["sufixo_dup"])
    itens.sort(key=ordenacao)

    digitos = max(2, len(str(len(itens))))
    for n, it in enumerate(itens, 1):
        novo = f"{str(n).zfill(digitos)}- {it['base']}{it['sufixo_dup']}"
        if it["ac_med"]:
            novo += SUFIXO_AC_MED
        it["novo"] = novo + ".pdf"
    return itens


def renomear(pasta, pares):
    """Renomeia em 2 etapas (nome temporário) para não colidir nomes."""
    temporarios = []
    for i, (de, para) in enumerate(pares):
        if de == para:
            continue
        tmp = f"__tmp_enum_{i}__.pdf"
        os.rename(os.path.join(pasta, de), os.path.join(pasta, tmp))
        temporarios.append((tmp, para))
    for tmp, para in temporarios:
        os.rename(os.path.join(pasta, tmp), os.path.join(pasta, para))


def executar(pasta, simular):
    print("Pasta:", pasta)
    itens = montar_plano(pasta)
    print()
    for it in itens:
        extra = []
        if it["ac_med"]:
            extra.append(f"AC. MED. ({it['motivo']}: {reais(it['valor'])})")
        elif it["motivo"] == "guia de honorários" and it["valor"] and it["valor"] > LIMITE_AC_MED:
            extra.append("honorários - não é AC. MED.")
        elif it["motivo"].startswith("sem valor"):
            extra.append(it["motivo"].upper())
        if it["posicao"] is None:
            extra.append("NÃO ACHEI NA PLANILHA")
        print(f"  {it['novo']}" + (f"   <-- {' | '.join(extra)}" if extra else ""))
    print()
    fora = [it for it in itens if it["posicao"] is None]
    if fora:
        print(f"ATENÇÃO: {len(fora)} PDF(s) não encontrado(s) na coluna H da planilha (foram para o final).")
    if simular:
        print("SIMULAÇÃO: nada foi renomeado.")
        return

    pares = [(it["arquivo"], it["novo"]) for it in itens]
    renomear(pasta, pares)
    log = os.path.join(pasta, NOME_LOG)
    with open(log, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["NOME ANTIGO", "NOME NOVO", "VALOR USADO", "AC. MED.", "NATUREZA (COL. A)", "CRITERIO", "DATA"])
        agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        for it in itens:
            w.writerow([it["arquivo"], it["novo"],
                        "" if it["valor"] is None else f"{it['valor']:.2f}".replace(".", ","),
                        "SIM" if it["ac_med"] else "", it["natureza"], it["motivo"], agora])
    print(f"OK: {len(itens)} PDF(s) renomeado(s). Log salvo em {NOME_LOG}")


def desfazer(pasta):
    log = os.path.join(pasta, NOME_LOG)
    if not os.path.exists(log):
        raise SystemExit("Não há log para desfazer nesta pasta.")
    with open(log, encoding="utf-8-sig") as f:
        linhas = list(csv.reader(f, delimiter=";"))[1:]
    pares = [(novo, antigo) for antigo, novo, *_ in linhas if os.path.exists(os.path.join(pasta, novo))]
    renomear(pasta, pares)
    os.remove(log)
    print(f"Desfeito: {len(pares)} arquivo(s) voltaram ao nome original.")


def escolher_pasta():
    try:
        import tkinter
        from tkinter import filedialog
        raiz = tkinter.Tk()
        raiz.withdraw()
        pasta = filedialog.askdirectory(title="Escolha a pasta do MALOTE (PDFs + planilha)")
        raiz.destroy()
        return pasta
    except Exception:
        return input("Cole o caminho da pasta do malote: ").strip().strip('"')


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a.lower() for a in sys.argv[1:] if a.startswith("--")}
    dependencias()
    pasta = args[0] if args else escolher_pasta()
    if not pasta or not os.path.isdir(pasta):
        raise SystemExit("Pasta inválida.")
    if "--desfazer" in flags:
        desfazer(pasta)
    else:
        executar(pasta, simular="--simular" in flags)


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0):
            print(e.code)
    except Exception as e:
        print("ERRO:", e)
    if sys.stdin and sys.stdin.isatty() and os.environ.get("ENUM_SEM_PAUSA") != "1":
        input("\nPressione ENTER para fechar...")

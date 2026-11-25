# -*- coding: utf-8 -*-
"""
Triagem automática de publicações jurídicas (aba RELATÓRIO -> coluna A: MANTER / DESCARTE).

Todos os nomes de clientes, escritórios e advogados deste arquivo são FICTÍCIOS.

Uso:
    Dois cliques no arquivo: usa a PLANILHA DE TRATAMENTO e a BASE PROCESSOS mais recentes da mesma pasta.
    Ou pelo terminal:
    python triagem_publicacoes.py "PLANILHA DE TRATAMENTO - 06.11.2026.xlsx" "BASE PROCESSOS 06.11.2026.xlsx" [saida.xlsx]

- Lê os valores já calculados da planilha de tratamento (aba ROBÔ com o teor das publicações).
- Lê a base de processos cadastrados direto do Excel (não depende de link externo).
- Classifica cada publicação por regras de negócio (cliente, tribunal, advogado citado no teor,
  número CNJ de processo originário citado no teor).
- Grava uma CÓPIA da planilha (o original não é alterado) via Excel, preservando fórmulas,
  validações e links:
    * coluna A preenchida (MANTER / DESCARTE; em branco = revisar à mão)
    * coluna D trocada para o cliente correto em casos específicos
    * colunas O:T com o que a automação fez e por quê
    * aba "RESUMO AUTOMAÇÃO"
"""
import collections
import json
import os
import re
import subprocess
import sys
import unicodedata
import warnings

import openpyxl

warnings.filterwarnings("ignore")

CLIENTE_ESPECIAL = "ÁGUAS DO VALE S.A."
NAO_CAD = "PROCESSO NÃO CADASTRADO NO SISTEMA"

# Clientes da casa que, achados em STJ/STF/agravo, indicam que a publicação é do escritório
CLIENTES_ALVO = {
    "LUZ NORTE": ["LUZ NORTE DISTRIBUIDORA", "LUZ NORTE S.A", "LUZ NORTE S/A"],
    "FORCA SUL": ["FORCA SUL ENERGIA", "FORCA SUL S.A"],
    "ENERGIA CENTRO": ["ENERGIA CENTRO DISTRIBUICAO", "ENERGIA CENTRO"],
    "AGUAS DO VALE": ["AGUAS DO VALE", "COMPANHIA AGUAS DO VALE"],
}

ADVS_CASA = {"SOCIO A - NÃO CADASTRADO", "SOCIO B - NÃO CADASTRADO",
             "SOCIO C - NÃO CADASTRADO", "OUTROS ADVS DA CASA - NÃO CADASTRADO"}
OUTROS_ESCRITORIOS = {"ESCRITORIO ALFA", "ESCRITORIO BETA", "ESCRITORIO GAMA", "ESCRITORIO DELTA"}

# Advogados do escritório procurados no teor (sem acento) - nomes fictícios
ADVS_ESCRITORIO = [
    "SOCIO A", "SOCIO B", "ANA BEATRIZ EXEMPLO", "BRUNO FICTICIO SANTOS", "CARLA MODELO RIBEIRO",
    "DANIEL TESTE OLIVEIRA", "EDUARDA AMOSTRA LIMA", "FABIO DEMO PEREIRA", "GABRIELA SIMULADA COSTA",
    "HENRIQUE PROTOTIPO ALVES",
]


def sem_acento(s):
    s = unicodedata.normalize("NFKD", str(s or ""))
    return "".join(c for c in s if not unicodedata.combining(c)).upper()


def adv_escritorio_no_teor(teor):
    t = sem_acento(teor)
    for n in ADVS_ESCRITORIO:
        if n in t:
            return n
    return None


# Coluna E (advogado localizado) com esses valores -> DESCARTE direto, antes de qualquer outra regra
DESCARTE_DIRETO_ADV = {"ESCRITORIO BETA", "ESCRITORIO DELTA", "ENERGIA CENTRO", "ESCRITORIO ALFA"}

# Número CNJ: NNNNNNN-DD.AAAA.J.TR.OOOO (aceita espaços/pontos faltando)
CNJ_RE = re.compile(r"(?<!\d)(\d{7})[-.\s]?(\d{2})[.\s]?(\d{4})[.\s]?(\d)[.\s]?(\d{2})[.\s]?(\d{4})(?!\d)")


def digitos(s):
    return re.sub(r"\D", "", str(s or ""))


def fmt_cnj(d):
    return f"{d[0:7]}-{d[7:9]}.{d[9:13]}.{d[13]}.{d[14:16]}.{d[16:20]}"


def normaliza_diario(d):
    """DJEN-TJRJ -> TJRJ, DJ-RJ -> TJRJ, DJEN-TRF1-GO -> TRF1, DJE-STF -> STF."""
    d = sem_acento(d).strip()
    for p in ("DJEN-", "DJE-"):
        if d.startswith(p):
            d = d[len(p):]
    if d.startswith("DJ-"):
        d = "TJ" + d[3:]
    elif re.fullmatch(r"DJ[A-Z]{2}", d):  # DJRJ, DJGO (sem hífen)
        d = "TJ" + d[2:]
    if d.startswith("TRF") and "-" in d:
        d = d.split("-")[0]
    return d


def cliente_alvo(texto):
    t = sem_acento(texto)
    for nome, chaves in CLIENTES_ALVO.items():
        if any(k in t for k in chaves):
            return nome
    return None


def categoria(proc, diario_n, cliente, adv):
    """Classifica a publicação (mesma lógica da coluna de conferência da planilha)."""
    if cliente and cliente != NAO_CAD:
        return "OK - LOCALIZADO"
    if str(proc).strip().endswith("0000"):
        return "AGRAVO - CONFERIR"
    if diario_n == "STJ":
        return "CONFERIR STJ"
    if diario_n == "STF":
        return "CONFERIR STF"
    if adv in ADVS_CASA:
        return "CONFERIR PROCESSO - COM ADVS DA CASA"
    if adv == "ESCRITORIO PARCEIRO - NÃO CADASTRADO":
        return "CHECK PARCEIRO"
    if adv == "ENERGIA CENTRO":
        return "DESCARTE - ENERGIA CENTRO"
    if diario_n in ("TJGO", "TJDFT", "TJDF", "DFT", "TRF1"):
        return "TJGO/DFT - VERIFICAR INCLUSÃO"
    if adv in OUTROS_ESCRITORIOS:
        return "DESCARTE - OUTROS ESCRITÓRIOS"
    if adv == "ESCRITORIO ZETA":
        return "INCLUIR - ZETA"
    if diario_n.startswith("TRT") or diario_n.startswith("TST"):
        return "DESCARTE - PROCESSO TRABALHISTA"
    return "DESCARTE - SEM ADVS DA CASA, SEM ADVS OUTROS ESCRITÓRIOS"


def carrega_base(path):
    print("Lendo base:", os.path.basename(path))
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Relatório"] if "Relatório" in wb.sheetnames else wb.worksheets[0]
    base = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 6 or row[1] is None:
            continue
        k = digitos(row[1])
        if k and k not in base:  # igual ao PROCV: fica a primeira ocorrência
            base[k] = {"cnj": str(row[1]).strip(), "cliente": row[2] or "", "carteira": row[3] or "", "resp": row[5] or ""}
    wb.close()
    print(f"  {len(base)} processos na base")
    return base


def processos_no_teor(teor, proprio, base):
    """Números CNJ citados no teor (exceto o próprio) que existem na base."""
    achados, vistos = [], {digitos(proprio)}
    for m in CNJ_RE.finditer(str(teor or "")):
        d = "".join(m.groups())
        if d in vistos:
            continue
        vistos.add(d)
        if d in base:
            achados.append(base[d])
    return achados


def decide(cat, proc, teor, adv, diario_n, base, cliente=""):
    """Retorna (decisao, motivo, processo_teor, cliente_teor, novo_D). decisao: MANTER / DESCARTE / REVISAR."""
    if cat == "OK - LOCALIZADO":
        return "MANTER", "Processo cadastrado na base", "", "", None

    if cat in ("CONFERIR STJ", "CONFERIR STF", "AGRAVO - CONFERIR"):
        achados = processos_no_teor(teor, proc, base)
        t = sem_acento(teor)
        # 1) originário de cliente que não é mais atendido -> descarte
        fora = [a for a in achados if "ENERGIA CENTRO" in sem_acento(a["cliente"])]
        if fora:
            a = fora[0]
            return "DESCARTE", "Originário da ENERGIA CENTRO - descarte", a["cnj"], a["cliente"], None
        if "ENERGIA CENTRO" in t:
            return "DESCARTE", "Teor cita ENERGIA CENTRO - descarte", "", "", None
        # 2) originário achado na base (qualquer cliente) -> MANTER
        if achados:
            a = achados[0]
            return "MANTER", "Originário achado na base pelo teor", a["cnj"], a["cliente"], None
        # 3) não achou: advogado do escritório no teor -> MANTER
        nome = adv_escritorio_no_teor(teor)
        if nome:
            return "MANTER", f"Originário não achado; teor cita advogado do escritório ({nome})", "", "", None
        # 4) dúvida: teor cita cliente da casa -> fica em branco para revisão
        c = cliente_alvo(teor)
        if c:
            return "REVISAR", f"Originário não achado e sem advogado do escritório, mas teor cita {c} - conferir à mão", "", "", None
        if not str(teor or "").strip():
            return "REVISAR", "Teor vazio - conferir à mão", "", "", None
        # 5) nada encontrado -> descarte
        return "DESCARTE", "Originário não achado e teor sem advogado do escritório", "", "", None

    if cat == "CHECK PARCEIRO":
        return "DESCARTE", "Escritório parceiro - sempre descarte", "", "", None

    if cat in ("DESCARTE - OUTROS ESCRITÓRIOS", "DESCARTE - ENERGIA CENTRO", "DESCARTE - PROCESSO TRABALHISTA"):
        return "DESCARTE", cat, "", "", None

    if cat.startswith("TJGO/DFT"):
        return "DESCARTE", "TJGO/DFT - sempre descarte", "", "", None

    if cat == "INCLUIR - ZETA":
        return "MANTER", f"Escritório Zeta - cliente alterado para {CLIENTE_ESPECIAL}", "", "", CLIENTE_ESPECIAL

    if cat.startswith("DESCARTE - SEM ADVS"):
        return "DESCARTE", "Sem advs do escritório - descarte direto", "", "", None

    if cat == "CONFERIR PROCESSO - COM ADVS DA CASA":
        # TRT, TST e TJGO -> descarte direto
        if diario_n.startswith(("TRT", "TST")) or diario_n == "TJGO":
            return "DESCARTE", f"Com advs da casa, mas diário {diario_n} - descarte", "", "", None
        nome = adv_escritorio_no_teor(teor)
        if nome:
            return "MANTER", f"Teor cita advogado do escritório ({nome}) - conferir", "", "", None
        return "DESCARTE", "Advogado do escritório não encontrado no teor - conferir", "", "", None

    return "REVISAR", "Categoria não prevista", "", "", None


DATA_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})")


def mais_recente(pasta, prefixo):
    """Arquivo .xlsx mais recente da pasta que começa com o prefixo (ignora saídas e temporários do Excel)."""
    achados = [os.path.join(pasta, f) for f in os.listdir(pasta)
               if f.upper().startswith(prefixo) and f.lower().endswith(".xlsx")
               and "AUTOMACAO" not in f.upper() and not f.startswith("~$")]
    if not achados:
        return None

    # prefere a data escrita no nome (dd.mm.aaaa); se não tiver, usa a data de modificação
    def chave(p):
        m = DATA_RE.search(os.path.basename(p))
        d = (int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else (0, 0, 0)
        return (d, os.path.getmtime(p))

    return max(achados, key=chave)


def data_do_nome(p):
    m = DATA_RE.search(os.path.basename(p or ""))
    return m.group(0) if m else None


def main():
    if len(sys.argv) >= 3:
        plan, base_path = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
    else:
        # dois cliques: procura os arquivos na mesma pasta do script
        pasta = os.path.dirname(os.path.abspath(__file__))
        plan = mais_recente(pasta, "PLANILHA DE TRATAMENTO")
        base_path = mais_recente(pasta, "BASE PROCESSOS")
        if not plan or not base_path:
            print("Coloque na mesma pasta do script a 'PLANILHA DE TRATAMENTO ... .xlsx' e a 'BASE PROCESSOS ... .xlsx'.")
            print("Pasta:", pasta)
            sys.exit(1)
        print("Planilha:", os.path.basename(plan))
        print("Base:    ", os.path.basename(base_path))
    dp, db = data_do_nome(plan), data_do_nome(base_path)
    if dp and db and dp != db:
        print(f"\n*** ATENÇÃO: planilha é de {dp} e a base é de {db}. Confira se é isso mesmo. ***\n")
    saida = os.path.abspath(sys.argv[3]) if len(sys.argv) > 3 else os.path.splitext(plan)[0] + " - AUTOMACAO.xlsx"
    # nunca sobrescreve: se já existir, cria "(2)", "(3)"...
    raiz, k = os.path.splitext(saida)[0], 2
    while os.path.exists(saida):
        saida = f"{raiz} ({k}).xlsx"
        k += 1

    base = carrega_base(base_path)

    print("Lendo planilha:", os.path.basename(plan))
    wb = openpyxl.load_workbook(plan, read_only=True, data_only=True)
    robo = list(wb["ROBÔ"].iter_rows(min_row=2, values_only=True))
    rel = list(wb["RELATÓRIO"].iter_rows(min_row=2, max_col=14, values_only=True))
    wb.close()
    n = 0
    for i, r in enumerate(robo):
        if r and r[0] not in (None, ""):
            n = i + 1
    print(f"  {n} publicações no ROBÔ")

    linhas, resumo, comp, cat_dif = [], collections.Counter(), collections.Counter(), collections.Counter()
    for i in range(n):
        r_robo = robo[i]
        r_rel = rel[i] if i < len(rel) else (None,) * 14
        a_orig, proc, diario, cliente, adv, f_planilha = (r_rel + (None,) * 14)[:6]
        if not proc:
            linhas.append(None)
            continue
        teor = r_robo[9] if len(r_robo) > 9 else ""
        diario_n = normaliza_diario(diario)
        # cliente: busca direta na base; se não achar, respeita o que estiver na planilha
        b = base.get(digitos(proc))
        cli = b["cliente"] if b else (cliente if cliente and cliente != NAO_CAD else NAO_CAD)
        cat = categoria(proc, diario_n, cli, adv or "")
        if (adv or "") in DESCARTE_DIRETO_ADV and cat != "OK - LOCALIZADO":
            dec, motivo, p_teor, c_teor, novo_d = "DESCARTE", f"Advogado localizado {adv} - descarte direto", "", "", None
        else:
            dec, motivo, p_teor, c_teor, novo_d = decide(cat, proc, teor, adv or "", diario_n, base, cli)
            # o que a planilha marcou como SEM ADVS nunca fica em branco
            if dec == "REVISAR" and str(f_planilha or "").startswith("DESCARTE - SEM ADVS"):
                dec, motivo = "DESCARTE", "Planilha marcou SEM ADVS - descarte (" + motivo + ")"
        if cat != (f_planilha or ""):
            cat_dif[(f_planilha, cat)] += 1
        resumo[(cat, dec)] += 1
        comp[(a_orig or "(vazio)", dec)] += 1
        linhas.append({"linha": i + 2, "A": dec if dec != "REVISAR" else None, "A_orig": a_orig,
                       "D": novo_d, "cat": cat, "dec": dec, "motivo": motivo, "p_teor": p_teor, "c_teor": c_teor})

    # ---- relatório no console
    print("\nRESULTADO POR CATEGORIA")
    for (cat, dec), q in sorted(resumo.items()):
        print(f"  {q:5}  {dec:8}  {cat}")
    tot = collections.Counter()
    for (cat, dec), q in resumo.items():
        tot[dec] += q
    print("\nTOTAL:", dict(tot))
    print("\nCOMPARAÇÃO coluna A atual x automação")
    for (o, d), q in sorted(comp.items()):
        print(f"  {q:5}  atual={o:10} automação={d}")
    print("\nCATEGORIAS QUE MUDARAM (planilha -> automação)")
    for (a, b), q in cat_dif.most_common():
        print(f"  {q:5}  {a}  ->  {b}")

    # ---- grava a cópia pelo Excel (script PowerShell ao lado)
    dados = {"origem": plan, "saida": saida, "linhas": [l for l in linhas if l],
             "resumo": [[c, d, q] for (c, d), q in sorted(resumo.items())]}
    js = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_resultado.json")
    with open(js, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False)
    ps1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aplicar_excel.ps1")
    print("\nGravando cópia pelo Excel (pode levar 1-2 min)...")
    rc = subprocess.call(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1, js])
    if rc != 0:
        print("ERRO ao gravar pelo Excel. Código", rc)
        sys.exit(rc)
    print("Pronto:", saida)


if __name__ == "__main__":
    dois_cliques = len(sys.argv) < 3
    codigo = 0
    try:
        main()
    except SystemExit as e:
        codigo = e.code or 0
    except Exception:
        import traceback

        traceback.print_exc()
        codigo = 1
    if dois_cliques:
        input("\nPressione Enter para fechar...")
    sys.exit(codigo)

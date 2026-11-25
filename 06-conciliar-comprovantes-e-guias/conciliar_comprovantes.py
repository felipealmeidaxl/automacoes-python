"""
AUTOMAÇÃO DE COMPROVANTES  (é o código empacotado no executável "Automação de Comprovantes.exe")

O que faz: na pasta onde estiver, junta cada COMPROVANTE (PDF com "COMPROVANTE" no nome) à sua guia/boleto,
renomeia o comprovante: se a guia tiver código após a indexação ("01_01_01_21500_...") fica só "21500.pdf";
senão, usa o nome da guia sem a indexação "01_02_" e move para a pasta RESULTADO.

Ordem de pareamento:
  1. Código de barras (comprovante x texto da guia)
  2. Valor + data  (boleto: vencimento igual | PIX: pago de 15 dias antes a 5 dias depois do vencimento)
  3. Valor + data + nome do beneficiário (quando sobra mais de uma guia)
  4. Eliminação: cada guia só é usada uma vez
  5. Empate/dúvida: NÃO move, só avisa
Também: nunca apaga arquivo em RESULTADO (numera "(2)", "(3)") e mostra o motivo de cada erro.

Rodar sem gerar exe:   python conciliar_comprovantes.py      (precisa: pip install PyPDF2)
Gerar o .exe:          pip install pyinstaller PyPDF2
                       python -m PyInstaller --onefile --console --name "Automação de Comprovantes" conciliar_comprovantes.py
                       (o .exe sai na pasta "dist")
"""
import ctypes
import glob
import os
import re
import shutil
import sys
import traceback
import unicodedata
from datetime import date

import PyPDF2

# Janela aceita entre a data de um PIX e o vencimento do boleto/guia
PIX_DIAS_ANTES_DO_VENC = 15   # pagou até 15 dias antes do vencimento
PIX_DIAS_DEPOIS_DO_VENC = 5   # ou até 5 dias depois (atraso)

PALAVRAS_IGNORADAS = {
    "BOLETO", "AUTOR", "AUTORIZACAO", "PARA", "PAGTO", "PAGAMENTO", "VENC", "EM",
    "LTDA", "SA", "S/A", "ME", "EIRELI", "DE", "DA", "DO", "DAS", "DOS", "E", "CNPJ",
}


# ------------------------------------------------------------------ leitura

def ler_texto_pdf(caminho_pdf):
    """Lê o texto do PDF. Retorna (texto, motivo_do_erro). Se deu certo, motivo_do_erro é None."""
    try:
        with open(caminho_pdf, 'rb') as arquivo:
            leitor = PyPDF2.PdfReader(arquivo)

            if leitor.is_encrypted:
                try:
                    if not leitor.decrypt(""):
                        return "", "PDF protegido por senha"
                except Exception:
                    return "", "PDF protegido por senha (não foi possível abrir)"

            texto_completo = ""
            for pagina in leitor.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_completo += texto + " "
    except PermissionError:
        return "", "arquivo aberto em outro programa ou sem permissão de leitura"
    except PyPDF2.errors.PdfReadError:
        return "", "arquivo PDF corrompido ou inválido"
    except Exception as e:
        return "", f"erro inesperado ao ler o PDF ({type(e).__name__}: {e})"

    if not texto_completo.strip():
        return "", "PDF sem texto (provavelmente é imagem/escaneado)"
    return texto_completo, None


def buscar_codigo_barras(texto_completo):
    """'Busca Agressiva' do código de barras (47 ou 48 números)."""
    # Divide o texto do PDF onde houver letras (A-Z); sobram blocos de números, espaços e símbolos.
    for bloco in re.split(r'[a-zA-Z]', texto_completo):
        numeros_limpos = re.sub(r'\D', '', bloco)
        if len(numeros_limpos) in (47, 48):
            return numeros_limpos
        elif len(numeros_limpos) > 48:
            ultimos_48 = numeros_limpos[-48:]
            ultimos_47 = numeros_limpos[-47:]
            return ultimos_48 if ultimos_48.startswith('8') else ultimos_47
    return None


def extrair_codigo_comprovante(caminho_pdf):
    """Mantido por compatibilidade. Retorna (codigo, motivo_do_erro)."""
    texto, erro = ler_texto_pdf(caminho_pdf)
    if erro:
        return None, erro
    codigo = buscar_codigo_barras(texto)
    if codigo:
        return codigo, None
    return None, "o PDF tem texto, mas nenhuma sequência de 47/48 números (código de barras) foi encontrada"


def extrair_todos_numeros_guia(caminho_pdf):
    """Lê a guia inteira e transforma TUDO em uma única sequência de números. Retorna (numeros, erro)."""
    texto, erro = ler_texto_pdf(caminho_pdf)
    if erro:
        return "", erro
    return re.sub(r'\D', '', texto), None


# ------------------------------------------------------------------ dados auxiliares

def valor_em_centavos(texto_valor):
    """'19.900,00' -> 1990000"""
    try:
        return int(texto_valor.replace('.', '').replace(',', ''))
    except (ValueError, AttributeError):
        return None


def formatar_valor(centavos):
    if centavos is None:
        return "?"
    reais = f"{centavos // 100:,}".replace(",", ".")
    return f"R$ {reais},{centavos % 100:02d}"


def para_data(d, m, a):
    try:
        return date(int(a), int(m), int(d))
    except ValueError:
        return None


def sem_acentos(texto):
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode().upper()


def palavras_nome(texto):
    palavras = re.findall(r"[A-Z0-9]+", sem_acentos(texto or ""))
    return [p for p in palavras if len(p) >= 2 and not p.isdigit() and p not in PALAVRAS_IGNORADAS]


def pontuacao_nome(nome_a, nome_b):
    """Conta quantas palavras batem entre os nomes ('TELEC' bate com 'TELECOMUNICACAO')."""
    a, b = palavras_nome(nome_a), palavras_nome(nome_b)
    pontos = 0
    for pa in a:
        for pb in b:
            curta, longa = sorted((pa, pb), key=len)
            if len(curta) >= 3 and longa.startswith(curta) or pa == pb:
                pontos += 1
                break
    return pontos


def dados_do_comprovante(texto):
    """Tira valor, vencimento, data de pagamento e beneficiário do texto do comprovante."""
    valor = None
    for padrao in (r'Valor Cobrado:\s*(?:R\$)?\s*([\d\.]+,\d{2})',   # layout de pagamento em lote
                   r'Valor total:\s*(?:R\$)?\s*([\d\.]+,\d{2})',     # layout de boleto do internet banking
                   r'Valor:?\s*R\$\s*([\d\.]+,\d{2})',               # PIX / "ValorR$ 1.240,90"
                   r'Valor do\s*Documento:\s*([\d\.]+,\d{2})'):
        m = re.search(padrao, texto, re.IGNORECASE)
        if m:
            valor = valor_em_centavos(m.group(1))
            break

    venc = None
    m = re.search(r'Vencimento:\s*(\d{2})/(\d{2})/(\d{4})', texto, re.IGNORECASE)
    if m:
        venc = para_data(*m.groups())

    pagamento = None
    for padrao in (r'Data de\s*Pagamento:\s*(\d{2})/(\d{2})/(\d{4})',
                   r'Data da opera\S*:\s*(\d{2})/(\d{2})/(\d{4})'):
        m = re.search(padrao, texto)
        if m:
            pagamento = para_data(*m.groups())
            break

    beneficiario = None
    m = re.search(r'Benefici\S*:\s*(.+?)\s*(?:CNPJ|CPF|\n)', texto)
    if not m:
        m = re.search(r'Nome:\s*(.+)', texto)
    if m:
        beneficiario = m.group(1).strip()

    return valor, venc, pagamento, beneficiario


def dados_da_guia(nome_arquivo, texto):
    """Valor, vencimento e nome tirados do NOME do arquivo (e, se faltar valor, do texto da guia)."""
    base = os.path.splitext(nome_arquivo)[0]

    valor = None
    m = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})\s*$', base)
    if m:
        valor = valor_em_centavos(m.group(1))
    elif texto:
        m = re.search(r'Valor (?:do\s*Documento|Cobrado|total)[^\d]{0,30}?([\d\.]+,\d{2})', texto, re.IGNORECASE)
        if m:
            valor = valor_em_centavos(m.group(1))

    venc = None
    m = re.search(r'VENC\.?\s*EM\s*(\d{2})[./](\d{2})[./](\d{4})', base, re.IGNORECASE)
    if m:
        venc = para_data(*m.groups())

    nome = re.sub(r'^(\d+_)+', '', base)
    nome = re.sub(r'VENC\.?\s*EM\s*\d{2}[./]\d{2}[./]\d{4}', ' ', nome, flags=re.IGNORECASE)
    nome = re.sub(r'\d{1,3}(?:\.\d{3})*,\d{2}\s*$', ' ', nome)
    return valor, venc, nome.replace('_', ' ').strip()


# ------------------------------------------------------------------ pareamento

def validar_par(codigo_comprovante, numeros_guia):
    """Validação inteligente para driblar números intrometidos da guia."""
    if not codigo_comprovante or not numeros_guia:
        return False

    if codigo_comprovante in numeros_guia:
        return True

    if len(codigo_comprovante) >= 45:
        p1 = codigo_comprovante[:15]
        p2 = codigo_comprovante[15:30]
        p3 = codigo_comprovante[30:]

        if p1 in numeros_guia and p2 in numeros_guia and p3 in numeros_guia:
            return True

    return False


def data_compativel(comp, guia):
    """Critério 2: data. Vencimento igual, OU pago perto da data 'VENC. EM' do nome do arquivo.
    (A data do nome do arquivo nem sempre é o vencimento impresso no boleto.)"""
    if guia["venc"] is None:
        return True  # guia sem data no nome: não dá para descartar por data
    if comp["venc"] is not None and comp["venc"] == guia["venc"]:
        return True
    if comp["pagamento"] is not None:
        dias = (guia["venc"] - comp["pagamento"]).days
        return -PIX_DIAS_DEPOIS_DO_VENC <= dias <= PIX_DIAS_ANTES_DO_VENC
    return True


def escolher_guia(comp, guias_livres):
    """Critérios 2 e 3. Retorna (guia, motivo) ou (None, explicação)."""
    if comp["valor"] is None:
        return None, "não encontrei o valor pago no comprovante"

    mesmo_valor = [g for g in guias_livres if g["valor"] == comp["valor"]]
    if not mesmo_valor:
        return None, f"nenhuma guia livre com o valor {formatar_valor(comp['valor'])}"

    # Critério 2: valor + data (vencimento igual, ou PIX pago perto do vencimento)
    candidatos = [g for g in mesmo_valor if data_compativel(comp, g)]
    if not candidatos:
        datas = ", ".join(sorted({g["venc"].strftime("%d/%m/%Y") for g in mesmo_valor if g["venc"]}))
        return None, (f"há guia(s) com {formatar_valor(comp['valor'])}, mas a data não bate "
                      f"(vencimento da guia: {datas or '?'})")
    if len(candidatos) > 1 and comp["venc"] is not None:
        # vencimento exatamente igual tem prioridade sobre "pago perto da data"
        exatos = [g for g in candidatos if g["venc"] == comp["venc"]]
        if exatos:
            candidatos = exatos
    if len(candidatos) == 1:
        return candidatos[0], "valor + data"

    # Critério 3: valor + data + nome do beneficiário
    pontos = sorted(((pontuacao_nome(comp["beneficiario"], g["nome"]), g) for g in candidatos),
                    key=lambda x: x[0], reverse=True)
    if pontos[0][0] > 0 and pontos[0][0] > pontos[1][0]:
        return pontos[0][1], "valor + data + nome do beneficiário"

    return None, "EMPATE: " + " | ".join(f"'{g['arquivo']}'" for g in candidatos)


def caminho_sem_sobrescrever(pasta, nome):
    """Se 'nome' já existir na pasta, devolve 'nome (2).pdf', 'nome (3).pdf'... (nunca apaga nada)."""
    base, ext = os.path.splitext(nome)
    destino = os.path.join(pasta, nome)
    n = 2
    while os.path.exists(destino):
        destino = os.path.join(pasta, f"{base} ({n}){ext}")
        n += 1
    return destino


def mover_para_resultado(comp, guia, pasta_resultado, criterio):
    # Se, depois da indexação ("01_01_01_"), a guia tiver um código numérico (ex.: "21500_"),
    # o comprovante fica SÓ com esse código:
    #   "01_01_01_21500_0001234-56.2020.8.00.0001.pdf" -> "21500.pdf"
    # Sem código: remove a indexação e mantém o resto do nome, como antes:
    #   "01_01_01_0009876-54.2019.8.00.0002.pdf"       -> "0009876-54.2019.8.00.0002.pdf"
    codigo = re.match(r'^(?:\d{1,3}_)+(\d{4,6})_', guia["arquivo"])
    if codigo:
        novo_nome_comp = codigo.group(1) + os.path.splitext(guia["arquivo"])[1]
    else:
        novo_nome_comp = re.sub(r'^(\d+_)+', '', guia["arquivo"])

    # NUNCA apaga um arquivo que já está em RESULTADO: se o nome existir, numera
    caminho_comp_destino = caminho_sem_sobrescrever(pasta_resultado, novo_nome_comp)
    nome_final = os.path.basename(caminho_comp_destino)

    print(f"✔️ Par encontrado! (por {criterio})")
    print(f"   Comprovante: {comp['arquivo']} -> {nome_final}")
    print(f"   Guia original (mantida): {guia['arquivo']}")
    if nome_final != novo_nome_comp:
        print(f"   ⚠️ Já existia '{novo_nome_comp}' em RESULTADO — salvo como '{nome_final}' para não apagar o anterior.")
    print(f"   * Movendo APENAS o comprovante limpo para a pasta 'RESULTADO'...\n")

    try:
        shutil.move(comp["caminho"], caminho_comp_destino)
        return True
    except PermissionError:
        print(f"   [!] Não consegui mover '{comp['arquivo']}': o arquivo está aberto em outro programa "
              f"ou sem permissão. Feche o PDF e rode de novo.\n")
    except Exception as e:
        print(f"   [!] Erro ao mover '{comp['arquivo']}': {type(e).__name__}: {e}\n")
    return False


def processar_comprovantes(diretorio, mover=True):
    comprovantes = []
    guias = []

    pasta_resultado = os.path.join(diretorio, "RESULTADO")
    if mover:
        os.makedirs(pasta_resultado, exist_ok=True)

    print("Lendo arquivos na pasta...\n")

    for arquivo in sorted(os.listdir(diretorio)):
        caminho_completo = os.path.join(diretorio, arquivo)
        if not (os.path.isfile(caminho_completo) and arquivo.lower().endswith('.pdf')):
            continue

        texto, erro = ler_texto_pdf(caminho_completo)

        if "COMPROVANTE" in arquivo.upper():
            if erro:
                print(f"⚠️ Não consegui ler o comprovante: {arquivo}")
                print(f"   Motivo: {erro}\n")
                continue
            codigo = buscar_codigo_barras(texto)
            valor, venc, pagamento, beneficiario = dados_do_comprovante(texto)
            if not codigo and valor is None:
                print(f"⚠️ Não consegui achar o código de barras nem o valor no comprovante: {arquivo}\n")
                continue
            comprovantes.append({"arquivo": arquivo, "caminho": caminho_completo, "codigo": codigo,
                                 "valor": valor, "venc": venc, "pagamento": pagamento,
                                 "beneficiario": beneficiario})
        else:
            valor, venc, nome = dados_da_guia(arquivo, texto)
            if erro:
                print(f"ℹ️ Guia sem texto legível: {arquivo}")
                print(f"   Motivo: {erro} — será pareada só pelo valor/data/nome do arquivo.\n")
            guias.append({"arquivo": arquivo, "caminho": caminho_completo,
                          "numeros": re.sub(r'\D', '', texto) if not erro else "",
                          "valor": valor, "venc": venc, "nome": nome})

    print("\n" + "-" * 40)
    print("Iniciando cruzamento de dados e movimentação...\n")

    usadas = set()       # Critério 4: cada guia só pode ser usada uma vez
    pendentes = []

    # Critério 1: código de barras
    for comp in comprovantes:
        guia = next((g for g in guias if validar_par(comp["codigo"], g["numeros"])), None)
        if guia:
            usadas.add(guia["arquivo"])
            if mover:
                mover_para_resultado(comp, guia, pasta_resultado, "código de barras")
            else:
                print(f"✔️ [código de barras] {comp['arquivo']} -> {guia['arquivo']}")
        else:
            pendentes.append(comp)

    # Critérios 2, 3 e 4: valor + data + nome, repetindo enquanto a eliminação resolver novos casos
    motivos = {}
    progresso = True
    while pendentes and progresso:
        progresso = False
        for comp in list(pendentes):
            livres = [g for g in guias if g["arquivo"] not in usadas]
            guia, motivo = escolher_guia(comp, livres)
            if guia:
                usadas.add(guia["arquivo"])
                pendentes.remove(comp)
                progresso = True
                if mover:
                    mover_para_resultado(comp, guia, pasta_resultado, motivo)
                else:
                    print(f"✔️ [{motivo}] {comp['arquivo']} -> {guia['arquivo']}")
            else:
                motivos[comp["arquivo"]] = motivo

    # Critério 5: se ainda houver dúvida, não adivinha — avisa
    for comp in pendentes:
        tipo = "boleto" if comp["codigo"] else "PIX/transferência"
        print(f"❌ Aviso: Nenhuma guia definida para '{comp['arquivo']}' ({tipo}, {formatar_valor(comp['valor'])}).")
        if comp["codigo"]:
            print(f"   Código procurado: {comp['codigo']}")
        print(f"   Motivo: {motivos.get(comp['arquivo'], 'sem candidatos')}")
        print(f"   → Separe este manualmente.\n")


# ================================================================== pop-ups e execução


TITULO = "Automação de Comprovantes"
MB_OK, MB_OKCANCEL = 0x0, 0x1
MB_ICONERROR, MB_ICONWARNING, MB_ICONINFO = 0x10, 0x30, 0x40
MB_TOPMOST, MB_SETFOREGROUND = 0x40000, 0x10000
IDOK = 1


def popup(texto, titulo, flags):
    return ctypes.windll.user32.MessageBoxW(None, texto, titulo, flags | MB_TOPMOST | MB_SETFOREGROUND)


def contar(padrao):
    return len([f for f in glob.glob(padrao) if os.path.isfile(f)])


def main():
    for fluxo in (sys.stdout, sys.stderr):
        try:
            fluxo.reconfigure(encoding="utf-8")
        except Exception:
            pass

    # Sempre trabalha na pasta onde o .exe está (não importa de onde foi aberto)
    pasta = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
    os.chdir(pasta)
    ctypes.windll.kernel32.SetConsoleTitleW(TITULO)

    pendentes_antes = contar("*COMPROVANTE*.pdf")  # no Windows o glob ignora maiúsculas/minúsculas
    if pendentes_antes == 0:
        popup("Nenhum comprovante (arquivo PDF com \"COMPROVANTE\" no nome) foi encontrado nesta pasta:\n\n"
              f"{pasta}", TITULO, MB_OK | MB_ICONWARNING)
        return 0

    resp = popup(f"A automação de comprovantes vai começar.\n\n"
                 f"Comprovantes encontrados: {pendentes_antes}\n\n"
                 "Clique em OK para iniciar ou Cancelar para sair.",
                 "Iniciando automação", MB_OKCANCEL | MB_ICONINFO)
    if resp != IDOK:
        return 0

    antes = contar(os.path.join("RESULTADO", "*.pdf"))
    print("=" * 40)
    print("  AUTOMAÇÃO DE COMPROVANTES EM EXECUÇÃO")
    print("=" * 40 + "\n")

    try:
        processar_comprovantes(".")
    except Exception:
        traceback.print_exc()
        popup("A automação terminou com ERRO.\n\nVeja os detalhes na janela preta.",
              "Automação com erro", MB_OK | MB_ICONERROR)
        input("\nPressione Enter para fechar...")
        return 1

    movidos = max(0, contar(os.path.join("RESULTADO", "*.pdf")) - antes)
    sem_par = contar("*COMPROVANTE*.pdf")
    icone = MB_ICONINFO if sem_par == 0 else MB_ICONWARNING
    popup("Automação finalizada!\n\n"
          f"Comprovantes movidos para RESULTADO: {movidos}\n"
          f"Comprovantes não processados (sem guia ou com erro): {sem_par}"
          + ("\n\nVeja o motivo de cada um na janela preta." if sem_par else ""),
          "Automação finalizada", MB_OK | icone)
    input("\nPressione Enter para fechar...")
    return 0


if __name__ == "__main__":
    sys.exit(main())

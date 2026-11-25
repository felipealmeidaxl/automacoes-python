"""
Pareia comprovantes de pagamento com as guias/boletos pelo código de barras.

Na pasta onde for executado:
  1. Lê todos os PDFs; os que têm "COMPROVANTE" na 1ª página são comprovantes, o resto são guias
  2. Extrai a linha digitável (47/48 dígitos) do comprovante e procura dentro dos números da guia
  3. Renomeia o comprovante com o nome da guia (sem a indexação "123__") e move para RESULTADO
  4. Nunca sobrescreve arquivos e grava um relatório .txt com pendências e erros
"""
import os
import re
import PyPDF2
import shutil
import datetime
import traceback
import tkinter as tk
from tkinter import messagebox


def ler_texto_pdf(caminho_pdf):
    """Lê o PDF uma única vez e devolve o texto de todas as páginas.
    Se o arquivo não puder ser lido, a exceção sobe para ser registrada no relatório."""
    with open(caminho_pdf, 'rb') as arquivo:
        leitor = PyPDF2.PdfReader(arquivo)
        return [(pagina.extract_text() or "") for pagina in leitor.pages]


def eh_comprovante(texto_primeira_pagina):
    """Identifica se o PDF é um comprovante pela primeira página."""
    return "COMPROVANTE" in texto_primeira_pagina.upper()


def extrair_codigo_comprovante(texto_completo):
    """Lê o comprovante usando uma 'Busca Agressiva' para driblar a miopia do PyPDF2."""
    blocos_sem_letras = re.split(r'[a-zA-Z]', texto_completo)

    for bloco in blocos_sem_letras:
        numeros_limpos = re.sub(r'\D', '', bloco)

        if len(numeros_limpos) == 47 or len(numeros_limpos) == 48:
            return numeros_limpos
        elif len(numeros_limpos) > 48:
            ultimos_48 = numeros_limpos[-48:]
            ultimos_47 = numeros_limpos[-47:]
            if ultimos_48.startswith('8'):
                return ultimos_48
            else:
                return ultimos_47
    return None


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


def caminho_destino_livre(pasta, nome_arquivo):
    """Nunca sobrescreve: se o nome já existir, usa 'nome_2.pdf', 'nome_3.pdf'..."""
    destino = os.path.join(pasta, nome_arquivo)
    base, extensao = os.path.splitext(nome_arquivo)
    contador = 2
    while os.path.exists(destino):
        destino = os.path.join(pasta, f"{base}_{contador}{extensao}")
        contador += 1
    return destino


def processar_comprovantes(diretorio):
    comprovantes = []
    guias = []

    relatorio = {
        "pareados": [],          # (comprovante, nome final)
        "sem_guia": [],          # comprovantes sem guia correspondente
        "sem_codigo": [],        # tinham "COMPROVANTE" mas nenhum código; tratados como guia
        "renomeados": [],        # nome já existia em RESULTADO; recebeu sufixo
        "erros": [],             # (arquivo, mensagem)
    }

    pasta_resultado = os.path.join(diretorio, "RESULTADO")
    os.makedirs(pasta_resultado, exist_ok=True)

    for arquivo in os.listdir(diretorio):
        caminho_completo = os.path.join(diretorio, arquivo)

        if not (os.path.isfile(caminho_completo) and arquivo.lower().endswith('.pdf')):
            continue

        try:
            paginas = ler_texto_pdf(caminho_completo)
        except Exception as erro:
            relatorio["erros"].append((arquivo, f"não foi possível ler o PDF: {erro}"))
            continue

        texto_completo = " ".join(paginas)
        primeira_pagina = paginas[0] if paginas else ""

        if eh_comprovante(primeira_pagina):
            codigo = extrair_codigo_comprovante(texto_completo)
            if codigo:
                comprovantes.append((arquivo, caminho_completo, codigo))
                continue
            # Tem a palavra "COMPROVANTE" mas nenhum código: provavelmente é uma
            # guia com essa palavra no texto. Em vez de sumir, entra como guia.
            relatorio["sem_codigo"].append(arquivo)

        numeros_guia = re.sub(r'\D', '', texto_completo)
        guias.append((arquivo, caminho_completo, numeros_guia))

    for comp_nome, comp_caminho, codigo in comprovantes:
        par_encontrado = False

        for guia_nome, guia_caminho, numeros_guia in guias:
            if validar_par(codigo, numeros_guia):
                par_encontrado = True

                # Remove os números seguidos de "__" no começo do nome da guia, se existirem
                novo_nome_comp = re.sub(r'^\d+__', '', guia_nome)
                caminho_comp_destino = caminho_destino_livre(pasta_resultado, novo_nome_comp)
                nome_final = os.path.basename(caminho_comp_destino)

                try:
                    shutil.move(comp_caminho, caminho_comp_destino)
                    relatorio["pareados"].append((comp_nome, nome_final))
                    if nome_final != novo_nome_comp:
                        relatorio["renomeados"].append((comp_nome, nome_final))
                except Exception as erro:
                    relatorio["erros"].append((comp_nome, f"falha ao mover para RESULTADO: {erro}"))

                break

        if not par_encontrado:
            relatorio["sem_guia"].append((comp_nome, codigo))

    return relatorio, pasta_resultado


def salvar_relatorio(relatorio, pasta_resultado):
    agora = datetime.datetime.now()
    caminho = os.path.join(pasta_resultado, f"relatorio_{agora:%Y-%m-%d_%H-%M-%S}.txt")

    linhas = [f"Relatório da automação - {agora:%d/%m/%Y %H:%M:%S}", ""]

    linhas.append(f"PAREADOS E MOVIDOS ({len(relatorio['pareados'])}):")
    for comp, final in relatorio["pareados"]:
        linhas.append(f"  {comp}  ->  {final}")
    linhas.append("")

    linhas.append(f"COMPROVANTES SEM GUIA CORRESPONDENTE ({len(relatorio['sem_guia'])}) - continuam na pasta original:")
    for comp, codigo in relatorio["sem_guia"]:
        linhas.append(f"  {comp}  (código lido: {codigo})")
    linhas.append("")

    linhas.append(f"NOME JÁ EXISTIA EM RESULTADO, SALVO COM SUFIXO ({len(relatorio['renomeados'])}) - conferir:")
    for comp, final in relatorio["renomeados"]:
        linhas.append(f"  {comp}  ->  {final}")
    linhas.append("")

    linhas.append(f"TÊM 'COMPROVANTE' NO TEXTO MAS SEM CÓDIGO - TRATADOS COMO GUIA ({len(relatorio['sem_codigo'])}):")
    for arq in relatorio["sem_codigo"]:
        linhas.append(f"  {arq}")
    linhas.append("")

    linhas.append(f"ERROS ({len(relatorio['erros'])}):")
    for arq, msg in relatorio["erros"]:
        linhas.append(f"  {arq}: {msg}")

    with open(caminho, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))
    return caminho


# ==========================================
pasta_alvo = r"."

if __name__ == "__main__":
    # Cria uma janela invisível apenas para ancorar o pop-up
    root = tk.Tk()
    root.withdraw()

    try:
        relatorio, pasta_resultado = processar_comprovantes(pasta_alvo)
        caminho_relatorio = salvar_relatorio(relatorio, pasta_resultado)

        resumo = (
            f"Pareados e movidos: {len(relatorio['pareados'])}\n"
            f"Comprovantes sem guia: {len(relatorio['sem_guia'])}\n"
            f"Salvos com sufixo (nome repetido): {len(relatorio['renomeados'])}\n"
            f"Erros: {len(relatorio['erros'])}\n\n"
            f"Detalhes em:\n{os.path.basename(caminho_relatorio)} (pasta RESULTADO)"
        )

        houve_problema = relatorio["sem_guia"] or relatorio["erros"] or relatorio["renomeados"]
        if houve_problema:
            messagebox.showwarning(title="Automação concluída - verificar pendências", message=resumo)
        else:
            messagebox.showinfo(title="Automação concluída", message=resumo)

    except Exception:
        messagebox.showerror(
            title="Automação falhou",
            message="Ocorreu um erro inesperado:\n\n" + traceback.format_exc()
        )

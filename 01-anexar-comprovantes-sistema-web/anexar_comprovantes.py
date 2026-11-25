"""
Anexa comprovantes de pagamento (PDF) aos processos de um sistema jurídico web.

Fluxo:
  1. Faz login no sistema com as credenciais do arquivo .env
  2. Para cada PDF da pasta de entrada, lê o número do processo no início do nome do arquivo
     (ex.: "123456 - comprovante.pdf" -> processo 123456)
  3. Abre o processo, entra em "Documentos", preenche a descrição, escolhe o tipo
     e envia o arquivo
  4. Ao final mostra um relatório (sucesso, parada pelo ESC ou queda do sistema)

A tecla ESC interrompe a automação com segurança entre um arquivo e outro.
"""
import os
import re
import time
import tkinter as tk
from tkinter import messagebox

import keyboard
from dotenv import find_dotenv, load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# ================= CONFIGURAÇÕES (vêm do arquivo .env) =================
load_dotenv(find_dotenv())
URL_SISTEMA = os.getenv("SISTEMA_URL", "https://sistema-juridico.exemplo.com.br")
USUARIO = os.getenv("SISTEMA_USUARIO", "")
SENHA = os.getenv("SISTEMA_SENHA", "")
PASTA_ARQUIVOS = os.getenv("PASTA_COMPROVANTES", "./entrada")

DESCRICAO_DOCUMENTO = "COMPROVANTE DE PAGAMENTO"
TIPO_DOCUMENTO = os.getenv("SISTEMA_TIPO_DOCUMENTO", "COMPROVANTE")

# Seletores dos campos do sistema (ajuste conforme o sistema usado)
SELETOR = {
    "login_usuario": (By.ID, "login"),
    "login_senha": (By.ID, "password"),
    "busca_processo": (By.NAME, "numeroProcesso"),
    "botao_consultar": (By.XPATH, "//input[@value='Consultar']"),
    "aba_documentos": (By.XPATH, "//*[contains(text(), 'Documentos')]"),
    "descricao": (By.NAME, "descricao"),
    "tipo_documento": (By.ID, "tipoDocumento"),
    "arquivo": (By.NAME, "arquivo"),
    "botao_enviar": (By.XPATH, "//input[@value='Enviar Arquivo' or @type='submit']"),
    "pagina_inicial": (By.XPATH, "//*[contains(text(), 'Página Inicial')]"),
}

parar_automacao = False
# =======================================================================


def acionar_parada():
    global parar_automacao
    parar_automacao = True


keyboard.add_hotkey("esc", acionar_parada)


def cacar_elemento(driver, by, value, timeout=10):
    """Procura o elemento na página e dentro de todos os frames/iframes até o tempo acabar."""
    tempo_limite = time.time() + timeout
    while time.time() < tempo_limite:
        try:
            driver.switch_to.default_content()
            return driver.find_element(by, value)
        except Exception:
            pass

        try:
            driver.switch_to.default_content()
            frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
            for frame in frames:
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(frame)
                    return driver.find_element(by, value)
                except Exception:
                    pass
        except Exception:
            pass

        time.sleep(0.5)
    raise Exception(f"Tempo esgotado. Não encontrou o elemento: '{value}'")


def main():
    global parar_automacao

    root = tk.Tk()
    root.withdraw()

    if not USUARIO or not SENHA:
        messagebox.showerror("Configuração", "Preencha SISTEMA_USUARIO e SISTEMA_SENHA no arquivo .env")
        return

    servico = Service(ChromeDriverManager().install())
    opcoes = webdriver.ChromeOptions()
    # Esconde as mensagens técnicas do Chrome no terminal
    opcoes.add_experimental_option("excludeSwitches", ["enable-logging"])
    driver = webdriver.Chrome(service=servico, options=opcoes)

    driver.get(URL_SISTEMA)
    driver.maximize_window()

    # ================= LOGIN AUTOMÁTICO =================
    try:
        campo_usuario = WebDriverWait(driver, 15).until(EC.presence_of_element_located(SELETOR["login_usuario"]))
        campo_usuario.clear()
        campo_usuario.send_keys(USUARIO)

        campo_senha = driver.find_element(*SELETOR["login_senha"])
        campo_senha.clear()
        campo_senha.send_keys(SENHA)
        campo_senha.send_keys(Keys.RETURN)

        time.sleep(5)
    except Exception as e:
        messagebox.showerror("Erro de Login", f"Não foi possível fazer o login automático.\nErro: {e}")
        driver.quit()
        return
    # ====================================================

    arquivos = [f for f in os.listdir(PASTA_ARQUIVOS) if f.lower().endswith(".pdf")]

    if not arquivos:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF encontrado na pasta.")
        driver.quit()
        return

    comprovantes_inseridos = []
    motivo_parada = "SUCESSO"  # controla o que será exibido no relatório final
    erro_detalhe = ""

    for nome_arquivo in arquivos:
        if parar_automacao:
            motivo_parada = "ESC"
            break

        caminho_completo = os.path.abspath(os.path.join(PASTA_ARQUIVOS, nome_arquivo))

        match = re.search(r"^(\d+)", nome_arquivo)
        if not match:
            continue

        numero_processo = match.group(1)

        try:
            janela_principal = driver.current_window_handle

            campo_busca = cacar_elemento(driver, *SELETOR["busca_processo"])
            campo_busca.clear()
            campo_busca.send_keys(numero_processo)

            cacar_elemento(driver, *SELETOR["botao_consultar"]).click()
            cacar_elemento(driver, *SELETOR["aba_documentos"]).click()

            # a tela de documentos abre em uma nova janela
            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
            for janela in driver.window_handles:
                if janela != janela_principal:
                    driver.switch_to.window(janela)
                    break

            campo_descricao = cacar_elemento(driver, *SELETOR["descricao"])
            campo_descricao.clear()
            campo_descricao.send_keys(DESCRICAO_DOCUMENTO)

            Select(cacar_elemento(driver, *SELETOR["tipo_documento"])).select_by_visible_text(TIPO_DOCUMENTO)

            cacar_elemento(driver, *SELETOR["arquivo"]).send_keys(caminho_completo)
            cacar_elemento(driver, *SELETOR["botao_enviar"]).click()
            time.sleep(3)

            if len(driver.window_handles) > 1:
                driver.close()

            driver.switch_to.window(janela_principal)
            comprovantes_inseridos.append(numero_processo)

            cacar_elemento(driver, *SELETOR["pagina_inicial"]).click()
            time.sleep(2)

        except Exception as e:
            # tenta fechar janelas perdidas se o sistema travou no meio do processo
            try:
                if len(driver.window_handles) > 1:
                    driver.switch_to.window(driver.window_handles[-1])
                    driver.close()
            except Exception:
                pass

            detalhe = str(e).split("Stacktrace")[0].strip()
            if len(detalhe) > 400:
                detalhe = detalhe[:400] + "..."
            erro_detalhe = f"Arquivo: {nome_arquivo}\nProcesso: {numero_processo}\nErro: {detalhe}"

            # para imediatamente para não seguir com o sistema instável
            motivo_parada = "QUEDA_SITE"
            break

    driver.quit()

    # ================= RELATÓRIO FINAL =================
    if not comprovantes_inseridos:
        texto_inseridos = "Nenhum comprovante foi inserido nesta sessão."
    else:
        texto_inseridos = "Processos atualizados com sucesso:\n" + "\n".join(comprovantes_inseridos)

    if motivo_parada == "ESC":
        messagebox.showwarning("Automação Interrompida", f"Você parou o robô apertando a tecla ESC!\n\n{texto_inseridos}")
    elif motivo_parada == "QUEDA_SITE":
        messagebox.showerror(
            "Falha de conexão com o sistema",
            "O sistema demorou muito para responder ou caiu. A automação foi pausada por segurança.\n\n"
            f"Onde parou:\n{erro_detalhe}\n\n{texto_inseridos}",
        )
    else:
        messagebox.showinfo("Automação Concluída", f"Todos os arquivos da pasta foram processados!\n\n{texto_inseridos}")


if __name__ == "__main__":
    try:
        main()
    except Exception as erro:
        # qualquer erro não previsto aparece em pop-up em vez de ficar só no terminal
        import traceback

        janela = tk.Tk()
        janela.withdraw()
        messagebox.showerror(
            "Erro inesperado na automação",
            f"A automação parou por um erro inesperado:\n\n{erro}\n\nDetalhes técnicos:\n{traceback.format_exc()[-800:]}",
        )
        janela.destroy()

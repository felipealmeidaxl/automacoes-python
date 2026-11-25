"""
Lança despesas de custas judiciais em um sistema jurídico web a partir de uma planilha Excel.

Fluxo:
  0. Calcula o próximo dia útil e pergunta se é feriado (pula feriados seguidos)
  1. Abre a planilha .xlsx mais recente da pasta configurada
  2. Faz login no sistema com as credenciais do arquivo .env
  3. Para cada linha: pesquisa o processo, abre o Financeiro, preenche data, valor,
     histórico e o tipo de movimento (via tabela DE/PARA)
  4. Se a natureza da despesa não estiver no DE/PARA, pausa para o usuário escolher à mão

MODO_TESTE = True preenche tudo, mas não clica em "Incluir" (para conferir antes de rodar valendo).

Colunas da planilha usadas: D = nº do processo | J = valor | K = natureza | L = histórico
"""
import glob
import os
import time
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox

import openpyxl
from dotenv import find_dotenv, load_dotenv
from selenium import webdriver
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# --- MODO TESTE ---
# True  = preenche tudo, mas NÃO clica no botão "Incluir"
# False = funcionamento normal (clica em "Incluir")
MODO_TESTE = True

# --- CONFIGURAÇÕES (vêm do arquivo .env) ---
load_dotenv(find_dotenv())
URL_SISTEMA = os.getenv("SISTEMA_URL", "https://sistema-juridico.exemplo.com.br")
NOME_USUARIO = os.getenv("SISTEMA_USUARIO", "")
SENHA_USUARIO = os.getenv("SISTEMA_SENHA", "")
DIRETORIO_PLANILHAS = os.getenv("PASTA_PLANILHAS_CUSTAS", "./planilhas")

# Seletores dos campos do sistema (ajuste conforme o sistema usado)
SELETOR = {
    "login_usuario": (By.ID, "login"),
    "login_senha": (By.ID, "password"),
    "login_entrar": (By.NAME, "entrar"),
    "busca_processo": (By.NAME, "numeroCNJ"),
    "botao_consultar": (By.XPATH, "//input[@value='Consultar']"),
    "menu_financeiro": (By.XPATH, "//*[contains(text(), 'Financeiro')]"),
    "data_movimento": (By.ID, "dataMovimento"),
    "valor_despesa": (By.ID, "valorDespesa"),
    "historico": (By.NAME, "historico"),
    "tipo_movimento": (By.ID, "tipoMovimento"),
    "botao_incluir": (By.ID, "btnIncluir"),
    "pagina_inicial": (By.XPATH, "//*[contains(text(), 'Página Inicial')]"),
}

# --- DE/PARA (NATUREZA DA PLANILHA -> OPÇÃO DO SISTEMA) ---
DE_PARA_MOVIMENTOS = {
    "CUSTAS FINAIS": "Custas - Finais",
    "CUSTAS DE IMPUGNAÇÃO AO CUMPRIMENTO DE SENTENÇA": "Custas - Atos/Citação/Intimação",
    "CUSTAS DE APELAÇÃO": "Custas - Apelação",
    "CUSTAS DE EXPEDIÇÃO DE MANDADO DE PAGAMENTO": "Custas - Mandado de Pagamento",
    "CUSTAS PARA EXPEDIÇÃO DO MANDADO DE PAGAMENTO": "Custas - Mandado de Pagamento",
    "CUSTAS PARA EXPEDIÇÃO DE MANDADO DE PAGAMENTO": "Custas - Mandado de Pagamento",
    "CUSTAS DE DESARQUIVAMENTO": "Custas - Desarquivamento",
    "CUSTAS PARA DESARQUIVAMENTO": "Custas - Desarquivamento",
    "CUSTAS DE AGRAVO INTERNO": "Custas - Agravo",
    "CUSTAS DE AGRAVO DE INSTRUMENTO": "Custas - Agravo",
    "CUSTAS DE REMESSA AO CONTADOR": "Custas - Atos/Citação/Intimação",
    "CUSTAS DE IMPUGNAÇÃO": "Custas - Atos/Citação/Intimação",
    "CUSTAS DE BAIXA": "Custas - Finais",
    "CUSTAS DE PENHORA": "Custas - Atos/Citação/Intimação",
    "CUSTAS DE CONSULTA INFOJUD E RENAJUD": "Custas - Atos/Citação/Intimação",
}

# Opções genéricas que exigem colocar a natureza da planilha no histórico
MOVIMENTOS_GENERICOS = ["Custas - Atos/Citação/Intimação"]

DIAS_SEMANA = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]


def proximo_dia_util(data):
    """Retorna o próximo dia (a partir do dia seguinte) que não seja sábado nem domingo."""
    data = data + timedelta(days=1)
    while data.weekday() >= 5:
        data = data + timedelta(days=1)
    return data


def definir_data_lancamento():
    """
    Calcula o próximo dia útil e pergunta numa janela se ele é feriado.
    Se for, pula para o próximo dia útil e pergunta de novo (cobre feriados seguidos, como Carnaval).
    Retorna a data final ou None se o usuário cancelar.
    """
    janela = tk.Tk()
    janela.withdraw()
    janela.attributes("-topmost", True)

    data = proximo_dia_util(datetime.now())

    while True:
        resposta = messagebox.askyesnocancel(
            "Verificação de Feriado",
            f"O lançamento será feito em:\n\n"
            f"    {data.strftime('%d/%m/%Y')} ({DIAS_SEMANA[data.weekday()]})\n\n"
            f"Esse dia é FERIADO?\n\n"
            f"Sim = pular para o próximo dia útil\n"
            f"Não = lançar nesta data\n"
            f"Cancelar = encerrar a automação",
            parent=janela,
        )
        if resposta is None:
            janela.destroy()
            return None
        if resposta:
            data = proximo_dia_util(data)
            continue
        break

    modo = "MODO TESTE (não vai clicar em Incluir)" if MODO_TESTE else "MODO NORMAL"
    confirmado = messagebox.askokcancel(
        "Confirmar Data",
        f"Os lançamentos serão feitos com a data:\n\n"
        f"    {data.strftime('%d/%m/%Y')} ({DIAS_SEMANA[data.weekday()]})\n\n"
        f"{modo}\n\nDeseja iniciar a automação?",
        parent=janela,
    )
    janela.destroy()
    return data if confirmado else None


def obter_planilha_mais_recente(diretorio):
    arquivos = glob.glob(os.path.join(diretorio, "*.xlsx"))
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo Excel (.xlsx) foi encontrado na pasta: {diretorio}")
    return max(arquivos, key=os.path.getmtime)


def iniciar_navegador():
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico)
    driver.maximize_window()
    return driver


def fechar_alertas(driver):
    try:
        alerta = driver.switch_to.alert
        print(f"Alerta detectado e fechado: {alerta.text}")
        alerta.accept()
        time.sleep(0.5)
    except NoAlertPresentException:
        pass


def procurar_elemento_visivel_em_frames(driver, by, valor):
    """Procura um elemento visível e habilitado na página e em todos os frames/iframes."""
    driver.switch_to.default_content()
    try:
        for el in driver.find_elements(by, valor):
            if el.is_displayed() and el.is_enabled():
                return el
    except Exception:
        pass

    iframes = driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame")
    for frame in iframes:
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(frame)
            for el in driver.find_elements(by, valor):
                if el.is_displayed() and el.is_enabled():
                    return el
        except Exception:
            continue

    driver.switch_to.default_content()
    raise Exception(f"Erro: O elemento '{valor}' não foi encontrado ou não está visível.")


def formatar_valor_br(valor):
    """Números viram '1234,56'; textos são mantidos como estão."""
    if isinstance(valor, (int, float)):
        return f"{valor:.2f}".replace(".", ",")
    return str(valor).strip()


def main():
    if not NOME_USUARIO or not SENHA_USUARIO:
        print("Preencha SISTEMA_USUARIO e SISTEMA_SENHA no arquivo .env")
        return

    # --- PASSO 0: DATA DO LANÇAMENTO (PERGUNTA SOBRE FERIADO) ---
    data_escolhida = definir_data_lancamento()
    if data_escolhida is None:
        print("Automação cancelada pelo usuário.")
        return
    data_lancamento = data_escolhida.strftime("%d/%m/%Y")
    print(f"Data de lançamento definida: {data_lancamento}")
    if MODO_TESTE:
        print(">>> MODO TESTE ATIVO: o botão 'Incluir' NÃO será clicado. <<<")

    caminho_planilha = obter_planilha_mais_recente(DIRETORIO_PLANILHAS)
    print(f"Planilha encontrada: {caminho_planilha}")

    planilha = openpyxl.load_workbook(caminho_planilha).active

    driver = iniciar_navegador()
    wait = WebDriverWait(driver, 15)
    driver.get(URL_SISTEMA)

    # --- PASSO 1: LOGIN ---
    wait.until(EC.presence_of_element_located(SELETOR["login_usuario"])).send_keys(NOME_USUARIO)
    driver.find_element(*SELETOR["login_senha"]).send_keys(SENHA_USUARIO)
    driver.find_element(*SELETOR["login_entrar"]).click()

    print("Aguardando carregamento da página principal...")
    time.sleep(5)

    linha_atual = 2

    while True:
        processo = planilha[f"D{linha_atual}"].value
        if processo is None:
            print("Fim da planilha encontrado. Encerrando automação.")
            break

        valor_despesa = planilha[f"J{linha_atual}"].value
        natureza_bruta = planilha[f"K{linha_atual}"].value
        historico_base = planilha[f"L{linha_atual}"].value

        natureza_planilha = str(natureza_bruta).strip().upper() if natureza_bruta else ""
        print(f"\n--- Processando linha {linha_atual}: Processo {processo} ---")

        # --- PASSO 2: PESQUISA DO PROCESSO ---
        fechar_alertas(driver)
        janela_principal = driver.current_window_handle

        campo_pesquisa = procurar_elemento_visivel_em_frames(driver, *SELETOR["busca_processo"])
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(str(processo))
        procurar_elemento_visivel_em_frames(driver, *SELETOR["botao_consultar"]).click()
        time.sleep(3)

        # --- PASSO 3: ENTRAR NO FINANCEIRO (abre um pop-up) ---
        procurar_elemento_visivel_em_frames(driver, *SELETOR["menu_financeiro"]).click()
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[-1])

        # --- PASSO 4: PREENCHER O NOVO MOVIMENTO ---
        wait.until(EC.presence_of_element_located(SELETOR["data_movimento"]))
        campo_data = procurar_elemento_visivel_em_frames(driver, *SELETOR["data_movimento"])
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_data)
        time.sleep(1)
        driver.execute_script(f"arguments[0].value = '{data_lancamento}';", campo_data)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", campo_data)
        time.sleep(1)
        fechar_alertas(driver)

        valor_formatado = formatar_valor_br(valor_despesa)
        campo_valor = procurar_elemento_visivel_em_frames(driver, *SELETOR["valor_despesa"])
        campo_valor.clear()
        campo_valor.send_keys(valor_formatado)

        tipo_movimento = DE_PARA_MOVIMENTOS.get(natureza_planilha)
        historico_final = str(historico_base)
        if tipo_movimento in MOVIMENTOS_GENERICOS:
            historico_final += f" - {natureza_planilha}"

        campo_historico = procurar_elemento_visivel_em_frames(driver, *SELETOR["historico"])
        campo_historico.clear()
        campo_historico.send_keys(historico_final)

        dropdown = Select(procurar_elemento_visivel_em_frames(driver, *SELETOR["tipo_movimento"]))

        if tipo_movimento:
            dropdown.select_by_visible_text(tipo_movimento)
            print(f"-> Movimento '{tipo_movimento}' selecionado automaticamente.")
        else:
            # natureza fora do DE/PARA -> intervenção manual
            print(f"\n[ATENÇÃO] A natureza '{natureza_planilha}' NÃO foi encontrada no DE/PARA.")
            print("=" * 50)
            print("PAUSA DA AUTOMAÇÃO (INTERVENÇÃO MANUAL)!")
            print("Vá até o navegador e selecione o 'Tipo de Movimento' manualmente.")
            input("Depois, pressione ENTER aqui no terminal para continuar...")
            print("=" * 50 + "\n")

        if MODO_TESTE:
            print(f"-> [MODO TESTE] 'Incluir' NÃO clicado. Data={data_lancamento} | Valor={valor_formatado}")
            time.sleep(3)  # tempo para conferir o formulário na tela
        else:
            procurar_elemento_visivel_em_frames(driver, *SELETOR["botao_incluir"]).click()
            print("-> Botão 'Incluir' clicado com sucesso.")

        time.sleep(2)

        # --- PASSO 5: FECHAR O POP-UP E VOLTAR PARA A PÁGINA INICIAL ---
        driver.close()
        driver.switch_to.window(janela_principal)
        procurar_elemento_visivel_em_frames(driver, *SELETOR["pagina_inicial"]).click()
        time.sleep(2)
        linha_atual += 1

    print("Processo finalizado com sucesso!")
    driver.quit()


if __name__ == "__main__":
    main()

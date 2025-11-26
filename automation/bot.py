import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from automation.helpers import (
    esperar_overlay_sumir,
    verificar_contratante,
    # esperar_estado_adicionar,               # ❌ não usamos mais aqui
    esperar_painel_contrato_pronto,           # ✅ alias no helpers garante compatibilidade
    esperar_flash_processando,                # ✅ segura o popup "Processando..."
    esperar_refresh_topo,                     # ✅ confirma refresh + retorno ao topo
)
from tkinter import messagebox
import time
import unicodedata
from utils.config_handler import salvar_config, carregar_config
import calendar
import datetime

# Dicionário para controle de parada da execução
controle_parada = {"parar": False}

# ====== CONTROLE DE ART ======
LIMITE_POR_ART = 100


def ler_art_id(driver):
    """Lê um identificador textual da ART exibido no topo da tela (label do cabeçalho)."""
    try:
        el = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo"]/div[1]/label'))
        )
        return (el.text or "").strip()
    except Exception:
        return ""


def esperar_troca_de_art_infinita(driver, art_atual, poll=0.75):
    """Espera INDEFINIDAMENTE até o cabeçalho da ART mudar."""
    while True:
        novo = ler_art_id(driver)
        if novo and novo != art_atual:
            return True
        time.sleep(poll)


def ler_total_contratos_ui(driver, timeout=15):
    """
    (Mantido para compatibilidade, mas não usado no fluxo principal agora.)
    Lê o total de contratos mostrado no rodapé da tabela:
    ex.: 'Mostrando de 1 até 45 de 100 registros' -> retorna 100
    """
    XPS = [
        '//*[contains(translate(normalize-space(.) ,'
        ' "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÀÂÃÉÊÍÓÔÕÚÇ",'
        ' "abcdefghijklmnopqrstuvwxyzáàâãéêíóôõúç"), "mostrando")'
        ' and contains(translate(normalize-space(.) ,'
        ' "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÀÂÃÉÊÍÓÔÕÚÇ",'
        ' "abcdefghijklmnopqrstuvwxyzáàâãéêíóôõúç"), "registros")]',
        '//*[contains(translate(normalize-space(.) ,'
        ' "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÀÂÃÉÊÍÓÔÕÚÇ",'
        ' "abcdefghijklmnopqrstuvwxyzáàâãéêíóôõúç"), "registros")]',
    ]
    end = time.time() + timeout
    while time.time() < end:
        try:
            candidatos = []
            for xp in XPS:
                candidatos.extend(driver.find_elements(By.XPATH, xp))
            textos = [(el.text or "").strip() for el in candidatos if (el.text or "").strip()]
            textos.sort(key=len, reverse=True)
            for txt in textos:
                m = re.search(r'de\s+(\d+)\s+registros', txt, flags=re.IGNORECASE) or \
                    re.search(r'de\s+(\d+)\s+registro', txt, flags=re.IGNORECASE)
                if m:
                    try:
                        return int(m.group(1))
                    except Exception:
                        pass
        except Exception:
            pass
        time.sleep(0.3)
    return None


# =============================

def limpar_documento(valor):
    """Limpa e formata um número de documento (CPF ou CNPJ)."""
    try:
        doc = str(valor).strip().replace(".", "").replace("-", "").replace("/", "")
        if doc.lower() == "nan" or not doc.isdigit():
            return "", None
        if len(doc) <= 11:
            return doc.zfill(11), "CPF"
        return doc.zfill(14), "CNPJ"
    except Exception:
        return "", None


def normalizar(txt):
    """Remove acentos, espaços e converte para maiúsculas."""
    return ''.join(
        c for c in unicodedata.normalize('NFD', str(txt))
        if unicodedata.category(c) != 'Mn'
    ).replace(' ', '').upper()


def formatar_data_simples(valor):
    """Tenta formatar um valor para o formato de data dd/mm/YYYY."""
    if pd.isna(valor) or str(valor).strip().lower() == "nan" or str(valor).strip() == "":
        return ""

    if isinstance(valor, (datetime.datetime, datetime.date)):
        return valor.strftime("%d/%m/%Y")

    # Tenta converter de formato numérico do Excel
    try:
        n = float(valor)
        if 30000 < n < 60000:
            data = (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=n)).date()
            return data.strftime("%d/%m/%Y")
    except Exception:
        pass

    valor_str = str(valor).strip().replace("\\", "/").replace("-", "/").replace(".", "/")

    # Tenta múltiplos formatos
    try:
        return pd.to_datetime(valor_str).strftime("%d/%m/%Y")
    except Exception:
        pass

    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y"):
        try:
            dt = datetime.datetime.strptime(valor_str, fmt)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass

    # Lógica para datas ambíguas (dd/mm vs mm/dd)
    if "/" in valor_str:
        partes = valor_str.split("/")
        if len(partes) == 3 and all(p.isdigit() for p in partes):
            d1, d2, ano = partes
            if int(d1) > 12:
                return f"{d1.zfill(2)}/{d2.zfill(2)}/{ano}"
            elif int(d2) > 12:
                return f"{d2.zfill(2)}/{d1.zfill(2)}/{ano}"
            else:
                return f"{d1.zfill(2)}/{d2.zfill(2)}/{ano}"

    # Última tentativa com pandas
    try:
        dt = pd.to_datetime(valor_str, dayfirst=True, errors="raise")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass

    return valor_str


def selecionar_fazenda(driver, nome_fazenda, log):
    """Seleciona a fazenda correspondente na página do CREA."""
    nome_fazenda_n = normalizar(nome_fazenda)
    while True:
        time.sleep(1)
        try:
            area = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="NDOCU"]'))
            )
            radios = area.find_elements(By.XPATH, './/input[@type="radio"]')
            fazenda_encontrada = False

            for radio in radios:
                label_text = ""
                try:
                    label = radio.find_element(By.XPATH, 'following-sibling::label')
                    label_text = label.text
                except:
                    pass
                if not label_text:
                    try:
                        label = radio.find_element(By.XPATH, 'preceding-sibling::label')
                        label_text = label.text
                    except:
                        pass

                texto_n = normalizar(label_text)
                if nome_fazenda_n in texto_n:
                    radio.click()
                    fazenda_encontrada = True
                    log(f"✅ Fazenda '{nome_fazenda}' selecionada com sucesso!")
                    break

            if fazenda_encontrada:
                return

            log(f"❌ Fazenda '{nome_fazenda}' NÃO encontrada para o CPF/CNPJ informado!")
            messagebox.showwarning(
                "Fazenda não encontrada",
                "Fazenda não existe no banco de dados.\nCadastre a fazenda no sistema CREA e aperte OK para continuar."
            )

            try:
                campo_cpf = driver.find_element(By.ID, "contratante0_CampoContratantePF")
                documento = campo_cpf.get_attribute("value")
                campo_cpf.clear()
                campo_cpf.send_keys(documento)
            except:
                try:
                    campo_cnpj = driver.find_element(By.ID, "contratante0_CampoContratantePJ")
                    documento = campo_cnpj.get_attribute("value")
                    campo_cnpj.clear()
                    campo_cnpj.send_keys(documento)
                except:
                    log("❌ Não foi possível reenviar CPF/CNPJ para recarregar as fazendas.")
                    raise Exception("Não foi possível reenviar CPF/CNPJ.")
        except StaleElementReferenceException:
            log("⚠️ StaleElementReferenceException detectada! Tentando de novo...")
            continue


def ler_unica_celula(df, col):
    """Lê um valor único de uma coluna do DataFrame."""
    try:
        val = df[col].iloc[0]
        if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan":
            return ""
        return val
    except Exception:
        return ""


def preencher_contrato_com_linha(index, driver, df, log, quantidade=1, marcador_global=None):
    """Preenche os dados de um contrato na página do CREA."""
    try:
        # 🔢 Loga a quantidade desta linha ANTES de iniciar o cadastro
        try:
            qtd_prevista = int(quantidade)
        except Exception:
            qtd_prevista = 1
        marc = marcador_global if marcador_global is not None else (index + 1)
        log(f"🔢 [{marc}] Quantidade: {qtd_prevista}")

        # ✅ Seleciona EXATAMENTE o botão correto "Cadastrar Contrato" (e NÃO o de bloco)
        botao_cadastrar = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//*[starts-with(@id, 'cadastrarContratoArt') and not(contains(@id, 'Bloco'))]"
            ))
        )

        # Scroll + clique seguro
        driver.execute_script(
            "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
            botao_cadastrar
        )
        time.sleep(1.5)
        driver.execute_script("arguments[0].click();", botao_cadastrar)
        log(f"✅ [{marc}] Iniciando cadastro do contrato")

        # ⏳ Espera o popup "Processando..." sumir (tolerante a pisca) + overlay
        esperar_flash_processando(driver, max_espera=8, quiet_ms=450, poll=0.06)
        esperar_overlay_sumir(driver)

        # ✅ Espera adaptativa rápida para os campos-chave ficarem prontos
        try:
            esperar_painel_contrato_pronto(driver, timeout=3)
        except Exception:
            time.sleep(0.5)  # fallback curtinho

        # Preenchimento dos campos da ART
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "NIVEL00"))).click()
        for opt in driver.find_elements(By.XPATH, '//*[@id="NIVEL00"]/option'):
            if "EXECUÇÃO" in opt.text.upper():
                opt.click()
                break

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ATIVIDADEPROFISSIONAL00"))).click()
        for opt in driver.find_elements(By.XPATH, '//*[@id="ATIVIDADEPROFISSIONAL00"]/option'):
            if "EXECUÇÃO DE SERVIÇO TÉCNICO" in opt.text.upper():
                opt.click()
                break

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "LABELATUACAO00"))
        ).send_keys("39.24")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="listaAtividadeEscolherATUACAO00"]/ul/li[1]'))
        ).click()

        unidade = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "UNIDADEMEDIDA00")))
        unidade.click()
        for opt in driver.find_elements(By.XPATH, '//*[@id="UNIDADEMEDIDA00"]/option'):
            if "UNIDADE" in opt.text.upper():
                opt.click()
                break

        campo_qtd = driver.find_element(By.ID, "QUANTIDADE00")
        campo_qtd.clear()
        campo_qtd.send_keys(f"{int(quantidade)},00")

        # Preenchimento dos dados do contratante
        documento_bruto = str(df.loc[index, "CPF_CNPJ"]).strip().replace(".", "").replace("-", "").replace("/", "")
        if len(documento_bruto) <= 11:
            documento = documento_bruto.zfill(11)
            driver.find_element(By.ID, "contratante0_ContratantePF").click()
            campo_id = "contratante0_CampoContratantePF"
        else:
            documento = documento_bruto.zfill(14)
            driver.find_element(By.ID, "contratante0_ContratantePJ").click()
            campo_id = "contratante0_CampoContratantePJ"

        campo_doc = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, campo_id)))
        campo_doc.clear()
        campo_doc.send_keys(documento)
        time.sleep(1)

        # Pausa para cadastro de contratante não existente
        try:
            time.sleep(1)
            ndocu_exists = False
            try:
                driver.find_element(By.XPATH, '//*[@id="NDOCU"]')
                ndocu_exists = True
            except:
                ndocu_exists = False

            result_pesquisa = driver.find_elements(By.XPATH, '//*[@id="contratante0_ResultPesquisa"]/a')
            if (not ndocu_exists) and result_pesquisa:
                log("⏸️ Contratante não cadastrado! Aguarde cadastro manual.")
                messagebox.showwarning(
                    "Contratante não cadastrado",
                    "O contratante (CPF/CNPJ) não está cadastrado no CREA.\n"
                    "Cadastre o contratante manualmente e clique OK para continuar."
                )
                while True:
                    time.sleep(2)
                    try:
                        driver.find_element(By.XPATH, '//*[@id="NDOCU"]')
                        log("✅ Contratante cadastrado! Continuando o processo.")
                        break
                    except:
                        pass
        except Exception as e:
            log(f"⚠️ Erro ao checar contratante cadastrado: {e}")

        verificar_contratante(driver, campo_doc, documento, index)
        selecionar_fazenda(driver, df.loc[index, "FAZENDA"], log)

        # Preenchimento dos dados do contrato
        driver.find_element(By.XPATH, '//input[contains(@id, "CONTRATO_NUMERO")]').send_keys(
            str(df.loc[index, "NUMERO DO CONTRATO"])
        )

        raw_dr = df.loc[index, "DATA DO REGISTRO"] if "DATA DO REGISTRO" in df.columns else ""
        dr = formatar_data_simples(raw_dr)
        if dr:
            campo_data_reg = driver.find_element(By.XPATH, '//*[@id="CONTRATO_DATA0"]')
            campo_data_reg.clear()
            campo_data_reg.send_keys(dr)
            log(f"DATA REGISTRO: {dr}")

        raw_di = ler_unica_celula(df, "DATA_INICIO")
        di = formatar_data_simples(raw_di) if raw_di and str(raw_di).strip() and str(raw_di).strip().lower() != "nan" else \
            f"01/{datetime.datetime.now().month:02d}/{datetime.datetime.now().year}"
        campo_data_inicio = driver.find_element(By.XPATH, '//*[@id="CONTRATO_DATAINICIO0"]')
        campo_data_inicio.clear()
        campo_data_inicio.send_keys(di)
        log(f"DATA INICIO: {di}")

        raw_dfim = ler_unica_celula(df, "DATA_FIM")
        dfim = formatar_data_simples(raw_dfim) if raw_dfim and str(raw_dfim).strip() and str(raw_dfim).strip().lower() != "nan" else \
            f"{calendar.monthrange(datetime.datetime.now().year, datetime.datetime.now().month)[1]:02d}/{datetime.datetime.now().month:02d}/{datetime.datetime.now().year}"
        campo_data_fim = driver.find_element(By.XPATH, '//*[@id="CONTRATO_DATAFIM0"]')
        campo_data_fim.clear()
        campo_data_fim.send_keys(dfim)
        log(f"DATA FIM: {dfim}")

        valor_receita = ler_unica_celula(df, "VALOR_RECEITA")
        if not valor_receita or str(valor_receita).lower() == "nan":
            valor_receita = "2,00"
        campo_valor = driver.find_element(By.XPATH, '//*[@id="CONTRATO_VALOR0"]')
        campo_valor.clear()
        campo_valor.send_keys(valor_receita)
        log(f"VALOR: {valor_receita}")

        # Seleção do endereço do cliente
        try:
            endereco_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="evtContratoEnderecoContainerSpecific0"]/div[3]/input[1]'))
            )
            endereco_input.click()
            log("✅ Input do endereço do cliente selecionado com sucesso!")
            time.sleep(1.5)
        except Exception as e:
            log(f"⚠️ Não foi possível selecionar o input de endereço do cliente: {e}")

        # Checagem de ART cheia (mantida para fallback)
        try:
            botao_salvar = None
            try:
                botao_salvar = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "save"))
                )
            except Exception:
                botao_salvar = None

            if not botao_salvar:
                log(f"⚠️ ART cheia ou sistema bloqueou cadastro! Aguarde troca de ART manual.")
                resposta = messagebox.askquestion(
                    "Cadastro interrompido - ART cheia",
                    "O limite de contratos dessa ART foi atingido.\n\n"
                    "Cadastre uma nova ART manualmente e clique em 'Sim' para continuar, ou 'Não' para encerrar."
                )
                if resposta == "no":
                    log("❌ Execução encerrada pelo usuário (ART cheia, sem nova ART disponível).")
                    return False
                else:
                    log("⏳ Aguardando indefinidamente até o botão 'Cadastrar Contrato' reaparecer com nova ART...")
                    while True:
                        botoes = driver.find_elements(By.XPATH, '//*[starts-with(@id, "cadastrarContratoArt")]')
                        if botoes:
                            log("✅ Nova ART detectada. Continuando execução...")
                            break
                        time.sleep(2)
                    return False
        except Exception as e:
            log(f"⚠️ Erro ao checar ART cheia: {e}")

        # Salvar (Adicionar)
        botao_salvar = WebDriverWait(driver, 25).until(EC.element_to_be_clickable((By.ID, "save")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", botao_salvar)
        driver.execute_script("arguments[0].click();", botao_salvar)

        # ⏳ Regra: espera o "Processando..." sumir e a página voltar pro topo (refresh)
        esperar_flash_processando(driver, max_espera=8, quiet_ms=450, poll=0.06)
        esperar_refresh_topo(driver, timeout=10, poll=0.06)

        log(f"✅ [{marc}] Contrato salvo com sucesso")
        return True

    except Exception as e:
        log(f"❌ [{(marcador_global or index+1)}] Erro ao preencher contrato: {e}")
        return False


def comparar_datas_sem_hora(dt1, dt2):
    """Compara duas datas ignorando a parte de hora."""
    try:
        d1 = pd.to_datetime(dt1, dayfirst=True, errors="coerce").date()
    except Exception:
        d1 = None
    try:
        d2 = pd.to_datetime(dt2, dayfirst=True, errors="coerce").date()
    except Exception:
        d2 = None
    return (d1 is not None) and (d2 is not None) and (d1 == d2)


def executar_lote(
    df,
    inicio,
    log,
    numero_inicial=None,
    cpf_cnpj_inicial=None,
    data_registro_inicial=None,
    callback_atualizar_contrato=None
):
    """Função principal que executa o processo de automação em lote."""
    if not isinstance(df, pd.DataFrame):
        df = pd.read_excel(df, dtype=str)

    # Agrupamento de contratos para contar quantidades
    colunas_chave = ["NUMERO DO CONTRATO", "CPF_CNPJ", "DATA DO REGISTRO"]
    df_agrupado = (
        df.groupby(colunas_chave, as_index=False)
        .size()
        .rename(columns={'size': 'QUANTIDADE'})
        .merge(df.drop_duplicates(subset=colunas_chave), on=colunas_chave, how='left')
    )
    df = df_agrupado.reset_index(drop=True)

    # Encontrar a posição inicial
    if numero_inicial and cpf_cnpj_inicial and data_registro_inicial:
        pos = None
        for idx, row in df.iterrows():
            if (
                str(row["NUMERO DO CONTRATO"]) == str(numero_inicial) and
                str(row["CPF_CNPJ"]) == str(cpf_cnpj_inicial) and
                comparar_datas_sem_hora(row["DATA DO REGISTRO"], data_registro_inicial)
            ):
                pos = idx
                break
        if pos is None:
            log(f"Contrato {numero_inicial}/{cpf_cnpj_inicial}/{data_registro_inicial} não encontrado após agrupamento.")
            return
    else:
        pos = 0

    # Configuração do WebDriver
    cpf = str(df["CPF_LOGIN"].dropna().iloc[0]).strip()
    senha = str(df["SENHA_LOGIN"].dropna().iloc[0]).strip()
    art = str(df["ARTCREA"].dropna().iloc[0]).strip()

    chrome_options = Options()
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")

    log("🌐 Iniciando navegador...")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get("https://servicos-crea-mg.sitac.com.br/index.php")

    # Verificação de parada antes do login
    if controle_parada["parar"]:
        if messagebox.askquestion("Confirmação de parada", "Deseja realmente encerrar a execução?") == "yes":
            log("🛑 Parada solicitada antes do login.")
            return
        else:
            controle_parada["parar"] = False
            log("▶️ Usuário optou por continuar.")

    # Tentativas de login
    tentativas = 0
    max_tentativas = 3
    login_ok = False
    while tentativas < max_tentativas and not login_ok:
        if controle_parada["parar"]:
            if messagebox.askquestion("Confirmação de parada", "Deseja realmente encerrar a execução?") == "yes":
                log("🛑 Parada solicitada durante tentativa de login.")
                return
            else:
                controle_parada["parar"] = False
                log("▶️ Usuário optou por continuar.")

        try:
            log(f"⏳ Tentando login... (tentativa {tentativas + 1})")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "login"))).send_keys(cpf)
            driver.find_element(By.ID, "senha").send_keys(senha)
            driver.find_element(By.XPATH, '//button[contains(text(),"Entrar")]').click()
            log("✅ Login realizado com sucesso!")
            login_ok = True
        except Exception as e:
            tentativas += 1
            log(f"⚠️ Erro no login (tentativa {tentativas}): {e}")
            if tentativas < max_tentativas:
                log("🔄 Recarregando a página...")
                driver.refresh()
                time.sleep(5)
            else:
                log("❌ Falha no login após várias tentativas.")
                driver.quit()
                return

    if controle_parada["parar"]:
        if messagebox.askquestion("Confirmação de parada", "Deseja realmente encerrar a execução?") == "yes":
            log("🛑 Parada solicitada após login.")
            return
        else:
            controle_parada["parar"] = False
            log("▶️ Usuário optou por continuar.")

    # Navegação até a ART
    try:
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sidebar"]/ul/li[5]/a'))).click()
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="PesquisarART"]/a'))).click()
        log("✅ Menu 'Pesquisar ART' acessado")

        if controle_parada["parar"]:
            if messagebox.askquestion("Confirmação de parada", "Deseja realmente encerrar a execução?") == "yes":
                log("🛑 Parada solicitada antes de pesquisar ART.")
                return
            else:
                controle_parada["parar"] = False
                log("▶️ Usuário optou por continuar.")

        campo_art = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "NUMEROART")))
        campo_art.clear()
        campo_art.send_keys(art)
        time.sleep(1)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Result_NumeroART"]/a'))).click()
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo"]/div[1]/label')))
        log("✅ ART localizada com sucesso")
    except Exception as e:
        log(f"❌ Erro ao localizar ART: {e}")
        driver.quit()
        return

    if controle_parada["parar"]:
        if messagebox.askquestion("Confirmação de parada", "Deseja realmente encerrar a execução?") == "yes":
            log("🛑 Parada solicitada antes de processar contratos.")
            return
        else:
            controle_parada["parar"] = False
            log("▶️ Usuário optou por continuar.")

    # Identificação da ART atual
    art_id_atual = ler_art_id(driver)

    # Loop principal de processamento dos contratos
    contratos_lancados = 0
    for i, (_, linha) in enumerate(df.iloc[pos:].iterrows()):
        idx = i + pos
        n_global = idx + 1  # referência única (1,2,3,...)

        # Portão: sempre que for começar 101/201/301/... PAUSA antes de lançar
        if n_global > 1 and ((n_global - 1) % LIMITE_POR_ART == 0):
            # Terminou 100/200/300 no item anterior; exigir nova ART agora.
            log(f"⛔ Ciclo fechado no [{n_global - 1}]. Cadastre uma NOVA ART para continuar com [{n_global}].")
            continuar = messagebox.askyesno(
                "Nova ART necessária",
                f"Você finalizou {LIMITE_POR_ART} lançamentos (até [{n_global - 1}]).\n\n"
                f"1) Cadastre agora uma NOVA ART no sistema.\n"
                f"2) Depois de cadastrar, clique em 'Sim' para continuar.\n\n"
                "Se quiser encerrar, clique em 'Não'."
            )
            if not continuar:
                log("🛑 Execução encerrada pelo usuário ao fechar o ciclo da ART.")
                driver.quit()
                return
            log("⏳ Aguardando troca de ART...")
            esperar_troca_de_art_infinita(driver, art_id_atual, poll=0.75)
            art_id_atual = ler_art_id(driver)
            log(f"✅ Nova ART detectada: '{art_id_atual}'. Continuando do [{n_global}]...")

        if controle_parada["parar"]:
            if messagebox.askquestion("Confirmação de parada", "Deseja realmente encerrar a execução?") == "yes":
                log("🛑 Parada solicitada durante o processamento.")
                break
            else:
                controle_parada["parar"] = False
                log("▶️ Usuário optou por continuar.")

        contrato = str(linha["NUMERO DO CONTRATO"]).strip()
        log(f"[{n_global}] Preparando lançamento nesta ART")
        log(f"▶️ Processando contrato {contrato}...")

        try:
            sucesso = preencher_contrato_com_linha(
                idx, driver, df, log,
                quantidade=linha["QUANTIDADE"],
                marcador_global=n_global
            )
            if not sucesso:
                log(f"⚠️ Contrato {contrato} pausado (possivelmente ART cheia).")
                continue
            else:
                log(f"✅ Contrato {contrato} lançado com sucesso.")

            try:
                config = carregar_config()
                config["ultimo_contrato"] = contrato
                salvar_config(config)
                if callback_atualizar_contrato:
                    callback_atualizar_contrato(contrato)
            except Exception as e:
                log(f"⚠️ Erro ao salvar progresso: {e}")

            contratos_lancados += 1

            # Recarrega a página a cada 10 contratos para estabilidade
            if contratos_lancados % 10 == 0:
                log("♻️ Recarregando página para evitar erros acumulados...")
                driver.refresh()
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo"]/div[1]/label'))
                )
                art_id_atual = ler_art_id(driver)

            if controle_parada["parar"]:
                if messagebox.askyesno("Confirmação de parada", "Deseja realmente encerrar a execução?"):
                    log("🛑 Parada solicitada após contrato.")
                    break
                else:
                    controle_parada["parar"] = False
                    log("▶️ Usuário optou por continuar.")

        except Exception as e:
            log(f"❌ Erro no contrato {contrato}: {e}")
            continue

        progresso = (i + 1) / len(df.iloc[pos:])
        log(f"Progresso: {progresso:.0%}")


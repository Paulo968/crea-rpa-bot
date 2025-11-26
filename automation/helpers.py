from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from tkinter import messagebox
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# =======================
# Esperas gerais
# =======================

def esperar_overlay_sumir(driver, timeout=30):
    """Espera o overlay/ajax sumir (ID ajax-overlay)."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "ajax-overlay"))
        )
    except Exception:
        print("⚠️ Timeout esperando overlay desaparecer.")

def verificar_contratante(driver, campo_doc, documento, index):
    """Confere se apareceu a área do contratante (NDOCU); se não, pede cadastro manual."""
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="NDOCU"]'))
        )
        print(f"✅ [{index}] Contratante {documento} encontrado.")
    except Exception:
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="contratante0_ResultPesquisa"]/a'))
            )
            print(f"⏳ [{index}] Contratante {documento} não encontrado. Cadastro manual necessário.")
            messagebox.showinfo(
                title="Contratante não encontrado",
                message=(
                    f"O contratante '{documento}' não existe no banco de dados.\n\n"
                    "Cadastre o contratante no sistema CREA e aperte OK para continuar."
                )
            )
            campo_doc.clear()
            campo_doc.send_keys(documento)
            time.sleep(3)
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="NDOCU"]'))
            )
        except Exception:
            print(f"⚠️ [{index}] Erro ao verificar contratante {documento}. Verificação manual pode ser necessária.")

# =======================
# “Adicionar/Salvar” (Aguardando Processamento)
# =======================

WAIT_PADRAO = 25
WAIT_LENTO  = 45

def _get_btn_save(driver):
    """Rebusca o botão de salvar/adicionar (SPA pode re-renderizar)."""
    try:
        return WebDriverWait(driver, WAIT_PADRAO).until(
            EC.presence_of_element_located((By.ID, "save"))
        )
    except TimeoutException:
        return WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//button[@id="save" or @name="save"]'))
        )

def _text_of(driver, el):
    """Texto visível do botão, robusto contra re-render."""
    try:
        return (driver.execute_script(
            "return (arguments[0].innerText || arguments[0].textContent || '').trim();", el
        ) or "").upper()
    except StaleElementReferenceException:
        el2 = _get_btn_save(driver)
        return (driver.execute_script(
            "return (arguments[0].innerText || arguments[0].textContent || '').trim();", el2
        ) or "").upper()

def _enabled(driver, el):
    """Verifica se o botão está habilitado, considerando attrs/classes."""
    try:
        disabled_attr = driver.execute_script("return arguments[0].disabled === true;", el)
        cls = (driver.execute_script("return (arguments[0].className || '').toString();", el) or "").lower()
        aria_busy = (driver.execute_script("return arguments[0].getAttribute('aria-busy');", el) or "").lower()
        return (not disabled_attr) and ('disabled' not in cls) and (aria_busy not in ('true', '1'))
    except StaleElementReferenceException:
        el2 = _get_btn_save(driver)
        return _enabled(driver, el2)

# ---------- NOVO: espera compacta pro popup "Processando..." (com pisca) ----------
def esperar_flash_processando(driver, max_espera=10, quiet_ms=400, poll=0.05):
    """
    Espera o modal/overlay 'Processando...' terminar, mesmo que pisque várias vezes.
    Libera só após ficar quieto por quiet_ms contínuos; se estourar max_espera, libera.
    """
    deadline = time.time() + float(max_espera)
    ultimo_visto = None
    ja_apareceu = False

    XPS = [
        # Texto/legenda do popup
        '//*[contains(translate(normalize-space(.),'
        ' "ABCDEFGHIJKLMNOPQRSTUVWXYZÇÁÀÂÃÉÊÍÓÔÕÚ",'
        ' "abcdefghijklmnopqrstuvwxyzçáàâãéêíóôõú"), "processando")]',
        '//*[contains(translate(normalize-space(.),'
        ' "ABCDEFGHIJKLMNOPQRSTUVWXYZÇÁÀÂÃÉÊÍÓÔÕÚ",'
        ' "abcdefghijklmnopqrstuvwxyzçáàâãéêíóôõú"), "aguarde")]',
        # Containers típicos
        '//*[contains(translate(@class,"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"modal")]',
        '//*[contains(translate(@class,"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"overlay")]',
        '//*[contains(translate(@class,"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"blockui")]',
    ]

    def _visivel():
        try:
            for xp in XPS:
                for el in driver.find_elements(By.XPATH, xp):
                    try:
                        if el.is_displayed():
                            txt = (el.text or "").lower()
                            cls = (el.get_attribute("class") or "").lower()
                            if ("processando" in txt) or ("aguarde" in txt) or ("por favor" in txt):
                                return True
                            if any(k in cls for k in ("modal", "overlay", "blockui")):
                                return True
                    except Exception:
                        continue
            return False
        except Exception:
            return False

    while time.time() < deadline:
        v = _visivel()
        now = time.time()

        if v:
            ja_apareceu = True
            ultimo_visto = now
        else:
            if not ja_apareceu:
                return True  # nunca detectou; não segura fluxo
            if ultimo_visto is not None and (now - ultimo_visto) >= (float(quiet_ms) / 1000.0):
                return True

        time.sleep(float(poll))

    return True  # segurança: não travar

# ---------- NOVO: detecta "refresh + voltou ao topo" ----------
def esperar_refresh_topo(driver, timeout=12, poll=0.06):
    """
    Libera quando a página 'recarrega' e a rolagem está no topo.
    Critérios:
      - window.pageYOffset ~ 0
      - e reaparecer a lista/botões de 'cadastrarContratoArt'
    """
    deadline = time.time() + float(timeout)

    def _no_topo():
        try:
            y = driver.execute_script(
                "return (window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0);"
            )
            return (y is None) or (float(y) <= 4.0)
        except Exception:
            return False

    while time.time() < deadline:
        try:
            botoes = driver.find_elements(By.XPATH, '//*[starts-with(@id, "cadastrarContratoArt")]')
            if botoes and _no_topo():
                return True
        except Exception:
            pass
        time.sleep(float(poll))

    # Não travar o fluxo se não detectar exatamente—segue
    return True

def esperar_estado_adicionar(driver, log=None, timeout_aguardar=WAIT_PADRAO, timeout_processar=WAIT_LENTO):
    """
    Espera o ciclo do botão #save sem sleep fixo (fallback antigo/compatível).
    Preferir hoje o par:
      - esperar_flash_processando (pós clique)
      - esperar_refresh_topo (refresh + topo) 
    """
    if log is None:
        log = lambda *a, **k: None

    # -------- Fase A: confirmar que entrou em processamento --------
    t_desabilitar = min(5, timeout_aguardar)  # curto e responsivo
    deadline = time.time() + float(t_desabilitar)
    entrou = False
    while time.time() < deadline:
        btn = None
        try:
            btn = _get_btn_save(driver)
            txt = _text_of(driver, btn)
            if "AGUARDANDO PROCESSAMENTO" in txt or not _enabled(driver, btn):
                entrou = True
                break
        except Exception:
            pass

        # sinais alternativos (já mudou de tela/estado)
        if driver.find_elements(By.XPATH, '//*[starts-with(@id, "cadastrarContratoArt")]'):
            return True
        if driver.find_elements(
            By.XPATH,
            '//*[contains(translate(normalize-space(.), "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÀÂÃÉÊÍÓÔÕÚÇ", "abcdefghijklmnopqrstuvwxyzáàâãéêíóôõúç"), "sucesso")]'
        ):
            return True

        # se aparecer o popup, espera ele sumir (curto)
        esperar_flash_processando(driver, max_espera=3, quiet_ms=300, poll=0.05)
        time.sleep(0.15)

    if not entrou:
        log("ℹ️ Não vi desabilitar/texto; seguirei pelo estado geral da página.")

    # -------- Fase B: aguardar a conclusão do processamento --------
    deadline2 = time.time() + float(timeout_processar)
    while time.time() < deadline2:
        # sinais de conclusão
        if driver.find_elements(By.XPATH, '//*[starts-with(@id, "cadastrarContratoArt")]'):
            return True
        if driver.find_elements(
            By.XPATH,
            '//*[contains(translate(normalize-space(.), "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÀÂÃÉÊÍÓÔÕÚÇ", "abcdefghijklmnopqrstuvwxyzáàâãéêíóôõúç"), "sucesso")]'
        ):
            return True

        # reforço via botão
        try:
            btn = _get_btn_save(driver)
            txt = _text_of(driver, btn)
            if "AGUARDANDO PROCESSAMENTO" not in txt and _enabled(driver, btn):
                # antes de liberar, garante que o popup não está piscando
                esperar_flash_processando(driver, max_espera=3, quiet_ms=350, poll=0.06)
                return True
        except Exception:
            pass

        # se apareceu o popup novamente, respeita o ciclo
        esperar_flash_processando(driver, max_espera=3, quiet_ms=350, poll=0.06)
        time.sleep(0.15)

    raise TimeoutException("Tempo excedido aguardando o ciclo do botão Adicionar/Salvar.")

# =======================
# Processando: esperar sumir DEFINITIVAMENTE (com tolerância a pisca)
# =======================

# XPaths que tentam capturar tanto texto quanto ícones/imagens comuns de "processando"
XPATHS_PROCESSANDO = [
    # Qualquer elemento onde o texto contenha "processando" (case-insensitive)
    '//*[contains(translate(normalize-space(.), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "processando")]',
    # Ícones/spinners por atributos comuns
    '//img[contains(translate(@alt, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "processando") or '
    '     contains(translate(@title, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "processando") or '
    '     contains(translate(@aria-label, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "processando") or '
    '     contains(translate(@src, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "processando") or '
    '     contains(translate(@class, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "spinner") or '
    '     contains(translate(@src, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "spinner") or '
    '     contains(translate(@class, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "loading") or '
    '     contains(translate(@src, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "loading")]',
    # Qualquer elemento com classes típicas de loading/spinner
    '//*[contains(translate(@class, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "spinner") or '
    '   contains(translate(@class, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "loading") or '
    '   contains(translate(@class, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "progress")]',
]

def _processando_visivel(driver):
    """Retorna True se algum seletor de 'processando' estiver visível na tela."""
    try:
        for xp in XPATHS_PROCESSANDO:
            els = driver.find_elements(By.XPATH, xp)
            for el in els:
                try:
                    if el.is_displayed():
                        return True
                except Exception:
                    continue
        return False
    except Exception:
        return False

def esperar_processando_sumir_definitivo(driver, timeout=30, quiet_ms=600, poll=0.05):
    """
    Espera enquanto a imagem/mensagem 'processando' estiver visível.
    Só libera quando ficar AUSENTE de forma CONTÍNUA por quiet_ms (tolerância a “piscadas”/reaparecimentos).
    - timeout: tempo máximo total
    - quiet_ms: janelinha de silêncio exigida (padrão 600ms)
    - poll: intervalo entre checagens
    """
    deadline = time.time() + float(timeout)
    ultimo_visto = None

    while time.time() < deadline:
        visivel = _processando_visivel(driver)
        now = time.time()

        if visivel:
            ultimo_visto = now
        else:
            if ultimo_visto is None:
                return True
            if (now - ultimo_visto) >= (float(quiet_ms) / 1000.0):
                return True

        time.sleep(float(poll))

    return False

# =======================
# Retrocompatibilidade: alias para o nome antigo
# =======================

def esperar_painel_contrato_pronto(driver, timeout=30):
    """
    Alias para manter compatibilidade com versões antigas do bot.
    Mapeia para esperar_estado_adicionar com timeouts razoáveis.
    """
    timeout_aguardar = min(max(int(timeout), 10), WAIT_LENTO)
    timeout_processar = max(WAIT_PADRAO, int(timeout) + 15)
    return esperar_estado_adicionar(
        driver,
        log=None,
        timeout_aguardar=timeout_aguardar,
        timeout_processar=timeout_processar
    )

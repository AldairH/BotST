import re
import json
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

URL = "https://seguimientotitulacion.unam.mx/control/login"
SEL_SEGUIMIENTO = (By.XPATH, "//a[normalize-space()='Seguimiento']")
SEL_FILAS   = (By.CSS_SELECTOR, "table.tab-docs tbody tr")
SEL_DETALLE = (By.CSS_SELECTOR, "#expediente, .detalle-expediente")
SEL_FILAS_TABLA  = (By.CSS_SELECTOR, "table.tab-docs tbody tr")
SEL_COL_ESTADO   = (By.CSS_SELECTOR, "table.tab-docs tbody tr td:nth-child(6)")
SEL_TBODY        = (By.CSS_SELECTOR, "table.tab-docs tbody")

# -----------------------------------------------------------------------------
# LOGIN
# -----------------------------------------------------------------------------

def esperar_login_e_ir_a_seguimiento(driver):
    driver.get(URL)
    print("->Inicia sesión manualmente")
    WebDriverWait(driver, 600).until(EC.presence_of_element_located(SEL_SEGUIMIENTO))
    print("->Login detectado, dando click en Seguimiento")
    driver.find_element(*SEL_SEGUIMIENTO).click()
    WebDriverWait(driver, 20).until(EC.url_contains("/listado/seguimiento"))
    print("->Estamos en la sección de Seguimiento")


def norm(s: str) -> str:
    return " ".join((s or "").split())

# -----------------------------------------------------------------------------
# FILTROS
# -----------------------------------------------------------------------------

def cambiar_mostrar_100(driver, timeout=8):
    wait = WebDriverWait(driver, timeout)

    select_el = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "select[wire\\:model='cantidad']"))
    )
    sel = Select(select_el)

    try:
        actual = sel.first_selected_option.get_attribute("value") or sel.first_selected_option.text
        if "100" in norm(actual):
            wait.until(lambda d: len(d.find_elements(*SEL_FILAS_TABLA)) > 10)
            print("-> El tamaño de la tabla ya estaba en 100")
            return
    except Exception:
        pass

    sel.select_by_value("100")
    driver.execute_script("""
        const el = arguments[0];
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.blur && el.blur();
    """, select_el)

    wait.until(lambda d: len(d.find_elements(*SEL_FILAS_TABLA)) > 10)
    print("-> El tamaño de la página fue ajustado a 100 registros")

def seleccionar_filtro_por_estado(driver, valor="Entrega electrónica y física de documentos", timeout=20):
    wait = WebDriverWait(driver, timeout)
    valor_norm = norm(valor)

    combo = Select(wait.until(EC.element_to_be_clickable((By.ID, "est_avance"))))
    try:
        combo.select_by_value(valor)
    except Exception:
        combo.select_by_visible_text(valor)

    def ok(d):
        filas = d.find_elements(*SEL_FILAS_TABLA)
        if not filas:
            return True
        for c in d.find_elements(*SEL_COL_ESTADO):
            if norm(c.text) != valor_norm:
                return False
        return True

    wait.until(ok)
    print("-> Filtro aplicado")

    cambiar_mostrar_100(driver, timeout=8)

# -----------------------------------------------------------------------------
# OBTENER URL DIRECTA
# -----------------------------------------------------------------------------

def _obtener_url_expediente_desde_fila(driver, fila):
    candidatos = fila.find_elements(By.CSS_SELECTOR, "td:last-child button.btn-accion")
    btn = None
    for b in candidatos:
        if b.find_elements(By.CSS_SELECTOR, "i.fa-file-alt"):
            btn = b
            break
    if btn is None:
        raise RuntimeError("No se encontro el botón de expediente en la fila seleccionada")

    attrs = driver.execute_script("""
        var el = arguments[0], out = {};
        for (const a of el.attributes) out[a.name]=a.value;
        return out;
    """, btn)
    onclick   = attrs.get("onclick", "") or ""
    data_href = attrs.get("data-href", "") or ""
    at_click  = attrs.get("@click", "") or attrs.get("x-on:click", "") or attrs.get("hx-get","") or ""

    url = None
    for texto in (onclick, data_href, at_click):
        m = (re.search(r"(https?://[^\s'\"<>]+/expediente[^\s'\"<>]*)", texto)
             or re.search(r"location\\.href\\s*=\\s*'([^']+)'", texto)
             or re.search(r'location\\.href\\s*=\\s*"([^"]+)"', texto))
        if m:
            url = m.group(1) if m.groups() else m.group(0)
            break

    if not url:
        raise RuntimeError("No se puede derivar la URL del expediente desde los atributos del boton")
    return url

# -----------------------------------------------------------------------------
# EXTRAER VALORES
# -----------------------------------------------------------------------------

def obtener_valor(driver, label_text, timeout=20):
    wait = WebDriverWait(driver, timeout)
    xp = f'//div[normalize-space()="{label_text}"]/following-sibling::div[1]'
    el = wait.until(EC.visibility_of_element_located((By.XPATH, xp)))
    return el.text.strip()


def extraer_expediente(driver):
    return {
        "numero_cuenta":     obtener_valor(driver, "Número de cuenta:"),
        "nombre":            obtener_valor(driver, "Nombre:"),
        "opcion_titulacion": obtener_valor(driver, "Opción de titulación:"),
        "correo":            obtener_valor(driver, "Correo electrónico:"),
        "plantel":           obtener_valor(driver, "Plantel:"),
        "carrera":           obtener_valor(driver, "Carrera:"),
        "plan_estudios":     obtener_valor(driver, "Plan de estudios:"),
    }

# -----------------------------------------------------------------------------
# EXTRAER CITA
# -----------------------------------------------------------------------------

def obtener_cita_programada_instant(driver):
    container_xpath = ("//div[contains(@class,'bg-emerald-50') "
                       "and .//text()[contains(.,'Cita programada')]]")
    conts = driver.find_elements(By.XPATH, container_xpath)
    if not conts:
        return None

    cont = conts[0]
    texto_cita = norm(cont.text)
    return {"cita_fecha": texto_cita}

# -----------------------------------------------------------------------------
# ABRIR EXPEDIENTE EN NUEVA PESTAÑA
# -----------------------------------------------------------------------------

def abrir_y_extraer_en_pestana_nueva(driver, url, timeout=20):
    original = driver.current_window_handle

    driver.switch_to.new_window('tab')
    driver.get(url)

    WebDriverWait(driver, timeout).until(
        EC.any_of(EC.url_contains("/expediente"), EC.presence_of_element_located(SEL_DETALLE))
    )

    cita = obtener_cita_programada_instant(driver)
    datos = None
    if cita:
        base = extraer_expediente(driver)
        base.update(cita)
        datos = base
    else:
        print("   -> Expediente sin cita programada, omitido")

    driver.close()
    driver.switch_to.window(original)
    return datos

# -----------------------------------------------------------------------------
# PAGINACION Y RECORRER EXPEDIENTES
# -----------------------------------------------------------------------------

def _ir_a_siguiente_pagina(driver, timeout=12):
    wait = WebDriverWait(driver, timeout)

    candidatos = driver.find_elements(
        By.CSS_SELECTOR,
        "button[rel='next'], button[wire\\:click^='nextPage']"
    )
    btn = next((b for b in candidatos if b.is_displayed() and b.is_enabled()), None)
    if not btn:
        return False

    try:
        tbody = wait.until(EC.presence_of_element_located(SEL_TBODY))
        filas = tbody.find_elements(*SEL_FILAS_TABLA)
        fila_ref = filas[0] if filas else None
        first_text = filas[0].text if filas else ""
    except Exception:
        fila_ref, first_text = None, ""

    try:
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click()", btn)

    if fila_ref:
        try:
            wait.until(EC.staleness_of(fila_ref))
        except TimeoutException:
            wait.until(lambda d: d.find_elements(*SEL_FILAS_TABLA)
                               and d.find_elements(*SEL_FILAS_TABLA)[0].text != first_text)
    else:
        wait.until(lambda d: len(d.find_elements(*SEL_FILAS_TABLA)) > 0)

    return True

def recorrer_expedientes(driver):
    resultados = []
    omitidos = 0
    vistos = set()
    pagina = 1
    while True:
        filas = driver.find_elements(*SEL_FILAS)
        total = len(filas)
        print(f"\n== Página {pagina}: {total} filas ==")

        for i in range(total):
            filas = driver.find_elements(*SEL_FILAS)
            fila = filas[i]

            url = _obtener_url_expediente_desde_fila(driver, fila)
            if url in vistos:
                continue

            datos = abrir_y_extraer_en_pestana_nueva(driver, url)
            if datos:
                resultados.append(datos)
                vistos.add(url)
            else:
                omitidos += 1

        avanzo = _ir_a_siguiente_pagina(driver)
        if not avanzo:
            break
        pagina += 1

    print(f"\n-> Expedientes guardados: {len(resultados)} | Omitidos (sin cita): {omitidos}")
    return resultados

# -----------------------------------------------------------------------------
# EXPORTAR A EXVEL
# -----------------------------------------------------------------------------

def exportar_excel(resultados, base="expedientes", tz="America/Mexico_City"):

    schema = [
        ("numero_cuenta", "Número de cuenta"),
        ("nombre", "Nombre completo"),
        ("opcion_titulacion", "Opción de titulación"),
        ("correo", "Correo"),
        ("plantel", "Plantel"),
        ("carrera", "Carrera"),
        ("plan_estudios", "Plan de estudios"),
        ("cita_fecha", "Cita programada"),
    ]
    raw_cols = [k for k, _ in schema]
    headers  = [h for _, h in schema]

    df = pd.DataFrame(resultados)

    for k in raw_cols:
        if k not in df.columns:
            df[k] = ""
    df = df[raw_cols]
    for k in raw_cols:
        df[k] = df[k].astype("string").fillna("")

    df.columns = headers

    try:
        from zoneinfo import ZoneInfo
        now = datetime.now(ZoneInfo(tz))
    except Exception:
        now = datetime.now()
    ts = now.strftime("%Y%m%d-%H%M%S")

    ruta_xlsx = f"{base}-{ts}.xlsx"
    try:
        df.to_excel(ruta_xlsx, index=False, sheet_name="Expedientes")
        print(f"-> Excel generado: {ruta_xlsx}")
    except ModuleNotFoundError:
        print("-> Falta 'openpyxl', instálalo con: pip install openpyxl")
        ruta_csv = f"{base}-{ts}.csv"
        df.to_csv(ruta_csv, index=False, encoding="utf-8-sig")
        print(f"-> CSV de respaldo generado: {ruta_csv}")

# -----------------------------------------------------------------------------
# ESTRUCTURA FINAL
# -----------------------------------------------------------------------------

def main():
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'eager'
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=options
    )
    try:
        esperar_login_e_ir_a_seguimiento(driver)

        seleccionar_filtro_por_estado(driver)
        print(f"-> Filas visibles: {len(driver.find_elements(*SEL_FILAS_TABLA))}")

        resultados = recorrer_expedientes(driver)
        exportar_excel(resultados, base="expedientes")

        with open("expedientes.json", "w", encoding="utf-8") as f:
            json.dump(resultados, f, ensure_ascii=False, indent=4)

        print("\n-> Datos guardados en expedientes.json")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════╗
║   DESCARGA AUTOMÁTICA DE TASAS MUNICIPALES — PERGAMINO          ║
║   Sitio: https://tasas.pergamino.gob.ar/tasas-web               ║
╠══════════════════════════════════════════════════════════════════╣
║  INSTALACIÓN (una sola vez en la terminal de VS Code):          ║
║      pip install playwright pandas openpyxl                     ║
║      playwright install chromium                                ║
║                                                                  ║
║  PARA EJECUTAR:                                                  ║
║      python tasas_pergamino.py                                  ║
║                                                                  ║
║  PARA REPROCESAR SOLO ALGUNOS (ej: los que dieron error):        ║
║      python tasas_pergamino.py --nombres "Aguirrebarrena,Barriola"║
║      python tasas_pergamino.py --nros "11540,66492"             ║
╚══════════════════════════════════════════════════════════════════╝
"""

import argparse
import asyncio
import os
import re
import shutil
from pathlib import Path
from datetime import datetime

import pandas as pd
from playwright.async_api import async_playwright


# ═══════════════════════════════════════════════════════════════════
#  ⚙  CONFIGURACIÓN
# ═══════════════════════════════════════════════════════════════════

EXCEL_PATH = r"C:\Users\54247\OneDrive\Documentos\inmobiliaria\tasas municipales\Administraciones.xlsx"
SHEET_NAME = "Tasas"
SAVE_PATH  = Path(r"C:\Users\54247\OneDrive\Documentos\inmobiliaria\tasas municipales")
HEADLESS   = False
URL        = "https://tasas.pergamino.gob.ar/tasas-web"
CUOTAS_A_DESCARGAR = 2

# ── Columnas del Excel ───────────────────────────────────────────
COL_LOCADOR   = "Locador/a"
COL_DOMICILIO = "Domicilio"
COL_INMUEBLE  = "Inmueble n°"

INVALIDOS = {"-", "a designar", "x", "?", "nan", "", "none"}

# Texto de las tasas en el sitio: "16 - Tasa por servicios sanitarios", etc.
PATRON_TASA = re.compile(r"\d+\s*-\s*Tasa", re.I)


# ═══════════════════════════════════════════════════════════════════
#  HELPERS GENERALES
# ═══════════════════════════════════════════════════════════════════

def limpiar(s) -> str:
    if pd.isna(s):
        return ""
    return str(s).strip().replace("*", "").strip()


def titulizar(s: str) -> str:
    """'avenida de mayo 855' → 'Avenida De Mayo 855'"""
    return s.strip().title() if s else ""


def parsear_inmueble(valor) -> tuple:
    """'66492 - 69223' → ('66492','69223') | '11540' → ('11540', None)"""
    s = limpiar(valor)
    if s.lower() in INVALIDOS:
        return None, None
    partes = [p.strip() for p in s.split(" - ")]
    n1 = partes[0] if partes[0].isdigit() else None
    n2 = partes[1] if len(partes) > 1 and partes[1].isdigit() else None
    return n1, n2


def parsear_locador(raw: str) -> tuple:
    """'García Otero Paula' → ('García Otero', 'Paula')"""
    s = limpiar(raw)
    if not s:
        return None, ""
    partes = s.split()
    if len(partes) == 1:
        return partes[0], ""
    nombre = ""
    apellido_p = list(partes)
    for i in range(len(partes) - 1, 0, -1):
        t = partes[i]
        if not (t.endswith(".") and len(t) <= 3) and t.lower() not in {"y","e","de","del","la","el"}:
            nombre = t
            apellido_p = partes[:i]
            break
    return (" ".join(apellido_p) if apellido_p else partes[0]), nombre


def calcular_bases(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_ap"]  = df["_lc"].apply(lambda x: parsear_locador(x)[0])
    df["_nom"] = df["_lc"].apply(lambda x: parsear_locador(x)[1])
    dup = (df.groupby("_ap")["_lc"].nunique().pipe(lambda s: set(s[s > 1].index)))
    bases = {}
    for ap, grp in df.groupby("_ap"):
        if ap not in dup:
            for idx in grp.index:
                bases[idx] = ap
        else:
            noms = grp.drop_duplicates("_lc")[["_lc","_nom"]]
            n_list = noms["_nom"].tolist()
            ml = 2
            while ml < max((len(n) for n in n_list), default=2):
                if len({n[:ml] for n in n_list}) == len(n_list):
                    break
                ml += 1
            mapa = {r["_lc"]: f"{ap} {r['_nom'][:ml]}" for _, r in noms.iterrows()}
            for idx, row in grp.iterrows():
                bases[idx] = mapa.get(row["_lc"], ap)
    df["_base"] = pd.Series(bases)
    return df


def cargar_excel(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET_NAME, header=1, dtype=str)
    df.columns = df.columns.str.strip()
    for col in [COL_LOCADOR, COL_DOMICILIO, COL_INMUEBLE]:
        if col not in df.columns:
            raise ValueError(f"❌ Columna '{col}' no encontrada. Disponibles: {list(df.columns)}")
    df = df.applymap(lambda x: str(x).strip() if pd.notna(x) else x)
    df = df[df[COL_LOCADOR].notna() & (df[COL_LOCADOR].str.strip() != "")].reset_index(drop=True)

    def no_descargar(row):
        return any(isinstance(v, str) and "no descargar" in v.lower() for v in row)

    df["_skip"] = df.apply(no_descargar, axis=1)
    df["_lc"]   = df[COL_LOCADOR].apply(limpiar)
    df[["_n1","_n2"]] = df[COL_INMUEBLE].apply(lambda x: pd.Series(parsear_inmueble(x)))
    df["_skip"] |= df["_n1"].isna()
    df = calcular_bases(df)
    multi = (df[~df["_skip"]].groupby("_lc").size().pipe(lambda s: set(s[s > 1].index)))
    df["_multi"] = df["_lc"].isin(multi)
    return df


# ═══════════════════════════════════════════════════════════════════
#  NOMENCLATURA DE ARCHIVOS
# ═══════════════════════════════════════════════════════════════════

def nombre_pdf(base: str, dom, sigla: str, fecha: str, coch: bool) -> str:
    """
    Mujica, sin cochera, 1 alquiler:        Mujica - TSS - 14-07-2026
    Mujica, cochera, 1 alquiler:            Mujica Coch - TSS - 14-07-2026
    Mujica, sin cochera, varios alquileres: Mujica - Avenida De Mayo 855 - TSS - 14-07-2026
    Mujica, cochera, varios alquileres:     Mujica Coch - Avenida De Mayo 855 - TSS - 14-07-2026
    """
    dom_f = titulizar(str(dom)) if dom else None
    if coch and dom_f:
        b = f"{base} Coch - {dom_f}"
    elif coch:
        b = f"{base} Coch"
    elif dom_f:
        b = f"{base} - {dom_f}"
    else:
        b = base
    return f"{b} - {sigla} - {fecha}"


# ═══════════════════════════════════════════════════════════════════
#  AUTOMATIZACIÓN WEB — selectores confirmados del sitio real
# ═══════════════════════════════════════════════════════════════════

async def navegar_y_buscar(page, nro: str):
    """Abre el sitio, selecciona 'Conozco el número de imponible', ingresa el nro y busca."""
    await page.goto(URL, wait_until="networkidle", timeout=45000)

    # Esperar que cargue el formulario
    await page.wait_for_selector('label[for="id_inmuebles"]', timeout=30000)
    await asyncio.sleep(0.5)

    # Hacer click en el label del radio "Conozco el número de imponible"
    # (se filtra por ":visible"/".first" porque el sitio deja un elemento
    #  oculto duplicado con el mismo selector, lo que provoca un "strict
    #  mode violation" si no se filtra)
    radio = page.locator('label[for="id_inmuebles"]:visible').first
    campo = page.locator('input[placeholder="Ingrese el número de imponible"]:visible').first
    ok = await click_y_confirmar(radio, campo, intentos=3, espera=3, timeout_confirmar=10000)
    if not ok:
        print("      ⚠  No apareció el campo para ingresar el número de imponible")

    # Ingresar el número de imponible
    await campo.fill(str(nro).strip())
    await asyncio.sleep(0.3)

    # Click en Buscar (button type=submit) — se confirma que aparezca la fila
    # con la celda <b>{nro}</b> antes de seguir; el sitio suele tardar.
    patron_nro = re.compile(rf"^\s*{re.escape(str(nro))}\s*$")
    fila_resultado = page.locator("td b:visible").filter(has_text=patron_nro).first
    buscar = page.locator('button[type="submit"]:visible').first
    ok = await click_y_confirmar(buscar, fila_resultado, intentos=4, espera=6, timeout_confirmar=12000)
    if not ok:
        print(f"      ⚠  El resultado de la búsqueda para {nro} está tardando más de lo esperado")
    await asyncio.sleep(1)


async def _fila_o_elemento(elemento):
    """Devuelve el <tr> ancestro del elemento si existe, o el elemento mismo."""
    fila = elemento.locator("xpath=ancestor::tr[1]")
    return fila if await fila.count() > 0 else elemento


async def click_y_confirmar(click_locator, confirmar_locator,
                             intentos: int = 4, espera: float = 6,
                             timeout_confirmar: int = 12000) -> bool:
    """
    Hace click en click_locator y espera a que confirmar_locator se haga
    visible (señal de que la página reaccionó al click). Si no aparece a
    tiempo, espera 'espera' segundos y vuelve a hacer click, hasta 'intentos'
    veces en total. El sitio suele tardar, así que conviene darle tiempo
    en vez de fallar al primer intento.
    """
    for intento in range(1, intentos + 1):
        try:
            await click_locator.click()
        except Exception:
            pass
        try:
            await confirmar_locator.wait_for(state="visible", timeout=timeout_confirmar)
            return True
        except Exception:
            if intento < intentos:
                await asyncio.sleep(espera)
    return False


async def expandir_inmueble(page, nro: str):
    """
    Hace click en la fila que contiene <b>{nro}</b> para expandir las tasas
    del inmueble. El sitio suele tardar en responder, así que reintenta el
    click con paciencia hasta que aparezcan las tasas.
    """
    patron_nro = re.compile(rf"^\s*{re.escape(str(nro))}\s*$")
    celda = page.locator("td b:visible").filter(has_text=patron_nro).first
    objetivo = await _fila_o_elemento(celda)
    tasas_visibles = page.locator("tr.fila-consolidado a.link-consolidado b:visible").filter(has_text=PATRON_TASA).first

    ok = await click_y_confirmar(objetivo, tasas_visibles, intentos=5, espera=6, timeout_confirmar=12000)
    if not ok:
        print(f"      ⚠  Las tasas de {nro} no terminaron de cargar después de varios intentos")


async def obtener_tasas(page) -> list:
    """
    Devuelve [(sigla, texto_tasa), ...] de las tasas disponibles, en el orden
    en que aparecen en la página (la primera es la de más arriba).
    Solo incluye las 4 tasas conocidas. Ignora 'Contribución de mejoras' y otras.
    """
    elementos = page.locator("tr.fila-consolidado a.link-consolidado b:visible").filter(has_text=PATRON_TASA)
    count = await elementos.count()
    resultado, vistos = [], set()

    for i in range(count):
        try:
            texto = (await elementos.nth(i).inner_text()).strip()
        except Exception:
            continue
        low = texto.lower()

        # TASS debe ir ANTES que TSS para no confundirse
        if "adicional" in low and "sanitari" in low:
            sigla = "TASS"
        elif "sanitari" in low:
            sigla = "TSS"
        elif "limpieza" in low or "conservaci" in low:
            sigla = "TLC"
        elif "alumbrado" in low:
            sigla = "TAP"
        else:
            continue  # saltar tasas desconocidas (ej: Contribución de mejoras)

        if sigla not in vistos:
            vistos.add(sigla)
            resultado.append((sigla, texto))

    return resultado


async def click_tasa(page, texto_tasa: str):
    """
    Hace click en el link de la tasa (<a class="link-consolidado"> dentro de
    <tr class="fila-consolidado">) para abrir su detalle, reintentando con
    paciencia hasta que aparezca la sección de años/cuotas (label[for^="a"]).
    """
    link = page.locator("tr.fila-consolidado a.link-consolidado:visible").filter(
        has_text=re.compile(re.escape(texto_tasa.strip()), re.I)
    ).first
    detalle_visible = page.locator('label[for^="a"]:visible').first

    ok = await click_y_confirmar(link, detalle_visible, intentos=5, espera=6, timeout_confirmar=12000)
    if not ok:
        print(f"         ⚠  El detalle de la tasa no terminó de cargar después de varios intentos")
    await asyncio.sleep(1)


async def seleccionar_ultimo_anio(page) -> bool:
    """
    Busca los checkboxes de año (id="a2022", "a2026", etc.) y expande el más
    reciente. El checkbox/label solo marca la selección: lo que realmente
    expande la fila es un click en su celda de importe (td.text-right).
    El sitio tarda en cargar las cuotas de ese año, por eso se reintenta con
    paciencia.
    Retorna True si encontró años, False si las cuotas ya son directamente visibles.
    """
    await asyncio.sleep(0.8)

    # Se piden todos los atributos "for" en una sola llamada al navegador
    # (mucho más rápido que consultarlos uno por uno) para no perder tiempo
    # antes de poder hacer el click ni bien esté disponible.
    fors = await page.locator("label[for]:visible").evaluate_all(
        "els => els.map(el => el.getAttribute('for'))"
    )
    años = {}
    for for_attr in fors:
        if for_attr and re.match(r'^a\d{4}$', for_attr):
            años[int(for_attr[1:])] = for_attr

    if not años:
        return False  # Las cuotas ya son directamente visibles (sin agrupación por año)

    ultimo = max(años.keys())
    label_anio = page.locator(f"label[for='a{ultimo}']:visible").first
    fila_anio = label_anio.locator("xpath=ancestor::tr[1]")
    celda_importe = fila_anio.locator("td.text-right").first
    cuota_visible = page.locator(f"label[for^='a{ultimo}c']:visible").first

    # El sitio suele tardar bastante (~20s) en traer las cuotas de ese año.
    # Se le da un primer margen largo para evitar un reintento innecesario
    # que podría volver a colapsar la fila en vez de ayudar.
    ok = await click_y_confirmar(celda_importe, cuota_visible, intentos=3, espera=5, timeout_confirmar=30000)
    if not ok:
        print(f"         ⚠  Las cuotas del año {ultimo} no terminaron de cargar")

    return True


async def obtener_cuotas(page) -> list:
    """
    Devuelve [((año, num_cuota), for_attr), ...] ordenadas de menor a mayor,
    a partir de los checkboxes id="a{año}c{num}" (ej: "a2026c7" → Cuota 7 de 2026).
    """
    patron = re.compile(r'^a(\d{4})c(\d+)$')
    fors = await page.locator("label[for]:visible").evaluate_all(
        "els => els.map(el => el.getAttribute('for'))"
    )
    cuotas = {}
    for for_attr in fors:
        m = patron.match(for_attr or "")
        if m:
            año, num = int(m.group(1)), int(m.group(2))
            cuotas[(año, num)] = for_attr

    return sorted(cuotas.items())


async def expandir_cuota(page, for_attr: str, confirmar: bool = False):
    """
    Expande o colapsa la fila de una cuota. El checkbox/label solo marca la
    selección: lo que realmente expande la fila es un click en su celda de
    importe (td.text-right). Si confirmar=True (al expandir, no al
    colapsar), reintenta con paciencia hasta que aparezca el ícono de descarga.
    """
    label = page.locator(f"label[for='{for_attr}']:visible").first
    fila = label.locator("xpath=ancestor::tr[1]")
    celda_importe = fila.locator("td.text-right").first

    if confirmar:
        descarga = page.locator("i.fa.fa-download:visible").first
        ok = await click_y_confirmar(celda_importe, descarga, intentos=3, espera=5, timeout_confirmar=25000)
        if not ok:
            print(f"         ⚠  No apareció el ícono de descarga después de expandir la cuota")
    else:
        await celda_importe.click()
        await asyncio.sleep(2)


async def obtener_fecha(page) -> str:
    """Extrae la fecha de vencimiento del comprobante expandido (formato DD/MM/YYYY en el sitio)."""
    patron = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")
    try:
        # Las filas de comprobante son tr.tree-slave.in
        filas = page.locator("tr.tree-slave.in")
        count = await filas.count()
        for i in range(count):
            texto = await filas.nth(i).inner_text()
            m = patron.search(texto)
            if m:
                d, mes, y = m.groups()
                return f"{d.zfill(2)}-{mes.zfill(2)}-{y}"
    except Exception:
        pass
    return datetime.now().strftime("%d-%m-%Y")


async def descargar_pdf(page, context, nombre: str, save_path: Path, tmp_path: Path) -> bool:
    """
    Hace click en i.fa.fa-download y guarda el PDF.
    Playwright intercepta automáticamente el diálogo 'Guardar como' con accept_downloads=True.
    """
    destino = save_path / f"{nombre}.pdf"
    icono = page.locator("i.fa.fa-download:visible").first

    if await icono.count() == 0:
        print(f"      ✗  Ícono de descarga no encontrado")
        return False

    try:
        async with page.expect_download(timeout=45000) as dl_info:
            await icono.click()
        dl = await dl_info.value
        tmp = tmp_path / (dl.suggested_filename or "tasa.pdf")
        await dl.save_as(tmp)
        shutil.move(str(tmp), str(destino))
        print(f"      ✓  {destino.name}")
        return True
    except Exception as e:
        print(f"      ✗  Error al descargar: {e}")
        return False


async def volver(page, nro: str):
    """
    Hace click en el ícono de flecha 'Volver' (tooltip='Volver'). Esto no
    lleva directo a la lista de tasas: vuelve a la vista de resultados de
    búsqueda, donde reaparece la celda <b>{nro}</b> del inmueble (hay que
    volver a expandirla para ver otra vez la lista de tasas).
    """
    icono = page.locator('i.fa.fa-arrow-circle-left:visible').first
    patron_nro = re.compile(rf"^\s*{re.escape(str(nro))}\s*$")
    fila_nro = page.locator("td b:visible").filter(has_text=patron_nro).first
    ok = await click_y_confirmar(icono, fila_nro, intentos=5, espera=6, timeout_confirmar=12000)
    if not ok:
        print("      ⚠  No se pudo confirmar el regreso a la vista del inmueble")


async def procesar_inmueble(page, context, nro: str, base: str,
                             dom, coch: bool, save_path: Path, tmp_path: Path) -> list:
    """
    Flujo completo para un número de imponible:
    1. Buscar → expandir inmueble → listar tasas (una sola vez)
    2. Por cada tasa: entrar, seleccionar último año, descargar últimas 2 cuotas,
       y volver (botón flecha) a la lista de tasas para pasar a la siguiente.

    Devuelve una lista de deudas anteriores encontradas (cuotas del mismo año
    que quedaron sin descargar por no estar entre las últimas dos), cada una
    como {"sigla": ..., "año": ..., "meses": [...]}.
    """
    sufijo = " [cochera]" if coch else ""
    print(f"\n      Nro {nro}  —  {base}{sufijo}")
    deudas_anteriores = []

    # ── Paso 1: buscar el inmueble y obtener lista de tasas disponibles ──
    # (navegar_y_buscar y expandir_inmueble ya reintentan con paciencia
    #  internamente si el sitio tarda en responder)
    await navegar_y_buscar(page, nro)
    await expandir_inmueble(page, nro)
    tasas = await obtener_tasas(page)

    if not tasas:
        print(f"      ⚠  Sin tasas para {nro}")
        return deudas_anteriores

    print(f"      Tasas disponibles: {[s for s,_ in tasas]}")

    # ── Paso 2: procesar cada tasa por separado, sin re-buscar ──────
    for i, (sigla, texto_tasa) in enumerate(tasas):
        if i > 0:
            # "volver" deja la vista de resultados (celda <b>{nro}</b>), no
            # la lista de tasas directamente: hay que expandir el inmueble
            # de nuevo (con paciencia) antes de elegir la siguiente tasa.
            await expandir_inmueble(page, nro)

        print(f"\n      ── {sigla}: {texto_tasa[:50]}")

        await click_tasa(page, texto_tasa)

        # Seleccionar el año más reciente (si aplica)
        await seleccionar_ultimo_anio(page)

        # Obtener cuotas
        cuotas = await obtener_cuotas(page)
        if not cuotas:
            print(f"         ⚠  Sin cuotas para {sigla}")
            await volver(page, nro)
            continue

        ultimas = cuotas[-CUOTAS_A_DESCARGAR:]
        print(f"         Cuotas: {[n for n,_ in cuotas]}  →  descargando: {[n for n,_ in ultimas]}")

        # Cuotas del mismo año que no están entre las últimas dos: son deuda
        # anterior sin descargar (el sitio solo lista en "Deuda" lo impago).
        anteriores = cuotas[:-CUOTAS_A_DESCARGAR] if len(cuotas) > CUOTAS_A_DESCARGAR else []
        if anteriores:
            año_deuda = anteriores[0][0][0]
            meses_deuda = sorted(num for (_, num), _ in anteriores)
            deudas_anteriores.append({"sigla": sigla, "año": año_deuda, "meses": meses_deuda})
            print(f"         ⚠  Adeuda además: {sigla} {año_deuda} cuotas {meses_deuda}")

        for (año, num_cuota), for_attr in ultimas:
            # Expandir la cuota (celda de importe), esperando con paciencia
            # a que aparezca el ícono de descarga
            await expandir_cuota(page, for_attr, confirmar=True)

            # Obtener la fecha de vencimiento
            fecha = await obtener_fecha(page)

            # Construir nombre del archivo y descargar (una descarga por cuota)
            nombre = nombre_pdf(base, dom, sigla, fecha, coch)
            await descargar_pdf(page, context, nombre, save_path, tmp_path)

            # Colapsar antes de la siguiente cuota
            await expandir_cuota(page, for_attr)
            await asyncio.sleep(0.3)

        # Volver a la vista del inmueble para pasar a la siguiente tasa
        await volver(page, nro)

    return deudas_anteriores


# ═══════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════

def parse_args():
    parser = argparse.ArgumentParser(description="Descarga de tasas municipales de Pergamino")
    parser.add_argument("--nombres", type=str, default=None,
                         help="Solo procesar locadores cuyo nombre contenga alguno de estos textos, separados por coma")
    parser.add_argument("--nros", type=str, default=None,
                         help="Solo procesar estos números de inmueble, separados por coma")
    return parser.parse_args()


async def main():
    args = parse_args()

    sep = "═" * 62
    print(f"\n{sep}")
    print("   TASAS MUNICIPALES — Municipalidad de Pergamino")
    print(f"   {datetime.now().strftime('%d/%m/%Y  %H:%M:%S')}")
    print(sep)

    SAVE_PATH.mkdir(parents=True, exist_ok=True)
    tmp_path = SAVE_PATH / "_tmp"
    tmp_path.mkdir(exist_ok=True)

    print("\n📋 Leyendo Excel...")
    df = cargar_excel(EXCEL_PATH)
    proc = df[~df["_skip"]]
    salt = df[df["_skip"]]

    if args.nombres:
        nombres = [n.strip().lower() for n in args.nombres.split(",") if n.strip()]
        proc = proc[proc[COL_LOCADOR].str.lower().apply(lambda x: any(n in x for n in nombres))]
        print(f"   🔎  Filtrando por nombre: {nombres}")

    if args.nros:
        nros = [n.strip() for n in args.nros.split(",") if n.strip()]
        proc = proc[proc["_n1"].isin(nros) | proc["_n2"].isin(nros)]
        print(f"   🔎  Filtrando por número de inmueble: {nros}")

    print(f"   ✓  {len(proc)} inmuebles a procesar")
    if not salt.empty:
        print(f"   ⚠  {len(salt)} filas salteadas (sin número o con 'No descargar')")

    errores = 0
    resumen_deudas = []  # [(base, sufijo, [deuda, ...]), ...]

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS)
        context = await browser.new_context(accept_downloads=True)
        page    = await context.new_page()

        total = len(proc)
        for i, (_, row) in enumerate(proc.iterrows(), start=1):
            base = row["_base"]
            dom  = titulizar(str(row[COL_DOMICILIO]).strip()) if row["_multi"] and pd.notna(row.get(COL_DOMICILIO)) else None

            print(f"\n[{i}/{total}] {limpiar(row[COL_LOCADOR])}")

            # Inmueble principal
            try:
                deudas = await procesar_inmueble(
                    page, context, row["_n1"],
                    base, dom, False, SAVE_PATH, tmp_path
                )
                if deudas:
                    resumen_deudas.append((base, "", deudas))
            except Exception as e:
                print(f"      ✗  Error en inmueble de {base}: {e}")
                errores += 1

            # Cochera (si existe)
            nro2 = str(row.get("_n2", "")).strip()
            if nro2 and nro2.lower() not in ("nan", "", "none"):
                try:
                    deudas_coch = await procesar_inmueble(
                        page, context, nro2,
                        base, dom, True, SAVE_PATH, tmp_path
                    )
                    if deudas_coch:
                        resumen_deudas.append((base, " [cochera]", deudas_coch))
                except Exception as e:
                    print(f"      ✗  Error en cochera de {base}: {e}")
                    errores += 1

        await browser.close()

    shutil.rmtree(tmp_path, ignore_errors=True)

    print(f"\n{sep}")
    print(f"   {'✅  Sin errores' if errores == 0 else f'⚠   {errores} error(es)'}")
    print(f"   📁  {SAVE_PATH}")
    print(f"{sep}\n")

    if resumen_deudas:
        print(f"{sep}")
        print("   ⚠  LOCADORES CON CUOTAS ANTERIORES ADEUDADAS")
        print("   (del mismo año descargado, sin contar las dos últimas)")
        print(sep)
        for base, sufijo, deudas in resumen_deudas:
            print(f"\n   {base}{sufijo}")
            for d in deudas:
                print(f"      {d['sigla']} {d['año']}  →  cuotas {d['meses']}")
        print(f"\n{sep}\n")
    else:
        print("   ✓  Nadie tiene cuotas anteriores pendientes (más allá de las descargadas)\n")


if __name__ == "__main__":
    asyncio.run(main())
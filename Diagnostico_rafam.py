#!/usr/bin/env python3
"""
SCRIPT DE DIAGNÓSTICO — Portal RAFAM Pergamino
================================================
Este script NO descarga nada. Solo abre el navegador,
entra al sitio con UN número de inmueble, y guarda:
  - capturas de pantalla en cada paso
  - el HTML completo de la página
  - todos los elementos interactivos que encuentra

Ejecutalo y mandame las imágenes + el archivo html_capturado.txt
"""

import asyncio
from pathlib import Path
from playwright.async_api import async_playwright

# ── ÚNICO VALOR QUE TENÉS QUE CAMBIAR ──────────────────────────
# Poné cualquier número de inmueble válido de tu Excel
NRO_PRUEBA = "11540"
# ────────────────────────────────────────────────────────────────

URL        = "https://tasas.pergamino.gob.ar/tasas-web"
CARPETA    = Path.home() / "Documents" / "diagnostico_rafam"

async def main():
    CARPETA.mkdir(exist_ok=True)
    print(f"\nGuardando diagnóstico en: {CARPETA}\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=500)
        context = await browser.new_context()
        page    = await context.new_page()

        # ── PASO 1: abrir el sitio ──────────────────────────────
        print("PASO 1: Abriendo el sitio...")
        await page.goto(URL, wait_until="networkidle", timeout=30000)
        await asyncio.sleep(3)
        await page.screenshot(path=CARPETA / "01_inicio.png", full_page=True)
        print("  ✓ Captura guardada: 01_inicio.png")

        # ── PASO 2: guardar el HTML inicial ─────────────────────
        html_inicial = await page.content()
        (CARPETA / "html_01_inicio.txt").write_text(html_inicial, encoding="utf-8")
        print("  ✓ HTML guardado: html_01_inicio.txt")

        # ── PASO 3: listar todos los elementos clickeables ──────
        print("\nPASO 2: Buscando elementos interactivos...")
        reporte = []

        # Botones
        botones = await page.locator("button").all()
        reporte.append(f"\n=== BOTONES ({len(botones)}) ===")
        for b in botones:
            try:
                txt  = (await b.inner_text()).strip()[:80]
                cls  = await b.get_attribute("class") or ""
                reporte.append(f"  BOTON: '{txt}'  class='{cls[:60]}'")
            except Exception:
                pass

        # Inputs
        inputs = await page.locator("input").all()
        reporte.append(f"\n=== INPUTS ({len(inputs)}) ===")
        for inp in inputs:
            try:
                tipo  = await inp.get_attribute("type") or "text"
                name  = await inp.get_attribute("name") or ""
                ph    = await inp.get_attribute("placeholder") or ""
                cls   = await inp.get_attribute("class") or ""
                reporte.append(f"  INPUT type='{tipo}' name='{name}' placeholder='{ph}' class='{cls[:50]}'")
            except Exception:
                pass

        # Textos de la página (primeros 3000 chars)
        texto_pagina = await page.locator("body").inner_text()
        reporte.append(f"\n=== TEXTO VISIBLE (primeros 2000 chars) ===")
        reporte.append(texto_pagina[:2000])

        (CARPETA / "elementos_paso1.txt").write_text("\n".join(reporte), encoding="utf-8")
        print("  ✓ Reporte guardado: elementos_paso1.txt")

        # ── PASO 4: intentar hacer click en opciones visibles ───
        print("\nPASO 3: Intentando interactuar...")

        # Buscar cualquier texto que contenga "imponible"
        texto_imponible = page.get_by_text("imponible", exact=False)
        count = await texto_imponible.count()
        print(f"  Elementos con 'imponible': {count}")

        if count > 0:
            for i in range(min(count, 5)):
                try:
                    t = await texto_imponible.nth(i).inner_text()
                    print(f"    [{i}] '{t.strip()[:80]}'")
                except Exception:
                    pass

        # Intentar click en el primero
        if count > 0:
            try:
                await texto_imponible.first.click()
                await asyncio.sleep(2)
                await page.screenshot(path=CARPETA / "02_despues_click_imponible.png", full_page=True)
                print("  ✓ Captura guardada: 02_despues_click_imponible.png")
                html2 = await page.content()
                (CARPETA / "html_02_despues_click.txt").write_text(html2, encoding="utf-8")
            except Exception as e:
                print(f"  ✗ No se pudo hacer click: {e}")

        # ── PASO 5: buscar campo de input y escribir el número ──
        print(f"\nPASO 4: Intentando ingresar el número {NRO_PRUEBA}...")
        campos = await page.locator("input[type='text'], input[type='number'], input:not([type])").all()
        print(f"  Campos de texto encontrados: {len(campos)}")

        if campos:
            try:
                await campos[0].fill(NRO_PRUEBA)
                await asyncio.sleep(1)
                await page.screenshot(path=CARPETA / "03_numero_ingresado.png", full_page=True)
                print("  ✓ Captura guardada: 03_numero_ingresado.png")
            except Exception as e:
                print(f"  ✗ Error: {e}")

        # ── PASO 6: buscar y clickear "Buscar" ──────────────────
        print("\nPASO 5: Buscando botón 'Buscar'...")
        buscar_btn = page.get_by_role("button", name="Buscar")
        count_buscar = await buscar_btn.count()
        print(f"  Botones 'Buscar' encontrados: {count_buscar}")

        # También buscar por texto
        cualquier_buscar = page.get_by_text("Buscar", exact=False)
        count_cualquier = await cualquier_buscar.count()
        print(f"  Elementos con texto 'Buscar': {count_cualquier}")

        if count_buscar > 0:
            try:
                await buscar_btn.first.click()
                await page.wait_for_load_state("networkidle")
                await asyncio.sleep(3)
                await page.screenshot(path=CARPETA / "04_resultados.png", full_page=True)
                html4 = await page.content()
                (CARPETA / "html_04_resultados.txt").write_text(html4, encoding="utf-8")
                print("  ✓ Capturas guardadas: 04_resultados.png y html_04_resultados.txt")

                # Listar qué apareció
                texto_resultado = await page.locator("body").inner_text()
                print(f"\n  TEXTO EN PANTALLA DESPUÉS DE BUSCAR:")
                print("  " + texto_resultado[:1500].replace("\n", "\n  "))

            except Exception as e:
                print(f"  ✗ Error al buscar: {e}")
        elif count_cualquier > 0:
            try:
                await cualquier_buscar.first.click()
                await page.wait_for_load_state("networkidle")
                await asyncio.sleep(3)
                await page.screenshot(path=CARPETA / "04_resultados.png", full_page=True)
                print("  ✓ Captura guardada: 04_resultados.png")
            except Exception as e:
                print(f"  ✗ Error: {e}")
        else:
            await page.screenshot(path=CARPETA / "04_sin_boton_buscar.png", full_page=True)
            print("  ✗ No se encontró botón Buscar. Captura guardada: 04_sin_boton_buscar.png")

        # ── PASO 7: HTML final ───────────────────────────────────
        html_final = await page.content()
        (CARPETA / "html_final.txt").write_text(html_final, encoding="utf-8")

        reporte2 = []
        botones2 = await page.locator("button").all()
        reporte2.append(f"\n=== BOTONES DESPUÉS DE BUSCAR ({len(botones2)}) ===")
        for b in botones2:
            try:
                txt = (await b.inner_text()).strip()[:80]
                cls = await b.get_attribute("class") or ""
                reporte2.append(f"  '{txt}'  class='{cls[:60]}'")
            except Exception:
                pass

        inputs2 = await page.locator("input").all()
        reporte2.append(f"\n=== INPUTS DESPUÉS DE BUSCAR ({len(inputs2)}) ===")
        for inp in inputs2:
            try:
                tipo = await inp.get_attribute("type") or "text"
                name = await inp.get_attribute("name") or ""
                cls  = await inp.get_attribute("class") or ""
                reporte2.append(f"  type='{tipo}' name='{name}' class='{cls[:60]}'")
            except Exception:
                pass

        (CARPETA / "elementos_paso2.txt").write_text("\n".join(reporte2), encoding="utf-8")

        print("\n" + "═"*55)
        print("  DIAGNÓSTICO COMPLETADO")
        print(f"  📁 Carpeta: {CARPETA}")
        print("  Archivos generados:")
        for f in sorted(CARPETA.iterdir()):
            print(f"    • {f.name}")
        print("═"*55)
        print("\n  ⚠  Mandame las IMÁGENES (01, 02, 03, 04)")
        print("     y el archivo elementos_paso2.txt")

        input("\n  Presioná ENTER para cerrar el navegador...")
        await browser.close()

if __name__ == "__main__":
    NRO_PRUEBA ="11540" 

    asyncio.run(main())
    
    
    
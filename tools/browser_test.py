#!/usr/bin/env python3
"""P47.2: verifica HTML autocontenido; render nativo Quarto requiere CLI.

La prueba carga el contenido completo en Chromium, sin insertar CSS/JS
externos; simula el uso del archivo movido de carpeta.
"""
from pathlib import Path
from playwright.sync_api import sync_playwright
import json
ROOT=Path(__file__).resolve().parent.parent
HTML=(ROOT/'ponencia.html').read_text(encoding='utf-8')
result={}
with sync_playwright() as pw:
  browser=pw.chromium.launch(executable_path='/usr/bin/chromium',headless=True,
                         args=['--no-sandbox','--disable-dev-shm-usage'])
  for width,height in [(1729,863),(1366,768)]:
    page=browser.new_page(viewport={'width':width,'height':height})
    errors=[]
    page.on('pageerror',lambda exc:errors.append(str(exc)))
    page.set_content(HTML,wait_until='domcontentloaded')
    assert page.locator('.slides>section').count()==20
    assert page.evaluate('window.CIIERoyal.total')==20
    assert page.evaluate("getComputedStyle(document.querySelector('.slide-canvas')).backgroundColor")=='rgb(255, 255, 255)'
    assert page.evaluate("getComputedStyle(document.querySelector('.ppt-text')).position")=='absolute'
    page.keyboard.press('ArrowRight')
    assert page.evaluate('window.CIIERoyal.current()')==2
    page.keyboard.press('Home')
    assert page.evaluate('window.CIIERoyal.current()')==1
    page.screenshot(path=str(ROOT/'capturas'/f'p47_2_portada_{width}.png'))
    page.keyboard.press('End')
    assert page.evaluate('window.CIIERoyal.current()')==20
    with page.expect_popup() as note:
      page.keyboard.press('s')
    assert note.value.locator('#speaker-title').inner_text()=='Diapositiva 20 de 20'
    assert page.locator('section.present .ref-link').get_attribute('href')=='referencias.html'
    note.value.close()
    assert not errors,errors
    result[str(width)]={'slides':20,'navegacion':True,'notas':True,'estilos_integrados':True,'js_errors':errors}
    page.close()
  browser.close()
(ROOT/'docs/P47_2_Validacion_Browser.json').write_text(json.dumps(result,ensure_ascii=False,indent=2))
print(json.dumps(result,ensure_ascii=False))

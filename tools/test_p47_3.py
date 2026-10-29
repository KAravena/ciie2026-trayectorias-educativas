from pathlib import Path
from playwright.sync_api import sync_playwright
import json, re
ROOT=Path(__file__).resolve().parents[1]
HTML=(ROOT/'ponencia.html').read_text(encoding='utf-8')
out={}
with sync_playwright() as pw:
  browser=pw.chromium.launch(executable_path='/usr/bin/chromium',headless=True,args=['--no-sandbox','--disable-dev-shm-usage'])
  for width,height in [(1729,863),(1366,768)]:
    page=browser.new_page(viewport={'width':width,'height':height})
    errors=[]
    page.on('pageerror',lambda e:errors.append(str(e)))
    page.set_content(HTML, wait_until='domcontentloaded')
    assert page.locator('.slides > section.academic-slide').count()==20
    assert page.locator('.offline-controls').count()==0
    assert page.locator('.offline-navigation .nav-arrow').count()==2
    assert page.locator('.offline-slide-number').inner_text()=='1 / 20'
    assert page.locator('.slides > section').first.locator('.cover-logo').count()==2
    assert page.locator('.slides > section').first.get_by_text('VII CONGRESO INTERDISCIPLINARIO DE INVESTIGACIÓN EN EDUCACIÓN').count()==0
    assert page.locator('#prev').is_disabled()
    assert page.locator('#next').is_enabled()
    page.screenshot(path=str(ROOT/'capturas'/f'p47_3_portada_{width}.png'))
    page.locator('#next').click()
    assert page.locator('#count').inner_text()=='2 / 20'
    page.keyboard.press('End')
    assert page.locator('#count').inner_text()=='20 / 20'
    assert page.locator('#next').is_disabled()
    page.screenshot(path=str(ROOT/'capturas'/f'p47_3_referencias_{width}.png'))
    page.keyboard.press('Home')
    with page.expect_popup() as speaker:
      page.keyboard.press('s')
    assert speaker.value.locator('#speaker-title').inner_text()=='Diapositiva 1 de 20'
    speaker.value.close()
    assert not errors,errors
    out[f'{width}x{height}']={'slides':20,'contador':'1 / 20 -> 2 / 20 -> 20 / 20','flechas_laterales':True,'notas_S':True,'errores_js':errors}
    page.close()
  browser.close()
(ROOT/'docs/P47_3_Validacion.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(out,ensure_ascii=False,indent=2))

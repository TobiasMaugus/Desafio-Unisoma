"""Visual checks against a local Streamlit server on port 8519 (no packaging).

Run with the project virtualenv; requires Playwright and installed Edge.
"""
from pathlib import Path
from io import BytesIO
import json

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, expect
from test_interface import synthetic

output = Path('tests/artifacts')

def capture(page, filename):
    page.wait_for_timeout(900)
    width = page.viewport_size['width']
    height = int(page.locator('.st-key-abc_page').bounding_box()['height']) + 160
    page.set_viewport_size({'width': width, 'height': height})
    page.locator('.st-key-abc_page').screenshot(path=str(output / filename))
    page.set_viewport_size({'width': width, 'height': 950})

with sync_playwright() as pw:
    browser = pw.chromium.launch(channel='msedge', headless=True)
    page = browser.new_page(viewport={'width': 1366, 'height': 950})
    errors = []
    page.on('pageerror', lambda error: errors.append(str(error)))
    page.goto('http://127.0.0.1:8519')
    page.get_by_role('heading', name='1 Vamos começar pela planilha').wait_for()
    page.get_by_role('button', name='Main menu').click()
    page.locator('[data-testid="stThemeSwitcher"]').get_by_role('menuitemradio', name='Dark').dispatch_event('click')
    page.wait_for_function("getComputedStyle(document.querySelector('.st-key-abc_upload')).backgroundColor === 'rgb(29, 41, 52)'")
    page.keyboard.press('Escape')
    page.get_by_role('button', name='Main menu').click()
    page.locator('[data-testid="stThemeSwitcher"]').get_by_role('menuitemradio', name='Light').dispatch_event('click')
    page.wait_for_function("getComputedStyle(document.querySelector('.st-key-abc_upload')).backgroundColor === 'rgb(255, 255, 255)'")
    page.keyboard.press('Escape')
    page.locator('input[type=file]').set_input_files({
        'name': 'conferencia.xlsx', 'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'buffer': synthetic()})
    page.get_by_role('button', name='Gerar ensalamento', exact=True).click()
    download = page.get_by_role('button', name='Baixar Excel completo (.xlsx)', exact=True)
    download.wait_for()
    page.get_by_text('Ensalamento concluído normalmente, sem inconsistências.', exact=True).wait_for()
    assert page.locator('.st-key-abc_distribution').bounding_box()['y'] < page.locator('.st-key-abc_results').bounding_box()['y']
    assert download.is_enabled()
    capture(page, 'interface_visual_desktop.png')
    page.get_by_role('button', name='Usar somente as salas necessárias', exact=True).click()
    expect(download).to_be_disabled()
    page.get_by_role('button', name='Recalcular ensalamento', exact=True).click()
    expect(download).to_be_enabled()
    book = load_workbook(BytesIO(synthetic()))
    book['Alunos']['A3'] = book['Alunos']['A2'].value
    content = BytesIO()
    book.save(content)
    page.locator('input[type=file]').set_input_files({
        'name': 'com_avisos.xlsx', 'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'buffer': content.getvalue()})
    page.get_by_role('button', name='Gerar ensalamento', exact=True).click()
    warning = page.get_by_text('Ensalamento concluído com 1 aviso(s) que precisam de atenção. A distribuição foi gerada; confira os registros antes de utilizá-la.', exact=True)
    warning.wait_for()
    assert page.get_by_text('Ensalamento concluído normalmente, sem inconsistências.', exact=True).count() == 0
    assert page.locator('[data-testid="stStatusWidget"]').count() == 0
    capture(page, 'interface_visual_avisos.png')
    for width in (1024, 768, 390):
        page.set_viewport_size({'width': width, 'height': 950})
        page.wait_for_timeout(400)
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
    capture(page, 'interface_visual_mobile.png')
    book['Alunos']['B3'] = 'INVALIDO'
    content = BytesIO()
    book.save(content)
    page.locator('input[type=file]').set_input_files({
        'name': 'bloqueado.xlsx', 'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'buffer': content.getvalue()})
    page.get_by_role('button', name='Gerar relatório de inconsistências', exact=True).click()
    page.get_by_role('button', name='Baixar planilha com inconsistências', exact=True).wait_for()
    assert page.get_by_text('Ensalamento concluído com', exact=False).count() == 0
    assert not errors, errors
    print(json.dumps({'status': 'OK', 'widths': [1366, 1024, 768, 390], 'checks': ['ordem das etapas', 'recalculo e downloads', 'verde sem avisos', 'amarelo com avisos', 'vermelho quando bloqueado'], 'browser_errors': errors}, ensure_ascii=False))
    browser.close()

"""Local browser checks, with synthetic data only.

Run a local Streamlit server on 8517, then execute with the project .venv.
Optional development dependency: playwright. Uses the installed Edge browser.
"""
from io import BytesIO
import json
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from test_interface import synthetic, workbook_digest

ARTIFACTS = Path(__file__).parent / 'artifacts'
URL = 'http://127.0.0.1:8517'


def capture(page, filename, selector=None):
    page.evaluate('document.fonts.ready')
    if selector:
        element = page.locator(selector)
        height = int(element.bounding_box()['height']) + 180
        page.set_viewport_size({'width':1366, 'height':max(900, height)})
        element.screenshot(path=str(ARTIFACTS/filename))
    else:
        page.screenshot(path=str(ARTIFACTS/filename))


def main():
    ARTIFACTS.mkdir(exist_ok=True)
    summary = {'sizes': [], 'downloads': [], 'external_requests': [], 'console_errors': []}
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(channel='msedge', headless=True)
        context = browser.new_context(viewport={'width':1366, 'height':900}, accept_downloads=True)
        def local_only(route):
            if route.request.url.startswith((URL, 'data:', 'blob:')):
                route.continue_()
            else:
                summary['external_requests'].append(route.request.url)
                route.abort()
        context.route('**/*', local_only)
        page = context.new_page()
        page.on('pageerror', lambda error: summary['console_errors'].append(str(error)))
        page.goto(URL)
        page.get_by_role('heading', name='1. Selecionar planilha').wait_for()
        brand = page.locator('img[alt="ABC Aprendiz — Centro Social de Educação para o Trabalho"]')
        assert brand.count() == 1 and brand.is_visible()
        assert page.locator('.st-key-abc_brand').evaluate('(e) => getComputedStyle(e).backgroundColor') == 'rgb(255, 255, 255)'
        # The native Streamlit menu keeps the two palettes one click away.
        page.get_by_role('button', name='Main menu').click()
        menu = page.locator('[data-testid="stThemeSwitcher"]')
        assert menu.get_by_role('menuitemradio', name='Light').is_visible()
        dark_option = menu.get_by_role('menuitemradio', name='Dark')
        assert dark_option.is_visible()
        dark_option.dispatch_event('click')
        page.wait_for_function("getComputedStyle(document.querySelector('.stApp')).backgroundColor === 'rgb(17, 26, 35)'")
        assert dark_option.get_attribute('aria-checked') == 'true'
        capture(page, 'tema_escuro.png')
        page.keyboard.press('Escape')
        page.get_by_role('button', name='Main menu').click()
        page.locator('[data-testid="stThemeSwitcher"]').get_by_role('menuitemradio', name='Light').dispatch_event('click')
        page.wait_for_function("getComputedStyle(document.querySelector('.stApp')).backgroundColor === 'rgb(245, 247, 249)'")
        assert page.evaluate('localStorage.length') > 0
        capture(page, 'depois_inicial.png')
        # Keyboard focus remains visible on the native upload button.
        page.get_by_role('button', name='Upload', exact=True).focus()
        page.keyboard.press('Tab')
        focused = page.evaluate('({tag:document.activeElement.tagName, outline:getComputedStyle(document.activeElement).outlineStyle})')
        summary['keyboard_focus'] = focused
        assert focused['outline'] != 'none'
        for width in (1366, 1024, 768, 390):
            page.set_viewport_size({'width':width, 'height':900})
            dimensions = page.evaluate('({viewport:innerWidth, document:document.documentElement.scrollWidth})')
            assert dimensions['document'] <= width
            summary['sizes'].append(dimensions)
            if width == 390:
                capture(page, 'depois_inicial_mobile.png')
        page.set_viewport_size({'width':1366, 'height':900})
        page.locator('input[type=file]').set_input_files({
            'name':'cenario_visual.xlsx', 'mimeType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'buffer':synthetic()})
        page.get_by_role('button', name='Gerar ensalamento', exact=True).wait_for()
        capture(page, 'depois_selecionado.png', '.st-key-abc_settings')
        page.get_by_role('button', name='Gerar ensalamento', exact=True).click()
        page.get_by_role('button', name='Baixar Excel completo (.xlsx)', exact=True).wait_for()
        page.get_by_role('heading', name='Distribuição por turma').wait_for()
        capture(page, 'depois_resultado.png', '.st-key-abc_results')
        for label, filename in [('Baixar Excel completo (.xlsx)', 'download_resultado.xlsx'),
                                ('Baixar resumo por turma (.csv)', 'download_resumo.csv')]:
            with page.expect_download() as pending:
                page.get_by_role('button', name=label, exact=True).click()
            download = pending.value
            download.save_as(ARTIFACTS/filename)
            summary['downloads'].append(download.suggested_filename)
        book = load_workbook(ARTIFACTS/'download_resultado.xlsx')
        assert book['Ensalamento'].max_row == 7
        assert list(next(book['Ensalamento'].values)) == ['Nome','Sexo','Disponibilidade','Turma']
        book.close()
        summary['workbook_digest'] = workbook_digest((ARTIFACTS/'download_resultado.xlsx').read_bytes())
        assert (ARTIFACTS/'download_resumo.csv').read_text(encoding='utf-8-sig').splitlines()[0] == 'Turma,Alunos,Meninas,Meninos,% meninas'
        page.get_by_role('tab', name='Lista de alunos', exact=True).click()
        page.get_by_role('textbox', name='Buscar aluno ou turma').fill('sem correspondencia')
        page.get_by_role('textbox', name='Buscar aluno ou turma').press('Enter')
        page.get_by_text('Nenhum aluno ou turma encontrado para essa busca. Tente outro nome ou letra.', exact=True).wait_for()
        page.get_by_role('tab', name='Visão geral', exact=True).click()
        page.set_viewport_size({'width':390, 'height':844})
        page.locator('.st-key-abc_results').scroll_into_view_if_needed()
        page.get_by_role('button', name='Baixar Excel completo (.xlsx)', exact=True).scroll_into_view_if_needed()
        capture(page, 'depois_resultado_mobile.png')
        assert page.evaluate('document.documentElement.scrollWidth') <= 390
        page.set_viewport_size({'width':1366, 'height':900})
        page.get_by_role('tab', name='Manual e ajuda', exact=True).click()
        with page.expect_download() as pending:
            page.get_by_role('button', name='Baixar manual (.pdf)', exact=True).click()
        pending.value.save_as(ARTIFACTS/'download_manual.pdf')
        assert (ARTIFACTS/'download_manual.pdf').read_bytes() == (Path(__file__).parents[1]/'MANUAL_USUARIO.pdf').read_bytes()
        summary['downloads'].append(pending.value.suggested_filename)
        page.get_by_role('tab', name='Ensalamento', exact=True).click()
        # Invalid data, using the same synthetic file, must expose diagnostics.
        book = load_workbook(BytesIO(synthetic()))
        book['Alunos']['B2'] = 'X'
        buffer = BytesIO()
        book.save(buffer)
        page.locator('input[type=file]').set_input_files({'name':'inconsistente.xlsx',
            'mimeType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'buffer':buffer.getvalue()})
        page.get_by_role('button', name='Gerar relatório de inconsistências', exact=True).wait_for()
        page.get_by_role('button', name='Gerar relatório de inconsistências', exact=True).click()
        page.get_by_role('button', name='Baixar planilha com inconsistências', exact=True).wait_for()
        capture(page, 'depois_bloqueado.png', '.st-key-abc_results')
        with page.expect_download() as pending:
            page.get_by_role('button', name='Baixar planilha com inconsistências', exact=True).click()
        pending.value.save_as(ARTIFACTS/'download_inconsistencias.xlsx')
        book = load_workbook(ARTIFACTS/'download_inconsistencias.xlsx')
        assert book['Ensalamento'].max_row == 1 and book['Inconsistências'].max_row > 1
        book.close()
        summary['downloads'].append(pending.value.suggested_filename)
        assert not summary['console_errors']
        (ARTIFACTS/'browser_checks.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        browser.close()


if __name__ == '__main__':
    main()

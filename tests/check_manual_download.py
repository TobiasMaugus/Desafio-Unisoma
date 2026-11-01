"""Check the updated manual and its download against Streamlit on port 8520."""
from io import BytesIO
from pathlib import Path

from playwright.sync_api import expect, sync_playwright
from pypdf import PdfReader

project = Path(__file__).resolve().parents[1]
source = (project / 'MANUAL_USUARIO.md').read_text(encoding='utf-8')
pdf = (project / 'MANUAL_USUARIO.pdf').read_bytes()
reader = PdfReader(BytesIO(pdf))
text = ' '.join(' '.join(page.extract_text().split()) for page in reader.pages)
for phrase in ('Vamos começar pela planilha', 'Confira os dados e gere as turmas',
               'Quer usar menos salas?', 'Confira e baixe o resultado',
               'barra horizontal no rodapé', 'barra vertical à direita',
               'Baixar Excel completo (.xlsx)', 'Baixar resumo da distribuição (PDF)',
               'Trocar dois alunos', 'Mover um aluno', 'Resolva os problemas mais comuns'):
    assert phrase in text, phrase
for obsolete in ('Outros formatos de download', '(.csv)', '\u2014', '\u2013'):
    assert obsolete not in source and obsolete not in text, obsolete
with sync_playwright() as pw:
    browser = pw.chromium.launch(channel='msedge', headless=True)
    page = browser.new_page(accept_downloads=True)
    page.goto('http://127.0.0.1:8520')
    page.get_by_role('tab', name='Manual e ajuda', exact=True).click()
    expect(page.get_by_role('heading', name='1. Conheça o aplicativo', exact=True)).to_be_visible()
    with page.expect_download() as pending:
        page.get_by_role('button', name='Baixar manual (.pdf)', exact=True).click()
    assert Path(pending.value.path()).read_bytes() == pdf
    expect(page.locator('[data-testid="stException"]')).to_have_count(0)
    browser.close()
print(f'PASS: manual updated, no em dashes or obsolete downloads; PDF with {len(reader.pages)} pages matches the app download.')

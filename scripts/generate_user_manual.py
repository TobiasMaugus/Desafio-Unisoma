"""Build the downloadable manual from MANUAL_USUARIO.md, without extra packages."""
from pathlib import Path
import re
import sys

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
from pdf_report import BLUE, MUTED, NAVY, TEAL, WHITE, WIDTH, _Document, _wrap


class ManualDocument(_Document):
    def __init__(self):
        super().__init__(adjusted=False)
        self.section = 'Como usar o aplicativo'

    def page(self, *args, **kwargs):
        self.commands = []
        self.pages.append(self.commands)
        self.rect(0, 0, 595, 66, NAVY)
        self.rect(0, 66, 595, 4, TEAL)
        self.text('ABC Aprendiz', 42, 32, 18, True, WHITE)
        self.text('MANUAL DO USUÁRIO', 42, 50, 9, color=WHITE)
        self.text('Ensalamento de alunos', 42, 94, 16, True)
        self.text(self.section, 42, 114, 10, color=MUTED)
        self.y = 140

    def section_heading(self, title, level):
        self.section = title
        self.space(85)
        self.y += 12
        size = 14 if level == 2 else 11
        for line in _wrap(title, WIDTH, size):
            self.text(line, 42, self.y, size, True, BLUE if level == 2 else NAVY)
            self.y += size + 5
        self.y += 5


def build_manual(source):
    document = ManualDocument()
    document.page()
    paragraph = []

    def flush():
        if paragraph:
            document.paragraph(' '.join(paragraph), size=10)
            document.y += 4
            paragraph.clear()

    for line in source.splitlines():
        line = line.strip()
        if line.startswith('# '):
            continue
        if not line:
            flush()
        elif line.startswith('##'):
            flush()
            document.section_heading(line.lstrip('#').strip(), len(line) - len(line.lstrip('#')))
        elif line.startswith('- ') or re.match(r'\d+\. ', line):
            flush()
            document.paragraph(line.replace('**', '').replace('- ', '• ', 1) if line.startswith('- ') else line.replace('**', ''), size=10)
            document.y += 2
        else:
            paragraph.append(line.replace('**', ''))
    flush()
    return document.finish()


if __name__ == '__main__':
    source = (PROJECT / 'MANUAL_USUARIO.md').read_text(encoding='utf-8')
    if '\u2014' in source or '\u2013' in source:
        raise ValueError('Revise os travessões do manual antes de gerar o PDF.')
    output = PROJECT / 'MANUAL_USUARIO.pdf'
    output.write_bytes(build_manual(source))
    print(f'Manual atualizado: {output.name}')

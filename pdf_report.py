"""Printable distribution reports, using standard PDF fonts without dependencies."""
import unicodedata


NAVY = (0.10, 0.20, 0.31)
BLUE = (0.055, 0.463, 0.737)
TEAL = (0.18, 0.58, 0.53)
MUTED = (0.39, 0.45, 0.51)
PALE = (0.94, 0.96, 0.98)
WHITE = (1, 1, 1)
WIDTH = 511
BOTTOM = 780


def _text_width(value, size):
    """Conservative Helvetica widths for wrapping, including accented letters."""
    width = 0
    for char in unicodedata.normalize('NFKD', str(value)):
        if unicodedata.combining(char):
            continue
        if char in 'ilI.,:;!|\' ':
            width += .30
        elif char in 'mwMW@%':
            width += .95
        elif char.isupper():
            width += .74
        else:
            width += .59
    return width * size


def _wrap(value, width, size):
    lines = []
    for paragraph in str(value).splitlines() or ['']:
        line = ''
        for word in paragraph.split():
            if line and _text_width(line + ' ' + word, size) > width:
                lines.append(line)
                line = ''
            while _text_width(word, size) > width:
                end = 1
                while end < len(word) and _text_width(word[:end + 1], size) <= width:
                    end += 1
                lines.append(word[:end])
                word = word[end:]
            line = (line + ' ' + word).strip()
        lines.append(line)
    return lines


class _Document:
    def __init__(self, adjusted):
        self.pages = []
        self.mode = 'Distribuição ajustada manualmente' if adjusted else 'Solução otimizada pelo modelo'

    def rect(self, x, top, width, height, color):
        rgb = ' '.join(str(item) for item in color)
        self.commands.append(f'q {rgb} rg {x} {842 - top - height} {width} {height} re f Q')

    def text(self, value, x, top, size=10, bold=False, color=NAVY):
        # A text object must contain a single line, with PDF delimiters escaped.
        value = str(value).replace('\r', ' ').replace('\n', ' ')
        value = value.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')
        rgb = ' '.join(str(item) for item in color)
        font = 'F2' if bold else 'F1'
        self.commands.append(
            f'BT /{font} {size} Tf {rgb} rg 1 0 0 1 {x} {842 - top} Tm ({value}) Tj ET')

    def page(self, title, subtitle='', continued=False):
        self.commands = []
        self.pages.append(self.commands)
        self.rect(0, 0, 595, 66, NAVY)
        self.rect(0, 66, 595, 4, TEAL)
        self.text('ABC Aprendiz', 42, 32, 18, True, WHITE)
        self.text('ENSALAMENTO · RELATÓRIO DE DISTRIBUIÇÃO', 42, 50, 8, color=WHITE)
        self.text(self.mode, 42, 92, 9, color=MUTED)
        top = 123
        for line in _wrap(title, WIDTH, 22):
            self.text(line, 42, top, 22, True)
            top += 27
        if continued:
            subtitle = 'Continuação · ' + subtitle
        for line in _wrap(subtitle, WIDTH, 10):
            self.text(line, 42, top, 10, color=MUTED)
            top += 14
        self.y = top + 14

    def space(self, height, title='Resumo da distribuição', subtitle='Continuação do relatório'):
        if self.y + height > BOTTOM:
            self.page(title, subtitle)

    def paragraph(self, value, size=10, color=NAVY):
        for line in _wrap(value, WIDTH, size):
            self.space(size + 6)
            self.text(line, 42, self.y, size, color=color)
            self.y += size + 5

    def heading(self, title, reserve=60):
        self.space(reserve)
        self.y += 22
        self.text(title, 42, self.y, 13, True)
        self.y += 20

    def cards(self, entries):
        self.space(78)
        gap = 10
        width = (WIDTH - gap * (len(entries) - 1)) / len(entries)
        for index, (label, value) in enumerate(entries):
            x = 42 + index * (width + gap)
            self.rect(x, self.y, width, 66, PALE)
            self.text(label, x + 12, self.y + 20, 8, color=MUTED)
            self.text(value, x + 12, self.y + 48, 22, True)
        self.y += 84

    def table_header(self, labels, widths):
        self.rect(42, self.y, WIDTH, 25, NAVY)
        x = 42
        for label, width in zip(labels, widths):
            self.text(label, x + 9, self.y + 16, 9, True, WHITE)
            x += width
        self.y += 25

    def table_row(self, values, widths, index, wrapped=None):
        if wrapped is None:
            wrapped = [_wrap(value, width - 18, 9) for value, width in zip(values, widths)]
        height = max(26, max(map(len, wrapped)) * 13 + 12)
        if index % 2 == 0:
            self.rect(42, self.y, WIDTH, height, PALE)
        x = 42
        for lines, width in zip(wrapped, widths):
            for offset, line in enumerate(lines):
                self.text(line, x + 9, self.y + 16 + offset * 13, 9)
            x += width
        self.rect(42, self.y + height - .5, WIDTH, .5, (.85, .89, .92))
        self.y += height
        return height

    def finish(self):
        for index, commands in enumerate(self.pages):
            # Draw each footer onto its own page after the final count is known.
            self.commands = commands
            self.rect(42, 800, WIDTH, .6, (.80, .85, .89))
            self.text('ABC Aprendiz · Organização das turmas', 42, 818, 8, color=MUTED)
            self.text(f'Página {index + 1} de {len(self.pages)}', 463, 818, 8, color=MUTED)

        objects = []

        def add(value):
            objects.append(value)
            return len(objects)

        catalog = add(b'')
        pages = add(b'')
        regular = add(b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>')
        bold = add(b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>')
        page_ids = []
        for commands in self.pages:
            stream = '\n'.join(commands).encode('cp1252', errors='replace')
            content = add(b'<< /Length ' + str(len(stream)).encode() + b' >>\nstream\n' + stream + b'\nendstream')
            page_ids.append(add(
                f'<< /Type /Page /Parent {pages} 0 R /MediaBox [0 0 595 842] '
                f'/Resources << /Font << /F1 {regular} 0 R /F2 {bold} 0 R >> >> '
                f'/Contents {content} 0 R >>'.encode()))
        objects[catalog - 1] = f'<< /Type /Catalog /Pages {pages} 0 R >>'.encode()
        objects[pages - 1] = (
            f'<< /Type /Pages /Kids [{" ".join(f"{page} 0 R" for page in page_ids)}] '
            f'/Count {len(page_ids)} >>'.encode())
        pdf = bytearray(b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n')
        offsets = [0]
        for index, obj in enumerate(objects, 1):
            offsets.append(len(pdf))
            pdf.extend(f'{index} 0 obj\n'.encode() + obj + b'\nendobj\n')
        xref = len(pdf)
        pdf.extend(f'xref\n0 {len(objects) + 1}\n0000000000 65535 f \n'.encode())
        pdf.extend(b''.join(f'{offset:010d} 00000 n \n'.encode() for offset in offsets[1:]))
        pdf.extend(f'trailer\n<< /Size {len(objects) + 1} /Root {catalog} 0 R >>\nstartxref\n{xref}\n%%EOF'.encode())
        return bytes(pdf)


def build_distribution_report(title, stats, capacity_by_room, issues, students_total,
                              adjusted=False, roster=()):
    """Create a summary followed by paginated class rosters from the same result."""
    roster = list(roster)
    document = _Document(adjusted)
    document.page(title, 'Visão geral da distribuição dos alunos')
    document.cards([
        ('ALUNOS ALOCADOS', str(students_total)),
        ('TURMAS OCUPADAS', str(sum(count > 0 for _, count, _, _ in stats))),
        ('VAGAS TOTAIS', str(sum(capacity_by_room.get(room, 0) for room, *_ in stats))),
    ])
    document.heading('Ocupação e composição por turma')
    labels = ['Turma', 'Alunos', 'Vagas', 'Meninas', 'Meninos', 'Ocupação']
    widths = [126, 61, 61, 80, 80, 103]
    document.table_header(labels, widths)
    by_period = {'Manhã': [], 'Tarde': []}
    for index, (room, count, girls, boys) in enumerate(stats):
        if document.y + 32 > BOTTOM:
            document.page('Resumo da distribuição', 'Ocupação e composição por turma', continued=True)
            document.table_header(labels, widths)
        capacity = capacity_by_room.get(room, 0)
        ratio = 100 * girls / count if count else 0
        document.table_row([
            room, count, capacity, girls, boys,
            f'{100 * count / capacity:.0f}%' if capacity else '—',
        ], widths, index)
        period = 'Manhã' if str(room).upper().startswith('MANHÃ') else 'Tarde'
        by_period[period].append((count, ratio))

    document.heading('Equilíbrio entre as turmas', reserve=90)
    for period, groups in by_period.items():
        if groups:
            sizes, ratios = zip(*groups)
            document.paragraph(f'{period}: diferença de {max(sizes) - min(sizes)} aluno(s) entre turmas; '
                               f'variação de meninas de {max(ratios) - min(ratios):.1f} pontos percentuais.')

    document.heading('Avisos e inconsistências', reserve=65)
    if not issues:
        document.paragraph('Nenhuma inconsistência encontrada.', color=TEAL)
    for level, message in issues:
        document.paragraph(f'{level}: {message}')
        document.y += 7

    def student_list(members, title, subtitle):
        labels, widths = ['Nº', 'Nome do aluno', 'Sexo', 'Disponibilidade'], [38, 303, 55, 115]
        document.heading('Lista de alunos', reserve=65)
        document.table_header(labels, widths)
        if not members:
            document.y += 18
            document.paragraph('Nenhum aluno alocado nesta turma.', color=MUTED)
        for index, (name, sex, period, _) in enumerate(members, 1):
            wrapped = [_wrap(value, width - 18, 9)
                       for value, width in zip([index, name, sex, period], widths)]
            height = max(26, max(map(len, wrapped)) * 13 + 12)
            if document.y + height > BOTTOM:
                document.page(title, subtitle, continued=True)
                document.table_header(labels, widths)
            # Exceptionally long rows continue across pages, with the same columns.
            offset = 0
            while offset < max(map(len, wrapped)):
                if document.y + 26 > BOTTOM:
                    document.page(title, subtitle, continued=True)
                    document.table_header(labels, widths)
                limit = int((BOTTOM - document.y - 12) // 13)
                fragment = [lines[offset:offset + limit] for lines in wrapped]
                document.table_row([], widths, index - 1, wrapped=fragment)
                offset += limit

    for room, count, girls, boys in stats:
        capacity = capacity_by_room.get(room, 0)
        period = 'Manhã' if str(room).upper().startswith('MANHÃ') else 'Tarde'
        document.page(str(room), f'Turno da {period.lower()} · Detalhamento da turma')
        document.cards([
            ('ALUNOS', str(count)), ('CAPACIDADE', str(capacity)),
            ('MENINAS', str(girls)), ('MENINOS', str(boys)),
        ])
        ratio = girls / count if count else 0
        document.text('Composição da turma', 42, document.y, 11, True)
        document.y += 14
        document.rect(42, document.y, WIDTH, 12, BLUE if count else PALE)
        if count:
            document.rect(42, document.y, WIDTH * ratio, 12, TEAL)
        document.y += 29
        document.paragraph(f'Meninas: {100 * ratio:.1f}%  ·  '
                           f'Meninos: {100 * boys / count if count else 0:.1f}%  ·  '
                           f'Vagas restantes: {max(0, capacity - count)}', size=9, color=MUTED)
        members = [row for row in roster if row[3] == room]
        if roster or count == 0:
            student_list(members, str(room), f'Turno da {period.lower()} · Lista de alunos')
        else:
            document.paragraph('Lista nominal não fornecida para este relatório.', color=MUTED)

    unassigned = [row for row in roster if not row[3]]
    if unassigned:
        document.page('Alunos sem turma', 'Registros que precisam de alocação após os ajustes manuais')
        document.paragraph(f'{len(unassigned)} aluno(s) sem turma definida.')
        student_list(unassigned, 'Alunos sem turma', 'Lista de alunos que precisam de alocação')
    return document.finish()

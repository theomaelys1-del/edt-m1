"""Convertit un emploi du temps Excel du Groupe 2 en calendrier abonné."""
import argparse
import json
import re
import sys
import unicodedata
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

ADDRESS = '2 Rue du Facteur Cheval, 91000 Évry-Courcouronnes'
PARIS = ZoneInfo('Europe/Paris')
HEADERS = ('dates', 'horaires', 'disciplines', 'intervenants', 'salle')
MONTHS = {'jan': 1, 'fev': 2, 'mar': 3, 'avr': 4, 'mai': 5, 'juin': 6,
          'juil': 7, 'aou': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
DAYS = ['lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi', 'dimanche']


def norm(value):
    return ''.join(c for c in unicodedata.normalize('NFD', str(value or '').lower())
                   if unicodedata.category(c) != 'Mn').strip()


def clean(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def parse_date(value):
    if isinstance(value, (datetime, date)):
        result = value.date() if isinstance(value, datetime) else value
    else:
        s = norm(value)
        match = re.search(r'(\d{1,2})\s+([a-z]+)(?:\.?\s+(\d{4}))?', s)
        if not match:
            raise ValueError(f'Date non reconnue : {value!r}')
        day, month_text, year = match.groups()
        month = next((v for k, v in MONTHS.items() if month_text.startswith(k)), None)
        if not month:
            raise ValueError(f'Mois non reconnu : {value!r}')
        expected_year = 2026 if month >= 9 else 2027
        if year and int(year) != expected_year:
            raise ValueError(f'Année hors 2026–2027 : {value!r}')
        result = date(expected_year, month, int(day))
        weekday = next((i for i, name in enumerate(DAYS) if re.search(r'\b' + name + r'\b', s)), None)
        if weekday is not None and result.weekday() != weekday:
            raise ValueError(f'Jour incohérent : {value!r} correspond à {DAYS[result.weekday()]} {result}')
    if result.year != (2026 if result.month >= 9 else 2027):
        raise ValueError(f'Date hors année universitaire : {result}')
    return result


def parse_hours(value, day):
    s = clean(value).lower().replace('–', '-').replace('—', '-')
    s = re.sub(r'(h\d{0,2})\s+(?=\d{1,2}h)', r'\1-', s).replace(' ', '')
    match = re.fullmatch(r'(\d{1,2})h(\d{2})?(?:-(\d{1,2})h(\d{2})?)?', s)
    if not match:
        raise ValueError(f'Horaires non reconnus : {value!r}')
    h, m, eh, em = match.groups()
    start = datetime.combine(day, time(int(h), int(m or 0)), PARIS)
    end = datetime.combine(day, time(int(eh), int(em or 0)), PARIS) if eh else start + timedelta(minutes=90)
    if end <= start:
        raise ValueError(f'Fin avant début : {value!r}')
    return start, end


def read_courses(path, sheet_name=None, confirmed_group=False):
    wb = load_workbook(path, data_only=True)
    candidates = []
    for ws in wb.worksheets:
        if sheet_name and ws.title != sheet_name:
            continue
        rows = list(ws.values)
        for index, row in enumerate(rows):
            mapping = {norm(v): i for i, v in enumerate(row) if norm(v) in HEADERS}
            if len(mapping) == 5:
                context = norm(ws.title + ' ' + ' '.join(clean(v) for r in rows[max(0, index-8):index+1] for v in r))
                groups = set(re.findall(r'groupe\s*([12])\b', context))
                if groups == {'2'} or (confirmed_group and sheet_name):
                    candidates.append((ws, rows, index, mapping))
    if len(candidates) != 1:
        raise ValueError('Groupe 2 ambigu ou absent. Inspecter le fichier ; préciser --sheet NOM --confirm-group-2 seulement après vérification visuelle du bloc sélectionné.')
    ws, rows, header, cols = candidates[0]
    # La colonne Enseignants est repérée dans le tableau à droite.
    teachers = {}
    right = max(cols.values()) + 1
    teacher_cols = {i for row in rows for i, v in enumerate(row)
                    if i >= right and norm(v) in {'enseignants', 'enseignant'}}
    full_names = set()
    for row in rows:
        for i in teacher_cols:
            name = clean(row[i])
            if name and len(name.split()) >= 2 and not re.match(r'^[A-Za-z]\.', name):
                full_names.add(name)
    for row in rows:
        values = [clean(v) for v in row[right:] if clean(v)]
        if len(values) == 2 and len(values[0]) <= 15 and len(values[1].split()) >= 2:
            key = norm(values[0]).rstrip('.')
            if key in teachers and teachers[key] != values[1]:
                raise ValueError(f'Correspondance enseignant ambiguë : {values[0]}')
            teachers[key] = values[1]
    def resolve_teacher(value):
        key = norm(value).rstrip('.')
        if key in teachers:
            return teachers[key]
        if any(norm(name) == key for name in full_names):
            return next(name for name in sorted(full_names) if norm(name) == key)
        match = re.fullmatch(r'([a-z])\.?\s+(.+)', key)
        if match:
            initial, surname = match.groups()
            matches = [name for name in sorted(full_names)
                       if norm(name).startswith(initial) and
                       norm(' '.join(name.split()[1:])).startswith(surname)]
            if len(matches) == 1:
                return matches[0]
            if len(matches) > 1:
                raise ValueError(f'Enseignant ambigu : {value} → {matches}')
        return value
    courses = []
    seen = set()
    for number, row in enumerate(rows[header+1:], header+2):
        if any(re.search(r'groupe\s*1\b', norm(v)) for v in row[:right]):
            break
        if norm(row[cols['dates']]) == 'dates':
            break
        raw = row[cols['dates']]
        if raw is None or not clean(raw):
            continue
        try:
            day = parse_date(raw)
            start, end = parse_hours(row[cols['horaires']], day)
            uid = f'{day.isoformat()}T{start:%H%M}-cours@edt-m1'
            if uid in seen:
                raise ValueError(f'Deux cours au même début : {uid}')
            seen.add(uid)
            teacher = clean(row[cols['intervenants']])
            teacher = resolve_teacher(teacher)
            courses.append(dict(uid=uid, start=start.isoformat(), end=end.isoformat(),
                                subject=clean(row[cols['disciplines']]) or 'A définir',
                                room=clean(row[cols['salle']]), teacher=teacher))
        except ValueError as error:
            raise ValueError(f'{ws.title}, ligne {number} : {error}') from error
    if not courses:
        raise ValueError('Aucun cours trouvé ; aucun fichier remplacé.')
    return sorted(courses, key=lambda c: c['start'])


def escape(value):
    return value.replace('\\', '\\\\').replace('\n', '\\n').replace(';', '\\;').replace(',', '\\,')


def fold(line):
    parts, current = [], ''
    for char in line:
        if len((current + char).encode('utf-8')) > 75:
            parts.append(current)
            current = ' '
        current += char
    return '\r\n'.join(parts + [current])


def utc(value):
    return value.astimezone(timezone.utc).strftime('%Y%m%dT%H%M%SZ')


def render(courses):
    stamp = utc(datetime.now(timezone.utc))
    lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//edt-m1//MGACL Groupe 2//FR',
             'CALSCALE:GREGORIAN', 'X-WR-CALNAME:EDT M1 MGACL - Groupe 2',
             'REFRESH-INTERVAL;VALUE=DURATION:PT1H', 'X-PUBLISHED-TTL:PT1H']
    trips = 0
    def event(uid, start, end, title, location):
        lines.extend(['BEGIN:VEVENT', 'UID:' + uid, 'DTSTAMP:' + stamp,
                      'DTSTART:' + utc(start), 'DTEND:' + utc(end),
                      'SUMMARY:' + escape(title), 'LOCATION:' + escape(location), 'END:VEVENT'])
    for c in courses:
        start, end = datetime.fromisoformat(c['start']), datetime.fromisoformat(c['end'])
        title = c['subject']
        if c['room']:
            title += ' — Salle ' + c['room']
        if c['teacher']:
            title += ' — ' + c['teacher']
        location = ('Salle ' + c['room'] + ' - ' if c['room'] else '') + ADDRESS
        event(c['uid'], start, end, title, location)
        if (start.hour, start.minute) in {(9, 0), (9, 30)}:
            trips += 1
            event(c['uid'].replace('-cours@', '-trajet@'), start-timedelta(hours=1), start,
                  '🚗 Trajet → Évry (≈1h) — ' + c['subject'], 'Domicile → ' + ADDRESS)
    lines.append('END:VCALENDAR')
    return ('\r\n'.join(fold(line) for line in lines) + '\r\n').encode('utf-8'), trips


def differences(old, new):
    before, after = {c['uid']: c for c in old}, {c['uid']: c for c in new}
    for uid in sorted(before.keys() & after.keys()):
        changes = [f'{field}: {before[uid][field]} → {after[uid][field]}'
                   for field in ('subject', 'room', 'teacher', 'end') if before[uid][field] != after[uid][field]]
        if changes:
            print(f'Modifié {uid} : ' + '; '.join(changes))
    removed = [before[k] for k in sorted(before.keys()-after.keys())]
    added = [after[k] for k in sorted(after.keys()-before.keys())]
    # Un déplacement est une hypothèse uniquement si la matière est unique des deux côtés.
    for c in removed[:]:
        matches = [n for n in added if n['subject'] == c['subject']]
        if len(matches) == 1 and sum(r['subject'] == c['subject'] for r in removed) == 1:
            n = matches[0]
            print(f'Déplacement probable (à vérifier) : {c["subject"]}, {c["start"]} → {n["start"]}; salle {c["room"]} → {n["room"]}; intervenant {c["teacher"]} → {n["teacher"]}')
            removed.remove(c)
            added.remove(n)
    for label, items in [('Supprimé', removed), ('Ajouté', added)]:
        for c in items:
            print(f'{label} : {c["start"]} — {c["subject"]} — {c["room"]} — {c["teacher"]}')
    if old == new:
        print('Aucune différence de cours.')


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('excel', type=Path)
    parser.add_argument('--sheet')
    parser.add_argument('--confirm-group-2', action='store_true')
    parser.add_argument('--output', type=Path, default=Path(__file__).parent / 'edt.ics')
    args = parser.parse_args()
    courses = read_courses(args.excel, args.sheet, args.confirm_group_2)
    data, trips = render(courses)
    state = args.output.with_suffix('.state.json')
    old = json.loads(state.read_text(encoding='utf-8')) if state.exists() else []
    differences(old, courses)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    if old != courses or not args.output.exists():
        temporary = args.output.with_suffix('.ics.tmp')
        temporary.write_bytes(data)
        temporary.replace(args.output)
        state.write_text(json.dumps(courses, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(f'Cours : {len(courses)} ; trajets : {trips} ; première date : {courses[0]["start"][:10]} ; dernière date : {courses[-1]["start"][:10]}')
    print('Cours sans salle :')
    for c in courses:
        if not c['room']:
            print(f'  {c["start"]} — {c["subject"]} — {c["teacher"]}')


if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError) as error:
        raise SystemExit(str(error))

"""Pilote de sessions mysql.exe réelles + rendu des captures d'écran (PNG)."""
import subprocess, threading, queue, time, re, os, json
from PIL import Image, ImageDraw, ImageFont

BASE = r'C:\Users\tomga\tp2_work\mysql'
BIN = BASE + r'\mysql-8.4.9-winx64\bin'
OUT = r'C:\Users\tomga\tp2_work\rapport\img'
os.makedirs(OUT, exist_ok=True)

FONT = ImageFont.truetype(r'C:\Windows\Fonts\consola.ttf', 17)
FONTB = ImageFont.truetype(r'C:\Windows\Fonts\consolab.ttf', 17)
FONT_T = ImageFont.truetype(r'C:\Windows\Fonts\segoeui.ttf', 16)
CW = FONT.getlength('M')
LH = 21

BG, GRID, HEAD = (24, 24, 28), (70, 70, 80), (40, 44, 56)
C_PROMPT, C_CMD, C_OUT = (90, 200, 110), (240, 240, 240), (200, 200, 205)
C_ERR, C_WAIT, C_NOTE, C_UNB = (255, 95, 95), (255, 200, 60), (100, 200, 255), (255, 160, 60)
SESS_COL = [(120, 190, 255), (255, 170, 90), (190, 140, 255), (120, 230, 180)]


# ---------------------------------------------------------------- serveur
def ping():
    return subprocess.run([BIN + r'\mysqladmin.exe', '-uroot', '-h', 'localhost', 'ping'],
                          capture_output=True).returncode == 0


def start_server():
    log = open(BASE + r'\server_restart.log', 'a')
    subprocess.Popen([BIN + r'\mysqld.exe', '--defaults-file=' + BASE + r'\my.ini', '--console'],
                     stdout=log, stderr=subprocess.STDOUT,
                     creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)
    for _ in range(60):
        if ping():
            return
        time.sleep(1)
    raise RuntimeError('serveur non démarré')


def shell(cmd_display, args, enc='cp850'):
    r = subprocess.run(args, capture_output=True)
    out = (r.stdout + r.stderr).decode(enc, errors='replace').replace('\r', '')
    return cmd_display, out.strip()


# ---------------------------------------------------------------- session
ECHO = re.compile(r'-{14}\n.*?\n-{14}\n\n?', re.S)


def clean(txt):
    txt = ECHO.sub('', txt)
    lines = []
    for l in txt.split('\n'):
        if 'Using a password on the command line' in l or l.strip() == 'Bye':
            continue
        lines.append(re.sub(r' at line \d+', '', l))
    return '\n'.join(lines).strip('\n')


class Session:
    def __init__(self, name, user='root', pwd='', db='tp2_mysql', label=None):
        self.name, self.user = name, user
        self.label = label or f'{name} ({user})'
        args = [BIN + r'\mysql.exe', '-h', 'localhost', '-u', user, '-t', '-vv', '-n', '--force',
                '--default-character-set=utf8mb4']
        if pwd:
            args.append('-p' + pwd)
        if db:
            args.append(db)
        self.p = subprocess.Popen(args, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                  stderr=subprocess.STDOUT, text=True, encoding='utf-8', bufsize=1)
        self.q = queue.Queue()
        self.blocked = False
        threading.Thread(target=self._read, daemon=True).start()
        time.sleep(0.4)
        self.drain()

    def _read(self):
        for line in self.p.stdout:
            self.q.put(line)

    def send(self, cmd):
        self.p.stdin.write(cmd.strip().rstrip(';') + ';\n')
        self.p.stdin.flush()

    def drain(self):
        s = ''
        while True:
            try:
                s += self.q.get_nowait()
            except queue.Empty:
                return s

    def close(self):
        try:
            self.p.stdin.write('exit\n'); self.p.stdin.flush()
        except Exception:
            pass
        try:
            self.p.wait(5)
        except Exception:
            self.p.kill()


# ---------------------------------------------------------------- scénario
class Scenario:
    def __init__(self, sessions, title):
        self.sessions = list(sessions)
        self.title = title
        self.rows = []      # chaque ligne : {session_name: [(texte, couleur), ...]}
        self.log = []       # journal texte (pour le rapport)
        self.n = 0

    def _collect(self, active, wait):
        t0, last, raw = time.time(), time.time(), {s.name: '' for s in self.sessions}
        while True:
            time.sleep(0.15)
            got = False
            for s in self.sessions:
                d = s.drain()
                if d:
                    raw[s.name] += d
                    got = True
            if got:
                last = time.time()
            if time.time() - t0 >= wait and time.time() - last >= 0.6:
                break
        return {k: clean(v) for k, v in raw.items()}

    def add_session(self, s):
        self.sessions.append(s)

    def run(self, s, cmd, wait=1.0, observe=(), obs_sql='SELECT * FROM T', row=None):
        """Exécute cmd dans la session s ; renvoie la sortie. observe = sessions qui font un SELECT ensuite."""
        if row is None:
            self.n += 1
            row = {'_n': self.n}
            self.rows.append(row)
        s.send(cmd)
        outs = self._collect(s, wait)
        cell = row.setdefault(s.name, [])
        cell.append(('mysql> ' + cmd.strip().rstrip(';') + ';', 'cmd'))
        mine = outs.pop(s.name)
        if mine.strip():
            cell += [(l, 'err' if l.startswith('ERROR') else 'out') for l in mine.split('\n')]
            s.blocked = False
        else:
            cell.append(('>>> EN ATTENTE (verrou détenu par une autre transaction)...', 'wait'))
            s.blocked = True
        self.log.append({'n': row['_n'], 'who': s.name, 'cmd': cmd, 'out': mine or '<<BLOQUÉ>>'})
        for name, o in outs.items():
            if o.strip():
                other = next(x for x in self.sessions if x.name == name)
                other.blocked = False
                c = row.setdefault(name, [])
                c.append(('>>> requête précédente DÉBLOQUÉE, résultat :', 'unb'))
                c += [(l, 'err' if l.startswith('ERROR') else 'out') for l in o.split('\n')]
                self.log.append({'n': row['_n'], 'who': name, 'cmd': '(débloquée)', 'out': o})
        for o in observe:
            if not o.blocked:
                self.run(o, obs_sql, wait=0.5, row=row)
        return mine

    def note(self, s_name, text, new_row=True):
        if new_row:
            self.n += 1
            row = {'_n': self.n}
            self.rows.append(row)
        else:
            row = self.rows[-1]
        row.setdefault(s_name, []).extend((l, 'note') for l in text.split('\n'))
        self.log.append({'n': row['_n'], 'who': s_name, 'cmd': 'NOTE', 'out': text})

    def snap(self, fname, title=None, cols=None):
        rows, self.rows = self.rows, []
        render_table(os.path.join(OUT, fname), title or self.title,
                     cols or [(s.name, s.label) for s in self.sessions], rows)
        return fname


COLORS = {'cmd': C_CMD, 'out': C_OUT, 'err': C_ERR, 'wait': C_WAIT, 'note': C_NOTE, 'unb': C_UNB}


def wrap(cell, maxc):
    res = []
    for t, k in cell:
        if len(t) <= maxc:
            res.append((t, k))
        else:
            first = True
            while t:
                res.append((t[:maxc] if first else '  ' + t[:maxc - 2], k))
                t = t[maxc:] if first else t[maxc - 2:]
                first = False
    return res


def render_table(path, title, cols, rows):
    used = [c for c in cols if any(c[0] in r for r in rows)]
    maxc = 110 if len(used) == 1 else (78 if len(used) == 2 else 60)
    for r in rows:
        for name, _ in used:
            if name in r:
                r[name] = wrap(r[name], maxc)
    widths = []
    for name, label in used:
        w = max([len(label) + 2] + [len(t) for r in rows for t, _ in r.get(name, [])] + [20])
        widths.append(int(w * CW) + 24)
    nw = int(4 * CW) + 16
    W = nw + sum(widths)
    heights = [max([1] + [len(r.get(n, [])) for n, _ in used]) * LH + 16 for r in rows]
    TB, HH = 36, 32
    H = TB + HH + sum(heights)
    img = Image.new('RGB', (W, H), BG)
    d = ImageDraw.Draw(img)
    # barre de titre type fenêtre
    d.rectangle([0, 0, W, TB], fill=(52, 52, 60))
    for i, c in enumerate([(255, 95, 86), (255, 189, 46), (39, 201, 63)]):
        d.ellipse([12 + i * 22, 11, 26 + i * 22, 25], fill=c)
    d.text((90, 7), title, font=FONT_T, fill=(230, 230, 230))
    # en-têtes
    y = TB
    d.rectangle([0, y, W, y + HH], fill=HEAD)
    d.text((8, y + 6), '#', font=FONTB, fill=(180, 180, 180))
    x = nw
    for i, (name, label) in enumerate(used):
        d.text((x + 10, y + 6), label, font=FONTB, fill=SESS_COL[cols.index((name, label)) % 4])
        x += widths[i]
    y += HH
    for r, h in zip(rows, heights):
        d.line([0, y, W, y], fill=GRID)
        d.text((8, y + 8), str(r['_n']), font=FONTB, fill=(150, 150, 150))
        x = nw
        for i, (name, _) in enumerate(used):
            yy = y + 8
            for t, k in r.get(name, []):
                if k == 'cmd' and t.startswith('mysql> '):
                    d.text((x + 10, yy), 'mysql>', font=FONTB, fill=C_PROMPT)
                    d.text((x + 10 + 7 * CW, yy), t[7:], font=FONTB, fill=C_CMD)
                else:
                    d.text((x + 10, yy), t, font=FONT, fill=COLORS[k])
                yy += LH
            x += widths[i]
        y += h
    x = nw
    d.line([nw - 1, TB + HH, nw - 1, H], fill=GRID)
    for w in widths[:-1]:
        x += w
        d.line([x, TB, x, H], fill=GRID)
    img.save(path)


def render_console(fname, title, blocks):
    """blocks = [(commande, sortie)] rendu façon PowerShell."""
    lines = []
    for cmd, out in blocks:
        lines.append(('PS C:\\Users\\tomga> ' + cmd, 'cmd'))
        lines += [(l, 'err' if 'ERROR' in l or 'Erreur' in l else 'out') for l in out.split('\n')]
        lines.append(('', 'out'))
    lines = wrap(lines, 120)
    W = int(max(60, max(len(t) for t, _ in lines)) * CW) + 30
    H = 36 + len(lines) * LH + 20
    img = Image.new('RGB', (W, H), (1, 36, 86))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, 36], fill=(52, 52, 60))
    d.text((12, 7), title, font=FONT_T, fill=(230, 230, 230))
    y = 46
    for t, k in lines:
        if k == 'cmd':
            d.text((12, y), t, font=FONTB, fill=(255, 255, 120))
        else:
            d.text((12, y), t, font=FONT, fill=(255, 120, 120) if k == 'err' else (238, 238, 238))
        y += LH
    img.save(os.path.join(OUT, fname))
    return fname

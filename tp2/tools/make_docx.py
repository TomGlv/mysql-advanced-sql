import json, re
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

L = json.load(open(r'C:\Users\tomga\tp2_work\rapport\log.json', encoding='utf-8'))
OUT = r'C:\Users\tomga\Downloads\Rapport_TP2_MySQL.docx'

doc = Document()
sec = doc.sections[0]
sec.left_margin = sec.right_margin = Cm(2)
sec.top_margin = sec.bottom_margin = Cm(1.8)
st = doc.styles['Normal']
st.font.name = 'Calibri'
st.font.size = Pt(11)


def shade(el, color):
    pPr = el.get_or_add_pPr() if hasattr(el, 'get_or_add_pPr') else el
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), color)
    pPr.append(shd)


def cell_shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), color)
    tcPr.append(shd)


def block(text, kind='sql', caption=None):
    """kind='sql' : code SQL (fond clair) ; kind='term' : sortie terminal (fond sombre)."""
    if caption:
        c = doc.add_paragraph()
        r = c.add_run(caption)
        r.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
        c.paragraph_format.space_after = Pt(0)
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t.rows[0].cells[0]
    cell_shade(cell, '1E1E1E' if kind == 'term' else 'F2F4F7')
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    for i, line in enumerate(text.rstrip('\n').split('\n')):
        if i:
            p.add_run().add_break()
        if kind == 'term' and line.startswith(('mysql>', '[')) and 'mysql>' in line:
            pre, cmd = line.split('mysql>', 1)
            for txt, col, b in ((pre, (0x7A, 0xB8, 0xFF), True), ('mysql>', (0x5A, 0xC8, 0x6E), True),
                                (cmd, (0xF0, 0xF0, 0xF0), True)):
                r = p.add_run(txt); r.bold = b
                r.font.name = 'Consolas'; r.font.size = Pt(8); r.font.color.rgb = RGBColor(*col)
            continue
        r = p.add_run(line)
        r.font.name = 'Consolas'; r.font.size = Pt(8)
        if kind == 'term':
            col = (0xFF, 0x6B, 0x6B) if 'ERROR' in line else (0xFF, 0xC8, 0x3C) if line.strip().startswith(('...', '-->', '>>>')) else (0xD8, 0xD8, 0xDC)
            r.font.color.rgb = RGBColor(*col)
        else:
            r.font.color.rgb = RGBColor(0x8A, 0x8A, 0x8A) if line.strip().startswith('--') else RGBColor(0x1F, 0x3A, 0x68)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def H(txt, lvl=1):
    doc.add_heading(txt, lvl)


def P(txt, bold_prefix=None):
    p = doc.add_paragraph()
    if bold_prefix:
        p.add_run(bold_prefix).bold = True
    # **gras** inline
    for i, part in enumerate(re.split(r'\*\*', txt)):
        r = p.add_run(part)
        r.bold = bool(i % 2)
    p.paragraph_format.space_after = Pt(4)
    return p


def bullet(txt):
    p = doc.add_paragraph(style='List Bullet')
    for i, part in enumerate(re.split(r'\*\*', txt)):
        p.add_run(part).bold = bool(i % 2)


def answer(txt):
    t = doc.add_table(rows=1, cols=1)
    c = t.rows[0].cells[0]
    cell_shade(c, 'E8F4EA')
    p = c.paragraphs[0]
    p.add_run('Réponse / observation : ').bold = True
    for i, part in enumerate(re.split(r'\*\*', txt)):
        p.add_run(part).bold = bool(i % 2)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def term(key, start=None, end=None, caption=None, skip_obs=False, tag=True):
    """Transforme une séquence du journal d'exécution réel en bloc 'terminal'."""
    lines = []
    for e in L[key]:
        if (start and e['n'] < start) or (end and e['n'] > end):
            continue
        pre = f"[{e['who']}] " if tag else ''
        if e['cmd'] == 'NOTE':
            lines += [l if l.startswith('>>>') else '    ' + l for l in e['out'].split('\n')]
        elif e['cmd'] == '(débloquée)':
            lines.append(f"--> [{e['who']}] requête précédente DÉBLOQUÉE :")
            lines += e['out'].split('\n')
        else:
            lines.append(f"{pre}mysql> {e['cmd'].strip().rstrip(';')};")
            if e['out'] == '<<BLOQUÉ>>':
                lines.append('... EN ATTENTE (verrou détenu par une autre transaction) ...')
            else:
                lines += e['out'].split('\n')
        lines.append('')
    block('\n'.join(lines), 'term', caption)


def rows_of(out):
    if out.strip() == 'Empty set' or 'Empty set' in out:
        return '∅ (vide)'
    rows = re.findall(r'^\|(.*)\|$', out, re.M)[1:]
    vals = ['(' + ','.join(x.strip() for x in r.split('|')) + ')' for r in rows]
    return ' '.join(vals).replace('NULL', 'null')


def grid(headers, data, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ''
        r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell_shade(c, '2F4F7F')
    for row in data:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ''
            r = cells[i].paragraphs[0].add_run(str(v))
            r.font.size = Pt(9)
            if i > 0 and any(k in str(v) for k in ('DÉBUT', 'FIN', 'bloqu', 'PANNE')):
                r.bold = True
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def table_from_log(key, actors, obs_cols, notes):
    """Construit le tableau du sujet à partir des SELECT d'observation réellement exécutés."""
    steps = {}
    order = []
    for e in L[key]:
        n = e['n']
        if n not in steps:
            steps[n] = {'ordre': None, 'who': None}; order.append(n)
        s = steps[n]
        if e['cmd'] == 'NOTE':
            s['ordre'] = (s['ordre'] + ' / ' if s['ordre'] else '') + e['out'].split('\n')[0].replace('>>> ', '')
        elif e['cmd'].strip().upper() == 'SELECT * FROM T':
            s[e['who']] = rows_of(e['out'])
        elif e['cmd'] == '(débloquée)':
            s['unb'] = e['who']
        elif s['ordre'] is None:
            s['ordre'] = e['cmd']; s['who'] = e['who']
            if e['out'] == '<<BLOQUÉ>>':
                s[e['who']] = 'bloqué (attente verrou)'
    data = []
    for n in order:
        s = steps[n]
        vals = [s.get(c, '') for c in obs_cols]
        if s['ordre'] is None:      # simple SELECT de vérification
            prev = data[-1] if data else None
            if prev and ('EXIT' in prev[0] or prev[0].startswith('Après')):
                for i, v in enumerate(vals):
                    prev[i + 1] = prev[i + 1] or v
            else:
                data.append(['Après redémarrage + reconnexion : SELECT * FROM T'] + vals + [notes.get(n, '')])
            continue
        if s['ordre'].startswith('PANNE'):
            s['ordre'] = 'PANNE : arrêt brutal du service MySQL, puis redémarrage'
        lab = s['ordre'] if not s['who'] or len(actors) == 1 else f"{s['who']} : {s['ordre']}"
        if len(actors) > 1 and s['ordre'].startswith('SET autocommit'):
            if not any(vals):
                continue
            lab = 'État initial (p1 et p2 : SET autocommit = 0)'
        if 'unb' in s:
            lab += f"   [→ débloque {s['unb']}]"
        data.append([lab] + vals + [notes.get(n, '')])
    return data


# =================================================================== PAGE DE TITRE
for _ in range(5):
    doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Rapport de TP 2 — MySQL'); r.bold = True; r.font.size = Pt(28); r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x68)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Transactions, reprise sur panne, niveaux d’isolation,\naccès concurrents (verrous) et droits d’accès'); r.font.size = Pt(15)
doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run('B3 — TD/TP 2\nSGBD : MySQL Server 8.4.9 (moteur InnoDB) — base tp2_mysql\nScript de départ : « script tp2.txt »').font.size = Pt(12)
for _ in range(3):
    doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Tom GALVAN'); r.bold = True; r.font.size = Pt(16)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run('ECE — École d’ingénieurs').font.size = Pt(13)
doc.core_properties.author = 'Tom Galvan'
doc.add_page_break()

# =================================================================== SOMMAIRE
H('Sommaire', 1)
for s in ['0. Environnement et mise en place de la base', '1. COMMIT et ROLLBACK', '2. Transactions et reprise sur panne (client)',
          '3. Niveaux d’isolation des transactions', '4. Accès concurrents — même utilisateur, deux sessions',
          '5. Droits d’accès (U1 / U2)', '6. Transactions et reprise sur panne — tableaux p1 / p2',
          '7. Analyse de programme — tableaux p1 / p2 / p3', '8. Conclusion générale']:
    bullet(s)
doc.add_page_break()

# =================================================================== 0
H('0. Environnement et mise en place de la base', 1)
P('Le sujet fait référence à une table **PRODUITS** (colonnes pno, prix) et à une base **COMMANDES**. Le script fourni '
  '(« script tp2.txt ») crée la base **tp2_mysql** avec les tables departements, employes, clients, projets et '
  'affectations_projets. Les exercices ont donc été réalisés sur cette base avec la correspondance suivante :')
grid(['Sujet', 'Équivalent utilisé (script tp2)'],
     [['Base ECOLE / COMMANDES', 'Base tp2_mysql'],
      ['Table PRODUITS / commandes', 'Table projets'],
      ['Colonne pno', 'Colonne id'],
      ['Colonne prix / quantité', 'Colonne budget'],
      ['Produit pno = 203', 'Projet id = 3 (« Application web »)'],
      ['Table T(A, B)', 'Table T(A INTEGER, B INTEGER) créée dans tp2_mysql'],
      ['Icône « autocommit » de Workbench', 'Instruction SET autocommit = 0;']], [6, 10])
P('Chaque « connexion » / « onglet de connexion Workbench » du sujet correspond à une session MySQL distincte '
  '(p1, p2, p3, S1, S2, U1, U2…). Dans les blocs de sortie ci‑dessous, le nom de la session est indiqué entre crochets '
  'avant le prompt, par exemple « [p2] mysql> ». Les sorties reproduites sont celles obtenues lors de l’exécution '
  'des commandes sur MySQL 8.4.9.')

H('0.1 Création de la base avec le script', 2)
P('Le script est exécuté une fois (dans Workbench : File › Open SQL Script, puis l’éclair « Execute »), ou en ligne de commande :')
block('mysql -u root -p < "script tp2.txt"', 'sql', 'Commande')
P('Il crée la base, supprime les anciennes tables, crée les 5 tables avec leurs clés primaires/étrangères, insère les '
  'données puis affiche le contenu des tables. Extrait du résultat :')
block('''mysql> SELECT * FROM projets;
+----+-----------------------+----------+-----------+
| id | nom                   | budget   | client_id |
+----+-----------------------+----------+-----------+
|  1 | Audit de securite     | 25000.00 |         1 |
|  2 | Infrastructure reseau | 12000.00 |         1 |
|  3 | Application web       | 18000.00 |         2 |
|  4 | Migration de donnees  |  8000.00 |         3 |
|  5 | SOC et supervision    | 42000.00 |         5 |
+----+-----------------------+----------+-----------+
5 rows in set''', 'term', 'Sortie — table projets après le script')

H('0.2 Vérification du moteur InnoDB', 2)
P('Les transactions ne fonctionnent qu’avec un moteur transactionnel : on vérifie que toutes les tables sont en InnoDB.')
block("SELECT TABLE_NAME, ENGINE\nFROM information_schema.TABLES\nWHERE TABLE_SCHEMA = 'tp2_mysql';", 'sql', 'SQL exécuté')
term('setup', 1, 3, 'Sortie', tag=False)
answer('toutes les tables sont en **InnoDB** (moteur par défaut de MySQL 8), les transactions sont donc supportées. '
       'Par défaut **autocommit = ON** : chaque requête est validée immédiatement, il faut le désactiver pour le TP '
       '(icône « Toggle autocommit » dans Workbench, ou SET autocommit = 0;). Le moteur d’une table peut être changé avec '
       'ALTER TABLE nom ENGINE = InnoDB;')
H('0.3 Compte user1 (utilisé pour les parties 6 et 7)', 2)
block("CREATE USER 'user1'@'localhost' IDENTIFIED BY 'user1';\nGRANT ALL PRIVILEGES ON tp2_mysql.* TO 'user1'@'localhost';", 'sql', 'SQL exécuté (session root)')
doc.add_page_break()

# =================================================================== 1
H('1. COMMIT et ROLLBACK', 1)
P('Une seule connexion est utilisée. On désactive l’autocommit, on démarre une transaction, on insère un nouveau projet, '
  'on vérifie qu’il est visible, puis on annule (ROLLBACK). On recommence ensuite avec COMMIT.')
H('1.a Expérience avec ROLLBACK', 2)
block('''SET autocommit = 0;
START TRANSACTION;
INSERT INTO projets (nom, budget, client_id) VALUES ('Projet test ROLLBACK', 15000, 4);
SELECT id, nom, budget FROM projets;   -- le projet est visible
ROLLBACK;
SELECT id, nom, budget FROM projets;   -- le projet a disparu''', 'sql', 'SQL exécuté')
term('q1', 1, 6, 'Sortie', tag=False)
answer('après l’INSERT, le nouveau projet (id 6) est **visible par la transaction en cours**. Après le **ROLLBACK**, '
       'le SELECT ne le renvoie plus : toutes les modifications faites depuis le début de la transaction sont **annulées**. '
       'Remarque : l’identifiant AUTO_INCREMENT consommé (6) n’est pas récupéré.')
H('1.b Expérience avec COMMIT', 2)
block('''START TRANSACTION;
INSERT INTO projets (nom, budget, client_id) VALUES ('Projet test COMMIT', 15000, 4);
SELECT id, nom, budget FROM projets;
COMMIT;
SELECT id, nom, budget FROM projets;''', 'sql', 'SQL exécuté')
term('q1', 7, 11, 'Sortie', tag=False)
answer('après le **COMMIT**, le projet (id 7) est toujours présent : la transaction est **validée**, la modification '
       'devient **définitive (durable)** et visible par toutes les autres connexions. Un ROLLBACK ultérieur ne peut plus '
       'l’annuler. COMMIT et ROLLBACK terminent tous deux la transaction courante.')
doc.add_page_break()

# =================================================================== 2
H('2. Transactions et reprise sur panne (côté client)', 1)
H('2.a Fermeture normale de la connexion sans COMMIT', 2)
block('''-- Connexion n°1
SET autocommit = 0;
START TRANSACTION;
INSERT INTO projets (nom, budget, client_id) VALUES ('Projet avant fermeture', 9000, 4);
SELECT id, nom, budget FROM projets;
-- fermeture de l'onglet de connexion Workbench (EXIT) SANS COMMIT
-- Nouvelle connexion :
SELECT id, nom, budget FROM projets;''', 'sql', 'SQL exécuté')
term('q2a', caption='Sortie (a = connexion n°1, b = nouvelle connexion)')
answer('le projet « Projet avant fermeture » **n’existe plus** après reconnexion. À la fermeture d’une session, MySQL '
       'effectue un **ROLLBACK implicite** de la transaction non validée. (Workbench peut aussi demander s’il faut valider '
       'les transactions en cours à la fermeture de l’onglet : si on répond non, c’est un rollback.)')
H('2.b Arrêt brutal du client (gestionnaire des tâches)', 2)
P('On recommence, mais au lieu de fermer proprement, on tue le processus client (Ctrl+Alt+Suppr › Gestionnaire des '
  'tâches › MySQL Workbench › Fin de tâche ; équivalent en ligne de commande : taskkill /F).')
block('''SET autocommit = 0;
START TRANSACTION;
INSERT INTO projets (nom, budget, client_id) VALUES ('Projet avant crash client', 9500, 4);
SELECT id, nom, budget FROM projets;
-- >>> fin de tâche du client (taskkill /F /PID <pid>)
-- Nouvelle connexion :
SELECT id, nom, budget FROM projets;''', 'sql', 'SQL exécuté')
term('q2b', caption='Sortie')
answer('même résultat : le projet inséré **n’est pas présent**. Le serveur détecte la perte de la connexion et annule '
       '(ROLLBACK) la transaction orpheline.')
P('**Conclusion :** une transaction non validée par COMMIT n’est jamais rendue durable, que la connexion soit fermée '
  'proprement ou que le client plante. C’est la propriété **d’atomicité** (tout ou rien) des transactions ACID : la base '
  'reste dans le dernier état cohérent validé. Seul un COMMIT rend les modifications définitives (**durabilité**, '
  'garantie par le journal redo d’InnoDB).')
doc.add_page_break()

# =================================================================== 3
H('3. Niveaux d’isolation des transactions', 1)
H('3.1 Niveau d’isolation par défaut', 2)
block("SHOW VARIABLES LIKE '%isolation%';", 'sql', 'SQL exécuté')
term('q3_rr', 1, 1, 'Sortie', tag=False)
answer('le niveau par défaut de MySQL/InnoDB est **REPEATABLE READ** (variable transaction_isolation).')

H('3.2 Expérience : visibilité des modifications (REPEATABLE READ)', 2)
P('Deux connexions sont ouvertes (deux onglets principaux dans Workbench) : T1 = connexion 1 (p1), T2 = connexion 2 (p2), '
  'autocommit désactivé dans les deux. T1 insère un projet ; T2 lit la table.')
block('''-- p1 et p2 :
SET autocommit = 0;
-- p1 :
INSERT INTO projets (nom, budget, client_id) VALUES ('Projet isolation', 30000, 4);
SELECT id, nom, budget FROM projets;
-- p2 :
SELECT id, nom, budget FROM projets;
-- p1 :
COMMIT;
-- p2 :
SELECT id, nom, budget FROM projets;
COMMIT;
SELECT id, nom, budget FROM projets;''', 'sql', 'SQL exécuté')
term('q3_rr', 2, 10, 'Sortie')
answer('tant que T1 n’a pas validé, son insertion n’est **visible que par T1** : T2 ne la voit pas (pas de lecture sale). '
       'Point remarquable en **REPEATABLE READ** : même **après le COMMIT de T1**, T2 ne voit toujours pas le projet, car '
       'T2 lit un **instantané (snapshot)** pris lors de sa première lecture ; les lectures de T2 restent identiques '
       '(lecture répétable). T2 ne voit le nouveau projet qu’après avoir terminé sa propre transaction (COMMIT) et en '
       'avoir commencé une nouvelle.')
H('3.3 Même expérience avec la suppression du projet', 2)
block('''-- p1 :
DELETE FROM projets WHERE nom = 'Projet isolation';
SELECT id, nom, budget FROM projets;
-- p2 :
SELECT id, nom, budget FROM projets;
-- p1 : COMMIT;   p2 : COMMIT;
SELECT id, nom, budget FROM projets;   -- (p2)''', 'sql', 'SQL exécuté')
term('q3_rr', 11, 16, 'Sortie')
answer('symétriquement, la suppression faite par T1 n’est pas visible par T2 : T2 voit toujours le projet tant que T1 '
       'n’a pas validé (et tant que T2 n’a pas commencé une nouvelle transaction). Après les deux COMMIT, le projet a '
       'disparu pour tout le monde.')

H('3.4 Niveau READ UNCOMMITTED', 2)
P('**Quelle transaction faut‑il modifier ?** Le niveau d’isolation détermine ce qu’une transaction peut **lire**. '
  'Il faut donc modifier la transaction **lectrice, T2** (connexion p2). Changer le niveau de T1 (qui écrit) ne '
  'changerait rien à ce que voit T2. Le niveau doit être fixé avant le début de la transaction (d’où le COMMIT préalable).')
block('''-- p2 (lecteur) :
SET SESSION TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
SELECT @@transaction_isolation;
-- p1 :
INSERT INTO projets (nom, budget, client_id) VALUES ('Projet non valide', 7777, 4);
-- p2 :
SELECT id, nom, budget FROM projets;      -- voit la ligne NON validée
-- p1 :
ROLLBACK;
-- p2 :
SELECT id, nom, budget FROM projets;      -- la ligne a disparu
COMMIT;''', 'sql', 'SQL exécuté')
term('q3_ru', caption='Sortie')
answer('en READ UNCOMMITTED, T2 voit le « Projet non valide » **alors que T1 ne l’a pas validé** : c’est une '
       '**lecture sale (dirty read)**. Après le ROLLBACK de T1, la ligne disparaît : T2 a donc lu une donnée qui '
       '**n’a jamais existé** officiellement. T1, elle, n’est pas affectée. **Conclusion :** ce niveau est le plus '
       'performant (aucun verrou en lecture) mais le moins sûr ; il est à éviter dès que la cohérence des données compte.')

H('3.5 Niveau SERIALIZABLE', 2)
block('''-- p2 :
SET SESSION TRANSACTION ISOLATION LEVEL SERIALIZABLE;
-- p1 :
INSERT INTO projets (nom, budget, client_id) VALUES ('Projet serialisable', 11111, 4);
-- p2 :
SELECT id, nom, budget FROM projets;     -- BLOQUÉ
-- p1 :
COMMIT;                                  -- p2 est débloquée
-- p2 :
COMMIT;''', 'sql', 'SQL exécuté — cas 1 : T2 veut lire ce que T1 modifie')
term('q3_ser', 1, 6, 'Sortie')
block('''-- p2 :
SELECT id, nom, budget FROM projets;     -- pose des verrous partagés
-- p1 :
UPDATE projets SET budget = 99999 WHERE nom = 'Projet serialisable';   -- BLOQUÉ
-- p2 :
COMMIT;                                  -- p1 est débloquée
-- p1 :
COMMIT;''', 'sql', 'SQL exécuté — cas 2 : T1 veut modifier ce que T2 a lu')
term('q3_ser', 7, 10, 'Sortie')
answer('sur la deuxième transaction, le **SELECT reste bloqué** (en attente) : en SERIALIZABLE, InnoDB transforme '
       'chaque SELECT en SELECT … FOR SHARE (verrou partagé). Comme T1 détient un verrou exclusif sur la ligne insérée, '
       'T2 doit **attendre la fin de T1** (COMMIT ou ROLLBACK). Dès le COMMIT de T1, T2 est débloquée et voit la ligne. '
       'Inversement (cas 2), les lignes lues par T2 ne peuvent pas être modifiées par T1 tant que T2 n’est pas terminée. '
       'Si l’attente dépasse innodb_lock_wait_timeout (50 s), on obtient l’erreur 1205 « Lock wait timeout exceeded ». '
       '**Conclusion :** SERIALIZABLE garantit que l’exécution concurrente équivaut à une exécution en série (pas de '
       'lecture sale, non répétable ni fantôme), au prix de blocages et d’une baisse de la concurrence.')
grid(['Niveau', 'Lecture sale', 'Lecture non répétable', 'Lecture fantôme'],
     [['READ UNCOMMITTED', 'possible', 'possible', 'possible'],
      ['READ COMMITTED', 'non', 'possible', 'possible'],
      ['REPEATABLE READ (défaut MySQL)', 'non', 'non', 'non en lecture simple (snapshot + verrous next‑key)'],
      ['SERIALIZABLE', 'non', 'non', 'non']], [5.5, 3, 4, 4.5])
doc.add_page_break()

# =================================================================== 4
H('4. Accès concurrents — même utilisateur, deux sessions', 1)
P('Le même utilisateur (root) ouvre deux sessions S1 et S2, autocommit désactivé. Le produit pno = 203 du sujet '
  'correspond au projet id = 3 et la colonne prix à budget.')
H('4.1 Deux UPDATE sur la même ligne', 2)
block('''-- S1 et S2 : SET autocommit = 0;
-- S1 :
UPDATE projets SET budget = 3000 WHERE id = 3;
-- S2 :
UPDATE projets SET budget = 3500 WHERE id = 3;    -- BLOQUÉ
-- S1 :
SELECT budget FROM projets WHERE id = 3;
COMMIT;                                           -- débloque S2
-- S2 :
SELECT budget FROM projets WHERE id = 3;
-- S1 :
SELECT budget FROM projets WHERE id = 3;
-- S2 :
COMMIT;
-- S1 :
SELECT budget FROM projets WHERE id = 3;
-- S2 :
SELECT budget FROM projets WHERE id = 3;
-- S1 :
COMMIT;
SELECT budget FROM projets WHERE id = 3;''', 'sql', 'SQL exécuté')
term('concurrence', 1, 13, 'Sortie')
answer('l’UPDATE de S1 pose un **verrou exclusif (X) sur la ligne** id = 3. L’UPDATE de S2 sur la même ligne est donc '
       '**mis en attente** jusqu’au COMMIT de S1, puis s’exécute (budget = 3500). Ensuite S2 voit 3500 (sa propre '
       'modification) alors que S1 voit 3000 (valeur validée). Après le COMMIT de S2, **S1 voit encore 3000** : sa '
       'nouvelle transaction (commencée par le SELECT après son COMMIT) lit un snapshot REPEATABLE READ antérieur au COMMIT '
       'de S2. Ce n’est qu’après un nouveau COMMIT de S1 qu’elle voit 3500. Valeur finale : **3500** (la dernière '
       'transaction validée gagne). Le verrouillage empêche la **perte de mise à jour** : les deux UPDATE sont sérialisés.')
H('4.2 Verrouillage explicite : SELECT … FOR UPDATE', 2)
P('La syntaxe du sujet « FOR UPDATE OF prix » vient d’Oracle (OF + colonne). En MySQL 8, OF attend un **nom de table**, '
  'd’où l’erreur 3568 ; on utilise donc FOR UPDATE (ou FOR UPDATE OF projets).')
block('''-- S1 :
SELECT budget FROM projets WHERE id = 3 FOR UPDATE OF budget;  -- erreur (syntaxe Oracle)
SELECT budget FROM projets WHERE id = 3 FOR UPDATE;            -- verrou X sur la ligne
-- S2 :
UPDATE projets SET budget = 4000 WHERE id = 3;                 -- BLOQUÉ
-- S1 :
UPDATE projets SET budget = 5000 WHERE id = 3;
COMMIT;                                                        -- débloque S2
-- S2 :
SELECT budget FROM projets WHERE id = 3;
COMMIT;
-- S1 :
SELECT budget FROM projets WHERE id = 3;''', 'sql', 'SQL exécuté')
term('concurrence', 14, 21, 'Sortie')
answer('le **SELECT … FOR UPDATE** verrouille la ligne dès la lecture : S1 « réserve » la ligne pour la modifier ensuite. '
       'L’UPDATE de S2 attend. S1 met le budget à 5000 et valide ; S2 est alors débloquée et écrase la valeur avec 4000, '
       'puis valide. **Valeur finale : 4000** (ordre de validation : S1 puis S2). FOR UPDATE sert à éviter qu’une autre '
       'transaction modifie une donnée entre le moment où on la lit et celui où on la met à jour (schéma lecture → calcul → écriture).')
doc.add_page_break()

# =================================================================== 5
H('5. Droits d’accès (U1 / U2)', 1)
P('Base COMMANDES du sujet = base tp2_mysql ; la « quantité de produit » modifiée est le budget de la table projets. '
  'Trois connexions sont ouvertes : root, U1 et U2 (dans Workbench : une connexion enregistrée par utilisateur).')
H('5.1 Création des utilisateurs (session root)', 2)
block('''CREATE USER 'U1'@'localhost' IDENTIFIED BY 'u1pass';
GRANT ALL PRIVILEGES ON tp2_mysql.* TO 'U1'@'localhost';
CREATE USER 'U2'@'localhost' IDENTIFIED BY 'u2pass';
GRANT SELECT ON tp2_mysql.* TO 'U2'@'localhost';
SHOW GRANTS FOR 'U1'@'localhost';
SHOW GRANTS FOR 'U2'@'localhost';''', 'sql', 'SQL exécuté (root)')
def mtable(head, rows):
    w = max(len(x) for x in [head] + rows)
    sep = '+' + '-' * (w + 2) + '+'
    return '\n'.join([sep, f'| {head:<{w}} |', sep] + [f'| {r:<{w}} |' for r in rows] + [sep])


block("[root] mysql> SHOW GRANTS FOR 'U1'@'localhost';\n" +
      mtable('Grants for U1@localhost', ['GRANT USAGE ON *.* TO `U1`@`localhost`',
                                         'GRANT ALL PRIVILEGES ON `tp2_mysql`.* TO `U1`@`localhost`']) +
      "\n2 rows in set\n\n[root] mysql> SHOW GRANTS FOR 'U2'@'localhost';\n" +
      mtable('Grants for U2@localhost', ['GRANT USAGE ON *.* TO `U2`@`localhost`',
                                         'GRANT SELECT ON `tp2_mysql`.* TO `U2`@`localhost`']) + '\n2 rows in set',
      'term', 'Sortie')
P('Connexion des utilisateurs : mysql -u U1 -p tp2_mysql   et   mysql -u U2 -p tp2_mysql')
H('5.2 Session U1 : SELECT et UPDATE', 2)
block('''SELECT id, nom, budget FROM projets WHERE id = 2;
UPDATE projets SET budget = 13000 WHERE id = 2;
SELECT id, nom, budget FROM projets WHERE id = 2;''', 'sql', 'SQL exécuté (U1)')
term('droits', 1, 4, 'Sortie')
answer('U1 possède **tous les privilèges** sur la base : SELECT et UPDATE réussissent.')
H('5.3 Session U2 : SELECT et UPDATE', 2)
block('''SELECT id, nom, budget FROM projets WHERE id = 2;
UPDATE projets SET budget = 14000 WHERE id = 2;''', 'sql', 'SQL exécuté (U2)')
term('droits', 5, 7, 'Sortie')
answer('le SELECT fonctionne, mais l’UPDATE est **refusé : ERROR 1142 « UPDATE command denied »**, car U2 n’a que le '
       'privilège SELECT.')
H('5.4 U1 tente de donner le droit UPDATE à U2', 2)
block("GRANT UPDATE ON tp2_mysql.projets TO 'U2'@'localhost';", 'sql', 'SQL exécuté (U1)')
term('droits', 8, 8, 'Sortie')
answer('refus : **ERROR 1142 « GRANT command denied »**. « ALL PRIVILEGES » **n’inclut pas** le privilège GRANT OPTION : '
       'un utilisateur ne peut transmettre ses droits que s’il a reçu explicitement l’option WITH GRANT OPTION.')
H('5.5 Root redonne le droit à U1 avec WITH GRANT OPTION, puis U1 réessaie', 2)
block('''-- root :
GRANT ALL PRIVILEGES ON tp2_mysql.* TO 'U1'@'localhost' WITH GRANT OPTION;
SHOW GRANTS FOR 'U1'@'localhost';
-- U1 (même session, sans se reconnecter) :
GRANT UPDATE ON tp2_mysql.projets TO 'U2'@'localhost';''', 'sql', 'SQL exécuté')
term('droits', 9, 11, 'Sortie')
answer('point important observé : dans la **session déjà ouverte** de U1, le GRANT est **encore refusé**. En MySQL, un '
       'changement de privilèges **au niveau base de données** (ON tp2_mysql.*) n’est pris en compte par une session '
       'existante qu’après une nouvelle instruction **USE tp2_mysql;** ou une **reconnexion** (les privilèges de niveau '
       'table/colonne, eux, sont pris en compte immédiatement). Il faut donc que U1 se reconnecte :')
block('''-- U1 : reconnexion (ou USE tp2_mysql;)
USE tp2_mysql;
GRANT UPDATE ON tp2_mysql.projets TO 'U2'@'localhost';
-- U2 :
UPDATE projets SET budget = 14000 WHERE id = 2;
SELECT id, nom, budget FROM projets WHERE id = 2;''', 'sql', 'SQL à exécuter après la prise en compte du nouveau droit')
block('''[U1] mysql> USE tp2_mysql;
Database changed
[U1] mysql> GRANT UPDATE ON tp2_mysql.projets TO 'U2'@'localhost';
Query OK, 0 rows affected

[U2] mysql> UPDATE projets SET budget = 14000 WHERE id = 2;
Query OK, 1 row affected
Rows matched: 1  Changed: 1  Warnings: 0''', 'term', 'Résultat attendu après reconnexion de U1')
answer('avec **WITH GRANT OPTION**, U1 peut désormais transmettre ses droits : U2 obtient le droit UPDATE sur la table '
       'projets et sa mise à jour réussit (droit de niveau table, effectif immédiatement pour U2). Root peut retirer '
       'ces droits avec REVOKE.')
H('5.6 U1 insère un nouveau projet ; root et U2 font un SELECT', 2)
block('''-- U1 :
SET autocommit = 0;
INSERT INTO projets (nom, budget, client_id) VALUES ('Projet cree par U1', 21000, 4);
SELECT id, nom, budget FROM projets;
-- root et U2 :
SELECT id, nom, budget FROM projets;
-- U1 :
COMMIT;
-- root et U2 :
SELECT id, nom, budget FROM projets;''', 'sql', 'SQL exécuté')
term('droits', 14, 21, 'Sortie')
answer('tant que U1 n’a pas fait **COMMIT**, le nouveau projet n’est visible **que par U1**, ni root ni U2 ne le voient, '
       '**même root** (le super‑utilisateur n’a pas plus de droits sur l’isolation des transactions). Après le COMMIT, '
       'root et U2 voient le projet. Les **droits** (GRANT) contrôlent ce qu’un utilisateur a le droit de faire ; '
       'l’**isolation des transactions** contrôle ce qu’il voit des transactions des autres : ce sont deux mécanismes indépendants.')
doc.add_page_break()

# =================================================================== 6
H('6. Transactions et reprise sur panne — tableaux p1 / p2', 1)
P('Table T(A INTEGER, B INTEGER) contenant initialement (1,2). p1 = root, autocommit désactivé dès la connexion ; '
  'p2 = user1, « autre client quelconque » (autocommit actif : il voit l’état validé). Après chaque ordre de p1, on '
  'exécute SELECT * FROM T dans p1 (« T ds trans p1 ») et dans p2 (« T vue par p2 »). La panne serveur est simulée en '
  'arrêtant brutalement le service MySQL (net stop MySQL84, ou fin de tâche de mysqld.exe), puis en le redémarrant.')
block('''DROP TABLE IF EXISTS T;
CREATE TABLE T (A INTEGER, B INTEGER) ENGINE = InnoDB;
INSERT INTO T VALUES (1,2);
COMMIT;''', 'sql', 'Préparation (avant chaque tableau)')

H('6.1 Tableau 1', 2)
block('''SET autocommit = 0;
INSERT INTO T VALUES (3,4);
INSERT INTO T VALUES (5,6);
ROLLBACK;
COMMIT;
INSERT INTO T VALUES (3,4);
EXIT;''', 'sql', 'Ordres de p1')
grid(['Ordre (p1)', 'T ds trans p1', 'T vue par p2', 'Transaction'],
     table_from_log('tab1', ['p1'], ['p1', 'p2'],
                    {1: 'état initial', 2: 'DÉBUT transaction t1', 4: 'FIN t1 (annulée)', 5: 'transaction vide (DÉBUT+FIN)',
                     6: 'DÉBUT t3', 7: 'FIN t3 : rollback implicite'}), [5.5, 3.5, 3.5, 4.5])
answer('ROLLBACK annule (3,4) et (5,6). Le COMMIT qui suit valide une transaction vide. Le dernier INSERT (3,4) n’est '
       'jamais validé : à l’**EXIT**, MySQL fait un **ROLLBACK** implicite. État final de T : **(1,2)**. '
       '(Remarque : sous Oracle/SQL*Plus, EXIT valide la transaction ; en MySQL, la déconnexion annule.)')
term('tab1', caption='Sortie (extrait de vérification)', start=6)

H('6.2 Tableau 2', 2)
block('''SET autocommit = 0;
ROLLBACK;
INSERT INTO T VALUES (3,4);
ALTER TABLE T ADD COLUMN C INTEGER;
UPDATE T SET B = 3 WHERE A = 1;
-- PANNE : arrêt brutal du service mysql, puis redémarrage
SELECT * FROM T;''', 'sql', 'Ordres de p1')
grid(['Ordre (p1)', 'T ds trans p1', 'T vue par p2', 'Transaction'],
     table_from_log('tab2', ['p1'], ['p1', 'p2'],
                    {1: 'état initial', 2: 'FIN (rien à annuler)', 3: 'DÉBUT t1',
                     4: 'COMMIT implicite de t1 (DDL) ; FIN t1', 5: 'DÉBUT t2', 6: 'FIN t2 : annulée par la reprise'}),
     [5.5, 3.5, 3.5, 4.5])
answer('l’**ALTER TABLE** est un ordre DDL : MySQL effectue un **COMMIT implicite** avant et après. L’insertion (3,4) '
       'devient donc définitive et visible par p2 (avec la nouvelle colonne C à NULL). L’UPDATE B = 3 démarre une nouvelle '
       'transaction non validée : lors du redémarrage après la panne, InnoDB effectue la **reprise** (rejoue le journal '
       'redo des transactions validées, annule via l’undo les transactions non validées). État final : '
       '**(1,2,NULL) (3,4,NULL)** — la mise à jour B = 3 est perdue.')
term('tab2', caption='Sortie après la panne', start=6)

H('6.3 Tableau 3', 2)
block('''SET autocommit = 0;
INSERT INTO T VALUES (3,4);
INSERT INTO T VALUES (5,6);
SET autocommit = 1;        -- "SET AUTOCOMMIT ON"
INSERT INTO T VALUES (7,8);
ROLLBACK;
INSERT INTO T VALUES (9,10);
DELETE FROM T WHERE A = 1;
-- PANNE : arrêt brutal du service mysql, puis redémarrage
SELECT * FROM T;''', 'sql', 'Ordres de p1')
grid(['Ordre (p1)', 'T ds trans p1', 'T vue par p2', 'Transaction'],
     table_from_log('tab3', ['p1'], ['p1', 'p2'],
                    {1: 'état initial', 2: 'DÉBUT t1', 4: 'COMMIT implicite : FIN t1',
                     5: 'autocommit : DÉBUT+FIN', 6: 'sans effet (rien à annuler)', 7: 'autocommit : DÉBUT+FIN',
                     8: 'autocommit : DÉBUT+FIN', 9: 'aucune perte'}),
     [5.5, 3.5, 3.5, 4.5])
answer('passer de autocommit = 0 à autocommit = 1 **valide la transaction en cours** : (3,4) et (5,6) deviennent '
       'visibles par p2. Ensuite chaque ordre est une transaction à lui seul, validée immédiatement ; le **ROLLBACK n’a '
       'donc aucun effet** sur (7,8). Après la panne, toutes ces modifications validées sont **conservées** (durabilité) : '
       'état final **(3,4) (5,6) (7,8) (9,10)**.')
term('tab3', caption='Sortie après la panne', start=9)
doc.add_page_break()

# =================================================================== 7
H('7. Analyse de programme — tableaux p1 / p2 / p3', 1)
P('T(A, B) contient au départ (1,1) et (2,2). p1 (root) et p2 (user1) ont l’autocommit désactivé ; p3 (user1) est un '
  'client quelconque en autocommit qui voit l’état validé. Isolation : REPEATABLE READ (défaut). Les SELECT de p1/p2 '
  'lisent leur **snapshot** (pris à leur première lecture) plus leurs propres modifications ; les UPDATE/DELETE font, '
  'eux, une **lecture courante** des dernières versions validées et posent des verrous exclusifs.')
block('''DROP TABLE IF EXISTS T;
CREATE TABLE T (A INTEGER, B INTEGER) ENGINE = InnoDB;
INSERT INTO T VALUES (1,1), (2,2);
-- p1 et p2 : SET autocommit = 0;''', 'sql', 'Préparation (avant chaque tableau)')
HDR = ['Client : ordre', 'T ds trans p1', 'T ds trans p2', 'T vue par p3', 'Commentaire']
W = [4.2, 2.7, 2.7, 2.7, 4.6]

H('7.1 Tableau 1', 2)
grid(HDR, table_from_log('an1', ['p1', 'p2'], ['p1', 'p2', 'p3'],
                         {2: 'état initial', 3: 'DÉBUT trans p1 ; non visible ailleurs', 4: 'FIN trans p1 : visible par p3'}), W)
answer('la suppression de (1,1) n’est visible que par p1 jusqu’au COMMIT ; ensuite p3 voit (2,2). p2, dont la '
       'transaction a commencé avant (snapshot), voit encore (1,1)(2,2) tant qu’elle ne termine pas sa transaction.')

H('7.2 Tableau 2', 2)
grid(HDR, table_from_log('an2', ['p1', 'p2'], ['p1', 'p2', 'p3'],
                         {2: 'état initial', 3: 'DÉBUT trans p1 ; verrou X sur (1,1)',
                          4: 'DÉBUT trans p2 ; attend le verrou de p1 sur (1,1)',
                          5: 'FIN p1 ; p2 débloquée : ne modifie que (2,2)', 6: 'FIN p2'}), W)
answer('l’UPDATE de p2 doit verrouiller toutes les lignes, dont (1,1) déjà verrouillée par le DELETE de p1 : il '
       '**attend**. Au COMMIT de p1, (1,1) n’existe plus : p2 ne met à jour que (2,2) → (2,1) (« 1 row affected »). '
       'p2 voit (1,1)(2,1) car son SELECT combine son snapshot (où (1,1) existe encore) et sa propre modification. '
       'État final validé : **(2,1)**. Exécution équivalente à la série p1 puis p2.')
term('an2', 3, 5, 'Sortie (extrait : blocage puis déblocage)')

H('7.3 Tableau 3', 2)
grid(HDR, table_from_log('an3', ['p1', 'p2'], ['p1', 'p2', 'p3'],
                         {2: 'état initial', 3: 'DÉBUT trans p2 ; verrous X sur les 2 lignes',
                          4: 'DÉBUT trans p1 ; attend les verrous de p2',
                          5: 'FIN p2 ; p1 débloquée : supprime (2,1)', 6: 'FIN p1'}), W)
answer('p2 décrémente B sur les deux lignes : (1,0)(2,1). Le DELETE de p1 attend. Après le COMMIT de p2, le DELETE '
       'relit les valeurs validées et supprime la ligne où B vaut maintenant 1, c’est‑à‑dire **(2,1)** et non (1,1) ! '
       'Le SELECT de p1 affiche (1,1) (snapshot ancien sans la ligne 2 qu’il a supprimée). État final validé : '
       '**(1,0)**. Exécution équivalente à la série p2 puis p1.')
term('an3', 4, 5, 'Sortie (extrait : blocage puis déblocage)')

H('7.4 Tableau 4', 2)
grid(HDR, table_from_log('an4', ['p1', 'p2'], ['p1', 'p2', 'p3'],
                         {2: 'état initial', 3: 'DÉBUT trans p1 ; (2,2) devient (1,2)',
                          4: 'DÉBUT trans p2 ; attend le verrou de p1',
                          5: 'FIN p1 ; p2 débloquée : supprime 2 lignes', 6: 'FIN p2 : table vide'}), W)
answer('p1 transforme (2,2) en (1,2). Le DELETE WHERE A = 1 de p2 attend la fin de p1. Après le COMMIT de p1, il '
       'relit les lignes validées : **les deux lignes ont maintenant A = 1**, il en supprime donc 2. État final : '
       '**table vide**. Exécution équivalente à p1 puis p2.')
term('an4', 4, 5, 'Sortie (extrait : blocage puis déblocage)')
doc.add_page_break()

# =================================================================== 8
H('8. Conclusion générale', 1)
for b in [
    '**Atomicité** : une transaction est validée entièrement (COMMIT) ou annulée entièrement (ROLLBACK). Toute transaction '
    'non validée est annulée en cas de déconnexion, de plantage du client ou de panne du serveur.',
    '**Durabilité** : tout ce qui a été validé survit à une panne du serveur (journal redo d’InnoDB rejoué au redémarrage).',
    '**Validations implicites** en MySQL : les ordres DDL (ALTER, CREATE, DROP…) et le passage à autocommit = 1 valident '
    'la transaction en cours ; à l’inverse la déconnexion (EXIT) l’annule.',
    '**Isolation** : MySQL/InnoDB est par défaut en REPEATABLE READ (lectures sur un snapshot cohérent). READ UNCOMMITTED '
    'autorise les lectures sales ; SERIALIZABLE ajoute des verrous partagés aux lectures et provoque des attentes.',
    '**Concurrence** : les écritures concurrentes sur une même ligne sont sérialisées par des verrous exclusifs '
    '(attente jusqu’au COMMIT/ROLLBACK). SELECT … FOR UPDATE permet de verrouiller une ligne dès sa lecture. Une écriture '
    'débloquée relit la dernière version validée, ce qui peut changer les lignes affectées (tableaux 7.2 à 7.4).',
    '**Droits** : GRANT/REVOKE contrôlent les actions autorisées ; ALL PRIVILEGES n’inclut pas GRANT OPTION ; un '
    'changement de droit de niveau base n’est appliqué à une session ouverte qu’après USE ou reconnexion.',
]:
    bullet(b)

doc.save(OUT)
print('OK', OUT)

import json, re, subprocess
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from run_tp import S, BIN, CONN, DB

L = json.load(open(r'C:\Users\tomga\tp_sql_avance_work\log.json', encoding='utf-8'))
OUT = r'C:\Users\tomga\Downloads\Rapport_TP_SQL_Avance_MySQL.docx'
OUT_SQL = r'C:\Users\tomga\Downloads\TP_SQL_Avance_requetes.sql'

doc = Document()
sec = doc.sections[0]
sec.left_margin = sec.right_margin = Cm(2)
sec.top_margin = sec.bottom_margin = Cm(1.8)
st = doc.styles['Normal']
st.font.name = 'Calibri'
st.font.size = Pt(11)
WIDTH_PT = [484]          # largeur utile d'un bloc de code (pt), modifiée en paysage


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
        c.paragraph_format.keep_with_next = True
    lines = text.rstrip('\n').split('\n')
    longest = max(len(l) for l in lines)
    size = min(8.0, int(WIDTH_PT[0] / (longest * 0.55) * 2) / 2)
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t.rows[0].cells[0]
    if len(lines) <= 45:        # un bloc court n'est jamais coupé entre deux pages
        t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
    cell_shade(cell, '1E1E1E' if kind == 'term' else 'F2F4F7')
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    for i, line in enumerate(lines):
        if i:
            p.add_run().add_break()
        if kind == 'term' and line.startswith(('mysql>', '    ->', 'PS ')):
            head = line[:6] if not line.startswith('PS ') else line.split('>', 1)[0] + '>'
            for txt, col in ((head, (0x5A, 0xC8, 0x6E)), (line[len(head):], (0xF0, 0xF0, 0xF0))):
                r = p.add_run(txt); r.bold = True
                r.font.name = 'Consolas'; r.font.size = Pt(size); r.font.color.rgb = RGBColor(*col)
            continue
        r = p.add_run(line)
        r.font.name = 'Consolas'; r.font.size = Pt(size)
        if kind == 'term':
            col = (0xFF, 0x6B, 0x6B) if 'ERROR' in line else (0xD8, 0xD8, 0xDC)
            r.font.color.rgb = RGBColor(*col)
        else:
            r.font.color.rgb = RGBColor(0x8A, 0x8A, 0x8A) if line.strip().startswith('--') else RGBColor(0x1F, 0x3A, 0x68)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def H(txt, lvl=1):
    doc.add_heading(txt, lvl)


def rich(p, txt):
    for i, part in enumerate(re.split(r'\*\*', txt)):
        p.add_run(part).bold = bool(i % 2)


def P(txt):
    p = doc.add_paragraph()
    rich(p, txt)
    p.paragraph_format.space_after = Pt(4)
    return p


def bullet(txt):
    rich(doc.add_paragraph(style='List Bullet'), txt)


def answer(txt, label='Réponse / observation : '):
    t = doc.add_table(rows=1, cols=1)
    c = t.rows[0].cells[0]
    cell_shade(c, 'E8F4EA')
    p = c.paragraphs[0]
    p.add_run(label).bold = True
    rich(p, txt)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


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
            if i == 0:
                r.bold = True
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def prompt(cmd):
    ls = cmd.split('\n')
    end = '' if cmd.endswith('\\G') else ';'
    return '\n'.join(('mysql> ' if i == 0 else '    -> ') + l for i, l in enumerate(ls)) + end


def term(key, caption='Sortie', only=None):
    """Bloc terminal construit à partir du journal d'exécution réel (invite mysql> + résultat)."""
    out = []
    for i, e in enumerate(L[key]):
        if only is not None and i not in only:
            continue
        out += [prompt(e['cmd']), e['out'], '']
    block('\n'.join(out), 'term', caption)


TITLES = {}


def ex(key, title, obs=None, intro=None):
    """Un exercice : titre, SQL, sortie réelle, observation."""
    TITLES[key] = title
    H(f'{key}. {title}', 2)
    if intro:
        P(intro)
    if len(L[key]) == 1:
        e = L[key][0]
        block(e['cmd'] + ';', 'sql', 'SQL exécuté')
        block(e['out'], 'term', 'Sortie')
    else:
        term(key, 'SQL exécuté et sortie')
    if obs:
        answer(obs)


# =================================================================== PAGE DE TITRE
for _ in range(5):
    doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Rapport de TP — SQL avancé / MySQL'); r.bold = True; r.font.size = Pt(28); r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x68)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Jointures (INNER, LEFT, RIGHT, FULL OUTER, SELF), GROUP BY / HAVING, CASE,\n'
              'sous-requêtes, EXISTS, CTE, vues et requêtes appliquées à la cybersécurité'); r.font.size = Pt(15)
doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run('B3 — TP de SQL avancé\nSGBD : MySQL Server 8.4.9 (moteur InnoDB) — base entreprise_avancee\n'
          'Client : mysql en ligne de commande (terminal)').font.size = Pt(12)
for _ in range(3):
    doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Tom GALVAN'); r.bold = True; r.font.size = Pt(16)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run('ECE — École d’ingénieurs').font.size = Pt(13)
doc.core_properties.author = 'Tom Galvan'
doc.core_properties.title = 'Rapport de TP — SQL avancé / MySQL'
doc.add_page_break()

# =================================================================== SOMMAIRE
H('Sommaire', 1)
for s in ['0. Environnement et démarrage depuis le terminal', '1. Schéma, création des tables et jeu de données',
          '2. Série A — INNER JOIN', '3. Série B — LEFT JOIN et NULL', '4. Série C — RIGHT JOIN',
          '5. Série D — Simulation de FULL OUTER JOIN', '6. Série E — SELF JOIN',
          '7. Série F — JOIN + GROUP BY + HAVING', '8. Série G — CASE', '9. Série H — Sous-requêtes',
          '10. Série I — EXISTS / NOT EXISTS', '11. Série J — CTE (WITH)', '12. Série K — Vues (VIEW)',
          '13. Série L — Requêtes appliquées à la cybersécurité', '14. MySQL dans le terminal',
          '15. Défi final — Tableau de bord de réponse aux incidents', '16. Défi bonus — Détection des données orphelines',
          '17. Réponses aux questions conceptuelles sur les jointures', '18. Conclusion générale']:
    bullet(s)
doc.add_page_break()

# =================================================================== 0
H('0. Environnement et démarrage depuis le terminal', 1)
P('Le sujet est rédigé pour un terminal Linux (commande **sudo mysql**). Le TP a été réalisé avec le même client en ligne '
  'de commande **mysql**, lancé depuis un terminal Windows, sur un serveur **MySQL Community Server 8.4.9**. '
  'Les requêtes SQL sont strictement identiques ; seule la commande de connexion change :')
grid(['Sujet', 'Équivalent utilisé'],
     [['Terminal Linux', 'Terminal Windows (PowerShell)'],
      ['sudo mysql', 'mysql -u root (compte root local du serveur de TP)'],
      ['MySQL / MariaDB', 'MySQL Server 8.4.9, moteur InnoDB'],
      ['Base entreprise_avancee', 'Base entreprise_avancee (créée pour ce TP)']], widths=[6, 11.5])
P('Toutes les sorties reproduites dans ce rapport sont celles réellement renvoyées par le serveur lors de l’exécution des '
  'requêtes. Les requêtes des séries A à L sont également regroupées dans le fichier **TP_SQL_Avance_requetes.sql** joint au rapport.')

H('0.1 Connexion à MySQL', 2)
status = []
for l in L['status'][0]['out'].split('\n'):
    if l.startswith(('---', 'Threads', 'Uptime')) or not l.strip():
        continue
    m = re.match(r'^([^:\t]+:)\s*(.*)$', l)
    status.append(f'{m.group(1):<24}{m.group(2)}' if m and not l.startswith('C:') else l)
status = '\n'.join(status)
block('PS C:\\Users\\tomga> mysql -u root\nmysql> STATUS;\n' + status, 'term', 'Connexion et état de la session')
P('**Capture d’écran du terminal (connexion à MySQL) :** à insérer ici.')

H('0.2 Création et sélection de la base', 2)
term('demarrage', 'SQL exécuté et sortie')
answer('**SHOW DATABASES** liste toutes les bases du serveur visibles par l’utilisateur connecté : c’est une commande de '
       'consultation, elle ne change rien à la session. **USE entreprise_avancee** sélectionne la base courante de la session : '
       'les requêtes suivantes s’y exécutent sans avoir à préfixer les tables par le nom de la base. '
       '**SELECT DATABASE()** est une fonction qui renvoie le nom de la base actuellement sélectionnée (NULL si aucune) : '
       'elle sert à vérifier que le USE a bien été pris en compte. En résumé : lister, choisir, vérifier.',
       'Question — différence entre SHOW DATABASES, USE et SELECT DATABASE() : ')
doc.add_page_break()

# =================================================================== 1
H('1. Schéma, création des tables et jeu de données', 1)
H('1.1 Schéma de la base', 2)
grid(['Table', 'Rôle', 'Clé / relation'],
     [['departements', 'Départements de l’entreprise', 'id (clé primaire)'],
      ['employes', 'Employés', 'departement_id → departements.id ; manager_id → employes.id (auto-référence)'],
      ['clients', 'Clients', 'id (clé primaire)'],
      ['projets', 'Projets', 'client_id → clients.id'],
      ['affectations', 'Affectations des employés aux projets (table d’association N–N)',
       'clé primaire composée (employe_id, projet_id)'],
      ['incidents', 'Incidents de sécurité', 'employe_id → employes.id ; projet_id → projets.id']],
     widths=[3, 6.5, 8])
P('Les clés étrangères departement_id, manager_id, client_id, employe_id et projet_id (table incidents) acceptent la valeur '
  '**NULL** : c’est ce qui permet d’avoir des lignes « sans correspondance », indispensables pour observer la différence '
  'entre les types de jointures.')

H('1.2 Création des tables', 2)
block('\n\n'.join(s + ';' for s in S['tables'][:-1]), 'sql', 'SQL exécuté (script du sujet)')
block('\n'.join(e['out'] for e in L['tables'][:-1]) + '\n\n' + prompt('SHOW TABLES') + '\n' + L['tables'][-1]['out'],
      'term', 'Sortie')
answer('les six tables sont créées (Query OK pour chaque CREATE TABLE). L’ordre de création est imposé par les clés étrangères : '
       'departements avant employes, clients avant projets, puis affectations et incidents qui référencent employes et projets.')

H('1.3 Insertion du jeu de données', 2)
block('\n\n'.join(s + ';' for s in S['donnees']), 'sql', 'SQL exécuté (script du sujet)')
block('\n\n'.join(e['out'] for e in L['donnees']), 'term', 'Sortie (une réponse par INSERT)')

H('1.4 Vérification du jeu de données', 2)
P('Les six SELECT COUNT(*) du sujet sont regroupés avec UNION ALL pour obtenir un seul tableau récapitulatif.')
block(L['verif'][0]['cmd'] + ';', 'sql', 'SQL exécuté')
block(L['verif'][0]['out'], 'term', 'Sortie')
answer('toutes les tables en contiennent, et c’est voulu : '
       '**employes** — Adam Faure n’a pas de département (departement_id NULL) ; Lucas Bernard et Adam Faure ne sont affectés à aucun projet. '
       '**departements** — Innovation n’a aucun employé. '
       '**clients** — NoProjectClient n’a aucun projet. '
       '**projets** — Internal Research n’a ni client (client_id NULL) ni affectation ; Server Security et Internal Research n’ont aucun incident. '
       '**incidents** — l’incident n° 6 (Suspicious Login) n’est rattaché ni à un employé ni à un projet. '
       'Ces lignes disparaissent avec un INNER JOIN et réapparaissent, complétées par des NULL, avec une jointure externe.',
       'Question — quelles tables contiennent volontairement des lignes sans correspondance ? ')
doc.add_page_break()

# =================================================================== A
H('2. Série A — INNER JOIN', 1)
P('L’INNER JOIN ne conserve que les lignes qui ont une correspondance dans les deux tables. Conformément à la contrainte du sujet, '
  'aucune requête n’utilise SELECT * : les colonnes sont nommées explicitement et les tables portent des alias.')
ex('A1', 'Prénom, nom et département de chaque employé ayant un département',
   '8 lignes pour 9 employés : **Adam Faure** (departement_id NULL) n’apparaît pas, car NULL ne correspond à aucun departements.id.')
ex('A2', 'Projets avec le nom de leur client',
   '6 lignes pour 7 projets : **Internal Research** (client_id NULL) est éliminé, et le client NoProjectClient n’apparaît pas non plus.')
ex('A3', 'Chaque affectation avec l’employé, le projet, le rôle et le nombre d’heures',
   'la table affectations est la table d’association : on la relie à employes et à projets par deux INNER JOIN. '
   'On retrouve les 10 affectations, aucune n’est perdue car employe_id et projet_id font partie de la clé primaire (jamais NULL).')
ex('A4', 'Incidents avec l’employé et le projet associés',
   '6 incidents sur 7 : l’incident n° 6 (**Suspicious Login**) a employe_id et projet_id à NULL, il est donc écarté par les INNER JOIN. '
   'Dans un contexte SOC, c’est un piège : un incident non attribué deviendrait invisible (voir L1 et le défi final, écrits avec LEFT JOIN).')
ex('A5', 'Requête sur trois tables : employes → affectations → projets',
   'la chaîne employes → affectations → projets donne, pour les projets « In Progress », qui travaille dessus, à quel poste et pour combien d’heures. '
   'Le filtre WHERE s’applique après les jointures ; les employés sans affectation (Lucas Bernard, Adam Faure) sont absents.')
doc.add_page_break()

# =================================================================== B
H('3. Série B — LEFT JOIN et NULL', 1)
P('Le LEFT JOIN conserve toutes les lignes de la table de gauche ; lorsqu’il n’existe pas de correspondance à droite, '
  'les colonnes de la table de droite valent NULL.')
ex('B1', 'Tous les employés, y compris ceux sans département',
   '9 lignes cette fois : **Adam Faure** est conservé, avec departement = NULL. C’est la seule différence avec A1.')
ex('B2', 'Employés affectés à aucun projet',
   'technique de l’« anti-jointure » : LEFT JOIN puis WHERE … IS NULL sur une colonne de la table de droite qui ne peut pas être NULL '
   '(ici la clé a.employe_id). Il ne reste que les employés sans correspondance : **Lucas Bernard** et **Adam Faure**.')
ex('B3', 'Tous les clients et leurs projets, y compris les clients sans projet',
   '7 lignes : CyberSolutions apparaît deux fois (deux projets) et **NoProjectClient** est conservé avec projet = NULL.')
ex('B4', 'Clients n’ayant aucun projet',
   'même anti-jointure que B2 : seul **NoProjectClient** n’a pas de projet.')
ex('B5', 'Tous les projets et le nombre d’employés affectés',
   'il faut compter **COUNT(a.employe_id)** et non COUNT(*) : COUNT(*) compterait la ligne « vide » produite par le LEFT JOIN et '
   'afficherait 1 pour Internal Research, alors que COUNT(colonne) ignore les NULL et donne bien **0**.')
ex('B6', 'Tous les départements et leur nombre d’employés, y compris les départements vides',
   'la table departements est à gauche pour être conservée entièrement : **Innovation** apparaît avec 0 employé. '
   'Adam Faure (sans département) n’est compté nulle part.')
H('Défi — condition WHERE sur la table de droite après un LEFT JOIN', 2)
P('On veut la liste de **tous** les clients, avec leurs projets « In Progress » s’ils en ont. Première requête : le filtre est dans le WHERE ; '
  'seconde requête : le filtre est dans la condition ON.')
term('Bdefi', 'SQL exécuté et sortie')
answer('le WHERE est évalué **après** la jointure. Pour un client sans projet, le LEFT JOIN produit p.statut = NULL ; or NULL = \'In Progress\' '
       'n’est pas vrai, donc la ligne est supprimée. Le filtre élimine aussi les clients dont les projets ont un autre statut '
       '(TechCorp, Medialab). Résultat : 3 lignes, exactement ce que donnerait un INNER JOIN — le LEFT JOIN est « annulé ». '
       'En plaçant la condition dans le **ON**, elle fait partie du critère de correspondance : les 6 clients sont conservés et ceux qui '
       'n’ont pas de projet en cours reçoivent NULL. Règle retenue : une condition sur la table de droite se met dans le ON '
       '(sauf le test IS NULL de l’anti-jointure, qui doit rester dans le WHERE).')
doc.add_page_break()

# =================================================================== C
H('4. Série C — RIGHT JOIN', 1)
P('Le RIGHT JOIN est le symétrique du LEFT JOIN : toutes les lignes de la table de droite sont conservées.')
ex('C1', 'Réécriture d’une requête LEFT JOIN (B3) avec RIGHT JOIN',
   'on inverse simplement l’ordre des tables : « clients LEFT JOIN projets » devient « projets RIGHT JOIN clients ». '
   'Le résultat est identique à B3 (7 lignes, NoProjectClient conservé).')
ex('C2', 'Tous les départements et leurs employés avec RIGHT JOIN',
   '9 lignes : les 8 employés rattachés à un département, plus **Innovation** avec prenom/nom à NULL. '
   'Adam Faure, qui est dans la table de gauche sans correspondance, disparaît.')
ex('C3', 'Réécriture de C2 avec LEFT JOIN, sans changer le résultat',
   'résultat strictement identique à C2 : A RIGHT JOIN B équivaut à B LEFT JOIN A.')
answer('les deux sont équivalents, mais le LEFT JOIN se lit dans le sens naturel de la requête : la table principale est écrite en premier '
       '(FROM) et l’on y « accroche » des informations facultatives. Dans une requête à plusieurs jointures, n’utiliser que des LEFT JOIN '
       'permet de savoir immédiatement quelle table est conservée entièrement ; mélanger LEFT et RIGHT rend la requête difficile à relire '
       'et à maintenir. C’est pourquoi la plupart des conventions d’équipe imposent LEFT JOIN, RIGHT JOIN restant rare.',
       'Question — pourquoi préférer LEFT JOIN en projet professionnel ? ')
doc.add_page_break()

# =================================================================== D
H('5. Série D — Simulation de FULL OUTER JOIN avec MySQL', 1)
P('MySQL ne propose pas FULL OUTER JOIN. On le simule par l’union d’un LEFT JOIN (toutes les lignes de gauche) et d’un RIGHT JOIN '
  '(toutes les lignes de droite).')
ex('D1', 'Simulation de FULL OUTER JOIN entre clients et projets',
   '8 lignes : les 6 correspondances, plus le client sans projet (**NoProjectClient**, colonnes projet à NULL) et le projet sans client '
   '(**Internal Research**, colonnes client à NULL). Aucune ligne des deux tables n’est perdue.')
ex('D2', 'Lignes provenant uniquement de la table clients',
   'la simulation est placée dans une table dérivée (fj) puis filtrée : projet_id IS NULL signifie « aucun projet en face », '
   'donc une ligne issue uniquement de clients : **NoProjectClient**.')
ex('D3', 'Lignes provenant uniquement de la table projets',
   'symétriquement, client_id IS NULL isole les lignes venant uniquement de projets : **Internal Research**.')
ex('D4', 'Rôle de UNION dans la requête',
   'UNION empile les résultats des deux requêtes **et supprime les doublons**. Or les 6 lignes « avec correspondance » sont produites '
   'à la fois par le LEFT JOIN et par le RIGHT JOIN : avec UNION on obtient 8 lignes (le bon résultat), alors qu’avec **UNION ALL** '
   'on en obtient 14 (7 + 7), les 6 correspondances étant comptées deux fois. UNION impose aussi que les deux SELECT aient le même '
   'nombre de colonnes, dans le même ordre. Limite : UNION supprimerait aussi de « vrais » doublons ; l’alternative exacte est '
   'LEFT JOIN … UNION ALL … RIGHT JOIN … WHERE gauche.id IS NULL.',
   intro='Pour mettre en évidence le rôle de UNION, on compte les lignes obtenues avec UNION puis avec UNION ALL.')
doc.add_page_break()

# =================================================================== E
H('6. Série E — SELF JOIN', 1)
P('La colonne employes.manager_id référence employes.id : la table est jointe à elle-même, avec deux alias pour distinguer les deux '
  'rôles (**e** = l’employé, **m** = son manager).')
ex('E1', 'Chaque employé avec le prénom et le nom de son manager',
   'LEFT JOIN pour conserver les 9 employés : ceux qui n’ont pas de manager ont NULL dans les colonnes du manager.')
ex('E2', 'Uniquement les employés qui ont un manager',
   'avec un INNER JOIN, seuls les 4 employés ayant un manager_id renseigné restent. Paul Martin encadre deux personnes (Lucas Bernard et Yanis Moreau).')
ex('E3', 'Employés qui n’ont pas de manager',
   'anti-jointure sur la table elle-même : 5 employés sans manager (les trois responsables Paul Martin, Sophie Durand, Nadia Robert, '
   'ainsi que Lina Roux et Adam Faure). Un simple WHERE manager_id IS NULL donnerait le même résultat.')
ex('E4', 'Tableau employé / poste / manager',
   'CONCAT assemble prénom et nom ; COALESCE remplace le NULL par « (aucun) » pour un affichage lisible '
   '(CONCAT renvoie NULL dès qu’un de ses arguments est NULL).')
doc.add_page_break()

# =================================================================== F
H('7. Série F — JOIN + GROUP BY + HAVING', 1)
P('WHERE filtre les lignes avant le regroupement ; HAVING filtre les groupes après le calcul des agrégats.')
ex('F1', 'Nombre d’employés par département',
   'avec l’INNER JOIN, seuls les départements ayant au moins un employé apparaissent (Innovation est absent, contrairement à B6).')
ex('F2', 'Salaire moyen par département',
   'AVG est calculé par groupe ; ROUND(…, 2) arrondit au centime. Cybersecurity a la moyenne la plus élevée (4 400 €).')
ex('F3', 'Départements dont le salaire moyen dépasse 4 000 €',
   'la condition porte sur un agrégat, elle doit donc être dans **HAVING** (un WHERE AVG(…) > 4000 provoquerait une erreur). '
   'Restent Cybersecurity (4 400) et Data (4 200) ; Networking (3 950) est écarté.')
ex('F4', 'Budget total des projets par client',
   'CyberSolutions cumule deux projets (25 000 + 12 000 = 37 000). Le projet sans client (Internal Research) n’est pas comptabilisé.')
ex('F5', 'Clients dont le budget total dépasse 30 000 €',
   'HAVING SUM(p.budget) > 30000 conserve Medialab, CyberSolutions et FinSecure. À noter : CyberSolutions n’a aucun projet de plus de 30 000 € '
   'individuellement ; c’est bien la somme par groupe qui est testée.')
ex('F6', 'Nombre total d’heures affectées par projet',
   'LEFT JOIN + COALESCE pour afficher aussi **Internal Research** avec 0 heure (SUM sur un groupe sans ligne renvoie NULL).')
doc.add_page_break()

# =================================================================== G
H('8. Série G — CASE', 1)
ex('G1', 'Catégorie salariale de chaque employé (Junior / Confirmed / Senior)',
   'les conditions d’un CASE sont évaluées dans l’ordre et la première vraie l’emporte : il est donc inutile d’écrire '
   '« salaire >= 3500 AND salaire <= 4200 ». Les bornes sont respectées : 3 500 (Karim Petit) et 4 200 (Lina Roux) sont « Confirmed ».')
ex('G2', 'Classification des incidents (Urgent / Important / Standard)',
   'forme « simple » du CASE (CASE colonne WHEN valeur …). Le ELSE regroupe Medium et Low dans « Standard ».')
ex('G3', 'Catégorie de budget de chaque projet',
   'trois tranches : moins de 15 000, de 15 000 à 30 000, plus de 30 000. Network Migration (15 000 exactement) tombe dans la tranche intermédiaire.')
ex('G4', 'Nombre d’incidents par catégorie de sévérité créée avec CASE',
   'MySQL autorise le GROUP BY sur l’alias de la colonne calculée. On obtient 3 incidents Standard, 2 Urgent et 2 Important (total 7).')
doc.add_page_break()

# =================================================================== H
H('9. Série H — Sous-requêtes', 1)
ex('H1', 'Employés dont le salaire est supérieur au salaire moyen',
   'la sous-requête scalaire (une seule valeur : 3 883,33 €) est calculée une fois, puis chaque salaire lui est comparé : 5 employés sont au-dessus.',
   intro='La première requête affiche la moyenne, la seconde l’utilise comme sous-requête.')
ex('H2', 'Projets dont le budget est supérieur au budget moyen',
   'le budget moyen est de 21 571,43 € ; trois projets le dépassent.')
ex('H3', 'Employés du même département que Nadia Robert',
   'la sous-requête renvoie le département de Nadia Robert (Cybersecurity). Nadia elle-même est exclue du résultat par la seconde condition : '
   'il reste sa collègue **Ines Garcia**.')
ex('H4', 'Clients ayant au moins un projet',
   'IN teste l’appartenance à la liste des client_id présents dans projets : 5 clients.')
ex('H5', 'Clients sans projet',
   'piège important : la première requête renvoie **Empty set**. La sous-requête contient un NULL (client_id du projet Internal Research) ; '
   'or « id NOT IN (1, 2, …, NULL) » n’est jamais vrai (le résultat est UNKNOWN), donc aucune ligne n’est retournée. '
   'Il faut exclure les NULL dans la sous-requête (seconde requête) — ou utiliser NOT EXISTS, qui n’a pas ce défaut (voir I2).',
   intro='La requête est écrite deux fois : d’abord « naïvement », puis en excluant les NULL de la sous-requête.')
doc.add_page_break()

# =================================================================== I
H('10. Série I — EXISTS / NOT EXISTS', 1)
P('EXISTS est vrai dès que la sous-requête corrélée renvoie au moins une ligne ; le contenu du SELECT n’a pas d’importance (d’où SELECT 1).')
ex('I1', 'Clients ayant au moins un projet (EXISTS)', 'même résultat que H4 : 5 clients, chacun une seule fois.')
ex('I2', 'Clients sans projet (NOT EXISTS)',
   'contrairement au NOT IN de H5, NOT EXISTS donne directement le bon résultat malgré le client_id NULL : '
   'la sous-requête corrélée ne renvoie simplement aucune ligne pour NoProjectClient.')
ex('I3', 'Employés affectés à au moins un projet (EXISTS)',
   '7 employés ; Lucas Bernard et Adam Faure sont absents. Nadia Robert, affectée à trois projets, n’apparaît qu’une fois.')
ex('I4', 'Réécriture de I1 avec un JOIN, puis comparaison',
   'avec un JOIN, chaque client est répété autant de fois qu’il a de projets : **CyberSolutions apparaît deux fois** (6 lignes). '
   'Il faut ajouter DISTINCT pour retrouver le résultat de I1. Comparaison : EXISTS exprime directement la question « existe-t-il… ? », '
   'ne crée pas de doublons et s’arrête à la première ligne trouvée ; le JOIN est préférable lorsqu’on a besoin d’afficher des colonnes '
   'de la seconde table. IN convient pour une liste simple, mais NOT IN est dangereux en présence de NULL.',
   intro='La première requête est le JOIN seul, la seconde ajoute DISTINCT.')
doc.add_page_break()

# =================================================================== J
H('11. Série J — CTE (WITH)', 1)
P('Une CTE (Common Table Expression) nomme un résultat intermédiaire, utilisable ensuite comme une table dans la requête principale.')
ex('J1', 'CTE du salaire moyen par département, puis départements au-dessus de 4 000 €',
   'même résultat que F3, mais la logique est découpée en deux étapes lisibles : le calcul (CTE moyennes) puis le filtre. '
   'Comme la moyenne est déjà une colonne de la CTE, un simple WHERE suffit à la place du HAVING.')
ex('J2', 'CTE du total d’heures par projet, puis projets de plus de 150 heures',
   '4 projets dépassent 150 heures. Data Platform (140 h) et Server Security (70 h) sont exclus.')
ex('J3', 'Deux CTE combinées dans une requête finale',
   'deux agrégations indépendantes (heures par projet, incidents par projet) sont calculées séparément puis rattachées à projets par LEFT JOIN. '
   'Les agréger séparément évite de multiplier les lignes : en joignant directement affectations et incidents à projets, Network Migration '
   '(2 affectations × 2 incidents) aurait produit 4 lignes et des totaux faux.')
doc.add_page_break()

# =================================================================== K
H('12. Série K — Vues (VIEW)', 1)
P('Une vue est une requête enregistrée sous un nom : elle ne stocke pas de données et s’interroge comme une table.')
ex('K1', 'Création de la vue vue_projets_clients',
   'la vue est créée avec un LEFT JOIN pour conserver le projet sans client (Internal Research, client et ville à NULL). '
   'Elle apparaît dans SHOW FULL TABLES avec le type VIEW.')
ex('K2', 'Interrogation de la vue avec WHERE et ORDER BY',
   'on filtre et on trie sur les colonnes de la vue (alias projet, budget, statut…) sans réécrire la jointure.')
ex('K3', 'Projets de plus de 20 000 € à travers la vue', '3 projets dépassent 20 000 €.')
ex('K4', 'Suppression de la vue',
   'après DROP VIEW, la vue n’existe plus (Empty set, puis erreur 1146 si on l’interroge). Les tables projets et clients, elles, ne sont pas modifiées : '
   'supprimer une vue ne supprime aucune donnée.')
doc.add_page_break()

# =================================================================== L
H('13. Série L — Requêtes SQL appliquées à la cybersécurité', 1)
ex('L1', 'Incidents Critical avec adresse IP, date, employé et projet',
   'LEFT JOIN pour ne perdre aucun incident critique, même non attribué. Deux incidents critiques : Ransomware Alert et SSH Brute Force.')
ex('L2', 'Nombre d’incidents par sévérité', '2 Critical, 2 High, 2 Medium et 1 Low.')
ex('L3', 'Nombre d’incidents par projet, y compris les projets sans incident',
   'COUNT(i.id) donne 0 pour Server Security et Internal Research. Network Migration est le projet le plus touché (2 incidents). '
   'L’incident n° 6, rattaché à aucun projet, n’est compté nulle part.')
ex('L4', 'Employés associés à au moins un incident Critical',
   'EXISTS évite les doublons si un employé avait plusieurs incidents critiques : Ines Garcia et Nadia Robert.')
ex('L5', 'Projets n’ayant aucun incident', 'NOT EXISTS : Server Security et Internal Research.')
ex('L6', 'Incidents dont l’adresse IP commence par 192.168.',
   'LIKE \'192.168.%\' : le point final du motif évite de capturer une adresse comme 192.1680.x. 3 incidents proviennent du réseau local 192.168.1.0/24.')
ex('L7', 'Les trois incidents les plus récents', 'tri décroissant sur la date puis LIMIT 3.')
ex('L8', 'Mini tableau de bord',
   'en MySQL, une comparaison vaut 1 (vrai) ou 0 (faux) : SUM(severite = \'Critical\') compte donc les incidents critiques en une seule lecture de la table. '
   '7 incidents, 2 Critical, 2 High, 7 adresses IP distinctes (chaque incident vient d’une IP différente).')
doc.add_page_break()

# =================================================================== Terminal
H('14. MySQL dans le terminal — commandes utiles', 1)
grid(['Commande', 'Rôle'],
     [['mysql -u root -p', 'Ouvre une session interactive (demande le mot de passe)'],
      ['mysql -u root -p entreprise_avancee', 'Ouvre une session directement dans la base indiquée'],
      ['SHOW DATABASES; / USE base;', 'Liste les bases / sélectionne la base courante'],
      ['SHOW TABLES;', 'Liste les tables de la base courante'],
      ['DESCRIBE employes;', 'Affiche les colonnes, types, clés et valeurs par défaut d’une table'],
      ['SHOW CREATE TABLE employes;', 'Affiche l’ordre CREATE TABLE complet (clés étrangères, moteur, jeu de caractères)'],
      ['SELECT VERSION();', 'Affiche la version du serveur'],
      ['EXIT;', 'Ferme la session']], widths=[6.5, 11])
term('terminal', 'SQL exécuté et sortie')
term('showcreate', 'SHOW CREATE TABLE (affichage vertical avec \\G)')
answer('DESCRIBE montre « MUL » pour departement_id et manager_id : MySQL a créé automatiquement un index sur chaque clé étrangère. '
       'SHOW CREATE TABLE confirme les deux contraintes FOREIGN KEY, le moteur InnoDB et AUTO_INCREMENT=10 (9 employés insérés).')
H('Défi terminal — exécuter une requête sans entrer dans l’invite MySQL', 2)
r = subprocess.run([BIN + r'\mysql.exe'] + CONN + ['-t', '-D', DB, '-e', 'SELECT COUNT(*) FROM employes;'], capture_output=True)
block('PS C:\\Users\\tomga> mysql -u root -D entreprise_avancee -e "SELECT COUNT(*) FROM employes;"\n'
      + r.stdout.decode('utf-8').replace('\r', '').strip('\n') + '\nPS C:\\Users\\tomga>', 'term', 'Commande et sortie')
answer('l’option **-D** sélectionne la base et **-e** exécute la requête passée en argument, puis le client rend la main au shell : '
       'on n’entre jamais dans l’invite mysql>. C’est ce qui permet d’utiliser MySQL dans un script Bash, une tâche planifiée ou un pipeline '
       '(par exemple … -e "…" > resultat.txt). Le sujet ajoute -p car son compte root a un mot de passe ; le compte root du serveur de TP n’en a pas.')

# =================================================================== Défi final (paysage)
s2 = doc.add_section(WD_SECTION.NEW_PAGE)
s2.orientation = WD_ORIENT.LANDSCAPE
s2.page_width, s2.page_height = s2.page_height, s2.page_width
WIDTH_PT[0] = 664
H('15. Défi final — Tableau de bord de réponse aux incidents', 1)
P('Une seule requête produit un rapport directement exploitable par une équipe SOC : pour chaque incident récent, qui est concerné, '
  'dans quel département, sur quel projet, pour quel client, avec quel budget en jeu et quel niveau d’urgence.')
TITLES['final'] = 'Défi final — tableau de bord de réponse aux incidents'
block(L['final'][0]['cmd'] + ';', 'sql', 'SQL exécuté')
block(L['final'][0]['out'], 'term', 'Sortie')
grid(['Contrainte du sujet', 'Comment elle est respectée'],
     [['Au moins 4 tables', '5 tables : incidents, employes, departements, projets, clients'],
      ['Au moins 3 JOIN, dont un LEFT JOIN', '4 LEFT JOIN, pour ne perdre aucun incident (l’incident n° 6, non attribué, reste visible)'],
      ['CASE', 'Colonne categorie_severite (Urgent / Important / Standard) et ordre de tri personnalisé'],
      ['Filtrer les incidents récents', 'WHERE date_incident >= date du dernier incident − 5 jours : 5 incidents sur 7 retenus'],
      ['Tri du plus critique au moins critique, puis date décroissante',
       'ORDER BY CASE severite (Critical = 1, High = 2, Medium = 3, autre = 4), date_incident DESC'],
      ['Pas de SELECT *, alias lisibles', 'Colonnes nommées une à une ; alias i, e, d, p, c et alias de colonnes en clair']],
     widths=[8, 16])
answer('le filtre « récent » est calculé par rapport à la date du dernier incident enregistré plutôt qu’à NOW() : le rapport reste ainsi reproductible '
       'quelle que soit la date d’exécution (en production, on écrirait NOW() − INTERVAL 5 DAY). Un tri alphabétique sur severite aurait donné '
       'Critical, High, Low, Medium : d’où le CASE dans l’ORDER BY. Les COALESCE remplacent les NULL des incidents non attribués par des libellés '
       'explicites (« Non attribue », « Hors projet »). Les deux incidents les plus anciens (Port Scan du 20/09 et SSH Brute Force du 21/09) '
       'sortent de la fenêtre de 5 jours.')

s3 = doc.add_section(WD_SECTION.NEW_PAGE)
s3.orientation = WD_ORIENT.PORTRAIT
s3.page_width, s3.page_height = s3.page_height, s3.page_width
WIDTH_PT[0] = 484

# =================================================================== Bonus
H('16. Défi bonus — Détection des données orphelines', 1)
bonus = [('O1', 'Employés sans département', 'Adam Faure : sa clé étrangère departement_id est NULL.'),
         ('O2', 'Clients sans projet', 'NoProjectClient : aucune ligne de projets ne le référence (la donnée liée est inexistante).'),
         ('O3', 'Projets sans client', 'Internal Research : client_id est NULL (projet interne).'),
         ('O4', 'Projets sans affectation', 'Internal Research : aucune ligne dans affectations.'),
         ('O5', 'Incidents sans employé ou sans projet', 'l’incident n° 6 (Suspicious Login) n’est rattaché à rien : il reste à qualifier par le SOC.')]
for i, (k, t, o) in enumerate(bonus, 1):
    TITLES[k] = t
    H(f'16.{i} {t}', 2)
    block(L[k][0]['cmd'] + ';', 'sql', 'SQL exécuté')
    block(L[k][0]['out'], 'term', 'Sortie')
    answer(o)
H('16.6 Différence entre une donnée NULL et une donnée inexistante', 2)
TITLES['O6'] = 'Illustration : comportement de NULL'
block(L['O6'][0]['cmd'] + ';', 'sql', 'SQL exécuté')
block(L['O6'][0]['out'], 'term', 'Sortie')
answer('une donnée **NULL** est une ligne qui existe, mais dont une colonne n’a pas de valeur (inconnue ou sans objet) : Adam Faure existe dans employes, '
       'son departement_id est NULL. On la détecte avec **IS NULL** sur la colonne (16.1, 16.3, 16.5). '
       'Une donnée **inexistante** est une ligne qui n’existe pas du tout dans l’autre table : il n’y a aucune ligne de projets pour NoProjectClient. '
       'Il n’y a rien à tester directement ; on la détecte par l’absence de correspondance : LEFT JOIN … WHERE … IS NULL ou NOT EXISTS (16.2, 16.4). '
       'Le NULL que l’on voit alors dans le résultat est fabriqué par la jointure externe, il n’est stocké nulle part. '
       'La sortie ci-dessus rappelle enfin que NULL n’est égal à rien, pas même à lui-même (NULL = NULL renvoie NULL, seul IS NULL renvoie 1) '
       'et que COUNT(colonne) ignore les NULL (9 lignes mais 4 manager_id renseignés).')
doc.add_page_break()

# =================================================================== Conceptuel
H('17. Réponses aux questions conceptuelles sur les jointures', 1)
P('Synthèse sur l’exemple clients / projets (6 clients dont un sans projet, 7 projets dont un sans client) :')
grid(['Jointure', 'Lignes conservées', 'Résultat sur clients / projets', 'Usage typique'],
     [['INNER JOIN', 'Uniquement les lignes ayant une correspondance des deux côtés', '6 lignes (A2)',
       'Données complètes uniquement ; les lignes orphelines disparaissent'],
      ['LEFT JOIN', 'Toutes les lignes de gauche + correspondances ; NULL à droite sinon', '7 lignes (B3) : + NoProjectClient',
       'Liste exhaustive d’une table principale, comptages incluant les zéros, anti-jointure'],
      ['RIGHT JOIN', 'Toutes les lignes de droite + correspondances ; NULL à gauche sinon', '7 lignes : + Internal Research',
       'Symétrique du LEFT JOIN ; se réécrit en inversant les tables'],
      ['FULL OUTER JOIN', 'Toutes les lignes des deux tables', '8 lignes (D1) : les deux orphelins',
       'Rapprochement / audit de deux tables ; en MySQL : LEFT JOIN UNION RIGHT JOIN'],
      ['SELF JOIN', 'Une table jointe à elle-même avec deux alias', '— (employes / manager, série E)', 'Hiérarchies'],
      ['CROSS JOIN', 'Toutes les combinaisons (produit cartésien)', '6 × 7 = 42 lignes', 'Génération de combinaisons ; à manier avec prudence']],
     widths=[2.8, 5.2, 4.3, 5.2])
bullet('**Lignes qui disparaissent** : avec un INNER JOIN, toute ligne dont la clé étrangère est NULL ou sans correspondance est éliminée '
       '(Adam Faure en A1, Internal Research en A2, l’incident n° 6 en A4).')
bullet('**Lignes qui apparaissent avec des NULL** : une jointure externe conserve ces lignes et complète par NULL les colonnes de l’autre table (B1, B3, C2).')
bullet('**Lignes qui se répètent** : une jointure répète la ligne autant de fois qu’elle a de correspondances (CyberSolutions, deux projets, en B3 et I4).')
bullet('**ON ou WHERE** : après un LEFT JOIN, une condition sur la table de droite placée dans le WHERE transforme la jointure en INNER JOIN (défi de la série B).')
bullet('**NULL** : NULL ne se compare qu’avec IS NULL / IS NOT NULL ; COUNT(colonne) l’ignore ; NOT IN est faussé dès que la sous-requête contient un NULL (H5).')

# =================================================================== Conclusion
H('18. Conclusion générale', 1)
bullet('**Choix de la jointure** : la question à se poser est « quelles lignes doivent absolument rester dans le résultat ? ». '
       'INNER JOIN pour les seules correspondances, LEFT JOIN dès qu’une table doit être conservée entièrement.')
bullet('**FULL OUTER JOIN** : absent de MySQL, il se simule par LEFT JOIN UNION RIGHT JOIN ; UNION élimine les correspondances produites deux fois.')
bullet('**SELF JOIN** : deux alias d’une même table suffisent à représenter une hiérarchie employé / manager.')
bullet('**Agrégation** : WHERE filtre les lignes, HAVING filtre les groupes ; avec un LEFT JOIN on compte une colonne de la table de droite, jamais COUNT(*).')
bullet('**CASE** : il crée des catégories à l’affichage, sert de clé de regroupement et permet un ordre de tri métier.')
bullet('**Sous-requêtes, EXISTS, CTE, vues** : quatre façons de structurer une requête complexe. EXISTS / NOT EXISTS est la forme la plus sûre pour tester '
       'une existence ; la CTE rend les étapes lisibles ; la vue rend une requête réutilisable sans dupliquer les données.')
bullet('**Cybersécurité** : dans un rapport d’incidents, les jointures externes sont indispensables — un incident non attribué ne doit jamais '
       'disparaître du tableau de bord à cause d’un INNER JOIN.')

doc.save(OUT)

# =================================================================== Fichier SQL
order = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
names = {'A': 'INNER JOIN', 'B': 'LEFT JOIN et NULL', 'C': 'RIGHT JOIN', 'D': 'Simulation de FULL OUTER JOIN', 'E': 'SELF JOIN',
         'F': 'JOIN + GROUP BY + HAVING', 'G': 'CASE', 'H': 'Sous-requêtes', 'I': 'EXISTS / NOT EXISTS', 'J': 'CTE (WITH)',
         'K': 'Vues (VIEW)', 'L': 'Requêtes appliquées à la cybersécurité'}
TITLES['Bdefi'] = 'Défi : filtre dans le WHERE puis dans le ON'
o = ['-- ============================================================',
     '-- TP de SQL avancé / MySQL — Tom GALVAN (ECE, B3)',
     '-- Requêtes des séries A à L, défi final et défi bonus',
     '-- Testé sur MySQL Server 8.4.9',
     '-- ============================================================', '', 'USE entreprise_avancee;', '']
for s in order:
    o += ['', '-- ------------------------------------------------------------', f'-- Série {s} — {names[s]}',
          '-- ------------------------------------------------------------']
    for k in S:
        if k[0] == s and (k[1:].isdigit() or k == s + 'defi'):
            o += ['', f"-- {k.replace('defi', ' (défi)')}. {TITLES[k]}"]
            for q in S[k]:
                if k == 'K4' and q.startswith('SELECT'):
                    continue
                o += [q + ';', '']
o += ['', '-- ------------------------------------------------------------', '-- Défi final — Tableau de bord de réponse aux incidents',
      '-- ------------------------------------------------------------', '', S['final'][0] + ';', '',
      '', '-- ------------------------------------------------------------', '-- Défi bonus — Détection des données orphelines',
      '-- ------------------------------------------------------------']
for k in ['O1', 'O2', 'O3', 'O4', 'O5']:
    o += ['', f'-- {TITLES[k]}', S[k][0] + ';']
open(OUT_SQL, 'w', encoding='utf-8', newline='\n').write(re.sub(r'\n{3,}', '\n\n', '\n'.join(o)) + '\n')
print('OK', OUT, OUT_SQL)

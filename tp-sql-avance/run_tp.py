"""Exécute toutes les requêtes du TP SQL avancé sur MySQL 8.4.9 (port 3307) et journalise les sorties réelles."""
import subprocess, re, json

BIN = r'C:\Users\tomga\tp2_work\mysql\mysql-8.4.9-winx64\bin'
CONN = ['-h', '127.0.0.1', '-P', '3307', '-u', 'root']
DB = 'entreprise_avancee'
ECHO = re.compile(r'-{14}\n.*?\n-{14}\n\n?', re.S)

S = {}   # clé -> liste d'instructions SQL

S['demarrage'] = ["SHOW DATABASES", "CREATE DATABASE IF NOT EXISTS entreprise_avancee", "USE entreprise_avancee",
                  "SELECT DATABASE()"]

S['tables'] = ["""CREATE TABLE departements (
    id INT PRIMARY KEY AUTO_INCREMENT,
    nom VARCHAR(80) NOT NULL,
    ville VARCHAR(50) NOT NULL
)""", """CREATE TABLE employes (
    id INT PRIMARY KEY AUTO_INCREMENT,
    nom VARCHAR(60) NOT NULL,
    prenom VARCHAR(60) NOT NULL,
    poste VARCHAR(80),
    salaire DECIMAL(10,2),
    departement_id INT NULL,
    manager_id INT NULL,
    date_embauche DATE,
    FOREIGN KEY (departement_id) REFERENCES departements(id),
    FOREIGN KEY (manager_id) REFERENCES employes(id)
)""", """CREATE TABLE clients (
    id INT PRIMARY KEY AUTO_INCREMENT,
    nom VARCHAR(100) NOT NULL,
    ville VARCHAR(60),
    secteur VARCHAR(60)
)""", """CREATE TABLE projets (
    id INT PRIMARY KEY AUTO_INCREMENT,
    nom VARCHAR(120) NOT NULL,
    budget DECIMAL(12,2),
    client_id INT NULL,
    statut VARCHAR(30),
    FOREIGN KEY (client_id) REFERENCES clients(id)
)""", """CREATE TABLE affectations (
    employe_id INT,
    projet_id INT,
    heures INT,
    role_projet VARCHAR(60),
    PRIMARY KEY (employe_id, projet_id),
    FOREIGN KEY (employe_id) REFERENCES employes(id),
    FOREIGN KEY (projet_id) REFERENCES projets(id)
)""", """CREATE TABLE incidents (
    id INT PRIMARY KEY AUTO_INCREMENT,
    type_incident VARCHAR(100),
    severite VARCHAR(20),
    adresse_ip VARCHAR(45),
    employe_id INT NULL,
    projet_id INT NULL,
    date_incident DATETIME,
    FOREIGN KEY (employe_id) REFERENCES employes(id),
    FOREIGN KEY (projet_id) REFERENCES projets(id)
)""", "SHOW TABLES"]

S['donnees'] = ["""INSERT INTO departements (nom, ville) VALUES
('Networking', 'Paris'),
('Cybersecurity', 'Paris'),
('Development', 'Lyon'),
('Support', 'Lille'),
('Data', 'Bordeaux'),
('Innovation', 'Nantes')""", """INSERT INTO employes
(nom, prenom, poste, salaire, departement_id, manager_id, date_embauche) VALUES
('Martin','Paul','Network Administrator',3600,1,NULL,'2021-09-01'),
('Durand','Sophie','Developer',3900,3,NULL,'2020-04-15'),
('Bernard','Lucas','Technician',2850,4,1,'2024-01-10'),
('Robert','Nadia','Cybersecurity Engineer',4700,2,NULL,'2021-06-20'),
('Petit','Karim','Developer',3500,3,2,'2024-03-01'),
('Garcia','Ines','SOC Analyst',4100,2,4,'2023-11-12'),
('Moreau','Yanis','Network Engineer',4300,1,1,'2022-05-18'),
('Roux','Lina','Data Analyst',4200,5,NULL,'2022-10-03'),
('Faure','Adam','Consultant',3800,NULL,NULL,'2025-01-15')""", """INSERT INTO clients (nom, ville, secteur) VALUES
('TechCorp','Paris','Industry'),
('CyberSolutions','Lyon','Cybersecurity'),
('DataPro','Lille','Data'),
('FinSecure','Nantes','Finance'),
('Medialab','Bordeaux','Healthcare'),
('NoProjectClient','Nice','Consulting')""", """INSERT INTO projets (nom, budget, client_id, statut) VALUES
('Network Migration',15000,1,'Completed'),
('Cybersecurity Audit',25000,2,'In Progress'),
('Web Application',18000,3,'In Progress'),
('Server Security',12000,2,'Completed'),
('SOC Deployment',32000,4,'In Progress'),
('Data Platform',40000,5,'Planned'),
('Internal Research',9000,NULL,'Planned')""", """INSERT INTO affectations VALUES
(1,1,80,'Lead'),
(7,1,100,'Network Engineer'),
(4,2,90,'Security Lead'),
(6,2,120,'SOC Analyst'),
(2,3,130,'Developer'),
(5,3,100,'Developer'),
(4,4,70,'Security Engineer'),
(6,5,160,'SOC Analyst'),
(4,5,110,'Security Lead'),
(8,6,140,'Data Analyst')""", """INSERT INTO incidents
(type_incident,severite,adresse_ip,employe_id,projet_id,date_incident) VALUES
('Port Scan','Low','192.168.1.20',1,1,'2026-09-20 09:15:00'),
('SSH Brute Force','Critical','192.168.1.50',6,2,'2026-09-21 11:42:00'),
('SQL Injection Attempt','High','10.0.0.25',2,3,'2026-09-22 14:05:00'),
('Network Scan','Medium','192.168.1.30',7,1,'2026-09-23 16:10:00'),
('Ransomware Alert','Critical','10.10.2.40',4,5,'2026-09-24 08:20:00'),
('Suspicious Login','Medium','172.16.5.10',NULL,NULL,'2026-09-25 22:45:00'),
('Malware Detected','High','10.10.4.15',8,6,'2026-09-26 13:30:00')"""]

S['verif'] = ["""SELECT 'departements' AS table_name, COUNT(*) AS nb FROM departements
UNION ALL SELECT 'employes', COUNT(*) FROM employes
UNION ALL SELECT 'clients', COUNT(*) FROM clients
UNION ALL SELECT 'projets', COUNT(*) FROM projets
UNION ALL SELECT 'affectations', COUNT(*) FROM affectations
UNION ALL SELECT 'incidents', COUNT(*) FROM incidents"""]

# ------------------------------------------------------------------ Série A
S['A1'] = ["""SELECT e.prenom, e.nom, d.nom AS departement
FROM employes e
INNER JOIN departements d ON e.departement_id = d.id
ORDER BY e.id"""]
S['A2'] = ["""SELECT p.nom AS projet, c.nom AS client
FROM projets p
INNER JOIN clients c ON p.client_id = c.id
ORDER BY p.id"""]
S['A3'] = ["""SELECT CONCAT(e.prenom, ' ', e.nom) AS employe, p.nom AS projet,
       a.role_projet, a.heures
FROM affectations a
INNER JOIN employes e ON a.employe_id = e.id
INNER JOIN projets p ON a.projet_id = p.id
ORDER BY p.id, e.id"""]
S['A4'] = ["""SELECT i.id, i.type_incident, i.severite,
       CONCAT(e.prenom, ' ', e.nom) AS employe, p.nom AS projet
FROM incidents i
INNER JOIN employes e ON i.employe_id = e.id
INNER JOIN projets p ON i.projet_id = p.id
ORDER BY i.id"""]
S['A5'] = ["""SELECT CONCAT(e.prenom, ' ', e.nom) AS employe, e.poste,
       p.nom AS projet, p.statut, a.heures
FROM employes e
INNER JOIN affectations a ON a.employe_id = e.id
INNER JOIN projets p ON p.id = a.projet_id
WHERE p.statut = 'In Progress'
ORDER BY p.nom, a.heures DESC"""]

# ------------------------------------------------------------------ Série B
S['B1'] = ["""SELECT e.prenom, e.nom, d.nom AS departement
FROM employes e
LEFT JOIN departements d ON e.departement_id = d.id
ORDER BY e.id"""]
S['B2'] = ["""SELECT e.id, e.prenom, e.nom, e.poste
FROM employes e
LEFT JOIN affectations a ON a.employe_id = e.id
WHERE a.employe_id IS NULL"""]
S['B3'] = ["""SELECT c.nom AS client, p.nom AS projet, p.statut
FROM clients c
LEFT JOIN projets p ON p.client_id = c.id
ORDER BY c.id, p.id"""]
S['B4'] = ["""SELECT c.id, c.nom AS client, c.ville
FROM clients c
LEFT JOIN projets p ON p.client_id = c.id
WHERE p.id IS NULL"""]
S['B5'] = ["""SELECT p.nom AS projet, COUNT(a.employe_id) AS nb_employes
FROM projets p
LEFT JOIN affectations a ON a.projet_id = p.id
GROUP BY p.id, p.nom
ORDER BY p.id"""]
S['B6'] = ["""SELECT d.nom AS departement, COUNT(e.id) AS nb_employes
FROM departements d
LEFT JOIN employes e ON e.departement_id = d.id
GROUP BY d.id, d.nom
ORDER BY d.id"""]
S['Bdefi'] = ["""SELECT c.nom AS client, p.nom AS projet, p.statut
FROM clients c
LEFT JOIN projets p ON p.client_id = c.id
WHERE p.statut = 'In Progress'
ORDER BY c.id""", """SELECT c.nom AS client, p.nom AS projet, p.statut
FROM clients c
LEFT JOIN projets p ON p.client_id = c.id AND p.statut = 'In Progress'
ORDER BY c.id"""]

# ------------------------------------------------------------------ Série C
S['C1'] = ["""SELECT c.nom AS client, p.nom AS projet, p.statut
FROM projets p
RIGHT JOIN clients c ON p.client_id = c.id
ORDER BY c.id, p.id"""]
S['C2'] = ["""SELECT d.nom AS departement, e.prenom, e.nom
FROM employes e
RIGHT JOIN departements d ON e.departement_id = d.id
ORDER BY d.id, e.id"""]
S['C3'] = ["""SELECT d.nom AS departement, e.prenom, e.nom
FROM departements d
LEFT JOIN employes e ON e.departement_id = d.id
ORDER BY d.id, e.id"""]

# ------------------------------------------------------------------ Série D
FULL = """SELECT c.id AS client_id, c.nom AS client, p.id AS projet_id, p.nom AS projet
    FROM clients c
    LEFT JOIN projets p ON p.client_id = c.id
    UNION
    SELECT c.id, c.nom, p.id, p.nom
    FROM clients c
    RIGHT JOIN projets p ON p.client_id = c.id"""
S['D1'] = ["""SELECT c.id AS client_id, c.nom AS client, p.id AS projet_id, p.nom AS projet
FROM clients c
LEFT JOIN projets p ON p.client_id = c.id
UNION
SELECT c.id, c.nom, p.id, p.nom
FROM clients c
RIGHT JOIN projets p ON p.client_id = c.id"""]
S['D2'] = [f"""SELECT fj.client_id, fj.client, fj.projet
FROM (
    {FULL}
) AS fj
WHERE fj.projet_id IS NULL"""]
S['D3'] = [f"""SELECT fj.client, fj.projet_id, fj.projet
FROM (
    {FULL}
) AS fj
WHERE fj.client_id IS NULL"""]
S['D4'] = ["""SELECT COUNT(*) AS lignes_avec_UNION FROM (
    SELECT c.id AS cid, p.id AS pid FROM clients c LEFT JOIN projets p ON p.client_id = c.id
    UNION
    SELECT c.id, p.id FROM clients c RIGHT JOIN projets p ON p.client_id = c.id
) AS t""", """SELECT COUNT(*) AS lignes_avec_UNION_ALL FROM (
    SELECT c.id AS cid, p.id AS pid FROM clients c LEFT JOIN projets p ON p.client_id = c.id
    UNION ALL
    SELECT c.id, p.id FROM clients c RIGHT JOIN projets p ON p.client_id = c.id
) AS t"""]

# ------------------------------------------------------------------ Série E
S['E1'] = ["""SELECT e.prenom, e.nom, m.prenom AS prenom_manager, m.nom AS nom_manager
FROM employes e
LEFT JOIN employes m ON e.manager_id = m.id
ORDER BY e.id"""]
S['E2'] = ["""SELECT e.prenom, e.nom, m.prenom AS prenom_manager, m.nom AS nom_manager
FROM employes e
INNER JOIN employes m ON e.manager_id = m.id
ORDER BY e.id"""]
S['E3'] = ["""SELECT e.id, e.prenom, e.nom, e.poste
FROM employes e
LEFT JOIN employes m ON e.manager_id = m.id
WHERE m.id IS NULL
ORDER BY e.id"""]
S['E4'] = ["""SELECT CONCAT(e.prenom, ' ', e.nom) AS employe, e.poste,
       COALESCE(CONCAT(m.prenom, ' ', m.nom), '(aucun)') AS manager
FROM employes e
LEFT JOIN employes m ON e.manager_id = m.id
ORDER BY manager, employe"""]

# ------------------------------------------------------------------ Série F
S['F1'] = ["""SELECT d.nom AS departement, COUNT(*) AS nb_employes
FROM employes e
INNER JOIN departements d ON e.departement_id = d.id
GROUP BY d.id, d.nom
ORDER BY nb_employes DESC, d.nom"""]
S['F2'] = ["""SELECT d.nom AS departement, ROUND(AVG(e.salaire), 2) AS salaire_moyen
FROM employes e
INNER JOIN departements d ON e.departement_id = d.id
GROUP BY d.id, d.nom
ORDER BY salaire_moyen DESC"""]
S['F3'] = ["""SELECT d.nom AS departement, ROUND(AVG(e.salaire), 2) AS salaire_moyen
FROM employes e
INNER JOIN departements d ON e.departement_id = d.id
GROUP BY d.id, d.nom
HAVING AVG(e.salaire) > 4000
ORDER BY salaire_moyen DESC"""]
S['F4'] = ["""SELECT c.nom AS client, COUNT(p.id) AS nb_projets, SUM(p.budget) AS budget_total
FROM clients c
INNER JOIN projets p ON p.client_id = c.id
GROUP BY c.id, c.nom
ORDER BY budget_total DESC"""]
S['F5'] = ["""SELECT c.nom AS client, SUM(p.budget) AS budget_total
FROM clients c
INNER JOIN projets p ON p.client_id = c.id
GROUP BY c.id, c.nom
HAVING SUM(p.budget) > 30000
ORDER BY budget_total DESC"""]
S['F6'] = ["""SELECT p.nom AS projet, COALESCE(SUM(a.heures), 0) AS heures_totales
FROM projets p
LEFT JOIN affectations a ON a.projet_id = p.id
GROUP BY p.id, p.nom
ORDER BY heures_totales DESC"""]

# ------------------------------------------------------------------ Série G
S['G1'] = ["""SELECT prenom, nom, salaire,
       CASE
           WHEN salaire < 3500 THEN 'Junior'
           WHEN salaire <= 4200 THEN 'Confirmed'
           ELSE 'Senior'
       END AS categorie
FROM employes
ORDER BY salaire"""]
S['G2'] = ["""SELECT id, type_incident, severite,
       CASE severite
           WHEN 'Critical' THEN 'Urgent'
           WHEN 'High' THEN 'Important'
           ELSE 'Standard'
       END AS categorie_severite
FROM incidents
ORDER BY id"""]
S['G3'] = ["""SELECT nom AS projet, budget,
       CASE
           WHEN budget < 15000 THEN '< 15000'
           WHEN budget <= 30000 THEN '15000 - 30000'
           ELSE '> 30000'
       END AS categorie_budget
FROM projets
ORDER BY budget"""]
S['G4'] = ["""SELECT CASE severite
           WHEN 'Critical' THEN 'Urgent'
           WHEN 'High' THEN 'Important'
           ELSE 'Standard'
       END AS categorie_severite,
       COUNT(*) AS nb_incidents
FROM incidents
GROUP BY categorie_severite
ORDER BY nb_incidents DESC"""]

# ------------------------------------------------------------------ Série H
S['H1'] = ["SELECT ROUND(AVG(salaire), 2) AS salaire_moyen FROM employes", """SELECT prenom, nom, salaire
FROM employes
WHERE salaire > (SELECT AVG(salaire) FROM employes)
ORDER BY salaire DESC"""]
S['H2'] = ["SELECT ROUND(AVG(budget), 2) AS budget_moyen FROM projets", """SELECT nom AS projet, budget
FROM projets
WHERE budget > (SELECT AVG(budget) FROM projets)
ORDER BY budget DESC"""]
S['H3'] = ["""SELECT prenom, nom, poste
FROM employes
WHERE departement_id = (SELECT departement_id
                        FROM employes
                        WHERE prenom = 'Nadia' AND nom = 'Robert')
  AND NOT (prenom = 'Nadia' AND nom = 'Robert')"""]
S['H4'] = ["""SELECT id, nom AS client
FROM clients
WHERE id IN (SELECT client_id FROM projets)
ORDER BY id"""]
S['H5'] = ["""SELECT id, nom AS client
FROM clients
WHERE id NOT IN (SELECT client_id FROM projets)""", """SELECT id, nom AS client
FROM clients
WHERE id NOT IN (SELECT client_id FROM projets WHERE client_id IS NOT NULL)"""]

# ------------------------------------------------------------------ Série I
S['I1'] = ["""SELECT c.id, c.nom AS client
FROM clients c
WHERE EXISTS (SELECT 1 FROM projets p WHERE p.client_id = c.id)
ORDER BY c.id"""]
S['I2'] = ["""SELECT c.id, c.nom AS client
FROM clients c
WHERE NOT EXISTS (SELECT 1 FROM projets p WHERE p.client_id = c.id)"""]
S['I3'] = ["""SELECT e.id, e.prenom, e.nom
FROM employes e
WHERE EXISTS (SELECT 1 FROM affectations a WHERE a.employe_id = e.id)
ORDER BY e.id"""]
S['I4'] = ["""SELECT c.id, c.nom AS client
FROM clients c
INNER JOIN projets p ON p.client_id = c.id
ORDER BY c.id""", """SELECT DISTINCT c.id, c.nom AS client
FROM clients c
INNER JOIN projets p ON p.client_id = c.id
ORDER BY c.id"""]

# ------------------------------------------------------------------ Série J
S['J1'] = ["""WITH moyennes AS (
    SELECT d.nom AS departement, AVG(e.salaire) AS salaire_moyen
    FROM departements d
    INNER JOIN employes e ON e.departement_id = d.id
    GROUP BY d.id, d.nom
)
SELECT departement, ROUND(salaire_moyen, 2) AS salaire_moyen
FROM moyennes
WHERE salaire_moyen > 4000
ORDER BY salaire_moyen DESC"""]
S['J2'] = ["""WITH heures_projet AS (
    SELECT projet_id, SUM(heures) AS heures_totales
    FROM affectations
    GROUP BY projet_id
)
SELECT p.nom AS projet, h.heures_totales
FROM heures_projet h
INNER JOIN projets p ON p.id = h.projet_id
WHERE h.heures_totales > 150
ORDER BY h.heures_totales DESC"""]
S['J3'] = ["""WITH heures_projet AS (
    SELECT projet_id, SUM(heures) AS heures_totales, COUNT(*) AS nb_employes
    FROM affectations
    GROUP BY projet_id
),
incidents_projet AS (
    SELECT projet_id, COUNT(*) AS nb_incidents
    FROM incidents
    WHERE projet_id IS NOT NULL
    GROUP BY projet_id
)
SELECT p.nom AS projet, p.statut,
       COALESCE(h.nb_employes, 0) AS nb_employes,
       COALESCE(h.heures_totales, 0) AS heures_totales,
       COALESCE(i.nb_incidents, 0) AS nb_incidents
FROM projets p
LEFT JOIN heures_projet h ON h.projet_id = p.id
LEFT JOIN incidents_projet i ON i.projet_id = p.id
ORDER BY nb_incidents DESC, heures_totales DESC"""]

# ------------------------------------------------------------------ Série K
S['K1'] = ["""CREATE VIEW vue_projets_clients AS
SELECT p.nom AS projet, p.budget, p.statut, c.nom AS client, c.ville
FROM projets p
LEFT JOIN clients c ON p.client_id = c.id""", "SHOW FULL TABLES WHERE Table_type = 'VIEW'",
           "SELECT * FROM vue_projets_clients"]
S['K2'] = ["""SELECT projet, budget, client, ville
FROM vue_projets_clients
WHERE statut = 'In Progress'
ORDER BY budget DESC"""]
S['K3'] = ["""SELECT projet, budget, statut, client
FROM vue_projets_clients
WHERE budget > 20000
ORDER BY budget DESC"""]
S['K4'] = ["DROP VIEW vue_projets_clients", "SHOW FULL TABLES WHERE Table_type = 'VIEW'",
           "SELECT * FROM vue_projets_clients"]

# ------------------------------------------------------------------ Série L
S['L1'] = ["""SELECT i.type_incident, i.adresse_ip, i.date_incident,
       CONCAT(e.prenom, ' ', e.nom) AS employe, p.nom AS projet
FROM incidents i
LEFT JOIN employes e ON i.employe_id = e.id
LEFT JOIN projets p ON i.projet_id = p.id
WHERE i.severite = 'Critical'
ORDER BY i.date_incident DESC"""]
S['L2'] = ["""SELECT severite, COUNT(*) AS nb_incidents
FROM incidents
GROUP BY severite
ORDER BY nb_incidents DESC, severite"""]
S['L3'] = ["""SELECT p.nom AS projet, COUNT(i.id) AS nb_incidents
FROM projets p
LEFT JOIN incidents i ON i.projet_id = p.id
GROUP BY p.id, p.nom
ORDER BY nb_incidents DESC, p.id"""]
S['L4'] = ["""SELECT e.id, e.prenom, e.nom, e.poste
FROM employes e
WHERE EXISTS (SELECT 1
              FROM incidents i
              WHERE i.employe_id = e.id AND i.severite = 'Critical')"""]
S['L5'] = ["""SELECT p.id, p.nom AS projet, p.statut
FROM projets p
WHERE NOT EXISTS (SELECT 1 FROM incidents i WHERE i.projet_id = p.id)"""]
S['L6'] = ["""SELECT id, type_incident, severite, adresse_ip
FROM incidents
WHERE adresse_ip LIKE '192.168.%'
ORDER BY id"""]
S['L7'] = ["""SELECT id, type_incident, severite, date_incident
FROM incidents
ORDER BY date_incident DESC
LIMIT 3"""]
S['L8'] = ["""SELECT COUNT(*) AS total_incidents,
       SUM(severite = 'Critical') AS nb_critical,
       SUM(severite = 'High') AS nb_high,
       COUNT(DISTINCT adresse_ip) AS ip_distinctes
FROM incidents"""]

# ------------------------------------------------------------------ Terminal
S['terminal'] = ["SHOW TABLES", "DESCRIBE employes", "SELECT VERSION()"]
S['showcreate'] = ["SHOW CREATE TABLE employes\\G"]

# ------------------------------------------------------------------ Défi final
S['final'] = ["""SELECT i.id AS incident_id,
       i.type_incident,
       i.severite,
       i.adresse_ip,
       i.date_incident,
       COALESCE(CONCAT(e.prenom, ' ', e.nom), 'Non attribue') AS employe,
       COALESCE(d.nom, '-') AS departement,
       COALESCE(p.nom, 'Hors projet') AS projet,
       COALESCE(c.nom, '-') AS client,
       p.budget AS budget_projet,
       CASE i.severite
           WHEN 'Critical' THEN 'Urgent'
           WHEN 'High' THEN 'Important'
           ELSE 'Standard'
       END AS categorie_severite
FROM incidents i
LEFT JOIN employes e ON i.employe_id = e.id
LEFT JOIN departements d ON e.departement_id = d.id
LEFT JOIN projets p ON i.projet_id = p.id
LEFT JOIN clients c ON p.client_id = c.id
WHERE i.date_incident >= (SELECT MAX(date_incident) FROM incidents) - INTERVAL 5 DAY
ORDER BY CASE i.severite
             WHEN 'Critical' THEN 1
             WHEN 'High' THEN 2
             WHEN 'Medium' THEN 3
             ELSE 4
         END,
         i.date_incident DESC"""]

# ------------------------------------------------------------------ Bonus
S['O1'] = ["""SELECT id, prenom, nom, poste
FROM employes
WHERE departement_id IS NULL"""]
S['O2'] = ["""SELECT c.id, c.nom AS client
FROM clients c
LEFT JOIN projets p ON p.client_id = c.id
WHERE p.id IS NULL"""]
S['O3'] = ["""SELECT id, nom AS projet, budget, statut
FROM projets
WHERE client_id IS NULL"""]
S['O4'] = ["""SELECT p.id, p.nom AS projet, p.statut
FROM projets p
LEFT JOIN affectations a ON a.projet_id = p.id
WHERE a.projet_id IS NULL"""]
S['O5'] = ["""SELECT id, type_incident, severite, adresse_ip, employe_id, projet_id
FROM incidents
WHERE employe_id IS NULL OR projet_id IS NULL"""]
S['O6'] = ["SELECT NULL = NULL AS null_egal_null, NULL IS NULL AS null_is_null, COUNT(*) AS lignes, COUNT(manager_id) AS managers_renseignes FROM employes"]


def run(key, stmts, db=DB):
    args = [BIN + r'\mysql.exe'] + CONN + ['-t', '-vv', '-n', '--force', '--default-character-set=utf8mb4']
    if db:
        args.append(db)
    sql = ''.join(s if s.endswith('\\G') else s + ';' for s in (x + '\n' for x in ()) ) if False else \
        '\n'.join(s if s.endswith('\\G') else s + ';' for s in stmts) + '\n'
    r = subprocess.run(args, input=sql.encode('utf-8'), capture_output=True)
    out = r.stdout.decode('utf-8', errors='replace').replace('\r', '')
    err = r.stderr.decode('utf-8', errors='replace').replace('\r', '')
    parts = ECHO.split(out)[1:]
    res = []
    errs = [l for l in err.split('\n') if l.startswith('ERROR')]
    for i, s in enumerate(stmts):
        o = parts[i] if i < len(parts) else ''
        o = '\n'.join(l for l in o.split('\n') if l.strip() != 'Bye').strip('\n')
        if not o.strip() and errs:
            o = re.sub(r' at line \d+', '', errs.pop(0))
        res.append({'cmd': s, 'out': o})
    return res


if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    subprocess.run([BIN + r'\mysql.exe'] + CONN + ['-e', 'DROP DATABASE IF EXISTS entreprise_avancee'], check=True)
    L = {}
    for k, st in S.items():
        L[k] = run(k, st, db=None if k == 'demarrage' else DB)
    # défi terminal : requête lancée directement depuis le shell
    r = subprocess.run([BIN + r'\mysql.exe'] + CONN + ['-D', DB, '-e', 'SELECT COUNT(*) FROM employes;'],
                       capture_output=True)
    L['bash'] = [{'cmd': 'mysql -u root -p -D entreprise_avancee -e "SELECT COUNT(*) FROM employes;"',
                  'out': r.stdout.decode('utf-8').replace('\r', '').strip('\n')}]
    r = subprocess.run([BIN + r'\mysql.exe'] + CONN + ['-t', '-e', 'STATUS'], capture_output=True)
    L['status'] = [{'cmd': 'STATUS', 'out': r.stdout.decode('utf-8', errors='replace').replace('\r', '').strip('\n')}]
    json.dump(L, open(r'C:\Users\tomga\tp_sql_avance_work\log.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    for k, v in L.items():
        if k in ('tables', 'donnees'):
            print('==', k, [e['out'].split('\n')[0] for e in v]); continue
        print('=====', k)
        for e in v:
            print(e['out']); print()

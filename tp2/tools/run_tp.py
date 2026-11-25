import sys, json, time, subprocess
sys.path.insert(0, r'C:\Users\tomga\tp2_work\tools')
from engine import *

SCRIPT = r'C:\Users\tomga\tp2_work\src\script tp2.txt'
LOG = {}
PROJ = 'SELECT id, nom, budget FROM projets'


def save(key, scn):
    LOG[key] = scn.log


if not ping():
    start_server()

# ============================================================ 0. Mise en place
r = subprocess.run([BIN + r'\mysql.exe', '-uroot', '-h', 'localhost', '-t'],
                   stdin=open(SCRIPT, 'rb'), capture_output=True)
render_console('00_script.png', 'Windows PowerShell — exécution du script tp2',
               [('Get-Content "script tp2.txt" | mysql -u root -t', r.stdout.decode('utf-8') + r.stderr.decode('utf-8'))])

adm = Session('root', label='Session root')
sc = Scenario([adm], 'MySQL — vérifications préalables')
sc.run(adm, "SHOW DATABASES LIKE 'tp2%'")
sc.run(adm, "SELECT TABLE_NAME, ENGINE FROM information_schema.TABLES WHERE TABLE_SCHEMA='tp2_mysql'")
sc.run(adm, "SHOW VARIABLES LIKE 'autocommit'")
sc.snap('01_innodb.png')
# table T (parties 5 et 6) + compte user1
for q in ["DROP USER IF EXISTS 'user1'@'localhost', 'U1'@'localhost', 'U2'@'localhost'",
          "CREATE USER 'user1'@'localhost' IDENTIFIED BY 'user1'",
          "GRANT ALL PRIVILEGES ON tp2_mysql.* TO 'user1'@'localhost'"]:
    sc.run(adm, q)
sc.snap('01b_user1.png', 'MySQL — création du compte user1 (utilisé pour les parties 5 et 6)')
save('setup', sc)

# ============================================================ 1. Commit / Rollback
s = Session('s1', label='Connexion unique (root)')
sc = Scenario([s], 'Q1 — COMMIT et ROLLBACK (une seule connexion)')
sc.run(s, 'SET autocommit = 0')
sc.run(s, 'START TRANSACTION')
sc.run(s, "INSERT INTO projets (nom, budget, client_id) VALUES ('Projet test ROLLBACK', 15000, 4)")
sc.run(s, PROJ)
sc.run(s, 'ROLLBACK')
sc.run(s, PROJ)
sc.snap('10_rollback.png', 'Q1-a — INSERT puis ROLLBACK')
sc.run(s, 'START TRANSACTION')
sc.run(s, "INSERT INTO projets (nom, budget, client_id) VALUES ('Projet test COMMIT', 15000, 4)")
sc.run(s, PROJ)
sc.run(s, 'COMMIT')
sc.run(s, PROJ)
sc.snap('11_commit.png', 'Q1-b — INSERT puis COMMIT')
save('q1', sc)
s.close()

# ============================================================ 2. Reprise sur panne (client)
a = Session('a', label='Connexion n°1 (root)')
sc = Scenario([a], 'Q2-a — fermeture de la connexion avec une transaction ouverte')
sc.run(a, 'SET autocommit = 0')
sc.run(a, 'START TRANSACTION')
sc.run(a, "INSERT INTO projets (nom, budget, client_id) VALUES ('Projet avant fermeture', 9000, 4)")
sc.run(a, PROJ)
sc.note('a', ">>> Fermeture de l'onglet de connexion (commande EXIT) sans COMMIT")
a.close()
b = Session('b', label='Nouvelle connexion (root)')
sc.add_session(b)
sc.run(b, PROJ)
sc.snap('20_fermeture.png')
save('q2a', sc)
b.close()

a = Session('a', label='Connexion n°1 (root)')
sc = Scenario([a], 'Q2-b — arrêt brutal du client (processus tué)')
sc.run(a, 'SET autocommit = 0')
sc.run(a, 'START TRANSACTION')
sc.run(a, "INSERT INTO projets (nom, budget, client_id) VALUES ('Projet avant crash client', 9500, 4)")
sc.run(a, PROJ)
pid = a.p.pid
cmd, out = shell(f'taskkill /F /PID {pid}', ['taskkill', '/F', '/PID', str(pid)])
sc.note('a', f'>>> PANNE CLIENT : {cmd}\n{out}')
time.sleep(1.5)
b = Session('b', label='Nouvelle connexion (root)')
sc.add_session(b)
sc.run(b, PROJ)
sc.snap('21_kill_client.png')
save('q2b', sc)
b.close()

# ============================================================ 3. Isolation
p1 = Session('p1', label='Transaction T1 — connexion 1 (root)')
p2 = Session('p2', label='Transaction T2 — connexion 2 (root)')
sc = Scenario([p1, p2], 'Q3 — niveau d’isolation par défaut')
sc.run(p1, "SHOW VARIABLES LIKE '%isolation%'")
sc.snap('30_show_isolation.png')
sc.run(p1, 'SET autocommit = 0')
sc.run(p2, 'SET autocommit = 0')
sc.run(p1, "INSERT INTO projets (nom, budget, client_id) VALUES ('Projet isolation', 30000, 4)")
sc.run(p1, PROJ)
sc.run(p2, PROJ)
sc.run(p1, 'COMMIT')
sc.run(p2, PROJ)
sc.run(p2, 'COMMIT')
sc.run(p2, PROJ)
sc.snap('31_rr_insert.png', 'Q3 — REPEATABLE READ : insertion d’un projet par T1')
sc.run(p1, "DELETE FROM projets WHERE nom = 'Projet isolation'")
sc.run(p1, PROJ)
sc.run(p2, PROJ)
sc.run(p1, 'COMMIT')
sc.run(p2, 'COMMIT')
sc.run(p2, PROJ)
sc.snap('32_rr_delete.png', 'Q3 — REPEATABLE READ : suppression du projet par T1')
save('q3_rr', sc)

sc = Scenario([p1, p2], 'Q3 — READ UNCOMMITTED')
sc.run(p2, 'SET SESSION TRANSACTION ISOLATION LEVEL READ UNCOMMITTED')
sc.run(p2, 'SELECT @@transaction_isolation')
sc.run(p1, "INSERT INTO projets (nom, budget, client_id) VALUES ('Projet non valide', 7777, 4)")
sc.run(p2, PROJ)
sc.run(p1, 'ROLLBACK')
sc.run(p2, PROJ)
sc.run(p2, 'COMMIT')
sc.snap('33_read_uncommitted.png', 'Q3 — READ UNCOMMITTED sur T2 : lecture sale (dirty read)')
save('q3_ru', sc)

sc = Scenario([p1, p2], 'Q3 — SERIALIZABLE')
sc.run(p2, 'SET SESSION TRANSACTION ISOLATION LEVEL SERIALIZABLE')
sc.run(p2, 'SELECT @@transaction_isolation')
sc.run(p1, "INSERT INTO projets (nom, budget, client_id) VALUES ('Projet serialisable', 11111, 4)")
sc.run(p2, PROJ, wait=2.5)
sc.run(p1, 'COMMIT', wait=1.5)
sc.run(p2, 'COMMIT')
sc.snap('34_serializable_a.png', 'Q3 — SERIALIZABLE (1) : T2 veut lire une donnée en cours de modification par T1')
sc.run(p2, PROJ)
sc.run(p1, "UPDATE projets SET budget = 99999 WHERE nom = 'Projet serialisable'", wait=2.5)
sc.run(p2, 'COMMIT', wait=1.5)
sc.run(p1, 'COMMIT')
sc.run(p1, "DELETE FROM projets WHERE nom = 'Projet serialisable'")
sc.run(p1, 'COMMIT')
sc.snap('35_serializable_b.png', 'Q3 — SERIALIZABLE (2) : T1 veut modifier une donnée lue par T2')
save('q3_ser', sc)
p1.close(); p2.close()

# ============================================================ 4. Accès concurrents (pno 203 -> projets.id 3, prix -> budget)
s1 = Session('S1', label='Session 1 (root)')
s2 = Session('S2', label='Session 2 (root)')
B = 'SELECT budget FROM projets WHERE id = 3'
sc = Scenario([s1, s2], 'Accès concurrents — même utilisateur, deux sessions')
sc.run(s1, 'SET autocommit = 0')
sc.run(s2, 'SET autocommit = 0')
sc.run(s1, 'UPDATE projets SET budget = 3000 WHERE id = 3')
sc.run(s2, 'UPDATE projets SET budget = 3500 WHERE id = 3', wait=2.5)
sc.run(s1, B)
sc.run(s1, 'COMMIT', wait=1.5)
sc.run(s2, B)
sc.run(s1, B)
sc.run(s2, 'COMMIT')
sc.run(s1, B)
sc.run(s2, B)
sc.run(s1, 'COMMIT')
sc.run(s1, B)
sc.snap('40_concurrence_update.png', 'Accès concurrents (1) — deux UPDATE sur la même ligne')
sc.run(s1, 'SELECT budget FROM projets WHERE id = 3 FOR UPDATE OF budget')
sc.run(s1, 'SELECT budget FROM projets WHERE id = 3 FOR UPDATE')
sc.run(s2, 'UPDATE projets SET budget = 4000 WHERE id = 3', wait=2.5)
sc.run(s1, 'UPDATE projets SET budget = 5000 WHERE id = 3')
sc.run(s1, 'COMMIT', wait=1.5)
sc.run(s2, B)
sc.run(s2, 'COMMIT')
sc.run(s1, B)
sc.snap('41_concurrence_forupdate.png', 'Accès concurrents (2) — verrouillage explicite SELECT ... FOR UPDATE')
save('concurrence', sc)
s1.close(); s2.close()

# ============================================================ 5. Droits d'accès
root = Session('root', label='Session root')
sc = Scenario([root], 'Droits d’accès — création des utilisateurs par root')
sc.run(root, "CREATE USER 'U1'@'localhost' IDENTIFIED BY 'u1pass'")
sc.run(root, "GRANT ALL PRIVILEGES ON tp2_mysql.* TO 'U1'@'localhost'")
sc.run(root, "CREATE USER 'U2'@'localhost' IDENTIFIED BY 'u2pass'")
sc.run(root, "GRANT SELECT ON tp2_mysql.* TO 'U2'@'localhost'")
sc.run(root, "SHOW GRANTS FOR 'U1'@'localhost'")
sc.run(root, "SHOW GRANTS FOR 'U2'@'localhost'")
sc.snap('50_create_users.png')
u1 = Session('U1', 'U1', 'u1pass', label='Session U1')
u2 = Session('U2', 'U2', 'u2pass', label='Session U2')
sc = Scenario([root, u1, u2], 'Droits d’accès — trois connexions')
sc.run(u1, 'SELECT CURRENT_USER()')
sc.run(u1, 'SELECT id, nom, budget FROM projets WHERE id = 2')
sc.run(u1, 'UPDATE projets SET budget = 13000 WHERE id = 2')
sc.run(u1, 'SELECT id, nom, budget FROM projets WHERE id = 2')
sc.snap('51_u1_select_update.png', 'Droits — étape 2 : SELECT et UPDATE par U1')
sc.run(u2, 'SELECT CURRENT_USER()')
sc.run(u2, 'SELECT id, nom, budget FROM projets WHERE id = 2')
sc.run(u2, 'UPDATE projets SET budget = 14000 WHERE id = 2')
sc.snap('52_u2_select_update.png', 'Droits — étape 3 : SELECT et UPDATE par U2')
sc.run(u1, "GRANT UPDATE ON tp2_mysql.projets TO 'U2'@'localhost'")
sc.snap('53_u1_grant_refuse.png', 'Droits — étape 4 : U1 tente de donner UPDATE à U2')
sc.run(root, "GRANT ALL PRIVILEGES ON tp2_mysql.* TO 'U1'@'localhost' WITH GRANT OPTION")
sc.run(root, "SHOW GRANTS FOR 'U1'@'localhost'")
sc.run(u1, "GRANT UPDATE ON tp2_mysql.projets TO 'U2'@'localhost'")
sc.run(u2, 'UPDATE projets SET budget = 14000 WHERE id = 2')
sc.run(u2, 'SELECT id, nom, budget FROM projets WHERE id = 2')
sc.snap('54_grant_option.png', 'Droits — étape 5 : WITH GRANT OPTION puis nouvelle tentative')
sc.run(u1, 'SET autocommit = 0')
sc.run(u1, "INSERT INTO projets (nom, budget, client_id) VALUES ('Projet cree par U1', 21000, 4)")
sc.run(u1, PROJ)
sc.run(root, PROJ)
sc.run(u2, PROJ)
sc.run(u1, 'COMMIT')
sc.run(root, PROJ)
sc.run(u2, PROJ)
sc.snap('55_u1_insert.png', 'Droits — étape 6 : insertion par U1, visibilité pour root et U2')
save('droits', sc)
u1.close(); u2.close()


# ============================================================ 6. Tables T — reprise sur panne
def reset_T(sess):
    for q in ['DROP TABLE IF EXISTS T', 'CREATE TABLE T (A INTEGER, B INTEGER) ENGINE=InnoDB',
              'INSERT INTO T VALUES (1,2)']:
        sess.send(q)
    sess.send('COMMIT')
    time.sleep(0.8)
    sess.drain()


def crash_and_restart(sc, who):
    cmd, out = shell('taskkill /F /IM mysqld.exe', ['taskkill', '/F', '/IM', 'mysqld.exe'])
    sc.note(who, f'>>> PANNE SERVEUR : {cmd}\n{out}')
    time.sleep(2)
    start_server()
    sc.note(who, '>>> Redémarrage du serveur MySQL (mysqld) OK', new_row=False)
    return cmd, out


# tableau 1
reset_T(root)
p1 = Session('p1', label='p1 (root) — T dans trans p1')
p2 = Session('p2', 'user1', 'user1', label='p2 (user1) — T vue par p2')
sc = Scenario([p1, p2], 'Panne — tableau 1')
sc.run(p1, 'SET autocommit = 0', observe=[p1, p2])
for q in ['INSERT INTO T VALUES (3,4)', 'INSERT INTO T VALUES (5,6)', 'ROLLBACK', 'COMMIT',
          'INSERT INTO T VALUES (3,4)']:
    sc.run(p1, q, observe=[p1, p2])
sc.note('p1', '>>> EXIT (déconnexion de p1)')
p1.close()
sc.run(p2, 'SELECT * FROM T')
sc.snap('60_tab1.png')
save('tab1', sc)
p2.close()

# tableau 2
reset_T(root)
root.close()
p1 = Session('p1', label='p1 (root) — T dans trans p1')
p2 = Session('p2', 'user1', 'user1', label='p2 (user1) — T vue par p2')
sc = Scenario([p1, p2], 'Panne — tableau 2')
sc.run(p1, 'SET autocommit = 0', observe=[p1, p2])
for q in ['ROLLBACK', 'INSERT INTO T VALUES (3,4)', 'ALTER TABLE T ADD COLUMN C INTEGER',
          'UPDATE T SET B = 3 WHERE A = 1']:
    sc.run(p1, q, observe=[p1, p2])
crash_and_restart(sc, 'p1')
p1.p.kill(); p2.p.kill()
p1 = Session('p1', label='p1 (root) — T dans trans p1')
p2 = Session('p2', 'user1', 'user1', label='p2 (user1) — T vue par p2')
sc.sessions = [p1, p2]
sc.run(p1, 'SELECT * FROM T')
sc.run(p2, 'SELECT * FROM T')
sc.snap('61_tab2.png')
save('tab2', sc)
p2.close()

# tableau 3
reset_T(p1)
p1.close()
p1 = Session('p1', label='p1 (root) — T dans trans p1')
p2 = Session('p2', 'user1', 'user1', label='p2 (user1) — T vue par p2')
sc = Scenario([p1, p2], 'Panne — tableau 3')
sc.run(p1, 'SET autocommit = 0', observe=[p1, p2])
for q in ['INSERT INTO T VALUES (3,4)', 'INSERT INTO T VALUES (5,6)', 'SET autocommit = 1',
          'INSERT INTO T VALUES (7,8)', 'ROLLBACK', 'INSERT INTO T VALUES (9,10)', 'DELETE FROM T WHERE A = 1']:
    sc.run(p1, q, observe=[p1, p2])
crash_and_restart(sc, 'p1')
p1.p.kill(); p2.p.kill()
p1 = Session('p1', label='p1 (root) — T dans trans p1')
p2 = Session('p2', 'user1', 'user1', label='p2 (user1) — T vue par p2')
sc.sessions = [p1, p2]
sc.run(p1, 'SELECT * FROM T')
sc.run(p2, 'SELECT * FROM T')
sc.snap('62_tab3.png')
save('tab3', sc)
p1.close(); p2.close()


# ============================================================ 7. Analyse de programme
def analyse(fname, title, orders):
    root = Session('root', label='root')
    root.send('DROP TABLE IF EXISTS T'); root.send('CREATE TABLE T (A INTEGER, B INTEGER) ENGINE=InnoDB')
    root.send('INSERT INTO T VALUES (1,1),(2,2)'); time.sleep(0.8); root.close()
    p1 = Session('p1', label='p1 (root) — T ds trans p1')
    p2 = Session('p2', 'user1', 'user1', label='p2 (user1) — T ds trans p2')
    p3 = Session('p3', 'user1', 'user1', label='p3 (user1) — T vue par p3')
    S = {'p1': p1, 'p2': p2}
    sc = Scenario([p1, p2, p3], title)
    sc.run(p1, 'SET autocommit = 0')
    sc.run(p2, 'SET autocommit = 0', observe=[p1, p2, p3])
    for who, q in orders:
        w = 2.5 if q.startswith(('UPDATE', 'DELETE')) else 1.5
        sc.run(S[who], q, wait=w, observe=[p1, p2, p3])
    sc.snap(fname)
    for x in (p1, p2, p3):
        x.close()
    return sc


LOG['an1'] = analyse('70_analyse1.png', 'Analyse 1', [('p1', 'DELETE FROM T WHERE B=1'), ('p1', 'COMMIT')]).log
LOG['an2'] = analyse('71_analyse2.png', 'Analyse 2', [('p1', 'DELETE FROM T WHERE B=1'), ('p2', 'UPDATE T SET B=B-1'),
                                                     ('p1', 'COMMIT'), ('p2', 'COMMIT')]).log
LOG['an3'] = analyse('72_analyse3.png', 'Analyse 3', [('p2', 'UPDATE T SET B=B-1'), ('p1', 'DELETE FROM T WHERE B=1'),
                                                     ('p2', 'COMMIT'), ('p1', 'COMMIT')]).log
LOG['an4'] = analyse('73_analyse4.png', 'Analyse 4', [('p1', 'UPDATE T SET A=1 WHERE B=2'), ('p2', 'DELETE FROM T WHERE A=1'),
                                                     ('p1', 'COMMIT'), ('p2', 'COMMIT')]).log

json.dump(LOG, open(r'C:\Users\tomga\tp2_work\rapport\log.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('OK')

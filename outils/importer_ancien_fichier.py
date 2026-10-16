"""Remplit un nouveau fichier avec les données saisies dans l'ancien format (.xlsx ou .csv).

Usage : python3 outils/importer_ancien_fichier.py ANCIEN.xlsx|ANCIEN.csv Vacance_Prenom_Nom.xlsx "Prénom Nom"
"""
import csv, datetime, re, sys
sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
import generer_fichiers as G

def heure(v):
    if v in (None, ""): return None
    if isinstance(v, datetime.datetime): return v.time()
    if isinstance(v, datetime.time): return v
    m = re.fullmatch(r"\s*(\d{1,2})\s*[:hH]\s*(\d{2})?\s*", str(v))
    return datetime.time(int(m.group(1)), int(m.group(2) or 0)) if m else str(v)

def txt(v):
    if v is None: return None
    v = str(v).strip()
    return v or None

def code(v):
    v = txt(v)
    return int(v) if v and v.isdigit() else v

def depuis_csv(path):
    rows = list(csv.reader(open(path, encoding="utf-8-sig").read().splitlines(), delimiter=";"))
    out = []
    for r in rows[1:]:
        r = (r + [""] * 11)[:11]
        if not any(x.strip() for x in r): continue
        days = [x.strip().upper() == "OUI" for x in r[4:9]]
        out.append([heure(r[0]), txt(r[1]), code(r[2]), txt(r[3]), None, *[True if d else None for d in days], txt(r[9]), None, txt(r[10])])
    return out

def depuis_xlsx(path):
    out = []
    for r in G.read_clients(path):
        r[0] = heure(r[0]); r[1] = txt(r[1]); r[2] = code(r[2]); r[3] = txt(r[3])
        r[4:10] = [True if v is True else None for v in r[4:10]]
        r[10] = txt(r[10]); r[12] = txt(r[12])
        out.append(r)
    return out

if __name__ == "__main__":
    src, dest, name = sys.argv[1], sys.argv[2], sys.argv[3]
    clients = depuis_csv(src) if src.endswith(".csv") else depuis_xlsx(src)
    G.build(dest, name, clients)
    print(dest, len(clients), "clients")

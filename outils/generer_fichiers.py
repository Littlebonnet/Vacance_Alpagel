"""Reconstruit les classeurs Vacance_*.xlsx (Excel + Google Sheets).

Formules volontairement classiques (INDEX, MATCH, SMALL, IF, AND, OR, ISNUMBER,
TIMEVALUE, IFERROR) : pas de tableau dynamique, pas de référence structurée,
pas de "" utilisé dans un calcul (Google Sheets le traite comme 0).
"""
import datetime, os, re, sys
import openpyxl, xlsxwriter

MAXROW = 1001                        # dernière ligne client possible dans BASECLIENTS
NLIST = MAXROW - 1                   # lignes calculées par feuille jour
DAYS = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI"]
DAYCOL = dict(zip(DAYS, "FGHIJ"))
CONSIGNES = ["Commande habituelle", "Attention impayé", "Ne pas relancer", "Client sensible"]
TIMES = [datetime.time(h, m) for h in range(8, 19) for m in (0, 15, 30, 45) if (h, m) <= (18, 30)]
BASE_HDR = ["HEURE D'APPEL", "NOM CLIENT", "CODE CLIENT", "VILLE", "TOUS LES JOURS",
            *DAYS, "TÉLÉPHONE", "CONSIGNE", "COMMENTAIRE"]
EXEMPLE = [datetime.time(8, 30), "EXEMPLE - NE PAS SUPPRIMER", 123456, "VILLARD DE LANS", False,
           True, False, True, False, False, "0476000000", "Commande habituelle",
           "Demander Paul - livraison le mardi"]
RED, DARK, GREY = "#C0504D", "#943634", "#808080"
CONSIGNE_FMT = {"Commande habituelle": ("#DDEBF7", "#1F4E78"), "Attention impayé": ("#FFC7CE", "#9C0006"),
                "Ne pas relancer": ("#D9D9D9", "#404040"), "Client sensible": ("#FFEB9C", "#9C5700")}


def read_clients(path):
    """Lit les clients d'un fichier, ancienne (A..K) ou nouvelle (A..M) structure."""
    ws = openpyxl.load_workbook(path)["BASECLIENTS"]
    hdr = [c.value for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in r):
            continue
        if "TOUS LES JOURS" in hdr:                       # déjà au nouveau format
            r = list(r[:13]) + [None] * (13 - len(r[:13]))
            if str(r[1] or "").startswith("EXEMPLE"):
                continue
            rows.append(r)
        else:                                             # ancien format
            days = [str(v or "").strip().upper() in ("OUI", "TRUE", "VRAI", "X") or v is True for v in r[4:9]]
            rows.append([r[0], r[1], r[2], r[3], False, *days, r[9], None, r[10]])
    return rows


def build(path, commercial, clients):
    wb = xlsxwriter.Workbook(path)
    F = lambda **k: wb.add_format({"font_name": "Calibri", "font_size": 11, "valign": "vcenter", **k})

    # ---------------- formats
    hdr = F(bold=True, font_color="white", bg_color=RED, align="center", border=1, border_color="white", text_wrap=True)
    time_f, txt_f, code_f = F(num_format="hh:mm", align="center"), F(num_format="@"), F(align="left")
    box = F(checkbox=True, align="center")
    ex = dict(italic=True, font_color=GREY, bg_color="#EDEDED")
    ex_time, ex_txt, ex_code = F(num_format="hh:mm", align="center", **ex), F(num_format="@", **ex), F(align="left", **ex)
    ex_box = F(checkbox=True, align="center", **ex)

    # ---------------- LISTES (masquée) : sources des menus déroulants
    # ---------------- A SAVOIR
    sav = wb.add_worksheet("A SAVOIR")
    base = wb.add_worksheet("BASECLIENTS")
    days_ws = {d: wb.add_worksheet(d) for d in DAYS}
    lst = wb.add_worksheet("LISTES")

    lst.write(0, 0, "HEURES", hdr); lst.write(0, 1, "CONSIGNES", hdr)
    for i, t in enumerate(TIMES, 1):
        lst.write_datetime(i, 0, datetime.datetime.combine(datetime.date(1899, 12, 31), t), time_f)
    for i, c in enumerate(CONSIGNES, 1):
        lst.write(i, 1, c)
    lst.set_column(0, 1, 22)
    lst.hide()

    write_savoir(wb, sav, F, commercial)

    # ---------------- BASECLIENTS
    widths = [12, 40, 13, 26, 10, 10, 10, 11, 10, 11, 15, 21, 42]
    for i, w in enumerate(widths):
        base.set_column(i, i, w)
    base.set_row(0, 32)
    base.set_default_row(18)
    rows = [EXEMPLE] + clients
    last = max(1 + len(rows), 2)
    base.add_table(0, 0, last - 1, 12, {
        "name": "BASECLIENTS", "style": "Table Style Medium 3",
        "columns": [{"header": h, "header_format": hdr} for h in BASE_HDR]})

    def put(ws, r, c, v, fmt):
        if v is None or v == "":
            ws.write_blank(r, c, None, fmt)
        elif isinstance(v, bool):
            ws.write_boolean(r, c, v, fmt)
        elif isinstance(v, datetime.time):
            ws.write_datetime(r, c, datetime.datetime.combine(datetime.date(1899, 12, 31), v), fmt)
        elif isinstance(v, (int, float)):
            ws.write_number(r, c, v, fmt)
        else:
            ws.write_string(r, c, str(v), fmt)

    for i in range(1, MAXROW):                             # lignes 2..1001
        row = rows[i - 1] if i - 1 < len(rows) else [None] * 13
        is_ex = i == 1
        fmts = ([ex_time, ex_txt, ex_code, ex_txt] + [ex_box] * 6 + [ex_txt] * 3) if is_ex else \
               ([time_f, txt_f, code_f, txt_f] + [box] * 6 + [txt_f] * 3)
        for c in range(13):
            v = row[c]
            if c in range(4, 10):
                v = True if v is True else (v.strip() if isinstance(v, str) and v.strip() else None)  # vide = décochée
            if c == 0 and isinstance(v, datetime.datetime):
                v = v.time()
            put(base, i, c, v, fmts[c])

    base.data_validation(1, 0, MAXROW - 1, 0, {
        "validate": "list", "source": f"=LISTES!$A$2:$A${len(TIMES)+1}", "error_type": "warning",
        "input_title": "Heure d'appel", "input_message": "Choisissez une heure (08:00 à 18:30).",
        "error_title": "Heure inhabituelle", "error_message": "Cette heure n'est pas dans la liste. La garder quand même ?"})
    base.data_validation(1, 11, MAXROW - 1, 11, {
        "validate": "list", "source": f"=LISTES!$B$2:$B${len(CONSIGNES)+1}", "error_type": "warning",
        "error_title": "Consigne", "error_message": "Cette consigne n'est pas dans la liste. La garder quand même ?"})
    base.conditional_format(1, 5, MAXROW - 1, 9, {
        "type": "formula", "criteria": "=$E2=TRUE", "format": wb.add_format({"bg_color": "#C6EFCE"})})
    for txt, (bg, fg) in CONSIGNE_FMT.items():
        base.conditional_format(1, 11, MAXROW - 1, 11, {
            "type": "formula", "criteria": f'=$L2="{txt}"', "format": wb.add_format({"bg_color": bg, "font_color": fg})})
    base.freeze_panes(1, 0)

    # ---------------- feuilles jours
    title = F(bold=True, font_size=16, font_color="white", bg_color=DARK, align="left", indent=1)
    cell = F(text_wrap=True)
    for d, ws in days_ws.items():
        build_day(wb, ws, d, F, title, hdr, time_f, cell)

    sav.activate()
    wb.close()


def build_day(wb, ws, day, F, title, hdr, time_f, cell):
    center = F(align="center")
    X = DAYCOL[day]
    widths = [8, 11, 32, 22, 13, 19, 38, 34]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    ws.set_column(8, 9, 10, None, {"hidden": True})
    ws.set_row(0, 30)
    ws.set_row(1, 22)
    ws.merge_range(0, 0, 0, 7, "", title)
    ws.write_formula(0, 0, f"=\"{day}\"&IF('A SAVOIR'!$D$4=\"\",\"\",\"   -   \"&'A SAVOIR'!$D$4)", title, day)
    for c, h in enumerate(["HEURE", "CODE CLIENT", "NOM CLIENT", "VILLE", "TÉLÉPHONE", "CONSIGNE", "COMMENTAIRE", "NOTES"]):
        ws.write(1, c, h, hdr)
    ws.write(1, 8, "CLÉ (calcul)", hdr)
    ws.write(1, 9, "LIGNE (calcul)", hdr)

    def B(col):
        return f"INDEX(BASECLIENTS!${col}:${col},ROW()-1)"
    n, t, e, dd = B("B"), B("A"), B("E"), B(X)
    checked = f'OR({e}=TRUE,{dd}=TRUE,{e}="X",{dd}="X",{e}="OUI",{dd}="OUI")'
    key = (f'=IF(AND({n}<>"",LEFT({n},7)<>"EXEMPLE",{checked}),'
           f'IF(ISNUMBER({t}),MOD({t},1),IFERROR(TIMEVALUE({t}),1))+ROW()/10000000,"")')
    src = ["A", "C", "B", "D", "K", "L", "M"]
    rng = f"$I$3:$I${NLIST+2}"
    for r in range(2, NLIST + 2):                       # lignes Excel 3..1002
        R = r + 1
        ws.set_row(r, 21)
        ws.write_formula(r, 8, key, None, "")
        ws.write_formula(r, 9, f'=IFERROR(MATCH(SMALL({rng},ROW()-2),{rng},0)+1,"")', None, "")
        for c, s in enumerate(src):
            ix = f"INDEX(BASECLIENTS!${s}:${s},$J{R})"
            ws.write_formula(r, c, f'=IF($J{R}="","",IF({ix}="","",{ix}))', time_f if c == 0 else (center if c == 1 else cell), "")
    last = NLIST + 1
    # Google Sheets et LibreOffice n'appliquent que la 1re règle vraie : la plus précise d'abord
    for txt, (bg, fg) in CONSIGNE_FMT.items():
        ws.conditional_format(2, 5, last, 5, {"type": "formula", "criteria": f'=$F3="{txt}"',
                              "format": wb.add_format({"bg_color": bg, "font_color": fg, "bold": True,
                                                       "border": 1, "border_color": "#BFBFBF"})})
    ws.conditional_format(2, 0, last, 7, {"type": "formula", "criteria": "=AND($C3<>\"\",MOD(ROW(),2)=0)",
                          "format": wb.add_format({"bg_color": "#F2DCDB", "border": 1, "border_color": "#BFBFBF"})})
    ws.conditional_format(2, 0, last, 7, {"type": "formula", "criteria": "=$C3<>\"\"",
                          "format": wb.add_format({"border": 1, "border_color": "#BFBFBF"})})
    ws.freeze_panes(2, 0)
    # impression : A4 paysage, 1 page en largeur, titres répétés, zone limitée aux clients du jour
    ws.set_landscape(); ws.set_paper(9); ws.fit_to_pages(1, 0); ws.repeat_rows(0, 1)
    ws.set_margins(left=0.4, right=0.4, top=0.6, bottom=0.6); ws.center_horizontally()
    ws.set_footer("&LImprimé le &D&RPage &P / &N")
    wb.define_name(f"{day}!_xlnm.Print_Area", f"=OFFSET({day}!$A$1,0,0,MAX(3,COUNT({day}!$J:$J)+2),8)")


def write_savoir(wb, ws, F, commercial):
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2); ws.set_column(1, 1, 5); ws.set_column(2, 2, 10); ws.set_column(3, 3, 30)
    ws.set_column(4, 9, 9); ws.set_column(10, 10, 24)
    big = F(bold=True, font_size=20, font_color="white", bg_color=DARK, indent=1)
    sub = F(italic=True, font_color=GREY)
    sec = F(bold=True, font_size=13, font_color=DARK, bottom=2, bottom_color=DARK)
    num = F(bold=True, font_size=14, font_color="white", bg_color=RED, align="center")
    txt = F(text_wrap=True)
    lab = F(bold=True, align="right")
    inp = F(bold=True, font_size=13, bg_color="#FFF2CC", border=1, border_color="#BF9000")
    note = F(italic=True, font_color=GREY)
    mh = F(bold=True, font_color="white", bg_color=RED, align="center", border=1, border_color="white", font_size=9)
    mc = F(border=1, border_color="#BFBFBF", align="center")
    mcl = F(border=1, border_color="#BFBFBF")
    mg = F(border=1, border_color="#BFBFBF", align="center", bg_color="#C6EFCE")
    arrow = F(bold=True, font_color="#375623")
    dont_h = F(bold=True, font_size=13, font_color="#9C0006", bg_color="#FFC7CE", indent=1)
    dont = F(font_color="#9C0006", bg_color="#FFE5E8", text_wrap=True, indent=1)
    g_h = F(bold=True, font_size=13, font_color="#1F4E78", bg_color="#DDEBF7", indent=1)
    g = F(font_color="#1F4E78", bg_color="#EEF5FB", text_wrap=True, indent=1)

    ws.set_row(0, 40)
    ws.merge_range("B1:K1", "FICHIER D'APPELS - MODE D'EMPLOI", big)
    ws.merge_range("B2:K2", "Fonctionne de la même manière dans Excel et dans Google Sheets.", sub)
    ws.set_row(3, 24)
    ws.merge_range("B4:C4", "Commercial :", lab)
    ws.merge_range("D4:F4", commercial, inp)
    ws.merge_range("G4:K4", "← votre nom : il s'affiche en haut des listes imprimées", note)

    r = 5
    ws.merge_range(r, 1, r, 10, "COMMENT REMPLIR LE FICHIER", sec); r += 1
    steps = [
        "Ouvrez la feuille BASECLIENTS. Une ligne = un client. La ligne grise « EXEMPLE » en haut montre comment remplir : ne la supprimez pas, elle n'apparaît jamais dans les listes.",
        "Choisissez l'heure d'appel dans la liste déroulante (de 08:00 à 18:30, toutes les 15 minutes).",
        "Cochez les jours où le client doit être appelé. Pour un client appelé tous les jours, cochez seulement « TOUS LES JOURS » : les 5 jours se colorent en vert.",
        "Si besoin, choisissez une CONSIGNE pour le remplaçant (Commande habituelle, Attention impayé, Ne pas relancer, Client sensible) et ajoutez un commentaire.",
        "C'est tout : les feuilles LUNDI à VENDREDI se remplissent toutes seules, classées par heure d'appel. Les clients sans heure sont en bas de liste.",
        "Pour imprimer une journée : ouvrez la feuille du jour puis Fichier > Imprimer. La colonne NOTES sert à écrire à la main.",
    ]
    for i, s in enumerate(steps, 1):
        r += 1
        ws.set_row(r, 36)
        ws.write(r, 1, i, num)
        ws.merge_range(r, 2, r, 10, s, txt)

    r += 2
    ws.merge_range(r, 1, r, 10, "EXEMPLE", sec); r += 2
    for c, h in enumerate(["HEURE", "NOM CLIENT", "TOUS", "LUN", "MAR", "MER", "JEU", "VEN", "CONSIGNE"], 2):
        ws.write(r, c, h, mh)
    r += 1
    for c, v in enumerate(["08:30", "LE BISTROT", "☐", "☑", "☐", "☑", "☐", "☐", "Commande habituelle"], 2):
        ws.write(r, c, v, mcl if c in (3, 10) else mc)
    r += 1
    ws.merge_range(r, 2, r, 10, "➜ apparaît automatiquement dans les feuilles LUNDI et MERCREDI, à 08:30.", arrow)
    r += 2
    for c, h in enumerate(["HEURE", "NOM CLIENT", "TOUS", "LUN", "MAR", "MER", "JEU", "VEN", "CONSIGNE"], 2):
        ws.write(r, c, h, mh)
    r += 1
    for c, v in enumerate(["10:00", "HOTEL DU COL", "☑", "", "", "", "", "", "Client sensible"], 2):
        ws.write(r, c, v, mg if 5 <= c <= 9 else (mcl if c in (3, 10) else mc))
    r += 1
    ws.merge_range(r, 2, r, 10, "➜ apparaît dans les 5 feuilles, à 10:00 (les jours se colorent en vert).", arrow)

    r += 2
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 10, "À NE PAS FAIRE", dont_h)
    for s in ["Ne rien taper dans les feuilles LUNDI à VENDREDI : elles sont calculées automatiquement.",
              "Ne pas supprimer ni déplacer de colonnes, ne pas modifier les formules.",
              "Ne pas supprimer la ligne grise EXEMPLE. Maximum 1000 clients par fichier."]:
        r += 1
        ws.set_row(r, 20)
        ws.merge_range(r, 1, r, 10, "✖  " + s, dont)

    r += 2
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 10, "DANS GOOGLE SHEETS", g_h)
    for s in ["Si les cases à cocher n'apparaissent pas : sélectionnez les colonnes TOUS LES JOURS à VENDREDI de BASECLIENTS, puis Insertion > Case à cocher.",
              "En attendant, taper X (ou OUI) dans la case du jour fonctionne aussi.",
              "Pour imprimer : Fichier > Imprimer, puis choisissez la feuille du jour."]:
        r += 1
        ws.set_row(r, 32)
        ws.merge_range(r, 1, r, 10, "•  " + s, g)
    ws.set_landscape(); ws.set_paper(9); ws.fit_to_pages(1, 1)


def commercial_from(path):
    name = os.path.basename(path)[len("Vacance_"):-len(".xlsx")]
    return "" if name == "BASE_VIDE" else name.replace("_", " ")


if __name__ == "__main__":
    src_dir, out_dir = sys.argv[1], sys.argv[2]
    os.makedirs(out_dir, exist_ok=True)
    for f in sorted(os.listdir(src_dir)):
        if re.match(r"Vacance_.*\.xlsx$", f):
            clients = read_clients(os.path.join(src_dir, f))
            build(os.path.join(out_dir, f), commercial_from(f), clients)
            print(f"{f}: {len(clients)} clients")

"""
Gestion des règles d'anomalies externalisées dans un fichier Excel.

But : permettre à des utilisateurs NON techniques de modifier les règles
« simples » (listes blanches, longueurs, protocoles, plages) sans toucher au
code Python. Les règles complexes (format FP2E, cohérence année/diamètre
déduite du numéro) restent dans logique_controles.py.

Fonctionnement :
  - Au 1er lancement, si le fichier regles_anomalies.xlsx n'existe pas à côté
    de l'application, il est créé automatiquement avec les valeurs par défaut
    (celles historiquement codées en dur).
  - À chaque lancement, l'application relit ce fichier.
  - Si le fichier est absent, illisible ou un onglet est mal rempli, on
    retombe sur les valeurs par défaut pour la partie concernée (jamais de
    plantage) et on mémorise un avertissement.
"""

import os
import sys

import pandas as pd

NOM_FICHIER = "regles_anomalies.xlsx"

# =====================================================================
#  VALEURS PAR DÉFAUT (= comportement historique du code)
# =====================================================================

DEFAUT_MARQUES = {
    "Radio": ["SAPPEL (H)", "SAPPEL (C)", "KAMSTRUP", "U Kamstrup"],
    "Tele": ["INTEGRA", "ITRON", "KAIFA", "KAMSTRUP", "SAPPEL (C)",
             "SAPPEL (H)", "SENSUS", "SOCAM", "U Kamstrup"],
}

DEFAUT_TYPES_COMPTEUR = [
    "ALTO", "AQ4", "CF", "CL", "CS", "CV", "GV", "GW", "GY", "HH",
    "HL", "HS", "HT", "HU", "HV", "HY", "HZ", "IB", "II", "IJ",
    "IK", "IL", "INT3", "KAI8", "KI22", "KI31", "KI32", "KM21", "LD22", "MAG1",
    "MAG3", "MAG6", "MAG8", "RTKD", "SEN3", "SEN4", "UH", "UJ", "UK", "YR",
]

# Préfixes de Traité qui imposent le protocole LRA en télérelève
# (tous les autres attendent SGX).
DEFAUT_TRAITES_LRA = ["312", "455", "863", "895", "903", "956"]

# Plage de diamètre autorisée par marque (compteurs classiques).
# Ne s'applique PAS à « U Kamstrup » (compteurs FP2E), seulement à « KAMSTRUP ».
DEFAUT_PLAGE_DIAMETRE = {"KAMSTRUP": (15, 80)}

# Longueur de tête attendue selon (Mode, Marque, Type Compteur).
# Type Compteur vide = s'applique à toutes les valeurs de la marque.
# Règle la plus spécifique (Marque + Type) prioritaire sur la règle générale
# (Marque seule).
# Colonnes : Mode, Marque, Type Compteur, Longueur
DEFAUT_LONGUEUR_TETE = [
    ("Radio", "SAPPEL (C)", "SEN4", 16),
    ("Radio", "SAPPEL (H)", "SEN4", 16),
    ("Tele", "SAPPEL (C)", "", 16),
    ("Tele", "SAPPEL (H)", "", 16),
    ("Tele", "SAPPEL (C)", "SEN3", 15),
    ("Tele", "SAPPEL (H)", "SEN3", 15),
    ("Tele", "SAPPEL (C)", "INT3", 15),
    ("Tele", "SAPPEL (H)", "INT3", 15),
    ("Tele", "ITRON", "", 8),
]

# Correspondance lettre FP2E (4e caractère du n° de compteur) -> diamètre(s).
# Une lettre peut accepter plusieurs diamètres (ex. G = 60 ou 65).
# Le 1er diamètre de la liste sert de correction proposée par défaut.
DEFAUT_DIAMETRE_FP2E = {
    "A": [15], "U": [15], "B": [20], "C": [25], "D": [30], "E": [40], "F": [50],
    "G": [60, 65], "H": [80], "I": [100], "J": [125], "K": [150],
    "L": [200], "M": [250], "N": [300], "O": [350], "P": [400],
}


# =====================================================================
#  Outils
# =====================================================================

def _norm(valeur) -> str:
    """Normalise pour comparaison : MAJUSCULES, sans espaces, sans NaN."""
    s = "" if valeur is None else str(valeur)
    s = s.strip().upper().replace(" ", "")
    return "" if s == "NAN" else s


def _norm_mode(valeur) -> str:
    """Ramène un libellé de mode à 'Radio', 'Tele' ou 'Manuelle'."""
    s = _norm(valeur)
    if s.startswith("RADIO"):
        return "Radio"
    if s.startswith("TELE") or s.startswith("TÉLÉ"):
        return "Tele"
    if s.startswith("MANUEL"):
        return "Manuelle"
    return ""


def dossier_base() -> str:
    """Dossier où chercher/créer le fichier de règles.

    - En .exe (PyInstaller) : à côté de l'exécutable.
    - En développement : à côté de ce fichier .py.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def chemin_fichier() -> str:
    return os.path.join(dossier_base(), NOM_FICHIER)


# =====================================================================
#  Objet Config
# =====================================================================

class ReglesConfig:
    """Règles chargées, prêtes à l'emploi + éventuels avertissements."""

    def __init__(self):
        self.avertissements = []

        # Valeurs (initialisées aux défauts, écrasées par le fichier si OK)
        self.marques = {k: list(v) for k, v in DEFAUT_MARQUES.items()}
        self.types_valides = list(DEFAUT_TYPES_COMPTEUR)
        self.traites_lra = list(DEFAUT_TRAITES_LRA)
        self.plage_diametre = dict(DEFAUT_PLAGE_DIAMETRE)
        self.longueur_tete = list(DEFAUT_LONGUEUR_TETE)
        # Dict prêt à l'emploi {LETTRE: [diamètres]} (utilisé par les règles FP2E)
        self.diametre_fp2e = {k.upper(): list(v) for k, v in DEFAUT_DIAMETRE_FP2E.items()}

    # ---- Vues normalisées (utilisées par le moteur) ----

    def marques_autorisees_norm(self, mode: str) -> set:
        return {_norm(m) for m in self.marques.get(mode, [])}

    def types_valides_norm(self) -> set:
        return {_norm(t) for t in self.types_valides}

    def traites_lra_tuple(self) -> tuple:
        return tuple(str(p).strip() for p in self.traites_lra if str(p).strip())

    def diametre_min_max(self, marque: str, defaut=(15, 400)) -> tuple:
        return self.plage_diametre.get(_norm(marque), defaut)

    def regles_tete(self, mode: str) -> list:
        """Retourne [(marque_norm, type_norm|None, longueur, message), ...]
        triées : règles générales d'abord, spécifiques ensuite (priorité)."""
        regles = []
        for (m, marque, type_c, longueur) in self.longueur_tete:
            if _norm_mode(m) != mode:
                continue
            type_norm = _norm(type_c)
            if type_norm:
                msg = f"{marque} {type_c}: Tête ≠ {longueur} caractères"
            else:
                msg = f"{marque}: Tête ≠ {longueur} caractères"
            regles.append((_norm(marque), type_norm or None, int(longueur), msg))
        # général (type None) avant spécifique (type défini) => spécifique gagne
        regles.sort(key=lambda r: r[1] is not None)
        return regles


# =====================================================================
#  Chargement
# =====================================================================

def charger_config(creer_si_absent: bool = True) -> ReglesConfig:
    """Charge les règles depuis le fichier Excel (avec repli sur défauts)."""
    cfg = ReglesConfig()
    chemin = chemin_fichier()

    if not os.path.exists(chemin):
        if creer_si_absent:
            try:
                creer_fichier_defaut(chemin)
                cfg.avertissements.append(
                    f"Fichier de règles créé avec les valeurs par défaut : {chemin}"
                )
            except Exception as e:  # pragma: no cover
                cfg.avertissements.append(
                    f"Impossible de créer {NOM_FICHIER} ({e}). Règles par défaut utilisées."
                )
        return cfg

    try:
        feuilles = pd.read_excel(chemin, sheet_name=None, dtype=str)
    except Exception as e:
        cfg.avertissements.append(
            f"Lecture de {NOM_FICHIER} impossible ({e}). Règles par défaut utilisées."
        )
        return cfg

    _charger_marques(cfg, feuilles)
    _charger_types(cfg, feuilles)
    _charger_lra(cfg, feuilles)
    _charger_diametre(cfg, feuilles)
    _charger_tete(cfg, feuilles)
    _charger_diametre_fp2e(cfg, feuilles)
    return cfg


def _feuille(feuilles, nom):
    """Retrouve une feuille par nom (tolérant à la casse/espaces)."""
    cible = nom.strip().lower()
    for k, df in feuilles.items():
        if str(k).strip().lower() == cible:
            return df
    return None


def _charger_marques(cfg, feuilles):
    df = _feuille(feuilles, "Marques_autorisees")
    if df is None or "Mode" not in df.columns or "Marque" not in df.columns:
        cfg.avertissements.append("Onglet 'Marques_autorisees' absent/incomplet : défaut utilisé.")
        return
    res = {"Radio": [], "Tele": []}
    for _, row in df.iterrows():
        mode = _norm_mode(row.get("Mode"))
        marque = ("" if pd.isna(row.get("Marque")) else str(row.get("Marque"))).strip()
        if mode in res and marque:
            res[mode].append(marque)
    if res["Radio"] or res["Tele"]:
        for mode in ("Radio", "Tele"):
            if res[mode]:
                cfg.marques[mode] = res[mode]


def _charger_types(cfg, feuilles):
    df = _feuille(feuilles, "Type_Compteur_autorises")
    if df is None or "Type Compteur" not in df.columns:
        cfg.avertissements.append("Onglet 'Type_Compteur_autorises' absent/incomplet : défaut utilisé.")
        return
    vals = [str(v).strip() for v in df["Type Compteur"] if not pd.isna(v) and str(v).strip()]
    if vals:
        cfg.types_valides = vals


def _charger_lra(cfg, feuilles):
    df = _feuille(feuilles, "Traites_LRA_tele")
    if df is None or "Prefixe Traite" not in df.columns:
        cfg.avertissements.append("Onglet 'Traites_LRA_tele' absent/incomplet : défaut utilisé.")
        return
    vals = [str(v).strip() for v in df["Prefixe Traite"] if not pd.isna(v) and str(v).strip()]
    if vals:
        cfg.traites_lra = vals


def _charger_diametre(cfg, feuilles):
    df = _feuille(feuilles, "Plage_diametre")
    if df is None or not {"Marque", "Min", "Max"}.issubset(df.columns):
        cfg.avertissements.append("Onglet 'Plage_diametre' absent/incomplet : défaut utilisé.")
        return
    res = {}
    for _, row in df.iterrows():
        marque = _norm(row.get("Marque"))
        try:
            mn = int(float(row.get("Min")))
            mx = int(float(row.get("Max")))
        except (TypeError, ValueError):
            continue
        if marque:
            res[marque] = (mn, mx)
    if res:
        cfg.plage_diametre = res


def _charger_tete(cfg, feuilles):
    df = _feuille(feuilles, "Longueur_tete")
    cols = {"Mode", "Marque", "Type Compteur", "Longueur"}
    if df is None or not cols.issubset(df.columns):
        cfg.avertissements.append("Onglet 'Longueur_tete' absent/incomplet : défaut utilisé.")
        return
    res = []
    for _, row in df.iterrows():
        mode = ("" if pd.isna(row.get("Mode")) else str(row.get("Mode"))).strip()
        marque = ("" if pd.isna(row.get("Marque")) else str(row.get("Marque"))).strip()
        type_c = ("" if pd.isna(row.get("Type Compteur")) else str(row.get("Type Compteur"))).strip()
        try:
            longueur = int(float(row.get("Longueur")))
        except (TypeError, ValueError):
            continue
        if _norm_mode(mode) and marque:
            res.append((mode, marque, type_c, longueur))
    if res:
        cfg.longueur_tete = res


def _charger_diametre_fp2e(cfg, feuilles):
    df = _feuille(feuilles, "Diametre_FP2E")
    if df is None or not {"Lettre", "Diametre"}.issubset(df.columns):
        cfg.avertissements.append("Onglet 'Diametre_FP2E' absent/incomplet : défaut utilisé.")
        return
    res = {}
    for _, row in df.iterrows():
        lettre = _norm(row.get("Lettre"))
        try:
            diam = int(float(row.get("Diametre")))
        except (TypeError, ValueError):
            continue
        if len(lettre) == 1 and lettre.isalpha():
            res.setdefault(lettre, []).append(diam)
    if res:
        cfg.diametre_fp2e = res


# =====================================================================
#  Génération du fichier par défaut
# =====================================================================

NOTICE = [
    ["NOTICE — Fichier de règles des anomalies"],
    [""],
    ["Ce fichier permet de modifier certaines règles SANS toucher au code."],
    ["Après modification : enregistrez le fichier, fermez-le, puis relancez l'application."],
    [""],
    ["Onglets :"],
    ["  • Marques_autorisees   : marques acceptées par mode (Radio / Tele)."],
    ["  • Type_Compteur_autorises : codes Type Compteur acceptés (tous modes)."],
    ["  • Traites_LRA_tele      : préfixes de Traité en LRA (télé). Le reste = SGX."],
    ["  • Plage_diametre        : diamètre min/max autorisé par marque."],
    ["  • Longueur_tete         : longueur de tête attendue selon Mode/Marque/Type."],
    ["  • Diametre_FP2E         : diamètre(s) correspondant à chaque lettre FP2E."],
    [""],
    ["Règles :"],
    ["  - Mode = Radio, Tele ou Manuelle (selon l'onglet)."],
    ["  - Une valeur vide dans 'Type Compteur' (onglet Longueur_tete) = s'applique"],
    ["    à toutes les valeurs de la marque. Une ligne avec un Type précis est"],
    ["    prioritaire sur la ligne générale."],
    ["  - Onglet Diametre_FP2E : une ligne par (Lettre, Diametre). Une lettre qui"],
    ["    accepte plusieurs diamètres a plusieurs lignes (ex. G -> 60 et G -> 65)."],
    ["    Le 1er diamètre listé pour une lettre sert de correction proposée."],
    ["  - Ne renommez PAS les onglets ni les colonnes (en-têtes)."],
    ["  - En cas d'erreur de saisie, l'application ignore la partie concernée et"],
    ["    utilise les valeurs par défaut (elle ne plante pas)."],
    [""],
    ["Astuce : pour repartir de zéro, supprimez ce fichier et relancez l'application :"],
    ["il sera recréé avec les valeurs par défaut."],
]


def creer_fichier_defaut(chemin: str = None):
    """Écrit un regles_anomalies.xlsx pré-rempli avec les valeurs par défaut."""
    if chemin is None:
        chemin = chemin_fichier()

    df_marques = pd.DataFrame(
        [(mode, m) for mode, lst in DEFAUT_MARQUES.items() for m in lst],
        columns=["Mode", "Marque"],
    )
    df_types = pd.DataFrame({"Type Compteur": DEFAUT_TYPES_COMPTEUR})
    df_lra = pd.DataFrame({"Prefixe Traite": DEFAUT_TRAITES_LRA})
    df_diam = pd.DataFrame(
        [(m, mn, mx) for m, (mn, mx) in DEFAUT_PLAGE_DIAMETRE.items()],
        columns=["Marque", "Min", "Max"],
    )
    df_tete = pd.DataFrame(DEFAUT_LONGUEUR_TETE,
                           columns=["Mode", "Marque", "Type Compteur", "Longueur"])
    df_diam_fp2e = pd.DataFrame(
        [(lettre, d) for lettre, diams in DEFAUT_DIAMETRE_FP2E.items() for d in diams],
        columns=["Lettre", "Diametre"],
    )
    df_notice = pd.DataFrame(NOTICE)

    with pd.ExcelWriter(chemin, engine="openpyxl") as writer:
        df_notice.to_excel(writer, sheet_name="Notice", index=False, header=False)
        df_marques.to_excel(writer, sheet_name="Marques_autorisees", index=False)
        df_types.to_excel(writer, sheet_name="Type_Compteur_autorises", index=False)
        df_lra.to_excel(writer, sheet_name="Traites_LRA_tele", index=False)
        df_diam.to_excel(writer, sheet_name="Plage_diametre", index=False)
        df_tete.to_excel(writer, sheet_name="Longueur_tete", index=False)
        df_diam_fp2e.to_excel(writer, sheet_name="Diametre_FP2E", index=False)

        # Largeur de colonnes lisible
        for ws in writer.book.worksheets:
            for col in ws.columns:
                lettre = col[0].column_letter
                largeur = max((len(str(c.value)) for c in col if c.value is not None), default=10)
                ws.column_dimensions[lettre].width = min(largeur + 2, 70)

    return chemin


# =====================================================================
#  Singleton (chargé une seule fois par exécution)
# =====================================================================

_CONFIG = None


def get_config(recharger: bool = False) -> ReglesConfig:
    global _CONFIG
    if _CONFIG is None or recharger:
        _CONFIG = charger_config()
    return _CONFIG

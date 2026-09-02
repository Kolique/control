# 📚 Guide d'installation - Analyses Anomalies

## 🎯 Pour les utilisateurs (NON-TECHNIQUES)

### Option simple : Utiliser le fichier .exe

1. **Télécharger** le fichier `AnalysesAnomalies.exe`
2. **Double-cliquer** dessus pour lancer l'application
3. **C'est tout !** Aucune installation nécessaire

---

## 🛠️ Pour l'administrateur (INSTALLATION INITIALE)

### Prérequis
- Python 3.8 ou supérieur installé
- Windows 10/11

### Étape 1 : Installation des dépendances

Ouvrez un terminal (CMD ou PowerShell) dans le dossier contenant les fichiers, puis :

```bash
pip install -r requirements.txt
```

### Étape 2 : Tester l'application en mode développement

```bash
python app_analyses_anomalies.py
```

L'interface graphique devrait s'ouvrir. Testez avec vos fichiers.

### Étape 3 : Compiler en .exe (pour distribution)

```bash
python compiler_exe.py
```

Cela va créer un dossier `dist/` contenant `AnalysesAnomalies.exe`

---

## 📖 Guide d'utilisation

### 1️⃣ Lancer l'application
- Double-cliquez sur `AnalysesAnomalies.exe`

### 2️⃣ Sélectionner les fichiers
Deux méthodes au choix :

**Méthode A : Glisser-déposer**
- Faites glisser votre fichier Excel dans la zone correspondante
  - 📻 Radiorelève
  - 📡 Télérelève
  - ✍️ Manuelle

**Méthode B : Parcourir**
- Cliquez sur le bouton "Parcourir"
- Sélectionnez votre fichier Excel

### 3️⃣ Choisir le dossier de sortie
- Cliquez sur "Choisir" à côté de "📁 Dossier de sortie"
- Sélectionnez où vous voulez sauvegarder les rapports

### 4️⃣ Analyser
- Cliquez sur le gros bouton vert **"🚀 ANALYSER"**
- Attendez la fin de l'analyse (barre de progression)
- Une fenêtre s'ouvre pour ouvrir le dossier des résultats

---

## 📊 Résultats générés

L'application crée automatiquement :

### Rapports par mode
- `Rapport_Radioreleve_2025_Janvier.xlsx`
- `Rapport_Telereleve_2025_Janvier.xlsx`
- `Rapport_Manuelle_2025_Janvier.xlsx`

**Contenu :**
- Onglet "Récapitulatif" avec liens cliquables
- Onglet "Toutes_Anomalies" avec toutes les lignes
- Un onglet par type d'anomalie (surlignage automatique)

### Rapports par Traité (dossier "Traites/")
- `Anomalies_Traite_965_2025_Janvier.xlsx`
- `Anomalies_Traite_455_2025_Janvier.xlsx`
- `Anomalies_Traite_899_2025_Janvier.xlsx`
- etc.

**Contenu :** Même structure, regroupé par les 3 premiers chiffres du Traité

---

## ⚙️ Modifier les règles d'anomalies (SANS toucher au code)

Certaines règles peuvent être modifiées directement dans un **fichier Excel**,
sans aucune connaissance en programmation.

### Comment faire

1. Dans l'application, cliquez sur le bouton **« ⚙️ Règles »** (en haut à droite).
   - Le fichier `regles_anomalies.xlsx` s'ouvre (il est créé automatiquement au
     besoin, à côté de l'application).
2. Modifiez les onglets voulus, **enregistrez** et **fermez** le fichier.
3. Relancez une analyse : les nouvelles règles sont prises en compte.

### Onglets disponibles

| Onglet | Rôle | Exemple d'ajout |
|---|---|---|
| **Marques_autorisees** | Marques acceptées par mode | Ligne `Tele` / `DIEHL` |
| **Type_Compteur_autorises** | Codes Type Compteur acceptés (tous modes) | Ligne `ZZ99` |
| **Traites_LRA_tele** | Préfixes de Traité en LRA (télé), le reste en SGX | Ligne `777` |
| **Plage_diametre** | Diamètre min/max autorisé par marque | `KAMSTRUP` / `15` / `400` |
| **Longueur_tete** | Longueur de tête attendue selon Mode/Marque/Type | `Radio` / `KAMSTRUP` / `KM21` / `10` |

Pour l'onglet **Longueur_tete**, laisser la colonne *Type Compteur* vide = la
règle s'applique à **toutes** les valeurs de la marque. Une ligne avec un Type
précis est **prioritaire** sur la règle générale.

### Sécurité

- En cas d'erreur de saisie (onglet renommé, colonne manquante, valeur invalide),
  l'application **ignore la partie concernée** et utilise les valeurs par défaut :
  elle ne plante pas.
- Pour **repartir de zéro**, supprimez `regles_anomalies.xlsx` et relancez :
  il sera recréé avec les valeurs par défaut.

### Ce qui n'est PAS modifiable ici

Les règles techniques et stables restent dans le code : format **FP2E** du numéro
de compteur, cohérence **année/diamètre** déduite du numéro, déduction du Type
Compteur, règles KAMSTRUP compteur = tête. Pour celles-ci, contactez la personne
en charge de la maintenance du code.

---

## ❓ FAQ / Dépannage

### L'application ne se lance pas
- **Solution 1 :** Vérifiez que vous avez les droits administrateur
- **Solution 2 :** Désactivez temporairement l'antivirus
- **Solution 3 :** Relancez en mode développement : `python app_analyses_anomalies.py`

### "Erreur : Veuillez sélectionner au moins un fichier"
- Vous devez sélectionner **au moins un** des trois fichiers (Radio, Télé ou Manuelle)
- Vous n'êtes **pas obligé** de tous les remplir

### "Erreur : Veuillez sélectionner un dossier de sortie"
- Cliquez sur "Choisir" à côté de "📁 Dossier de sortie"
- Sélectionnez un dossier existant ou créez-en un nouveau

### L'analyse est très lente
- C'est normal pour les gros fichiers (>10 000 lignes)
- L'application traite en arrière-plan
- **Ne fermez pas** pendant l'analyse !

### Les rapports ne s'ouvrent pas automatiquement
- Ouvrez manuellement le dossier de sortie que vous avez choisi
- Les fichiers Excel sont dedans

---

## 🔄 Mise à jour

Pour mettre à jour l'application :

1. Remplacez les fichiers :
   - `logique_controles.py`
   - `main.py`
   - `app_analyses_anomalies.py`

2. Recompilez :
   ```bash
   python compiler_exe.py
   ```

3. Redistribuez le nouveau `AnalysesAnomalies.exe`

---

## 📞 Support

Pour toute question ou problème :
- Contactez l'administrateur système
- Ou envoyez un email avec une capture d'écran de l'erreur

---

## 📝 Notes de version

### Version 2.0 (Décembre 2024)
- ✨ Interface graphique moderne
- 🎯 Drag & drop des fichiers
- 📊 Support Traités 965/455/899 avec lettre finale
- 🔧 Support KAMSTRUP FP2E (commence par U)
- 📁 Sélection libre des fichiers et dossier de sortie
- ⚡ Analyse multi-threadée (pas de freeze)
- 🎨 Rapports Excel améliorés avec liens cliquables

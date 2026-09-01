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

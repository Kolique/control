"""
Script de compilation en .exe
Transforme l'application Python en exécutable Windows
"""

import os
import subprocess
import sys

def compiler_en_exe():
    """Compile l'application en fichier .exe"""
    
    print("=" * 60)
    print("🚀 Compilation de Analyses Anomalies en .exe")
    print("=" * 60)
    print()
    
    # Vérifier que PyInstaller est installé
    try:
        import PyInstaller
    except ImportError:
        print("❌ PyInstaller n'est pas installé.")
        print("Installation en cours...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("✅ PyInstaller installé avec succès")
        print()
    
    # Commande PyInstaller pour un seul fichier .exe portable
    commande = [
        sys.executable, "-m", "PyInstaller",  
        "--onefile",                          # UN SEUL FICHIER .exe (portable !)
        "--windowed",                         # Mode fenêtre (pas de console)
        "--name=AnalysesAnomalies",          # Nom de l'exécutable
        "--icon=NONE",                        # Pas d'icône personnalisée
        "--add-data=logique_controles.py;.", # Inclure le module de logique
        "--add-data=main.py;.",              # Inclure main.py
        "--add-data=regles_config.py;.",     # Inclure le module de configuration des règles

        # Imports critiques pour le fonctionnement
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=tkinterdnd2",
        "--hidden-import=xlrd",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",
        "--hidden-import=threading",
        "--hidden-import=os",
        "--hidden-import=sys",
        "--hidden-import=logique_controles",
        "--hidden-import=main",
        "--hidden-import=regles_config",
        
        # Exclusions pour réduire la taille
        "--exclude-module=matplotlib",
        "--exclude-module=scipy", 
        "--exclude-module=PIL",
        "--exclude-module=setuptools",
        "--exclude-module=unittest",
        "--exclude-module=test",
        
        # Options d'optimisation
        "--noconfirm",                        # Pas de confirmation (écrase les fichiers)
        "--clean",                            # Nettoie le cache
        "app_analyses_anomalies.py"
    ]
    
    print("📦 Compilation en cours...")
    print(f"Commande : {' '.join(commande)}")
    print()
    
    try:
        subprocess.run(commande, check=True)
        print()
        print("=" * 60)
        print("✅ COMPILATION RÉUSSIE !")
        print("=" * 60)
        print()
        print(f"📁 L'exécutable se trouve dans : {os.path.abspath('dist')}")
        print(f"📝 Fichier : AnalysesAnomalies.exe")
        print()
        print("🎉 FICHIER UNIQUE ET PORTABLE !")
        print("   ✅ Vous pouvez copier ce seul fichier .exe n'importe où")
        print("   ✅ Aucune installation Python requise sur la machine cible")
        print("   ✅ Fonctionne sur n'importe quel Windows")
        print()
        print("⚠️  Note : Le drag & drop pourrait ne pas marcher")
        print("   → Utilisez les boutons 'Sélectionner fichier' dans ce cas")
        print()
        
    except subprocess.CalledProcessError as e:
        print()
        print("=" * 60)
        print("❌ ERREUR LORS DE LA COMPILATION")
        print("=" * 60)
        print(f"Erreur : {e}")
        print()
        print("💡 Assurez-vous que tous les fichiers nécessaires sont présents :")
        print("   - app_analyses_anomalies.py")
        print("   - logique_controles.py")
        print("   - main.py")
        print("   - regles_config.py")
        print()


if __name__ == "__main__":
    compiler_en_exe()

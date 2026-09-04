"""
Korrect'eau - Interface Graphique
Version 2.0
Permet l'analyse des fichiers de compteurs (Radio, Télé, Manuelle)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import sys

# Classe pour les tooltips
class ToolTip:
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.tipwindow = None

    def enter(self, event=None):
        self.show_tooltip()

    def leave(self, event=None):
        self.hide_tooltip()

    def show_tooltip(self):
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 25
        y = y + cy + self.widget.winfo_rooty() + 25
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                        font=("Segoe UI", 9, "normal"))
        label.pack(ipadx=1)

    def hide_tooltip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

# Imports optimisés - chargement à la demande
def lazy_import_pandas():
    """Import pandas seulement quand nécessaire"""
    global pd
    import pandas as pd
    return pd

def lazy_import_datetime():
    """Import datetime seulement quand nécessaire"""
    from datetime import datetime
    return datetime

def lazy_import_subprocess():
    """Import subprocess seulement quand nécessaire"""
    import subprocess
    return subprocess

def lazy_import_tkinterdnd():
    """Import tkinterdnd2 seulement quand nécessaire avec gestion d'erreur pour exe"""
    global DND_FILES, TkinterDnD
    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
        return DND_FILES, TkinterDnD
    except ImportError as e:
        print(f"⚠️ Erreur import tkinterdnd2: {e}")
        print("💡 Le drag & drop ne fonctionnera pas, utilisez les boutons 'Parcourir'")
        return None, None

def lazy_import_logique():
    """Import logique_controles seulement quand nécessaire"""
    global logique_controles
    import logique_controles
    return logique_controles


class AnalysesAnomaliesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Korrect'eau")
        self.root.geometry("950x750")
        self.root.resizable(False, False)
        
        # Test de tkinterdnd2 pour l'exe
        DND_FILES, TkinterDnD = lazy_import_tkinterdnd()
        self.dnd_available = DND_FILES is not None and TkinterDnD is not None
        
        # PLUS DE RE-CRÉATION DE FENÊTRE ! On utilise celle passée en paramètre
        
        # Fichiers avec statut et tooltips
        self.fichier_radio = None
        self.fichier_tele = None
        self.fichier_manuelle = None
        self.dossier_sortie = None
        
        # Status des fichiers pour icônes (None, 'selected', 'error')
        self.status_radio = None
        self.status_tele = None
        self.status_manuelle = None
        self.status_sortie = None
        
        # Couleurs modernes (Style Bootstrap moderne)
        self.bg_color = "#f8f9fa"          # Gris très clair moderne
        self.primary_color = "#198754"      # Vert Bootstrap success
        self.primary_dark = "#146c43"       # Vert Bootstrap foncé
        self.secondary_color = "#0d6efd"    # Bleu Bootstrap primary
        self.accent_color = "#fd7e14"       # Orange Bootstrap
        self.text_color = "#212529"         # Noir Bootstrap
        self.text_light = "#6c757d"         # Gris Bootstrap
        self.border_color = "#e9ecef"       # Bordure Bootstrap
        self.card_bg = "#ffffff"            # Blanc pur
        self.card_shadow = "#e9ecef"        # Ombre légère
        self.success_color = "#198754"      # Vert succès
        self.error_color = "#dc3545"        # Rouge erreur
        
        self.root.configure(bg=self.bg_color)
        
        self.create_widgets()
        
    def create_widgets(self):
        # === HEADER PREMIUM ===
        header_frame = tk.Frame(self.root, bg=self.primary_color, height=100)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        
        # Sous-header pour effet visuel
        header_inner = tk.Frame(header_frame, bg=self.primary_color)
        header_inner.pack(fill="both", expand=True, padx=30, pady=15)
        
        title_label = tk.Label(
            header_inner,
            text="📊 Korrect'eau",
            font=("Segoe UI", 28, "bold"),
            bg=self.primary_color,
            fg="white"
        )
        title_label.pack(side="left", anchor="w")
        
        # === BOUTON REFRESH MODERNE ===
        refresh_btn = tk.Button(
            header_inner,
            text="🔄",
            command=self.refresh_app,
            bg="white",
            fg=self.primary_color,
            font=("Segoe UI", 16, "bold"),
            relief="flat",
            cursor="hand2",
            bd=0,
            width=3,
            height=1,
            activebackground="#f8f9fa",
            activeforeground=self.primary_dark
        )
        refresh_btn.pack(side="right", anchor="e", padx=(10, 0))

        # === BOUTON RÈGLES (ouvre le fichier Excel de configuration) ===
        regles_btn = tk.Button(
            header_inner,
            text="⚙️ Règles",
            command=self.ouvrir_regles,
            bg="white",
            fg=self.primary_color,
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=12,
            height=1,
            activebackground="#f8f9fa",
            activeforeground=self.primary_dark
        )
        regles_btn.pack(side="right", anchor="e", padx=(10, 0))
        ToolTip(regles_btn, "Ouvre le fichier Excel des règles d'anomalies.\n"
                            "Modifiez-le, enregistrez, fermez, puis relancez l'analyse.")
        
        # === CONTENU PRINCIPAL ===
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill="both", expand=True, padx=25, pady=20)
        
        # Titre section fichiers
        files_title = tk.Label(
            main_frame,
            text="📁 Étape 1 : Sélectionner vos fichiers",
            font=("Segoe UI", 13, "bold"),
            bg=self.bg_color,
            fg=self.text_color
        )
        files_title.pack(anchor="w", pady=(0, 15))
        
        # Container pour les 3 fichiers avec meilleur espacement
        files_container = tk.Frame(main_frame, bg=self.bg_color)
        files_container.pack(fill="x", pady=(0, 30))
        
        # Configuration grille avec espacement
        for i in range(3):
            files_container.grid_columnconfigure(i, weight=1, pad=15)
        
        # === FICHIER RADIO ===
        self.create_file_dropzone(
            files_container,
            "📻",
            "Radiorelève",
            "fichier_radio",
            "#E3F2FD",
            0
        )
        
        # === FICHIER TELE ===
        self.create_file_dropzone(
            files_container,
            "📡",
            "Télérelève",
            "fichier_tele",
            "#FFF3E0",
            1
        )
        
        # === FICHIER MANUELLE ===
        self.create_file_dropzone(
            files_container,
            "📝",
            "Manuelle",
            "fichier_manuelle",
            "#F3E5F5",
            2
        )
        
        # === SÉPARATEUR ===
        separator = tk.Frame(main_frame, height=2, bg=self.border_color)
        separator.pack(fill="x", pady=20)
        
        # === DOSSIER DE SORTIE ===
        output_title = tk.Label(
            main_frame,
            text="💾 Étape 2 : Dossier de sortie",
            font=("Segoe UI", 13, "bold"),
            bg=self.bg_color,
            fg=self.text_color
        )
        output_title.pack(anchor="w", pady=(0, 15))
        
        sortie_frame = tk.Frame(main_frame, bg=self.card_bg, relief="flat", bd=0)
        sortie_frame.pack(fill="x", pady=(0, 20), padx=8)
        sortie_frame.configure(highlightthickness=1, highlightbackground=self.border_color)
        
        sortie_inner = tk.Frame(sortie_frame, bg=self.card_bg)
        sortie_inner.pack(fill="x", padx=20, pady=15)
        
        sortie_label = tk.Label(
            sortie_inner,
            text="📂 Dossier de destination :",
            font=("Segoe UI", 11, "bold"),
            bg=self.card_bg,
            fg=self.text_color
        )
        sortie_label.pack(side="left", anchor="w")
        
        self.sortie_path_label = tk.Label(
            sortie_inner,
            text="Aucun dossier sélectionné",
            font=("Segoe UI", 10),
            bg=self.card_bg,
            fg=self.text_light,
            anchor="w"
        )
        self.sortie_path_label.pack(side="left", fill="x", expand=True, padx=15)
        
        sortie_btn = tk.Button(
            sortie_inner,
            text="📂 Parcourir",
            command=self.choisir_dossier_sortie,
            bg=self.secondary_color,
            fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=8,
            activebackground="#1565C0",
            activeforeground="white"
        )
        # Tooltip pour le dossier de sortie
        ToolTip(sortie_frame, "Sélectionnez où sauvegarder les rapports d'analyse\nLes fichiers Excel seront créés dans ce dossier")
        
        sortie_btn.pack(side="right")
        
        # Tooltip pour le dossier de sortie
        ToolTip(sortie_frame, "Sélectionnez où sauvegarder les rapports d'analyse\nLes fichiers Excel seront créés dans ce dossier")
        
        # === BARRE DE PROGRESSION AMÉLIORÉE ===
        self.progress_frame = tk.Frame(main_frame, bg=self.bg_color)
        self.progress_frame.pack(fill="x", pady=(0, 20))
        
        # Titre et pourcentage
        progress_header = tk.Frame(self.progress_frame, bg=self.bg_color)
        progress_header.pack(fill="x")
        
        progress_title = tk.Label(
            progress_header,
            text="Analyse en cours...",
            font=("Segoe UI", 11, "bold"),
            bg=self.bg_color,
            fg=self.text_color
        )
        progress_title.pack(side="left")
        
        self.progress_percent = tk.Label(
            progress_header,
            text="0%",
            font=("Segoe UI", 11, "bold"),
            bg=self.bg_color,
            fg=self.primary_color
        )
        self.progress_percent.pack(side="right")
        
        # Barre de progression
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            mode="determinate",  # Mode déterministe pour afficher le pourcentage
            length=400
        )
        self.progress_bar.pack(fill="x", pady=(5, 0))
        
        # Statut détaillé
        self.status_label = tk.Label(
            self.progress_frame,
            text="",
            font=("Segoe UI", 9),
            bg=self.bg_color,
            fg=self.secondary_color
        )
        self.status_label.pack(anchor="w", pady=(5, 0))
        
        self.progress_frame.pack_forget()  # Caché par défaut
        
        # === BOUTON ANALYSER PRINCIPAL ===
        button_frame = tk.Frame(main_frame, bg=self.bg_color)
        button_frame.pack(fill="x", pady=(10, 0))
        
        self.analyze_btn = tk.Button(
            button_frame,
            text="🚀 ANALYSER LES DONNÉES",
            command=self.lancer_analyse,
            bg=self.primary_color,
            fg="white",
            font=("Segoe UI", 14, "bold"),
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=60,
            pady=18,
            activebackground=self.primary_dark,
            activeforeground="white"
        )
        self.analyze_btn.pack(anchor="center", expand=True)
        
        # Tooltip pour le bouton analyser
        ToolTip(self.analyze_btn, "Lance l'analyse des fichiers sélectionnés\nGénère des rapports Excel détaillés avec anomalies")
        
    def update_file_status(self, file_attr, status):
        """Met à jour le statut visuel d'un fichier"""
        status_icon = getattr(self, f"{file_attr}_status_icon", None)
        drop_frame = getattr(self, f"{file_attr}_drop_frame", None)
        
        if not status_icon or not drop_frame:
            return
            
        if status == 'selected':
            status_icon.config(text="✅", fg="#4CAF50")
            drop_frame.config(highlightthickness=2, highlightbackground="#4CAF50")
        elif status == 'error':
            status_icon.config(text="❌", fg="#F44336")
            drop_frame.config(highlightthickness=2, highlightbackground="#F44336")
        else:
            status_icon.config(text="")
            drop_frame.config(highlightthickness=1, highlightbackground=self.border_color)
        
    def create_file_dropzone(self, parent, icon, label_text, file_attr, bg_color, column):
        """Crée une zone de dépôt de fichier avec design card moderne"""
        # Conteneur avec colonne et espacement amélioré
        card_frame = tk.Frame(parent, bg=self.bg_color)
        card_frame.grid(row=0, column=column, padx=15, pady=10, sticky="nsew")
        parent.grid_columnconfigure(column, weight=1)
        
        # Carte principale avec ombre et bordures arrondies
        drop_frame = tk.Frame(
            card_frame,
            bg=self.card_bg,
            relief="flat",
            bd=0,
            height=240
        )
        drop_frame.pack(fill="both", expand=True)
        drop_frame.configure(highlightthickness=1, highlightbackground=self.border_color)
        drop_frame.pack_propagate(False)
        
        # Contenu avec meilleur spacing
        inner_frame = tk.Frame(drop_frame, bg=self.card_bg)
        inner_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Icône grande avec statut
        icon_frame = tk.Frame(inner_frame, bg=self.card_bg)
        icon_frame.pack(pady=(0, 12))
        
        icon_label = tk.Label(
            icon_frame,
            text=icon,
            font=("Segoe UI", 42),
            bg=self.card_bg,
            fg=self.primary_color
        )
        icon_label.pack(side="left")
        
        # Icône de statut (initialement cachée)
        status_icon = tk.Label(
            icon_frame,
            text="",
            font=("Segoe UI", 18),
            bg=self.card_bg
        )
        status_icon.pack(side="left", padx=(8, 0))
        
        # Titre avec meilleure typographie
        title = tk.Label(
            inner_frame,
            text=label_text,
            font=("Segoe UI", 13, "bold"),
            bg=self.card_bg,
            fg=self.text_color
        )
        title.pack(pady=(0, 8))
        
        # Chemin du fichier avec style amélioré
        path_label = tk.Label(
            inner_frame,
            text="Cliquez ou glissez un fichier Excel",
            font=("Segoe UI", 9),
            bg=self.card_bg,
            fg=self.text_light,
            wraplength=220
        )
        path_label.pack(pady=(0, 15))
        
        # Bouton parcourir taille équilibrée
        browse_btn = tk.Button(
            inner_frame,
            text="📁 Sélectionner fichier",
            command=lambda: self.choisir_fichier(file_attr, path_label),
            bg=self.primary_color,
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=25,
            pady=12,
            activebackground="#146c43",
            activeforeground="white"
        )
        browse_btn.pack(fill="x", pady=(5, 8))
        
        # Tooltip pour la carte
        tooltip_text = {
            "Radiorelève": "Glissez votre fichier radioreleve.xlsx ici\nAnalyse des compteurs en mode radio",
            "Télérelève": "Glissez votre fichier telereleve.xlsx ici\nAnalyse des compteurs en mode télé", 
            "Manuelle": "Glissez votre fichier manuelle.xlsx ici\nAnalyse des compteurs saisis manuellement"
        }
        ToolTip(drop_frame, tooltip_text.get(label_text, "Glissez votre fichier Excel ici"))
        
        # Drag and drop seulement si disponible (fonctionne dans l'exe)
        if self.dnd_available: 
            DND_FILES, TkinterDnD = lazy_import_tkinterdnd()  # Import à la demande
            if DND_FILES and TkinterDnD:
                drop_frame.drop_target_register(DND_FILES)
                drop_frame.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, file_attr, path_label))
                
                inner_frame.drop_target_register(DND_FILES)
                inner_frame.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, file_attr, path_label))
                
                icon_label.drop_target_register(DND_FILES)
                icon_label.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, file_attr, path_label))
                
                title.drop_target_register(DND_FILES)
                title.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, file_attr, path_label))
                
                path_label.drop_target_register(DND_FILES)
                path_label.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, file_attr, path_label))
        
        # Stocker les références
        setattr(self, f"{file_attr}_label", path_label)
        setattr(self, f"{file_attr}_status_icon", status_icon)
        setattr(self, f"{file_attr}_drop_frame", drop_frame)
    
    def ouvrir_regles(self):
        """Ouvre le fichier Excel de configuration des règles (le crée si absent)."""
        try:
            import regles_config
            chemin = regles_config.chemin_fichier()
            if not os.path.exists(chemin):
                regles_config.creer_fichier_defaut(chemin)
                messagebox.showinfo(
                    "⚙️ Fichier de règles créé",
                    "Le fichier de règles n'existait pas encore : il vient d'être créé "
                    "avec les valeurs par défaut.\n\n"
                    f"Emplacement :\n{chemin}\n\n"
                    "Il va s'ouvrir. Après modification : enregistrez, fermez, "
                    "puis relancez l'analyse."
                )
            if sys.platform == "win32":
                os.startfile(chemin)
            else:
                subprocess = lazy_import_subprocess()
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.run([opener, chemin])
        except Exception as e:
            messagebox.showerror(
                "❌ Erreur",
                f"Impossible d'ouvrir le fichier de règles :\n\n{e}"
            )

    def refresh_app(self):
        """Remet l'application à zéro tous les fichiers sélectionnés)"""
        # Réinitialiser les variables de fichiers
        self.fichier_radio = None
        self.fichier_tele = None
        self.fichier_manuelle = None
        self.dossier_sortie = None
        
        # Réinitialiser l'affichage des labels de fichiers
        if hasattr(self, 'fichier_radio_label'):
            self.fichier_radio_label.config(
                text="Cliquez ou glissez un fichier",
                fg=self.text_light,
                font=("Segoe UI", 9)
            )
        
        if hasattr(self, 'fichier_tele_label'):
            self.fichier_tele_label.config(
                text="Cliquez ou glissez un fichier",
                fg=self.text_light,
                font=("Segoe UI", 9)
            )
        
        if hasattr(self, 'fichier_manuelle_label'):
            self.fichier_manuelle_label.config(
                text="Cliquez ou glissez un fichier",
                fg=self.text_light,
                font=("Segoe UI", 9)
            )
        
        # Réinitialiser le dossier de sortie
        self.sortie_path_label.config(
            text="Aucun dossier sélectionné",
            fg=self.text_light,
            font=("Segoe UI", 10)
        )
        
        # Cacher la barre de progression si elle est affichée
        self.progress_frame.pack_forget()
        
        # Réinitialiser les statuts visuels
        self.status_radio = None
        self.status_tele = None
        self.status_manuelle = None
        self.status_sortie = None
        
        # Mettre à jour les icônes de statut
        self.update_file_status('fichier_radio', None)
        self.update_file_status('fichier_tele', None)
        self.update_file_status('fichier_manuelle', None)
        self.update_file_status('dossier_sortie', None)
        
        messagebox.showinfo("✅ Refresh", "L'application a été remise à zéro !")
    
    def on_drop(self, event, file_attr, label):
        """Gère le drop de fichier avec validation et feedback visuel"""
        filepath = event.data.strip('{}')
        if filepath.endswith(('.xlsx', '.xls')):
            # Validation basique du fichier
            try:
                # Test rapide d'ouverture
                pd = lazy_import_pandas()
                df_test = pd.read_excel(filepath, nrows=5)  # Lire juste 5 lignes pour tester
                
                # Fichier valide
                setattr(self, file_attr, filepath)
                filename = os.path.basename(filepath)
                label.config(text=f"✓ {filename}", fg=self.primary_color, font=("Segoe UI", 9, "bold"))
                self.update_file_status(file_attr, 'selected')
            except Exception as e:
                # Fichier invalide
                setattr(self, file_attr, None)
                label.config(text=f"⚠️ Fichier non valide", fg="#FF6F00", font=("Segoe UI", 9, "bold"))
                self.update_file_status(file_attr, 'error')
                messagebox.showwarning("⚠️ Fichier invalide", f"Le fichier ne peut pas être lu :\n{str(e)}")
        else:
            # Format incorrect
            self.update_file_status(file_attr, 'error')
            messagebox.showerror("❌ Format incorrect", "Veuillez déposer un fichier Excel (.xlsx ou .xls)")
            
    def choisir_fichier(self, file_attr, label):
        """Ouvre le dialogue de sélection de fichier avec validation"""
        filepath = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")]
        )
        if filepath:
            # Même logique de validation que le drop
            try:
                pd = lazy_import_pandas()
                df_test = pd.read_excel(filepath, nrows=5)
                
                setattr(self, file_attr, filepath)
                filename = os.path.basename(filepath)
                label.config(text=f"✓ {filename}", fg=self.primary_color, font=("Segoe UI", 9, "bold"))
                self.update_file_status(file_attr, 'selected')
                
            except Exception as e:
                setattr(self, file_attr, None)
                label.config(text=f"⚠️ Fichier non valide", fg="#FF6F00", font=("Segoe UI", 9, "bold"))
                self.update_file_status(file_attr, 'error')
                messagebox.showwarning("⚠️ Fichier invalide", f"Le fichier ne peut pas être lu :\n{str(e)}")
            
    def choisir_dossier_sortie(self):
        """Ouvre le dialogue de sélection de dossier"""
        dossier = filedialog.askdirectory(title="Sélectionner le dossier de sortie")
        if dossier:
            self.dossier_sortie = dossier
            foldername = os.path.basename(dossier)
            self.sortie_path_label.config(text=f"✓ {foldername}", fg=self.primary_color, font=("Segoe UI", 10, "bold"))
            
    def lancer_analyse(self):
        """Lance l'analyse dans un thread séparé"""
        # Vérifications
        if not self.fichier_radio and not self.fichier_tele and not self.fichier_manuelle:
            messagebox.showerror(
                "❌ Sélection manquante",
                "Veuillez sélectionner au moins un fichier à analyser"
            )
            return
            
        if not self.dossier_sortie:
            messagebox.showerror(
                "❌ Dossier manquant",
                "Veuillez sélectionner un dossier de sortie"
            )
            return
        
        # Désactiver le bouton et afficher la progression améliorée
        self.analyze_btn.config(state="disabled")
        self.progress_frame.pack(fill="x", pady=20)
        self.progress_bar.config(value=0, maximum=100)
        self.progress_percent.config(text="0%")
        self.status_label.config(text="⏳ Initialisation de l'analyse...", fg=self.secondary_color)
        
        # Lancer l'analyse dans un thread
        thread = threading.Thread(target=self.executer_analyse)
        thread.start()
        
    def executer_analyse(self):
        """Exécute l'analyse (dans un thread séparé)"""
        try:
            # Import des modules lourds seulement quand nécessaire
            pd = lazy_import_pandas()
            datetime = lazy_import_datetime()
            logique_controles = lazy_import_logique()
            
            date_actuelle = datetime.now()
            date_str = date_actuelle.strftime('%Y_%B')
            
            # Créer le dossier de sortie
            os.makedirs(self.dossier_sortie, exist_ok=True)
            
            anomalies_tous_modes = []
            
            # Liste des tâches
            tasks = []
            if self.fichier_radio:
                tasks.append({
                    "nom": "Radioreleve",
                    "fichier": self.fichier_radio,
                    "fonction": logique_controles.check_data_radio,
                    "tab_type": "radio"
                })
            if self.fichier_tele:
                tasks.append({
                    "nom": "Telereleve",
                    "fichier": self.fichier_tele,
                    "fonction": logique_controles.check_data_tele,
                    "tab_type": "tele"
                })
            if self.fichier_manuelle:
                tasks.append({
                    "nom": "Manuelle",
                    "fichier": self.fichier_manuelle,
                    "fonction": logique_controles.check_data_manuelle,
                    "tab_type": "manuelle"
                })
            
            for i, task in enumerate(tasks):
                # Calcul du pourcentage
                progress = int((i / len(tasks)) * 100)
                self.update_progress(progress, f"📊 Analyse {task['nom']}... ({i+1}/{len(tasks)})")
                
                # Lire le fichier
                df = pd.read_excel(task["fichier"])
                df = df.iloc[:-2].copy()  # Supprimer les 2 dernières lignes
                
                # Analyser
                anomalies_df, anomaly_counter = task["fonction"](df)
                
                if not anomalies_df.empty:
                    # Rapport détaillé par mode
                    nom_rapport = f"Rapport_{task['nom']}_{date_str}.xlsx"
                    chemin_rapport = os.path.join(self.dossier_sortie, nom_rapport)
                    logique_controles.creer_rapport_excel_detaille(
                        chemin_rapport, anomalies_df, anomaly_counter, task["tab_type"]
                    )
                    
                    # Collecte pour regroupement Traités
                    tmp = anomalies_df.copy()
                    if "Traité" not in tmp.columns:
                        tmp["Traité"] = "NON_RENSEIGNE"
                    tmp["_Source_Mode"] = task["nom"]
                    anomalies_tous_modes.append(tmp)
            
            # Générer les rapports par Traité si nécessaire
            if anomalies_tous_modes:
                self.update_progress(85, "📁 Génération des rapports par Traité...")
                self.generer_rapports_traites(anomalies_tous_modes, date_str)
            
            # Finalisation
            self.update_progress(100, "✅ Analyse terminée !")
            
            # Succès
            self.root.after(0, self.on_analyse_complete)
            
        except Exception as e:
            self.root.after(0, lambda: self.on_analyse_error(str(e)))
            
    def generer_rapports_traites(self, anomalies_tous_modes, date_str):
        """Génère les rapports groupés par Traité"""
        from main import traite_key, create_excel_traite, COLONNES_A_SUPPR_TRAITES
        
        df_all = pd.concat(anomalies_tous_modes, ignore_index=True, sort=False)
        out_dir = os.path.join(self.dossier_sortie, "Traites")
        os.makedirs(out_dir, exist_ok=True)
        
        for code3, grp in df_all.groupby(df_all["Traité"].map(traite_key), dropna=False):
            drop_cols = ["_Source_Mode"] + [c for c in COLONNES_A_SUPPR_TRAITES if c in grp.columns]
            grp_to_save = grp.drop(columns=[c for c in drop_cols if c in grp.columns])
            
            sort_cols = [c for c in ["Traité", "Index original"] if c in grp_to_save.columns]
            if sort_cols:
                grp_to_save = grp_to_save.sort_values(sort_cols)
            
            nom_fic = f"Anomalies_Traite_{code3}_{date_str}.xlsx"
            chemin_out = os.path.join(out_dir, nom_fic)
            
            create_excel_traite(chemin_out, grp_to_save)
    
    def update_status(self, message):
        """Met à jour le message de statut (thread-safe)"""
        self.root.after(0, lambda: self.status_label.config(text=message, fg=self.secondary_color))
        
    def update_progress(self, percentage, message):
        """Met à jour la barre de progression avec pourcentage"""
        self.root.after(0, lambda: [
            self.progress_bar.config(value=percentage),
            self.progress_percent.config(text=f"{percentage}%"),
            self.status_label.config(text=message, fg=self.secondary_color)
        ])
        
    def on_analyse_complete(self):
        """Appelé quand l'analyse est terminée avec succès"""
        self.progress_bar.config(value=0)  # Reset de la barre
        self.progress_frame.pack_forget()
        self.analyze_btn.config(state="normal")
        
        # Message de succès avec design amélioré
        result = messagebox.askyesno(
            "✅ Analyse terminée !",
            f"Les rapports ont été générés avec succès !\n\n"
            f"Dossier : {self.dossier_sortie}\n\n"
            f"Voulez-vous ouvrir le dossier ?"
        )
        
        if result:
            # Ouvrir le dossier
            subprocess = lazy_import_subprocess()  # Import à la demande
            if sys.platform == "win32":
                os.startfile(self.dossier_sortie)
            elif sys.platform == "darwin":
                subprocess.run(["open", self.dossier_sortie])
            else:
                subprocess.run(["xdg-open", self.dossier_sortie])
                
    def on_analyse_error(self, error_message):
        """Appelé quand une erreur survient"""
        self.progress_bar.config(value=0)  # Reset de la barre
        self.progress_frame.pack_forget()
        self.analyze_btn.config(state="normal")
        
        messagebox.showerror(
            "❌ Erreur lors de l'analyse",
            f"Une erreur est survenue :\n\n{error_message}"
        )


def main():
    """Point d'entrée de l'application avec gestion d'erreur pour exe"""
    try:
        # Essayer d'importer tkinterdnd2
        DND_FILES, TkinterDnD = lazy_import_tkinterdnd()
        if DND_FILES and TkinterDnD:
            root = TkinterDnD.Tk()
        else:
            # Fallback sur tkinter standard si tkinterdnd2 ne fonctionne pas
            import tkinter as tk
            root = tk.Tk()
            print("⚠️ Mode sans drag & drop - Utilisez les boutons 'Parcourir'")
        
        app = AnalysesAnomaliesApp(root)
        root.mainloop()
        
    except Exception as e:
        import tkinter as tk
        import tkinter.messagebox as messagebox
        
        root = tk.Tk()
        root.withdraw()  # Cacher la fenêtre principale
        
        messagebox.showerror(
            "Erreur de démarrage",
            f"Erreur lors du lancement de l'application :\n\n{e}\n\n"
            f"Essayez d'exécuter en tant qu'administrateur ou vérifiez que tous les fichiers sont présents."
        )
        root.destroy()


if __name__ == "__main__":
    main()

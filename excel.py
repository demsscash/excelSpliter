import pandas as pd
import math
import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import threading

class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Splitter Avancé")
        self.root.geometry("900x900")
        self.root.resizable(True, True)
        
        # Configuration pour améliorer la réactivité
        self.root.config(takefocus=True)
        
        # Variables
        self.input_file_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()
        self.rows_per_file = tk.IntVar(value=500)  # Valeur par défaut: 500 lignes
        self.file_prefix = tk.StringVar(value="Template")
        self.split_mode = tk.StringVar(value="rows")  # Par défaut: division par lignes
        self.selected_column = tk.StringVar()
        
        # Données
        self.data = None
        self.columns = []
        
        # Définir le dossier Documents comme dossier de sortie par défaut
        documents_path = os.path.join(Path.home(), "Documents", "ExcelSplitter")
        self.output_folder_path.set(documents_path)
        
        # Création de l'interface
        self.create_widgets()
        
        # Définir le focus sur la fenêtre principale
        self.root.focus_force()
        
        # Capture des événements de clic pour s'assurer que toute la fenêtre est cliquable
        self.root.bind("<Button-1>", self.handle_click)
        
        # Création d'un événement virtuel pour la propagation du focus
        self.root.event_add('<<PropagateFocus>>', '<Button-1>')
        
        # Configuration pour améliorer la réactivité des boutons
        self.root.bind_class('Button', '<ButtonPress-1>', lambda e: e.widget.invoke())
        self.root.bind_class('TButton', '<ButtonPress-1>', lambda e: e.widget.invoke())
        
    def handle_click(self, event):
        """Gère les clics dans la fenêtre pour s'assurer que le focus reste sur la fenêtre"""
        # Assurez-vous que le clic est traité normalement
        # puis redonner le focus à la fenêtre principale
        self.root.focus_set()
        
        # Propagation de l'événement pour assurer que tous les widgets reçoivent le clic
        event.widget.event_generate('<<PropagateFocus>>', when='tail')
    
    def create_widgets(self):
        # Style
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=5)
        style.configure("TLabelframe", padding=10)
        style.configure("TRadiobutton", padding=5)
        
        # Créer un style accentué pour le bouton Charger
        style.configure("Accent.TButton", background="#4e73df", foreground="white")
        
        # Création d'un cadre principal avec défilement
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas pour permettre le défilement
        main_canvas = tk.Canvas(main_frame)
        main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # S'assurer que le canvas reçoit aussi les clics
        main_canvas.bind("<Button-1>", self.handle_click)
        
        # Barre de défilement
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=main_canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        main_canvas.configure(yscrollcommand=scrollbar.set)
        main_canvas.bind('<Configure>', lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        
        # Frame interne pour le contenu
        content_frame = ttk.Frame(main_canvas)
        main_canvas.create_window((0, 0), window=content_frame, anchor="nw", width=880)
        
        # S'assurer que le contenu reçoit aussi les clics
        content_frame.bind("<Button-1>", self.handle_click)
        
        # Section sélection de fichier d'entrée
        input_frame = ttk.LabelFrame(content_frame, text="Fichier Excel à diviser")
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Fichier:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        # Créer un sous-cadre pour le champ de saisie et les boutons
        input_buttons_frame = ttk.Frame(input_frame)
        input_buttons_frame.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        input_buttons_frame.columnconfigure(0, weight=1)  # Le champ de saisie s'étendra
        
        ttk.Entry(input_buttons_frame, textvariable=self.input_file_path,width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(input_buttons_frame, text="Parcourir...", command=self.browse_input_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(input_buttons_frame, text="Charger", command=self.load_file).pack(side=tk.LEFT)
        
        # Section mode de division
        self.split_mode_frame = ttk.LabelFrame(content_frame, text="Mode de division")
        self.split_mode_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Radiobutton(self.split_mode_frame, text="Division par nombre de lignes", variable=self.split_mode, 
                        value="rows", command=self.update_ui_for_mode).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Radiobutton(self.split_mode_frame, text="Division par valeurs de colonne", variable=self.split_mode, 
                        value="column", command=self.update_ui_for_mode).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Section configuration par lignes
        self.rows_config_frame = ttk.LabelFrame(content_frame, text="Configuration (division par lignes)")
        self.rows_config_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.rows_config_frame, text="Nombre de lignes par fichier:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Spinbox(self.rows_config_frame, from_=1, to=10000, textvariable=self.rows_per_file, width=10).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Section configuration par colonne
        self.column_config_frame = ttk.LabelFrame(content_frame, text="Configuration (division par colonne)")
        
        ttk.Label(self.column_config_frame, text="Sélectionnez la colonne:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.column_combobox = ttk.Combobox(self.column_config_frame, textvariable=self.selected_column, state="readonly", width=30)
        self.column_combobox.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Tableau d'aperçu des valeurs de colonne
        preview_frame = ttk.LabelFrame(self.column_config_frame, text="Aperçu des valeurs distinctes")
        preview_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # Création d'un Treeview pour afficher les valeurs distinctes
        columns = ("valeur", "nombre")
        self.preview_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=8)
        self.preview_tree.heading("valeur", text="Valeur distincte")
        self.preview_tree.heading("nombre", text="Nombre d'occurrences")
        self.preview_tree.column("valeur", width=450)  # Augmentation de la largeur
        self.preview_tree.column("nombre", width=200, anchor=tk.CENTER)  # Augmentation de la largeur
        
        # Ajouter une barre de défilement pour l'aperçu
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=preview_scroll.set)
        
        # Placement du Treeview et de la barre de défilement
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        # Section préfixe de fichier (commun aux deux modes)
        prefix_frame = ttk.LabelFrame(content_frame, text="Paramètres communs")
        prefix_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(prefix_frame, text="Préfixe des fichiers:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(prefix_frame, textvariable=self.file_prefix, width=20).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Section dossier de sortie
        output_frame = ttk.LabelFrame(content_frame, text="Dossier de sortie")
        output_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(output_frame, text="Dossier:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        # Créer un sous-cadre pour le champ de saisie et le bouton
        output_buttons_frame = ttk.Frame(output_frame)
        output_buttons_frame.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        output_buttons_frame.columnconfigure(0, weight=1)  # Le champ de saisie s'étendra
        
        ttk.Entry(output_buttons_frame, textvariable=self.output_folder_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        browse_output_button = ttk.Button(output_buttons_frame, text="Parcourir...", command=self.browse_output_folder)
        browse_output_button.pack(side=tk.LEFT)
        browse_output_button.bind("<ButtonRelease-1>", lambda e: self.browse_output_folder())
        
        # Barre de progression
        progress_frame = ttk.Frame(content_frame)
        progress_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        # Section statut
        status_frame = ttk.Frame(content_frame)
        status_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_var = tk.StringVar(value="Prêt. Veuillez charger un fichier Excel.")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor=tk.W)
        self.status_label.pack(fill=tk.X, padx=5)
        
        # Bouton pour exécuter
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.split_button = ttk.Button(button_frame, text="Diviser le fichier Excel", command=self.split_excel, state=tk.DISABLED)
        self.split_button.pack(side=tk.RIGHT, padx=5)
        self.split_button.bind("<ButtonRelease-1>", lambda e: self.split_excel())
        
        # Configuration initiale de l'interface selon le mode
        self.update_ui_for_mode()
    
    def update_ui_for_mode(self):
        """Met à jour l'interface selon le mode de division sélectionné"""
        if self.split_mode.get() == "rows":
            # Afficher la configuration par lignes
            if hasattr(self, 'column_config_frame'):
                self.column_config_frame.pack_forget()
            self.rows_config_frame.pack(fill=tk.X, padx=5, pady=5, after=self.split_mode_frame)
        else:
            # Afficher la configuration par colonne
            self.rows_config_frame.pack_forget()
            self.column_config_frame.pack(fill=tk.X, padx=5, pady=5, after=self.split_mode_frame)
    
    def browse_input_file(self):
        """Ouvre une boîte de dialogue pour sélectionner le fichier Excel d'entrée"""
        filetypes = [("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")]
        filename = filedialog.askopenfilename(filetypes=filetypes, title="Sélectionner un fichier Excel")
        if filename:
            self.input_file_path.set(filename)
            # Remettre le focus sur la fenêtre principale après la sélection du fichier
            self.root.focus_force()
    
    def browse_output_folder(self):
        """Ouvre une boîte de dialogue pour sélectionner le dossier de sortie"""
        folder = filedialog.askdirectory(title="Sélectionner un dossier de sortie")
        if folder:
            self.output_folder_path.set(folder)
            # Remettre le focus sur la fenêtre principale après la sélection du dossier
            self.root.focus_force()
    
    def load_file(self):
        """Charge le fichier Excel et extrait les colonnes"""
        if not self.input_file_path.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel d'entrée.")
            return
        
        try:
            self.update_status("Chargement du fichier Excel...")
            
            # Charger le fichier dans un DataFrame pandas
            self.data = pd.read_excel(self.input_file_path.get())
            self.columns = list(self.data.columns)
            
            # Mettre à jour la liste déroulante des colonnes
            self.column_combobox['values'] = self.columns
            if self.columns:
                self.column_combobox.current(0)  # Sélectionner la première colonne par défaut
                self.update_column_preview()  # Mettre à jour l'aperçu pour la colonne sélectionnée
            
            # Activer le bouton de division
            self.split_button['state'] = tk.NORMAL
            
            # Mise à jour du statut
            num_rows = len(self.data)
            num_cols = len(self.columns)
            self.update_status(f"Fichier chargé : {num_rows} lignes, {num_cols} colonnes")
            
            # Lier l'événement de changement de colonne
            self.column_combobox.bind("<<ComboboxSelected>>", lambda e: self.update_column_preview())
            
            # Remettre le focus sur la fenêtre principale
            self.root.after(100, self.root.focus_force)
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger le fichier : {str(e)}")
            self.update_status(f"Erreur : {str(e)}")
            # Remettre le focus sur la fenêtre principale après l'erreur
            self.root.focus_force()
    
    def update_column_preview(self):
        """Met à jour l'aperçu des valeurs distinctes pour la colonne sélectionnée"""
        if not self.data is None and self.selected_column.get():
            try:
                # Effacer le tableau d'aperçu
                for item in self.preview_tree.get_children():
                    self.preview_tree.delete(item)
                
                # Obtenir les valeurs distinctes et leur nombre d'occurrences
                col_name = self.selected_column.get()
                value_counts = self.data[col_name].value_counts().reset_index()
                value_counts.columns = ['value', 'count']
                
                # Limiter à 100 valeurs pour des raisons de performance
                max_display = 100
                if len(value_counts) > max_display:
                    displayed_counts = value_counts.iloc[:max_display]
                    total_distinct = len(value_counts)
                    
                    # Ajouter les valeurs au tableau
                    for _, row in displayed_counts.iterrows():
                        self.preview_tree.insert("", tk.END, values=(row['value'], row['count']))
                    
                    # Ajouter une entrée indiquant qu'il y a plus de valeurs
                    more_values = total_distinct - max_display
                    self.preview_tree.insert("", tk.END, values=(f"... {more_values} autres valeurs ...", ""))
                else:
                    # Ajouter toutes les valeurs au tableau
                    for _, row in value_counts.iterrows():
                        self.preview_tree.insert("", tk.END, values=(row['value'], row['count']))
                
                num_distinct = len(value_counts)
                self.update_status(f"Colonne '{col_name}' a {num_distinct} valeurs distinctes")
                
                # Remettre le focus sur la fenêtre principale
                self.root.focus_force()
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'analyse de la colonne : {str(e)}")
                self.update_status(f"Erreur : {str(e)}")
                # Remettre le focus sur la fenêtre principale après l'erreur
                self.root.focus_force()
    
    def update_status(self, message):
        """Met à jour le message de statut"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def split_by_rows(self, output_folder, date_aujourdhui):
        """Divise le fichier par nombre de lignes"""
        try:
            count = len(self.data)
            rows_per_file = self.rows_per_file.get()
            no_of_files = math.ceil(count/rows_per_file)
            
            self.update_status(f"Divisant {count} lignes en {no_of_files} fichiers...")
            self.progress_var.set(0)
            
            for x in range(no_of_files):
                # Calcul des indices de début et de fin
                start_row = x * rows_per_file
                end_row = min((x + 1) * rows_per_file, count)
                
                # Extraction des données pour ce fichier
                new_data = self.data.iloc[start_row:end_row]
                
                # Création du nom de fichier
                output_file = os.path.join(
                    output_folder, 
                    f"{self.file_prefix.get()}_{date_aujourdhui}_{x}.xlsx"
                )
                
                # Sauvegarde du fichier
                new_data.to_excel(output_file, index=False)
                
                # Mise à jour de la barre de progression
                progress = (x + 1) / no_of_files * 100
                self.progress_var.set(progress)
                self.update_status(f"Création du fichier {x+1}/{no_of_files} : {os.path.basename(output_file)}")
                self.root.update_idletasks()
            
            return no_of_files
        
        except Exception as e:
            raise Exception(f"Erreur lors de la division par lignes : {str(e)}")
    
    def split_by_column(self, output_folder, date_aujourdhui):
        """Divise le fichier par valeurs de colonne"""
        try:
            col_name = self.selected_column.get()
            unique_values = self.data[col_name].unique()
            no_of_files = len(unique_values)
            
            self.update_status(f"Divisant selon les {no_of_files} valeurs de la colonne '{col_name}'...")
            self.progress_var.set(0)
            
            for i, value in enumerate(unique_values):
                # Filtrer les données pour cette valeur
                filtered_data = self.data[self.data[col_name] == value]
                
                # Création du nom de fichier (valider le nom de fichier)
                safe_value = str(value).replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace("\"", "_").replace("<", "_").replace(">", "_").replace("|", "_")
                if len(safe_value) > 50:  # Limiter la longueur du nom de fichier
                    safe_value = safe_value[:50]
                
                output_file = os.path.join(
                    output_folder, 
                    f"{self.file_prefix.get()}_{date_aujourdhui}_{safe_value}.xlsx"
                )
                
                # Sauvegarde du fichier
                filtered_data.to_excel(output_file, index=False)
                
                # Mise à jour de la barre de progression
                progress = (i + 1) / no_of_files * 100
                self.progress_var.set(progress)
                self.update_status(f"Création du fichier {i+1}/{no_of_files} : {os.path.basename(output_file)}")
                self.root.update_idletasks()
            
            return no_of_files
        
        except Exception as e:
            raise Exception(f"Erreur lors de la division par colonne : {str(e)}")
    
    def split_excel_thread(self):
        """Fonction exécutée dans un thread séparé pour diviser le fichier Excel"""
        try:
            # Désactiver le bouton pendant le traitement
            self.split_button['state'] = tk.DISABLED
            
            # Vérification des entrées
            if not self.input_file_path.get():
                messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel d'entrée.")
                return
            
            if self.data is None:
                messagebox.showerror("Erreur", "Veuillez d'abord charger le fichier.")
                return
            
            # Vérifier le mode et les paramètres associés
            if self.split_mode.get() == "column" and not self.selected_column.get():
                messagebox.showerror("Erreur", "Veuillez sélectionner une colonne pour la division.")
                return
            
            # Création du dossier de sortie s'il n'existe pas
            output_folder = self.output_folder_path.get()
            os.makedirs(output_folder, exist_ok=True)
            
            # Obtenir la date d'aujourd'hui au format AAAAMMJJ
            date_aujourdhui = datetime.datetime.now().strftime("%Y%m%d")
            
            # Division selon le mode sélectionné
            if self.split_mode.get() == "rows":
                no_of_files = self.split_by_rows(output_folder, date_aujourdhui)
                mode_str = "par nombre de lignes"
            else:  # column
                no_of_files = self.split_by_column(output_folder, date_aujourdhui)
                mode_str = f"par valeurs de la colonne '{self.selected_column.get()}'"
            
            # Terminé
            self.update_status(f"Terminé ! {no_of_files} fichiers créés dans {output_folder}")
            messagebox.showinfo("Succès", f"Division {mode_str} terminée !\n{no_of_files} fichiers ont été créés dans le dossier :\n{output_folder}")
            
            # Remettre le focus sur la fenêtre principale après la division
            self.root.focus_force()
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur est survenue : {str(e)}")
            self.update_status(f"Erreur : {str(e)}")
            # Remettre le focus sur la fenêtre principale après l'erreur
            self.root.focus_force()
        
        finally:
            # Réactiver le bouton
            self.split_button['state'] = tk.NORMAL
    
    def split_excel(self):
        """Lance la division du fichier Excel dans un thread séparé"""
        threading.Thread(target=self.split_excel_thread, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    # Définir le focus sur la fenêtre principal au lancement
    root.focus_force()
    root.mainloop()
# Librairies
import pandas as pd
import customtkinter
from customtkinter import * 
from tkinter import filedialog as fd
from PIL import Image
import os
import sys

# Change le répertoire courant pour être celui du fichier en cours
os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

# Création de l'application
app = customtkinter.CTk()
app.title("ESME Market")
app.geometry("1000x720")
set_appearance_mode("light")

df = None  # Variable pour stocker les données

# Création de l'image de fond
fond = customtkinter.CTkImage(Image.open("fond.png"),size=(1000,720)) 
image = customtkinter.CTkLabel(master=app, text="",image=fond)
image.pack()
image.place(x=0, y=0)

# Création de l'image de fond blanc
gris = customtkinter.CTkImage(Image.open("fond blanc.png"),size=(860,660)) 
image2 = customtkinter.CTkLabel(master=app, text="",image=gris)
image2.pack()
image2.place(x=70, y=30)

# Création du logo ESME
logo = customtkinter.CTkImage(Image.open("Logo ESME.png"),size=(80,80)) 
image3 = customtkinter.CTkLabel(master=app, text="",image=logo)
image3.pack()
image3.place(x=75, y=34)

# Création du logo ESME Market
logo_esme_market = customtkinter.CTkImage(Image.open("ESME Market.png"),size=(160,120)) 
image4 = customtkinter.CTkLabel(master=app, text="",image=logo_esme_market)
image4.pack()
image4.place(x=758, y=40)

# Création de la textbox pour afficher le contenu du fichier
textbox = customtkinter.CTkTextbox(app, width=800, height=400,fg_color="#908383",text_color="white")
textbox.place(relx=0.5, rely=0.65, anchor="center")

bienvenue= customtkinter.CTkLabel(app, text="Bienvenue sur ESME Market !", font=("Arial", 30))
bienvenue.pack(pady=47, padx=10,anchor="center")

# Création de la boîte déroulante (cachée au début)
frame_options = customtkinter.CTkFrame(app, width=400, height=200)
label_options = customtkinter.CTkLabel(frame_options, text="Que voulez-vous faire ?", font=("Arial", 16))

# Création de la liste déroulante
options_menu = customtkinter.CTkOptionMenu(frame_options, values=["Rechercher une donnée", "Modifier une vente"], width=300, height=50, fg_color="#908383", button_color="#908383",button_hover_color="#6f6565")
options_menu.place(relx=0.5, rely=0.1, anchor="center")



# Cacher la boîte déroulante au démarrage
frame_options.pack_forget()
label_options.pack_forget()
options_menu.pack_forget()

# Fonction déclenchée lors de la sélection d'une option
def option_choisie(choix):
    if choix == "Rechercher une donnée":
        ouvrir_fenetre_recherche()
    elif choix == "Modifier une vente":
        ouvrir_fenetre_modification()
    elif choix == "Statistiques" :
        afficher_statistiques(df)

# Associer la fonction à la sélection d'une option
options_menu.configure(command=option_choisie)

# Fonction pour ouvrir et charger un fichier
def selection_fichier():
    """Charge un fichier CSV contenant les données des ventes et les stocke dans un DataFrame global."""
    global df  
    filetypes = (('Excel files', '*.xlsx'), ('CSV', '*.csv'), ('All files', '*.*'))
    filename = fd.askopenfilename(title='Ouvrir le fichier', initialdir='/', filetypes=filetypes)
    
    if filename.endswith('.xlsx'):
        df = pd.read_excel(filename)
    elif filename.endswith('.csv'):
        df = pd.read_csv(filename)

    df['Order ID'] = df['Order ID'].astype(str)  # Convertir Order ID en chaîne

    # Affiche les données dans la textbox
    textbox.delete("1.0", customtkinter.END) 
    textbox.insert(customtkinter.END, df.to_string())


    # Cacher le bouton "Ouvrir le fichier"
    bouton_selection_fichier.place_forget()

    # Afficher la liste déroulante au centre
    frame_options.place(relx=0.5, rely=0.25, anchor="center")
    label_options.pack(pady=10)
    options_menu.pack(pady=10)

# Création du bouton pour ouvrir un fichier
bouton_selection_fichier = customtkinter.CTkButton(app, text='Ouvrir le fichier', font=("helvetica",16), command=selection_fichier, width=300, height=60, fg_color="#908383",hover_color="#6f6565")
bouton_selection_fichier.place(relx=0.5, rely=0.25, anchor="center")  # Affiché au début

# Fenêtre pour modifier une donnée
def ouvrir_fenetre_modification():
    fenetre_modification = customtkinter.CTkToplevel(app)
    fenetre_modification.title("Modifier une donnée")
    fenetre_modification.geometry("1000x720")
    
    # Création de l'image de fond
    fond = customtkinter.CTkImage(Image.open("fond.png"),size=(1000,720)) 
    image = customtkinter.CTkLabel(master=fenetre_modification, text="",image=fond)
    image.pack()
    image.place(x=0, y=0)

    # Création de l'image de fond blanc
    gris = customtkinter.CTkImage(Image.open("fond blanc.png"),size=(860,660)) 
    image2 = customtkinter.CTkLabel(master=fenetre_modification, text="",image=gris)
    image2.pack()
    image2.place(x=70, y=30)

    # Création du logo ESME
    logo = customtkinter.CTkImage(Image.open("Logo ESME.png"),size=(80,80)) 
    image3 = customtkinter.CTkLabel(master=fenetre_modification, text="",image=logo)
    image3.pack()
    image3.place(x=75, y=34)

    # Création du logo ESME Market
    logo_esme_market = customtkinter.CTkImage(Image.open("ESME Market.png"),size=(160,120)) 
    image4 = customtkinter.CTkLabel(master=fenetre_modification, text="",image=logo_esme_market)
    image4.pack()
    image4.place(x=758, y=40)

    # Création du label pour apporter une précision
    label_id = customtkinter.CTkLabel(fenetre_modification, text="Entrez l'Order ID :", font=("Helvetica", 16))
    label_id.pack(pady=32)

    # Création de l'entrée pour l'ID
    entry_id = customtkinter.CTkEntry(fenetre_modification)
    entry_id.pack(pady=10)
    texte_ligne = customtkinter.CTkTextbox(fenetre_modification, width=600, height=100,fg_color="#908383",text_color="white")
    texte_ligne.pack(pady=10)
    
    # Création du menu Option
    ligne_selectionnee = StringVar(None)
    ligne_menu = customtkinter.CTkOptionMenu(fenetre_modification, values=[], variable=ligne_selectionnee)
    ligne_menu.pack(pady=15)
    ligne_menu.pack_forget()


    def afficher_ligne():
        if df is not None:
            order_id = entry_id.get().strip()
            if not order_id.isdigit():  # Vérifier que l'ID est un nombre
                texte_ligne.delete("1.0", customtkinter.END)
                texte_ligne.insert(customtkinter.END, "L'ID doit être un nombre")
                return
            
            resultats = df[df["Order ID"] == order_id]
            
            if not resultats.empty:  # Si l'ID est trouvé
                texte_ligne.delete("1.0", customtkinter.END)  # Effacer le contenu actuel
                texte_ligne.insert(customtkinter.END, resultats.to_string())  # Afficher les résultats essentiels
                
                # Crée une liste des lignes pour le menu déroulant
                valeur_lignes = [f"Ligne {zizi}: {row['Product']} - {row['Order Date']}" for zizi, row in resultats.iterrows()]
                ligne_menu.configure(values=valeur_lignes)
                ligne_menu.pack(pady=10)
                ligne_selectionnee.set(valeur_lignes[0])  # sélectionner la première ligne par défaut
            else:
                texte_ligne.delete("1.0", customtkinter.END)
                texte_ligne.insert(customtkinter.END, "ID non trouvé")
                ligne_menu.pack_forget()

    bouton_afficher = customtkinter.CTkButton(fenetre_modification, text="Afficher", command=afficher_ligne, fg_color="#908383",hover_color="#6f6565", width=200, height=40)
    bouton_afficher.pack(pady=5)

    label_quantite = customtkinter.CTkLabel(fenetre_modification, text="Nouvelle quantité :", font=("Helvetica", 16))
    label_quantite.pack(pady=8)
    entry_quantite = customtkinter.CTkEntry(fenetre_modification, width=200, height=40)
    entry_quantite.pack(pady=8)

    label_prix = customtkinter.CTkLabel(fenetre_modification, text="Nouveau prix :", font=("Helvetica", 16))
    label_prix.pack(pady=8)
    entry_prix = customtkinter.CTkEntry(fenetre_modification, width=200, height=40)
    entry_prix.pack(pady=8)    

    def modifier_donnees():
        if df is not None:
            ligne_selectionnee_valeur = ligne_selectionnee.get()
            index_ligne = int(ligne_selectionnee_valeur.split(":")[0].replace("Ligne", "").strip())
            if index_ligne in df.index:
                if entry_quantite.get():
                    df.at[index_ligne, "Quantity Ordered"] = int(entry_quantite.get())
                if entry_prix.get():
                    df.at[index_ligne, "Price Each"] = float(entry_prix.get())
                
                afficher_ligne()  # Rafraîchir l'affichage

    bouton_modifier = customtkinter.CTkButton(fenetre_modification, text="Modifier", command=modifier_donnees, width=200, height=40, fg_color="#908383",hover_color="#6f6565")
    bouton_modifier.pack(pady=8)

    def ajouter_vente(): # Ajouter une nouvelle vente
        if df is not None:
            resultats = df[df["Order ID"] == entry_id.get().strip()] # Vérifier si l'ID existe déjà
            if not resultats.empty:
                nouvelle_ligne = {
                    "Order ID": entry_id.get().strip(),
                    "Product": resultats.iloc[0]["Product"],
                    "Quantity Ordered": int(entry_quantite.get()),
                    "Price Each": float(entry_prix.get()),
                    "Order Date": resultats.iloc[0]["Order Date"],
                    "Purchase Address": resultats.iloc[0]["Purchase Address"]
                }
                df.loc[len(df)] = nouvelle_ligne
                afficher_ligne()

    bouton_ajouter = customtkinter.CTkButton(fenetre_modification, text="Ajouter vente", command=ajouter_vente, width=200, height=40, fg_color="#908383",hover_color="#6f6565")
    bouton_ajouter.pack(pady=10)

    def sauvegarder_donnees():
        if df is not None:
            df.to_csv("ventes_modifiees.csv", index=False)
            customtkinter.CTkLabel(fenetre_modification, text="Données sauvegardées !", text_color="green").pack(pady=10)

    bouton_sauvegarder = customtkinter.CTkButton(fenetre_modification, text="Sauvegarder", command=sauvegarder_donnees, width=200, height=40, fg_color="#908383",hover_color="#6f6565")
    bouton_sauvegarder.pack(pady=10)



def ouvrir_fenetre_recherche():
    """
    Ouvre une fenêtre de recherche permettant à l'utilisateur d'effectuer des recherches selon la date,
    un produit spécifique ou d'afficher le produit le plus vendu.
    """
    global df  # On vérifie que df est bien accessible

    # Fenêtre de recherche
    fenetre_recherche = customtkinter.CTkToplevel(app)
    fenetre_recherche.title("Chercher une donnée")
    fenetre_recherche.geometry("1000x720")

    fond = customtkinter.CTkImage(Image.open("fond.png"),size=(1000,720)) 
    image = customtkinter.CTkLabel(master=fenetre_recherche, text="",image=fond)
    image.pack()
    image.place(x=0, y=0)


    gris = customtkinter.CTkImage(Image.open("fond blanc.png"),size=(860,660)) 
    image2 = customtkinter.CTkLabel(master=fenetre_recherche, text="",image=gris)
    image2.pack()
    image2.place(x=70, y=30)

    logo = customtkinter.CTkImage(Image.open("Logo ESME.png"),size=(80,80)) 
    image3 = customtkinter.CTkLabel(master=fenetre_recherche, text="",image=logo)
    image3.pack()
    image3.place(x=75, y=34)

    logo_esme_market = customtkinter.CTkImage(Image.open("ESME Market.png"),size=(160,120)) 
    image4 = customtkinter.CTkLabel(master=fenetre_recherche, text="",image=logo_esme_market)
    image4.pack()
    image4.place(x=758, y=40)



    # Zone où les résultats seront affichés
    frame_recherche = customtkinter.CTkScrollableFrame(fenetre_recherche, height=200, width=510, border_color="#484646", border_width=2, fg_color="#C4C3C3", corner_radius=10)
    frame_recherche.place(relx=0.5, rely=0.7,anchor="center")

    # Label d'instruction
    texte_recherche = customtkinter.CTkLabel(fenetre_recherche, text='Choisissez un type de recherche :',font=("helvetica",18))
    texte_recherche.place(relx=0.5, rely=0.1,anchor="center")

    # Liste déroulante pour choisir le type de recherche
    options_recherche = ["Recherche selon la date", "Recherche selon un produit", "Produit le plus vendu"]
    entree_produit = customtkinter.CTkOptionMenu(fenetre_recherche, width=400, height=50, values=options_recherche, fg_color="#908383", button_color="#908383",button_hover_color="#6f6565")
    entree_produit.place(relx=0.5, rely=0.2,anchor="center")



    # Champ de saisie (qui apparaît dynamiquement)
    entry_saisie = customtkinter.CTkEntry(fenetre_recherche, width=400, height=40)
    entry_saisie.place(relx=0.5, rely=0.3,anchor="center")
    
    # Fonction pour afficher le champ de saisie selon l'option choisie
    def afficher_champ_saisie(option):
        """
        Affiche ou cache le champ de saisie en fonction de l'option sélectionnée.
        """
        if option == "Produit le plus vendu":
            entry_saisie.place_forget()
        else :
            entry_saisie.place(relx=0.5, rely=0.25,anchor="center")  # On affiche l'entrée à une position spécifique
            entry_saisie.delete(0, customtkinter.END)  # On vide l'entrée à chaque changement

    # Associer la fonction à la sélection d'une option
    entree_produit.configure(command=afficher_champ_saisie)

    def produit_le_plus_vendu(df, frame_recherche):
        """
        Affiche le produit le plus vendu en fonction des données du DataFrame.
        """
    # Vérification des colonnes nécessaires
        if not all(col in df.columns for col in ['Product', 'Price Each', 'Quantity Ordered']):
            print("Les colonnes nécessaires sont absentes. Colonnes disponibles :", df.columns)
            return None

        # Nettoyer les colonnes 'Price Each' et 'Quantity Ordered' pour gérer les valeurs non numériques
        df['Price Each'] = pd.to_numeric(df['Price Each'], errors='coerce')  # Remplace les valeurs invalides par NaN
        df['Quantity Ordered'] = pd.to_numeric(df['Quantity Ordered'], errors='coerce')

        # Supprimer les lignes où 'Price Each' ou 'Quantity Ordered' est NaN
        df = df.dropna(subset=['Price Each', 'Quantity Ordered'])
        
        # Ajouter une colonne 'Revenue' pour stocker le chiffre d'affaire
        df.loc[:, 'Revenue'] = df['Price Each'] * df['Quantity Ordered']

        # Calculer le chiffre d'affaire pour chaque ligne (Price Each * Quantity Ordered)
        df['Revenue'] = df['Price Each'] * df['Quantity Ordered']

        # Agréger les données par produit : somme des quantités et du chiffre d'affaire
        produit_stats = df.groupby('Product').agg(
            total_ventes=pd.NamedAgg(column='Quantity Ordered', aggfunc='sum'),
            chiffre_affaire=pd.NamedAgg(column='Revenue', aggfunc='sum'),
            prix_unitaire=pd.NamedAgg(column='Price Each', aggfunc='first')  # Le prix est constant pour chaque produit
        ).reset_index()

        # Trouver le produit avec le plus grand nombre de ventes
        produit_stats = produit_stats.sort_values(by='total_ventes', ascending=False)

        # Sélectionner le produit avec le plus grand nombre de ventes
        produit_le_plus_vendu = produit_stats.iloc[0]

        for widget in frame_recherche.winfo_children():
            widget.destroy()  # Vider la zone des résultats avant d'afficher les nouveaux résultats

        # Ajouter les résultats dans le frame_recherche
        label_resultat = customtkinter.CTkLabel(frame_recherche, text=f"Produit le plus vendu : {produit_le_plus_vendu['Product']}", wraplength=480, text_color="black")
        label_resultat.pack(pady=2)

        label_quantite = customtkinter.CTkLabel(frame_recherche, text=f"Nombre de fois vendu : {produit_le_plus_vendu['total_ventes']}", wraplength=480, text_color="black")
        label_quantite.pack(pady=2)

        label_prix = customtkinter.CTkLabel(frame_recherche, text=f"Prix unitaire : {produit_le_plus_vendu['prix_unitaire']}€", wraplength=480, text_color="black")
        label_prix.pack(pady=2)

        label_chiffre_affaire = customtkinter.CTkLabel(frame_recherche, text=f"Chiffre d'affaire : {produit_le_plus_vendu['chiffre_affaire']}€", wraplength=480, text_color="black")
        label_chiffre_affaire.pack(pady=2)

    def rechercher_donnees():
        """
        Effectue une recherche en fonction du critère sélectionné et affiche les résultats.
        """
        resultats = pd.DataFrame()
        if df is None:
            return  # Vérifie que le fichier a bien été chargé

        critere = entree_produit.get()  # Type de recherche sélectionné
        valeur_recherche = entry_saisie.get().strip()

        if critere == "Produit le plus vendu":
            produit_le_plus_vendu(df, frame_recherche)
            return
        
        if not valeur_recherche:
            return  # On ne fait rien si le champ est vide

    # On vide la zone des résultats
        for widget in frame_recherche.winfo_children():
            widget.destroy()

        # Recherche selon le type sélectionné
        if critere == "Recherche selon la date":
            resultats = df[df["Order Date"].astype(str) == valeur_recherche]

        elif critere == "Recherche selon un produit":
            resultats = df[df["Product"].str.contains(valeur_recherche, case=False, na=False)]


        # Affichage des résultats dans frame_recherche
        if not resultats.empty:
            for index, row in resultats.iterrows():
                label_resultat = customtkinter.CTkLabel(frame_recherche, text=f"{row['Order ID']} - {row['Product']} - {row['Quantity Ordered']}- {row['Price Each']}- {row['Order Date']}- {row['Purchase Address']}", wraplength=480,text_color="black")
                label_resultat.pack(pady=2)
        else:
            label_vide = customtkinter.CTkLabel(frame_recherche, text="Aucun résultat trouvé.", text_color="red")
            label_vide.pack(pady=2)

    # Bouton pour déclencher la recherche
    bouton_recherche = customtkinter.CTkButton(fenetre_recherche, text="Rechercher", command=rechercher_donnees,width=300,height=60,fg_color="#908383",hover_color="#6f6565",font=("helvetica",18))
    bouton_recherche.place(relx=0.5, rely=0.4,anchor="center")

    

app.mainloop()  # Boucle principale pour afficher l'interface
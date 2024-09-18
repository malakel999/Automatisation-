import pandas as pd
import matplotlib.pyplot as plt

# Chemin du fichier Excel
fichier = 'Factures PR MFS 2024 4 check du 24 07 2024 (1).xlsx'

# Charger la feuille spécifique du fichier Excel contenant la colonne RA_SOC
df = pd.read_excel(fichier, sheet_name='encours 2024')  # Remplacer par le nom correct de la feuille

# Vérifier que la colonne RA_SOC existe dans le DataFrame
if 'RA_SOCL' not in df.columns:
    raise Exception("La colonne 'RA_SOCL' n'existe pas dans le fichier Excel.")

# Calculer le pourcentage de chaque valeur dans la colonne RA_SOC
pourcentages = df['RA_SOCL'].value_counts(normalize=True) * 100

# Personnalisation des couleurs
couleurs = plt.get_cmap('tab20').colors  # Palette de couleurs (20 couleurs distinctes)

# Mettre en avant le segment avec le plus grand pourcentage (explode)
explode = [0.1 if i == pourcentages.idxmax() else 0 for i in pourcentages.index]

# Création du graphique circulaire sans labels internes
plt.figure(figsize=(10, 7))  # Taille du graphique
plt.pie(pourcentages, labels=None, autopct='%1.1f%%', startangle=90,
        colors=couleurs, explode=explode, shadow=True, pctdistance=0.85)

# Ajouter un titre stylé
plt.title("Répartition des commandes par RA_SOCL", fontsize=16, fontweight='bold')

# Assurer que le graphique est circulaire
plt.axis('equal')

# Ajouter une légende externe avec les noms et les couleurs
plt.legend(pourcentages.index, title="RA_SOCL", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

# Sauvegarder le graphique en PDF
plt.savefig("repartition_commandes_RA_SOCL.pdf", format='pdf')

# Afficher le graphique avec la légende à côté
plt.tight_layout()  # Ajuster les marges pour éviter le chevauchement
plt.show()

import pandas as pd

# Créer un fichier Excel factice de 100 000 lignes
data = {'Colonne1': range(1, 100001), 'Colonne2': ['Valeur'] * 100000}
df = pd.DataFrame(data)
df.to_excel("test.xlsx", index=False)
print("Fichier test.xlsx créé avec succès.")
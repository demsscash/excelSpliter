# Excel Splitter Avancé

Une application GUI pour diviser des fichiers Excel en plusieurs fichiers plus petits, selon différents critères.

## Fonctionnalités

- **Deux modes de division** :
  - Par nombre de lignes (découpage fixe)
  - Par valeurs d'une colonne (un fichier par valeur unique)
- Aperçu des valeurs distinctes avec comptage
- Barre de progression pour les opérations longues
- Interface intuitive avec thème moderne

## Prérequis

- Python 3.6+
- Packages requis :
pandas
openpyxl
tkinter


## Installation

1. Clonez le dépôt ou téléchargez le fichier `excel_splitter.py`
2. Installez les dépendances :
pip install pandas openpyxl


## Utilisation

1. Lancez l'application :
python excel_splitter.py

2. Sélectionnez votre fichier Excel
3. Choisissez le mode de division :
- **Par lignes** : Définissez le nombre de lignes par fichier
- **Par colonne** : Sélectionnez la colonne à utiliser
4. Spécifiez le dossier de sortie et le préfixe des fichiers
5. Cliquez sur "Diviser le fichier Excel"

## Options avancées

- Le dossier de sortie par défaut est `Documents/ExcelSplitter/`
- Vous pouvez personnaliser le préfixe des fichiers générés
- L'aperçu des colonnes montre les 50 premières valeurs uniques

## Captures d'écran

*(Insérez ici des captures si disponibles)*

## Licence

MIT - Libre d'utilisation et de modification

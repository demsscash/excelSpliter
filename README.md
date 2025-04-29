# Excel Splitter Avancé

![Python](https://img.shields.io/badge/Python-3.6%2B-blue)
![Pandas](https://img.shields.io/badge/Pandas-1.0%2B-orange)
![License](https://img.shields.io/badge/License-MIT-green)

Une application GUI intuitive pour diviser des fichiers Excel volumineux en plusieurs fichiers plus petits selon différents critères.

## ✨ Fonctionnalités

- **Deux modes de division** :
  - 🔢 Par nombre de lignes (découpage fixe)
  - 📊 Par valeurs d'une colonne (un fichier par valeur unique)
- 👀 Aperçu des valeurs distinctes avec comptage
- 📊 Barre de progression pour les opérations longues
- 🎨 Interface moderne et intuitive
- 🚀 Traitement en arrière-plan pour ne pas bloquer l'interface
- 📂 Sélection automatique du dossier Documents/ExcelSplitter par défaut

## 📦 Prérequis

- Python 3.6 ou supérieur
- Packages requis :
  ```bash
  pandas >= 1.0.0
  openpyxl >= 3.0.0
  tkinter (inclus dans Python standard)

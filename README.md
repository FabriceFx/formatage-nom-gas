# Boîte à outils de formatage de texte GAS

![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

Ce projet est un script Google Apps Script conçu pour Google Sheets. Il déploie un menu "Outils Texte" permettant aux utilisateurs de formater rapidement des plages de données textuelles.

## Fonctionnalités clés

* **TOUT EN MAJUSCULES** : Convertit le texte sélectionné en majuscules.
* **tout en minuscules** : Convertit le texte sélectionné en minuscules.
* **Nom Propre** : Capitalise la première lettre de chaque mot (gère intelligemment les tirets `Jean-Pierre` et apostrophes `l'arbre`).
* **Nettoyage** : Supprime les espaces de début/fin (trim) et réduit les doubles espaces intra-mots.
* **Performance** : Utilise des opérations par lots (`getValues` / `setValues`) pour minimiser le temps d'exécution sur les grandes plages de données.
* **Robustesse** : Gestion des erreurs et feedback utilisateur via des messages "Toast" (notifications en bas à droite).

## Installation manuelle

1.  Ouvrez votre fichier **Google Sheets**.
2.  Allez dans le menu **Extensions** > **Apps Script**.
3.  Supprimez tout code existant dans le fichier `Code.gs` par défaut.
4.  Copiez et collez l'intégralité du script fourni dans `Code.js`.
5.  Enregistrez le projet (Icône disquette ou `Ctrl + S`).
6.  Rechargez votre feuille Google Sheets (F5).
7.  Le menu **"Outils Texte"** apparaîtra à droite du menu "Aide" après quelques secondes.

## Utilisation

1.  Sélectionnez une ou plusieurs cellules contenant du texte.
2.  Cliquez sur **Outils Texte** dans la barre de menu.
3.  Choisissez la transformation souhaitée.
4.  Une notification confirmera le nombre de cellules modifiées.

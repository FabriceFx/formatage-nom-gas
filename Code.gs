/**
 * @fileoverview Boîte à outils de formatage de texte pour Google Sheets.
 * Ce script ajoute un menu personnalisé permettant de modifier la casse 
 * et de nettoyer le contenu des cellules sélectionnées.
 * * @author Fabrice Faucheux
 * @version 1.1.0
 */

/**
 * Crée le menu personnalisé à l'ouverture du classeur.
 * Cette fonction s'exécute automatiquement.
 */
function onOpen() {
  const interfaceUtilisateur = SpreadsheetApp.getUi();
  interfaceUtilisateur.createMenu("Outils Texte")
    .addItem("TOUT EN MAJUSCULES", "passerEnMajuscules")
    .addItem("tout en minuscules", "passerEnMinuscules")
    .addItem("Nom Propre (Première lettre maj)", "passerEnNomPropre")
    .addSeparator()
    .addItem("Supprimer les espaces inutiles", "nettoyerEspaces")
    .addToUi();
}

/**
 * Convertit le texte des cellules sélectionnées en majuscules.
 */
function passerEnMajuscules() {
  appliquerTransformation_(texte => texte.toUpperCase());
}

/**
 * Convertit le texte des cellules sélectionnées en minuscules.
 */
function passerEnMinuscules() {
  appliquerTransformation_(texte => texte.toLowerCase());
}

/**
 * Convertit le texte des cellules sélectionnées en format "Nom Propre".
 * Gère les accents, les tirets et les apostrophes.
 */
function passerEnNomPropre() {
  appliquerTransformation_(texte => {
    const texteMinuscule = texte.toLowerCase();
    // Regex : Trouve le début de ligne ou un séparateur (espace, tiret, apostrophe)
    // suivi d'un caractère alphanumérique (y compris accents).
    return texteMinuscule.replace(/(?:^|[\s\-\'\"])([\w\u00C0-\u00FF])/g, caractere => caractere.toUpperCase());
  });
}

/**
 * Supprime les espaces en début/fin et réduit les espaces multiples à un seul.
 */
function nettoyerEspaces() {
  appliquerTransformation_(texte => texte.trim().replace(/\s+/g, ' '));
}

/**
 * Moteur principal de transformation de texte.
 * Applique une fonction de rappel à chaque cellule de la plage active.
 * Utilise des opérations de lecture/écriture par lots pour la performance.
 *
 * @param {function(string): string} fonctionDeTransformation - La fonction à appliquer sur chaque chaîne.
 * @private
 */
function appliquerTransformation_(fonctionDeTransformation) {
  try {
    const classeur = SpreadsheetApp.getActiveSpreadsheet();
    const plageActive = classeur.getActiveRange();

    if (!plageActive) {
      classeur.toast("Veuillez sélectionner une cellule ou une plage.", "Attention");
      return;
    }

    const valeurs = plageActive.getValues();
    let compteurModifications = 0;

    // Transformation des données en mémoire
    const nouvellesValeurs = valeurs.map(ligne => 
      ligne.map(cellule => {
        if (typeof cellule === 'string' && cellule !== "") {
          const nouvelleValeur = fonctionDeTransformation(cellule);
          
          if (nouvelleValeur !== cellule) {
            compteurModifications++;
            return nouvelleValeur;
          }
        }
        // Retourne la cellule intacte si ce n'est pas du texte ou s'il n'y a pas de changement
        return cellule;
      })
    );

    // Écriture par lot uniquement si nécessaire
    if (compteurModifications > 0) {
      plageActive.setValues(nouvellesValeurs);
      classeur.toast(`Opération terminée avec succès. ${compteurModifications} cellule(s) modifiée(s).`, "Succès", 4);
    } else {
      classeur.toast("Aucune modification n'était nécessaire sur la sélection.", "Info", 4);
    }

  } catch (erreur) {
    console.error(`Erreur dans la fonction appliquerTransformation_ : ${erreur.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Une erreur est survenue. Consultez les logs.", "Erreur");
  }
}

function creerBrouillonsRelances() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Réponses au formulaire 1");

  const donnees = feuille.getDataRange().getValues();
  const titres = donnees[0];

  const colPrenom = titres.indexOf("Prénom");
  const colNom = titres.indexOf("Nom");
  const colEmail = titres.indexOf("Email");
  const colARégler = titres.indexOf("A régler");
  const colRelance = titres.indexOf("Relance");

  if (colRelance === -1) {
    throw new Error("La colonne 'Relance' est introuvable.");
  }

  let nombreBrouillons = 0;

  for (let i = 1; i < donnees.length; i++) {

    const prenom = donnees[i][colPrenom];
    const nom = donnees[i][colNom];
    const email = donnees[i][colEmail];
    const aRegler = Number(donnees[i][colARégler]) || 0;
    const relance = donnees[i][colRelance];

    // On ignore :
    // - ceux qui ne doivent rien
    // - ceux qui n'ont pas d'adresse email
    // - ceux pour lesquels une relance a déjà été préparée
    if (aRegler <= 0 || !email || relance) {
      continue;
    }

    const sujet = "Règlement de la cotisation";

    const message =
      "Bonjour " + prenom + ",\n\n" +
      "Je me permets de revenir vers vous concernant le règlement de la cotisation.\n\n" +
      "D'après nos informations, il reste actuellement " +
      aRegler + " € à régler.\n\n" +
      "Si le règlement a déjà été effectué, merci de ne pas tenir compte de ce message.\n\n" +
      "Bonne journée,\n" +
      "Clément Vignaud";

    // Création du brouillon uniquement
    GmailApp.createDraft(email, sujet, message);

    // On indique dans le tableau qu'un brouillon a été préparé
    feuille.getRange(i + 1, colRelance + 1).setValue("Brouillon");

    nombreBrouillons++;

    Logger.log(
      "Brouillon créé pour : " +
      prenom + " " + nom +
      " | " + aRegler + " €"
    );
  }

  Logger.log(
    "===== TERMINÉ =====\n" +
    nombreBrouillons +
    " nouveau(x) brouillon(s) créé(s)."
  );
}
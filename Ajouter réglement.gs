function ajouterReglement(prenom, nom, montant) {

  const feuille = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Réponses au formulaire 1");

  const donnees = feuille.getDataRange().getValues();
  const titres = donnees[0];

  const colPrenom = titres.indexOf("Prénom");
  const colNom = titres.indexOf("Nom");
  const colRegle = titres.indexOf("Réglé");

  montant = Number(montant);

  if (!montant || montant <= 0) {
    throw new Error("Le montant doit être supérieur à 0.");
  }

  for (let i = 1; i < donnees.length; i++) {

    const prenomLigne =
      String(donnees[i][colPrenom] || "").trim();

    const nomLigne =
      String(donnees[i][colNom] || "").trim();

    if (
      prenomLigne.toLowerCase() === prenom.toLowerCase() &&
      nomLigne.toLowerCase() === nom.toLowerCase()
    ) {

      const ancienRegle =
        Number(donnees[i][colRegle]) || 0;

      const nouveauRegle =
        ancienRegle + montant;

      feuille
        .getRange(i + 1, colRegle + 1)
        .setValue(nouveauRegle);

      return {
        ancienRegle: ancienRegle,
        montantAjoute: montant,
        nouveauRegle: nouveauRegle
      };
    }
  }

  throw new Error("Adhérent introuvable.");
}
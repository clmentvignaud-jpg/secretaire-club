function creerHistoriqueReglements() {

  const classeur = SpreadsheetApp.getActiveSpreadsheet();

  let historique =
    classeur.getSheetByName("HISTORIQUE RÈGLEMENTS");

  if (!historique) {

    historique =
      classeur.insertSheet("HISTORIQUE RÈGLEMENTS");

    historique.getRange("A1:E1").setValues([[
      "Date",
      "Prénom",
      "Nom",
      "Montant",
      "Mode de paiement"
    ]]);

    historique.getRange("A1:E1")
      .setFontWeight("bold");

    historique.getRange("A:A")
      .setNumberFormat("dd/MM/yyyy HH:mm");

    historique.getRange("D:D")
      .setNumberFormat('0.00 "€"');

    historique.autoResizeColumns(1, 5);
  }

  Logger.log("Onglet HISTORIQUE RÈGLEMENTS prêt.");
}
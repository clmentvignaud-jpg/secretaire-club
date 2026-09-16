function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === "rechercher") {
    const resultats = rechercherAdherent(data.recherche);

    return ContentService
      .createTextOutput(JSON.stringify(resultats))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (data.action === "reglement") {
    const resultat = ajouterReglement(
      data.prenom,
      data.nom,
      data.montant
    );

    return ContentService
      .createTextOutput(JSON.stringify(resultat))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(
      JSON.stringify({
        erreur: "Action inconnue"
      })
    )
    .setMimeType(ContentService.MimeType.JSON);
}
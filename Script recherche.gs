function doGet(e) {

  if (e && e.parameter && e.parameter.api === "bridge") {

    const recherche = e.parameter.recherche || "";

    const donnees = rechercherAdherent(recherche);

    const modele =
      HtmlService.createTemplateFromFile("Pont");

    modele.donnees = donnees;

    return modele
      .evaluate()
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );
  }

  return HtmlService
    .createHtmlOutputFromFile("Recherche")
    .setTitle("Gestion des adhérents");
}


function rechercherAdherent(recherche) {

  const feuille = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Réponses au formulaire 1");

  const donnees = feuille.getDataRange().getValues();
  const titres = donnees[0];

  const colPrenom = titres.indexOf("Prénom");
  const colNom = titres.indexOf("Nom");
  const colRegle = titres.indexOf("Réglé");
  const colARégler = titres.indexOf("A régler");
  const colCertif = titres.indexOf("Certif médical");
  const colLicence = titres.indexOf("Licence");
  const colEmail = titres.indexOf("Email");
  const colTelephone = titres.indexOf("Numéro de téléphone");
  const colSalle = titres.indexOf("Quelle salle?");
  const colTarif = titres.indexOf("Tarifs");

  const rechercheMinuscule =
    recherche.toString().trim().toLowerCase();

  const resultats = [];

  for (let i = 1; i < donnees.length; i++) {

    const prenom =
      String(donnees[i][colPrenom] || "").trim();

    const nom =
      String(donnees[i][colNom] || "").trim();

    if (!prenom || !nom) {
      continue;
    }

    const nomComplet =
      (prenom + " " + nom).toLowerCase();

    if (nomComplet.includes(rechercheMinuscule)) {

      resultats.push({

        prenom: prenom,

        nom: nom,

        tarif:
          Number(donnees[i][colTarif]) || 0,

        regle:
          Number(donnees[i][colRegle]) || 0,

        aRegler:
          Number(donnees[i][colARégler]) || 0,

        certif:
          donnees[i][colCertif] || "",

        licence:
          donnees[i][colLicence] || "",

        email:
          donnees[i][colEmail] || "",

        telephone:
          donnees[i][colTelephone] || "",

        salle:
          donnees[i][colSalle] || ""
      });
    }
  }

  return resultats;
}


function doPost(e) {

  const data = JSON.parse(e.postData.contents);

  if (data.action === "rechercher") {

    const resultats =
      rechercherAdherent(data.recherche);

    return ContentService
      .createTextOutput(JSON.stringify(resultats))
      .setMimeType(ContentService.MimeType.JSON);
  }


  if (data.action === "reglement") {

    const resultat =
      ajouterReglement(
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
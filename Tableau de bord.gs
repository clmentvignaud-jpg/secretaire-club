function creerTableauDeBord() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = classeur.getSheetByName("Réponses au formulaire 1");

  let tableau = classeur.getSheetByName("TABLEAU DE BORD");

  if (!tableau) {
    tableau = classeur.insertSheet("TABLEAU DE BORD");
  }

  tableau.clear();

  const donnees = feuille.getDataRange().getValues();
  const titres = donnees[0];

  const colPrenom = titres.indexOf("Prénom");
  const colNom = titres.indexOf("Nom");
  const colARégler = titres.indexOf("A régler");
  const colReglé = titres.indexOf("Réglé");
  const colCertif = titres.indexOf("Certif médical");
  const colLicence = titres.indexOf("Licence");

  let nombreAdherents = 0;
  let totalRegle = 0;
  let totalARégler = 0;
  let personnesARelancer = 0;
  let certificatsManquants = 0;
  let licencesManquantes = 0;

  for (let i = 1; i < donnees.length; i++) {

    const prenom = String(donnees[i][colPrenom]).trim();
    const nom = String(donnees[i][colNom]).trim();

    // On ne compte comme adhérent que si prénom ET nom sont renseignés
    if (prenom && nom) {
      nombreAdherents++;
    }

    const regle = Number(donnees[i][colReglé]) || 0;
    const aRegler = Number(donnees[i][colARégler]) || 0;

    totalRegle += regle;
    totalARégler += aRegler;

    if (aRegler > 0 && prenom && nom) {
      personnesARelancer++;
    }

    const certif = String(donnees[i][colCertif]).trim().toLowerCase();
    const licence = String(donnees[i][colLicence]).trim().toLowerCase();

    if (certif === "non" && prenom && nom) {
      certificatsManquants++;
    }

    if (licence === "non" && prenom && nom) {
      licencesManquantes++;
    }
  }

  const totalCotisations = totalRegle + totalARégler;

  // Titre
  tableau.getRange("A1:B1").merge();
  tableau.getRange("A1")
    .setValue("TABLEAU DE BORD — COTISATIONS")
    .setFontWeight("bold")
    .setFontSize(16);

  // Indicateurs
  const indicateurs = [
    ["Nombre d'adhérents", nombreAdherents],
    ["Total des cotisations", totalCotisations],
    ["Total réglé", totalRegle],
    ["Total restant à régler", totalARégler],
    ["Personnes à relancer", personnesARelancer],
    ["Certificats manquants", certificatsManquants],
    ["Licences manquantes", licencesManquantes]
  ];

  tableau.getRange(3, 1, indicateurs.length, 2).setValues(indicateurs);

  tableau.getRange("A3:A9").setFontWeight("bold");

  // Format monétaire
  tableau.getRange("B4:B6").setNumberFormat('0.00 "€"');

  tableau.autoResizeColumns(1, 2);

  Logger.log("Tableau de bord mis à jour.");
}
function miseAJourAutomatique() {
  creerTableauDeBord();
}
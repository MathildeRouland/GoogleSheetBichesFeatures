
// Déclaration des globale variables
var fichier
var objet
var dossier
var corps
var reponse
var msgBoxAlert
var aliases
var ui
var me
var n
var docID = "https://docs.google.com/spreadsheets/d/IDduDocument/";  // NE PAS OUBLIER LE / à la fin
var d
const doc = SpreadsheetApp.getActive();

var  feuille = 'Convocation';                                // feuille ou chercher le titre du document
var  feuilleID = 'IDdeLaFeuilleAEnvoyer';                                //Feuille à envoyer 
const sheetConvocation = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(feuille);

var feuilleDonneesID = 'IdDeLaFeuilleDeDonnees'                 //feuille de données noms et emails ID
var feuilleDonneesName = 'RH Technique'            //feuille de données noms et emails Name
const sheetRHTechnique = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(feuilleDonneesName);
var plageName = [];     

// Récupération de l'adresse de l'expéditeur
try {
aliases = GmailApp.getAliases();                  // récupère l'email d'alias pour l'envoie
} catch{
// on ne fait rien
}

function pause(milliseconds) {                      // Attendre jusqu'à ce que le délai soit écoulé
  var start = new Date().getTime();
  while (new Date().getTime() - start < milliseconds) {
  }
}

function onOpen() {                               // création des menus
  ui = SpreadsheetApp.getUi();                    // Or DocumentApp or FormApp.
  ui.createMenu('Send as PDF')
    .addItem('Envoyer Convoc', 'SendFdrAsPDF')
    .addItem('Envoyer Convoc à tous', 'SendALL')
    .addToUi();
}

function GetEmail(){

  email = sheetConvocation.getRange('C31').getValue() ; // récupère l'email type

}

// Création de l'email
function SendFdrAsPDF() {

  n=0                                             // compteur

  // variables utilisateur

  GetEmail()

  var choix = "no"                                 // de base, on ne souhaite pas envoyer à l'adresse type

  // Récupération de l'email
  if (email!=""){                                 // regarde si j'ai une adresse type
    choix = Browser.msgBox("Envoyer la convocation à l'adresse email : ", email , Browser.Buttons.YES_NO_CANCEL);
  }

  switch(choix){
    case "yes" : //l'email sera envoyé à l'adresse type
      break
    case "no" : 
      email= Browser.inputBox('Send Convoc as PDF', "Entrez l'adresse email du destinataire :", Browser.Buttons.OK_CANCEL);
      break
    case "cancel" : 
    return
  }

  PDFcreat()

}

function PDFcreat() {
  // Préparation de l'email
  objet = feuille + " " +sheetConvocation.getRange('D3').getValue() +" @ Biches festival" ;
  fichier = objet +".pdf";
  dossier = DriveApp.getFolderById('IDduDossierDuPdf'); 
  corps = "Bonjour, <br><br>Merci de trouver en pièce jointe le document cité en objet et de me <u>confirmer dès que possible la bonne réception du document.</u><br><br><b>Martin Mignon</b><br>www.martin-technique.fr";

  PDF();
  Send(email);
}

function SendALL(){

  n=0
  var sentEmails = {}
                                    
  plageName = sheetRHTechnique.getRange('B11:B150').getValues();        //variablequi récupère les valeurs des cellules noms

  //supprime les noms en doublons du array plageName
  function RemoveDuplicates(plageName) {                          
    var unique = {};                            // crée un objet vide "unique" pour stocker les éléments uniques du tableau plageName.
    plageName.forEach(function(i) {             // itère sur chaque élément i du tableau plageName. 
      if(!unique[i]) {                          // vérifie si i n'existe pas déjà comme clé dans l'objet unique.                 
      unique[i] = true;                         // Si non, cela signifie que i est un élément unique et il est ajouté à l'objet unique en utilisant unique[i] = true;.
      }
    });
    return Object.keys(unique);
  }
  uniqueName = RemoveDuplicates(plageName)
  console.log(uniqueName);

  var checked = sheetRHTechnique.getRange('D11:D150').getValues();      // récupère les valeurs des cellules checkbox
  var cellName = sheetConvocation.getRange('D3');                           
 
  plageName.forEach(function(name, index) {
    var isChecked = checked [index][0];                       // Accède à la valeur spécifique de la case correspondante checkbox

     console.log(name, index, isChecked);

    if (isChecked) {                                          // Si checké
      cellName.setValue(name);                               // cellName prend la valeur du nom

      GetEmail();                                        // récupère l'adresse mail du nom actuel

      if(!sentEmails[name]) {                          // vérifie si i n'existe pas déjà comme clé dans l'objet unique. 
        PDFcreat();                                     // il utilise PDFcreat() pour faire l'envoie                
        sentEmails[name] = true;    
      }
    } else {
      return
    }
  });
 
}

// Création du PDF
function PDF(){

  // Création du fichier pdf
  const url = docID + 'export?';
  const exportOptions =
    'gid=IDduDocumentAEnvoyer'+ 
    '&exportFormat=pdf&format=pdf' + 
    '&size=A4' + 
    '&portrait=true' +                     // orientation portrait, false pour paysage
    '&scale=2' +
    '&sheetnames=false&printtitle=false' + // pas de nom ni de titre à l'impression
    '&pagenumbers=false&gridlines=false' ; // pas de numérotation, pas de grille                       

  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  reponse = UrlFetchApp.fetch(url + exportOptions, params).getBlob();

  pause(10000);                           // pour éviter la saturation du serveur

}

// Envoi email avec fichier attaché
function Send(email) {
  try {
    var aliases = GmailApp.getAliases();
    var senderName = "Martin-technique.fr";
    var options = {
      htmlBody: corps,
      attachments: [{
        fileName: fichier,
        content: reponse.getBytes(),
        mimeType: "application/pdf"
      }]
    };
    
    if (aliases.length > 0) {
      options.from = aliases[0];
      options.name = senderName;
    }

   GmailApp.sendEmail(email, objet, corps, options);

    n = n + 1;
  } catch (error) {
    Browser.msgBox("Attention", "L'email n'est pas parti, désolé", Browser.Buttons.CANCEL);
  }
}

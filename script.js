// Vérification que les bibliothèques sont chargées
if (typeof grist === 'undefined') {
  console.error('Grist API n\'est pas chargée');  // [web:59][web:48]
}
if (typeof JSZip === 'undefined') {
  console.error('JSZip n\'est pas chargée');  // [web:7][web:19]

// Configuration du widget Grist
grist.ready({
  requiredAccess: 'full',
  columns: [
    {
      name: 'AttachmentColumns',
      title: 'Colonnes de pièces jointes',
      type: 'Attachments',
      optional: false,
      allowMultiple: true,
      description: 'Sélectionnez toutes les colonnes contenant des pièces jointes'  // [web:59][web:60]
    },
    {
      name: 'ZipName',
      title: 'Nom du fichier ZIP',
      type: 'Text',
      optional: false,
      description: 'Colonne contenant le nom pour le fichier ZIP (ex: Identité)'  // [web:59][web:60]
    }
  ]
});

// Références aux éléments DOM
const btn = document.getElementById('downloadBtn');
const msg = document.getElementById('msg');
const icon = btn.querySelector('.icon');
const spinner = btn.querySelector('.spinner');
const text = btn.querySelector('.text');
let currentRecord = null;
let currentMappings = null;  // on stocke les mappings de colonnes  [web:59][web:48]

/**
 * Fonction principale de téléchargement des pièces jointes en ZIP
 */
async function downloadAllAttachments() {
  if (!currentRecord) {
    msg.textContent = '⚠️ Aucun enregistrement sélectionné';
    return;  // [web:59][web:48]
  }
  
  // Activer l'état de chargement
  btn.classList.add('loading');
  icon.style.display = 'none';
  spinner.style.display = 'block';
  text.textContent = 'Création du ZIP...';  // [web:28][web:83]
  
  // Récupérer les colonnes mappées
  const mapped = grist.mapColumnNames(currentRecord);  // [web:59][web:48]
  
  // Vérifier que toutes les colonnes sont mappées
  if (!mapped || !mapped.AttachmentColumns || !mapped.ZipName) {
    resetButton();
    msg.textContent = '⚠️ Veuillez mapper toutes les colonnes';
    return;  // [web:59][web:48]
  }
  
  const allAttachments = mapped.AttachmentColumns;
  const identity = String(mapped.ZipName || 'sans_nom').trim();
  let totalCount = 0;  // [web:59][web:48]
  
  // Compter le total de fichiers
  for (const attachmentList of allAttachments) {
    if (Array.isArray(attachmentList)) {
      totalCount += attachmentList.length;
    }
  }
  
  // Vérifier qu'il y a des fichiers à télécharger
  if (totalCount === 0) {
    resetButton();
    msg.textContent = '⚠️ Aucune pièce jointe à télécharger';
    return;  // [web:59][web:48]
  }
  
  try {
    // Obtenir le token d'accès Grist
    const { token, baseUrl } = await grist.docApi.getAccessToken({ readOnly: true });  // [web:43][web:48]
    const zip = new JSZip();
    let processedCount = 0;  // [web:19][web:28]
    
    // Récupérer les vrais noms de colonnes d’attachements
    let realAttachmentCols = [];
    if (currentMappings && currentMappings.AttachmentColumns) {
      // AttachmentColumns est un tableau de noms de colonnes réelles  [web:59][web:48]
      realAttachmentCols = currentMappings.AttachmentColumns;
    }

    // Parcourir toutes les colonnes et fichiers
    for (let colIndex = 0; colIndex < allAttachments.length; colIndex++) {
      const attachmentList = allAttachments[colIndex];
      const colName = realAttachmentCols[colIndex] || `Col${colIndex + 1}`;  // [web:59][web:48]
      
      if (Array.isArray(attachmentList)) {
        for (let fileIndex = 0; fileIndex < attachmentList.length; fileIndex++) {
          const attId = attachmentList[fileIndex];
          const url = `${baseUrl}/attachments/${attId}/download?auth=${token}`;  // [web:43][web:48]
          
          // Mettre à jour le message de progression
          text.textContent = `Ajout ${processedCount + 1}/${totalCount}...`;  // [web:59][web:65]
          
          try {
            // Récupérer le fichier comme blob
            const response = await fetch(url);
            
            if (!response.ok) {
              console.error(`Erreur lors du téléchargement du fichier ${attId}: ${response.status}`);
              continue;  // [web:25][web:87]
            }
            
            const blob = await response.blob();  // [web:25][web:84]

            // Nouveau schéma de nommage en .pdf :
            // <NomColonne> - <Identité> - <index>.pdf
            const safeColName = colName.replace(/[^a-z0-9_\-\s]/gi, '_');
            const safeIdentity = identity.replace(/[^a-z0-9_\-\s]/gi, '_');
            const filename = `${safeColName} - ${safeIdentity} - ${fileIndex + 1}.pdf`;  // [web:19][web:28]

            // Ajouter le fichier au ZIP
            zip.file(filename, blob);  // [web:19][web:7][web:28]
            processedCount++;
            
          } catch (fetchError) {
            console.error(`Erreur lors du téléchargement du fichier ${attId}:`, fetchError);
            continue;  // [web:25][web:87]
          }
        }
      }
    }
    
    // Vérifier qu'au moins un fichier a été traité
    if (processedCount === 0) {
      resetButton();
      msg.textContent = '❌ Aucun fichier n\'a pu être téléchargé';
      return;  // [web:25][web:87]
    }
    
    // Générer le ZIP
    text.textContent = 'Génération du ZIP...';
    
    const zipBlob = await zip.generateAsync({ 
      type: 'blob',
      streamFiles: true,
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });  // [web:19][web:7][web:28]
    
    // Nettoyer le nom du fichier ZIP (supprimer les caractères spéciaux)
    const cleanZipName = identity.replace(/[^a-z0-9_\-\s]/gi, '_') || 'attachments';  // [web:31][web:33]
    
    // Créer le lien de téléchargement
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = `${cleanZipName}.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // Libérer la mémoire
    URL.revokeObjectURL(link.href);  // [web:28][web:83]
    
    // Message de succès
    msg.textContent = `✅ ${processedCount} fichier(s) téléchargé(s) dans ${cleanZipName}.zip`;
    
  } catch (error) {
    msg.textContent = `❌ Erreur lors de la création du ZIP`;
    console.error('Erreur complète:', error);  // [web:28][web:83]
  }
  
  // Réinitialiser le bouton
  resetButton();
}

/**
 * Réinitialiser l'état du bouton
 */
function resetButton() {
  btn.classList.remove('loading');
  icon.style.display = 'block';
  spinner.style.display = 'none';
  text.textContent = 'Télécharger en ZIP';  // [web:28][web:83]
}

// Ajouter l'écouteur d'événement au bouton
btn.addEventListener('click', downloadAllAttachments);  // [web:59][web:48]

/**
 * Écouter les changements d'enregistrement dans Grist
 * On récupère aussi `mappings` pour connaître les vrais noms de colonnes.
 */
grist.onRecord((record, mappings) => {
  currentRecord = record;
  currentMappings = mappings || currentMappings;  // [web:59][web:48]

  const mapped = grist.mapColumnNames(record);  // [web:59][web:48]
  
  if (mapped && mapped.AttachmentColumns) {
    // Compter le nombre total de fichiers
    let totalCount = 0;
    for (const attachmentList of mapped.AttachmentColumns) {
      if (Array.isArray(attachmentList)) {
        totalCount += attachmentList.length;
      }
    }
    
    // Convertir le nom en string
    const zipName = String(mapped.ZipName || 'sans nom');
    
    // Afficher le message d'information
    msg.textContent = `📎 ${totalCount} fichier(s) → ${zipName}.zip`;
  } else {
    msg.textContent = '⚙️ Configurez les colonnes dans les paramètres du widget';
  }
});  // [web:59][web:48]

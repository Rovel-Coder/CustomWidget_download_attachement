// Vérification que les bibliothèques sont chargées
if (typeof grist === 'undefined') {
  console.error('Grist API n\'est pas chargée');
}
if (typeof JSZip === 'undefined') {
  console.error('JSZip n\'est pas chargée');
}

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
      description: 'Sélectionnez toutes les colonnes contenant des pièces jointes'
    },
    {
      name: 'ZipName',
      title: 'Nom du fichier ZIP',
      type: 'Text',
      optional: false,
      description: 'Colonne contenant le nom pour le fichier ZIP (ex: Identité)'
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
let currentMappings = null;   // configuration de mapping reçue via onRecord

/**
 * Fonction principale de téléchargement des pièces jointes en ZIP
 */
async function downloadAllAttachments() {
  if (!currentRecord) {
    msg.textContent = '⚠️ Aucun enregistrement sélectionné';
    return;
  }
  
  // Activer l'état de chargement
  btn.classList.add('loading');
  icon.style.display = 'none';
  spinner.style.display = 'block';
  text.textContent = 'Création du ZIP...';
  
  // Récupérer les colonnes mappées
  const mapped = grist.mapColumnNames(currentRecord);
  
  // Vérifier que toutes les colonnes sont mappées
  if (!mapped || !mapped.AttachmentColumns || !mapped.ZipName) {
    resetButton();
    msg.textContent = '⚠️ Veuillez mapper toutes les colonnes';
    return;
  }
  
  const allAttachments = mapped.AttachmentColumns;
  const identity = String(mapped.ZipName || 'sans_nom').trim();
  let totalCount = 0;
  
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
    return;
  }
  
  try {
    // Obtenir le token d'accès Grist
    const { token, baseUrl } = await grist.docApi.getAccessToken({ readOnly: true });
    const zip = new JSZip();
    let processedCount = 0;

    // Récupérer les vrais noms de colonnes d’attachements depuis mappings.AttachmentColumns
    let realAttachmentCols = [];
    if (currentMappings && currentMappings.AttachmentColumns) {
      realAttachmentCols = currentMappings.AttachmentColumns;
    }

    // Parcourir toutes les colonnes et fichiers
    for (let colIndex = 0; colIndex < allAttachments.length; colIndex++) {
      const attachmentList = allAttachments[colIndex];
      const colName = realAttachmentCols[colIndex] || `Col${colIndex + 1}`;
      
      if (Array.isArray(attachmentList)) {
        const hasMultipleInCell = attachmentList.length > 1;

        for (let fileIndex = 0; fileIndex < attachmentList.length; fileIndex++) {
          const attId = attachmentList[fileIndex];
          const url = `${baseUrl}/attachments/${attId}/download?auth=${token}`;
          
          // Mettre à jour le message de progression
          text.textContent = `Ajout ${processedCount + 1}/${totalCount}...`;
          
          try {
            // Récupérer le fichier comme blob
            const response = await fetch(url);
            
            if (!response.ok) {
              console.error(`Erreur lors du téléchargement du fichier ${attId}: ${response.status}`);
              continue;
            }
            
            const blob = await response.blob();

            // Nom du fichier dans le ZIP
            const safeColName = colName.replace(/[^a-z0-9\-\\s]/gi, '_');
            const safeIdentity = identity.replace(/[^a-z0-9\-\\s]/gi, '_');

            let filename;
            if (hasMultipleInCell) {
              // Plusieurs fichiers pour cette cellule -> on garde l’index
              filename = `${safeColName}_${safeIdentity}_${fileIndex + 1}.pdf`;
            } else {
              // Un seul fichier pour cette cellule -> pas d’index ni de "_" final
              filename = `${safeColName}_${safeIdentity}.pdf`;
            }

            // Ajouter le fichier au ZIP
            zip.file(filename, blob);
            processedCount++;
            
          } catch (fetchError) {
            console.error(`Erreur lors du téléchargement du fichier ${attId}:`, fetchError);
            continue;
          }
        }
      }
    }
    
    // Vérifier qu'au moins un fichier a été traité
    if (processedCount === 0) {
      resetButton();
      msg.textContent = '❌ Aucun fichier n\'a pu être téléchargé';
      return;
    }
    
    // Générer le ZIP
    text.textContent = 'Génération du ZIP...';
    
    const zipBlob = await zip.generateAsync({ 
      type: 'blob',
      streamFiles: true,
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    // Nettoyer le nom du fichier ZIP (supprimer les caractères spéciaux)
    const cleanZipName = identity.replace(/[^a-z0-9_\-\s]/gi, '_') || 'attachments';
    
    // Créer le lien de téléchargement
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = `${cleanZipName}.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // Libérer la mémoire
    URL.revokeObjectURL(link.href);
    
    // Message de succès
    msg.textContent = `✅ ${processedCount} fichier(s) téléchargé(s) dans ${cleanZipName}.zip`;
    
  } catch (error) {
    msg.textContent = `❌ Erreur lors de la création du ZIP`;
    console.error('Erreur complète:', error);
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
  text.textContent = 'Télécharger en ZIP';
}

// Ajouter l'écouteur d'événement au bouton
btn.addEventListener('click', downloadAllAttachments);

/**
 * Écouter les changements d'enregistrement dans Grist
 * Le 2ᵉ paramètre `mappings` contient la configuration de mapping.
 */
grist.onRecord((record, mappings) => {
  currentRecord = record;
  currentMappings = mappings || currentMappings;

  const mapped = grist.mapColumnNames(record);
  
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
});

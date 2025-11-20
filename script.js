// Vérification que les bibliothèques sont chargées
if (typeof grist === 'undefined') {
  console.error('Grist API n\'est pas chargée');  // [web:11][web:13]
}
if (typeof JSZip === 'undefined') {
  console.error('JSZip n\'est pas chargée');  // [web:19][web:22]
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
      description: 'Sélectionnez toutes les colonnes contenant des pièces jointes'  // [web:11][web:13]
    },
    {
      name: 'ZipName',
      title: 'Nom du fichier ZIP',
      type: 'Text',
      optional: false,
      description: 'Colonne contenant le nom pour le fichier ZIP (ex: Identité)'  // [web:11][web:13]
    }
  ]
});

// Références aux éléments DOM
const btn = document.getElementById('downloadBtn');
const msg = document.getElementById('msg');
const icon = btn.querySelector('.icon');
const spinner = btn.querySelector('.spinner');
const text = btn.querySelector('.text');
let currentRecord = null;  // [web:11][web:13]

/**
 * Fonction principale de téléchargement des pièces jointes en ZIP
 */
async function downloadAllAttachments() {
  if (!currentRecord) {
    msg.textContent = '⚠️ Aucun enregistrement sélectionné';
    return;  // [web:11][web:13]
  }
  
  // Activer l'état de chargement
  btn.classList.add('loading');
  icon.style.display = 'none';
  spinner.style.display = 'block';
  text.textContent = 'Création du ZIP...';  // [web:22][web:28]
  
  // Récupérer les colonnes mappées
  const mapped = grist.mapColumnNames(currentRecord);  // [web:11][web:13]
  
  // Vérifier que toutes les colonnes sont mappées
  if (!mapped || !mapped.AttachmentColumns || !mapped.ZipName) {
    resetButton();
    msg.textContent = '⚠️ Veuillez mapper toutes les colonnes';
    return;  // [web:11][web:13]
  }
  
  const allAttachments = mapped.AttachmentColumns;
  const zipName = String(mapped.ZipName || 'attachments').trim();
  let totalCount = 0;  // [web:11][web:13]
  
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
    return;  // [web:11][web:13]
  }
  
  try {
    // Obtenir le token d'accès Grist
    const { token, baseUrl } = await grist.docApi.getAccessToken({ readOnly: true });
    const zip = new JSZip();
    let processedCount = 0;  // [web:11][web:13][web:19]
    
    // Parcourir toutes les colonnes et fichiers
    for (let colIndex = 0; colIndex < allAttachments.length; colIndex++) {
      const attachmentList = allAttachments[colIndex];
      
      if (Array.isArray(attachmentList)) {
        for (let fileIndex = 0; fileIndex < attachmentList.length; fileIndex++) {
          const attId = attachmentList[fileIndex];
          const url = `${baseUrl}/attachments/${attId}/download?auth=${token}`;
          
          // Mettre à jour le message de progression
          text.textContent = `Ajout ${processedCount + 1}/${totalCount}...`;  // [web:21]
          
          try {
            // Récupérer le fichier comme blob
            const response = await fetch(url);
            
            if (!response.ok) {
              console.error(`Erreur lors du téléchargement du fichier ${attId}: ${response.status}`);
              continue;  // [web:21][web:25]
            }
            
            const blob = await response.blob();  // [web:25]
            
            // Extraire le nom du fichier depuis les headers
            const contentDisposition = response.headers.get('content-disposition');
            let filename = `fichier_${colIndex + 1}_${fileIndex + 1}`;  // [web:14][web:20]
            
            if (contentDisposition) {
              try {
                let candidate = null;

                // 1) Essayer filename* (UTF-8'')
                const fnStarMatch = contentDisposition.match(/filename\*\s*=\s*([^;]+)/i);
                if (fnStarMatch && fnStarMatch[1]) {
                  candidate = fnStarMatch[1].trim();
                  candidate = candidate.replace(/^utf-8''/i, '');
                  candidate = candidate.replace(/['"]/g, '');
                  candidate = decodeURIComponent(candidate);
                } else {
                  // 2) Repli sur filename classique
                  const fnMatch = contentDisposition.match(/filename[^;=\n]*=\s*([^;\n]*)/i);
                  if (fnMatch && fnMatch[1]) {
                    candidate = fnMatch[1].trim().replace(/['"]/g, '');
                    try {
                      candidate = decodeURIComponent(candidate);
                    } catch (e) {
                      // on garde tel quel
                    }
                  }
                }

                if (candidate) {
                  filename = candidate;
                }
              } catch (e) {
                console.warn('Impossible de lire le nom de fichier depuis Content-Disposition:', e);
              }
            }  // [web:14][web:20][web:31]
            
            // Ajouter le fichier au ZIP
            zip.file(filename, blob);  // [web:19][web:22][web:28]
            processedCount++;
            
          } catch (fetchError) {
            console.error(`Erreur lors du téléchargement du fichier ${attId}:`, fetchError);
            continue;  // [web:25][web:38]
          }
        }
      }
    }
    
    // Vérifier qu'au moins un fichier a été traité
    if (processedCount === 0) {
      resetButton();
      msg.textContent = '❌ Aucun fichier n\'a pu être téléchargé';
      return;  // [web:25]
    }
    
    // Générer le ZIP
    text.textContent = 'Génération du ZIP...';
    
    const zipBlob = await zip.generateAsync({ 
      type: 'blob',
      streamFiles: true,
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });  // [web:19][web:22][web:28]
    
    // Nettoyer le nom du fichier ZIP (supprimer les caractères spéciaux)
    const cleanZipName = zipName.replace(/[^a-z0-9_\-\s]/gi, '_');  // [web:31][web:33]
    
    // Créer le lien de téléchargement
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = `${cleanZipName}.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // Libérer la mémoire
    URL.revokeObjectURL(link.href);  // [web:22][web:28]
    
    // Message de succès
    msg.textContent = `✅ ${processedCount} fichier(s) téléchargé(s) dans ${cleanZipName}.zip`;
    
  } catch (error) {
    msg.textContent = `❌ Erreur lors de la création du ZIP`;
    console.error('Erreur complète:', error);  // [web:22][web:28]
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
  text.textContent = 'Télécharger en ZIP';  // [web:22][web:28]
}

// Ajouter l'écouteur d'événement au bouton
btn.addEventListener('click', downloadAllAttachments);  // [web:11][web:13]

/**
 * Écouter les changements d'enregistrement dans Grist
 */
grist.onRecord(record => {
  currentRecord = record;
  const mapped = grist.mapColumnNames(record);  // [web:11][web:13]
  
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
});  // [web:11][web:13]

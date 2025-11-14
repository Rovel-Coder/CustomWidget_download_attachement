// Attendre que les bibliothèques soient chargées
if (typeof grist === 'undefined' || typeof JSZip === 'undefined') {
  console.error('Les bibliothèques requises ne sont pas chargées');
}

// Configuration du widget
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

const btn = document.getElementById('downloadBtn');
const msg = document.getElementById('msg');
const icon = btn.querySelector('.icon');
const spinner = btn.querySelector('.spinner');
const text = btn.querySelector('.text');
let currentRecord = null;

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
  
  const mapped = grist.mapColumnNames(currentRecord);
  
  if (!mapped || !mapped.AttachmentColumns || !mapped.ZipName) {
    resetButton();
    msg.textContent = '⚠️ Veuillez mapper toutes les colonnes';
    return;
  }
  
  const allAttachments = mapped.AttachmentColumns;
  const zipName = mapped.ZipName || 'attachments';
  let totalCount = 0;
  
  // Compter le total de fichiers
  for (const attachmentList of allAttachments) {
    if (Array.isArray(attachmentList)) {
      totalCount += attachmentList.length;
    }
  }
  
  if (totalCount === 0) {
    resetButton();
    msg.textContent = '⚠️ Aucune pièce jointe à télécharger';
    return;
  }
  
  try {
    const { token, baseUrl } = await grist.docApi.getAccessToken({ readOnly: true });
    const zip = new JSZip();
    let processedCount = 0;
    
    // Parcourir toutes les colonnes et fichiers
    for (let colIndex = 0; colIndex < allAttachments.length; colIndex++) {
      const attachmentList = allAttachments[colIndex];
      
      if (Array.isArray(attachmentList)) {
        for (let fileIndex = 0; fileIndex < attachmentList.length; fileIndex++) {
          const attId = attachmentList[fileIndex];
          const url = `${baseUrl}/attachments/${attId}/download?auth=${token}`;
          
          text.textContent = `Ajout ${processedCount + 1}/${totalCount}...`;
          
          // Récupérer le fichier comme blob
          const response = await fetch(url);
          
          if (!response.ok) {
            console.error(`Erreur lors du téléchargement du fichier ${attId}`);
            continue;
          }
          
          const blob = await response.blob();
          
          // Extraire le nom du fichier depuis les headers
          const contentDisposition = response.headers.get('content-disposition');
          let filename = `fichier_${colIndex + 1}_${fileIndex + 1}`;
          
          if (contentDisposition) {
            const filenameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
            if (filenameMatch && filenameMatch[1]) {
              filename = filenameMatch[1].replace(/['"]/g, '');
              // Décoder les caractères encodés
              filename = decodeURIComponent(filename);
            }
          }
          
          // Ajouter le fichier au ZIP
          zip.file(filename, blob);
          processedCount++;
        }
      }
    }
    
    if (processedCount === 0) {
      resetButton();
      msg.textContent = '❌ Aucun fichier n\'a pu être téléchargé';
      return;
    }
    
    text.textContent = 'Génération du ZIP...';
    
    // Générer le ZIP
    const zipBlob = await zip.generateAsync({ 
      type: 'blob',
      streamFiles: true,
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    // Télécharger le ZIP
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    // Nettoyer le nom du fichier
    const cleanZipName = zipName.replace(/[^a-z0-9_\-]/gi, '_');
    link.download = `${cleanZipName}.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
    
    msg.textContent = `✅ ${processedCount} fichier(s) téléchargé(s) dans ${cleanZipName}.zip`;
  } catch (error) {
    msg.textContent = `❌ Erreur lors de la création du ZIP`;
    console.error('Erreur:', error);
  }
  
  resetButton();
}

function resetButton() {
  btn.classList.remove('loading');
  icon.style.display = 'block';
  spinner.style.display = 'none';
  text.textContent = 'Télécharger en ZIP';
}

btn.addEventListener('click', downloadAllAttachments);

grist.onRecord(record => {
  currentRecord = record;
  const mapped = grist.mapColumnNames(record);
  
  if (mapped && mapped.AttachmentColumns) {
    let totalCount = 0;
    for (const attachmentList of mapped.AttachmentColumns) {
      if (Array.isArray(attachmentList)) {
        totalCount += attachmentList.length;
      }
    }
    const zipName = mapped.ZipName || 'sans nom';
    msg.textContent = `📎 ${totalCount} fichier(s) → ${zipName}.zip`;
  } else {
    msg.textContent = '⚙️ Configurez les colonnes dans les paramètres du widget';
  }
});

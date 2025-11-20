// Vérification que les bibliothèques sont chargées
if (typeof grist === 'undefined') {
  console.error('Grist API n\'est pas chargée');
}
if (typeof JSZip === 'undefined') {
  console.error('JSZip n\'est pas chargée');
}
if (typeof window.jspdf === 'undefined') {
  console.error('jsPDF n\'est pas chargé');
}

// Fonction utilitaire pour convertir PNG/JPG -> PDF
async function imageBlobToPdf(blob, mimeType = "image/jpeg") {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      const imgData = e.target.result;
      const pdf = new window.jspdf.jsPDF();
      // Dimensions ajustables selon tes besoins
      pdf.addImage(imgData, mimeType === "image/png" ? "PNG" : "JPEG", 10, 10, 180, 240);
      // output('blob') disponible en jsPDF 2.x
      const pdfBlob = pdf.output('blob');
      resolve(pdfBlob);
    };
    reader.readAsDataURL(blob);
  });
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
let currentMappings = null;

const cleanPart = (str) => {
  return String(str || '')
    .replace(/[^a-zA-Z0-9]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .trim();
};

/**
 * Fonction principale de téléchargement des pièces jointes en ZIP
 */
async function downloadAllAttachments() {
  if (!currentRecord) {
    msg.textContent = '⚠️ Aucun enregistrement sélectionné';
    return;
  }
  
  btn.classList.add('loading');
  icon.style.display = 'none';
  spinner.style.display = 'block';
  text.textContent = 'Création du ZIP...';

  const mapped = grist.mapColumnNames(currentRecord);

  const rawIdentity = String(mapped.ZipName || 'None').trim();

  const allAttachments = mapped.AttachmentColumns;
  let totalCount = 0;
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

    let realAttachmentCols = [];
    if (currentMappings && currentMappings.AttachmentColumns) {
      realAttachmentCols = currentMappings.AttachmentColumns;
    }

    const identity = cleanPart(rawIdentity);

    for (let colIndex = 0; colIndex < allAttachments.length; colIndex++) {
      const attachmentList = allAttachments[colIndex];
      const rawColName = realAttachmentCols[colIndex] || `Col${colIndex + 1}`;
      const colName = cleanPart(rawColName);

      if (Array.isArray(attachmentList) && attachmentList.length > 0) {
        const hasMultipleInCell = attachmentList.length > 1;
        for (let fileIndex = 0; fileIndex < attachmentList.length; fileIndex++) {
          const attId = attachmentList[fileIndex];
          const url = `${baseUrl}/attachments/${attId}/download?auth=${token}`;
          text.textContent = `Ajout ${processedCount + 1}/${totalCount}...`;

          try {
            const response = await fetch(url);
            if (!response.ok) {
              console.error(`Erreur lors du téléchargement du fichier ${attId}: ${response.status}`);
              continue;
            }
            const blob = await response.blob();

            // Détecter le type/extension original
            let extension = 'pdf'; // par défaut, PDF
            let finalBlob = blob;

            const contentDisposition = response.headers.get('content-disposition');
            if (contentDisposition) {
              const match = contentDisposition.match(/filename[^;=\n]*=\s*(['"])?([^'";\n]+)\1?/);
              if (match && match[2]) {
                const fname = match[2];
                const extMatch = fname.match(/\.[a-z0-9]+$/i);
                if (extMatch) {
                  extension = extMatch[0].replace('.', '').toLowerCase();
                }
              }
            }

            // Si PNG ou JPG : convertir en PDF
            if (blob.type === "image/png" || blob.type === "image/jpeg" || extension === "png" || extension === "jpg" || extension === "jpeg") {
              finalBlob = await imageBlobToPdf(blob, blob.type);
              extension = "pdf";
            }

            // Nom du fichier
            let filename;
            if (hasMultipleInCell) {
              filename = `${colName}_${identity}_${fileIndex + 1}.${extension}`;
            } else {
              filename = `${colName}_${identity}.${extension}`;
            }

            zip.file(filename, finalBlob);
            processedCount++;
          } catch (fetchError) {
            console.error(`Erreur lors du téléchargement du fichier ${attId}:`, fetchError);
            continue;
          }
        }
      }
    }

    if (processedCount === 0) {
      resetButton();
      msg.textContent = '❌ Aucun fichier n\'a pu être téléchargé';
      return;
    }

    text.textContent = 'Génération du ZIP...';

    const cleanZipName = identity || 'attachments';
    const zipBlob = await zip.generateAsync({ 
      type: 'blob',
      streamFiles: true,
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = `${cleanZipName}.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);

    msg.textContent = `✅ ${processedCount} fichier(s) téléchargé(s) dans ${cleanZipName}.zip`;

  } catch (error) {
    msg.textContent = `❌ Erreur lors de la création du ZIP`;
    console.error('Erreur complète:', error);
  }

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

btn.addEventListener('click', downloadAllAttachments);

grist.onRecord((record, mappings) => {
  currentRecord = record;
  currentMappings = mappings || currentMappings;

  const mapped = grist.mapColumnNames(record);
  if (mapped && mapped.AttachmentColumns) {
    let totalCount = 0;
    for (const attachmentList of mapped.AttachmentColumns) {
      if (Array.isArray(attachmentList)) {
        totalCount += attachmentList.length;
      }
    }
    const zipName = String(mapped.ZipName || 'None');
    msg.textContent = `📎 ${totalCount} fichier(s) → ${zipName}.zip`;
  } else {
    msg.textContent = '⚙️ Configurez les colonnes dans les paramètres du widget';
  }
});

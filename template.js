/**
 * Ejemplo de creación de presentación en Google Slides mediante Node.js en base a una presentación Template 
 * Usamos el paquete googleapis para autenticarnos y enviar un batchUpdate con las instrucciones,
 * y se elimina el slide default automáticamente.
 *
 * Requisitos:
 * - Habilitar la API de Google Slides y Google Drive en tu proyecto de Google Cloud.
 * - Configurar la cuenta de servicio con delegación de autoridad para suplantar a mok@plunkton.com.
 *
 * Para instalar la dependencia:
 *   npm install googleapis
 */

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

// Define the log file path with a timestamp
const logFilePath = path.join(__dirname, `log-${Date.now()}.log`);

// Function to write logs to the file
function writeLog(message) {
  fs.appendFileSync(logFilePath, message + '\n');
}

// Autenticación con cuenta de servicio con delegación para suplantar a mok@plunkton.com
const auth = new google.auth.JWT({
  keyFile: process.env.KEY_FILE,
  scopes: [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive'
  ],
  subject: process.env.SUBJECT_EMAIL  // Impersona a mok@plunkton.com
});

const slides = google.slides({ version: 'v1', auth });
const drive = google.drive({ version: 'v3', auth });

// Load external variables from JSON file
const jsonFilePath = process.env.JSON_FILE_PATH;
const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, 'utf8'));

/**
 * JSON con las instrucciones para crear la presentación:
 * Se generan 4 slides con diferentes elementos, dejando el slide default que se crea automáticamente.
 */
const updates = [
  {
    deleteText: { //Borra el texto por defecto del slide default
      objectId: 'p2_i7171',
      textRange: {
        type: 'ALL'
      }
    }
  },
  {     
    insertText: { 
      objectId: 'p2_i7171',
      text: jsonData.teams[0].name + ' vs ' + jsonData.teams[1].name,
      insertionIndex: 0 // Inserts text at the beginning
    }
  },
  {
    deleteText: { //Borra el texto por defecto del slide default
      objectId: 'p2_i7172',
      textRange: {
        type: 'ALL'
      }
    }
  },
  {
    insertText: { 
      objectId: 'p2_i7172',
      text: jsonData.championship + ' - ' + jsonData.gameDate,
      insertionIndex: 0 // Inserts text at the beginning
    }
  },
  { //Esto para mostrar como se puede crear un slide en blanco
    createSlide: {
      objectId: 'T1-Snap0',
      slideLayoutReference: { predefinedLayout: 'BLANK' }
    }
  },
  {
    replaceImage: { //Logo del Equipo 1
      imageObjectId: 'g359de4ee1dd_1_1',
      imageReplaceMethod: 'CENTER_INSIDE',
      url: jsonData.teams[0].logo
    }
  },
  {
    replaceImage: { //Logo del Equipo 2
      imageObjectId: 'g359de4ee1dd_1_0',
      imageReplaceMethod: 'CENTER_INSIDE',
      url: jsonData.teams[1].logo
    }
  },
  {
    duplicateObject: { // Agregar instrucción para duplicar el slide p7
      objectId: 'p7', // ID del slide que quieres duplicar
      objectIds: {
        'p7': 'p7_copy1' // Nuevo ID válido para el slide duplicado. 
      } // El slide nuevo tendrá los mismos elementos que el original, pero los object ID serán desconocidos. 
    }
  }
];

/**
 * Función principal que crea la presentación, aplica las actualizaciones y elimina el slide default.
 * La presentación se crea con 720 puntos por 405 puntos
 */
async function createPresentation() {
  try {
    
    // Copy the template presentation.
    const copyResponse = await drive.files.copy({
      fileId: process.env.TEMPLATE_ID,  // Cambia el nombre de la presentación copiada
      requestBody: {
        name: 'Lions'
      }
    });
    
    const presentationId = copyResponse.data.id;
    console.log('ID de la presentación: ' + presentationId);

    //Envía las instrucciones en batchUpdate para configurar la presentación.
    await slides.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: updates } 
    });
    
    const presentationResponse = await slides.presentations.get({ presentationId });
    
    const slidesArray = presentationResponse.data.slides;


    slidesArray.forEach(slide => {
      writeLog('Slide ID: ' + slide.objectId);

      if (slide.pageElements) {
        slide.pageElements.forEach(element => {
          writeLog('  Element ID: ' + element.objectId);
          // Print some general properties of the element in a readable JSON format:
          writeLog('  Element properties: ' + JSON.stringify(element, null, 2));

          // If the element is a table, it will have a "table" property.
          if (element.table) {
            writeLog('    This element is a table:');
            writeLog('      Rows: ' + element.table.rows);
            writeLog('      Columns: ' + element.table.columns);

            // Iterate over each row and cell if available.
            if (element.table.tableRows) {
              element.table.tableRows.forEach((row, rowIndex) => {
                row.tableCells.forEach((cell, colIndex) => {
                  // If the cell contains text, it usually comes under cell.text.textElements.
                  // This example concatenates all text run content if available.
                  let cellText = '';
                  if (cell.text && cell.text.textElements) {
                    cell.text.textElements.forEach(te => {
                      if (te.textRun && te.textRun.content) {
                        cellText += te.textRun.content;
                      }
                    });
                  }
                  writeLog(`      Cell [${rowIndex}, ${colIndex}] text: "${cellText.trim()}"`);
                });
              });
            }
          }
        });
      } else {
        writeLog('  This slide has no elements.');
      }
    });

    slidesArray.forEach(slide => {
      writeLog("Slide ID: " + slide.objectId);
      if (slide.pageElements) {
        slide.pageElements.forEach(element => {
          writeLog("  Element ID: " + element.objectId);
        });
      } else {
        writeLog("  Este slide no tiene elementos.");
      }
    });

    // Establece los permisos y transfiere la propiedad.
    await setPermissionsAndTransferOwnership(presentationId);

    // Actualiza el slide duplicado
    await updateDuplicatedSlide(presentationId, 'p7_copy1');
  } catch (error) {
    console.error('Error al crear la presentación:', error);
  }
}

createPresentation();

/**
 * Función para transferir permisos y la propiedad del archivo
 */
async function setPermissionsAndTransferOwnership(fileId) {
  try {
    // Otorga permisos de edición a toda la organización.
    await drive.permissions.create({
      fileId,
      requestBody: {
        role: 'writer',
        type: 'domain',
        domain: 'plunkton.com'
      }
    });
    writeLog('Permisos de edición para la organización asignados.');

    // Transfiere la propiedad a mok@plunkton.com.
    await drive.permissions.create({
      fileId,
      transferOwnership: true,
      requestBody: {
        role: 'owner',
        type: 'user',
        emailAddress: 'mok@plunkton.com'
      }
    });
    writeLog('Propiedad transferida a mok@plunkton.com.');
  } catch (error) {
    console.error('Error al establecer los permisos o transferir la propiedad:', error);
  }
}

/**
 * Después de duplicar el slide, accede a los elementos del nuevo slide y modifica el texto del objeto deseado
 */
async function updateDuplicatedSlide(presentationId, duplicatedSlideId) {
  try {
    // Obtén los detalles del slide duplicado
    const presentation = await slides.presentations.get({ presentationId });
    const duplicatedSlide = presentation.data.slides.find(slide => slide.objectId === duplicatedSlideId);

    if (!duplicatedSlide) {
      console.error('Slide duplicado no encontrado');
      return;
    }

    // Busca el objeto con el título "Timer 1/4"
    const targetElement = duplicatedSlide.pageElements.find(element => element.title === 'Timer 1/4');

    if (!targetElement) {
      console.error('Elemento con título "Timer 1/4" no encontrado');
      return;
    }

    // Reemplaza el texto del objeto encontrado
    const requests = [
      {
        deleteText: {
          objectId: targetElement.objectId,
          textRange: { type: 'ALL' }
        }
      },
      {
        insertText: {
          objectId: targetElement.objectId,
          text: '25:00', // Nuevo texto para Timer 1/4
          insertionIndex: 0
        }
      }
    ];

    await slides.presentations.batchUpdate({
      presentationId,
      requestBody: { requests }
    });

    console.log('Texto del objeto actualizado correctamente');
  } catch (error) {
    console.error('Error al actualizar el slide duplicado:', error);
  }
}

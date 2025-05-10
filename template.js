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

// Autenticación con cuenta de servicio con delegación para suplantar a mok@plunkton.com
const auth = new google.auth.JWT({
  keyFile: 'i-monolith-453913-r1-b1bb363852b5.json',
  scopes: [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive'
  ],
  subject: 'mok@plunkton.com'  // Impersona a mok@plunkton.com
});

const slides = google.slides({ version: 'v1', auth });
const drive = google.drive({ version: 'v3', auth });

/**
 * JSON con las instrucciones para crear la presentación:
 * Se generan 4 slides con diferentes elementos, dejando el slide default que se crea automáticamente.
 */
const requests = [
  {
    createImage: {
      objectId: 'marca_g346dcc353d0_0_104',
      url: 'https://plunkton.com/images/clients/lions/marca-bancamiga.png',
      elementProperties: {
        pageObjectId: 'g346dcc353d0_0_104',
        size: { height: { magnitude: 58, unit: 'PT' }, width: { magnitude: 177, unit: 'PT' } },
        transform: {
          scaleX: 1, scaleY: 1,
          translateX: 53, translateY: 191,
          unit: 'PT'
        }
      }
    }
  },
  {
    createSlide: {
      objectId: 'T1-Snap1',
      slideLayoutReference: { predefinedLayout: 'BLANK' }
    }
  },
  // Asigna fondo con imagen a slide1
  {
    updatePageProperties: {
      objectId: 'T1-Snap1',
      pageProperties: {
        pageBackgroundFill: {
          stretchedPictureFill: {
            contentUrl: 'https://drive.google.com/uc?export=view&id=1npMt_cqKH7fV-3H5LEEaQKVzAy8Kowcp'
          }
        }
      },
      fields: 'pageBackgroundFill'
    }
  },
  {
    createImage: {
      objectId: 'fulbo_T1-Snap1',
      url: 'https://plunkton.com/images/clients/lions/lions.png',
      elementProperties: {
        pageObjectId: 'T1-Snap1',
        size: { height: { magnitude: 45, unit: 'PT' }, width: { magnitude: 95, unit: 'PT' } },
        transform: {
          scaleX: 1, scaleY: 1,
          translateX: 6, translateY: 348,
          unit: 'PT'
        }
      }
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
      fileId: '153SzG1Dapj2SO1PfQSAFObgTno8QZD4kosHJmbnJDEA',
      requestBody: {
        name: 'Lions'
      }
    });
    
    const presentationId = copyResponse.data.id;
    console.log('ID de la presentación:', presentationId);

    // Envía las instrucciones en batchUpdate para configurar la presentación.
    await slides.presentations.batchUpdate({
      presentationId,
      requestBody: { requests }
    });
    
    const presentationResponse = await slides.presentations.get({ presentationId });
    
    const slidesArray = presentationResponse.data.slides;


    slidesArray.forEach(slide => {
      console.log('Slide ID:', slide.objectId);

      if (slide.pageElements) {
        slide.pageElements.forEach(element => {
          console.log('  Element ID:', element.objectId);
          // Print some general properties of the element:
          console.log('  Element properties:', element);

          // If the element is a table, it will have a "table" property.
          if (element.table) {
            console.log('    This element is a table:');
            console.log('      Rows:', element.table.rows);
            console.log('      Columns:', element.table.columns);

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
                  console.log(`      Cell [${rowIndex}, ${colIndex}] text: "${cellText.trim()}"`);
                });
              });
            }
          }
        });
      } else {
        console.log('  This slide has no elements.');
      }
    });

    slidesArray.forEach(slide => {
      console.log("Slide ID:", slide.objectId);
      if (slide.pageElements) {
        slide.pageElements.forEach(element => {
          console.log("  Element ID:", element.objectId);
        });
      } else {
        console.log("  Este slide no tiene elementos.");
      }
    });

    // Establece los permisos y transfiere la propiedad.
    await setPermissionsAndTransferOwnership(presentationId);
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
    console.log('Permisos de edición para la organización asignados.');

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
    console.log('Propiedad transferida a mok@plunkton.com.');
  } catch (error) {
    console.error('Error al establecer los permisos o transferir la propiedad:', error);
  }
}

/**
 * Ejemplo de creación de presentación en Google Slides mediante Node.js
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
  // === PRIMER SLIDE ===
  {
    createSlide: {
      objectId: 'slide1',
      slideLayoutReference: { predefinedLayout: 'BLANK' }
    }
  },
  // Asigna fondo con imagen a slide1
  {
    updatePageProperties: {
      objectId: 'p',
      pageProperties: {
        pageBackgroundFill: {
          stretchedPictureFill: {
            contentUrl: 'https://plunkton.com/images/clients/lions/bg-fullhd.jpg'
          }
        }
      },
      fields: 'pageBackgroundFill'
    }
  },
  // Agrega un cuadro de texto para el Título en Slide
  {
    createShape: {
      objectId: 'title_slide1',
      shapeType: 'TEXT_BOX',
      elementProperties: {
        pageObjectId: 'slide1',
        size: { 
          height: { magnitude: 39, unit: 'PT' }, 
          width: { magnitude: 170, unit: 'PT' } },
        transform: {
          scaleX: 1, scaleY: 1,
          translateX: 41, translateY: 138,
          unit: 'PT'
        }
      }
    }
  },
  {
    insertText: {
      objectId: 'title_slide1',
      text: 'REPORTE',
    }
  },
  {
    updateTextStyle: {
      objectId: 'title_slide1',
      style: {
        bold: true,
        fontFamily: 'Inter',
        fontSize: {
          magnitude: 32,
          unit: 'PT'
        },
        foregroundColor: {
          opaqueColor: {
            rgbColor: { red: 1, green: 1, blue: 1 }
          }
        }
      },
      textRange: { type: 'ALL' },
      fields: 'bold,fontFamily,fontSize,foregroundColor'
    }
  },
  // Agrega dos logos en slide1
  {
    createImage: {
      objectId: 'fulbo_slide1',
      url: 'https://plunkton.com/images/clients/lions/fulbo.png',
      elementProperties: {
        pageObjectId: 'slide1',
        size: { height: { magnitude: 37, unit: 'PT' }, width: { magnitude: 37, unit: 'PT' } },
        transform: {
          scaleX: 1, scaleY: 1,
          translateX: 41, translateY: 85,
          unit: 'PT'
        }
      }
    }
  },
  {
    createImage: {
      objectId: 'marca_slide1',
      url: 'https://plunkton.com/images/clients/lions/marca-bancamiga.png',
      elementProperties: {
        pageObjectId: 'slide1',
        size: { height: { magnitude: 58, unit: 'PT' }, width: { magnitude: 177, unit: 'PT' } },
        transform: {
          scaleX: 1, scaleY: 1,
          translateX: 41, translateY: 191,
          unit: 'PT'
        }
      }
    }
  },

  {
    createImage: {
      objectId: 'pelota-red',
      url: 'https://plunkton.com/images/clients/lions/pelota-red.png',
      elementProperties: {
        pageObjectId: 'slide1',
        size: { height: { magnitude: 405, unit: 'PT' }, width: { magnitude: 205, unit: 'PT' } },
        transform: {
          scaleX: 1, scaleY: 1,
          translateX: 515, translateY: 0,
          unit: 'PT'
        }
      }
    }
  },

  // === SEGUNDO SLIDE ===
  {
    createSlide: {
      objectId: 'slide2',
      slideLayoutReference: { predefinedLayout: 'BLANK' }
    }
  },
  // Agrega título "Primer Tiempo" en slide2
  {
    createShape: {
      objectId: 'title_slide2',
      shapeType: 'TEXT_BOX',
      elementProperties: {
        pageObjectId: 'slide2',
        size: { height: { magnitude: 50, unit: 'PT' }, width: { magnitude: 400, unit: 'PT' } },
        transform: { scaleX: 1, scaleY: 1, translateX: 50, translateY: 20, unit: 'PT' }
      }
    }
  },
  {
    insertText: {
      objectId: 'title_slide2',
      text: 'Primer Tiempo'
    }
  },
  // Crea una tabla (cabecera + 5 filas = 6 filas, 3 columnas)
  {
    createTable: {
      objectId: 'table_slide2',
      elementProperties: {
        pageObjectId: 'slide2',
        size: { height: { magnitude: 200, unit: 'PT' }, width: { magnitude: 500, unit: 'PT' } },
        transform: { scaleX: 1, scaleY: 1, translateX: 50, translateY: 100, unit: 'PT' }
      },
      rows: 6,
      columns: 3
    }
  },
  // Inserta texto en la cabecera de la tabla
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 0, columnIndex: 0 },
      text: 'cliente'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 0, columnIndex: 1 },
      text: 'minuto'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 0, columnIndex: 2 },
      text: 'duración'
    }
  },
  // Inserta datos en las 5 filas siguientes
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 1, columnIndex: 0 },
      text: 'Cliente A'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 1, columnIndex: 1 },
      text: '10'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 1, columnIndex: 2 },
      text: '30 min'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 2, columnIndex: 0 },
      text: 'Cliente B'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 2, columnIndex: 1 },
      text: '20'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 2, columnIndex: 2 },
      text: '25 min'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 3, columnIndex: 0 },
      text: 'Cliente C'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 3, columnIndex: 1 },
      text: '30'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 3, columnIndex: 2 },
      text: '40 min'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 4, columnIndex: 0 },
      text: 'Cliente D'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 4, columnIndex: 1 },
      text: '40'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 4, columnIndex: 2 },
      text: '35 min'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 5, columnIndex: 0 },
      text: 'Cliente E'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 5, columnIndex: 1 },
      text: '50'
    }
  },
  {
    insertText: {
      objectId: 'table_slide2',
      cellLocation: { rowIndex: 5, columnIndex: 2 },
      text: '45 min'
    }
  },

  // === TERCER SLIDE (Slide 3) ===
  {
    createSlide: {
      objectId: 'slide3',
      slideLayoutReference: { predefinedLayout: 'BLANK' }
    }
  },
  {
    updatePageProperties: {
      objectId: 'slide3',
      pageProperties: {
        pageBackgroundFill: {
          stretchedPictureFill: {
            contentUrl: 'https://fakeimg.pl/800x600?text=partido'
          }
        }
      },
      fields: 'pageBackgroundFill'
    }
  },
  {
    createShape: {
      objectId: 'title_slide3',
      shapeType: 'TEXT_BOX',
      elementProperties: {
        pageObjectId: 'slide3',
        size: { height: { magnitude: 50, unit: 'PT' }, width: { magnitude: 300, unit: 'PT' } },
        transform: { scaleX: 1, scaleY: 1, translateX: 20, translateY: 20, unit: 'PT' }
      }
    }
  },
  {
    insertText: {
      objectId: 'title_slide3',
      text: 'Título Slide 3'
    }
  },
  {
    createImage: {
      objectId: 'logo_empresa_slide3',
      url: 'https://fakeimg.pl/150x70?text=Empresa',
      elementProperties: {
        pageObjectId: 'slide3',
        size: { height: { magnitude: 50, unit: 'PT' }, width: { magnitude: 50, unit: 'PT' } },
        transform: { scaleX: 1, scaleY: 1, translateX: 680, translateY: 20, unit: 'PT' }
      }
    }
  },

  // === CUARTO SLIDE (Slide 4) ===
  {
    createSlide: {
      objectId: 'slide4',
      slideLayoutReference: { predefinedLayout: 'BLANK' }
    }
  },
  {
    updatePageProperties: {
      objectId: 'slide4',
      pageProperties: {
        pageBackgroundFill: {
          stretchedPictureFill: {
            contentUrl: 'https://fakeimg.pl/800x600?text=BG'
          }
        }
      },
      fields: 'pageBackgroundFill'
    }
  },
  {
    createShape: {
      objectId: 'title_slide4',
      shapeType: 'TEXT_BOX',
      elementProperties: {
        pageObjectId: 'slide4',
        size: { height: { magnitude: 50, unit: 'PT' }, width: { magnitude: 300, unit: 'PT' } },
        transform: { scaleX: 1, scaleY: 1, translateX: 20, translateY: 20, unit: 'PT' }
      }
    }
  },
  {
    insertText: {
      objectId: 'title_slide4',
      text: 'Título Slide 4'
    }
  },
  {
    createImage: {
      objectId: 'logo_empresa_slide4',
      url: 'https://fakeimg.pl/150x70?text=empresa',
      elementProperties: {
        pageObjectId: 'slide4',
        size: { height: { magnitude: 50, unit: 'PT' }, width: { magnitude: 50, unit: 'PT' } },
        transform: { scaleX: 1, scaleY: 1, translateX: 680, translateY: 20, unit: 'PT' }
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
    // Crea la presentación base con un título.
    const createResponse = await slides.presentations.create({
      requestBody: { title: 'Lions' }
    });
    const presentationId = createResponse.data.presentationId;
    console.log('ID de la presentación:', presentationId);

    // Envía las instrucciones en batchUpdate para configurar la presentación.
    const batchUpdateResponse = await slides.presentations.batchUpdate({
      presentationId,
      requestBody: { requests }
    });
    //console.log('Actualizaciones aplicadas:', batchUpdateResponse.data.replies);
    const presentationResponse = await slides.presentations.get({ presentationId });
    const slidesArray = presentationResponse.data.slides;
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

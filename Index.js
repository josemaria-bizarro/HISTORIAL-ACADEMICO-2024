const tablaCalif =SpreadsheetApp.openById('1VxNaZwRntqPYN12MhOqzPMlCBWl-vQU7BRb3JLdYXsE');
const tabCalifOrria=tablaCalif.getSheetByName('CALIFICACIONES');
const bdIkasle= SpreadsheetApp.openById('1kzmgvrbkQ7ehUslRxR4ghkpSeI2yT99M3h-Dyo-EEVY');
const bdIkasleOrria=bdIkasle.getSheetByName('ACTIVOS_FORMATEADO');
const catAcade=SpreadsheetApp.openById('1T5KQCvXiYsGewSoc_A8qugUwSs-ft1NNds4oaNgQxkw');
const catBoleta = SpreadsheetApp.openById("1PAxLZjZ1QYCwMRJ3U12bIzO1mwNma9CkSBhRQkSfDgE").getSheetByName("CATALOGO");
const asignaOrria=catAcade.getSheetByName('TABLA ASIGNATURAS');
const opcEduOrria=catAcade.getSheetByName('OPCIONES EDUCATIVAS')
const periEduca=catAcade.getSheetByName('PERIODOS EDUCATIVOS')
const plantilla=DocumentApp.openById('1TYItUFXhupC5pD6gGBi8AxmCXAOAadSt-9HLWKeIAcQ');
const sheetAct=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARATULA');
const sheetPlant=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PLANTILLA');
const sheetCalif=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CALIFICACIONES');
const sheetPaso=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PASO2');
const SheetDB=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB");
const prueba=tablaCalif.getSheetByName('prueba');
const plantillaBTA="https://docs.google.com/document/d/1TYItUFXhupC5pD6gGBi8AxmCXAOAadSt-9HLWKeIAcQ"
const plantillaBTC="https://docs.google.com/document/d/1-7o-nPLqBNLFEsS9RMGWMzednRswEm2a8mRAlAKsV_o";
const plantillaBTG="https://docs.google.com/document/d/1aYNnfunAQh70Zw24Z0n7AzDi0Q-IIpHIgtrk--fJaco";
const plantillaLAE="https://docs.google.com/document/d/1gTLvU7CrOQQ6x4rs-Ahwxk_ASJEMI5EANL_cF6XNWb4";
const plantillaBTD="https://docs.google.com/document/d/1cHv98Q5iWH9y_uNg2dSgJZOzfsD8mQFjiUZ6v0s7rcM";
const pdfFolderBTA="https://drive.google.com/drive/folders/1xy6gy4coE4GxmOpCEhgTR8b7sWBkhc0R";
const pdfFolderBTG="https://drive.google.com/drive/folders/1OaBnyqkG-QupoPfGZ8nXak5-hr_fq8ka";
const pdfFolderBTC="https://drive.google.com/drive/folders/1yOj7M_k8OQbn4pEsGyrXeiIsw-rpEtH0";
const pdfFolderLAD="https://drive.google.com/drive/folders/19JQ_5N2VnQf6o2pNGsEEtOrZQ7DKs2B2";
const pdfFolderBTD="https://drive.google.com/drive/folders/1SchfNryJIZ-IBV0sGgYQi7KolNpjx-i1";
const folderBta="https://drive.google.com/drive/folders/1xy6gy4coE4GxmOpCEhgTR8b7sWBkhc0R";

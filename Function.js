// @ts-nocheck
// ESTA ES LA FUNCION PARA GENERAR LOS PDF DESDE UNA PLANTILLA DE DOCS
let regUrl=[];
let histCal=[];
let arrayIkasleak=[];
let wal=0;

function onOpen()
 {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('OPCIONES')
   
   menu
      .addItem("Actualiza", 'actualiza')
      .addItem("Plantilla",'cargaPlantilla')
      .addItem("Historiales",'historiales')
  menu.addToUi();
}

function historiales()
{
  
//var carpetaBoletas =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARATULA").getRange(4, 21).getValue(); 
var lsr=sheetPlant.getLastRow();
//const tempFolder = DriveApp.getFolderById("1Hg3PUYw6XV6StmdRVog_61z0PMQESNzI"); //FOLDER PARA LOS ARCHIVOS DE TRABAJO

var hist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("urls") //HOJA DONDE SE ESCRIBEN LOS DATOS DE HISTORIALES
hist.getRange("A2:C").clearContent()

                        //***************CARGA LA URL DEL FOLDER DE SALIDA DE ACUER
var opceduca=sheetAct.getRange("E1").getDisplayValue();       //OBTIENE LA OPCION EDUCATIVA A TRABAJAR
switch(opceduca)
  {
    case "BACH. ALIMENTOS Y BEBIDAS":
        var pdfFolderA=pdfFolderBTG;
        var url_plantilla=plantillaBTG;
        armaPlantilla(pdfFolderA,url_plantilla,opceduca);
        break;
    case "BACH. DISEÑO GRÁFICO Y ARTE":
        var pdfFolderA=pdfFolderBTD;
        var url_plantilla=plantillaBTD;
        armaPlantilla(pdfFolderA,url_plantilla,opceduca);
        break;
    case "SAETI":
        var pdfFolderA=pdfFolderBTA;
        var url_plantilla=plantillaBTA;
        armaPlantilla(pdfFolderA,url_plantilla,opceduca);
        break;
    case "BACH. COMUNICACIÓN DIGITAL":
        var pdfFolderA=pdfFolderBTC;
        var url_plantilla=plantillaBTC;
        armaPlantilla(pdfFolderA,url_plantilla,opceduca);
        break;
    case "LIC. ADMINISTRACIÓN":
        var pdfFolderA=pdfFolderLAD;
        var url_plantilla=plantillaLAE;
        armaPlantilla(pdfFolderA,url_plantilla,opceduca);
        break;
    default:
        console.log("error en folder a trabajar");
        return;
  }
}

function armaPlantilla(pdfFolderA,url_plantilla,opceduca)
{
var encuentra = pdfFolderA.lastIndexOf("/")+1   // Rutina que obtiene y recorta solo la ID *********************
var pdfFolderI =pdfFolderA.substr(encuentra,50);

const pdfFolder = DriveApp.getFolderById(pdfFolderI); //folder donde se depositan los historiales

const currenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANTILLA");

var FECEMI = Utilities.formatDate(new Date(), "GMT-6", "dd/MMM/yyyy")

//*******************************  DETERMINA EL PERIODO DE PROCESO */
//const PERIODO = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANTILLA").getRange("B4").getValue();
//calcula el ultimo periodo en base a la fecha de elaboracion que sea mayor a la fecha de fin de periodo
var wperiodos=periEduca.getDataRange().getValues();
var wperidosF=wperiodos.filter(ilara=>ilara[2]==opceduca);
var indPer=0;
if (wperidosF.length>0)
{
  var date1=new Date(FECEMI).getTime();
  wperidosF.forEach(ilara=>
  {
   var date2=new Date(ilara[5]).getTime();
    if (date1>date2)
    {
      indPer++
    }
    else
    {
      console.log("menorn que "+FECEMI+" "+ilara[5])
      console.log(date1+" "+date2)
    }
  } 
  )
  PERIODO=wperidosF[indPer-1][1];
} 
else
{
  console.log("error en busqueda de periodos")
  return
}  
//*******************obtiene url de plantilla segun Opc Edu */
var encuentra = url_plantilla.lastIndexOf("/")+1
var plantillaUrl =url_plantilla.substr(encuentra,50);
var docFile=DriveApp.getFileById(plantillaUrl)                //plantilla de historial segun opc edu

const tempFolder = DriveApp.getFolderById("1Hg3PUYw6XV6StmdRVog_61z0PMQESNzI");

//***************************************************** */
// CATALOGO PARA POSTERIORMENTE ENVIAR HISTORIAL
//***************************************************** */
catHist = SpreadsheetApp.openById("1PAxLZjZ1QYCwMRJ3U12bIzO1mwNma9CkSBhRQkSfDgE").getSheetByName("CATALOGO");
catHistLr=catHist.getLastRow()+1


var wcstLr=currenSheet.getLastRow();
var wcsLc=currenSheet.getLastColumn();
const dataW = currenSheet.getRange(6,1,wcstLr,wcsLc).getDisplayValues();  //OBTIENE DATOS A PROCESAR DESDE HOJA PLANTILLA

var data=dataW.filter(ilara=>ilara[0]=="*");
if(data.length>0)
{
      wal=0;
      var indx=(data[0].length)-1
      var sumax=0;
      
      for(var u=0;u<data[0].length;u++)
      {
      
        if(data[0][indx]=="")
        {
          
          sumax++;
          indx=indx-1
        }
      }
        wal=(data[0].length-1)-sumax;
  
 
 //******************************************************** */

    switch(opceduca)
    {
      case "BACH. ALIMENTOS Y BEBIDAS":
        createPDFBTG(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal)
          break;
      case "BACH. DISEÑO GRÁFICO Y ARTE":
        createPDFBTD(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal);
          break;
      case "SAETI":
            createPDFBTA(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal); 
          break;
      case "BACH. COMUNICACIÓN DIGITAL":
          createPDFBTc(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal);
          break;
      case "LIC. ADMINISTRACIÓN":
        createPDFLAD(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal);
          break;
      default:
          console.log("error en folder a trabajar");
          return;
    }

    //******************************************************** */             

    var histS=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("urls")
    histS.getRange(2,1,regUrl.length,3).setValues(regUrl)

    catBoleta.getRange(catHistLr,1,histCal.length,10).setValues(histCal)
  }
  else
  {
    console.log("no hay alumnos a procesar")
  }
}

function createPDFBTA(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal)
{
var ikasleDatuak=bdIkasleOrria.getDataRange().getDisplayValues(); //datos alumnos
data.forEach(ilara=>            //PROCESA CADA ALUMNO SELECCIONADO 
  {
          const tempFile = docFile.makeCopy(tempFolder);
          const tempDocFile = DocumentApp.openById(tempFile.getId());
          const body = tempDocFile.getBody();
          const head = tempDocFile.getHeader();

          var ikasleW = ikasleDatuak.filter(fila=>fila[0]==ilara[1])
          
          if (ikasleW.length>0)
          {
            console.log("ikaslew ok")
            NOMBRE=ikasleW[0][0];
            GRUPO=ikasleW[0][5];
            PLANTEL=ikasleW[0][22];
            TURNO=ikasleW[0][8];
            var FING=ikasleW[0][10];
            ID=ikasleW[0][2];
          }
          else
          {
            console.log(ilara[1])
          }
        
        //*** ARMA EL HEADER */

        var CURSADAS= ilara[wal];
        var PROMEDIO= ilara[wal-1];
        var crdts= ilara[wal-2];
          head.replaceText("{OPCION}",opceduca);
          head.replaceText("{ID}",ID);
          head.replaceText("{NOMBRE}",NOMBRE);
          head.replaceText("{GRUPO}",GRUPO);
          head.replaceText("{PLANTEL}",PLANTEL);
          head.replaceText("{TURNO}",TURNO);
          head.replaceText("{PERIODO}",PERIODO);  
          head.replaceText("{PROMG}",PROMEDIO);
          head.replaceText("{FECEMI}",FECEMI);
          head.replaceText("{CURSADAS}",CURSADAS);  
          head.replaceText("{FING}",FING);
          head.replaceText("{CRDTS}",crdts);
        
      //***********************************************************************************

          body.replaceText("{ID01}",ilara[2]);
          body.replaceText("{NAM01}",ilara[3]);
          body.replaceText("{CAL01}",ilara[4]);
          body.replaceText("{TIP01}",ilara[5]);
          body.replaceText("{FCURS1}",ilara[6]);

          body.replaceText("{ID02}", ilara[7]);
          body.replaceText("{NAM02}",ilara[8]);
          body.replaceText("{CAL02}",ilara[9]);
          body.replaceText("{TIP02}",ilara[10]);
          body.replaceText("{FCURS2}",ilara[11]);

          body.replaceText("{ID03}", ilara[12]);
          body.replaceText("{NAM03}",ilara[13]);
          body.replaceText("{CAL03}",ilara[14]);
          body.replaceText("{TIP03}",ilara[15]);
          body.replaceText("{FCURS3}",ilara[16]);

          body.replaceText("{ID04}", ilara[17]);
          body.replaceText("{NAM04}",ilara[18]);
          body.replaceText("{CAL04}",ilara[19]);
          body.replaceText("{TIP04}",ilara[20]);
          body.replaceText("{FCURS4}",ilara[21]);

          body.replaceText("{ID05}", ilara[22]);
          body.replaceText("{NAM05}",ilara[23]);
          body.replaceText("{CAL05}",ilara[24]);
          body.replaceText("{TIP05}",ilara[25]);
          body.replaceText("{FCURS5}",ilara[26]);

          body.replaceText("{ID06}", ilara[27]);
          body.replaceText("{NAM06}",ilara[28]);
          body.replaceText("{CAL06}",ilara[29]);
          body.replaceText("{TIP06}",ilara[30]);
          body.replaceText("{FCURS6}",ilara[31]);

          body.replaceText("{ID07}", ilara[32]);
          body.replaceText("{NAM07}",ilara[33]);
          body.replaceText("{CAL07}",ilara[34]);
          body.replaceText("{TIP07}",ilara[35]);
          body.replaceText("{FCURS7}",ilara[36]);

          body.replaceText("{ID08}", ilara[37]);
          body.replaceText("{NAM08}",ilara[38]);
          body.replaceText("{CAL08}",ilara[39]);
          body.replaceText("{TIP08}",ilara[40]);
          body.replaceText("{FCURS8}",ilara[41]);

          body.replaceText("{ID09}", ilara[42]);
          body.replaceText("{NAM09}",ilara[43]);
          body.replaceText("{CAL09}",ilara[44]);
          body.replaceText("{TIP09}",ilara[45]);
          body.replaceText("{FCURS9}",ilara[46]);

          body.replaceText("{ID10}", ilara[47]);
          body.replaceText("{NAM10}",ilara[48]);
          body.replaceText("{CAL10}",ilara[49]);
          body.replaceText("{TIP10}",ilara[50]);
          body.replaceText("{FCURS10}",ilara[51]);

          body.replaceText("{ID11}", ilara[52]);
          body.replaceText("{NAM11}",ilara[53]);
          body.replaceText("{CAL11}",ilara[54]);
          body.replaceText("{TIP11}",ilara[55]);
          body.replaceText("{FCURS11}",ilara[56]);

          body.replaceText("{ID12}", ilara[57]);
          body.replaceText("{NAM12}",ilara[58]);
          body.replaceText("{CAL12}",ilara[59]);
          body.replaceText("{TIP12}",ilara[60]);
          body.replaceText("{FCURS12}",ilara[61]);

          body.replaceText("{ID13}", ilara[62]);
          body.replaceText("{NAM13}",ilara[63]);
          body.replaceText("{CAL13}",ilara[64]);
          body.replaceText("{TIP13}",ilara[65]);
          body.replaceText("{FCURS13}",ilara[66]);

          body.replaceText("{ID14}", ilara[67]);
          body.replaceText("{NAM14}",ilara[68]);
          body.replaceText("{CAL14}",ilara[69]);
          body.replaceText("{TIP14}",ilara[70]);
          body.replaceText("{FCURS14}",ilara[71]);

          body.replaceText("{ID15}", ilara[72]);
          body.replaceText("{NAM15}",ilara[73]);
          body.replaceText("{CAL15}",ilara[74]);
          body.replaceText("{TIP15}",ilara[75]);
          body.replaceText("{FCURS15}",ilara[76]);

          body.replaceText("{ID16}", ilara[77]);
          body.replaceText("{NAM16}",ilara[78]);
          body.replaceText("{CAL16}",ilara[79]);
          body.replaceText("{TIP16}",ilara[80]);
          body.replaceText("{FCURS16}",ilara[81]);

          body.replaceText("{ID17}", ilara[82]);
          body.replaceText("{NAM17}",ilara[83]);
          body.replaceText("{CAL17}",ilara[84]);
          body.replaceText("{TIP17}",ilara[85]);
          body.replaceText("{FCURS17}",ilara[86]);

          body.replaceText("{ID18}", ilara[87]);
          body.replaceText("{NAM18}",ilara[88]);
          body.replaceText("{CAL18}",ilara[89]);
          body.replaceText("{TIP18}",ilara[90]);
          body.replaceText("{FCURS18}",ilara[91]);

          body.replaceText("{ID19}", ilara[92]);
          body.replaceText("{NAM19}",ilara[93]);
          body.replaceText("{CAL19}",ilara[94]);
          body.replaceText("{TIP19}",ilara[95]);
          body.replaceText("{FCURS19}",ilara[96]);

          body.replaceText("{ID20}", ilara[97]);
          body.replaceText("{NAM20}",ilara[98]);
          body.replaceText("{CAL20}",ilara[99]);
          body.replaceText("{TIP20}",ilara[100]);
          body.replaceText("{FCURS20}",ilara[101]);

          body.replaceText("{ID21}", ilara[102]);
          body.replaceText("{NAM21}",ilara[103]);
          body.replaceText("{CAL21}",ilara[104]);
          body.replaceText("{TIP21}",ilara[105]);
          body.replaceText("{FCURS21}",ilara[106]);
          
          body.replaceText("{ID22}", ilara[107]);
          body.replaceText("{NAM22}",ilara[108]);
          body.replaceText("{CAL22}",ilara[109]);
          body.replaceText("{TIP22}",ilara[110]);
          body.replaceText("{FCURS22}",ilara[111]);

          body.replaceText("{ID23}", ilara[112]);
          body.replaceText("{NAM23}",ilara[113]);
          body.replaceText("{CAL23}",ilara[114]);
          body.replaceText("{TIP23}",ilara[115]);
          body.replaceText("{FCURS23}",ilara[116]);

          body.replaceText("{ID24}",ilara[117]);
          body.replaceText("{NAM24}",ilara[118]);
          body.replaceText("{CAL24}",ilara[119]);
          body.replaceText("{TIP24}",ilara[120]);
          body.replaceText("{FCURS24}",ilara[121]);
          
          body.replaceText("{ID25}", ilara[122]);
          body.replaceText("{NAM25}",ilara[123]);
          body.replaceText("{CAL25}",ilara[124]);
          body.replaceText("{TIP25}",ilara[125]);
          body.replaceText("{FCURS25}",ilara[126]);

          body.replaceText("{ID26}", ilara[127]);
          body.replaceText("{NAM26}",ilara[128]);
          body.replaceText("{CAL26}",ilara[129]);
          body.replaceText("{TIP26}",ilara[130]);
          body.replaceText("{FCURS26}",ilara[131]);

          body.replaceText("{ID27}", ilara[132]);
          body.replaceText("{NAM27}",ilara[133]);
          body.replaceText("{CAL27}",ilara[134]);
          body.replaceText("{TIP27}",ilara[135]);
          body.replaceText("{FCURS27}",ilara[136]);

          body.replaceText("{ID28}", ilara[137]);
          body.replaceText("{NAM28}",ilara[138]);
          body.replaceText("{CAL28}",ilara[139]);
          body.replaceText("{TIP28}",ilara[140]);
          body.replaceText("{FCURS28}",ilara[141]);

          body.replaceText("{ID29}", ilara[142]);
          body.replaceText("{NAM29}",ilara[143]);
          body.replaceText("{CAL29}",ilara[144]);
          body.replaceText("{TIP29}",ilara[145]);
          body.replaceText("{FCURS29}",ilara[146]);

          body.replaceText("{ID30}", ilara[147]);
          body.replaceText("{NAM30}",ilara[148]);
          body.replaceText("{CAL30}",ilara[149]);
          body.replaceText("{TIP30}",ilara[150]);
          body.replaceText("{FCURS30}",ilara[151]);

          body.replaceText("{ID31}", ilara[152]);
          body.replaceText("{NAM31}",ilara[153]);
          body.replaceText("{CAL31}",ilara[154]);
          body.replaceText("{TIP31}",ilara[155]);
          body.replaceText("{FCURS31}",ilara[156]);

          body.replaceText("{ID32}", ilara[157]);
          body.replaceText("{NAM32}",ilara[158]);
          body.replaceText("{CAL32}",ilara[159]);
          body.replaceText("{TIP32}",ilara[160]);
          body.replaceText("{FCURS32}",ilara[161]);

          body.replaceText("{ID33}", ilara[162]);
          body.replaceText("{NAM33}",ilara[163]);
          body.replaceText("{CAL33}",ilara[164]);
          body.replaceText("{TIP33}",ilara[165]);
          body.replaceText("{FCURS33}",ilara[166]);

          body.replaceText("{ID34}", ilara[167]);
          body.replaceText("{NAM34}",ilara[168]);
          body.replaceText("{CAL34}",ilara[169]);
          body.replaceText("{TIP34}",ilara[170]);
          body.replaceText("{FCURS34}",ilara[171]);

          body.replaceText("{ID35}", ilara[172]);
          body.replaceText("{NAM35}",ilara[173]);
          body.replaceText("{CAL35}",ilara[174]);
          body.replaceText("{TIP35}",ilara[175]);
          body.replaceText("{FCURS35}",ilara[176]);

          body.replaceText("{ID36}", ilara[177]);
          body.replaceText("{NAM36}",ilara[178]);
          body.replaceText("{CAL36}",ilara[179]);
          body.replaceText("{TIP36}",ilara[180]);
          body.replaceText("{FCURS36}",ilara[181]);

          body.replaceText("{ID37}", ilara[182]);
          body.replaceText("{NAM37}",ilara[183]);
          body.replaceText("{CAL37}",ilara[184]);
          body.replaceText("{TIP37}",ilara[185]);
          body.replaceText("{FCURS37}",ilara[186]);

          body.replaceText("{ID38}", ilara[187]);
          body.replaceText("{NAM38}",ilara[188]);
          body.replaceText("{CAL38}",ilara[189]);
          body.replaceText("{TIP38}",ilara[190]);
          body.replaceText("{FCURS38}",ilara[191]);

          body.replaceText("{ID39}", ilara[192]);
          body.replaceText("{NAM39}",ilara[193]);
          body.replaceText("{CAL39}",ilara[194]);
          body.replaceText("{TIP39}",ilara[195]);
          body.replaceText("{FCURS39}",ilara[196]);

          body.replaceText("{ID40}", ilara[197]);
          body.replaceText("{NAM40}",ilara[198]);
          body.replaceText("{CAL40}",ilara[199]);
          body.replaceText("{TIP40}",ilara[200]);
          body.replaceText("{FCURS40}",ilara[201]);

          body.replaceText("{ID41}", ilara[202]);
          body.replaceText("{NAM41}",ilara[203]);
          body.replaceText("{CAL41}",ilara[204]);
          body.replaceText("{TIP41}",ilara[205]);
          body.replaceText("{FCURS41}",ilara[206]);

          body.replaceText("{ID42}", ilara[207]);
          body.replaceText("{NAM42}",ilara[208]);
          body.replaceText("{CAL42}",ilara[209]);
          body.replaceText("{TIP42}",ilara[210]);
          body.replaceText("{FCURS42}",ilara[211]);

          body.replaceText("{ID43}", ilara[212]);
          body.replaceText("{NAM43}",ilara[213]);
          body.replaceText("{CAL43}",ilara[214]);
          body.replaceText("{TIP43}",ilara[215]);
          body.replaceText("{FCURS43}",ilara[216]);

          body.replaceText("{ID44}", ilara[217]);
          body.replaceText("{NAM44}",ilara[218]);
          body.replaceText("{CAL44}",ilara[219]);
          body.replaceText("{TIP44}",ilara[220]);
          body.replaceText("{FCURS44}",ilara[221]);

          body.replaceText("{ID45}", ilara[222]);
          body.replaceText("{NAM45}",ilara[223]);
          body.replaceText("{CAL45}",ilara[224]);
          body.replaceText("{TIP45}",ilara[225]);
          body.replaceText("{FCURS45}",ilara[226]);

          body.replaceText("{ID46}", ilara[227]);
          body.replaceText("{NAM46}",ilara[228]);
          body.replaceText("{CAL46}",ilara[229]);
          body.replaceText("{TIP46}",ilara[230]);
          body.replaceText("{FCURS46}",ilara[231]);

          body.replaceText("{ID47}", ilara[232]);
          body.replaceText("{NAM47}",ilara[233]);
          body.replaceText("{CAL47}",ilara[234]);
          body.replaceText("{TIP47}",ilara[235]);
          body.replaceText("{FCURS47}",ilara[236]);

          finalizaImpresion(NOMBRE,GRUPO,ID,opceduca,PLANTEL,PERIODO,tempDocFile,pdfFolder,tempFolder,tempFile)
    })
}
//*********************************************** */
//*          PLANTILLA DE GASTRO

function createPDFBTG(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal)
{
var ikasleDatuak=bdIkasleOrria.getDataRange().getDisplayValues(); //datos alumnos
data.forEach(ilara=>            //PROCESA CADA ALUMNO SELECCIONADO 
  {
          const tempFile = docFile.makeCopy(tempFolder);
          const tempDocFile = DocumentApp.openById(tempFile.getId());
          const body = tempDocFile.getBody();
          const head = tempDocFile.getHeader();

          var ikasleW = ikasleDatuak.filter(fila=>fila[0]==ilara[1])
          
          if (ikasleW.length>0)
          {
            console.log("ikaslew ok")
            NOMBRE=ikasleW[0][0];
            GRUPO=ikasleW[0][5];
            PLANTEL=ikasleW[0][22];
            TURNO=ikasleW[0][8];
            var FING=ikasleW[0][10];
            ID=ikasleW[0][2];
          }
          else
          {
            console.log(ilara[1])
          }
        
        //*** ARMA EL HEADER */

        var CURSADAS= ilara[wal];
        var PROMEDIO= ilara[wal-1];
        var especialidad="CON ESPECIALIDAD EN ALIMENTOS Y BEBIDAS"
        var CUATI="CUATRIMESTRE 1";
        var CUATII="CUATRIMESTRE 2";
        var CUATIII="CUATRIMESTRE 3";
        var CUATIV="CUATRIMESTRE 4";
        var CUATV="CUATRIMESTRE 5";
        var CUATVI="CUATRIMESTRE 6";
        var CUATVII="CUATRIMESTRE 7";
          head.replaceText("{OPCION}",opceduca);
          head.replaceText("{ID}",ID);
          head.replaceText("{NOMBRE}",NOMBRE);
          head.replaceText("{GRUPO}",GRUPO);
          head.replaceText("{PLANTEL}",PLANTEL);
          head.replaceText("{TURNO}",TURNO);
          head.replaceText("{PERIODO}",PERIODO);  
          head.replaceText("{PROMG}",PROMEDIO);
          head.replaceText("{FECEMI}",FECEMI);
          head.replaceText("{CURSADAS}",CURSADAS);  
          head.replaceText("{FING}",FING);
          head.replaceText("{especialidad}",especialidad);
        
      //***********************************************************************************

          body.replaceText("{CUATI}",CUATI);
          body.replaceText("{CUATII}",CUATII);
          body.replaceText("{CUATIII}",CUATIII);
          body.replaceText("{CUATIV}",CUATIV);
          body.replaceText("{CUATV}",CUATV);
          body.replaceText("{CUATVI}",CUATVI);
          body.replaceText("{CUATVII}",CUATVII);


          body.replaceText("{ID01}",ilara[2]);
          body.replaceText("{NAM01}",ilara[3]);
          body.replaceText("{CAL01}",ilara[4]);
          body.replaceText("{TIP01}",ilara[5]);
          body.replaceText("{FCURS1}",ilara[6]);

          body.replaceText("{ID02}", ilara[7]);
          body.replaceText("{NAM02}",ilara[8]);
          body.replaceText("{CAL02}",ilara[9]);
          body.replaceText("{TIP02}",ilara[10]);
          body.replaceText("{FCURS2}",ilara[11]);

          body.replaceText("{ID03}", ilara[12]);
          body.replaceText("{NAM03}",ilara[13]);
          body.replaceText("{CAL03}",ilara[14]);
          body.replaceText("{TIP03}",ilara[15]);
          body.replaceText("{FCURS3}",ilara[16]);

          body.replaceText("{ID04}", ilara[17]);
          body.replaceText("{NAM04}",ilara[18]);
          body.replaceText("{CAL04}",ilara[19]);
          body.replaceText("{TIP04}",ilara[20]);
          body.replaceText("{FCURS4}",ilara[21]);

          body.replaceText("{ID05}", ilara[22]);
          body.replaceText("{NAM05}",ilara[23]);
          body.replaceText("{CAL05}",ilara[24]);
          body.replaceText("{TIP05}",ilara[25]);
          body.replaceText("{FCURS5}",ilara[26]);

          body.replaceText("{ID06}", ilara[27]);
          body.replaceText("{NAM06}",ilara[28]);
          body.replaceText("{CAL06}",ilara[29]);
          body.replaceText("{TIP06}",ilara[30]);
          body.replaceText("{FCURS6}",ilara[31]);

          body.replaceText("{ID07}", ilara[32]);
          body.replaceText("{NAM07}",ilara[33]);
          body.replaceText("{CAL07}",ilara[34]);
          body.replaceText("{TIP07}",ilara[35]);
          body.replaceText("{FCURS7}",ilara[36]);

          body.replaceText("{ID08}", ilara[37]);
          body.replaceText("{NAM08}",ilara[38]);
          body.replaceText("{CAL08}",ilara[39]);
          body.replaceText("{TIP08}",ilara[40]);
          body.replaceText("{FCURS8}",ilara[41]);

          body.replaceText("{ID09}", ilara[42]);
          body.replaceText("{NAM09}",ilara[43]);
          body.replaceText("{CAL09}",ilara[44]);
          body.replaceText("{TIP09}",ilara[45]);
          body.replaceText("{FCURS9}",ilara[46]);

          body.replaceText("{ID10}", ilara[47]);
          body.replaceText("{NAM10}",ilara[48]);
          body.replaceText("{CAL10}",ilara[49]);
          body.replaceText("{TIP10}",ilara[50]);
          body.replaceText("{FCURS10}",ilara[51]);

          body.replaceText("{ID11}", ilara[52]);
          body.replaceText("{NAM11}",ilara[53]);
          body.replaceText("{CAL11}",ilara[54]);
          body.replaceText("{TIP11}",ilara[55]);
          body.replaceText("{FCURS11}",ilara[56]);

          body.replaceText("{ID12}", ilara[57]);
          body.replaceText("{NAM12}",ilara[58]);
          body.replaceText("{CAL12}",ilara[59]);
          body.replaceText("{TIP12}",ilara[60]);
          body.replaceText("{FCURS12}",ilara[61]);

          body.replaceText("{ID13}", ilara[62]);
          body.replaceText("{NAM13}",ilara[63]);
          body.replaceText("{CAL13}",ilara[64]);
          body.replaceText("{TIP13}",ilara[65]);
          body.replaceText("{FCURS13}",ilara[66]);

          body.replaceText("{ID14}", ilara[67]);
          body.replaceText("{NAM14}",ilara[68]);
          body.replaceText("{CAL14}",ilara[69]);
          body.replaceText("{TIP14}",ilara[70]);
          body.replaceText("{FCURS14}",ilara[71]);

          body.replaceText("{ID15}", ilara[72]);
          body.replaceText("{NAM15}",ilara[73]);
          body.replaceText("{CAL15}",ilara[74]);
          body.replaceText("{TIP15}",ilara[75]);
          body.replaceText("{FCURS15}",ilara[76]);

          body.replaceText("{ID16}", ilara[77]);
          body.replaceText("{NAM16}",ilara[78]);
          body.replaceText("{CAL16}",ilara[79]);
          body.replaceText("{TIP16}",ilara[80]);
          body.replaceText("{FCURS16}",ilara[81]);

          body.replaceText("{ID17}", ilara[82]);
          body.replaceText("{NAM17}",ilara[83]);
          body.replaceText("{CAL17}",ilara[84]);
          body.replaceText("{TIP17}",ilara[85]);
          body.replaceText("{FCURS17}",ilara[86]);

          body.replaceText("{ID18}", ilara[87]);
          body.replaceText("{NAM18}",ilara[88]);
          body.replaceText("{CAL18}",ilara[89]);
          body.replaceText("{TIP18}",ilara[90]);
          body.replaceText("{FCURS18}",ilara[91]);

          body.replaceText("{ID19}", ilara[92]);
          body.replaceText("{NAM19}",ilara[93]);
          body.replaceText("{CAL19}",ilara[94]);
          body.replaceText("{TIP19}",ilara[95]);
          body.replaceText("{FCURS19}",ilara[96]);

          body.replaceText("{ID20}", ilara[97]);
          body.replaceText("{NAM20}",ilara[98]);
          body.replaceText("{CAL20}",ilara[99]);
          body.replaceText("{TIP20}",ilara[100]);
          body.replaceText("{FCURS20}",ilara[101]);

          body.replaceText("{ID21}", ilara[102]);
          body.replaceText("{NAM21}",ilara[103]);
          body.replaceText("{CAL21}",ilara[104]);
          body.replaceText("{TIP21}",ilara[105]);
          body.replaceText("{FCURS21}",ilara[106]);
          
          body.replaceText("{ID22}", ilara[107]);
          body.replaceText("{NAM22}",ilara[108]);
          body.replaceText("{CAL22}",ilara[109]);
          body.replaceText("{TIP22}",ilara[110]);
          body.replaceText("{FCURS22}",ilara[111]);

          body.replaceText("{ID23}", ilara[112]);
          body.replaceText("{NAM23}",ilara[113]);
          body.replaceText("{CAL23}",ilara[114]);
          body.replaceText("{TIP23}",ilara[115]);
          body.replaceText("{FCURS23}",ilara[116]);

          body.replaceText("{ID24}",ilara[117]);
          body.replaceText("{NAM24}",ilara[118]);
          body.replaceText("{CAL24}",ilara[119]);
          body.replaceText("{TIP24}",ilara[120]);
          body.replaceText("{FCURS24}",ilara[121]);
          
          body.replaceText("{ID25}", ilara[122]);
          body.replaceText("{NAM25}",ilara[123]);
          body.replaceText("{CAL25}",ilara[124]);
          body.replaceText("{TIP25}",ilara[125]);
          body.replaceText("{FCURS25}",ilara[126]);

          body.replaceText("{ID26}", ilara[127]);
          body.replaceText("{NAM26}",ilara[128]);
          body.replaceText("{CAL26}",ilara[129]);
          body.replaceText("{TIP26}",ilara[130]);
          body.replaceText("{FCURS26}",ilara[131]);

          body.replaceText("{ID27}", ilara[132]);
          body.replaceText("{NAM27}",ilara[133]);
          body.replaceText("{CAL27}",ilara[134]);
          body.replaceText("{TIP27}",ilara[135]);
          body.replaceText("{FCURS27}",ilara[136]);

          body.replaceText("{ID28}", ilara[137]);
          body.replaceText("{NAM28}",ilara[138]);
          body.replaceText("{CAL28}",ilara[139]);
          body.replaceText("{TIP28}",ilara[140]);
          body.replaceText("{FCURS28}",ilara[141]);

          body.replaceText("{ID29}", ilara[142]);
          body.replaceText("{NAM29}",ilara[143]);
          body.replaceText("{CAL29}",ilara[144]);
          body.replaceText("{TIP29}",ilara[145]);
          body.replaceText("{FCURS29}",ilara[146]);

          body.replaceText("{ID30}", ilara[147]);
          body.replaceText("{NAM30}",ilara[148]);
          body.replaceText("{CAL30}",ilara[149]);
          body.replaceText("{TIP30}",ilara[150]);
          body.replaceText("{FCURS30}",ilara[151]);

          body.replaceText("{ID31}", ilara[152]);
          body.replaceText("{NAM31}",ilara[153]);
          body.replaceText("{CAL31}",ilara[154]);
          body.replaceText("{TIP31}",ilara[155]);
          body.replaceText("{FCURS31}",ilara[156]);

          body.replaceText("{ID32}", ilara[157]);
          body.replaceText("{NAM32}",ilara[158]);
          body.replaceText("{CAL32}",ilara[159]);
          body.replaceText("{TIP32}",ilara[160]);
          body.replaceText("{FCURS32}",ilara[161]);

          body.replaceText("{ID33}", ilara[162]);
          body.replaceText("{NAM33}",ilara[163]);
          body.replaceText("{CAL33}",ilara[164]);
          body.replaceText("{TIP33}",ilara[165]);
          body.replaceText("{FCURS33}",ilara[166]);

          body.replaceText("{ID34}", ilara[167]);
          body.replaceText("{NAM34}",ilara[168]);
          body.replaceText("{CAL34}",ilara[169]);
          body.replaceText("{TIP34}",ilara[170]);
          body.replaceText("{FCURS34}",ilara[171]);

          body.replaceText("{ID35}", ilara[172]);
          body.replaceText("{NAM35}",ilara[173]);
          body.replaceText("{CAL35}",ilara[174]);
          body.replaceText("{TIP35}",ilara[175]);
          body.replaceText("{FCURS35}",ilara[176]);

          body.replaceText("{ID36}", ilara[177]);
          body.replaceText("{NAM36}",ilara[178]);
          body.replaceText("{CAL36}",ilara[179]);
          body.replaceText("{TIP36}",ilara[180]);
          body.replaceText("{FCURS36}",ilara[181]);

          body.replaceText("{ID37}", ilara[182]);
          body.replaceText("{NAM37}",ilara[183]);
          body.replaceText("{CAL37}",ilara[184]);
          body.replaceText("{TIP37}",ilara[185]);
          body.replaceText("{FCURS37}",ilara[186]);

          body.replaceText("{ID38}", ilara[187]);
          body.replaceText("{NAM38}",ilara[188]);
          body.replaceText("{CAL38}",ilara[189]);
          body.replaceText("{TIP38}",ilara[190]);
          body.replaceText("{FCURS38}",ilara[191]);

          body.replaceText("{ID39}", ilara[192]);
          body.replaceText("{NAM39}",ilara[193]);
          body.replaceText("{CAL39}",ilara[194]);
          body.replaceText("{TIP39}",ilara[195]);
          body.replaceText("{FCURS39}",ilara[196]);

          body.replaceText("{ID40}", ilara[197]);
          body.replaceText("{NAM40}",ilara[198]);
          body.replaceText("{CAL40}",ilara[199]);
          body.replaceText("{TIP40}",ilara[200]);
          body.replaceText("{FCURS40}",ilara[201]);

          body.replaceText("{ID41}", ilara[202]);
          body.replaceText("{NAM41}",ilara[203]);
          body.replaceText("{CAL41}",ilara[204]);
          body.replaceText("{TIP41}",ilara[205]);
          body.replaceText("{FCURS41}",ilara[206]);

          body.replaceText("{ID42}", ilara[207]);
          body.replaceText("{NAM42}",ilara[208]);
          body.replaceText("{CAL42}",ilara[209]);
          body.replaceText("{TIP42}",ilara[210]);
          body.replaceText("{FCURS42}",ilara[211]);

          body.replaceText("{ID43}", ilara[212]);
          body.replaceText("{NAM43}",ilara[213]);
          body.replaceText("{CAL43}",ilara[214]);
          body.replaceText("{TIP43}",ilara[215]);
          body.replaceText("{FCURS43}",ilara[216]);

          body.replaceText("{ID44}", ilara[217]);
          body.replaceText("{NAM44}",ilara[218]);
          body.replaceText("{CAL44}",ilara[219]);
          body.replaceText("{TIP44}",ilara[220]);
          body.replaceText("{FCURS44}",ilara[221]);

          body.replaceText("{ID45}", ilara[222]);
          body.replaceText("{NAM45}",ilara[223]);
          body.replaceText("{CAL45}",ilara[224]);
          body.replaceText("{TIP45}",ilara[225]);
          body.replaceText("{FCURS45}",ilara[226]);

          body.replaceText("{ID46}", ilara[227]);
          body.replaceText("{NAM46}",ilara[228]);
          body.replaceText("{CAL46}",ilara[229]);
          body.replaceText("{TIP46}",ilara[230]);
          body.replaceText("{FCURS46}",ilara[231]);

          body.replaceText("{ID47}", ilara[232]);
          body.replaceText("{NAM47}",ilara[233]);
          body.replaceText("{CAL47}",ilara[234]);
          body.replaceText("{TIP47}",ilara[235]);
          body.replaceText("{FCURS47}",ilara[236]);


          body.replaceText("{ID48}", ilara[237]);
          body.replaceText("{NAM48}",ilara[238]);
          body.replaceText("{CAL48}",ilara[239]);
          body.replaceText("{TIP48}",ilara[240]);
          body.replaceText("{FCURS48}",ilara[241]);

          body.replaceText("{ID49}", ilara[242]);
          body.replaceText("{NAM49}",ilara[243]);
          body.replaceText("{CAL49}",ilara[244]);
          body.replaceText("{TIP49}",ilara[245]);
          body.replaceText("{FCURS49}",ilara[246]);

          body.replaceText("{ID50}", ilara[247]);
          body.replaceText("{NAM50}",ilara[248]);
          body.replaceText("{CAL50}",ilara[249]);
          body.replaceText("{TIP50}",ilara[250]);
          body.replaceText("{FCURS50}",ilara[251]);

          body.replaceText("{ID51}", ilara[252]);
          body.replaceText("{NAM51}",ilara[253]);
          body.replaceText("{CAL51}",ilara[254]);
          body.replaceText("{TIP51}",ilara[255]);
          body.replaceText("{FCURS51}",ilara[256]);

          body.replaceText("{ID52}", ilara[257]);
          body.replaceText("{NAM52}",ilara[258]);
          body.replaceText("{CAL52}",ilara[259]);
          body.replaceText("{TIP52}",ilara[260]);
          body.replaceText("{FCURS52}",ilara[261]);

          body.replaceText("{ID53}", ilara[262]);
          body.replaceText("{NAM53}",ilara[263]);
          body.replaceText("{CAL53}",ilara[264]);
          body.replaceText("{TIP53}",ilara[265]);
          body.replaceText("{FCURS53}",ilara[266]);

          body.replaceText("{ID54}", ilara[267]);
          body.replaceText("{NAM54}",ilara[268]);
          body.replaceText("{CAL54}",ilara[269]);
          body.replaceText("{TIP54}",ilara[270]);
          body.replaceText("{FCURS54}",ilara[271]);

          body.replaceText("{ID55}", ilara[272]);
          body.replaceText("{NAM55}",ilara[273]);
          body.replaceText("{CAL55}",ilara[274]);
          body.replaceText("{TIP55}",ilara[275]);
          body.replaceText("{FCURS55}",ilara[276]);

          body.replaceText("{ID56}", ilara[277]);
          body.replaceText("{NAM56}",ilara[278]);
          body.replaceText("{CAL56}",ilara[279]);
          body.replaceText("{TIP56}",ilara[280]);
          body.replaceText("{FCURS56}",ilara[281]);

          body.replaceText("{ID57}", ilara[282]);
          body.replaceText("{NAM57}",ilara[283]);
          body.replaceText("{CAL57}",ilara[284]);
          body.replaceText("{TIP57}",ilara[285]);
          body.replaceText("{FCURS57}",ilara[286]);

          body.replaceText("{ID58}", ilara[287]);
          body.replaceText("{NAM58}",ilara[288]);
          body.replaceText("{CAL58}",ilara[289]);
          body.replaceText("{TIP58}",ilara[290]);
          body.replaceText("{FCURS58}",ilara[291]);

          body.replaceText("{ID59}", ilara[292]);
          body.replaceText("{NAM59}",ilara[293]);
          body.replaceText("{CAL59}",ilara[294]);
          body.replaceText("{TIP59}",ilara[295]);
          body.replaceText("{FCURS59}",ilara[296]);

          body.replaceText("{ID60}", ilara[297]);
          body.replaceText("{NAM60}",ilara[298]);
          body.replaceText("{CAL60}",ilara[299]);
          body.replaceText("{TIP60}",ilara[300]);
          body.replaceText("{FCURS60}",ilara[301]);

          body.replaceText("{ID61}", ilara[302]);
          body.replaceText("{NAM61}",ilara[303]);
          body.replaceText("{CAL61}",ilara[304]);
          body.replaceText("{TIP61}",ilara[305]);
          body.replaceText("{FCURS61}",ilara[306]);

          body.replaceText("{ID62}", ilara[307]);
          body.replaceText("{NAM62}",ilara[308]);
          body.replaceText("{CAL62}",ilara[309]);
          body.replaceText("{TIP62}",ilara[310]);
          body.replaceText("{FCURS62}",ilara[311]);

          body.replaceText("{ID63}", ilara[312]);
          body.replaceText("{NAM63}",ilara[313]);
          body.replaceText("{CAL63}",ilara[314]);
          body.replaceText("{TIP63}",ilara[315]);
          body.replaceText("{FCURS63}",ilara[316]);

          body.replaceText("{ID64}", ilara[317]);
          body.replaceText("{NAM64}",ilara[318]);
          body.replaceText("{CAL64}",ilara[319]);
          body.replaceText("{TIP64}",ilara[320]);
          body.replaceText("{FCURS64}",ilara[321]);

          body.replaceText("{ID65}", ilara[322]);
          body.replaceText("{NAM65}",ilara[323]);
          body.replaceText("{CAL65}",ilara[324]);
          body.replaceText("{TIP65}",ilara[325]);
          body.replaceText("{FCURS65}",ilara[326]);

          body.replaceText("{ID66}", ilara[327]);
          body.replaceText("{NAM66}",ilara[328]);
          body.replaceText("{CAL66}",ilara[329]);
          body.replaceText("{TIP66}",ilara[330]);
          body.replaceText("{FCURS66}",ilara[331]);

          finalizaImpresion(NOMBRE,GRUPO,ID,opceduca,PLANTEL,PERIODO,tempDocFile,pdfFolder,tempFolder,tempFile)
    })
}

//*********************************************** */

//*********************************************** */
//*          PLANTILLA DE COMUNICACIÓN

function createPDFBTc(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal)
{
var ikasleDatuak=bdIkasleOrria.getDataRange().getDisplayValues(); //datos alumnos
data.forEach(ilara=>            //PROCESA CADA ALUMNO SELECCIONADO 
  {
          const tempFile = docFile.makeCopy(tempFolder);
          const tempDocFile = DocumentApp.openById(tempFile.getId());
          const body = tempDocFile.getBody();
          const head = tempDocFile.getHeader();

          var ikasleW = ikasleDatuak.filter(fila=>fila[0]==ilara[1])
          
          if (ikasleW.length>0)
          {
            console.log("ikaslew ok")
            NOMBRE=ikasleW[0][0];
            GRUPO=ikasleW[0][5];
            PLANTEL=ikasleW[0][22];
            TURNO=ikasleW[0][8];
            var FING=ikasleW[0][10];
            ID=ikasleW[0][2];
          }
          else
          {
            console.log(ilara[1])
          }
        
        //*** ARMA EL HEADER */

        var CURSADAS= ilara[wal];
        var PROMEDIO= ilara[wal-1];
        var especialidad="CON ESPECIALIDAD EN COMUNICACIÓN DIGITAL"
        var CUATI="CUATRIMESTRE 1";
        var CUATII="CUATRIMESTRE 2";
        var CUATIII="CUATRIMESTRE 3";
        var CUATIV="CUATRIMESTRE 4";
        var CUATV="CUATRIMESTRE 5";
        var CUATVI="CUATRIMESTRE 6";
        var CUATVII="CUATRIMESTRE 7";
          head.replaceText("{OPCION}",opceduca);
          head.replaceText("{ID}",ID);
          head.replaceText("{NOMBRE}",NOMBRE);
          head.replaceText("{GRUPO}",GRUPO);
          head.replaceText("{PLANTEL}",PLANTEL);
          head.replaceText("{TURNO}",TURNO);
          head.replaceText("{PERIODO}",PERIODO);  
          head.replaceText("{PROMG}",PROMEDIO);
          head.replaceText("{FECEMI}",FECEMI);
          head.replaceText("{CURSADAS}",CURSADAS);  
          head.replaceText("{FING}",FING);
          head.replaceText("{especialidad}",especialidad);
        
      //***********************************************************************************

          body.replaceText("{CUATI}",CUATI);
          body.replaceText("{CUATII}",CUATII);
          body.replaceText("{CUATIII}",CUATIII);
          body.replaceText("{CUATIV}",CUATIV);
          body.replaceText("{CUATV}",CUATV);
          body.replaceText("{CUATVI}",CUATVI);
          body.replaceText("{CUATVII}",CUATVII);


          body.replaceText("{ID01}",ilara[2]);
          body.replaceText("{NAM01}",ilara[3]);
          body.replaceText("{CAL01}",ilara[4]);
          body.replaceText("{TIP01}",ilara[5]);
          body.replaceText("{FCURS1}",ilara[6]);

          body.replaceText("{ID02}", ilara[7]);
          body.replaceText("{NAM02}",ilara[8]);
          body.replaceText("{CAL02}",ilara[9]);
          body.replaceText("{TIP02}",ilara[10]);
          body.replaceText("{FCURS2}",ilara[11]);

          body.replaceText("{ID03}", ilara[12]);
          body.replaceText("{NAM03}",ilara[13]);
          body.replaceText("{CAL03}",ilara[14]);
          body.replaceText("{TIP03}",ilara[15]);
          body.replaceText("{FCURS3}",ilara[16]);

          body.replaceText("{ID04}", ilara[17]);
          body.replaceText("{NAM04}",ilara[18]);
          body.replaceText("{CAL04}",ilara[19]);
          body.replaceText("{TIP04}",ilara[20]);
          body.replaceText("{FCURS4}",ilara[21]);

          body.replaceText("{ID05}", ilara[22]);
          body.replaceText("{NAM05}",ilara[23]);
          body.replaceText("{CAL05}",ilara[24]);
          body.replaceText("{TIP05}",ilara[25]);
          body.replaceText("{FCURS5}",ilara[26]);

          body.replaceText("{ID06}", ilara[27]);
          body.replaceText("{NAM06}",ilara[28]);
          body.replaceText("{CAL06}",ilara[29]);
          body.replaceText("{TIP06}",ilara[30]);
          body.replaceText("{FCURS6}",ilara[31]);

          body.replaceText("{ID07}", ilara[32]);
          body.replaceText("{NAM07}",ilara[33]);
          body.replaceText("{CAL07}",ilara[34]);
          body.replaceText("{TIP07}",ilara[35]);
          body.replaceText("{FCURS7}",ilara[36]);

          body.replaceText("{ID08}", ilara[37]);
          body.replaceText("{NAM08}",ilara[38]);
          body.replaceText("{CAL08}",ilara[39]);
          body.replaceText("{TIP08}",ilara[40]);
          body.replaceText("{FCURS8}",ilara[41]);

          body.replaceText("{ID09}", ilara[42]);
          body.replaceText("{NAM09}",ilara[43]);
          body.replaceText("{CAL09}",ilara[44]);
          body.replaceText("{TIP09}",ilara[45]);
          body.replaceText("{FCURS9}",ilara[46]);

          body.replaceText("{ID10}", ilara[47]);
          body.replaceText("{NAM10}",ilara[48]);
          body.replaceText("{CAL10}",ilara[49]);
          body.replaceText("{TIP10}",ilara[50]);
          body.replaceText("{FCURS10}",ilara[51]);

          body.replaceText("{ID11}", ilara[52]);
          body.replaceText("{NAM11}",ilara[53]);
          body.replaceText("{CAL11}",ilara[54]);
          body.replaceText("{TIP11}",ilara[55]);
          body.replaceText("{FCURS11}",ilara[56]);

          body.replaceText("{ID12}", ilara[57]);
          body.replaceText("{NAM12}",ilara[58]);
          body.replaceText("{CAL12}",ilara[59]);
          body.replaceText("{TIP12}",ilara[60]);
          body.replaceText("{FCURS12}",ilara[61]);

          body.replaceText("{ID13}", ilara[62]);
          body.replaceText("{NAM13}",ilara[63]);
          body.replaceText("{CAL13}",ilara[64]);
          body.replaceText("{TIP13}",ilara[65]);
          body.replaceText("{FCURS13}",ilara[66]);

          body.replaceText("{ID14}", ilara[67]);
          body.replaceText("{NAM14}",ilara[68]);
          body.replaceText("{CAL14}",ilara[69]);
          body.replaceText("{TIP14}",ilara[70]);
          body.replaceText("{FCURS14}",ilara[71]);

          body.replaceText("{ID15}", ilara[72]);
          body.replaceText("{NAM15}",ilara[73]);
          body.replaceText("{CAL15}",ilara[74]);
          body.replaceText("{TIP15}",ilara[75]);
          body.replaceText("{FCURS15}",ilara[76]);

          body.replaceText("{ID16}", ilara[77]);
          body.replaceText("{NAM16}",ilara[78]);
          body.replaceText("{CAL16}",ilara[79]);
          body.replaceText("{TIP16}",ilara[80]);
          body.replaceText("{FCURS16}",ilara[81]);

          body.replaceText("{ID17}", ilara[82]);
          body.replaceText("{NAM17}",ilara[83]);
          body.replaceText("{CAL17}",ilara[84]);
          body.replaceText("{TIP17}",ilara[85]);
          body.replaceText("{FCURS17}",ilara[86]);

          body.replaceText("{ID18}", ilara[87]);
          body.replaceText("{NAM18}",ilara[88]);
          body.replaceText("{CAL18}",ilara[89]);
          body.replaceText("{TIP18}",ilara[90]);
          body.replaceText("{FCURS18}",ilara[91]);

          body.replaceText("{ID19}", ilara[92]);
          body.replaceText("{NAM19}",ilara[93]);
          body.replaceText("{CAL19}",ilara[94]);
          body.replaceText("{TIP19}",ilara[95]);
          body.replaceText("{FCURS19}",ilara[96]);

          body.replaceText("{ID20}", ilara[97]);
          body.replaceText("{NAM20}",ilara[98]);
          body.replaceText("{CAL20}",ilara[99]);
          body.replaceText("{TIP20}",ilara[100]);
          body.replaceText("{FCURS20}",ilara[101]);

          body.replaceText("{ID21}", ilara[102]);
          body.replaceText("{NAM21}",ilara[103]);
          body.replaceText("{CAL21}",ilara[104]);
          body.replaceText("{TIP21}",ilara[105]);
          body.replaceText("{FCURS21}",ilara[106]);
          
          body.replaceText("{ID22}", ilara[107]);
          body.replaceText("{NAM22}",ilara[108]);
          body.replaceText("{CAL22}",ilara[109]);
          body.replaceText("{TIP22}",ilara[110]);
          body.replaceText("{FCURS22}",ilara[111]);

          body.replaceText("{ID23}", ilara[112]);
          body.replaceText("{NAM23}",ilara[113]);
          body.replaceText("{CAL23}",ilara[114]);
          body.replaceText("{TIP23}",ilara[115]);
          body.replaceText("{FCURS23}",ilara[116]);

          body.replaceText("{ID24}",ilara[117]);
          body.replaceText("{NAM24}",ilara[118]);
          body.replaceText("{CAL24}",ilara[119]);
          body.replaceText("{TIP24}",ilara[120]);
          body.replaceText("{FCURS24}",ilara[121]);
          
          body.replaceText("{ID25}", ilara[122]);
          body.replaceText("{NAM25}",ilara[123]);
          body.replaceText("{CAL25}",ilara[124]);
          body.replaceText("{TIP25}",ilara[125]);
          body.replaceText("{FCURS25}",ilara[126]);

          body.replaceText("{ID26}", ilara[127]);
          body.replaceText("{NAM26}",ilara[128]);
          body.replaceText("{CAL26}",ilara[129]);
          body.replaceText("{TIP26}",ilara[130]);
          body.replaceText("{FCURS26}",ilara[131]);

          body.replaceText("{ID27}", ilara[132]);
          body.replaceText("{NAM27}",ilara[133]);
          body.replaceText("{CAL27}",ilara[134]);
          body.replaceText("{TIP27}",ilara[135]);
          body.replaceText("{FCURS27}",ilara[136]);

          body.replaceText("{ID28}", ilara[137]);
          body.replaceText("{NAM28}",ilara[138]);
          body.replaceText("{CAL28}",ilara[139]);
          body.replaceText("{TIP28}",ilara[140]);
          body.replaceText("{FCURS28}",ilara[141]);

          body.replaceText("{ID29}", ilara[142]);
          body.replaceText("{NAM29}",ilara[143]);
          body.replaceText("{CAL29}",ilara[144]);
          body.replaceText("{TIP29}",ilara[145]);
          body.replaceText("{FCURS29}",ilara[146]);

          body.replaceText("{ID30}", ilara[147]);
          body.replaceText("{NAM30}",ilara[148]);
          body.replaceText("{CAL30}",ilara[149]);
          body.replaceText("{TIP30}",ilara[150]);
          body.replaceText("{FCURS30}",ilara[151]);

          body.replaceText("{ID31}", ilara[152]);
          body.replaceText("{NAM31}",ilara[153]);
          body.replaceText("{CAL31}",ilara[154]);
          body.replaceText("{TIP31}",ilara[155]);
          body.replaceText("{FCURS31}",ilara[156]);

          body.replaceText("{ID32}", ilara[157]);
          body.replaceText("{NAM32}",ilara[158]);
          body.replaceText("{CAL32}",ilara[159]);
          body.replaceText("{TIP32}",ilara[160]);
          body.replaceText("{FCURS32}",ilara[161]);

          body.replaceText("{ID33}", ilara[162]);
          body.replaceText("{NAM33}",ilara[163]);
          body.replaceText("{CAL33}",ilara[164]);
          body.replaceText("{TIP33}",ilara[165]);
          body.replaceText("{FCURS33}",ilara[166]);

          body.replaceText("{ID34}", ilara[167]);
          body.replaceText("{NAM34}",ilara[168]);
          body.replaceText("{CAL34}",ilara[169]);
          body.replaceText("{TIP34}",ilara[170]);
          body.replaceText("{FCURS34}",ilara[171]);

          body.replaceText("{ID35}", ilara[172]);
          body.replaceText("{NAM35}",ilara[173]);
          body.replaceText("{CAL35}",ilara[174]);
          body.replaceText("{TIP35}",ilara[175]);
          body.replaceText("{FCURS35}",ilara[176]);

          body.replaceText("{ID36}", ilara[177]);
          body.replaceText("{NAM36}",ilara[178]);
          body.replaceText("{CAL36}",ilara[179]);
          body.replaceText("{TIP36}",ilara[180]);
          body.replaceText("{FCURS36}",ilara[181]);

          body.replaceText("{ID37}", ilara[182]);
          body.replaceText("{NAM37}",ilara[183]);
          body.replaceText("{CAL37}",ilara[184]);
          body.replaceText("{TIP37}",ilara[185]);
          body.replaceText("{FCURS37}",ilara[186]);

          body.replaceText("{ID38}", ilara[187]);
          body.replaceText("{NAM38}",ilara[188]);
          body.replaceText("{CAL38}",ilara[189]);
          body.replaceText("{TIP38}",ilara[190]);
          body.replaceText("{FCURS38}",ilara[191]);

          body.replaceText("{ID39}", ilara[192]);
          body.replaceText("{NAM39}",ilara[193]);
          body.replaceText("{CAL39}",ilara[194]);
          body.replaceText("{TIP39}",ilara[195]);
          body.replaceText("{FCURS39}",ilara[196]);

          body.replaceText("{ID40}", ilara[197]);
          body.replaceText("{NAM40}",ilara[198]);
          body.replaceText("{CAL40}",ilara[199]);
          body.replaceText("{TIP40}",ilara[200]);
          body.replaceText("{FCURS40}",ilara[201]);

          body.replaceText("{ID41}", ilara[202]);
          body.replaceText("{NAM41}",ilara[203]);
          body.replaceText("{CAL41}",ilara[204]);
          body.replaceText("{TIP41}",ilara[205]);
          body.replaceText("{FCURS41}",ilara[206]);

          body.replaceText("{ID42}", ilara[207]);
          body.replaceText("{NAM42}",ilara[208]);
          body.replaceText("{CAL42}",ilara[209]);
          body.replaceText("{TIP42}",ilara[210]);
          body.replaceText("{FCURS42}",ilara[211]);

          body.replaceText("{ID43}", ilara[212]);
          body.replaceText("{NAM43}",ilara[213]);
          body.replaceText("{CAL43}",ilara[214]);
          body.replaceText("{TIP43}",ilara[215]);
          body.replaceText("{FCURS43}",ilara[216]);

          body.replaceText("{ID44}", ilara[217]);
          body.replaceText("{NAM44}",ilara[218]);
          body.replaceText("{CAL44}",ilara[219]);
          body.replaceText("{TIP44}",ilara[220]);
          body.replaceText("{FCURS44}",ilara[221]);

          body.replaceText("{ID45}", ilara[222]);
          body.replaceText("{NAM45}",ilara[223]);
          body.replaceText("{CAL45}",ilara[224]);
          body.replaceText("{TIP45}",ilara[225]);
          body.replaceText("{FCURS45}",ilara[226]);

          body.replaceText("{ID46}", ilara[227]);
          body.replaceText("{NAM46}",ilara[228]);
          body.replaceText("{CAL46}",ilara[229]);
          body.replaceText("{TIP46}",ilara[230]);
          body.replaceText("{FCURS46}",ilara[231]);

          body.replaceText("{ID47}", ilara[232]);
          body.replaceText("{NAM47}",ilara[233]);
          body.replaceText("{CAL47}",ilara[234]);
          body.replaceText("{TIP47}",ilara[235]);
          body.replaceText("{FCURS47}",ilara[236]);


          body.replaceText("{ID48}", ilara[237]);
          body.replaceText("{NAM48}",ilara[238]);
          body.replaceText("{CAL48}",ilara[239]);
          body.replaceText("{TIP48}",ilara[240]);
          body.replaceText("{FCURS48}",ilara[241]);

          body.replaceText("{ID49}", ilara[242]);
          body.replaceText("{NAM49}",ilara[243]);
          body.replaceText("{CAL49}",ilara[244]);
          body.replaceText("{TIP49}",ilara[245]);
          body.replaceText("{FCURS49}",ilara[246]);

          body.replaceText("{ID50}", ilara[247]);
          body.replaceText("{NAM50}",ilara[248]);
          body.replaceText("{CAL50}",ilara[249]);
          body.replaceText("{TIP50}",ilara[250]);
          body.replaceText("{FCURS50}",ilara[251]);

          body.replaceText("{ID51}", ilara[252]);
          body.replaceText("{NAM51}",ilara[253]);
          body.replaceText("{CAL51}",ilara[254]);
          body.replaceText("{TIP51}",ilara[255]);
          body.replaceText("{FCURS51}",ilara[256]);

          body.replaceText("{ID52}", ilara[257]);
          body.replaceText("{NAM52}",ilara[258]);
          body.replaceText("{CAL52}",ilara[259]);
          body.replaceText("{TIP52}",ilara[260]);
          body.replaceText("{FCURS52}",ilara[261]);

          body.replaceText("{ID53}", ilara[262]);
          body.replaceText("{NAM53}",ilara[263]);
          body.replaceText("{CAL53}",ilara[264]);
          body.replaceText("{TIP53}",ilara[265]);
          body.replaceText("{FCURS53}",ilara[266]);

          body.replaceText("{ID54}", ilara[267]);
          body.replaceText("{NAM54}",ilara[268]);
          body.replaceText("{CAL54}",ilara[269]);
          body.replaceText("{TIP54}",ilara[270]);
          body.replaceText("{FCURS54}",ilara[271]);

          body.replaceText("{ID55}", ilara[272]);
          body.replaceText("{NAM55}",ilara[273]);
          body.replaceText("{CAL55}",ilara[274]);
          body.replaceText("{TIP55}",ilara[275]);
          body.replaceText("{FCURS55}",ilara[276]);

          body.replaceText("{ID56}", ilara[277]);
          body.replaceText("{NAM56}",ilara[278]);
          body.replaceText("{CAL56}",ilara[279]);
          body.replaceText("{TIP56}",ilara[280]);
          body.replaceText("{FCURS56}",ilara[281]);

          body.replaceText("{ID57}", ilara[282]);
          body.replaceText("{NAM57}",ilara[283]);
          body.replaceText("{CAL57}",ilara[284]);
          body.replaceText("{TIP57}",ilara[285]);
          body.replaceText("{FCURS57}",ilara[286]);

          body.replaceText("{ID58}", ilara[287]);
          body.replaceText("{NAM58}",ilara[288]);
          body.replaceText("{CAL58}",ilara[289]);
          body.replaceText("{TIP58}",ilara[290]);
          body.replaceText("{FCURS58}",ilara[291]);

          body.replaceText("{ID59}", ilara[292]);
          body.replaceText("{NAM59}",ilara[293]);
          body.replaceText("{CAL59}",ilara[294]);
          body.replaceText("{TIP59}",ilara[295]);
          body.replaceText("{FCURS59}",ilara[296]);

          body.replaceText("{ID60}", ilara[297]);
          body.replaceText("{NAM60}",ilara[298]);
          body.replaceText("{CAL60}",ilara[299]);
          body.replaceText("{TIP60}",ilara[300]);
          body.replaceText("{FCURS60}",ilara[301]);

          body.replaceText("{ID61}", ilara[302]);
          body.replaceText("{NAM61}",ilara[303]);
          body.replaceText("{CAL61}",ilara[304]);
          body.replaceText("{TIP61}",ilara[305]);
          body.replaceText("{FCURS61}",ilara[306]);

          body.replaceText("{ID62}", ilara[307]);
          body.replaceText("{NAM62}",ilara[308]);
          body.replaceText("{CAL62}",ilara[309]);
          body.replaceText("{TIP62}",ilara[310]);
          body.replaceText("{FCURS62}",ilara[311]);

          body.replaceText("{ID63}", ilara[312]);
          body.replaceText("{NAM63}",ilara[313]);
          body.replaceText("{CAL63}",ilara[314]);
          body.replaceText("{TIP63}",ilara[315]);
          body.replaceText("{FCURS63}",ilara[316]);

          body.replaceText("{ID64}", ilara[317]);
          body.replaceText("{NAM64}",ilara[318]);
          body.replaceText("{CAL64}",ilara[319]);
          body.replaceText("{TIP64}",ilara[320]);
          body.replaceText("{FCURS64}",ilara[321]);

          body.replaceText("{ID65}", ilara[322]);
          body.replaceText("{NAM65}",ilara[323]);
          body.replaceText("{CAL65}",ilara[324]);
          body.replaceText("{TIP65}",ilara[325]);
          body.replaceText("{FCURS65}",ilara[326]);

          body.replaceText("{ID66}", ilara[327]);
          body.replaceText("{NAM66}",ilara[328]);
          body.replaceText("{CAL66}",ilara[329]);
          body.replaceText("{TIP66}",ilara[330]);
          body.replaceText("{FCURS66}",ilara[331]);
          
          finalizaImpresion(NOMBRE,GRUPO,ID,opceduca,PLANTEL,PERIODO,tempDocFile,pdfFolder,tempFolder,tempFile)
    })
}

//*********************************************** */

//*********************************************** */
//*          PLANTILLA DE LICENCIATURA ADMINISTRACIÓN

function createPDFLAD(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal)
{
var ikasleDatuak=bdIkasleOrria.getDataRange().getDisplayValues(); //datos alumnos
data.forEach(ilara=>            //PROCESA CADA ALUMNO SELECCIONADO 
  {
          const tempFile = docFile.makeCopy(tempFolder);
          const tempDocFile = DocumentApp.openById(tempFile.getId());
          const body = tempDocFile.getBody();
          const head = tempDocFile.getHeader();

          var ikasleW = ikasleDatuak.filter(fila=>fila[0]==ilara[1])
          
          if (ikasleW.length>0)
          {
            console.log("ikaslew ok")
            NOMBRE=ikasleW[0][0];
            GRUPO=ikasleW[0][5];
            PLANTEL=ikasleW[0][22];
            TURNO=ikasleW[0][8];
            var FING=ikasleW[0][10];
            ID=ikasleW[0][2];
          }
          else
          {
            console.log(ilara[1])
          }
        
        //*** ARMA EL HEADER */

        var CURSADAS= ilara[wal];
        var PROMEDIO= ilara[wal-1];
        var especialidad=""
        var CUATI="CUATRIMESTRE 1";
        var CUATII="CUATRIMESTRE 2";
        var CUATIII="CUATRIMESTRE 3";
        var CUATIV="CUATRIMESTRE 4";
        var CUATV="CUATRIMESTRE 5";
        var CUATVI="CUATRIMESTRE 6";
        var CUATVII="CUATRIMESTRE 7";
        var CUATVIII="CUATRIMESTRE 8";
        var CUATIX="CUATRIMESTRE 9";
          head.replaceText("{OPCION}",opceduca);
          head.replaceText("{ID}",ID);
          head.replaceText("{NOMBRE}",NOMBRE);
          head.replaceText("{GRUPO}",GRUPO);
          head.replaceText("{PLANTEL}",PLANTEL);
          head.replaceText("{TURNO}",TURNO);
          head.replaceText("{PERIODO}",PERIODO);  
          head.replaceText("{PROMG}",PROMEDIO);
          head.replaceText("{FECEMI}",FECEMI);
          head.replaceText("{CURSADAS}",CURSADAS);  
          head.replaceText("{FING}",FING);
          head.replaceText("{especialidad}",especialidad);
        
      //***********************************************************************************

          body.replaceText("{CUATI}",CUATI);
          body.replaceText("{CUATII}",CUATII);
          body.replaceText("{CUATIII}",CUATIII);
          body.replaceText("{CUATIV}",CUATIV);
          body.replaceText("{CUATV}",CUATV);
          body.replaceText("{CUATVI}",CUATVI);
          body.replaceText("{CUATVII}",CUATVII);
          body.replaceText("{CUATVIII}",CUATVIII);
          body.replaceText("{CUATIX}",CUATIX);


          body.replaceText("{ID01}",ilara[2]);
          body.replaceText("{NAM01}",ilara[3]);
          body.replaceText("{CAL01}",ilara[4]);
          body.replaceText("{TIP01}",ilara[5]);
          body.replaceText("{FCURS1}",ilara[6]);

          body.replaceText("{ID02}", ilara[7]);
          body.replaceText("{NAM02}",ilara[8]);
          body.replaceText("{CAL02}",ilara[9]);
          body.replaceText("{TIP02}",ilara[10]);
          body.replaceText("{FCURS2}",ilara[11]);

          body.replaceText("{ID03}", ilara[12]);
          body.replaceText("{NAM03}",ilara[13]);
          body.replaceText("{CAL03}",ilara[14]);
          body.replaceText("{TIP03}",ilara[15]);
          body.replaceText("{FCURS3}",ilara[16]);

          body.replaceText("{ID04}", ilara[17]);
          body.replaceText("{NAM04}",ilara[18]);
          body.replaceText("{CAL04}",ilara[19]);
          body.replaceText("{TIP04}",ilara[20]);
          body.replaceText("{FCURS4}",ilara[21]);

          body.replaceText("{ID05}", ilara[22]);
          body.replaceText("{NAM05}",ilara[23]);
          body.replaceText("{CAL05}",ilara[24]);
          body.replaceText("{TIP05}",ilara[25]);
          body.replaceText("{FCURS5}",ilara[26]);

          body.replaceText("{ID06}", ilara[27]);
          body.replaceText("{NAM06}",ilara[28]);
          body.replaceText("{CAL06}",ilara[29]);
          body.replaceText("{TIP06}",ilara[30]);
          body.replaceText("{FCURS6}",ilara[31]);

          body.replaceText("{ID07}", ilara[32]);
          body.replaceText("{NAM07}",ilara[33]);
          body.replaceText("{CAL07}",ilara[34]);
          body.replaceText("{TIP07}",ilara[35]);
          body.replaceText("{FCURS7}",ilara[36]);

          body.replaceText("{ID08}", ilara[37]);
          body.replaceText("{NAM08}",ilara[38]);
          body.replaceText("{CAL08}",ilara[39]);
          body.replaceText("{TIP08}",ilara[40]);
          body.replaceText("{FCURS8}",ilara[41]);

          body.replaceText("{ID09}", ilara[42]);
          body.replaceText("{NAM09}",ilara[43]);
          body.replaceText("{CAL09}",ilara[44]);
          body.replaceText("{TIP09}",ilara[45]);
          body.replaceText("{FCURS9}",ilara[46]);

          body.replaceText("{ID10}", ilara[47]);
          body.replaceText("{NAM10}",ilara[48]);
          body.replaceText("{CAL10}",ilara[49]);
          body.replaceText("{TIP10}",ilara[50]);
          body.replaceText("{FCURS10}",ilara[51]);

          body.replaceText("{ID11}", ilara[52]);
          body.replaceText("{NAM11}",ilara[53]);
          body.replaceText("{CAL11}",ilara[54]);
          body.replaceText("{TIP11}",ilara[55]);
          body.replaceText("{FCURS11}",ilara[56]);

          body.replaceText("{ID12}", ilara[57]);
          body.replaceText("{NAM12}",ilara[58]);
          body.replaceText("{CAL12}",ilara[59]);
          body.replaceText("{TIP12}",ilara[60]);
          body.replaceText("{FCURS12}",ilara[61]);

          body.replaceText("{ID13}", ilara[62]);
          body.replaceText("{NAM13}",ilara[63]);
          body.replaceText("{CAL13}",ilara[64]);
          body.replaceText("{TIP13}",ilara[65]);
          body.replaceText("{FCURS13}",ilara[66]);

          body.replaceText("{ID14}", ilara[67]);
          body.replaceText("{NAM14}",ilara[68]);
          body.replaceText("{CAL14}",ilara[69]);
          body.replaceText("{TIP14}",ilara[70]);
          body.replaceText("{FCURS14}",ilara[71]);

          body.replaceText("{ID15}", ilara[72]);
          body.replaceText("{NAM15}",ilara[73]);
          body.replaceText("{CAL15}",ilara[74]);
          body.replaceText("{TIP15}",ilara[75]);
          body.replaceText("{FCURS15}",ilara[76]);

          body.replaceText("{ID16}", ilara[77]);
          body.replaceText("{NAM16}",ilara[78]);
          body.replaceText("{CAL16}",ilara[79]);
          body.replaceText("{TIP16}",ilara[80]);
          body.replaceText("{FCURS16}",ilara[81]);

          body.replaceText("{ID17}", ilara[82]);
          body.replaceText("{NAM17}",ilara[83]);
          body.replaceText("{CAL17}",ilara[84]);
          body.replaceText("{TIP17}",ilara[85]);
          body.replaceText("{FCURS17}",ilara[86]);

          body.replaceText("{ID18}", ilara[87]);
          body.replaceText("{NAM18}",ilara[88]);
          body.replaceText("{CAL18}",ilara[89]);
          body.replaceText("{TIP18}",ilara[90]);
          body.replaceText("{FCURS18}",ilara[91]);

          body.replaceText("{ID19}", ilara[92]);
          body.replaceText("{NAM19}",ilara[93]);
          body.replaceText("{CAL19}",ilara[94]);
          body.replaceText("{TIP19}",ilara[95]);
          body.replaceText("{FCURS19}",ilara[96]);

          body.replaceText("{ID20}", ilara[97]);
          body.replaceText("{NAM20}",ilara[98]);
          body.replaceText("{CAL20}",ilara[99]);
          body.replaceText("{TIP20}",ilara[100]);
          body.replaceText("{FCURS20}",ilara[101]);

          body.replaceText("{ID21}", ilara[102]);
          body.replaceText("{NAM21}",ilara[103]);
          body.replaceText("{CAL21}",ilara[104]);
          body.replaceText("{TIP21}",ilara[105]);
          body.replaceText("{FCURS21}",ilara[106]);
          
          body.replaceText("{ID22}", ilara[107]);
          body.replaceText("{NAM22}",ilara[108]);
          body.replaceText("{CAL22}",ilara[109]);
          body.replaceText("{TIP22}",ilara[110]);
          body.replaceText("{FCURS22}",ilara[111]);

          body.replaceText("{ID23}", ilara[112]);
          body.replaceText("{NAM23}",ilara[113]);
          body.replaceText("{CAL23}",ilara[114]);
          body.replaceText("{TIP23}",ilara[115]);
          body.replaceText("{FCURS23}",ilara[116]);

          body.replaceText("{ID24}",ilara[117]);
          body.replaceText("{NAM24}",ilara[118]);
          body.replaceText("{CAL24}",ilara[119]);
          body.replaceText("{TIP24}",ilara[120]);
          body.replaceText("{FCURS24}",ilara[121]);
          
          body.replaceText("{ID25}", ilara[122]);
          body.replaceText("{NAM25}",ilara[123]);
          body.replaceText("{CAL25}",ilara[124]);
          body.replaceText("{TIP25}",ilara[125]);
          body.replaceText("{FCURS25}",ilara[126]);

          body.replaceText("{ID26}", ilara[127]);
          body.replaceText("{NAM26}",ilara[128]);
          body.replaceText("{CAL26}",ilara[129]);
          body.replaceText("{TIP26}",ilara[130]);
          body.replaceText("{FCURS26}",ilara[131]);

          body.replaceText("{ID27}", ilara[132]);
          body.replaceText("{NAM27}",ilara[133]);
          body.replaceText("{CAL27}",ilara[134]);
          body.replaceText("{TIP27}",ilara[135]);
          body.replaceText("{FCURS27}",ilara[136]);

          body.replaceText("{ID28}", ilara[137]);
          body.replaceText("{NAM28}",ilara[138]);
          body.replaceText("{CAL28}",ilara[139]);
          body.replaceText("{TIP28}",ilara[140]);
          body.replaceText("{FCURS28}",ilara[141]);

          body.replaceText("{ID29}", ilara[142]);
          body.replaceText("{NAM29}",ilara[143]);
          body.replaceText("{CAL29}",ilara[144]);
          body.replaceText("{TIP29}",ilara[145]);
          body.replaceText("{FCURS29}",ilara[146]);

          body.replaceText("{ID30}", ilara[147]);
          body.replaceText("{NAM30}",ilara[148]);
          body.replaceText("{CAL30}",ilara[149]);
          body.replaceText("{TIP30}",ilara[150]);
          body.replaceText("{FCURS30}",ilara[151]);

          body.replaceText("{ID31}", ilara[152]);
          body.replaceText("{NAM31}",ilara[153]);
          body.replaceText("{CAL31}",ilara[154]);
          body.replaceText("{TIP31}",ilara[155]);
          body.replaceText("{FCURS31}",ilara[156]);

          body.replaceText("{ID32}", ilara[157]);
          body.replaceText("{NAM32}",ilara[158]);
          body.replaceText("{CAL32}",ilara[159]);
          body.replaceText("{TIP32}",ilara[160]);
          body.replaceText("{FCURS32}",ilara[161]);

          body.replaceText("{ID33}", ilara[162]);
          body.replaceText("{NAM33}",ilara[163]);
          body.replaceText("{CAL33}",ilara[164]);
          body.replaceText("{TIP33}",ilara[165]);
          body.replaceText("{FCURS33}",ilara[166]);

          body.replaceText("{ID34}", ilara[167]);
          body.replaceText("{NAM34}",ilara[168]);
          body.replaceText("{CAL34}",ilara[169]);
          body.replaceText("{TIP34}",ilara[170]);
          body.replaceText("{FCURS34}",ilara[171]);

          body.replaceText("{ID35}", ilara[172]);
          body.replaceText("{NAM35}",ilara[173]);
          body.replaceText("{CAL35}",ilara[174]);
          body.replaceText("{TIP35}",ilara[175]);
          body.replaceText("{FCURS35}",ilara[176]);

          body.replaceText("{ID36}", ilara[177]);
          body.replaceText("{NAM36}",ilara[178]);
          body.replaceText("{CAL36}",ilara[179]);
          body.replaceText("{TIP36}",ilara[180]);
          body.replaceText("{FCURS36}",ilara[181]);

          body.replaceText("{ID37}", ilara[182]);
          body.replaceText("{NAM37}",ilara[183]);
          body.replaceText("{CAL37}",ilara[184]);
          body.replaceText("{TIP37}",ilara[185]);
          body.replaceText("{FCURS37}",ilara[186]);

          body.replaceText("{ID38}", ilara[187]);
          body.replaceText("{NAM38}",ilara[188]);
          body.replaceText("{CAL38}",ilara[189]);
          body.replaceText("{TIP38}",ilara[190]);
          body.replaceText("{FCURS38}",ilara[191]);

          body.replaceText("{ID39}", ilara[192]);
          body.replaceText("{NAM39}",ilara[193]);
          body.replaceText("{CAL39}",ilara[194]);
          body.replaceText("{TIP39}",ilara[195]);
          body.replaceText("{FCURS39}",ilara[196]);

          body.replaceText("{ID40}", ilara[197]);
          body.replaceText("{NAM40}",ilara[198]);
          body.replaceText("{CAL40}",ilara[199]);
          body.replaceText("{TIP40}",ilara[200]);
          body.replaceText("{FCURS40}",ilara[201]);

          body.replaceText("{ID41}", ilara[202]);
          body.replaceText("{NAM41}",ilara[203]);
          body.replaceText("{CAL41}",ilara[204]);
          body.replaceText("{TIP41}",ilara[205]);
          body.replaceText("{FCURS41}",ilara[206]);

          body.replaceText("{ID42}", ilara[207]);
          body.replaceText("{NAM42}",ilara[208]);
          body.replaceText("{CAL42}",ilara[209]);
          body.replaceText("{TIP42}",ilara[210]);
          body.replaceText("{FCURS42}",ilara[211]);

          body.replaceText("{ID43}", ilara[212]);
          body.replaceText("{NAM43}",ilara[213]);
          body.replaceText("{CAL43}",ilara[214]);
          body.replaceText("{TIP43}",ilara[215]);
          body.replaceText("{FCURS43}",ilara[216]);

          body.replaceText("{ID44}", ilara[217]);
          body.replaceText("{NAM44}",ilara[218]);
          body.replaceText("{CAL44}",ilara[219]);
          body.replaceText("{TIP44}",ilara[220]);
          body.replaceText("{FCURS44}",ilara[221]);

          body.replaceText("{ID45}", ilara[222]);
          body.replaceText("{NAM45}",ilara[223]);
          body.replaceText("{CAL45}",ilara[224]);
          body.replaceText("{TIP45}",ilara[225]);
          body.replaceText("{FCURS45}",ilara[226]);

          body.replaceText("{ID46}", ilara[227]);
          body.replaceText("{NAM46}",ilara[228]);
          body.replaceText("{CAL46}",ilara[229]);
          body.replaceText("{TIP46}",ilara[230]);
          body.replaceText("{FCURS46}",ilara[231]);

          body.replaceText("{ID47}", ilara[232]);
          body.replaceText("{NAM47}",ilara[233]);
          body.replaceText("{CAL47}",ilara[234]);
          body.replaceText("{TIP47}",ilara[235]);
          body.replaceText("{FCURS47}",ilara[236]);


          body.replaceText("{ID48}", ilara[237]);
          body.replaceText("{NAM48}",ilara[238]);
          body.replaceText("{CAL48}",ilara[239]);
          body.replaceText("{TIP48}",ilara[240]);
          body.replaceText("{FCURS48}",ilara[241]);

          body.replaceText("{ID49}", ilara[242]);
          body.replaceText("{NAM49}",ilara[243]);
          body.replaceText("{CAL49}",ilara[244]);
          body.replaceText("{TIP49}",ilara[245]);
          body.replaceText("{FCURS49}",ilara[246]);

          body.replaceText("{ID50}", ilara[247]);
          body.replaceText("{NAM50}",ilara[248]);
          body.replaceText("{CAL50}",ilara[249]);
          body.replaceText("{TIP50}",ilara[250]);
          body.replaceText("{FCURS50}",ilara[251]);

          body.replaceText("{ID51}", ilara[252]);
          body.replaceText("{NAM51}",ilara[253]);
          body.replaceText("{CAL51}",ilara[254]);
          body.replaceText("{TIP51}",ilara[255]);
          body.replaceText("{FCURS51}",ilara[256]);

          body.replaceText("{ID52}", ilara[257]);
          body.replaceText("{NAM52}",ilara[258]);
          body.replaceText("{CAL52}",ilara[259]);
          body.replaceText("{TIP52}",ilara[260]);
          body.replaceText("{FCURS52}",ilara[261]);

          body.replaceText("{ID53}", ilara[262]);
          body.replaceText("{NAM53}",ilara[263]);
          body.replaceText("{CAL53}",ilara[264]);
          body.replaceText("{TIP53}",ilara[265]);
          body.replaceText("{FCURS53}",ilara[266]);

          body.replaceText("{ID54}", ilara[267]);
          body.replaceText("{NAM54}",ilara[268]);
          body.replaceText("{CAL54}",ilara[269]);
          body.replaceText("{TIP54}",ilara[270]);
          body.replaceText("{FCURS54}",ilara[271]);

          body.replaceText("{ID55}", ilara[272]);
          body.replaceText("{NAM55}",ilara[273]);
          body.replaceText("{CAL55}",ilara[274]);
          body.replaceText("{TIP55}",ilara[275]);
          body.replaceText("{FCURS55}",ilara[276]);

          body.replaceText("{ID56}", ilara[277]);
          body.replaceText("{NAM56}",ilara[278]);
          body.replaceText("{CAL56}",ilara[279]);
          body.replaceText("{TIP56}",ilara[280]);
          body.replaceText("{FCURS56}",ilara[281]);

          body.replaceText("{ID57}", ilara[282]);
          body.replaceText("{NAM57}",ilara[283]);
          body.replaceText("{CAL57}",ilara[284]);
          body.replaceText("{TIP57}",ilara[285]);
          body.replaceText("{FCURS57}",ilara[286]);

          body.replaceText("{ID58}", ilara[287]);
          body.replaceText("{NAM58}",ilara[288]);
          body.replaceText("{CAL58}",ilara[289]);
          body.replaceText("{TIP58}",ilara[290]);
          body.replaceText("{FCURS58}",ilara[291]);

          body.replaceText("{ID59}", ilara[292]);
          body.replaceText("{NAM59}",ilara[293]);
          body.replaceText("{CAL59}",ilara[294]);
          body.replaceText("{TIP59}",ilara[295]);
          body.replaceText("{FCURS59}",ilara[296]);

          body.replaceText("{ID60}", ilara[297]);
          body.replaceText("{NAM60}",ilara[298]);
          body.replaceText("{CAL60}",ilara[299]);
          body.replaceText("{TIP60}",ilara[300]);
          body.replaceText("{FCURS60}",ilara[301]);

          body.replaceText("{ID61}", ilara[302]);
          body.replaceText("{NAM61}",ilara[303]);
          body.replaceText("{CAL61}",ilara[304]);
          body.replaceText("{TIP61}",ilara[305]);
          body.replaceText("{FCURS61}",ilara[306]);

          body.replaceText("{ID62}", ilara[307]);
          body.replaceText("{NAM62}",ilara[308]);
          body.replaceText("{CAL62}",ilara[309]);
          body.replaceText("{TIP62}",ilara[310]);
          body.replaceText("{FCURS62}",ilara[311]);

          body.replaceText("{ID63}", ilara[312]);
          body.replaceText("{NAM63}",ilara[313]);
          body.replaceText("{CAL63}",ilara[314]);
          body.replaceText("{TIP63}",ilara[315]);
          body.replaceText("{FCURS63}",ilara[316]);

          body.replaceText("{ID64}", ilara[317]);
          body.replaceText("{NAM64}",ilara[318]);
          body.replaceText("{CAL64}",ilara[319]);
          body.replaceText("{TIP64}",ilara[320]);
          body.replaceText("{FCURS64}",ilara[321]);

          body.replaceText("{ID65}", ilara[322]);
          body.replaceText("{NAM65}",ilara[323]);
          body.replaceText("{CAL65}",ilara[324]);
          body.replaceText("{TIP65}",ilara[325]);
          body.replaceText("{FCURS65}",ilara[326]);

          body.replaceText("{ID66}", ilara[327]);
          body.replaceText("{NAM66}",ilara[328]);
          body.replaceText("{CAL66}",ilara[329]);
          body.replaceText("{TIP66}",ilara[330]);
          body.replaceText("{FCURS66}",ilara[331]);

        finalizaImpresion(NOMBRE,GRUPO,ID,opceduca,PLANTEL,PERIODO,tempDocFile,pdfFolder,tempFolder,tempFile)
        })
}
    
function finalizaImpresion(NOMBRE,GRUPO,ID,opceduca,PLANTEL,PERIODO,tempDocFile,pdfFolder,tempFolder,tempFile)
{
          var opcEduDatuak =opcEduOrria.getDataRange().getDisplayValues();
          var opcEduCveData=opcEduDatuak.filter(ilara=> ilara[0]==opceduca);
          if (opcEduCveData.length>0)
          {
            var opcEduCve=opcEduCveData[0][2];
          }
          else
          {
            var opcEduCve=NaN;
          }

          var pdfName=NOMBRE + "." +GRUPO+opcEduCve+ "." +PERIODO
          tempDocFile.saveAndClose();


          const pdfContentBlob = tempDocFile.getAs(MimeType.PDF);
              
          pdfFile = pdfFolder.createFile(pdfContentBlob).setName(pdfName);
              
          const pdfURL = pdfFile.getUrl();
          
          tempFolder.removeFile(tempFile);
          

      regUrl.push([NOMBRE,pdfURL,GRUPO])

      var IDPER=""
      var CICLO=""
      var SEMESTRE=""
      var diaCre = new Date();
      const dataC = [NOMBRE,ID,opceduca,GRUPO,SEMESTRE,PLANTEL,PERIODO,CICLO,pdfURL,diaCre];
      //catBoleta.appendRow(dataC);
      histCal.push(dataC);
}

//*********************************************** */



//*********************************************** */
//*          PLANTILLA DE DISEÑO GRAFICO

function createPDFBTD(docFile,tempFolder,pdfFolder,FECEMI,data,opceduca,PERIODO,wal)
{
var ikasleDatuak=bdIkasleOrria.getDataRange().getDisplayValues(); //datos alumnos
data.forEach(ilara=>            //PROCESA CADA ALUMNO SELECCIONADO 
  {
          const tempFile = docFile.makeCopy(tempFolder);
          const tempDocFile = DocumentApp.openById(tempFile.getId());
          const body = tempDocFile.getBody();
          const head = tempDocFile.getHeader();

          var ikasleW = ikasleDatuak.filter(fila=>fila[0]==ilara[1])
          
          if (ikasleW.length>0)
          {
            NOMBRE=ikasleW[0][0];
            GRUPO=ikasleW[0][5];
            PLANTEL=ikasleW[0][22];
            TURNO=ikasleW[0][8];
            var FING=ikasleW[0][10];
            ID=ikasleW[0][2];
          }
          else
          {
            console.log(ilara[1])
          }
        
        //*** ARMA EL HEADER */

        var CURSADAS= ilara[wal];
        var PROMEDIO= ilara[wal-1];
        var especialidad="CON ESPECIALIDAD EN DISEÑO GRÁFICO"
        var CUATI="SEMESTRE 1";
        var CUATII="SEMESTRE 2";
        var CUATIII="SEMESTRE 3";
        var CUATIV="SEMESTRE 4";
        var CUATV="SEMESTRE 5";
        var CUATVI="SEMESTRE 6";
        
          head.replaceText("{OPCION}",opceduca);
          head.replaceText("{ID}",ID);
          head.replaceText("{NOMBRE}",NOMBRE);
          head.replaceText("{GRUPO}",GRUPO);
          head.replaceText("{PLANTEL}",PLANTEL);
          head.replaceText("{TURNO}",TURNO);
          head.replaceText("{PERIODO}",PERIODO);  
          head.replaceText("{PROMG}",PROMEDIO);
          head.replaceText("{FECEMI}",FECEMI);
          head.replaceText("{CURSADAS}",CURSADAS);  
          head.replaceText("{FING}",FING);
          head.replaceText("{especialidad}",especialidad);
        
      //***********************************************************************************

          body.replaceText("{CUATI}",CUATI);
          body.replaceText("{CUATII}",CUATII);
          body.replaceText("{CUATIII}",CUATIII);
          body.replaceText("{CUATIV}",CUATIV);
          body.replaceText("{CUATV}",CUATV);
          body.replaceText("{CUATVI}",CUATVI);


          body.replaceText("{ID01}",ilara[2]);
          body.replaceText("{NAM01}",ilara[3]);
          body.replaceText("{CAL01}",ilara[4]);
          body.replaceText("{TIP01}",ilara[5]);
          body.replaceText("{FCURS1}",ilara[6]);

          body.replaceText("{ID02}", ilara[7]);
          body.replaceText("{NAM02}",ilara[8]);
          body.replaceText("{CAL02}",ilara[9]);
          body.replaceText("{TIP02}",ilara[10]);
          body.replaceText("{FCURS2}",ilara[11]);

          body.replaceText("{ID03}", ilara[12]);
          body.replaceText("{NAM03}",ilara[13]);
          body.replaceText("{CAL03}",ilara[14]);
          body.replaceText("{TIP03}",ilara[15]);
          body.replaceText("{FCURS3}",ilara[16]);

          body.replaceText("{ID04}", ilara[17]);
          body.replaceText("{NAM04}",ilara[18]);
          body.replaceText("{CAL04}",ilara[19]);
          body.replaceText("{TIP04}",ilara[20]);
          body.replaceText("{FCURS4}",ilara[21]);

          body.replaceText("{ID05}", ilara[22]);
          body.replaceText("{NAM05}",ilara[23]);
          body.replaceText("{CAL05}",ilara[24]);
          body.replaceText("{TIP05}",ilara[25]);
          body.replaceText("{FCURS5}",ilara[26]);

          body.replaceText("{ID06}", ilara[27]);
          body.replaceText("{NAM06}",ilara[28]);
          body.replaceText("{CAL06}",ilara[29]);
          body.replaceText("{TIP06}",ilara[30]);
          body.replaceText("{FCURS6}",ilara[31]);

          body.replaceText("{ID07}", ilara[32]);
          body.replaceText("{NAM07}",ilara[33]);
          body.replaceText("{CAL07}",ilara[34]);
          body.replaceText("{TIP07}",ilara[35]);
          body.replaceText("{FCURS7}",ilara[36]);

          body.replaceText("{ID08}", ilara[37]);
          body.replaceText("{NAM08}",ilara[38]);
          body.replaceText("{CAL08}",ilara[39]);
          body.replaceText("{TIP08}",ilara[40]);
          body.replaceText("{FCURS8}",ilara[41]);

          body.replaceText("{ID09}", ilara[42]);
          body.replaceText("{NAM09}",ilara[43]);
          body.replaceText("{CAL09}",ilara[44]);
          body.replaceText("{TIP09}",ilara[45]);
          body.replaceText("{FCURS9}",ilara[46]);

          body.replaceText("{ID10}", ilara[47]);
          body.replaceText("{NAM10}",ilara[48]);
          body.replaceText("{CAL10}",ilara[49]);
          body.replaceText("{TIP10}",ilara[50]);
          body.replaceText("{FCURS10}",ilara[51]);

          body.replaceText("{ID11}", ilara[52]);
          body.replaceText("{NAM11}",ilara[53]);
          body.replaceText("{CAL11}",ilara[54]);
          body.replaceText("{TIP11}",ilara[55]);
          body.replaceText("{FCURS11}",ilara[56]);

          body.replaceText("{ID12}", ilara[57]);
          body.replaceText("{NAM12}",ilara[58]);
          body.replaceText("{CAL12}",ilara[59]);
          body.replaceText("{TIP12}",ilara[60]);
          body.replaceText("{FCURS12}",ilara[61]);

          body.replaceText("{ID13}", ilara[62]);
          body.replaceText("{NAM13}",ilara[63]);
          body.replaceText("{CAL13}",ilara[64]);
          body.replaceText("{TIP13}",ilara[65]);
          body.replaceText("{FCURS13}",ilara[66]);

          body.replaceText("{ID14}", ilara[67]);
          body.replaceText("{NAM14}",ilara[68]);
          body.replaceText("{CAL14}",ilara[69]);
          body.replaceText("{TIP14}",ilara[70]);
          body.replaceText("{FCURS14}",ilara[71]);

          body.replaceText("{ID15}", ilara[72]);
          body.replaceText("{NAM15}",ilara[73]);
          body.replaceText("{CAL15}",ilara[74]);
          body.replaceText("{TIP15}",ilara[75]);
          body.replaceText("{FCURS15}",ilara[76]);

          body.replaceText("{ID16}", ilara[77]);
          body.replaceText("{NAM16}",ilara[78]);
          body.replaceText("{CAL16}",ilara[79]);
          body.replaceText("{TIP16}",ilara[80]);
          body.replaceText("{FCURS16}",ilara[81]);

          body.replaceText("{ID17}", ilara[82]);
          body.replaceText("{NAM17}",ilara[83]);
          body.replaceText("{CAL17}",ilara[84]);
          body.replaceText("{TIP17}",ilara[85]);
          body.replaceText("{FCURS17}",ilara[86]);

          body.replaceText("{ID18}", ilara[87]);
          body.replaceText("{NAM18}",ilara[88]);
          body.replaceText("{CAL18}",ilara[89]);
          body.replaceText("{TIP18}",ilara[90]);
          body.replaceText("{FCURS18}",ilara[91]);

          body.replaceText("{ID19}", ilara[92]);
          body.replaceText("{NAM19}",ilara[93]);
          body.replaceText("{CAL19}",ilara[94]);
          body.replaceText("{TIP19}",ilara[95]);
          body.replaceText("{FCURS19}",ilara[96]);

          body.replaceText("{ID20}", ilara[97]);
          body.replaceText("{NAM20}",ilara[98]);
          body.replaceText("{CAL20}",ilara[99]);
          body.replaceText("{TIP20}",ilara[100]);
          body.replaceText("{FCURS20}",ilara[101]);

          body.replaceText("{ID21}", ilara[102]);
          body.replaceText("{NAM21}",ilara[103]);
          body.replaceText("{CAL21}",ilara[104]);
          body.replaceText("{TIP21}",ilara[105]);
          body.replaceText("{FCURS21}",ilara[106]);
          
          body.replaceText("{ID22}", ilara[107]);
          body.replaceText("{NAM22}",ilara[108]);
          body.replaceText("{CAL22}",ilara[109]);
          body.replaceText("{TIP22}",ilara[110]);
          body.replaceText("{FCURS22}",ilara[111]);

          body.replaceText("{ID23}", ilara[112]);
          body.replaceText("{NAM23}",ilara[113]);
          body.replaceText("{CAL23}",ilara[114]);
          body.replaceText("{TIP23}",ilara[115]);
          body.replaceText("{FCURS23}",ilara[116]);

          body.replaceText("{ID24}",ilara[117]);
          body.replaceText("{NAM24}",ilara[118]);
          body.replaceText("{CAL24}",ilara[119]);
          body.replaceText("{TIP24}",ilara[120]);
          body.replaceText("{FCURS24}",ilara[121]);
          
          body.replaceText("{ID25}", ilara[122]);
          body.replaceText("{NAM25}",ilara[123]);
          body.replaceText("{CAL25}",ilara[124]);
          body.replaceText("{TIP25}",ilara[125]);
          body.replaceText("{FCURS25}",ilara[126]);

          body.replaceText("{ID26}", ilara[127]);
          body.replaceText("{NAM26}",ilara[128]);
          body.replaceText("{CAL26}",ilara[129]);
          body.replaceText("{TIP26}",ilara[130]);
          body.replaceText("{FCURS26}",ilara[131]);

          body.replaceText("{ID27}", ilara[132]);
          body.replaceText("{NAM27}",ilara[133]);
          body.replaceText("{CAL27}",ilara[134]);
          body.replaceText("{TIP27}",ilara[135]);
          body.replaceText("{FCURS27}",ilara[136]);

          body.replaceText("{ID28}", ilara[137]);
          body.replaceText("{NAM28}",ilara[138]);
          body.replaceText("{CAL28}",ilara[139]);
          body.replaceText("{TIP28}",ilara[140]);
          body.replaceText("{FCURS28}",ilara[141]);

          body.replaceText("{ID29}", ilara[142]);
          body.replaceText("{NAM29}",ilara[143]);
          body.replaceText("{CAL29}",ilara[144]);
          body.replaceText("{TIP29}",ilara[145]);
          body.replaceText("{FCURS29}",ilara[146]);

          body.replaceText("{ID30}", ilara[147]);
          body.replaceText("{NAM30}",ilara[148]);
          body.replaceText("{CAL30}",ilara[149]);
          body.replaceText("{TIP30}",ilara[150]);
          body.replaceText("{FCURS30}",ilara[151]);

          body.replaceText("{ID31}", ilara[152]);
          body.replaceText("{NAM31}",ilara[153]);
          body.replaceText("{CAL31}",ilara[154]);
          body.replaceText("{TIP31}",ilara[155]);
          body.replaceText("{FCURS31}",ilara[156]);

          body.replaceText("{ID32}", ilara[157]);
          body.replaceText("{NAM32}",ilara[158]);
          body.replaceText("{CAL32}",ilara[159]);
          body.replaceText("{TIP32}",ilara[160]);
          body.replaceText("{FCURS32}",ilara[161]);

          body.replaceText("{ID33}", ilara[162]);
          body.replaceText("{NAM33}",ilara[163]);
          body.replaceText("{CAL33}",ilara[164]);
          body.replaceText("{TIP33}",ilara[165]);
          body.replaceText("{FCURS33}",ilara[166]);

          body.replaceText("{ID34}", ilara[167]);
          body.replaceText("{NAM34}",ilara[168]);
          body.replaceText("{CAL34}",ilara[169]);
          body.replaceText("{TIP34}",ilara[170]);
          body.replaceText("{FCURS34}",ilara[171]);

          body.replaceText("{ID35}", ilara[172]);
          body.replaceText("{NAM35}",ilara[173]);
          body.replaceText("{CAL35}",ilara[174]);
          body.replaceText("{TIP35}",ilara[175]);
          body.replaceText("{FCURS35}",ilara[176]);

          body.replaceText("{ID36}", ilara[177]);
          body.replaceText("{NAM36}",ilara[178]);
          body.replaceText("{CAL36}",ilara[179]);
          body.replaceText("{TIP36}",ilara[180]);
          body.replaceText("{FCURS36}",ilara[181]);

          body.replaceText("{ID37}", ilara[182]);
          body.replaceText("{NAM37}",ilara[183]);
          body.replaceText("{CAL37}",ilara[184]);
          body.replaceText("{TIP37}",ilara[185]);
          body.replaceText("{FCURS37}",ilara[186]);

          body.replaceText("{ID38}", ilara[187]);
          body.replaceText("{NAM38}",ilara[188]);
          body.replaceText("{CAL38}",ilara[189]);
          body.replaceText("{TIP38}",ilara[190]);
          body.replaceText("{FCURS38}",ilara[191]);

          body.replaceText("{ID39}", ilara[192]);
          body.replaceText("{NAM39}",ilara[193]);
          body.replaceText("{CAL39}",ilara[194]);
          body.replaceText("{TIP39}",ilara[195]);
          body.replaceText("{FCURS39}",ilara[196]);

          body.replaceText("{ID40}", ilara[197]);
          body.replaceText("{NAM40}",ilara[198]);
          body.replaceText("{CAL40}",ilara[199]);
          body.replaceText("{TIP40}",ilara[200]);
          body.replaceText("{FCURS40}",ilara[201]);

          body.replaceText("{ID41}", ilara[202]);
          body.replaceText("{NAM41}",ilara[203]);
          body.replaceText("{CAL41}",ilara[204]);
          body.replaceText("{TIP41}",ilara[205]);
          body.replaceText("{FCURS41}",ilara[206]);

          body.replaceText("{ID42}", ilara[207]);
          body.replaceText("{NAM42}",ilara[208]);
          body.replaceText("{CAL42}",ilara[209]);
          body.replaceText("{TIP42}",ilara[210]);
          body.replaceText("{FCURS42}",ilara[211]);

          body.replaceText("{ID43}", ilara[212]);
          body.replaceText("{NAM43}",ilara[213]);
          body.replaceText("{CAL43}",ilara[214]);
          body.replaceText("{TIP43}",ilara[215]);
          body.replaceText("{FCURS43}",ilara[216]);

          body.replaceText("{ID44}", ilara[217]);
          body.replaceText("{NAM44}",ilara[218]);
          body.replaceText("{CAL44}",ilara[219]);
          body.replaceText("{TIP44}",ilara[220]);
          body.replaceText("{FCURS44}",ilara[221]);

          body.replaceText("{ID45}", ilara[222]);
          body.replaceText("{NAM45}",ilara[223]);
          body.replaceText("{CAL45}",ilara[224]);
          body.replaceText("{TIP45}",ilara[225]);
          body.replaceText("{FCURS45}",ilara[226]);

          body.replaceText("{ID46}", ilara[227]);
          body.replaceText("{NAM46}",ilara[228]);
          body.replaceText("{CAL46}",ilara[229]);
          body.replaceText("{TIP46}",ilara[230]);
          body.replaceText("{FCURS46}",ilara[231]);

          body.replaceText("{ID47}", ilara[232]);
          body.replaceText("{NAM47}",ilara[233]);
          body.replaceText("{CAL47}",ilara[234]);
          body.replaceText("{TIP47}",ilara[235]);
          body.replaceText("{FCURS47}",ilara[236]);


          body.replaceText("{ID48}", ilara[237]);
          body.replaceText("{NAM48}",ilara[238]);
          body.replaceText("{CAL48}",ilara[239]);
          body.replaceText("{TIP48}",ilara[240]);
          body.replaceText("{FCURS48}",ilara[241]);

          body.replaceText("{ID49}", ilara[242]);
          body.replaceText("{NAM49}",ilara[243]);
          body.replaceText("{CAL49}",ilara[244]);
          body.replaceText("{TIP49}",ilara[245]);
          body.replaceText("{FCURS49}",ilara[246]);

          body.replaceText("{ID50}", ilara[247]);
          body.replaceText("{NAM50}",ilara[248]);
          body.replaceText("{CAL50}",ilara[249]);
          body.replaceText("{TIP50}",ilara[250]);
          body.replaceText("{FCURS50}",ilara[251]);

          body.replaceText("{ID51}", ilara[252]);
          body.replaceText("{NAM51}",ilara[253]);
          body.replaceText("{CAL51}",ilara[254]);
          body.replaceText("{TIP51}",ilara[255]);
          body.replaceText("{FCURS51}",ilara[256]);

          body.replaceText("{ID52}", ilara[257]);
          body.replaceText("{NAM52}",ilara[258]);
          body.replaceText("{CAL52}",ilara[259]);
          body.replaceText("{TIP52}",ilara[260]);
          body.replaceText("{FCURS52}",ilara[261]);

          body.replaceText("{ID53}", ilara[262]);
          body.replaceText("{NAM53}",ilara[263]);
          body.replaceText("{CAL53}",ilara[264]);
          body.replaceText("{TIP53}",ilara[265]);
          body.replaceText("{FCURS53}",ilara[266]);

          body.replaceText("{ID54}", ilara[267]);
          body.replaceText("{NAM54}",ilara[268]);
          body.replaceText("{CAL54}",ilara[269]);
          body.replaceText("{TIP54}",ilara[270]);
          body.replaceText("{FCURS54}",ilara[271]);

          body.replaceText("{ID55}", ilara[272]);
          body.replaceText("{NAM55}",ilara[273]);
          body.replaceText("{CAL55}",ilara[274]);
          body.replaceText("{TIP55}",ilara[275]);
          body.replaceText("{FCURS55}",ilara[276]);

          body.replaceText("{ID56}", ilara[277]);
          body.replaceText("{NAM56}",ilara[278]);
          body.replaceText("{CAL56}",ilara[279]);
          body.replaceText("{TIP56}",ilara[280]);
          body.replaceText("{FCURS56}",ilara[281]);

          body.replaceText("{ID57}", ilara[282]);
          body.replaceText("{NAM57}",ilara[283]);
          body.replaceText("{CAL57}",ilara[284]);
          body.replaceText("{TIP57}",ilara[285]);
          body.replaceText("{FCURS57}",ilara[286]);

          body.replaceText("{ID58}", ilara[287]);
          body.replaceText("{NAM58}",ilara[288]);
          body.replaceText("{CAL58}",ilara[289]);
          body.replaceText("{TIP58}",ilara[290]);
          body.replaceText("{FCURS58}",ilara[291]);

          body.replaceText("{ID59}", ilara[292]);
          body.replaceText("{NAM59}",ilara[293]);
          body.replaceText("{CAL59}",ilara[294]);
          body.replaceText("{TIP59}",ilara[295]);
          body.replaceText("{FCURS59}",ilara[296]);

          body.replaceText("{ID60}", ilara[297]);
          body.replaceText("{NAM60}",ilara[298]);
          body.replaceText("{CAL60}",ilara[299]);
          body.replaceText("{TIP60}",ilara[300]);
          body.replaceText("{FCURS60}",ilara[301]);

          body.replaceText("{ID61}", ilara[302]);
          body.replaceText("{NAM61}",ilara[303]);
          body.replaceText("{CAL61}",ilara[304]);
          body.replaceText("{TIP61}",ilara[305]);
          body.replaceText("{FCURS61}",ilara[306]);

          body.replaceText("{ID62}", ilara[307]);
          body.replaceText("{NAM62}",ilara[308]);
          body.replaceText("{CAL62}",ilara[309]);
          body.replaceText("{TIP62}",ilara[310]);
          body.replaceText("{FCURS62}",ilara[311]);

          body.replaceText("{ID63}", ilara[312]);
          body.replaceText("{NAM63}",ilara[313]);
          body.replaceText("{CAL63}",ilara[314]);
          body.replaceText("{TIP63}",ilara[315]);
          body.replaceText("{FCURS63}",ilara[316]);

          body.replaceText("{ID64}", ilara[317]);
          body.replaceText("{NAM64}",ilara[318]);
          body.replaceText("{CAL64}",ilara[319]);
          body.replaceText("{TIP64}",ilara[320]);
          body.replaceText("{FCURS64}",ilara[321]);

          finalizaImpresion(NOMBRE,GRUPO,ID,opceduca,PLANTEL,PERIODO,tempDocFile,pdfFolder,tempFolder,tempFile)
    })
}

//*********************************************** */

//*********************************************************************** */
//TOMA LOS DATOS EN LA HOJA PASO2 Y RESUME LAS MATERIAS DE VARIAS CALIFICACIONES
//EN UN SOLO REGISTRO PROMEDIADO, EN CASO DE GLOBAL TOMA LA CALIFICACION DEL GLOBAL
//*********************************************************************** */

function junta(arrayIkasleak)
{
  // *************** ARRAY ORDENADA POR ALUMNO Y CVE MATERIA Y PARCIAL

//var pasoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PASO")
const califSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CALIFICACIONES")

califSheet.getRange("A2:T").clearContent()

var fila=1
var corrAnt=""
var cvematAnt=""
var pvez=0
var ind=0
var calif=0
var promCalif=0
var sumCalif=0
var Lrow=1
let nom=""
let matri=""
let opcedu=""
let plantelorg=""
let periorg=""
let prom=""
let idasigorg=""
var legalorg=""
var nomInterno=""
var ordinario=""
var fecha=""
var cveprofe=""
var urlasig=""
var C14=""
var C15=""
var creditos=""
var fechafin=""
var grupikasle=""
var turnoik=""
var C18=""
var docenteorg=""
var parcialAnt=""
var cveAsigna="";
var indGbl=0;

let regSalida =[]

var ikasleDatuak=bdIkasleOrria.getDataRange().getDisplayValues(); //datos alumnos 040424
var asignaDatuak=asignaOrria.getDataRange().getDisplayValues();   //datos asignatura 040424

for (var k=0;k<arrayIkasleak.length;k++)
{ //ID ASIGNATURA,	PARCIAL,	TIPO,	CALIFICACION,	PERIODO,	INSITUTCIONAL,	FECHA,	ID PROFESOR.	URL LISTA,	CREDITOS,	FIN PERIODO
  //    0             1        2       3            4          5              6          7           8           9        10
  var pase=0;
  arrayIkasleak[k].forEach(row => 
  { 
    
    if(corrAnt == row[5]&& cvematAnt == row[0])
      {
            ind++
        if (indGbl==0)  
          {
                if (row[2] == "gbl")
                  {
                    sumCalif=parseInt(row[3]) //5
                    ind=1
                    indGbl=1
                  }
                  else
                  {
                    sumCalif=sumCalif+parseInt(row[3])      //ACUMULA CALIFICACIONES
                  }
              
                pase=0;
               if (parseInt(row[3])==6)
                  {
                      pase=1;
                  }
                
              
               wparcial=row[1];
               parcialAnt = wparcial;
               fecha=row[6];
               fechafin= row[10];
               ordinario=row[2];
               cveprofe=row[7];
               urlasig=row[8];
               docenteNom=""; //hy que conseguirlo??? se usa o no????
         }
         else
          { 
              //revisa si es asignatura especial
              if (row[1]==9)
              {
                sumCalif=parseInt(row[3]);      //SI ES ESPECIAL ESTE VALOR DOMINA  9
                ind=1;   
                wparcial=9;
                if (parseInt(row[3])==6)
                  {
                      pase=1;
                  }
              }
              else
              {
              sumCalif=sumCalif+parseInt(row[3])      //ACUMULA CALIFICACIONES de globales ******
              wparcial=row[1];
              }
              parcialAnt = wparcial;
            }
      }
    else
      {
        if (pvez==0)
          {
            pvez=1
          } 
        else
          // NO ES PRIMERA VEZ
          {
                  promCalif=sumCalif/ ind;
                  calif=Math.round(promCalif);
                  if (calif<7)
                  {
                    if (pase==0)
                      {
                        calif=5
                      }
                  }

                //*************************    asienta registro en calificaciones
                //var Lrow=califSheet.getLastRow()
                Lrow++

                regSalida.push([nom,matri,opcedu,plantelorg,periorg,parcialAnt,cveAsigna,calif,legalorg,nomInterno,ordinario,fecha,cveprofe,urlasig,creditos,fechafin,grupikasle,turnoik,C18,docenteorg])

                promCalif=0
                sumCalif=0
                ind=0
          } 
          // FIN PVEZ
                                               
            if(corrAnt != row[5])
              {                                 //OBTIENE EL NOMBRE DEL ALUMNO
                  //************************************** */040424
                  var ikasleDatuakF=ikasleDatuak.filter(ilara=>ilara[1]==row[5])
                  if (ikasleDatuakF.length>0)
                  {
                    nom=ikasleDatuakF[0][0];
                    opcedu=ikasleDatuakF[0][9];    //OBTIENE OPC EDU
                    turnoik=ikasleDatuakF[0][8];    //OBTIENE TURNO
                    grupikasle=ikasleDatuakF[0][5]; //OBTIENE GRUPO
                  }
                  else
                  {
                    nom="nombre invalido";
                    opcedu="alumno invalido"
                  }

                  matri=row[5];
                  corrAnt=row[5];
                  //******************************************* */
              }                             
            
            cvematAnt = row[0];

            plantelorg="CAFETALES";
            periorg=row[4];
            wparcial=row[1];
            cveAsigna=row[0];
            //**************************************** UNA ASIGNATURA PUEDE SER GLOBAL EN TRES PARCIALES */
            if (row[2] == "gbl")
              {
                sumCalif=parseInt(row[3])
                ind=1
                indGbl=1;
              }
              else
              {
                sumCalif=sumCalif+parseInt(row[3])      //ACUMULA CALIFICACIONES
                pase=0;
               if (parseInt(row[3])==6)
                  {
                      pase=1;
                  }
                indGbl=0;
                ind++
              }

            //OBTENER NOMBRES DE ASIGNATURA
            
            var asignaDatuakF=asignaDatuak.filter(ilara=>ilara[0]==cveAsigna&& ilara[3]==opcedu);
            if (asignaDatuakF.length>0)
            {
              legalorg=asignaDatuakF[0][1];
              nomInterno=asignaDatuakF[0][2];
              creditos=asignaDatuakF[0][17];
            }
            else
            {
              legalorg="Asignatura invalida" ;
              nomInterno="Asignatura invalida"; 
              creditos=0;
            }

            
            ordinario=row[2]
            fecha=row[6];
            cveprofe=row[7];
            urlasig=row[8];
            fechafin=row[10];
            
            C18=""
            docenteNom ="" //hay que obtener
            parcialAnt = wparcial;
    }
          //FIN REGISTROS DIFERENTES     
  })// termina forEach

} //termina for

promCalif=sumCalif/ ind;
              
calif=Math.round(promCalif);

if (calif<7)
{
    calif=5

}
//*************************    asienta registro en calificaciones
//var Lrow=califSheet.getLastRow()

regSalida.push([nom,matri,opcedu,plantelorg,periorg,parcialAnt,cveAsigna,calif,legalorg,nomInterno,ordinario,fecha,cveprofe,urlasig,creditos,fechafin,grupikasle,turnoik,C18,docenteorg])

// ESCRIBE BLOQUE DE INFORMACION en CALIFICACIONES
califSheet.getRange(2,1,regSalida.length,20).setValues(regSalida)
  

}
//*************************************CARGA DATOS EN HOJA PLANTILLA */
//
//******************************************************************* */
function cargaPlantilla()
{
                                                                      //DETERMINA OPC EDU A TRABAJAR
  var wopcedu=sheetAct.getRange("E1").getDisplayValue();
                                                                      //OBTIENE ASIGNATURAS
  var tablaAsigna=asignaOrria.getDataRange().getDisplayValues();
                                                                     //FILTRA TABLA ASIGNATURAS POR OPC EDUCA
  var tablaAsignaF=tablaAsigna.filter(ilara=>ilara[3]==wopcedu);
  if (tablaAsignaF.length>0)
  {
    //DEBE ESTAR ORDENADA POR OPC EDU, PERIODO,TIPO ASIGNA, CVE ASIGNA
    var array1 =[];
    for (var i=0;i<tablaAsignaF.length;i++)
    {
                  //periodo           tipo asigna         clave asigna        nombre legal
      array1.push([tablaAsignaF[i][8],tablaAsignaF[i][18],tablaAsignaF[i][0],tablaAsignaF[i][1]])
    
    }
    array1.sort();
  }
  else
  {
    console.log("error gacho");
    return;
  }

   //***ARMA PLANTILLA LINEA 6 Y SUBSECUENTES */
      //OBTIENE LOS ALUMNOS A PROCESAR DE CARATULA CON MARCA "SI"
  var wdatuak= sheetAct.getDataRange().getDisplayValues();  //datos de alumno en hoja caratula
  //var wdatuakF=wdatuak.filter(ilara=>ilara[3]=="SI");
  var wdatuakF=wdatuak.filter(ilara=>ilara[3]=="SI"&&ilara[2]!="");
  if (wdatuakF.length>0)
  {
    var arrayAlum=[];
    var wlr=sheetPlant.getLastRow();
    var wlc=sheetPlant.getLastColumn();
    sheetPlant.getRange(6,1,wlr,wlc).clearContent();
    wlr=6;
  }
  else
  {
    console.log("no hay datos a procesar")
    return;
  }

  
var califDatua =sheetCalif.getDataRange().getDisplayValues(); //obtiene datos de base de calificaciones
var warray=[];
  //FILTRO POR CADA ALUMNO
  for (var k=0;k<wdatuakF.length;k++)
  {
    var numCdtos=0;
    var promCalif=0;
    var wikasle =wdatuakF[k][2];                    //NOMBRE ALUMNO
    arrayAlum.push("*",wikasle);

      //BARRE CADA ASIGNATURA DE CATALOGO (array1)
      var m=0;
      var indP=0;
      for (var j=0;j<array1.length;j++)
      {
        //COMPARA VS FILTRO POR CADA ASIGNATURA  DE HOJA CALIFICACIONES
          //SI EXISTE
              //ASIGNA valores de calificacion y tipo Y FECHA EN ARRAY
          //SI NOOO EXISTE
            //DEJA ESPACIOS EN CAMPOS DE ARRAY
        
        //                                        nombre calif /nombre    cve asigna
        var califDatuaF = califDatua.filter(ilara=>ilara[0]==wikasle&&ilara[6]==array1[j][2]);
        if (califDatuaF.length>0)
        {
            m++;
            var wcalif=califDatuaF[0][7]        //calificacion
            var wtipoCal=califDatuaF[0][10]     //tipo
            var wfechaCal=califDatuaF[0][11]    //fecha
            if (array1[j][1]="ASIG")
            {                                   //PROMEDIA SOLO LAS ASIGNATURAS "SAETI"
              promCalif=(promCalif+Number(wcalif))            
              numCdtos=Number(califDatuaF[0][14])+numCdtos //creditos
              indP++
            }
        }
        else
        {
          var wcalif=""        //calificacion
          var wtipoCal=""     //tipo
          var wfechaCal=""    //fecha
        }
        //asignar valores de CVE ASIG y NOM ASIG DE CATALOGO EN ARRAY
        arrayAlum.push(array1[j][2],array1[j][3],wcalif,wtipoCal,wfechaCal);
      }
     var wprom =(promCalif/indP).toFixed(1);
      arrayAlum.push(numCdtos,wprom,m);
      //AL FINALIZAR CADA ALUMNO
          // ESCRIBE EN PLANTILLA

      //var warray=[];
      warray.push(arrayAlum);
      //sheetPlant.getRange(wlr,2,1,arrayAlum.length).setValues(warray);
      //arrayAlum=[];
      arrayAlum=[];
  }
  wlr=sheetPlant.getLastColumn();
  
  sheetPlant.getRange(6,1,warray.length,warray[0].length).setValues(warray);
}

//**************************************************** */
// OBTIENE INFORMACION DE CALIFICACIONES DE TABLA DE CALIFICACIONES FILTRANDO POR OPC EDU Y GRUPO
// ORDENA EL ARRAY POR ALUMNO Y CVE ASIGNATURA PARA PROCESAR DESPUES
// OPCION DE TRAABAJARLO EN MEMORIA, PERO SE ESCRIBE EN HOJA PASO2 DE ARCHIVO DE HISTORIALES POR SI 
// SE INTERRUMPE EL PROCESO SE REINICIA EN ESE PASO
//**************************************************** */
function actualiza()
{
  var opceduca=sheetAct.getRange("E1").getDisplayValue();       //OBTIENE LA OPCION EDUCATIVA A TRABAJAR
                                                               //OBTIENE LA CLAVE NUMERICA DE LA OPCIÓN EDUCATIVA *********************************
  var wopceduca=opcEduOrria.getDataRange().getDisplayValues(); 
  var wopceducaF=wopceduca.filter(ilara=>ilara[0]==opceduca);

  if(wopceducaF.length>0)
  {
    var opcEduNum=wopceducaF[0][8];
  
  }
  else 
  {
    console.log("error al obtener opc educativa: "+wopceduca);
    return
  }

  var wtaldea=sheetAct.getRange("E2").getDisplayValue();          //OBTIENE EL GRUPO




                                          //obtener DATOS DE TABLA CALIFICACIONES ************************************
  var califDatua=tabCalifOrria.getRange("B2:N").getDisplayValues();
  if (wtaldea.length>0)
  {
  var califDatuaF=califDatua.filter(ilara=>ilara[2]==opcEduNum&&ilara[1]==wtaldea); //FILTRA SOLO LOS DE LA OPC EDU y GRUPO SELECCIONADO A PROCESAR
  }
  else
  {
    var califDatuaF=califDatua.filter(ilara=>ilara[2]==opcEduNum); //FILTRA SOLO LOS DE LA OPC EDU A PROCESAR
  }
  if (califDatuaF.length>0)
  {
    var ordenatu1=califDatuaF.sort();
    var ikasleant="";
    var ikasleArray=[];
    var wuno=0;

    //*********************************************** */040424
      var wlrow=sheetPaso.getLastRow();
      var wlcol=sheetPaso.getLastColumn();
      sheetPaso.getRange(2,1,wlrow,wlcol).clearContent();
//******************************************** */



    //para hacer un segundo sort crea array alumno por alumno con cve asigna en primer lugar
    for (var i=0;i<califDatuaF.length;i++)
    {
      if(ikasleant!=ordenatu1[i][0])
      {
        
        if (wuno==0)
        {
          wuno=1;

        }
        else
        {
          //carga array en otra aray ordenada
          var ikasleArrayO=ikasleArray.sort()

          //***************************************** */040424
          var wlrow=sheetPaso.getLastRow()+1;
          sheetPaso.getRange(wlrow,1,ikasleArrayO.length,11).setValues(ikasleArrayO);
          //***************************************** */
          
          arrayIkasleak.push(ikasleArrayO)
          ikasleArray=[]
          
        }
        ikasleant=ordenatu1[i][0];
      }
      
    ikasleArray.push([ordenatu1[i][5],ordenatu1[i][4],ordenatu1[i][7],ordenatu1[i][6],ordenatu1[i][3],ordenatu1[i][0],ordenatu1[i][8],ordenatu1[i][9],ordenatu1[i][10],ordenatu1[i][11],ordenatu1[i][12]])
      
    } //fin for
    var ikasleArrayO=ikasleArray.sort()

    //***************************************** */ 040424
    var wlrow=sheetPaso.getLastRow()+1;
    sheetPaso.getRange(wlrow,1,ikasleArrayO.length,11).setValues(ikasleArrayO);
    //***************************************** */
    
    arrayIkasleak.push(ikasleArrayO)
    
  }
  else
  {
    console.log("error al obtener datos de tabla calificaciones: opcEduNum "+opcEduNum);
    return
  }
  //TIENE QUE OBTENER CADA ALUMNO ARRAY MULTINIVEL

/*var filara=2;
  prueba.getRange("A2:K").clearContent();
  for (var k=0;k<arrayIkasleak.length;k++)
  {
    prueba.getRange(filara,1,arrayIkasleak[k].length,11).setValues(arrayIkasleak[k]);
    filara=filara+arrayIkasleak[k].length
  }
  */

 junta(arrayIkasleak);
}

function bddata()
{
  var opedu = catAcade.getSheetByName("OPCIONES EDUCATIVAS");
  var d1opedu = opedu.getRange("A1:A").getValues();
  var d2opedu = opedu.getRange("C1:C").getValues();
  var d3opedu = opedu.getRange("F1:F").getValues();
 
  SheetDB.getRange("H1:J").clearContent();
  
 SheetDB.getRange(1, 8, d1opedu.length, 1).setValues(d1opedu);
 SheetDB.getRange(1, 9, d2opedu.length, 1).setValues(d2opedu);
 SheetDB.getRange(1, 10, d3opedu.length, 1).setValues(d3opedu);
  //var plant = catAcade.getSheetByName("PLANTELES");
  //var d1plant = plant.getRange("A1:B").getValues();
  SheetDB.getRange("M1:N").clearContent();
  
 SheetDB.getRange(1, 13, 1, 1).setValue("CAFETALES");
 
 var taldeak= catAcade.getSheetByName('GRUPOS')
 var d1taldeak =taldeak.getRange("A1:B").getDisplayValues();
 SheetDB.getRange("C1:D").clearContent();
 SheetDB.getRange(1,3,d1taldeak.length,2).setValues(d1taldeak);
 
 var asign = catAcade.getSheetByName("TABLA ASIGNATURAS");
   var d1asign = asign.getRange("A1:C").getValues();
   var d2asign = asign.getRange("I1:I").getValues();
   var d3asign = asign.getRange("D1:D").getValues();
   var d4asign = asign.getRange("P1:P").getValues();
   
 SheetDB.getRange("P1:W").clearContent(); 
 
 SheetDB.getRange(1, 16, d1asign.length, 3).setValues(d1asign);
 SheetDB.getRange(1, 19, d1asign.length, 1).setValues(d2asign);
 SheetDB.getRange(1, 20, d1asign.length, 1).setValues(d3asign);
 SheetDB.getRange(1, 21, d1asign.length, 1).setValues(d4asign);
 
   var modest = catAcade.getSheetByName("MODALIDADES DE ESTUDIO");
   var d1modest = modest.getRange("A1:A").getValues();
  SheetDB.getRange("W1:W").clearContent();
  
 SheetDB.getRange(1, 23, d1modest.length, 1).setValues(d1modest);
   var perso = catAcade.getSheetByName("PERSONAL");
   var d1perso = perso.getRange("A1:A").getValues();
   var d2perso = perso.getRange("E1:E").getValues();
 SheetDB.getRange("Z1:AA").clearContent();
  
 SheetDB.getRange(1, 26, d1perso.length, 1).setValues(d1perso);
 SheetDB.getRange(1, 27, d1perso.length, 1).setValues(d2perso);
   var alumn = bdIkasle.getSheetByName("ACTIVOS_FORMATEADO");
   var d1alumn = alumn.getRange("A1:K").getValues();
   var d2alumn = alumn.getRange("R1:R").getValues();
   var d3alumn = alumn.getRange("N1:N").getValues();
   var tutoname = alumn.getRange("Q1:Q").getValues();
   var plantel = alumn.getRange("W1:W").getValues();
  SheetDB.getRange("AB1:AK").clearContent();
 SheetDB.getRange("AN1:AN").clearContent();
 SheetDB.getRange("AP1:AP").clearContent();
 SheetDB.getRange("AQ1:AQ").clearContent();
  
  SheetDB.getRange(1, 28, d1alumn.length, 11).setValues(d1alumn);
  SheetDB.getRange(1, 39, d2alumn.length, 1).setValues(d2alumn);
  SheetDB.getRange(1, 40, d3alumn.length, 1).setValues(d3alumn);
  SheetDB.getRange(1, 42, tutoname.length, 1).setValues(tutoname);
  SheetDB.getRange(1, 43, plantel.length, 1).setValues(plantel)
 }
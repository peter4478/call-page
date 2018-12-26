var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/18FQCLVkIVVd1Z9wImbUODJYS-7kYxps45Kz2BvCVIM8/edit#gid=1455322467");
var ssCallPageDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1mAZdPPuGiD03hFQ5-sTfEKWzWJ-oxhnC3KJ3D6xtjp0/edit#gid=1984671374");
var ssSPI = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1FDs2QVTHqdEJ6HzynpZl4hg_qFLzMRC0nPVSf_3e_5g/edit#gid=0");

var sheetDatabaza = ss.getSheetByName("Databáza");
var sheetHistoria = ss.getSheetByName("Call_História");
var historiaSPI = ssSPI.getSheetByName("Call_História");
var sheetSPI = ssSPI.getSheetByName("Databáza");






function deleteZmluva(_idOsoba, _idZmluva){  
  var shDat = SpreadsheetApp.openById("1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o").getSheetByName("Zmluvy");
  var data = shDat.getDataRange().getValues();

  for(var i = 0; i < data.length; i++){  
    var idZmluva = data[i][data[0].indexOf("ID - Zmluva")];
    var idOsoba = data[i][data[0].indexOf("ID - Osoba")];
    
    if(idZmluva == _idZmluva && idOsoba == _idOsoba){
      shDat.getRange(i + 1, data[0].indexOf("Delete") + 1).setValue("ok");
    }
  }  
}


function deleteIO(_idOsoba, _idIO){  //IO = ina osoba
  var shDat = SpreadsheetApp.openById("1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o").getSheetByName("Iné osoby");
  var data = shDat.getDataRange().getValues();

  for(var i = 0; i < data.length; i++){  
    var idIO = data[i][data[0].indexOf("ID - Iná osoba")];
    var idOsoba = data[i][data[0].indexOf("ID - Databáza")];
    
    if(idIO == _idIO && idOsoba == _idOsoba){
      shDat.getRange(i + 1, data[0].indexOf("Delete") + 1).setValue("ok");
    }
  }  
}

function deleteDatabaza(_idDat){
  var shDat = SpreadsheetApp.openById("1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o").getSheetByName("Databáza");
  var data = shDat.getDataRange().getValues();

  for(var i = 0; i < data.length; i++){ 
    var idDad = data[i][data[0].indexOf("ID")];
    
    if(idDad == _idDat){
      shDat.getRange(i + 1, data[0].indexOf("Delete") + 1).setValue("ok");
    }
  }  
}

function deleteIK(_idDat){
  var shDat = SpreadsheetApp.openById("1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0").getSheetByName("IK");
  var data = shDat.getDataRange().getValues();

  for(var i = 0; i < data.length; i++){ 
    var idDad = data[i][data[0].indexOf("ID")];
    
    if(idDad == _idDat){
      shDat.getRange(i + 1, data[0].indexOf("Delete") + 1).setValue("ok");
    }
  }  
}

function deleteDohoda(_idIK, _idDohoda){  
  var shDat = SpreadsheetApp.openById("1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0").getSheetByName("Dohody");
  var data = shDat.getDataRange().getValues();

  for(var i = 0; i < data.length; i++){  
    var idIK = data[i][data[0].indexOf("ID - IK")];
    var idDohoda = data[i][data[0].indexOf("ID - Dohoda")];
    
    if(idIK == _idIK && idDohoda == _idDohoda){
      shDat.getRange(i + 1, data[0].indexOf("Delete") + 1).setValue("ok");
    }
  }  
}





//=================================================================
//doGet
//=================================================================
function doGet(e){  
  if(e.parameter.track == "ik"){ 
    var htmlTemplate = HtmlService.createTemplateFromFile('IK.html');
    htmlTemplate.dataFromServerTemplate = { id: e.parameter.id };  
    var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle('IK');
    return htmlOutput;
  }
  else if (e.parameter.track == "dohoda") {
    var htmlTemplate = HtmlService.createTemplateFromFile('Dohoda.html');
    htmlTemplate.dataFromServerTemplate = { id: e.parameter.id };  
    var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle('Dohoda'); 
    return htmlOutput;
  }else{
    return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Databáza');
  }
}


//=================================================================
//Include - Integruje html subory do "Page"
//=================================================================
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}



//=================================================================
//Potvrdenie IK
//=================================================================
function submitIK(_id){
  var sheet = SpreadsheetApp.openById("1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0").getSheetByName("IK");
  var data = sheet.getDataRange().getValues();

  for(var i = 0; i < data.length; i++){  
    var id = data[i][data[0].indexOf("ID")];
    
    if(id == _id){
      sheet.getRange(i + 1, data[0].indexOf("Stav") + 1).setValue("odoslané, potvrdené");
    }
  }  
}



//=================================================================
//Potvrdenie IK
//=================================================================
function getDohoda(_id){
  var sheet = SpreadsheetApp.openById("1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0").getSheetByName("Dohody");
  var data = sheet.getDataRange().getValues();
  
  var index;

  for(var i = 0; i < data.length; i++){  
    var id = data[i][data[0].indexOf("ID - Dohoda")];
    
    if(id == _id){
      index = i + 1;
    }
  }
  
  var dateStart = sheet.getRange(index, 1).getValue();
  var dateEnd = sheet.getRange(index, 2).getValue();
  var dohoda = sheet.getRange(index, 3).getValue();

  var timestamp_format = "dd.MM.yyy"; var timezone = "GMT+2";
  var _dateStart = "";
  var _dateEnd = "";
  try 
  {
    dateStart = Utilities.formatDate(dateStart, timezone, timestamp_format);
    dateEnd = Utilities.formatDate(dateEnd, timezone, timestamp_format);
  } catch (error) {}
  
  var text = { "dateStart":dateStart, "dateEnd":dateEnd, "dohoda":dohoda }; 
  return text;

}


//=================================================================
//odosle info o priebehu dohody
//=================================================================
function setDohoda(_id, stav, dateNewEnd, info){
  var sheet = SpreadsheetApp.openById("1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0").getSheetByName("Dohody");
  var data = sheet.getDataRange().getValues();
  
  var index;

  for(var i = 0; i < data.length; i++){  
    var id = data[i][data[0].indexOf("ID - Dohoda")];    
    if(id == _id){
      index = i + 1;      
    }
  }
  sheet.getRange(index, 2).setValue(dateNewEnd);
  sheet.getRange(index, 4).setValue(stav);
  sheet.getRange(index, 5).setValue(info);
  
  data = sheet.getDataRange().getValues();  
  var idIK = sheet.getRange(index, 7).getValue();
  var sumDohody = 0;
  var sumDohodySplnene = 0;
  for(var i = 0; i < data.length; i++){
    var id = data[i][data[0].indexOf("ID - IK")];
    if(id == idIK){
      sumDohody++;
      var check = sheet.getRange(i + 1, 4).getValue();
      if(check == "splnená"){
        sumDohodySplnene++;        
      }
    }
  }
  
  var sheet = SpreadsheetApp.openById("1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0").getSheetByName("IK");
  var data = sheet.getDataRange().getValues();
  for(var i = 0; i < data.length; i++){
    var _idIK = data[i][data[0].indexOf("ID")];
    if(idIK == _idIK){
      var percenta = (sumDohodySplnene / sumDohody) * 100;
      sheet.getRange(i + 1, 4).setValue(sumDohody + "/" + sumDohodySplnene + " - " + Math.round(percenta) + "%");
    }
  }
}





//=================================================================
//Resetne CP a NK
//=================================================================
function resetCPNK(){  
  var sheet = SpreadsheetApp.openById("1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak").getSheetByName("Call-Page");
  var hlavicka = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var rowLast = sheet.getLastRow() + 1;  
  for(var i = 2; i < rowLast; i++){    
    sheet.getRange(i, hlavicka[0].indexOf("CP") + 1).setValue(0);
    sheet.getRange(i, hlavicka[0].indexOf("NK") + 1).setValue(0);
  }
}










//=================================================================
//Firmy - Vráti cislo riadku pouzitelneho kontaktu
//=================================================================
function firmyGetIndex(sheetName){  
  var sheet = SpreadsheetApp.openById("1Jpv1Iw6qlVU-oPJKkWdJuCbqww6V8GqeM7o6JNZcvMA").getSheetByName(sheetName);
  var index = sheet.getRange(1, 1).getValue();  
  index++; 
  var temp = sheet.getRange(index, 6).getValue();
  while (temp == "číslo neexistuje" || temp == "nemá záujem - starobný dôchodca" || temp == "klient" || temp == "nekontaktovať")
  {
    index++;
    var temp = sheet.getRange(index, 6).getValue();
  }
  var lastRow = sheet.getLastRow();
  if(index > lastRow)//na konci resetne zoznam
  {
    index = 2;
    sheet.getRange(1, 1).setValue(index);  
    var temp = sheet.getRange(1, 2).getValue(); temp++;
    sheet.getRange(1, 2).setValue(temp);
  }  
  sheet.getRange(1, 1).setValue(index);
 
  var name = sheet.getRange(index, 1).getValue();
  var phone = sheet.getRange(index, 2).getValue();
  var email = sheet.getRange(index, 4).getValue();
  var web = sheet.getRange(index, 5).getValue();
  var ico = sheet.getRange(index, 6).getValue();
  var kategoria = sheet.getRange(index, 7).getValue();
  var info = sheet.getRange(index, 8).getValue();
  
  
  //Logger.log(poslednyKontaktDatum)
  //Logger.log(poslednyKontaktOsoba)
  //Logger.log(poslednyKontaktStav)
  
  var timestamp_format = "yyy-MM-dd"; var timezone = "GMT+2";
  var timestamp_format1 = "HH:mm"; var timezone = "GMT+2"; 
  var timestamp_format2 = "dd.MM.yyy HH:mm"; var timezone = "GMT+2";
  var datum = "";
  var cas = "";
  try 
  {
    datum = Utilities.formatDate(date, timezone, timestamp_format);
    cas = Utilities.formatDate(time, timezone, timestamp_format1);
    var poslednyKontaktDatum = Utilities.formatDate(sheet.getRange(index, 3).getValue(), timezone, timestamp_format2);
  } catch (error) {}
  
  var text = { "name":name, "phone":phone, "email":email, "index":index, "web":web, "ico":ico, "kategoria":kategoria, "info":info, "index":index }; 
  Logger.log(text);
  return text;
}















//=================================================================
//Vráti cislo riadku pouzitelneho kontaktu
//=================================================================
function getIndex(sheetName){  
  var sheet = ssCallPageDat.getSheetByName(sheetName);
  var index = sheet.getRange(1, 1).getValue();  
  index++; 
  var temp = sheet.getRange(index, 6).getValue();
  while (temp == "číslo neexistuje" || temp == "nemá záujem - starobný dôchodca" || temp == "klient" || temp == "nekontaktovať")
  {
    index++;
    var temp = sheet.getRange(index, 6).getValue();
  }
  var lastRow = sheet.getLastRow();
  if(index > lastRow)//na konci resetne zoznam
  {
    index = 2;
    sheet.getRange(1, 1).setValue(index);  
    var temp = sheet.getRange(1, 2).getValue(); temp++;
    sheet.getRange(1, 2).setValue(temp);
  }  
  sheet.getRange(1, 1).setValue(index);
 
  var name = sheet.getRange(index, 1).getValue();
  var phone = sheet.getRange(index, 2).getValue();
  var okres = sheet.getRange(index, 5).getValue();
  var vysledok = sheet.getRange(index, 6).getValue();
  var date = sheet.getRange(index, 7).getValue();
  var time = sheet.getRange(index, 8).getValue();
  var poznamka = sheet.getRange(index, 9).getValue();
  var produkty = sheet.getRange(index, 10).getValue();  
  var oblast = sheet.getRange(index, 11).getValue();
  var specifikacia = sheet.getRange(index, 12).getValue();
  var email = sheet.getRange(index, 13).getValue();
  var heslo = "-";
  

  var poslednyKontaktOsoba = sheet.getRange(index, 4).getValue();
  var poslednyKontaktStav = sheet.getRange(index, 6).getValue();
  
  //Logger.log(poslednyKontaktDatum)
  //Logger.log(poslednyKontaktOsoba)
  //Logger.log(poslednyKontaktStav)
  
  var timestamp_format = "yyy-MM-dd"; var timezone = "GMT+2";
  var timestamp_format1 = "HH:mm"; var timezone = "GMT+2"; 
  var timestamp_format2 = "dd.MM.yyy HH:mm"; var timezone = "GMT+2";
  var datum = "";
  var cas = "";
  try 
  {
    var poslednyKontaktDatum = Utilities.formatDate(sheet.getRange(index, 3).getValue(), timezone, timestamp_format2);
    datum = Utilities.formatDate(date, timezone, timestamp_format);
    cas = Utilities.formatDate(time, timezone, timestamp_format1);
  } catch (error) {}  
  
  
  var text = { "name":name, "phone":phone, "okres":okres, "vysledok":vysledok, "date":datum, "time":cas, "tempProd":produkty, "poznamka":poznamka, "oblast":oblast, "specifikacia":specifikacia, "email":email, "heslo":heslo, "index":index, "poslednyKontaktDatum":poslednyKontaktDatum, "poslednyKontaktOsoba":poslednyKontaktOsoba, "poslednyKontaktStav":poslednyKontaktStav }; 
  Logger.log(text);
  return text;
}




//=================================================================
//Vráti cislo riadku pouzitelneho kontaktu
//=================================================================
function NKgetIndex(sheetName, skratka){

  var access = checkNKaccess(skratka);

  if(access == true){  
  
    var sheet = SpreadsheetApp.openById("1mMTkZRwyw7W_w1Qm-FopzkMv2bvHvw7YgA8wCxZ_74k").getSheetByName(sheetName);
    var index = sheet.getRange(1, 1).getValue();  
    index++; 
    var temp = sheet.getRange(index, 6).getValue();
    while (temp == "číslo neexistuje" || temp == "nemá záujem - starobný dôchodca" || temp == "klient" || temp == "nekontaktovať")
    {
      index++;
      var temp = sheet.getRange(index, 6).getValue();
    }
    var lastRow = sheet.getLastRow();
    if(index > lastRow)//na konci resetne zoznam
    {
      index = 2;
      sheet.getRange(1, 1).setValue(index);  
      var temp = sheet.getRange(1, 2).getValue(); temp++;
      sheet.getRange(1, 2).setValue(temp);
    }  
    sheet.getRange(1, 1).setValue(index);
    
    var name = sheet.getRange(index, 1).getValue();
    var phone = sheet.getRange(index, 2).getValue();
    
    var okres = sheet.getRange(index, 5).getValue(); 
    
    var poznamka = sheet.getRange(index, 9).getValue();
    var email = sheet.getRange(index, 12).getValue();
    var spec = sheet.getRange(index, 13).getValue();
    var oblast = sheet.getRange(index, 14).getValue();
  
    var poslednyKontaktOsoba = sheet.getRange(index, 4).getValue();
    var poslednyKontaktStav = sheet.getRange(index, 6).getValue();
    
    var timestamp_format = "yyy-MM-dd"; var timezone = "GMT+2";
    var timestamp_format1 = "HH:mm"; var timezone = "GMT+2"; 
    var timestamp_format2 = "dd.MM.yyy HH:mm"; var timezone = "GMT+2";
    var timestamp_format3 = "dd.MM.yyy"; var timezone = "GMT+2";
  
    try 
    {
      var datNar = Utilities.formatDate(sheet.getRange(index, 10).getValue(), timezone, timestamp_format3);
      var poslednyKontaktDatum = Utilities.formatDate(sheet.getRange(index, 3).getValue(), timezone, timestamp_format2); 
    } catch (error) {} 
    
    
    var text = { "index":index, "name":name, "phone":phone, "okres":okres, "datNar":datNar, "poznamka":poznamka, "email":email, "spec":spec, "oblast":oblast, "poslednyKontaktOsoba":poslednyKontaktOsoba, "poslednyKontaktStav":poslednyKontaktStav, "poslednyKontaktDatum":poslednyKontaktDatum }
    Logger.log(text);
    return text;  
  } else return access;
}

function TeeeeeSt(){
  Logger.log(checkNKaccess("BAD"));
}



//=================================================================
//Limit nepriradených kontaktov
//=================================================================
function checkNKaccess(_skratka){
  var minCP = 15; //minimalne mnozstvo CP potrebne pre pridelenie NK
  var maxNK = 10; //maximalne mnozstvo pridelenych NK za den

  var access = false;
  var sheet = SpreadsheetApp.openById("1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak").getSheetByName("Call-Page");
  var hlavicka = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var rowLast = sheet.getLastRow() + 1;
  
  var data = sheet.getDataRange().getValues();
  
  for(var i = 0; i < data.length; i++){
    var skratka = data[i][data[0].indexOf("Skratka")];
    if(skratka == _skratka){    
      var cp = data[i][data[0].indexOf("CP")];
      var nk = data[i][data[0].indexOf("NK")];
      if(cp >= minCP && nk < maxNK){
        access = true;
      }
    }
  }
  return access;
}


//=================================================================
//Ulozi udalost do kalendara
//=================================================================
function setEventCalendar(skratka, sheetName, menoKlienta, cisloKlienta, dateTime){ 
    var cal = CalendarApp.getCalendarById("m5snno83t1vpc5nn3n82bumtjo@group.calendar.google.com");
    cal.createEvent(skratka + " " + sheetName + " " + menoKlienta + " " + cisloKlienta, minutePlus(dateTime, 0), minutePlus(dateTime, 30));
}




//=================================================================
//Ulozi kontakt
//=================================================================
function setContact(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka)
{
  var timestamp_format = "yyy.MM.dd HH:mm"; var timezone = "GMT+2";
  var timestamp_format1 = "dd.MM.yyy"; var timezone = "GMT+2";
  var timestamp_format3 = "HH:mm"; var timezone = "GMT+2";
  var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
  var date1 = Utilities.formatDate(new Date(), timezone, timestamp_format1);
  var timestamp_format2 = "yyy.MM.dd"; var timezone = "GMT+2";
  var date2 = Utilities.formatDate(new Date(), timezone, timestamp_format2);
  var dateTime = datum;
  var cas = Utilities.formatDate(new Date(dateTime), timezone, timestamp_format3);
  datum = Utilities.formatDate(new Date(dateTime), timezone, timestamp_format2);
  //------------------------------------------------
  var sheet = ssCallPageDat.getSheetByName(sheetName);
  sheet.getRange(index, 3).setValue(date);
  sheet.getRange(index, 4).setValue(volajuci);
  sheet.getRange(index, 5).setValue(sheetName);
  sheet.getRange(index, 6).setValue(vysledok);
  sheet.getRange(index, 7).setValue(datum);
  sheet.getRange(index, 8).setValue(cas); 
  sheet.getRange(index, 9).setValue(poznamka);   
  sheet.getRange(index, 10).setValue(produkty);
  sheet.getRange(index, 11).setValue(specifikacia);
  sheet.getRange(index, 12).setValue(oblast);
  sheet.getRange(index, 13).setValue(email);  
  //------------------------------------------------  
  var lastRowHist = sheetHistoria.getLastRow();//ja
  sheetHistoria.getRange(lastRowHist + 1, 1).setValue(menoKlienta);
  sheetHistoria.getRange(lastRowHist + 1, 2).setValue(cisloKlienta);
  sheetHistoria.getRange(lastRowHist + 1, 3).setValue(date);
  sheetHistoria.getRange(lastRowHist + 1, 4).setValue(volajuci);
  sheetHistoria.getRange(lastRowHist + 1, 5).setValue(sheetName);
  sheetHistoria.getRange(lastRowHist + 1, 6).setValue(vysledok);
  sheetHistoria.getRange(lastRowHist + 1, 7).setValue(datum);
  sheetHistoria.getRange(lastRowHist + 1, 8).setValue(cas); 
  sheetHistoria.getRange(lastRowHist + 1, 9).setValue(poznamka);   
  sheetHistoria.getRange(lastRowHist + 1, 10).setValue(produkty);
  sheetHistoria.getRange(lastRowHist + 1, 11).setValue(specifikacia);
  sheetHistoria.getRange(lastRowHist + 1, 12).setValue(oblast);
  sheetHistoria.getRange(lastRowHist + 1, 13).setValue(email);  
  //------------------------------------------------ 
  historiaSPI.insertRows(2);
  historiaSPI.getRange(2, 1).setValue(menoKlienta);
  historiaSPI.getRange(2, 2).setValue(cisloKlienta);
  historiaSPI.getRange(2, 3).setValue(date);
  historiaSPI.getRange(2, 4).setValue(volajuci);
  historiaSPI.getRange(2, 5).setValue(sheetName);
  historiaSPI.getRange(2, 6).setValue(vysledok);
  historiaSPI.getRange(2, 7).setValue(datum);
  historiaSPI.getRange(2, 8).setValue(cas); 
  historiaSPI.getRange(2, 9).setValue(poznamka);   
  historiaSPI.getRange(2, 10).setValue(produkty);
  historiaSPI.getRange(2, 11).setValue(specifikacia);
  historiaSPI.getRange(2, 12).setValue(oblast);
  historiaSPI.getRange(2, 13).setValue(email);  
  //------------------------------------------------ 
  
  addCP(skratka); //funkcia poznaci kolko bolo prevolanych za den kontaktov z CP

  if(volajuci == "Špilberger Lukáš" || volajuci == "Bosák Rastislav")  {
    addContactSpilbergerLukas(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka, date2, dateTime);
  }
  else if(volajuci == "Baláž Vladimír")  {
    addContactBalazVladimir(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka, date2, dateTime);
  }
  else if(volajuci == "Dzurendová Dominika")  {
    addContactDzurendovaDominika(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka, date2, dateTime);
  } 
  else   {
    addContactCentralnaDatabaza(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka, date2, dateTime);        
  }
}


//=================================================================
//Add CP - //funkcia poznaci kolko bolo prevolanych za den kontaktov z CP
//=================================================================
function addCP(_skratka){
  var sheet = SpreadsheetApp.openById("1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak").getSheetByName("Call-Page");
  var hlavicka = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var rowLast = sheet.getLastRow() + 1;
  
  var data = sheet.getDataRange().getValues();
  
  for(var i = 0; i < data.length; i++){
    var skratka = data[i][data[0].indexOf("Skratka")];
    if(skratka == _skratka){    
      var cpIndex = sheet.getRange(i + 1, data[0].indexOf("CP") + 1).getValue();
      cpIndex++;
      sheet.getRange(i + 1, data[0].indexOf("CP") + 1).setValue(cpIndex);
    }
  }
}


//=================================================================
//Add NK - //funkcia poznaci kolko bolo prevolanych za den kontaktov z NK
//=================================================================
function addNK(_skratka){
  var sheet = SpreadsheetApp.openById("1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak").getSheetByName("Call-Page");
  var hlavicka = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var rowLast = sheet.getLastRow() + 1;
  
  var data = sheet.getDataRange().getValues();
  
  for(var i = 0; i < data.length; i++){
    var skratka = data[i][data[0].indexOf("Skratka")];
    if(skratka == _skratka){    
      var cpIndex = sheet.getRange(i + 1, data[0].indexOf("NK") + 1).getValue();
      cpIndex++;
      sheet.getRange(i + 1, data[0].indexOf("NK") + 1).setValue(cpIndex);
    }
  }
}





//=================================================================
//Ulozi zmeni v nepriradenom kotankte
//=================================================================
function setContactNK(okres, index, volajuci, vysledok, produkty, poznamka, datum, datNar, menoKlienta, cislo, specifikacia, oblast, email, skratka){

  var sheet = SpreadsheetApp.openById("1mMTkZRwyw7W_w1Qm-FopzkMv2bvHvw7YgA8wCxZ_74k").getSheetByName("NK");
  var hlavicka = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var rowLast = sheet.getLastRow() + 1;
  
  sheet.getRange(index, 1).setValue(menoKlienta);
  sheet.getRange(index, 2).setValue(cislo);
  sheet.getRange(index, hlavicka[0].indexOf("Dátum telefonátu") + 1).setValue(new Date());
  sheet.getRange(index, hlavicka[0].indexOf("Meno volajúceho") + 1).setValue(volajuci);
  sheet.getRange(index, hlavicka[0].indexOf("Okres") + 1).setValue(okres);
  sheet.getRange(index, hlavicka[0].indexOf("Výsledok telefonátu") + 1).setValue(vysledok);
  sheet.getRange(index, hlavicka[0].indexOf("Dátum") + 1).setValue(datum);
  sheet.getRange(index, hlavicka[0].indexOf("Poznámka") + 1).setValue(poznamka);
  sheet.getRange(index, hlavicka[0].indexOf("Dátum narodenia") + 1).setValue(datNar);
  sheet.getRange(index, hlavicka[0].indexOf("Email") + 1).setValue(email);
  sheet.getRange(index, hlavicka[0].indexOf("Zárobok - špecifikácia") + 1).setValue(specifikacia);
  sheet.getRange(index, hlavicka[0].indexOf("Zarobok -  Oblasť") + 1).setValue(oblast);
  
  addNK(skratka);

  if(vysledok == "kontaktovať inokedy")  
  {   
    var shBAD = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Databáza"); 
    var hlavicka = shBAD.getRange(1, 1, 1, shBAD.getLastColumn()).getValues();
    var rowLast = shBAD.getLastRow() + 1;
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(new Date());
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Odp.") + 1).setValue("NK");
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Skr.") + 1).setValue(skratka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("O.") + 1).setValue(okres);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(menoKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Mobil") + 1).setValue(cislo);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Úspešný kontakt") + 1).setValue(new Date());
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum akcie") + 1).setValue(datum);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Proces") + 1).setValue("2. Termín (48h)");      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Špecifikácia") + 1).setValue(specifikacia);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Oblasť") + 1).setValue(oblast);      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue("Produkty: " + produkty + ". Poznámka: " + poznamka);  
    shBAD.getRange(rowLast, hlavicka[0].indexOf("ID") + 1).setValue(menoKlienta + "(" + uniqueID() + ")");
 
  }
  if(vysledok == "dohodnuté stretnutie")
  {
    var shBAD = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Databáza");
    var hlavicka = shBAD.getRange(1, 1, 1, shBAD.getLastColumn()).getValues();
    var rowLast = shBAD.getLastRow() + 1;
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(new Date());
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Odp.") + 1).setValue("NK");
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Skr.") + 1).setValue(skratka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("O.") + 1).setValue(okres);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(menoKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Mobil") + 1).setValue(cislo);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Úspešný kontakt") + 1).setValue(new Date());
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum akcie") + 1).setValue(datum);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Proces") + 1).setValue("2. Termín (48h)");      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Špecifikácia") + 1).setValue(specifikacia);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Oblasť") + 1).setValue(oblast);      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue("TERMIN " + datum + "Produkty: " + produkty + ". Poznámka: " + poznamka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("ID") + 1).setValue(menoKlienta + "(" + uniqueID() + ")");
    
    //setEventCalendar(skratka, sheetName, menoKlienta, cisloKlienta, dateTime); 
  }  

}






//=================================================================
//Prida kontakt do databázy Badinka Peter
//=================================================================
function addContactCentralnaDatabaza(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka, date2, dateTime)
{
  if(vysledok == "kontaktovať inokedy")  
  {   
    var shBAD = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Databáza"); 
    var hlavicka = shBAD.getRange(1, 1, 1, shBAD.getLastColumn()).getValues();
    var rowLast = shBAD.getLastRow() + 1;
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(date2);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Odp.") + 1).setValue("B1");
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Skr.") + 1).setValue(skratka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("O.") + 1).setValue(sheetName);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(menoKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Mobil") + 1).setValue(cisloKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Úspešný kontakt") + 1).setValue(date2);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum akcie") + 1).setValue(datum);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Proces") + 1).setValue("1. Dohodni");      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Špecifikácia") + 1).setValue(specifikacia);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Oblasť") + 1).setValue(oblast);      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue("Produkty: " + produkty + ". Poznámka: " + poznamka);  
    shBAD.getRange(rowLast, hlavicka[0].indexOf("ID") + 1).setValue(menoKlienta + "(" + uniqueID() + ")");
 
  }
  if(vysledok == "dohodnuté stretnutie")
  {
    var shBAD = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Databáza");
    var hlavicka = shBAD.getRange(1, 1, 1, shBAD.getLastColumn()).getValues();
    var rowLast = shBAD.getLastRow() + 1;
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(date2);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Odp.") + 1).setValue("B1");
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Skr.") + 1).setValue(skratka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("O.") + 1).setValue(sheetName);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(menoKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Mobil") + 1).setValue(cisloKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Úspešný kontakt") + 1).setValue(date2);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum akcie") + 1).setValue(datum);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Proces") + 1).setValue("2. Termín (48h)");      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Špecifikácia") + 1).setValue(specifikacia);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Oblasť") + 1).setValue(oblast);      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue("TERMIN " + datum + " o " + cas +". " + "Produkty: " + produkty + ". Poznámka: " + poznamka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("ID") + 1).setValue(menoKlienta + "(" + uniqueID() + ")");
    
    //setEventCalendar(skratka, sheetName, menoKlienta, cisloKlienta, dateTime); 
  }  
}




//=================================================================
//Prida kontakt do databázy Badinka Peter
//=================================================================
function addContactBadinkaPeter(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka, date2, dateTime)
{
  if(vysledok == "kontaktovať inokedy")  
  {
    var sheetBadAktivity = ss.getSheetByName("Aktivity");    
    var shBAD = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Databáza"); 
    var hlavicka = shBAD.getRange(1, 1, 1, shBAD.getLastColumn()).getValues();
    var rowLast = shBAD.getLastRow() + 1;
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(date2);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Odp.") + 1).setValue("B1");
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Skr.") + 1).setValue(skratka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("O.") + 1).setValue(sheetName);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(menoKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Mobil") + 1).setValue(cisloKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Úspešný kontakt") + 1).setValue(date2);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum akcie") + 1).setValue(datum);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Proces") + 1).setValue("1. Dohodni");      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Špecifikácia") + 1).setValue(specifikacia);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Oblasť") + 1).setValue(oblast);      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue("Produkty: " + produkty + ". Poznámka: " + poznamka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("ID") + 1).setValue(menoKlienta + "(" + uniqueID() + ")");
    sheetBadAktivity.getRange(2, 2).setValue(sheetBadAktivity.getRange(2, 2).getValue() + 1);
  }
  if(vysledok == "dohodnuté stretnutie")
  {
    var sheetBadAktivity = ss.getSheetByName("Aktivity");
    var shBAD = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Databáza");
    var hlavicka = shBAD.getRange(1, 1, 1, shBAD.getLastColumn()).getValues();
    var rowLast = shBAD.getLastRow() + 1;
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(date2);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Odp.") + 1).setValue("B1");
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Skr.") + 1).setValue(skratka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("O.") + 1).setValue(sheetName);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(menoKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Mobil") + 1).setValue(cisloKlienta);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Úspešný kontakt") + 1).setValue(date2);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Dátum akcie") + 1).setValue(datum);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Proces") + 1).setValue("2. Termín (48h)");      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Špecifikácia") + 1).setValue(specifikacia);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Oblasť") + 1).setValue(oblast);      
    shBAD.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue("TERMIN " + datum + " o " + cas +". " + "Produkty: " + produkty + ". Poznámka: " + poznamka);
    shBAD.getRange(rowLast, hlavicka[0].indexOf("ID") + 1).setValue(menoKlienta + "(" + uniqueID() + ")");
    sheetBadAktivity.getRange(2, 2).setValue(sheetBadAktivity.getRange(2, 2).getValue() + 1);
    sheetBadAktivity.getRange(3, 2).setValue(sheetBadAktivity.getRange(3, 2).getValue() + 1);
    
    //setEventCalendar(skratka, sheetName, menoKlienta, cisloKlienta, dateTime); 
    
    if(volajuci == "Badinka Peter"){
      var cal = CalendarApp.getCalendarById("peter.badinka44@gmail.com");
      var idEvent =  cal.createEvent(skratka + " " + sheetName + " " + menoKlienta + " " + cisloKlienta, minutePlus(dateTime, 0), minutePlus(dateTime, 30)).getId();
      cal.getEventById(idEvent).setDescription("<b>Produkty:</b> " + produkty + "<br>" + "<b>Poznámka:</b> " + poznamka);    
    }
  }
}


//=================================================================
//Prida kontakt do databázy Spilberger Lukas
//=================================================================
function addContactSpilbergerLukas(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka, date2, dateTime)
{
  if(vysledok == "kontaktovať inokedy")  
  {
    sheetSPI.insertRows(2);
    sheetSPI.getRange(2, 1).setValue(date2);
    sheetSPI.getRange(2, 2).setValue("B1");
    sheetSPI.getRange(2, 3).setValue(skratka);
    sheetSPI.getRange(2, 4).setValue(sheetName);
    sheetSPI.getRange(2, 5).setValue(menoKlienta);
    sheetSPI.getRange(2, 6).setValue(cisloKlienta);
    sheetSPI.getRange(2, 8).setValue(date2);
    sheetSPI.getRange(2, 10).setValue(datum); 
    sheetSPI.getRange(2, 16).setValue(email); 
    sheetSPI.getRange(2, 17).setValue(specifikacia);
    sheetSPI.getRange(2, 18).setValue(oblast);
    sheetSPI.getRange(2, 11).setValue("Produkty: " + produkty + ". Poznámka: " + poznamka);
  }
  if(vysledok == "dohodnuté stretnutie")
  {
    sheetSPI.insertRows(2);
    sheetSPI.getRange(2, 1).setValue(date2);
    sheetSPI.getRange(2, 2).setValue("B1");
    sheetSPI.getRange(2, 3).setValue(skratka);
    sheetSPI.getRange(2, 4).setValue(sheetName);
    sheetSPI.getRange(2, 5).setValue(menoKlienta);
    sheetSPI.getRange(2, 6).setValue(cisloKlienta);
    sheetSPI.getRange(2, 8).setValue(date2);
    sheetSPI.getRange(2, 10).setValue(datum);   
    sheetSPI.getRange(2, 16).setValue(email);
    sheetSPI.getRange(2, 17).setValue(specifikacia);
    sheetSPI.getRange(2, 18).setValue(oblast);
    sheetSPI.getRange(2, 11).setValue("TERMIN "+ datum + " o " + cas +". " + "Produkty: " + produkty + ". Poznámka: " + poznamka);
    //setEventCalendar(skratka, sheetName, menoKlienta, cisloKlienta, dateTime);
  }
}


//=================================================================
//Prida kontakt do databázy Balaz Vladimir
//=================================================================
function addContactBalazVladimir(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka, date2, dateTime)
{
  var ssBAL = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1_w8eQIA03i8QJZz5DDGOATtI5tdISypmVsZOpb4xhYY/edit#gid=52422594");
  var sheetBAL = ssBAL.getSheetByName("ZH - ZV");
  if(vysledok == "kontaktovať inokedy")  
  {
    sheetBAL.insertRows(2);
    sheetBAL.getRange(2, 1).setValue(date2);
    sheetBAL.getRange(2, 2).setValue("B1");
    sheetBAL.getRange(2, 3).setValue(skratka);
    sheetBAL.getRange(2, 4).setValue(sheetName);
    sheetBAL.getRange(2, 5).setValue(menoKlienta);
    sheetBAL.getRange(2, 6).setValue(cisloKlienta);
    sheetBAL.getRange(2, 7).setValue(date2);
    sheetBAL.getRange(2, 9).setValue(datum);   
    sheetBAL.getRange(2, 16).setValue(email);
    sheetBAL.getRange(2, 17).setValue(specifikacia);
    sheetBAL.getRange(2, 18).setValue(oblast);
    sheetBAL.getRange(2, 10).setValue("Produkty: " + produkty + ". Poznámka: " + poznamka);
  }
  if(vysledok == "dohodnuté stretnutie")
  {
    sheetBAL.insertRows(2);
    sheetBAL.getRange(2, 1).setValue(date2);
    sheetBAL.getRange(2, 2).setValue("B1");
    sheetBAL.getRange(2, 3).setValue(skratka);
    sheetBAL.getRange(2, 4).setValue(sheetName);
    sheetBAL.getRange(2, 5).setValue(menoKlienta);
    sheetBAL.getRange(2, 6).setValue(cisloKlienta);
    sheetBAL.getRange(2, 7).setValue(date2);
    sheetBAL.getRange(2, 9).setValue(datum);   
    sheetBAL.getRange(2, 16).setValue(email);
    sheetBAL.getRange(2, 17).setValue(specifikacia);
    sheetBAL.getRange(2, 18).setValue(oblast);
    sheetBAL.getRange(2, 10).setValue("TERMIN "+ datum + " o " + cas +". " + "Produkty: " + produkty + ". Poznámka: " + poznamka);
    //setEventCalendar(skratka, sheetName, menoKlienta, cisloKlienta, dateTime);
  }
}


//=================================================================
//Prida kontakt do databázy Dzurendova Dominika
//=================================================================
function addContactDzurendovaDominika(sheetName, index, volajuci, vysledok, produkty, poznamka, datum, cas, heslo, menoKlienta, cisloKlienta, specifikacia, oblast, email, databaza, skratka, date2, dateTime)
{
  var ssDZU = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1eLdHOdoZY0vGvPLvjYVroioizcOw3RfrkrymnAZNPgA/edit#gid=788999471");
  var sheetDZU = ssDZU.getSheetByName("Databáza");
  if(vysledok == "kontaktovať inokedy")  
  {
    var lastRow = sheetDZU.getLastRow();
    sheetDZU.getRange(lastRow + 1, 1).setValue(date2);
    sheetDZU.getRange(lastRow + 1, 2).setValue("B1");
    sheetDZU.getRange(lastRow + 1, 3).setValue(skratka);
    sheetDZU.getRange(lastRow + 1, 4).setValue(sheetName);
    sheetDZU.getRange(lastRow + 1, 5).setValue(menoKlienta);
    sheetDZU.getRange(lastRow + 1, 6).setValue(cisloKlienta);
    sheetDZU.getRange(lastRow + 1, 9).setValue(date2);
    sheetDZU.getRange(lastRow + 1, 11).setValue(datum);   
    sheetDZU.getRange(lastRow + 1, 16).setValue(email);
    sheetDZU.getRange(lastRow + 1, 17).setValue(specifikacia);
    sheetDZU.getRange(lastRow + 1, 18).setValue(oblast);
    sheetDZU.getRange(lastRow + 1, 12).setValue("Produkty: " + produkty + ". Poznámka: " + poznamka);
  }
  if(vysledok == "dohodnuté stretnutie")
  {
    var lastRow = sheetDZU.getLastRow();
    sheetDZU.getRange(lastRow + 1, 1).setValue(date2);
    sheetDZU.getRange(lastRow + 1, 2).setValue("B1");
    sheetDZU.getRange(lastRow + 1, 3).setValue(skratka);
    sheetDZU.getRange(lastRow + 1, 4).setValue(sheetName);
    sheetDZU.getRange(lastRow + 1, 5).setValue(menoKlienta);
    sheetDZU.getRange(lastRow + 1, 6).setValue(cisloKlienta);
    sheetDZU.getRange(lastRow + 1, 9).setValue(date2);
    sheetDZU.getRange(lastRow + 1, 11).setValue(datum);   
    sheetDZU.getRange(lastRow + 1, 16).setValue(email);
    sheetDZU.getRange(lastRow + 1, 17).setValue(specifikacia);
    sheetDZU.getRange(lastRow + 1, 18).setValue(oblast);
    sheetDZU.getRange(lastRow + 1, 12).setValue("TERMIN "+ datum + " o " + cas +". " + "Produkty: " + produkty + ". Poznámka: " + poznamka);
    //setEventCalendar(skratka, sheetName, menoKlienta, cisloKlienta, dateTime);
  }
}


















//=================================================================
//Nacita kontakt podla telefonneho ceisla
//=================================================================
function searchContact(sheetName, phone)
{
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1mAZdPPuGiD03hFQ5-sTfEKWzWJ-oxhnC3KJ3D6xtjp0/edit#gid=769662034");
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  var name, phone, okres, vysledok, date, time, poznamka, produkty, oblast, specifikacia, email, heslo, datum, cas, index;  
  
  for(var i = 1; i < data.length; i++)
  {
    if(data[i][1] == phone)
    {    
      name = data[i][1 - 1];
      phone = data[i][2 - 1];
      okres = data[i][5 - 1];
      vysledok = data[i][6 - 1];
      date = data[i][7 - 1];
      time = data[i][8 - 1];
      poznamka = data[i][9 - 1];
      produkty = data[i][10 - 1];
      oblast = data[i][11];
      specifikacia = data[i][12 - 1];
      email = data[i][13 - 1];
      heslo = "-";
      index = i + 1; 
      
      var timestamp_format = "yyy-MM-dd"; var timezone = "GMT+2";
      var timestamp_format1 = "HH:mm"; var timezone = "GMT+2";
      datum = "";
      cas = "";
      try 
      {
        datum = Utilities.formatDate(date, timezone, timestamp_format);
        cas = Utilities.formatDate(time, timezone, timestamp_format1);
      } catch (error) {}   
    
      break;
    }
  }  
  var text = { "name":name, "phone":phone, "okres":okres, "vysledok":vysledok, "date":datum, "time":cas, "tempProd":produkty, "poznamka":poznamka, "oblast":oblast, "specifikacia":specifikacia, "email":email, "heslo":heslo, "index":index };  
  return text;
  
}
























//=================================================================
//Prida novy kontakt(ak nie je duplicitny)
//=================================================================
function pridaKontakt(sheetName, name, phone, menoVolajuceho)
{
  
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1mAZdPPuGiD03hFQ5-sTfEKWzWJ-oxhnC3KJ3D6xtjp0/edit#gid=769662034");
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var duplicate = false;
  
  var timestamp_format = "yyy.MM.dd HH:mm"; var timezone = "GMT+2";
  var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
  
  for(var i = 1; i < data.length; i++)
  {
    if(data[i][1] == phone)
    {
      duplicate = true;
      break; //Kontakt je duplicitny takže sa neuloží
    }
  }
  if(duplicate == false)
  {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(name);
    sheet.getRange(lastRow + 1, 2).setValue(phone);
    sheet.getRange(lastRow + 1, 3).setValue(date);
    sheet.getRange(lastRow + 1, 4).setValue(menoVolajuceho);
  }
}


//=================================================================
//Rozpis sluzieb
//=================================================================
function rozpisSluzieb() {
  var sluzba = [];

  var sheetRozpis = SpreadsheetApp.openById("1Dtmb7T0vF94ROixdoVGk_o81bb8_qKZghJPhUpsIsB4").getSheetByName("Rozpis služieb");
  var sheetSettings = SpreadsheetApp.openById("1Dtmb7T0vF94ROixdoVGk_o81bb8_qKZghJPhUpsIsB4").getSheetByName("Settings");
  var index = sheetSettings.getRange(1, 1).getValue();
  
  var sluzbaZvolen = sheetRozpis.getRange(index, 1).getValue();
  
  var sheetRozpis = SpreadsheetApp.openById("1ltn90FiTWJ7S1qzX7HWQUp4PB3gWqmW9CDNvhen6lkg").getSheetByName("Rozpis služieb");
  var sheetSettings = SpreadsheetApp.openById("1ltn90FiTWJ7S1qzX7HWQUp4PB3gWqmW9CDNvhen6lkg").getSheetByName("Settings");
  var index = sheetSettings.getRange(1, 1).getValue();
  
  var sluzbaZiar = sheetRozpis.getRange(index, 1).getValue();
  
  var text = { "ZV":sluzbaZvolen, "ZH":sluzbaZiar }
  
  return text;
}







function test123()
{
  var shTest = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/18FQCLVkIVVd1Z9wImbUODJYS-7kYxps45Kz2BvCVIM8/edit#gid=1455322467").getSheetByName("Databáza");
  
  var testIndex = 2064;
  var testMobil = "0901234567"



  var hlavicka = shTest.getRange(1, 1, 1, shTest.getLastColumn()).getValues();
  var colMobil = hlavicka[0].indexOf("Mobil");
  var dataMobil = shTest.getRange(1, colMobil + 1, shTest.getLastRow(), 1).getValues();
  var lastRow = shTest.getLastRow();
  
  Logger.log(dataMobil)
  
  for(var i = 0; i < lastRow; i++){
    var temp = dataMobil[i];
    if(temp == testMobil){
      Logger.log(i);
    }
  }
  
  //Logger.log(indexKontakt)


  //cal.createEvent("pokus pokusny", new Date("4/12/2018 08:00 AM"), new Date("4/12/2018 10:00 AM"), CalendarApp.setColor("#ff0000"))
  
}







//=================================================================
//Reset tyzdenneho planu
//=================================================================
function resetPlanWeek(_name){
  var shUsers = SpreadsheetApp.openById("1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak").getSheetByName("Users");
  var data = shUsers.getDataRange().getValues();
  for(var i = 2; i <= data.length; i++){
    var name = shUsers.getRange(i, 3).getValue();
    if(name == _name){
      shUsers.getRange(i, 6).setValue(10);
      shUsers.getRange(i, 7).setValue(5);
      shUsers.getRange(i, 8).setValue(10);
    }
  }
}










//=================================================================
//Mal by vratit celu volanu hystoriu - zatial nefunkcne
//=================================================================
function loadHistoryData(sheetName, meno, heslo){
  var timestamp_format = "yyy-MM-dd"; var timezone = "GMT+2";
  var timestamp_format1 = "HH:mm"; var timezone = "GMT+2";
  var text = "";
  var json = [];
  var indexJson = 0;
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  for(var i = 1; i < data.length; i++){
    if(data[i][3] == meno){
      if(data[i][13] == heslo){
        var name = data[i][0];
        Logger.log(name)
        var phone = data[i][1];
        var okres = data[i][4];
        var vysledok = data[i][5];
        var date = data[i][6];
        var time = data[i][7];
        var poznamka = data[i][8];
        var produkty = data[i][9];  
        var oblast = data[i][10];
        var specifikacia = data[i][11];
        var email = data[i][12];         

        var datum = "";
        var cas = "";
        try {
          datum = Utilities.formatDate(date, timezone, timestamp_format);
          cas = Utilities.formatDate(time, timezone, timestamp_format1);
        } catch (error) {}
        
        var index = i + 1;
        
        text += {"name":name, "phone":phone, "okres":okres, "vysledok":vysledok, "date":datum,"time":cas, "tempProd":produkty, "poznamka":poznamka, "oblast":oblast, "specifikacia":specifikacia, "email":email, "heslo":heslo, "index":index}
        

        json[indexJson] = [];
        for(var j = 0; j <= 13; j++){
          pDKlienti[i][j] = data[i][j-1];
          Logger.log(data)
        }
        indexJson++; 
      }
    }
  }

  Logger.log(json[indexJson]);
}


//=================================================================
//Nastavenie noveho planu
//=================================================================
function setPlanBeb(name, plan){
  var ssAccess = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=2095307398");
  var sheetName = ssAccess.getSheetByName("Produkčné mesiace");
  var indexProdMesiac = sheetName.getRange(2, 2).getValue();
  var nameProdMesiac = sheetName.getRange(indexProdMesiac, 1).getValue();
  var ssReport = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/18B65fgu2EQ0Yd0hDZjqaMVhvIVhVdfgzdbNcIhYfbOc/edit#gid=1331903965");
  sheetName = ssReport.getSheetByName(nameProdMesiac);  
  
  var headers = sheetName.getRange(1, 1, 1, sheetName.getLastColumn()).getValues();
  var lastRow = sheetName.getLastRow();
  var colName = headers[0].indexOf("Meno spolupracovníka") + 1;
  var colPlan = headers[0].indexOf("Plán") + 1;
  
  for(var i = 1; i <= lastRow - 1; i++){
    var temp = sheetName.getRange(i, colName).getValue();
    if(temp == name){
      
       sheetName.getRange(i, colPlan).setValue(plan);
      
    }
  }
}





function pokus(){  

  nastavitPokutu("Badinka Peter", "2", ":)", 1, "test")

 
  
}




//=================================================================
//Prideliť pokutu
//=================================================================
function nastavitPokutu(pokutuDal, indexPokutuDostal, typPokuty, vyskaPokuty, poznamka){

  var shUsers = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=2095307398").getSheetByName("Users");
  var sheetPokuty = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/18B65fgu2EQ0Yd0hDZjqaMVhvIVhVdfgzdbNcIhYfbOc/edit#gid=1986721057").getSheetByName("_Pokuty");
  var hlavicka = shUsers.getRange(1, 1, 1, shUsers.getLastColumn()).getValues();
  var email = shUsers.getRange(indexPokutuDostal, hlavicka[0].indexOf("eMail") + 1).getValue();
  var pokutuDostal = shUsers.getRange(indexPokutuDostal, hlavicka[0].indexOf("Meno") + 1).getValue();
  
  var timestamp_format = "yyy.MM.dd HH:mm"; var timezone = "GMT+2";
  var dateNow = Utilities.formatDate(new Date(), timezone, timestamp_format);
  
  setPukutu(email, vyskaPokuty, 4);
  
  sheetPokuty.insertRows(2);
  sheetPokuty.getRange(2, 1).setValue(dateNow);
  sheetPokuty.getRange(2, 2).setValue(pokutuDostal);
  sheetPokuty.getRange(2, 3).setValue(typPokuty);
  sheetPokuty.getRange(2, 4).setValue(vyskaPokuty);
  sheetPokuty.getRange(2, 5).setValue(poznamka);
  sheetPokuty.getRange(2, 6).setValue(pokutuDal);
  
  
  var predmet = "Dostal si pokutu " + dateNow;
  var text;
  var text_HTML = "<p><b>Pokutu si dostal za:</b> " + typPokuty + "</p>";  
  text_HTML += "<p><b>Pokutu zapísal</b> : " + pokutuDal + "</p>";
  text_HTML += "<p><b>Výška pokuty:</b> " + vyskaPokuty + " €</p>";
  text_HTML += "<p><b>Poznámka</b> : " + poznamka + "</p>";


  MailApp.sendEmail(email, predmet, text, { htmlBody: text_HTML });
  
}


//===========================================================================
//Pridelit pokutu v1.0.0
//===========================================================================
function setPukutu(email, value, riadokPokuty)
{  
  var sheetUsers = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=2095307398").getSheetByName("Users");
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/18B65fgu2EQ0Yd0hDZjqaMVhvIVhVdfgzdbNcIhYfbOc/edit#gid=1027388018"); 
  var date = getProdukcnyMesiac(); 
  var sheetReport = ss.getSheetByName(date);
  var data = sheetUsers.getDataRange().getValues();  
  var indexEmail;
  for(var i = 2; i <= data.length; i++){
    var _email = sheetUsers.getRange(i, 2).getValue();
    if(_email == email){
      indexEmail = i;
    }    
  }
  if(indexEmail > 1){
    var pokuta = sheetReport.getRange(indexEmail, 22).getValue(); pokuta = pokuta + value; //pripocita pokutu
    sheetReport.getRange(indexEmail, 22).setValue(pokuta); //zapise pokutu     
    
    //Pokuta archív
    try 
    {      
      var timestamp_format = "yyy.MM.dd"; var timezone = "GMT+2";
      var dateNow = Utilities.formatDate(new Date(), timezone, timestamp_format);
      var meno = sheetReport.getRange(indexEmail, 1).getValue();
      
    } catch (error)
    {           
      MailApp.sendEmail("peter.badinka44@gmail.com", "Error - Pokuty, historia" , "Error - Pokuty, historia");       
    } 
  }   
}



//===========================================================================
//Get Produkcny mesiac
//===========================================================================
function getProdukcnyMesiac(){
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=2095307398");
  var prodMesiac = ss.getSheetByName("Produkčné mesiace"); 
  var index = prodMesiac.getRange(2, 2).getValue();
  var date = prodMesiac.getRange(index, 1).getValue();
  return date;
}






//===========================================================================
//Nahodny citat
//===========================================================================
function nahodnyCitat()
{ 
  var sheetCitaty = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=2095307398").getSheetByName("Citáty");
  var data = sheetCitaty.getDataRange().getValues();
  return sheetCitaty.getRange(nahodneCislo(1, data.length), 1).getValue();
}


//=================================================================
//Generovanie nahodneho cisla
//=================================================================
function nahodneCislo(numMin, numMax)
{ 
  numMax = numMax - numMin + 1;
  return Math.floor(Math.random() * numMax) + numMin;
}








//=================================================================
//Zisti oprávnenia
//=================================================================
function checkAccess(login, password, mesto)
{
  var shAccess = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=1595525522").getSheetByName("Call-Page");
  var accessData = shAccess.getDataRange().getValues();
  var lastCol = shAccess.getLastColumn(); 

  var access = false, minPlan, povMatPlan, koeficient, databaza, skratka;
  for(var i = 0; i < accessData.length; i++){
    if(accessData[i][0] == login && accessData[i][1] == password){
      minPlan = accessData[i][2];
      povMatPlan = accessData[i][3];
      koeficient = accessData[i][4];
      databaza = accessData[i][5];
      skratka = accessData[i][6];
      access = true;      
    }
  }
  
  if(access == true){ //plan, skutocnost - body
    var ssAccess = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=2095307398");
    var sheetName = ssAccess.getSheetByName("Produkčné mesiace");
    var indexProdMesiac = sheetName.getRange(2, 2).getValue();
    var nameProdMesiac = sheetName.getRange(indexProdMesiac, 1).getValue();
    var ssReport = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/18B65fgu2EQ0Yd0hDZjqaMVhvIVhVdfgzdbNcIhYfbOc/edit#gid=1331903965");
    sheetName = ssReport.getSheetByName(nameProdMesiac);  
    
    var headers = sheetName.getRange(1, 1, 1, sheetName.getLastColumn()).getValues();
    var lastRow = sheetName.getLastRow();
    var colName = headers[0].indexOf("Meno spolupracovníka") + 1;
    
    for(var i = 1; i <= lastRow - 1; i++){
      
      var temp = sheetName.getRange(i, colName).getValue();
      
      if(temp == login){
        
        var plan = sheetName.getRange(i, headers[0].indexOf("Plán") + 1).getValue();        

        var skutocnost = sheetName.getRange(i, headers[0].indexOf("Body") + 1).getValue();             
        
      }
    }
  } 



  if(access == true){ //plan, skutocnost - register, terminy  
  
    var shAccess = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=2095307398").getSheetByName("Users");
    
    for(var i = 2; i <= lastRow; i++){
      var name = shAccess.getRange(i, 3).getValue();      
      if(name == login){
        var planReg = "2";
        var planTer = "2";
      }
    }
  
  
  }
  
  

  
  var sendData = [[access, plan, skutocnost, minPlan, povMatPlan, koeficient, databaza, skratka, nahodnyCitat(), planReg, planTer]];
  
  
  //Logguje uspesne aj neuspesne prihlasenie
  var shLogging = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=2095307398").getSheetByName("Logging"); 
  var timestamp_format = "yyy.MM.dd HH:mm"; var timezone = "GMT+2";
  var dateNow = Utilities.formatDate(new Date(), timezone, timestamp_format);
  shLogging.insertRows(2);
  shLogging.getRange(2, 1).setValue(dateNow);
  shLogging.getRange(2, 2).setValue(login);
  shLogging.getRange(2, 3).setValue(access);
  
  return sendData;
}








//=================================================================
//Vytvori pole ktore prekovertuje na JSON a odosle do javascriptu
//=================================================================
function getArrayFromTables(skratka, databaza){

  if(databaza != "Špilberger Lukáš"){
    var shDat;
    if(databaza == "Badinka Peter"){
      shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Databáza");
    }
    else { //centralna databaza
      shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Databáza");
    }

    var data = shDat.getDataRange().getValues();
    var lastCol = shDat.getLastColumn() - 1;
    var poleIndex = 1;
    
    var pole = [];
    pole[0] = [];
    
    for(var j = 0; j <= lastCol; j++){
      pole[0][j] = data[0][j];
    }
    
    for(var i = 1; i < data.length; i++){
      var _skratka = data[i][data[0].indexOf("Skr.")];
      var datumAkcie = data[i][data[0].indexOf("Dátum akcie")];
      var checkDel = data[i][data[0].indexOf("Delete")];
      
      if(_skratka == skratka && datumAkcie > 0 && checkDel != "ok"){      
        pole[poleIndex] = [];
        for(var j = 0; j <= lastCol; j++){        
          pole[poleIndex][j] = data[i][j];        
        }
        poleIndex++;      
      }    
    }
    return JSON.stringify(pole);
  } 
  
  else if(databaza == "Špilberger Lukáš"){
    var shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1FDs2QVTHqdEJ6HzynpZl4hg_qFLzMRC0nPVSf_3e_5g/edit#gid=0").getSheetByName("Databáza");
    var data = shDat.getDataRange().getValues();
    var lastCol = shDat.getLastColumn() - 1;
    var poleIndex = 0;
    
    var pole = [];
    for(var i = 0; i < data.length; i++){
      var _skratka = data[i][data[0].indexOf("Získateľ")];
      var datumAkcie = data[i][data[0].indexOf("Najbližší telefonát")];
      
      if(_skratka == skratka && datumAkcie > 0){
    
        pole[poleIndex] = [];
        for(var j = 0; j <= lastCol; j++){        
          pole[poleIndex][j] = data[i][j];        
        }
        poleIndex++;      
      }    
    }
    return JSON.stringify(pole);
  }
  
  else return null;
}



function test1111(){
  getZmluvy("FER", "Badinka Peter");
}



//=================================================================
//Vytvori pole ktore prekovertuje na JSON a odosle do javascriptu - IK
//=================================================================
function getDataIK(skratka, databaza){
  var shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0/edit#gid=0").getSheetByName("IK");
  
    var data = shDat.getDataRange().getValues();
    var lastCol = shDat.getLastColumn() - 1;
    var poleIndex = 1;
    
    var pole = [];
    pole[0] = [];
    
    for(var j = 0; j <= lastCol; j++){
      pole[0][j] = data[0][j];
    }

    for(var i = 1; i < data.length; i++){
      var _skratka = data[i][data[0].indexOf("Skratka")];
      var checkDel = data[i][data[0].indexOf("Delete")];
      
      if(_skratka == skratka && checkDel != "ok"){      
        pole[poleIndex] = [];
        for(var j = 0; j <= lastCol; j++){        
          pole[poleIndex][j] = data[i][j];        
        }
        poleIndex++;      
      }      
    }
    return JSON.stringify(pole);
}



//=================================================================
//Vytvori pole ktore prekovertuje na JSON a odosle do javascriptu - Dohody
//=================================================================
function getDataDohody(skratka){
  var shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0/edit#gid=0").getSheetByName("Dohody");
  
    var data = shDat.getDataRange().getValues();
    var lastCol = shDat.getLastColumn() - 1;
    var poleIndex = 1;
    
    var pole = [];
    pole[0] = [];
    
    for(var j = 0; j <= lastCol; j++){
      pole[0][j] = data[0][j];
    }

    for(var i = 1; i < data.length; i++){
      var _skratka = data[i][data[0].indexOf("Skratka")];
      var checkDel = data[i][data[0].indexOf("Delete")];
      
      if(_skratka == skratka && checkDel != "ok"){      
        pole[poleIndex] = [];
        for(var j = 0; j <= lastCol; j++){        
          pole[poleIndex][j] = data[i][j];        
        }
        poleIndex++;      
      }      
    }
    return JSON.stringify(pole);
}




//=================================================================
//Vytvori pole ktore prekovertuje na JSON a odosle do javascriptu - zmluvy
//=================================================================
function getZmluvy(skratka, databaza){
  if(databaza == "Badinka Peter" || databaza == "Centrálna databáza"){
    var shDat;
    if(databaza == "Badinka Peter"){
      shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Zmluvy");
    }
    if(databaza == "Centrálna databáza"){
      shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Zmluvy");
    }

    var data = shDat.getDataRange().getValues();
    var lastCol = shDat.getLastColumn() - 1;
    var poleIndex = 1;
    
    var pole = [];
    pole[0] = [];
    
    for(var j = 0; j <= lastCol; j++){
      pole[0][j] = data[0][j];
    }

    for(var i = 1; i < data.length; i++){
      var _skratka = data[i][data[0].indexOf("Skr.")];
      var checkDel = data[i][data[0].indexOf("Delete")];
      
      if(_skratka == skratka && checkDel != "ok"){      
        pole[poleIndex] = [];
        for(var j = 0; j <= lastCol; j++){        
          pole[poleIndex][j] = data[i][j];        
        }
        poleIndex++;      
      }      
    }
    return JSON.stringify(pole);
  } 
}


//=================================================================
//Load data from Call-Page
//=================================================================
function getCallPageData(login, okres){
  var shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1mAZdPPuGiD03hFQ5-sTfEKWzWJ-oxhnC3KJ3D6xtjp0/edit#gid=1884595478").getSheetByName(okres);
  var data = shDat.getDataRange().getValues();
  var lastCol = shDat.getLastColumn() - 1;
  var poleIndex = 0;
    
    var pole = [];
    for(var i = 0; i < data.length; i++){
      var _skratka = data[i][data[0].indexOf("Meno volajúceho")];
      var datumAkcie = data[i][data[0].indexOf("Dátum")];
      var vysledok = data[i][data[0].indexOf("Výsledok telefonátu")];
      
      if(_skratka == login && datumAkcie > 0 && vysledok == "kontaktovať inokedy - CallPage"){      
        pole[poleIndex] = [];
        for(var j = 0; j <= lastCol; j++){        
          pole[poleIndex][j] = data[i][j];        
        }
        pole[poleIndex][lastCol + 1] = i + 1;
        //Logger.log(pole[poleIndex][4]);
        poleIndex++;      
      }    
    }
    return JSON.stringify(pole);
}


function test1145411(){
  getCallPageData("Badinka Peter", "ZV");
}


//=================================================================
//Vrati ine osoby
//=================================================================
function getIneOsoby(skratka){
  var shDat = SpreadsheetApp.openById("1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o").getSheetByName("Iné osoby");
  
  var data = shDat.getDataRange().getValues();
  var lastCol = shDat.getLastColumn() - 1;
  var poleIndex = 1;
  
  var pole = [];
  pole[0] = [];
  
  for(var j = 0; j <= lastCol; j++){
    pole[0][j] = data[0][j];
  }
  
  for(var i = 1; i < data.length; i++){
    var _skratka = data[i][data[0].indexOf("Skr.")];
    var checkDel = data[i][data[0].indexOf("Delete")];
    
    if(_skratka == skratka && checkDel != "ok"){      
      pole[poleIndex] = [];
      for(var j = 0; j <= lastCol; j++){        
        pole[poleIndex][j] = data[i][j];        
      }
      poleIndex++;      
    }      
  }
  return JSON.stringify(pole);
} 










function test1112(){
  //saveData("Badinka Peter", "0915851509", "", "", "BAD", "");
  var shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Databáza");
  var last = shDat.getLastRow();
  Logger.log(last)
}



function generateID(){
  for(var i = 1; i <= 1000; i++){
    Logger.log(uniqueID())
  }
}








//=================================================================
//Ulozi data cez pole
//=================================================================
function saveDataNew(dataCallPage, indexID, databaza){
  var shDat;
  var shDatZ;
  if(databaza == "Badinka Peter"){
    shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Databáza");
    shDatZ = SpreadsheetApp.openById("1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I").getSheetByName("Zmluvy");
  }
  if(databaza == "Centrálna databáza"){
    shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Databáza");
    shDatZ = SpreadsheetApp.openById("1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o").getSheetByName("Zmluvy");
  }  
  var data = shDat.getDataRange().getValues();
  var json = JSON.parse(dataCallPage);

  for(var i = 0; i < data.length; i++){
    var tempID = data[i][data[0].indexOf("ID")];
    if(tempID == json[indexID]){    
      for(var j = 0; j < json.length; j++){
          shDat.getRange(i + 1, j + 1).setValue(json[j]);
      }
    }
  }  

  var data = shDatZ.getDataRange().getValues();
  var header = shDat.getRange(1, 1, 1, shDat.getLastColumn()).getValues();
  
  for(var i = 1; i < data.length; i++){
    var tempID = tempID = data[i][data[0].indexOf("ID - Osoba")];
    if(tempID == json[indexID]){
      shDatZ.getRange(i + 1, data[0].indexOf("Mobil") + 1).setValue(json[header[0].indexOf("Mobil")]);
      shDatZ.getRange(i + 1, data[0].indexOf("Email") + 1).setValue(json[header[0].indexOf("Email")]);
    }    
  }  
}



//=================================================================
//Ulozi data IO cez pole
//=================================================================
function saveDataIO(dataIO, indexID){
  var shDat = SpreadsheetApp.openById("1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o").getSheetByName("Iné osoby");

  var data = shDat.getDataRange().getValues();
  var json = JSON.parse(dataIO);

  for(var i = 0; i < data.length; i++){
    var tempID = data[i][data[0].indexOf("ID - Iná osoba")];
    if(tempID == json[indexID]){    
      for(var j = 0; j < json.length; j++){
          shDat.getRange(i + 1, j + 1).setValue(json[j]);
      }
    }
  }
}



//=================================================================
//Ulozi data IO cez pole
//=================================================================
function saveDataIK(dataIK, indexID){
  var shDat = SpreadsheetApp.openById("1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0").getSheetByName("IK");

  var data = shDat.getDataRange().getValues();
  var json = JSON.parse(dataIK);

  for(var i = 0; i < data.length; i++){
    var tempID = data[i][data[0].indexOf("ID")];
    if(tempID == json[indexID]){    
      for(var j = 0; j < json.length; j++){
          shDat.getRange(i + 1, j + 1).setValue(json[j]);
      }
    }
  }
}



//=================================================================
//Ulozi data Dohody cez pole
//=================================================================
function saveDataDohoda(dataDohoda, indexID){
  var shDat = SpreadsheetApp.openById("1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0").getSheetByName("Dohody");

  var data = shDat.getDataRange().getValues();
  var json = JSON.parse(dataDohoda);

  for(var i = 0; i < data.length; i++){
    var tempID = data[i][data[0].indexOf("ID - Dohoda")];
    if(tempID == json[indexID]){    
      for(var j = 0; j < json.length; j++){
          shDat.getRange(i + 1, j + 1).setValue(json[j]);
      }
    }
  }
}





//=================================================================
//Prida IO do databazy
//=================================================================
function addToDat(dataIO){
  var shDat = SpreadsheetApp.openById("1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o").getSheetByName("Databáza");

  var header = shDat.getRange(1, 1, 1, shDat.getLastColumn()).getValues();
  var lastRow = shDat.getLastRow();
  var json = JSON.parse(dataIO);

  for(var j = 0; j < json.length - 2; j++){ // -2 aby sa zbytocne neukladalo stare id(v dolnej casti sa vytvara nove)
    shDat.getRange(lastRow + 1, j + 1).setValue(json[j]);
  }
  
  var meno = shDat.getRange(lastRow + 1, header[0].indexOf("Priezvisko a meno") + 1).getValue();
  
  shDat.getRange(lastRow + 1, header[0].indexOf("ID") + 1).setValue(meno + "(" + uniqueID() + ")");
}



//=================================================================
//Ulozi data
//=================================================================
function saveData(databaza, selectNumber, dateReg, typOdp, ziskatel, pobocka, meno, mobil, datumTelefonatu, stav, datumNajblTelefonatu, poznamka, status, email, povolanie,
                                          ulica, obec, datNarodenia, rodneCislo, cisloOp, platnostOpOd, platnostOpDo, prijmy, vydavky, dokladVydal, cisloUctu, lekarMeno, lekarAdresa,
                                          lekarKontakt, vyska, hmotnost){
    
  if(databaza == "Špilberger Lukáš"){
    var shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1FDs2QVTHqdEJ6HzynpZl4hg_qFLzMRC0nPVSf_3e_5g/edit#gid=0").getSheetByName("Databáza");
    var lastRow = shDat.getLastRow();
    var colPhone = shDat.getRange(1, 6, lastRow, 1).getValues();
    var colZiskatel = shDat.getRange(1, 3, lastRow, 1).getValues();
    for(var i = 1; i <= colPhone.length; i++){
      if(colPhone[i - 1] == selectNumber && colZiskatel[i - 1] == ziskatel){       
        shDat.getRange(i, 1).setValue(dateReg);
        shDat.getRange(i, 2).setValue(typOdp);
        shDat.getRange(i, 3).setValue(ziskatel);
        shDat.getRange(i, 4).setValue(pobocka);
        shDat.getRange(i, 5).setValue(meno);
        shDat.getRange(i, 6).setValue(mobil);
        shDat.getRange(i, 8).setValue(datumTelefonatu);
        shDat.getRange(i, 9).setValue(stav);
        shDat.getRange(i, 10).setValue(datumNajblTelefonatu);
        shDat.getRange(i, 11).setValue(poznamka);
        shDat.getRange(i, 12).setValue(status);        
        shDat.getRange(i, 16).setValue(email);
        shDat.getRange(i, 17).setValue(povolanie);
        return null;
      }
    }
  }
  if(databaza == "Badinka Peter" || databaza == "Centrálna databáza"){  
    var shDat;
    if(databaza == "Badinka Peter"){
      shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Databáza");
    }
    if(databaza == "Centrálna databáza"){
      shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Databáza");
    } 
    var lastRow = shDat.getLastRow();
    var colPhone = shDat.getRange(1, 7, lastRow, 1).getValues();
    var colZiskatel = shDat.getRange(1, 3, lastRow, 1).getValues();
    var hlavicka = shDat.getRange(1, 1, 1, shDat.getLastColumn()).getValues();
    for(var i = 1; i <= colPhone.length; i++){
      var menoKl = shDat.getRange(i, hlavicka[0].indexOf("Priezvisko a meno") + 1).getValue();
      if(colPhone[i - 1] == selectNumber && colZiskatel[i - 1] == ziskatel && menoKl == meno){  
        shDat.getRange(i, hlavicka[0].indexOf("Dátum") + 1).setValue(dateReg);
        shDat.getRange(i, hlavicka[0].indexOf("Odp.") + 1).setValue(typOdp);
        shDat.getRange(i, hlavicka[0].indexOf("Skr.") + 1).setValue(ziskatel);
        shDat.getRange(i, hlavicka[0].indexOf("O.") + 1).setValue(pobocka);
        shDat.getRange(i, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(meno);
        shDat.getRange(i, hlavicka[0].indexOf("Mobil") + 1).setValue(mobil);
        shDat.getRange(i, hlavicka[0].indexOf("Úspešný kontakt") + 1).setValue(datumTelefonatu);
        shDat.getRange(i, hlavicka[0].indexOf("Proces") + 1).setValue(stav);
        shDat.getRange(i, hlavicka[0].indexOf("Dátum akcie") + 1).setValue(datumNajblTelefonatu);
        shDat.getRange(i, hlavicka[0].indexOf("Poznámka") + 1).setValue(poznamka);
        shDat.getRange(i, hlavicka[0].indexOf("Android Sync") + 1).setValue(status);        
        shDat.getRange(i, hlavicka[0].indexOf("Email") + 1).setValue(email);
        shDat.getRange(i, hlavicka[0].indexOf("Špecifikácia") + 1).setValue(povolanie); 
        shDat.getRange(i, hlavicka[0].indexOf("Ulica") + 1).setValue(ulica);
        shDat.getRange(i, hlavicka[0].indexOf("PSČ / Obec") + 1).setValue(obec);
        shDat.getRange(i, hlavicka[0].indexOf("Dátum narodenia") + 1).setValue(datNarodenia);
        shDat.getRange(i, hlavicka[0].indexOf("Rodné číslo") + 1).setValue(rodneCislo);
        shDat.getRange(i, hlavicka[0].indexOf("Číslo OP") + 1).setValue(cisloOp);
        shDat.getRange(i, hlavicka[0].indexOf("Platnosť od") + 1).setValue(platnostOpOd);
        shDat.getRange(i, hlavicka[0].indexOf("Platnosť do") + 1).setValue(platnostOpDo);
        shDat.getRange(i, hlavicka[0].indexOf("Doklad vydal") + 1).setValue(dokladVydal);
        shDat.getRange(i, hlavicka[0].indexOf("Príjmy") + 1).setValue(prijmy);
        shDat.getRange(i, hlavicka[0].indexOf("Výdavky") + 1).setValue(vydavky);
        shDat.getRange(i, hlavicka[0].indexOf("Číslo účtu") + 1).setValue(cisloUctu);
        shDat.getRange(i, hlavicka[0].indexOf("Obvodný lekár - Meno") + 1).setValue(lekarMeno);
        shDat.getRange(i, hlavicka[0].indexOf("Obvodný lekár - Adresa") + 1).setValue(lekarAdresa);
        shDat.getRange(i, hlavicka[0].indexOf("Obvodný lekár - Kontakt") + 1).setValue(lekarKontakt);
        shDat.getRange(i, hlavicka[0].indexOf("Výška") + 1).setValue(vyska);
        shDat.getRange(i, hlavicka[0].indexOf("Hmotnosť") + 1).setValue(hmotnost);
        return null;
      }
    }
  }
}



//=================================================================
//Ulozi zmluvy
//=================================================================
function saveDataZmluvy(dataCallPage, indexID, databaza){
  var shDat;
  if(databaza == "Badinka Peter"){
    shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Zmluvy");
  }
  if(databaza == "Centrálna databáza"){
    shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Zmluvy");
  }  
  var data = shDat.getDataRange().getValues();
  var json = JSON.parse(dataCallPage);

  for(var i = 0; i < data.length; i++){
    var tempID = data[i][data[0].indexOf("ID - Zmluva")];
    if(tempID == json[indexID]){    
      for(var j = 0; j < json.length; j++){
          shDat.getRange(i + 1, j + 1).setValue(json[j]);
      }
    }
  }
}





//=================================================================
//Prida data do hlavnej databazy
//=================================================================
function addData(databaza, selectNumber, dateReg, typOdp, ziskatel, pobocka, meno, mobil, datumTelefonatu, stav, datumNajblTelefonatu, poznamka, status, email, povolanie,
    ulica, obec, datNarodenia, miestoNarodenia, rodneCislo, cisloOp, platnostOpOd, platnostOpDo, prijmy, vydavky, dokladVydal, cisloUctu, lekarMeno, lekarAdresa, lekarKontakt, vyska, hmotnost, oslovenie,
    priezvisko, krstneMeno, meniny, interval, androidSync){
    
  if(databaza == "Špilberger Lukáš"){
    var shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1FDs2QVTHqdEJ6HzynpZl4hg_qFLzMRC0nPVSf_3e_5g/edit#gid=0").getSheetByName("Databáza");      
    shDat.insertRows(2);       
    shDat.getRange(2, 1).setValue(dateReg);
    shDat.getRange(2, 2).setValue(typOdp);
    shDat.getRange(2, 3).setValue(ziskatel);
    shDat.getRange(2, 4).setValue(pobocka);
    shDat.getRange(2, 5).setValue(meno);
    shDat.getRange(2, 6).setValue(mobil);
    shDat.getRange(2, 8).setValue(datumTelefonatu);
    shDat.getRange(2, 9).setValue(stav);
    shDat.getRange(2, 10).setValue(datumNajblTelefonatu);
    shDat.getRange(2, 11).setValue(poznamka);
    shDat.getRange(2, 12).setValue(status);        
    shDat.getRange(2, 16).setValue(email);
    shDat.getRange(2, 17).setValue(povolanie);
    return null;
  }
  if(databaza == "Badinka Peter" || databaza == "Centrálna databáza"){
    var shDat;
    if(databaza == "Badinka Peter"){
      shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Databáza");
    }
    if(databaza == "Centrálna databáza"){
      shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Databáza");
    }       
    var rowLast = shDat.getLastRow() + 1; 
    var hlavicka = shDat.getRange(1, 1, 1, shDat.getLastColumn()).getValues();
    shDat.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(dateReg);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Odp.") + 1).setValue(typOdp);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Skr.") + 1).setValue(ziskatel);
    shDat.getRange(rowLast, hlavicka[0].indexOf("O.") + 1).setValue(pobocka);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(meno);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Mobil") + 1).setValue(mobil);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Úspešný kontakt") + 1).setValue(datumTelefonatu);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Proces") + 1).setValue(stav);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Dátum akcie") + 1).setValue(datumNajblTelefonatu);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue(poznamka);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Android Sync") + 1).setValue(status);        
    shDat.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Špecifikácia") + 1).setValue(povolanie); 
    shDat.getRange(rowLast, hlavicka[0].indexOf("Ulica") + 1).setValue(ulica);
    shDat.getRange(rowLast, hlavicka[0].indexOf("PSČ / Obec") + 1).setValue(obec);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Dátum narodenia") + 1).setValue(datNarodenia);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Miesto narodenia") + 1).setValue(miestoNarodenia);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Rodné číslo") + 1).setValue(rodneCislo);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Číslo OP") + 1).setValue(cisloOp);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Platnosť od") + 1).setValue(platnostOpOd);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Platnosť do") + 1).setValue(platnostOpDo);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Doklad vydal") + 1).setValue(dokladVydal);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Príjmy") + 1).setValue(prijmy);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Výdavky") + 1).setValue(vydavky);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Číslo účtu") + 1).setValue(cisloUctu);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Obvodný lekár - Meno") + 1).setValue(lekarMeno);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Obvodný lekár - Adresa") + 1).setValue(lekarAdresa);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Obvodný lekár - Kontakt") + 1).setValue(lekarKontakt);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Výška") + 1).setValue(vyska);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Hmotnosť") + 1).setValue(hmotnost);
    
    shDat.getRange(rowLast, hlavicka[0].indexOf("Priezvisko") + 1).setValue(priezvisko);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Meno") + 1).setValue(krstneMeno);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Oslovenie") + 1).setValue(oslovenie);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Narodeniny / meniny") + 1).setValue(meniny);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Servisný email") + 1).setValue(interval);
    shDat.getRange(rowLast, hlavicka[0].indexOf("Android Sync") + 1).setValue(androidSync);
    
    shDat.getRange(rowLast, hlavicka[0].indexOf("ID") + 1).setValue(meno + "(" + uniqueID() + ")");
    return null;
  }
}


//=================================================================
//Prida data z IK do databazy
//=================================================================
function addDataIK(dateIK, menoIK, emailIK, poznamkaIK, vysledokIK, stavIK, skratka){
  var shDatIK = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0/edit#gid=0").getSheetByName("IK");  
  
  var rowLast = shDatIK.getLastRow() + 1; 
  var hlavicka = shDatIK.getRange(1, 1, 1, shDatIK.getLastColumn()).getValues();
  
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(dateIK);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Meno") + 1).setValue(menoIK);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(emailIK);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue(poznamkaIK);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Výsledok") + 1).setValue(vysledokIK);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Stav") + 1).setValue(stavIK);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Skratka") + 1).setValue(skratka);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("ID") + 1).setValue(IKUniqueID());

}


//=================================================================
//Prida data z Dohody do databazy
//=================================================================
function addDataDohoda(dateStart, dateEnd, dohoda, vysledok, info, skratka, id){
  var shDatIK = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1119X-R1RxEn4SK3xfZHdTmQKrvPddm5vXSz0Cmcepx0/edit#gid=0").getSheetByName("Dohody");  
  
  var rowLast = shDatIK.getLastRow() + 1; 
  var hlavicka = shDatIK.getRange(1, 1, 1, shDatIK.getLastColumn()).getValues();
  
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum začiatok") + 1).setValue(dateStart);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum koniec") + 1).setValue(dateEnd);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dohoda") + 1).setValue(dohoda);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Stav") + 1).setValue(vysledok);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Info") + 1).setValue(info);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("Skratka") + 1).setValue(skratka);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("ID - IK") + 1).setValue(id);
  shDatIK.getRange(rowLast, hlavicka[0].indexOf("ID - Dohoda") + 1).setValue(UniqueIDSimple());

}

//=================================================================
//Prida nepriradeneho klienta
//=================================================================
function addDataNK(okres, meno, cislo, datNar, vysledok, dateTime, email, zarobokSpec, zarobokOblast, id, ziskatel, poznamka){
  var shDatIK = SpreadsheetApp.openById("1mMTkZRwyw7W_w1Qm-FopzkMv2bvHvw7YgA8wCxZ_74k").getSheetByName("NK");  
  
  var rowLast = shDatIK.getLastRow() + 1; 
  var hlavicka = shDatIK.getRange(1, 1, 1, shDatIK.getLastColumn()).getValues();
  
  var data = shDatIK.getDataRange().getValues();
  var duplicate = false;
  for(var i = 1; i < data.length; i++)
  {
    var _id = data[i][data[0].indexOf("ID tiketu")];
    if(_id == id)
    {
      duplicate = true;
      break; //Kontakt je duplicitny takže sa neuloží
    }
  }
  
  if(duplicate == false){
    shDatIK.getRange(rowLast, 1).setValue(meno);
    shDatIK.getRange(rowLast, 2).setValue(cislo);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum telefonátu") + 1).setValue(dateTime);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Meno volajúceho") + 1).setValue(ziskatel);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum telefonátu") + 1).setValue(new Date());
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Okres") + 1).setValue(okres);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Výsledok telefonátu") + 1).setValue(vysledok);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(dateTime);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue(poznamka);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("ID tiketu") + 1).setValue(id);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Zárobok - špecifikácia") + 1).setValue(zarobokSpec);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Zarobok -  Oblasť") + 1).setValue(zarobokOblast);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum narodenia") + 1).setValue(datNar);    
  }
  
}


//=================================================================
//Prida nepriradeneho klienta
//=================================================================
function addDataCP(okres, meno, cislo, produkty, vysledok, dateTime, email, zarobokSpec, zarobokOblast, ziskatel, poznamka){
  var shDatIK = SpreadsheetApp.openById("1mAZdPPuGiD03hFQ5-sTfEKWzWJ-oxhnC3KJ3D6xtjp0").getSheetByName(okres);  
  
  var rowLast = shDatIK.getLastRow() + 1; 
  var hlavicka = shDatIK.getRange(1, 1, 1, shDatIK.getLastColumn()).getValues();
  
  var data = shDatIK.getDataRange().getValues();
  var duplicate = false;
  for(var i = 1; i < data.length; i++)
  {
    var _cislo = data[i][1];
    if(_cislo == cislo)
    {
      duplicate = true;
      break; //Kontakt je duplicitny takže sa neuloží
    }
  }
  
  if(duplicate == false){
    shDatIK.getRange(rowLast, 1).setValue(meno);
    shDatIK.getRange(rowLast, 2).setValue(cislo);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum telefonátu") + 1).setValue(dateTime);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Meno volajúceho") + 1).setValue(ziskatel);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum telefonátu") + 1).setValue(new Date());
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Okres") + 1).setValue(okres);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Výsledok telefonátu") + 1).setValue(vysledok);  
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(dateTime);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue(poznamka);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Specifikacia") + 1).setValue(zarobokSpec);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Oblast") + 1).setValue(zarobokOblast);
    shDatIK.getRange(rowLast, hlavicka[0].indexOf("Produkty") + 1).setValue(produkty);    
  }
  
}



//=================================================================
//Prida zmluvu
//=================================================================
function pridatZmluvu(databaza, idKlient, dateReg, ziskatel, meno, mobil, email, vlastnik, poisteny, spolocnost, produkt, cisloZmluvy, stav, poznamka, vyrDen,
                        datumUkoncenia, platba, interval, s, su, sud, tn, tnMaxPlnenie, pn, dOdskodne, kch, chz, hosp, zlom, inv40, inv70, osl, hodnotaNehnutelnosti,
                        poistenieNehnutelnosti, poistenieDomacnosti, zodpovednost, vyskaUveru, splatka, urokovaSadzba, rpmn, fixacia, mesInvesticia, hodnotaUctu, 
                        cielovaSuma, upozMakler, upozKlient, pocetTyzdnov){
  
  var shDat;
  if(databaza == "Badinka Peter"){
    shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1gQUhngN6oZmA_jVCTACJvEbKZKD1mBLq-a1bUdm-c7I/edit#gid=862465608").getSheetByName("Zmluvy");
  }
  if(databaza == "Centrálna databáza"){
    shDat = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o/edit#gid=862465608").getSheetByName("Zmluvy");
  }       
  var rowLast = shDat.getLastRow() + 1; 
  var hlavicka = shDat.getRange(1, 1, 1, shDat.getLastColumn()).getValues();
  
  shDat.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(dateReg);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Skr.") + 1).setValue(ziskatel);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(meno);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Mobil") + 1).setValue(mobil);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Poistník/Vlastník zmluvy") + 1).setValue(vlastnik);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Poistený") + 1).setValue(poisteny);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Produkt") + 1).setValue(produkt);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Spoločnosť") + 1).setValue(spolocnost);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Číslo zmluvy") + 1).setValue(cisloZmluvy);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Stav") + 1).setValue(stav);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue(poznamka);
  shDat.getRange(rowLast, hlavicka[0].indexOf("VD") + 1).setValue(vyrDen);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Dátum ukončenia") + 1).setValue(datumUkoncenia);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Platba") + 1).setValue(platba);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Interval platenia") + 1).setValue(interval);
  shDat.getRange(rowLast, hlavicka[0].indexOf("S") + 1).setValue(s);
  shDat.getRange(rowLast, hlavicka[0].indexOf("SÚ") + 1).setValue(su);
  shDat.getRange(rowLast, hlavicka[0].indexOf("SUD") + 1).setValue(sud);
  shDat.getRange(rowLast, hlavicka[0].indexOf("TN") + 1).setValue(tn);
  shDat.getRange(rowLast, hlavicka[0].indexOf("TN - Maximálne plnenie") + 1).setValue(tnMaxPlnenie);
  shDat.getRange(rowLast, hlavicka[0].indexOf("PN") + 1).setValue(pn);
  shDat.getRange(rowLast, hlavicka[0].indexOf("DO") + 1).setValue(dOdskodne);
  shDat.getRange(rowLast, hlavicka[0].indexOf("KCH") + 1).setValue(kch);
  shDat.getRange(rowLast, hlavicka[0].indexOf("CHZ") + 1).setValue(chz);
  shDat.getRange(rowLast, hlavicka[0].indexOf("HOSP") + 1).setValue(hosp);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Zlomeniny a popaleniny") + 1).setValue(zlom);
  shDat.getRange(rowLast, hlavicka[0].indexOf("INV 40") + 1).setValue(inv40);
  shDat.getRange(rowLast, hlavicka[0].indexOf("INV 70") + 1).setValue(inv70);
  shDat.getRange(rowLast, hlavicka[0].indexOf("OSL") + 1).setValue(osl);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Hodnota nehnuteľosti") + 1).setValue(hodnotaNehnutelnosti);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Poistenie - Nehnuteľnosť") + 1).setValue(poistenieNehnutelnosti);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Poistenie - domácnosť") + 1).setValue(poistenieDomacnosti);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Zodpov.") + 1).setValue(zodpovednost);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Výška úveru") + 1).setValue(vyskaUveru);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Splátka") + 1).setValue(splatka);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Uroková sadzba") + 1).setValue(urokovaSadzba);
  shDat.getRange(rowLast, hlavicka[0].indexOf("RPMN") + 1).setValue(rpmn);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Fixácie") + 1).setValue(fixacia);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Pravidelná mesačná investícia") + 1).setValue(mesInvesticia);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Hodnota účtu / Odkupná hodnota") + 1).setValue(hodnotaUctu);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Cieľová suma") + 1).setValue(cielovaSuma);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Upozornenie - maklér") + 1).setValue(upozMakler);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Upozornenie - klient") + 1).setValue(upozKlient);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Počet týždňov") + 1).setValue(pocetTyzdnov);
  shDat.getRange(rowLast, hlavicka[0].indexOf("ID - Osoba") + 1).setValue(idKlient);
  shDat.getRange(rowLast, hlavicka[0].indexOf("ID - Zmluva") + 1).setValue(uniqueID());
}





//=================================================================
//Prida inu osobu
//=================================================================
function addDataInaOsoba(databaza, dateReg, typOdp, ziskatel, pobocka, meno, mobil, datumTelefonatu, stav, datumNajblTelefonatu, poznamka, status, email, povolanie,
    ulica, obec, datNarodenia, miestoNarodenia, rodneCislo, cisloOp, platnostOpOd, platnostOpDo, prijmy, vydavky, dokladVydal, cisloUctu, lekarMeno, lekarAdresa, lekarKontakt, vyska, hmotnost, oslovenie,
    priezvisko, krstneMeno, meniny, interval, androidSync, vztah, selectPersonID){
    
  var shDat = SpreadsheetApp.openById("1pntnrFj_XwAZCgLhnRRmgAJ6nuBV1Y2nyVKrGHRGR0o").getSheetByName("Iné osoby");
  var rowLast = shDat.getLastRow() + 1; 
  var hlavicka = shDat.getRange(1, 1, 1, shDat.getLastColumn()).getValues();
  shDat.getRange(rowLast, hlavicka[0].indexOf("Dátum") + 1).setValue(dateReg);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Odp.") + 1).setValue(typOdp);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Skr.") + 1).setValue(ziskatel);
  shDat.getRange(rowLast, hlavicka[0].indexOf("O.") + 1).setValue(pobocka);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Priezvisko a meno") + 1).setValue(meno);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Mobil") + 1).setValue(mobil);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Úspešný kontakt") + 1).setValue(datumTelefonatu);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Proces") + 1).setValue(stav);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Dátum akcie") + 1).setValue(datumNajblTelefonatu);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Poznámka") + 1).setValue(poznamka);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Android Sync") + 1).setValue(status);        
  shDat.getRange(rowLast, hlavicka[0].indexOf("Email") + 1).setValue(email);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Špecifikácia") + 1).setValue(povolanie); 
  shDat.getRange(rowLast, hlavicka[0].indexOf("Ulica") + 1).setValue(ulica);
  shDat.getRange(rowLast, hlavicka[0].indexOf("PSČ / Obec") + 1).setValue(obec);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Dátum narodenia") + 1).setValue(datNarodenia);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Miesto narodenia") + 1).setValue(miestoNarodenia);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Rodné číslo") + 1).setValue(rodneCislo);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Číslo OP") + 1).setValue(cisloOp);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Platnosť od") + 1).setValue(platnostOpOd);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Platnosť do") + 1).setValue(platnostOpDo);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Doklad vydal") + 1).setValue(dokladVydal);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Príjmy") + 1).setValue(prijmy);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Výdavky") + 1).setValue(vydavky);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Číslo účtu") + 1).setValue(cisloUctu);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Obvodný lekár - Meno") + 1).setValue(lekarMeno);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Obvodný lekár - Adresa") + 1).setValue(lekarAdresa);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Obvodný lekár - Kontakt") + 1).setValue(lekarKontakt);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Výška") + 1).setValue(vyska);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Hmotnosť") + 1).setValue(hmotnost);
  
  shDat.getRange(rowLast, hlavicka[0].indexOf("Priezvisko") + 1).setValue(priezvisko);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Meno") + 1).setValue(krstneMeno);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Oslovenie") + 1).setValue(oslovenie);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Narodeniny / meniny") + 1).setValue(meniny);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Servisný email") + 1).setValue(interval);
  shDat.getRange(rowLast, hlavicka[0].indexOf("Android Sync") + 1).setValue(androidSync);
  
  shDat.getRange(rowLast, hlavicka[0].indexOf("Vzťah") + 1).setValue(vztah);
  
  shDat.getRange(rowLast, hlavicka[0].indexOf("ID - Iná osoba") + 1).setValue(uniqueID());
  shDat.getRange(rowLast, hlavicka[0].indexOf("ID - Databáza") + 1).setValue(selectPersonID);
  return null;
}





//=================================================================
//Ulozi denne aktivity
//=================================================================
function saveAktivity(userName, oslovenyKlienti, register, terminy, predstavenie, analyza, predaj, podpis, servis, odporucania, beby, pzpHp, poistenieMajetku, poistenieOsob, hypoteka, uver,
                                  investicie, druhyPilier, tretiPilier, ine){
  
  var shUser = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1anOfRjI3QivEcMpq7zk0uqgbde_lBUK5rq2OWOan_Ak/edit#gid=2095307398").getSheetByName("Users");
  var lastRow = shUser.getLastRow();
  var urlAktivity = "";
  for(var i = 2; i <= lastRow; i++){
    var user = shUser.getRange(i, 3).getValue();
    if(user == userName){
      urlAktivity = shUser.getRange(i, 1).getValue();  
      if(urlAktivity.length < 2){ //ak nie je tabulka na aktivity vytvorena tak sa skopiruje nova zo sablony      
        var idSablonaAktivity = "1Yclm-IV8z6n1C0moAgKDf9L5Hx49qwitbf7Qfkwp2tI";
        var driver = DriveApp.getFileById(idSablonaAktivity); //odkaz na šablónu
        var fileCopy = driver.makeCopy(userName);      
        urlAktivity = fileCopy.getUrl();
        shUser.getRange(i, 1).setValue(urlAktivity);        
      }
      var shAktivity = SpreadsheetApp.openByUrl(urlAktivity).getSheetByName("Aktivity");      
      shAktivity.getRange(1, 2).setValue(oslovenyKlienti);
      shAktivity.getRange(2, 2).setValue(register);
      shAktivity.getRange(3, 2).setValue(terminy);
      shAktivity.getRange(4, 2).setValue(predstavenie);
      shAktivity.getRange(5, 2).setValue(analyza);
      shAktivity.getRange(6, 2).setValue(predaj);
      shAktivity.getRange(7, 2).setValue(podpis);
      shAktivity.getRange(8, 2).setValue(servis);
      shAktivity.getRange(9, 2).setValue(odporucania);
      shAktivity.getRange(10, 2).setValue(beby);
      shAktivity.getRange(11, 2).setValue(pzpHp);
      shAktivity.getRange(12, 2).setValue(poistenieMajetku);
      shAktivity.getRange(13, 2).setValue(poistenieOsob);
      shAktivity.getRange(14, 2).setValue(hypoteka);
      shAktivity.getRange(15, 2).setValue(uver);
      shAktivity.getRange(16, 2).setValue(investicie);
      shAktivity.getRange(17, 2).setValue(druhyPilier);
      shAktivity.getRange(18, 2).setValue(tretiPilier);
      shAktivity.getRange(19, 2).setValue(ine);      
    }
  } 
}





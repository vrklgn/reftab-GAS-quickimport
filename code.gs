// Makes fun menus in the sheet

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Reftab')
      .addItem('Import Laptops', 'menuItem1')
      //.addSeparator()
      .addItem('Import Phones', 'menuItem2')
      .addItem('Import Monitors', 'menuItem3')
      .addToUi();
}

// Function to change "plain text" country to the designated location codes in Reftab

function getlocationCode (countryName) {
    if (locationCode.hasOwnProperty(countryName)) {
        return locationCode[countryName];
    } else {
        return "N/A";
    }
}

var locationCode = {
"Spain":15051,
"Sweden":14856,
"Poland":15047,
"Netherlands":15055,
"France":15053,
"Germany":15059,
"Norway":15060,
"England":15048,
"UK":15048,
"Italy":15052,
"Croatia":15319,
"Slovakia":15054,
"Denmark":15057,
"Portugal":15056,
"Finland":15058,
"US":44270,
"Mexico": 18384
};

function menuItem1() {
  SpreadsheetApp.getUi()
     .alert('Parsing Laptops... (Sheet will be wiped on refresh)');
     parselaptops();
}

function menuItem2() {
  SpreadsheetApp.getUi()
     .alert('Parsing Phones... (Sheet will be wiped on refresh)');
     parsePhones();
}

function menuItem3() {
  SpreadsheetApp.getUi()
     .alert('Parsing Monitors... (Sheet will be wiped on refresh)');
     parsemonitors();
}

function clearSheet() {
  var laptopSheet = SpreadsheetApp.getActive().getSheetByName("laptops");
  var laptopHeaders = ['Serial Number','Brand','Model', 'Screen Size', 'RAM', "Date Purchased", 'Purchase Cost', 'ADE','Architecture','Laptop Classification','Keyboard Layout','Location', 'Result'];
  laptopSheet.clear();
  laptopSheet.appendRow(laptopHeaders)
  var phoneSheet = SpreadsheetApp.getActive().getSheetByName("phones");
  var phoneHeaders = ['Serial Number','Brand','Model', 'Storage', "Date Purchased", 'Purchase Cost', 'ADE', 'Location','Status', 'Result'];
  phoneSheet.clear();
  phoneSheet.appendRow(phoneHeaders)
  var monitorSheet = SpreadsheetApp.getActive().getSheetByName("monitors");
  var monitorheaders = ['Serial Number','Brand','Model', 'Size', "Date Purchased", 'Purchase Cost', 'Location','Status', 'Result'];
  monitorSheet.clear();
  monitorSheet.appendRow(monitorheaders)
}


// Function from Reftab github
// https://github.com/Reftab/ReftabGAS

function Reftab(options) {
  var scriptProperties = PropertiesService.getScriptProperties();
  const secretKey = scriptProperties.getProperty("secretKey")
  const publicKey = scriptProperties.getProperty("publicKey")

 function signRequest (request) {
    request.headers['x-rt-date'] = new Date().toUTCString();
    request.contentType = 'application/json';
   
    let md5 = str =>
      Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str).reduce((str, chr) => {
        chr = (chr < 0 ? chr + 256 : chr).toString(16);
        return `${str}${chr.length == 1 ? '0' : ''}${chr}`;
      }, '');

    let token = Utilities.computeHmacSha256Signature(unescape(encodeURIComponent(
`${request.method}
${request.payload !== undefined ? md5(request.payload) : ''}
${request.payload !== undefined ? request.contentType : ''}
${request.headers['x-rt-date']}
${request.url}`
      )), secretKey);
   
    token = token.map(byte => {
      let v = (byte < 0) ? 256 + byte : byte;
      return ('0' + v.toString(16)).slice(-2);
    }).join('');
   
    request.headers.Authorization = `RT ${publicKey}:${Utilities.base64Encode(token)}`;
   
    return request;
  }

  this.request = function(method, endpoint, id, body) {
    if (id) {
      endpoint += '/' + id;
    }
    if (body) {
      body = JSON.stringify(body);
    }
    let options = {
      method: method,
      url: 'https://www.reftab.com/api/' + endpoint,
      payload: body,
      headers: {}
    };
    let response = UrlFetchApp.fetch(options.url, signRequest(options));
   
    return JSON.parse(response.getContentText());
  };
  
  this.get = function(endpoint, id) {
    return this.request('GET', endpoint, id);
  }
  
  this.put = function(endpoint, id, body) {
    return this.request('PUT', endpoint, id, body);
  }
  
  this.post = function(endpoint, body) {
    return this.request('POST', endpoint, undefined, body);
  }
  
  this.delete = function(endpoint, id) {
    return this.request('DELETE', endpoint, id);
  }

  return this;
}


function makeLaptop(serial,brand,model,size,ram,date,cost,ade,cpu,laptopclass,keylayout,location,type) {
  let assetpayload = {
    "title": serial,
    "cid": type,
    "clid": location,
    "statid": 44253,
    "details": {
      "RAM": ram,
      "Date Purchased": date,
      "Purchase Cost": cost,
      "Computer Model": model,
      "Laptop Brand": brand,
      "Screen Size": size,
      "Serial Number": serial,
      "ADE": ade,
      "Architecture": cpu,
      "Laptop Classification": laptopclass,
      "Keyboard Layout": keylayout
      }
  }
  const api = new Reftab();
  let asset = api.post('assets', assetpayload);
  return(asset)
  }

  function makePhone(serial,brand,model,storage,date,cost,ade,location,type,status) {
  if (status == "Storage") {
    var statid = 32677
  } else if (status == "Claimed") {
    var statid = 32676
  } else if (status == "Lost/Stolen") {
    var statid = 32679
  } else if (status == "Destroyed") {
    var statid = 33634
  } else {
    var statid = 33116
  } 
  let assetpayload = {
    "title": serial,
    "cid": type,
    "clid": location,
    "statid": statid,
    "details": {
      "Date Purchased": date,
      "Purchase Cost": cost,
      "Phone Model": model,
      "Phone Brand": brand,
      "Phone Storage": storage,
      "Serial Number": serial,
      "ADE": ade}
  }
  const api = new Reftab();
  let asset = api.post('assets', assetpayload);
  return(asset)
  }


  function makeMonitor(serial,brand,model,size,date,cost,location,type,status) {
  if (status == "Storage") {
    var statid = 32677
  } else {
    var statid = 33116
  } 
  var assetpayload = {
    "title": serial,
    "cid": type,
    "clid": location,
    "statid": statid,
    // statid: 44253 = Storage, 33116 = Deployed 
    "details": {
      "Date Purchased": date,
      "Purchase Cost": cost,
      "Monitor Model": model,
      "Monitor Brand": brand,
      "Monitor Size": size,
      "Serial Number": serial
      }
  }
  const api = new Reftab();
  let asset = api.post('assets', assetpayload);
  console.log (assetpayload)
  Logger.log (assetpayload)
  return(asset)
  }


 // Parse Phones from sheet
function parsePhones(){
  // Get the sheet called "laptops"
  var sheet = SpreadsheetApp.getActive().getSheetByName("phones");
  // Get the data and number of rows.
  var range = sheet.getDataRange()
  var numRows = range.getNumRows();
  var values = range.getValues();
  // For each row - place values into variables and send it to the "makeLaptop" function
  for (var asset = 1; asset < numRows; asset++){
    var item = values[asset]
    var serial = item[0]
    var brand = item[1]
    var model = item[2]
    var storage = item[3]
    var date = item[4]
    var cost = item[5]*100
    var ade = item[6]
    var type = 16024
    var location = getlocationCode(item[7])
    var status = item[8]
      try {var result = makePhone(serial,brand,model,storage,date,cost,ade,location,type,status)
      // Place asset ID in column K upon success!
      sheet.getRange(asset+1,10).setValue("https://www.reftab.com/assets/"+result.aid)
      }
      catch (e) {
        // Place error in column K upon failure
       sheet.getRange(asset+1,10).setValue(e)}
  }
}

// Parse laptops from sheet
function parselaptops(){
  // Get the sheet called "laptops"
  var sheet = SpreadsheetApp.getActive().getSheetByName("laptops");
  // Get the data and number of rows.
  var range = sheet.getDataRange()
  var numRows = range.getNumRows();
  var values = range.getValues();
  // For each row - place values into variables and send it to the "makeLaptop" function
  for (var asset = 1; asset < numRows; asset++){
    var item = values[asset]
    var serial = item[0]
    var brand = item[1]
    var model = item[2]
    var size = item[3]
    var ram = item[4]
    var date = item[5]
    var cost = item[6]*100
    var ade = item[7]
    var type = 16022
    var cpu = item[8]
    var laptopclass = item[9]
    var keylayout = item[10]
    var location = getlocationCode(item[11])
      try {var result = makeLaptop(serial,brand,model,size,ram,date,cost,ade,cpu,laptopclass,keylayout,location,type)
      // Place asset ID in column M upon success!
      sheet.getRange(asset+1,13).setValue("https://www.reftab.com/assets/"+result.aid)
      }
      catch (e) {
        // Place error in column m upon failure
       sheet.getRange(asset+1,13).setValue(e)}
  }
}


// Parse monitors from sheet
function parsemonitors(){
  // Get the sheet called "monitors"
  var sheet = SpreadsheetApp.getActive().getSheetByName("monitors");
  // Get the data and number of rows.
  var range = sheet.getDataRange()
  var numRows = range.getNumRows();
  var values = range.getValues();
  // For each row - place values into variables and send it to the "makeMonitor" function
  for (var asset = 1; asset < numRows; asset++){
    var item = values[asset]
    var serial = item[0]
    var brand = item[1]
    var model = item[2]
    var size = item[3]
    var date = item[4]
    var cost = item[5]*100
    var type = 16025
    var location = getlocationCode(item[6])
    var status = item[7]
      try {var result = makeMonitor(serial,brand,model,size,date,cost,location,type,status)
      Logger.log (result)
      console.log (result)
      // Place asset ID in column K upon success!
      sheet.getRange(asset+1,9).setValue("https://www.reftab.com/assets/"+result.aid)
      }
      catch (e) {
        // Place error in column K upon failure
       sheet.getRange(asset+1,9).setValue(e)}
  }
}


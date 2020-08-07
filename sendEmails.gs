function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Enviar')
  .addItem('Para Cliente', 'targetCliente')
  .addToUi();
}

//----------------------Menu--------------------
function targetCliente(){
  var clienteName = "cliente";
  var clienteEmail = "adrianalimafm@gmail.com";
  var clienteIndex = 5;
  setMaterials(clienteName, clienteEmail, clienteIndex);
}

// ------------- Busca as todas as informações da planilha --------------------
  function getData(sheetName) {
    var data = SpreadsheetApp.getActive().getSheetByName(sheetName).getDataRange().getValues();
    return data;
  }
  
// ------------- Busca o que foi selecionado na planilha --------------------
  function getSelection(sheetName){
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    var selection = sheet.getSelection(); 
    var result = selection.getActiveRangeList().getRanges();      
    var rangeList = [];
    
    for (var i = 0; i < result.length; i++) {
      var data = result[i].getValues();        
      rangeList = rangeList.concat(data);   
    }    
    return rangeList;
  }
  
// ------------- Renderiza os Templates --------------------
  function renderTemplate(template, data) {
    var output = template;
    var params = template.match(/\{\{(.*?)\}\}/g);
    params.forEach(function (param) {
      var propertyName = param.slice(2,-2); //Remove the {{ and the }}
      output = output.replace(param, data[propertyName] || "");
    });
    return output;
  }
  
// ------------- Converte as linhas em objetos --------------------
  function rowsToObjects(rows) {
    var headers = rows.shift();  
    var data = [];
    
    rows.forEach(function (row) {
      var object  = {}; //cria um objeto
      
      row.forEach(function (value, index) {   
        object[headers[index]] = value;
      });
      
      data.push(object);
    });
    return data;   
  }
  
  
//--------------------------------------------------------------  
//------------ Prepara Materiais para Envio --------------------
//------------------------------------------------------------- 
  
  function setMaterials(targetName, targetEmail, targetIndex){
  

//----------------Criar Item ---------------
    // pega as informações selecionadas
    var dataSelection = getSelection("Materiais prontos"); 
    
    // pega os valores do header
    var headerRange =  SpreadsheetApp.getActiveSheet().getRange(1, 1, 1, 4);
    var header = headerRange.getValues(); 
    
    // concatena o header com a seleção
    var concat = header.concat(dataSelection);
      
    // transforma o dataSelection em um objeto
    dataSelection = rowsToObjects(concat);
    
   
//----------------Criar lista---------------
    var objectGroupConfirm = "";
    var objectGroupEmail = "";
    
    // para cada item das informações selecionadas
    dataSelection.forEach(function (object) {      
      
       //formata dados no template da confirmação
       var templateData = getData("Templates");
       var objectTemplate = templateData[7][0];
       var objectItem = renderTemplate(objectTemplate, object);
       objectGroupConfirm = objectGroupConfirm + objectItem +  "\n\n\n";     
       
       //formata dados no template do email para o cliente
       var emailBodyTemplate = templateData[4][0];     
       var body = renderTemplate(emailBodyTemplate, object);    
       objectGroupEmail = objectGroupEmail + body + "\n-----------------------\n\n";              
    });    
    
//----------------Mensagem de confirmação ---------------    
      var ui = SpreadsheetApp.getUi();
      
      var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
      var alertTitle = "Confirma o envio dos materiais para " + targetName + "?";
      var alertMessage = "(Cota diária de email disponível: " + emailQuotaRemaining + ")\n\n" + objectGroupConfirm;   

      var confirm = ui.prompt(alertTitle, alertMessage , ui.ButtonSet.OK_CANCEL);
      var subject = "Recebimento de Material - " + confirm.getResponseText();
        
      if (confirm.getSelectedButton() == ui.Button.OK) {  
        MailApp.sendEmail({to: targetEmail, subject: subject, body: objectGroupEmail }); 
        sentDate(targetIndex);
      }  
  }  

//---------------- Colocar data de envio ---------------  
function sentDate(targetIndex) {  
  var sheet = SpreadsheetApp.getActiveSheet();
  var selection = sheet.getSelection(); 
  var result = selection.getActiveRangeList().getRanges(); 
  
  for (var i = 0; i < result.length; i++) {
    var data = result[i].getRowIndex();    
    var dateFormated = Utilities.formatDate(new Date(), "GMT-03:00", "dd/MM/yyyy HH:mm:ss");
    var dateCell = sheet.getRange(data, targetIndex);
    
    var dateCellValue = dateCell.isBlank();
    
    if (dateCellValue){
      dateCell.setValue("Enviado em: " + dateFormated).setBackground("#fffbd0");
    } else {
      dateCell.setValue("Reenviado em: " + dateFormated).setBackground("#e6e2ba");
    }   
    
  } 
}




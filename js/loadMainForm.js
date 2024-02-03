/* Lendo o Formulário */
function loadMainForm(){

    const htmlServ = HtmlService.createTemplateFromFile("main");
    const html = htmlServ.evaluate();
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(html,"Edit Customer Data");
  
  }
  
  /* Criando o Menu na Planilha */
  function createMenu_(){
  
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("Custom Menu");
    menu.addItem("Open Form", "loadMainForm");
    menu.addToUi();
  
  }
  
  /* Funcão na abertura da planilha */
  function onOpen(){
    createMenu_();
  }
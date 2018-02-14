function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Square')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
  showSidebar();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('page')
      .setTitle('Sidebar')
      .setWidth(200);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function startUp() {

}

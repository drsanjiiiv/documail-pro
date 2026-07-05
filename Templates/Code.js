// ==========================================
// DOCUMENT MASTER TEMPLATE - ONOPEN
// ==========================================
function onOpen(e) {
  try {
    // Create the clean, single entry point menu in the Document
    DocumentApp.getUi()
               .createMenu('DocuMail Pro')
               .addItem('Start Dynamic Doc Template', 'INITIALIZE_DOC_DESIGNER_SIDEBAR')
               .addToUi();

    // Fire the automatic guiding toast on-screen instantly when opened
    DocumentApp.getActiveDocument().toast(
      "To start creating your template rules, use: DocuMail Pro > Start Dynamic Doc Template from the menu above.", 
      "🚀 DocuMail Pro Engine", 
      12
    );
  } catch (err) {
    Logger.log("Error loading document menu layout: " + err.toString());
  }
}
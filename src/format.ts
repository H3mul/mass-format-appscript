/**
 * Reads doc IDs from script parameters and installs onOpen triggers for each.
 * Expects a comma-separated list of doc IDs in script properties under 'DOC_IDS'.
 */
function installTriggersFromScriptParameters() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const docIdsParam = scriptProperties.getProperty('DOC_IDS');
  if (!docIdsParam) {
    throw new Error('No DOC_IDS parameter found in script properties.');
  }
  const docIds = docIdsParam.split(',').map(id => id.trim()).filter(Boolean);
  docIds.forEach(installOnOpenTriggerForDoc);
}

/**
 * Installs an onOpen trigger for a Google Doc by its ID, calling the menu creation function.
 * @param {string} docId - The ID of the Google Doc to attach the trigger to.
 */
function installOnOpenTriggerForDoc(docId: string) {
  // Remove existing triggers for this function and doc to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (
      triggers[i].getHandlerFunction() === 'createDocHighlightMenu' &&
      triggers[i].getTriggerSourceId &&
      triggers[i].getTriggerSourceId() === docId
    ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('createDocHighlightMenu')
    .forDocument(docId)
    .onOpen()
    .create();
}

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Mass Format')
    .addItem('Highlight Selection everywhere', 'showColorPrompt')
    .addToUi();
}

function createDocHighlightMenu(event: GoogleAppsScript.Events.DocsOnOpen) {
    onOpen();
}

function showColorPrompt() {
  const selection = getSelectedText();
  if (!selection) {
    DocumentApp.getUi().alert('Please select a word or phrase to highlight first.');
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile('src/ColorPicker')
    .setWidth(200)

  DocumentApp.getUi().showModalDialog(html, 'Select a Color');
}

function processHighlight(formObject: any) {
  const color = formObject.color;

  if (color) {
    highlightSelectedText(color);
  } else {
    DocumentApp.getUi().alert('Please select a color.');
  }
}

function highlightSelectedYellow() {
  highlightSelectedText('#ffff00');
}

function highlightSelectedGreen() {
  highlightSelectedText('#00ff00');
}

function highlightSelectedCyan() {
  highlightSelectedText('#00ffff');
}

function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();

  if (!selection) {
    return "";
  }

  var selectedElements = selection.getRangeElements();
  var theText = "";

  for (var i = 0; i < selectedElements.length; i++) {
    var element = selectedElements[i].getElement();

    // Check if the element is editable text
    const maybeEdit = (element as any).editAsText;
    if (typeof maybeEdit === 'function') {
      var textElement = maybeEdit.call(element);
      if (selectedElements[i].isPartial()) {
        // If only part of the element is selected, get the specific range
        theText += textElement.getText().substring(
          selectedElements[i].getStartOffset(),
          selectedElements[i].getEndOffsetInclusive() + 1
        );
      } else {
        // If the entire element is selected
        theText += textElement.getText();
      }
    }
  }
  return theText;
}

function highlightSelectedText(color) {
  const ui = DocumentApp.getUi();
  const selection = getSelectedText();
  if (selection) {
    highlightWord(selection, color);
  } else {
    ui.alert('Please select a word or phrase to highlight first.');
  }
}

function highlightWord(wordToFind: string, highlightColor: string) {
  if (!wordToFind || !highlightColor) {
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const searchPattern = wordToFind;

  let foundElement = body.findText(searchPattern);
  let count = 0;

  while (foundElement) {
    const element = foundElement.getElement();
    const text = element.asText();
    const startIndex = foundElement.getStartOffset();
    const endIndex = foundElement.getEndOffsetInclusive();

    text.setBackgroundColor(startIndex, endIndex, highlightColor);
    count++;

    foundElement = body.findText(searchPattern, foundElement);
  }
}

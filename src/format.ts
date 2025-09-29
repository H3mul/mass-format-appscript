/**
 * NOTE: This file is generated from the repository at https://github.com/H3mul/mass-format-appscript
 *
 * To make changes, please modify the source code and push it to the repository.
 * This will ensure that the changes are not overwritten by a future deployment.
 */

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

  // Remove all existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  docIds.forEach(installOnOpenTriggerForDoc);
}

/**
 * Installs an onOpen trigger for a Google Doc by its ID, calling the menu creation function.
 * @param {string} docId - The ID of the Google Doc to attach the trigger to.
 */
function installOnOpenTriggerForDoc(docId: string) {
  ScriptApp.newTrigger('createDocHighlightMenu')
    .forDocument(docId)
    .onOpen()
    .create();
}

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Mass Format')
    .addItem('Highlight selection everywhere', 'showColorPrompt')
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
  
  const currentColor = getCurrentBackgroundColor();
  let html = HtmlService.createHtmlOutputFromFile('src/ColorPicker');
  
  if (currentColor) {
    const htmlContent = html.getContent();
    const modifiedContent = htmlContent.replace(
      '<script>',
      '<script>window.initialColor = "' + currentColor + '";'
    );
    html = HtmlService.createHtmlOutput(modifiedContent);
  }
  
  html.setWidth(200);
  DocumentApp.getUi().showModalDialog(html, 'Select a Color');
}

function processHighlight(formObject: any) {
  // formObject may be { color: '#rrggbb' } or { clear: true }
  if (formObject && formObject.clear) {
    // user requested clearing the background color
    highlightSelectedText('CLEAR');
    return;
  }

  const color = formObject && formObject.color;
  if (color) {
    highlightSelectedText(color);
  } else {
    DocumentApp.getUi().alert('Please select a color.');
  }
}

function processSelectedElements(callback: (textElement: any, rangeElement: any) => any) {
  var selection = DocumentApp.getActiveDocument().getSelection();

  if (!selection) {
    return null;
  }

  var selectedElements = selection.getRangeElements();

  for (var i = 0; i < selectedElements.length; i++) {
    var element = selectedElements[i].getElement();

    // Check if the element is editable text
    if (element.editAsText) {
      var textElement = element.editAsText();
      var result = callback(textElement, selectedElements[i]);
      if (result !== undefined) {
        return result;
      }
    }
  }
  return null;
}

function getSelectedText() {
  var theText = "";
  
  processSelectedElements((textElement, rangeElement) => {
    if (rangeElement.isPartial()) {
      // If only part of the element is selected, get the specific range
      theText += textElement.getText().substring(
        rangeElement.getStartOffset(),
        rangeElement.getEndOffsetInclusive() + 1
      );
    } else {
      // If the entire element is selected
      theText += textElement.getText();
    }
  });
  
  return theText;
}

function getCurrentBackgroundColor() {
  return processSelectedElements((textElement, rangeElement) => {
    var startOffset = rangeElement.isPartial() ? rangeElement.getStartOffset() : 0;
    
    // Get the background color of the first character in the selection
    var backgroundColor = textElement.getBackgroundColor(startOffset);
    
    if (backgroundColor && backgroundColor !== '#000000') {
      return backgroundColor;
    }
  });
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
    if (highlightColor === 'CLEAR') {
      // Clear the background color by setting the BACKGROUND_COLOR attribute to null
      const attrs: any = {};
      attrs[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
      text.setAttributes(startIndex, endIndex, attrs);
    } else {
      text.setBackgroundColor(startIndex, endIndex, highlightColor);
    }
    count++;

    foundElement = body.findText(searchPattern, foundElement);
  }
}

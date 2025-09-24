function onOpen() {
  DocumentApp.getUi()
    .createMenu('Highlight Word')
    .addItem('Highlight Selection everywhere', 'showColorPrompt')
    .addToUi();
}

function showColorPrompt() {
  const html = HtmlService.createHtmlOutputFromFile('ColorPicker')
    .setWidth(200)
    .setHeight(200);
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
    if (element.editAsText) {
      var textElement = element.editAsText();
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

  if (count > 0) {
    DocumentApp.getUi().alert(`Highlighted ${count} instance(s) of "${wordToFind}".`);
  } else {
    DocumentApp.getUi().alert(`No instances of "${wordToFind}" were found.`);
  }
}
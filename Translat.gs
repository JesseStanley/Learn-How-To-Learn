function myFunction() {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}
function onInstall(e) {
  onOpen(e);
}
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Translate');
  DocumentApp.getUi().showSidebar(ui);
}
function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var text = [];
    var elements = selection.getSelectedElements();
    for (var i = 0; i < elements.length; ++i) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        text.push(element.getText().substring(startIndex, endIndex + 1));
      } else {
        var element = elements[i].getElement()
        if (element.editAsText) {
          var elementText = element.asText().getText();
             if (elementText) {
            text.push(elementText);
          }
        }
      }
    }
    if (!text.length) {
      throw new Error('Please select some text.');
    }
    return text;
  } else {
    throw new Error('Please select some text.');
  }
}
function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  return {
    originLang: userProperties.getProperty('originLang'),
    destLang: userProperties.getProperty('destLang')
  };
}
function getTextAndTranslation(origin, dest, savePrefs) {
  if (savePrefs) {
    PropertiesService.getUserProperties()
        .setProperty('originLang', origin)
        .setProperty('destLang', dest);
  }
  var text = getSelectedText().join('\n');
  return {
    text: text,
    translation: translateText(text, origin, dest)
  };
}
function insertText(newText) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var replaced = false;
    var elements = selection.getSelectedElements();
    if (elements.length === 1 && elements[0].getElement().getType() ===
        DocumentApp.ElementType.INLINE_IMAGE) {
      throw new Error('Can\'t insert text into an image.');
    }
    for (var i = 0; i < elements.length; ++i) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();
        element.deleteText(startIndex, endIndex);
        if (!replaced) {
          element.insertText(startIndex, newText);
          replaced = true;
        } else {
          var parent = element.getParent();
          var remainingText = element.getText().substring(endIndex + 1);
          parent.getPreviousSibling().asText().appendText(remainingText);
         
          if (parent.getNextSibling()) {
            parent.removeFromParent();
          } else {
            element.removeFromParent();
          }
        }
      } else {
        var element = elements[i].getElement();
        if (!replaced && element.editAsText) {
          
          element.clear();
          element.asText().setText(newText);
          replaced = true;
        } else {
          if (element.getNextSibling()) {
            element.removeFromParent();
          } else {
            element.clear();
          }
        }
      }
    }
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var surroundingText = cursor.getSurroundingText().getText();
    var surroundingTextOffset = cursor.getSurroundingTextOffset();
    if (surroundingTextOffset > 0) {
      if (surroundingText.charAt(surroundingTextOffset - 1) != ' ') {
        newText = ' ' + newText;
      }
    }
    if (surroundingTextOffset < surroundingText.length) {
      if (surroundingText.charAt(surroundingTextOffset) != ' ') {
        newText += ' ';
      }
    }
    cursor.insertText(newText);
  }
}
function translateText(text, origin, dest) {
  if (origin === dest) return text;
  return LanguageApp.translate(text, origin, dest);
}

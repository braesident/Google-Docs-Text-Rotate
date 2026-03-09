function onOpen() {
  DocumentApp.getUi()
    .createMenu('Tabellen-Tools')
    .addItem('Text drehen', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Text-Rotation')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

function getSelectedText() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return null;

  const element = selection.getRangeElements()[0].getElement().asText();
  const offset = selection.getRangeElements()[0].getStartOffset();

  let fontFamily = element.getFontFamily(offset);
  if (!fontFamily) {
    const parentAttributes = element.getParent().getAttributes();
    fontFamily = parentAttributes[DocumentApp.Attribute.FONT_FAMILY] || "Arial";
  }

  let bgColor = "#ffffff";
  let parent = element.getParent();
  while (parent && parent.getType() !== DocumentApp.ElementType.TABLE_CELL) {
    parent = parent.getParent();
  }
  if (parent && parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
    bgColor = parent.asTableCell().getBackgroundColor() || "#ffffff";
  }

  return {
    text: element.getText().substring(
      selection.getRangeElements()[0].getStartOffset(),
      selection.getRangeElements()[0].getEndOffsetInclusive() + 1
    ),
    fontFamily: fontFamily,
    fontSize: element.getFontSize(offset) || 11,
    isBold: element.isBold(offset) || false,
    isItalic: element.isItalic(offset) || false,
    color: element.getForegroundColor(offset) || "#000000",
    background: bgColor
  };
  /*return {
    text: element.getText(),
    fontFamily: element.getFontFamily(offset) || "Arial",
    fontSize: element.getFontSize(offset) || 11,
    isBold: element.isBold(offset),
    isItalic: element.isItalic(offset)
  };*/
}

function insertRotatedImage(base64Data, targetWidth, targetHeight, formatData) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) return;

  const rangeElements = selection.getRangeElements();
  const element = rangeElements[0].getElement().asText();

  const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
  const blob = Utilities.newBlob(bytes, 'image/png', 'rotated_text.png');

  const paragraph = element.getParent().asParagraph();
  const offset = paragraph.getChildIndex(element);

  element.setText("");
  const image = paragraph.insertInlineImage(offset, blob);
  image.setAltDescription("ROTATED_JSON:" + JSON.stringify(formatData));

  paragraph
    .setSpacingBefore(0)
    .setSpacingAfter(0)
    .setLineSpacing(1.0);

  if (targetWidth && targetHeight) {
    image
      .setWidth(targetWidth)
      .setHeight(targetHeight);
  }
}

function restoreTextFromImage() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) {
    DocumentApp.getUi().alert('Bitte klicke zuerst das Bild an, das du zurückverwandeln möchtest.');
    return;
  }

  const rangeElements = selection.getRangeElements();
  const rangeElement = rangeElements[0];
  const element = rangeElement.getElement();

  let img = null;
  if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
    img = element.asInlineImage();
  } else if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
    const numChildren = element.asParagraph().getNumChildren();
    for (let i = 0; i < numChildren; i++) {
      if (element.asParagraph().getChild(i).getType() === DocumentApp.ElementType.INLINE_IMAGE) {
        img = element.asParagraph().getChild(i).asInlineImage();
        break;
      }
    }
  }

  if (img) {
    const alt = img.getAltDescription() || "";

    if (alt.indexOf("ROTATED_JSON:") === 0) {
      try {
        const jsonData = JSON.parse(alt.replace("ROTATED_JSON:", ""));
        const parent = img.getParent();
        const offset = parent.getChildIndex(img);

        img.removeFromParent();

        const newText = parent.asParagraph().insertText(offset, jsonData.text);

        if (jsonData.fontFamily) newText.setFontFamily(jsonData.fontFamily);
        if (jsonData.fontSize) newText.setFontSize(jsonData.fontSize);
        if (jsonData.color) newText.setForegroundColor(jsonData.color);
        newText.setBold(jsonData.isBold || false);
        newText.setItalic(jsonData.isItalic || false);

        parent.asParagraph().setSpacingBefore(null).setSpacingAfter(null).setLineSpacing(1.15);

      } catch (e) {
        DocumentApp.getUi().alert("Fehler beim Verarbeiten der Bild-Daten.");
      }
    } else {
      DocumentApp.getUi().alert("Dieses Bild enthält keine gespeicherten Text-Daten.");
    }
  } else {
    DocumentApp.getUi().alert("Kein Bild zum Wiederherstellen gefunden. Bitte klicke das Bild direkt an.");
  }
}

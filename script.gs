function onOpen(doc, startNum, endText, footer) {
  var sheet = SpreadsheetApp.openByUrl('/* sheet url */');  // master sheet of images

  // loops through diferent images
  var increment = startNum;
  while (true) {
    var imageID = sheet.getRange('C' + increment).getValue(); 
    var targetAltText = sheet.getRange('D' + increment).getValue();
    
    Logger.log(targetAltText);
    if (targetAltText === endText) {
      break;
    }

    refreshHeaderImage(doc, imageID, targetAltText);
    refreshBodyImage(doc, imageID, targetAltText);
    if (footer === true) {
      refreshFooterImage(doc, imageID, targetAltText);
    }
    increment++;
  }
}

function refreshHeaderImage(doc, imageID, targetAltText) {
  var header;
  var imageReplaced = false;

  for (var k = 0; k < 2; k++) {
    if (k === 0){ // first page header
      var temp = doc.getHeader();
      header = temp.getParent().getChild(2).asHeaderSection();
    } else { // rest of headers
      header = doc.getHeader();
    } 

    const numChildren = header.getNumChildren();

    // get images in headers if first page does not have different header
    for (var i = 0; i < numChildren; i++) {
      const child = header.getChild(i);

      // check for inline image
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const paragraph = child.asParagraph();
        const paragraphChildren = paragraph.getNumChildren();

        // Loop through the elements in the paragraph
        for (var j = 0; j < paragraphChildren; j++) {
          const element = paragraph.getChild(j);

          // Check for inline images
          if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
            Logger.log("Inline found in header paragraph");
            const image = element.asInlineImage();
            const altText = image.getAltDescription();

            // found the image and want to replace it with new one
            if (altText === targetAltText) {
              replaceImage(image,imageID,targetAltText,false,paragraph,element,j);
              imageReplaced = true;
            }
          }
        }
      }
    }
  }

  if (!imageReplaced) {
    Logger.log(`No image in header found with alt text "${targetAltText}".`);
  }
}

function refreshBodyImage(doc, imageID, targetAltText) {
  // get body
  const body = doc.getBody();
  const numChildren = body.getNumChildren();
  var imageReplaced = false;
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);

    // check for image in list (eg. bullet points)
    if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
      const list = child.asListItem();
      const listChildren = list.getNumChildren();

      // Loop through the elements in the paragraph
      for (var j = 0; j < listChildren; j++) {
        const element = list.getChild(j);

        // check for inline image
        if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          Logger.log("Inline found in body list");
          const image = element.asInlineImage();
          const altText = image.getAltDescription();

          // found the image and want to replace it with new one
          if (altText === targetAltText) {
            replaceImage(image,imageID,targetAltText,false,list,element, j);
            imageReplaced = true;
          }
        }

        // POSITIONED IMAGES - WORK IN PROGRESS (GET ALT DESC NOT WORKING)
        /*if (element.getType() === DocumentApp.ElementType.POSITIONED_IMAGE) {
          if (altText === targetAltText) {
          }
        }*/
      }
    }

    // Check for image in a paragraph
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = child.asParagraph();
      const paragraphChildren = paragraph.getNumChildren();

      // Loop through the elements in the paragraph
      for (var j = 0; j < paragraphChildren; j++) {
        const element = paragraph.getChild(j);

        // Check for inline images
        if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          Logger.log("Inline found in body paragraph");
          const image = element.asInlineImage();
          const altText = image.getAltDescription();

          // found the image and want to replace it with new one
          if (altText === targetAltText) {
            replaceImage(image,imageID,targetAltText,false, paragraph, element,j);
            imageReplaced = true;
          }
        }
      }
    }
  }

  if (!imageReplaced) {
    Logger.log(`No image found with alt text "${targetAltText}".`);
  }
}

function refreshFooterImage(doc, imageID, targetAltText) {
  var footer;
  var imageReplaced = false;

  for (var k = 0; k < 2; k++) {
    if (k === 0){ // first page footer;
      var temp = doc.getFooter();
      footer = temp.getParent().getChild(4).asFooterSection();
    } else { // rest of footers
      footer = doc.getFooter();
    } 
    
    const numChildren = footer.getNumChildren();

    // get images in footer if first page does not have different footer
    for (var i = 0; i < numChildren; i++) {
      const child = footer.getChild(i);

      // check for inline image
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const paragraph = child.asParagraph();
        const paragraphChildren = paragraph.getNumChildren();

        // Loop through the elements in the paragraph
        for (var j = 0; j < paragraphChildren; j++) {
          const element = paragraph.getChild(j);

          // Check for inline images
          if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
            Logger.log("Inline found in footer paragraph");
            const image = element.asInlineImage();
            const altText = image.getAltDescription();

            // found the image and want to replace it with new one
            if (altText === targetAltText) {
              replaceImage(image,imageID,targetAltText,false,paragraph,element,j);
              imageReplaced = true;
            }
          }
        }
      }
    }
  }

  if (!imageReplaced) {
    Logger.log(`No image in header found with alt text "${targetAltText}".`);
  }
}

function replaceImage(image, imageID, targetAltText, position, parent, child, j) {
  // get old dimensions and positioning
  var oldWidth = image.getWidth();
  var oldHeight = image.getHeight();
  var oldX; 
  var oldY;
  var oldLayout;
  var newImage;
      

  if (position === true) {
    oldX = image.getLeftOffset();
    oldY = image.getTopOffset();
    oldLayout = image.getLayout();

    // remove old image
    let id = image.getId();
    parent.removePositionedImage(id);

  } else {
    // remove old image
    parent.removeChild(child);
  }

  // add new image
  const newImageBlob = DriveApp.getFileById(imageID).getBlob();

  if (position === true) {
    newImage = parent.addPositionedImage(newImageBlob);
    newImage.setLeftOffset(oldX);
    newImage.setTopOffset(oldY);
    newImage.setLayout(oldLayout);
  } else {
    newImage = parent.insertInlineImage(j, newImageBlob);
  }

  // Set the new image's size to match the old image's dimensions
  newImage.setWidth(oldWidth);
  newImage.setHeight(oldHeight);
            
  // Set the alt text for the new image
  if (newImage.getType() ===  DocumentApp.ElementType.INLINE_IMAGE) {
    newImage.setAltDescription(targetAltText);
  } 
}

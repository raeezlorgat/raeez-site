/**
 * Google Apps Script: Google Doc Proxy
 * 
 * SETUP (5 minutes):
 * 
 * 1. Go to https://script.google.com
 * 2. Click "New Project"
 * 3. Delete everything in the editor and paste this entire file
 * 4. Click the disk icon to save (or Ctrl+S)
 * 5. Click "Deploy" → "New deployment"
 * 6. Click the gear icon next to "Select type" → choose "Web app"
 * 7. Set:
 *    - Description: "Google Doc Proxy"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 8. Click "Deploy"
 * 9. Click "Authorize access" and follow the prompts
 *    (You may see a warning - click "Advanced" → "Go to [project name]")
 * 10. Copy the Web app URL (looks like: https://script.google.com/macros/s/XXXX/exec)
 * 11. Use that URL in your HTML file
 * 
 * USAGE:
 * https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?id=GOOGLE_DOC_ID
 */

function doGet(e) {
  const docId = e.parameter.id;
  
  if (!docId) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Missing ?id= parameter' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const html = convertToHtml(body);
    
    return ContentService
      .createTextOutput(html)
      .setMimeType(ContentService.MimeType.HTML);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        error: 'Failed to fetch document', 
        message: error.message 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function convertToHtml(body) {
  const numChildren = body.getNumChildren();
  let html = '';
  let inList = false;
  let listTag = 'ul';
  
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    const type = child.getType();
    
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      if (inList) {
        html += '</' + listTag + '>\n';
        inList = false;
      }
      html += processParagraph(child.asParagraph());
      
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = child.asListItem();
      const glyphType = listItem.getGlyphType();
      const isOrdered = (glyphType === DocumentApp.GlyphType.NUMBER ||
                         glyphType === DocumentApp.GlyphType.LATIN_UPPER ||
                         glyphType === DocumentApp.GlyphType.LATIN_LOWER ||
                         glyphType === DocumentApp.GlyphType.ROMAN_UPPER ||
                         glyphType === DocumentApp.GlyphType.ROMAN_LOWER);
      const newListTag = isOrdered ? 'ol' : 'ul';
      
      if (!inList) {
        listTag = newListTag;
        html += '<' + listTag + '>\n';
        inList = true;
      } else if (listTag !== newListTag) {
        html += '</' + listTag + '>\n';
        listTag = newListTag;
        html += '<' + listTag + '>\n';
      }
      
      html += '  <li>' + processText(listItem) + '</li>\n';
      
    } else if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
      if (inList) {
        html += '</' + listTag + '>\n';
        inList = false;
      }
      html += '<hr>\n';
    }
  }
  
  if (inList) {
    html += '</' + listTag + '>\n';
  }
  
  return html;
}

function processParagraph(para) {
  const text = para.getText().trim();
  if (!text) return '';
  
  const heading = para.getHeading();
  const content = processText(para);
  
  if (heading === DocumentApp.ParagraphHeading.HEADING1) {
    return '<h1>' + content + '</h1>\n';
  } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
    return '<h2>' + content + '</h2>\n';
  } else if (heading === DocumentApp.ParagraphHeading.HEADING3) {
    return '<h3>' + content + '</h3>\n';
  } else if (heading === DocumentApp.ParagraphHeading.HEADING4) {
    return '<h4>' + content + '</h4>\n';
  } else {
    // Check if it's bold (likely a heading in Google Docs style)
    if (para.getNumChildren() > 0) {
      const firstChild = para.getChild(0);
      if (firstChild.getType() === DocumentApp.ElementType.TEXT) {
        const textEl = firstChild.asText();
        if (textEl.isBold(0) && text.length < 120) {
          return '<h2>' + content + '</h2>\n';
        }
      }
    }
    return '<p>' + content + '</p>\n';
  }
}

function processText(element) {
  let result = '';
  const numChildren = element.getNumChildren();
  
  for (let i = 0; i < numChildren; i++) {
    const child = element.getChild(i);
    
    if (child.getType() === DocumentApp.ElementType.TEXT) {
      result += processTextElement(child.asText());
    }
  }
  
  return result;
}

function processTextElement(textEl) {
  const text = textEl.getText();
  if (!text) return '';
  
  let result = '';
  let i = 0;
  
  while (i < text.length) {
    const startIndex = i;
    const isBold = textEl.isBold(i);
    const isItalic = textEl.isItalic(i);
    const linkUrl = textEl.getLinkUrl(i);
    
    while (i < text.length &&
           textEl.isBold(i) === isBold &&
           textEl.isItalic(i) === isItalic &&
           textEl.getLinkUrl(i) === linkUrl) {
      i++;
    }
    
    let segment = escapeHtml(text.substring(startIndex, i));
    
    if (isBold && !linkUrl) {
      segment = '<strong>' + segment + '</strong>';
    }
    if (isItalic) {
      segment = '<em>' + segment + '</em>';
    }
    if (linkUrl) {
      segment = '<a href="' + escapeHtml(linkUrl) + '" target="_blank" rel="noopener">' + segment + '</a>';
    }
    
    result += segment;
  }
  
  return result;
}

function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

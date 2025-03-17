// aiOps.gs

/**
 * Global mapping for heading configurations
 */
const HEADING_MAP = {
  [DocumentApp.ParagraphHeading.HEADING1]: { prefix: '<h1>', suffix: '</h1>', addSpace: true },
  [DocumentApp.ParagraphHeading.HEADING2]: { prefix: '<h2>', suffix: '</h2>', addSpace: true },
  [DocumentApp.ParagraphHeading.HEADING3]: { prefix: '<h3>', suffix: '</h3>', addSpace: true },
  [DocumentApp.ParagraphHeading.HEADING4]: { prefix: '<button>', suffix: '</button>', addSpace: false },
  [DocumentApp.ParagraphHeading.HEADING5]: { prefix: '<label>', suffix: '</label>', addSpace: false },
  [DocumentApp.ParagraphHeading.HEADING6]: { prefix: '<p>', suffix: '</p>', addSpace: false }
};

/**
 * Comment patterns to convert to React comments
 */
const COMMENT_PATTERNS = [
  /^\/\/(.*)/,      // JS single line
  /^\/\*(.*)/,      // JS multi line
  /^#(.*)/,         // Python/Ruby style
  /^<!--(.*)$/,     // HTML style
];

function convertToReactComment(text) {
  for (const pattern of COMMENT_PATTERNS) {
    const match = text.match(pattern);
    if (match) {
      const commentContent = match[1].trim();
      return `{/* ${commentContent} */}`;
    }
  }
  return text;
}

function getAllTags() {
  return Object.values(HEADING_MAP).reduce((tags, { prefix, suffix }) => {
    tags.push(prefix, suffix);
    return tags;
  }, []);
}


function dropHtml(params = {}) {
  const body = DocumentApp.getActiveDocument().getBody();
  const startTime = new Date().getTime();
  
  try {
    // Get filtered paragraphs
    const htmlParagraphs = body.getParagraphs()
      .filter(para => para.getHeading() !== DocumentApp.ParagraphHeading.NORMAL)
      .map(para => {
        const level = para.getHeading();
        let text = para.getText().trim();
        
        if (text && HEADING_MAP[level]) {
          // First check if it's a comment
          const isComment = COMMENT_PATTERNS.some(pattern => pattern.test(text));
          if (isComment) {
            return {
              text: convertToReactComment(text),
              level,
              addSpace: HEADING_MAP[level].addSpace,
              isComment: true
            };
          }
          
          // If not a comment, proceed with HTML tags
          const { prefix, suffix, addSpace } = HEADING_MAP[level];
          return {
            text: prefix + text + suffix,
            level,
            addSpace,
            isComment: false
          };
        }
        return null;
      })
      .filter(Boolean);

    // Handle clipboard copy if requested
    if (params.copyToClipboard) {
      const content = htmlParagraphs
        .map(p => p.text)
        .join('\n')
        .trim();
      
      return copyHtmlToClipboard(content);

      // return copyHtmlToClipboard();
      // const clipboardContent = htmlParagraphs
      //   .map(p => p.text)
      //   .join('\n')
      //   .trim();
      
      // return {
      //   success: true,
      //   clipboardContent,
      //   executionTime: new Date().getTime() - startTime
      // };
    }

    // Insert paragraphs and apply formatting
    const insertPosition = params.position === 'start' ? 0 : body.getNumChildren();
    let currentPosition = insertPosition;

    htmlParagraphs.forEach((para, index) => {
      // Add space before H1/H2 if needed
      if (para.addSpace && (para.level === DocumentApp.ParagraphHeading.HEADING1 || 
          para.level === DocumentApp.ParagraphHeading.HEADING2)) {
        body.insertParagraph(currentPosition++, '');
      }
      
      // Insert the paragraph
      const newPara = body.insertParagraph(currentPosition++, para.text);
      newPara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      
      // Apply bold formatting for H1-H3 (non-comment content only)
      if (!para.isComment && 
          (para.level === DocumentApp.ParagraphHeading.HEADING1 || 
           para.level === DocumentApp.ParagraphHeading.HEADING2 || 
           para.level === DocumentApp.ParagraphHeading.HEADING3)) {
        const text = newPara.editAsText();
        if (para.isComment) {
          text.setBold(0, para.text.length - 1, true);
        } else {
          const startTag = para.text.indexOf('>') + 1;
          const endTag = para.text.lastIndexOf('<');
          text.setBold(startTag, endTag - 1, true);
        }
      }
    });
    body.insertParagraph(currentPosition, '');  // Adds empty paragraph at the end


    return {
      success: true,
      executionTime: new Date().getTime() - startTime
    };

  } catch (error) {
    Logger.log('Error in dropHtml:', error);
    return {
      success: false,
      error: error.toString(),
      executionTime: new Date().getTime() - startTime
    };
  }
}

function stripHtml(params = {}) {
  const body = DocumentApp.getActiveDocument().getBody();
  const startTime = new Date().getTime();
  
  try {
    const mode = params.all ? 'all' : 'tags';
    
if (mode === 'all') {
  // Find paragraphs containing any HTML tags
  const htmlRegex = /<[^>]+>/;
  let removedCount = 0;
  let lastHtmlIndex = -1;
  let consecutiveEmpty = 0;
  
  for (let i = body.getNumChildren() - 1; i >= 0; i--) {
    const element = body.getChild(i);
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = element.asParagraph();
      const text = para.getText();
      const isEmpty = !text || text.trim() === '';
      const hasHtml = htmlRegex.test(text) || text.includes('{/*');

      if (hasHtml) {
        // Reset empty counter when we find HTML
        lastHtmlIndex = i;
        consecutiveEmpty = 0;
        body.removeChild(element);
        removedCount++;
      } else if (isEmpty) {
        // Remove if it's empty and either between HTML or part of consecutive empty paragraphs
        if (lastHtmlIndex !== -1 || consecutiveEmpty > 0) {
          body.removeChild(element);
          removedCount++;
        }
        consecutiveEmpty++;
      } else {
        // Non-empty, non-HTML paragraph resets consecutive counter
        consecutiveEmpty = 0;
      }
    }
  }
  
  return {
    success: true,
    mode: 'all',
    removedCount,
    executionTime: new Date().getTime() - startTime
  };
}
    // Handle strip tags only
    let replacedCount = 0;
    const paragraphs = body.getParagraphs();
    
    paragraphs.forEach(para => {
      const text = para.getText();
      const hasHtml = /<[^>]+>/.test(text);
      
      if (hasHtml || text.includes('{/*')) {
        // Remove all HTML tags and comments
        const cleanText = text
          .replace(/<[^>]+>/g, '')  // Remove HTML tags
          .replace(/\{\/\*.*?\*\/\}/g, ''); // Remove React comments
        
        if (cleanText.trim()) {
          para.setText(cleanText);
          replacedCount++;
        } else {
          // If only whitespace remains, remove the paragraph
          para.removeFromParent();
        }
      }
    });

    return {
      success: true,
      mode: 'tags',
      replacedCount,
      executionTime: new Date().getTime() - startTime
    };

  } catch (error) {
    Logger.log('Error in stripHtml:', error);
    return {
      success: false,
      error: error.toString(),
      executionTime: new Date().getTime() - startTime
    };
  }
}

// Menu helper functions
function stripHtmlTags() {
  return stripHtml({ tags: true });
}

function stripHtmlAll() {
  return stripHtml({ all: true });
}

function copyHtmlToClipboard(content) {
  if (!content) {
    DocumentApp.getUi().alert('No HTML content to copy.');
    return false;
  }
  
  const escapedContent = content.replace(/'/g, "\\'").replace(/\n/g, "\\n");
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      pre { 
        background: #f5f5f5; 
        padding: 10px; 
        border-radius: 4px;
        white-space: pre-wrap;
        word-wrap: break-word;
        margin-top: 10px;
      }
      .no-select {
        user-select: none;
        -webkit-user-select: none;
        -moz-user-select: none;
        -ms-user-select: none;
      }
      .copy-button {
        background: #4285f4;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 4px;
        cursor: pointer;
        margin-bottom: 10px;
      }
      .copy-button:hover {
        background: #3b78e7;
      }
    </style>

    <div class="no-select">
      <button class="copy-button no-select" onclick="copyContent()">Copy to Clipboard</button>
      <div class="instructions">Or manually select and copy (Cmd/Ctrl+C):</div>
    </div>
    
    <pre id="content">${content.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</pre>

    <script>
      const rawContent = '${escapedContent}';
      
      function copyContent() {
        navigator.clipboard.writeText(rawContent)
          .then(() => {
            const btn = document.querySelector('.copy-button');
            btn.textContent = 'Copied!';
            setTimeout(() => btn.textContent = 'Copy to Clipboard', 500);
          })
          .catch(err => {
            console.error('Failed to copy:', err);
            alert('Please use manual selection and copy instead');
          });
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(400);
  
  DocumentApp.getUi().showModalDialog(html, 'Copy HTML Content');
  return true;
}


// function copyHtmlToClipboard() {
//   // const result = generateHtml();
//   const result = dropHtml({ copyToClipboard: true });
  
//   if (result.success && result.clipboardContent) {
//     const escapedContent = result.clipboardContent.replace(/'/g, "\\'").replace(/\n/g, "\\n");
    
//     const html = HtmlService.createHtmlOutput(`
//       <style>
//         body { font-family: Arial, sans-serif; padding: 15px; }
//         pre { 
//           background: #f5f5f5; 
//           padding: 10px; 
//           border-radius: 4px;
//           white-space: pre-wrap;
//           word-wrap: break-word;
//           margin-top: 10px;
//         }
//         .no-select {
//           user-select: none;
//           -webkit-user-select: none;
//           -moz-user-select: none;
//           -ms-user-select: none;
//         }
//         .copy-button {
//           background: #4285f4;
//           color: white;
//           border: none;
//           padding: 8px 16px;
//           border-radius: 4px;
//           cursor: pointer;
//           margin-bottom: 10px;
//         }
//         .copy-button:hover {
//           background: #3b78e7;
//         }
//       </style>

//       <div class="no-select">
//         <button class="copy-button no-select" onclick="copyContent()">Copy to Clipboard</button>
//         <div class="instructions">Or manually select and copy (Cmd/Ctrl+C):</div>
//       </div>
      
//       <pre id="content">${result.clipboardContent.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</pre>

//       <script>
//         // Store the raw content in a variable to ensure we copy the unescaped version
//         const rawContent = '${escapedContent}';
        
//         function copyContent() {
//           navigator.clipboard.writeText(rawContent)
//             .then(() => {
//               const btn = document.querySelector('.copy-button');
//               btn.textContent = 'Copied!';
//               setTimeout(() => btn.textContent = 'Copy to Clipboard', 500);
//             })
//             .catch(err => {
//               console.error('Failed to copy:', err);
//               alert('Please use manual selection and copy instead');
//             });
//         }
//       </script>
//     `)
//     .setWidth(500)
//     .setHeight(400);
    
//     DocumentApp.getUi().showModalDialog(html, 'Copy HTML Content');
//     return true;
//   }
  
//   DocumentApp.getUi().alert('Failed to generate HTML content.');
//   return false;
// }

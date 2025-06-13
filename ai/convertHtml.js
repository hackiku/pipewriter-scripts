// convertHtml.gs

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

/**
 * NEW: Download HTML content as a complete HTML file
 * Generates a complete HTML document with the wireframe content
 */
function downloadHtmlFile() {
  const body = DocumentApp.getActiveDocument().getBody();
  const startTime = new Date().getTime();
  
  try {
    // Get filtered paragraphs (same logic as dropHtml)
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

    if (htmlParagraphs.length === 0) {
      DocumentApp.getUi().alert('No HTML content found to download. Please add some heading content first.');
      return {
        success: false,
        error: 'No HTML content found',
        executionTime: new Date().getTime() - startTime
      };
    }

    // Generate complete HTML document
    const htmlContent = htmlParagraphs.map(p => p.text).join('\n');
    const docTitle = DocumentApp.getActiveDocument().getName() || 'wireframe';
    
    const completeHtml = generateCompleteHtmlDocument(htmlContent, docTitle);
    
    // Show download dialog
    return showDownloadDialog(completeHtml, docTitle);

  } catch (error) {
    Logger.log('Error in downloadHtmlFile:', error);
    return {
      success: false,
      error: error.toString(),
      executionTime: new Date().getTime() - startTime
    };
  }
}

/**
 * Generate a complete HTML document with CSS and structure
 */
function generateCompleteHtmlDocument(htmlContent, title) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${title}</title>
    <style>
        /* Pipewriter Wireframe Styles */
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f8f9fa;
        }
        
        h1, h2, h3 {
            color: #333;
            margin-top: 2em;
            margin-bottom: 1em;
        }
        
        h1 {
            font-size: 2.5em;
            border-bottom: 3px solid #007bff;
            padding-bottom: 0.3em;
        }
        
        h2 {
            font-size: 2em;
            color: #495057;
        }
        
        h3 {
            font-size: 1.5em;
            color: #6c757d;
        }
        
        button {
            background-color: #007bff;
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 16px;
            margin: 10px 5px;
            transition: background-color 0.3s;
        }
        
        button:hover {
            background-color: #0056b3;
        }
        
        label {
            display: inline-block;
            background-color: #e9ecef;
            padding: 8px 16px;
            border-radius: 4px;
            margin: 5px;
            color: #495057;
            font-weight: 500;
        }
        
        p {
            color: #6c757d;
            margin: 1em 0;
        }
        
        /* Comment styles */
        .comment {
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 4px;
            padding: 10px;
            margin: 10px 0;
            color: #856404;
            font-style: italic;
        }
        
        /* Wireframe indication */
        .wireframe-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            text-align: center;
        }
        
        .wireframe-header h1 {
            margin: 0;
            border: none;
            color: white;
        }
        
        .wireframe-footer {
            margin-top: 40px;
            padding: 20px;
            background-color: #e9ecef;
            border-radius: 8px;
            text-align: center;
            color: #6c757d;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="wireframe-header">
        <h1>🎯 ${title}</h1>
        <p>Wireframe generated by Pipewriter</p>
    </div>
    
    <main>
${htmlContent}
    </main>
    
    <div class="wireframe-footer">
        <p>Generated on ${new Date().toLocaleDateString()} • Made with Pipewriter for Google Docs</p>
    </div>
</body>
</html>`;
}

/**
 * Show download dialog with the HTML file
 */
function showDownloadDialog(htmlContent, filename) {
  const escapedContent = htmlContent.replace(/'/g, "\\'").replace(/\n/g, "\\n");
  const safeFilename = filename.replace(/[^a-zA-Z0-9-_]/g, '_');
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { 
        font-family: Arial, sans-serif; 
        padding: 15px; 
        line-height: 1.5;
      }
      .preview {
        background: #f5f5f5; 
        padding: 15px; 
        border-radius: 4px;
        margin: 15px 0;
        max-height: 200px;
        overflow-y: auto;
        white-space: pre-wrap;
        word-wrap: break-word;
        font-size: 12px;
        border: 1px solid #ddd;
      }
      .download-button {
        background: #4285f4;
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 16px;
        margin: 10px 5px;
        display: inline-block;
      }
      .download-button:hover {
        background: #3b78e7;
      }
      .success-message {
        background: #d4edda;
        color: #155724;
        padding: 10px;
        border-radius: 4px;
        margin: 10px 0;
        display: none;
      }
      .instructions {
        background: #e2e6ea;
        padding: 10px;
        border-radius: 4px;
        margin: 10px 0;
        font-size: 14px;
        color: #383d41;
      }
      .filename-info {
        font-weight: bold;
        color: #495057;
        margin: 10px 0;
      }
    </style>

    <h3>📄 Download HTML File</h3>
    
    <div class="filename-info">
      File: ${safeFilename}.html
    </div>
    
    <div class="instructions">
      Click the download button below to save your wireframe as a complete HTML file. The file includes styling and is ready to view in any web browser.
    </div>
    
    <button class="download-button" onclick="downloadHtml()">
      📥 Download HTML File
    </button>
    
    <div class="success-message" id="success">
      ✅ Download started! Check your downloads folder.
    </div>
    
    <h4>Preview:</h4>
    <div class="preview">${htmlContent.replace(/</g, '&lt;').replace(/>/g, '&gt;').substring(0, 1000)}${htmlContent.length > 1000 ? '...' : ''}</div>

    <script>
      function downloadHtml() {
        try {
          // Create the complete HTML content
          const htmlContent = '${escapedContent}';
          const filename = '${safeFilename}.html';
          
          // Create blob and download
          const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
          const url = window.URL.createObjectURL(blob);
          
          // Create download link
          const a = document.createElement('a');
          a.href = url;
          a.download = filename;
          a.style.display = 'none';
          
          // Trigger download
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          
          // Clean up
          window.URL.revokeObjectURL(url);
          
          // Show success message
          document.getElementById('success').style.display = 'block';
          
          // Update button
          const btn = document.querySelector('.download-button');
          const originalText = btn.innerHTML;
          btn.innerHTML = '✅ Downloaded!';
          setTimeout(() => {
            btn.innerHTML = originalText;
          }, 2000);
          
        } catch (error) {
          console.error('Download failed:', error);
          alert('Download failed. Please try again or contact support.');
        }
      }
    </script>
  `)
  .setWidth(600)
  .setHeight(500);
  
  DocumentApp.getUi().showModalDialog(html, 'Download HTML File');
  
  return {
    success: true,
    message: 'Download dialog opened',
    filename: safeFilename + '.html'
  };
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
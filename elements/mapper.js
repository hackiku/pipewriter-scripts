// elements/mapper.js - Build element mapping from master documents

/**
 * Master document IDs for development/testing
 * These are the copies you provided to avoid affecting production
 */
const DEV_MASTER_DOCS = {
	light: "1yDx1RzvNqTHpPkdfml1iHn7H32a3kvvZqIm8NTmIKPk",
	dark: "1tD1NcdqhEyTy3N6K2Syu4f6Pn-aDYET9GTlB1hlBjxM"
};

/**
 * Scans master documents and builds element mapping
 * @param {boolean} useDevDocs - Whether to use development document IDs
 * @returns {Object} Mapping of element IDs to their locations and properties
 */
function scanMasterDocuments(useDevDocs = true) {
	const docIds = useDevDocs ? DEV_MASTER_DOCS : MASTER_DOCS;
	const elementMap = {};

	for (const theme of ['light', 'dark']) {
		try {
			Logger.log(`Scanning ${theme} theme document...`);
			const masterDoc = DocumentApp.openById(docIds[theme]);
			const masterBody = masterDoc.getBody();
			const numElements = masterBody.getNumChildren();

			let currentElementId = null;

			for (let i = 0; i < numElements; i++) {
				const element = masterBody.getChild(i);

				if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
					// Only consider paragraphs with text as potential element identifiers
					const text = element.getText().trim();
					if (text && !text.startsWith('#') && !text.startsWith('//')) {
						currentElementId = text;
						Logger.log(`Found potential element identifier: "${currentElementId}" at index ${i}`);
					}
				} else if (currentElementId && element.getType() == DocumentApp.ElementType.TABLE) {
					// Extract table properties
					const tableProperties = extractTableProperties(element);

					// Store element information in map
					elementMap[currentElementId] = {
						theme: theme,
						docId: docIds[theme],
						index: i,
						properties: tableProperties
					};

					Logger.log(`Mapped "${currentElementId}" to index ${i} with ${tableProperties.numRows} rows × ${tableProperties.numCols} columns`);

					// Reset current element ID to avoid duplicate mappings
					currentElementId = null;
				}
			}

			Logger.log(`Completed scan of ${theme} theme document`);
		} catch (error) {
			Logger.log(`Error scanning ${theme} document: ${error}`);
		}
	}

	return elementMap;
}

/**
 * Extracts useful properties from a table element
 * @param {TableElement} table - The table to analyze
 * @returns {Object} Table properties
 */
function extractTableProperties(table) {
	const properties = {
		numRows: table.getNumRows(),
		numCols: table.getRow(0).getNumCells(),
		hasBorders: table.getBorderWidth() > 0,
		backgroundColor: null,
		cellContents: []
	};

	// Sample first few cells for content types
	const maxSampleCells = 5;
	let sampleCount = 0;

	// Check first row for cell contents
	const firstRow = table.getRow(0);
	for (let j = 0; j < properties.numCols && sampleCount < maxSampleCells; j++) {
		try {
			const cell = firstRow.getCell(j);
			const cellText = cell.getText().trim();
			const hasImage = cellContainsImage(cell);

			properties.cellContents.push({
				position: `0,${j}`,
				hasText: cellText.length > 0,
				text: cellText.substring(0, 30) + (cellText.length > 30 ? "..." : ""),
				hasImage: hasImage,
				alignment: cell.getHorizontalAlignment()
			});

			sampleCount++;
		} catch (error) {
			Logger.log(`Error examining cell: ${error}`);
		}
	}

	// Check if there's background color on first cell
	try {
		properties.backgroundColor = firstRow.getCell(0).getBackgroundColor();
	} catch (error) {
		// Ignore errors getting background color
	}

	return properties;
}

/**
 * Checks if a table cell contains an image
 * @param {TableCell} cell - The cell to check
 * @returns {boolean} True if the cell contains an image
 */
function cellContainsImage(cell) {
	try {
		// Iterate through cell children looking for images
		for (let i = 0; i < cell.getNumChildren(); i++) {
			const child = cell.getChild(i);
			if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
				const para = child.asParagraph();
				// Check for inline images in paragraph
				for (let j = 0; j < para.getNumChildren(); j++) {
					if (para.getChild(j).getType() === DocumentApp.ElementType.INLINE_IMAGE) {
						return true;
					}
				}
			} else if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
				return true;
			}
		}
		return false;
	} catch (error) {
		Logger.log(`Error checking for images: ${error}`);
		return false;
	}
}

/**
 * Exports the element map as JSON for use in the application
 * @param {Object} elementMap - The element mapping to export
 * @returns {string} JSON representation of the element map
 */
function exportElementMapAsJson(elementMap) {
	// Create simplified version without document IDs for client use
	const clientMap = {};

	Object.keys(elementMap).forEach(elementId => {
		const element = elementMap[elementId];
		clientMap[elementId] = {
			theme: element.theme,
			index: element.index,
			numRows: element.properties.numRows,
			numCols: element.properties.numCols
		};
	});

	return JSON.stringify(clientMap, null, 2);
}

/**
 * Generates a JavaScript module with the element mapping
 * @param {Object} elementMap - The element mapping to export
 * @returns {string} JavaScript code for element mapping
 */
function generateElementMapModule(elementMap) {
	let jsCode = "// Auto-generated element mapping\n\n";
	jsCode += "// Master document IDs\n";
	jsCode += "const MASTER_DOCS = {\n";
	jsCode += "  light: \"" + DEV_MASTER_DOCS.light + "\",\n";
	jsCode += "  dark: \"" + DEV_MASTER_DOCS.dark + "\"\n";
	jsCode += "};\n\n";

	jsCode += "// Element mapping\n";
	jsCode += "const ELEMENT_MAPPING = {\n";

	Object.keys(elementMap).forEach(elementId => {
		const element = elementMap[elementId];
		jsCode += `  "${elementId}": { docId: MASTER_DOCS.${element.theme}, index: ${element.index} },\n`;
	});

	jsCode += "};\n\n";

	jsCode += "/**\n";
	jsCode += " * Get element from master document using the element mapping\n";
	jsCode += " * @param {string} elementId - The ID of the element to retrieve\n";
	jsCode += " * @param {string} theme - Theme ('light' or 'dark')\n";
	jsCode += " * @returns {TableElement|null} The table element or null if not found\n";
	jsCode += " */\n";
	jsCode += "function getElementFromMapping(elementId, theme) {\n";
	jsCode += "  try {\n";
	jsCode += "    // Adjust elementId for dark theme if needed\n";
	jsCode += "    const adjustedElementId = theme === \"dark\" ? `${elementId}-dark` : elementId;\n";
	jsCode += "    \n";
	jsCode += "    // Look up element information in mapping\n";
	jsCode += "    const elementInfo = ELEMENT_MAPPING[adjustedElementId];\n";
	jsCode += "    if (!elementInfo) {\n";
	jsCode += "      throw new Error(`Element mapping not found for: ${adjustedElementId}`);\n";
	jsCode += "    }\n";
	jsCode += "    \n";
	jsCode += "    // Open document and get element directly by index\n";
	jsCode += "    const masterDoc = DocumentApp.openById(elementInfo.docId);\n";
	jsCode += "    const masterBody = masterDoc.getBody();\n";
	jsCode += "    const element = masterBody.getChild(elementInfo.index);\n";
	jsCode += "    \n";
	jsCode += "    if (element && element.getType() == DocumentApp.ElementType.TABLE) {\n";
	jsCode += "      return element.copy();\n";
	jsCode += "    } else {\n";
	jsCode += "      throw new Error(`Element at index ${elementInfo.index} is not a table`);\n";
	jsCode += "    }\n";
	jsCode += "  } catch (error) {\n";
	jsCode += "    Logger.log('Failed to get element from master: ' + error);\n";
	jsCode += "    return null;\n";
	jsCode += "  }\n";
	jsCode += "}\n";

	return jsCode;
}

/**
 * Test function to run the scanner and log results
 */
function testElementMapper() {
	Logger.log("Starting element mapping scan...");
	const elementMap = scanMasterDocuments(true); // Use dev docs

	// Log summary of findings
	Logger.log("\n--- ELEMENT MAPPING SUMMARY ---");
	Logger.log(`Total elements found: ${Object.keys(elementMap).length}`);

	const lightElements = Object.keys(elementMap).filter(id => elementMap[id].theme === 'light').length;
	const darkElements = Object.keys(elementMap).filter(id => elementMap[id].theme === 'dark').length;

	Logger.log(`Light theme elements: ${lightElements}`);
	Logger.log(`Dark theme elements: ${darkElements}`);

	// List all elements
	Logger.log("\n--- ELEMENT LIST ---");
	Object.keys(elementMap).sort().forEach(id => {
		const element = elementMap[id];
		Logger.log(`${id} (${element.theme}): Table with ${element.properties.numRows} rows × ${element.properties.numCols} columns`);
	});

	// Generate JS module
	Logger.log("\n--- ELEMENT MAP MODULE ---");
	const jsModule = generateElementMapModule(elementMap);
	Logger.log(jsModule);

	return elementMap;
}

// Export functions for use in other modules
var elementMapper = {
	scanMasterDocuments: scanMasterDocuments,
	exportElementMapAsJson: exportElementMapAsJson,
	generateElementMapModule: generateElementMapModule,
	testElementMapper: testElementMapper
};
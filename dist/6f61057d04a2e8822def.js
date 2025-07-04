/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("insert-text").onclick = insertText;
        document.getElementById("get-selection").onclick = getSelectedText;
        document.getElementById("format-text").onclick = formatSelectedText;
        document.getElementById("insert-table").onclick = insertTable;
        document.getElementById("get-document-info").onclick = getDocumentInfo;
    }
});

/**
 * Inserts sample text into the document
 */
async function insertText() {
    try {
        await Word.run(async (context) => {
            // Insert text at the end of the document
            const body = context.document.body;
            body.insertText("Hello from Word Integration Add-in! This text was inserted using Office.js APIs.", Word.InsertLocation.end);
            
            // Sync the context to execute the API calls
            await context.sync();
            
            showResult("Text inserted successfully!", "success");
        });
    } catch (error) {
        showResult(`Error: ${error.message}`, "error");
    }
}

/**
 * Gets the currently selected text
 */
async function getSelectedText() {
    try {
        await Word.run(async (context) => {
            // Get the current selection
            const selection = context.document.getSelection();
            selection.load("text");
            
            await context.sync();
            
            if (selection.text) {
                showResult(`Selected text: "${selection.text}"`, "success");
            } else {
                showResult("No text is currently selected.", "success");
            }
        });
    } catch (error) {
        showResult(`Error: ${error.message}`, "error");
    }
}

/**
 * Formats the selected text with bold and blue color
 */
async function formatSelectedText() {
    try {
        await Word.run(async (context) => {
            // Get the current selection
            const selection = context.document.getSelection();
            
            // Check if there's a selection
            selection.load("text");
            await context.sync();
            
            if (!selection.text) {
                showResult("Please select some text first.", "error");
                return;
            }
            
            // Apply formatting
            selection.font.bold = true;
            selection.font.color = "#0078d4"; // Microsoft blue
            selection.font.size = 14;
            
            await context.sync();
            
            showResult("Selected text formatted successfully!", "success");
        });
    } catch (error) {
        showResult(`Error: ${error.message}`, "error");
    }
}

/**
 * Inserts a simple table
 */
async function insertTable() {
    try {
        await Word.run(async (context) => {
            // Insert a table at the end of the document
            const body = context.document.body;
            const table = body.insertTable(3, 3, Word.InsertLocation.end, [
                ["Header 1", "Header 2", "Header 3"],
                ["Row 1, Col 1", "Row 1, Col 2", "Row 1, Col 3"],
                ["Row 2, Col 1", "Row 2, Col 2", "Row 2, Col 3"]
            ]);
            
            // Format the header row
            const headerRow = table.rows.getFirst();
            headerRow.font.bold = true;
            headerRow.font.color = "white";
            headerRow.cells.items.forEach(cell => {
                cell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
            });
            
            // Set table style
            table.styleFirstColumn = false;
            table.styleBandedRows = true;
            table.style = "Grid Table 4 - Accent 1";
            
            await context.sync();
            
            showResult("Table inserted successfully!", "success");
        });
    } catch (error) {
        showResult(`Error: ${error.message}`, "error");
    }
}

/**
 * Gets basic information about the document
 */
async function getDocumentInfo() {
    try {
        await Word.run(async (context) => {
            // Get document properties
            const doc = context.document;
            const body = doc.body;
            const paragraphs = body.paragraphs;
            const tables = body.tables;
            
            // Load properties
            paragraphs.load("items");
            tables.load("items");
            body.load("text");
            
            await context.sync();
            
            const info = {
                paragraphCount: paragraphs.items.length,
                tableCount: tables.items.length,
                characterCount: body.text.length,
                wordCount: body.text.split(/\s+/).filter(word => word.length > 0).length
            };
            
            const infoHtml = `
                <strong>Document Statistics:</strong><br>
                • Paragraphs: ${info.paragraphCount}<br>
                • Tables: ${info.tableCount}<br>
                • Characters: ${info.characterCount}<br>
                • Words: ${info.wordCount}
            `;
            
            showResult(infoHtml, "success");
        });
    } catch (error) {
        showResult(`Error: ${error.message}`, "error");
    }
}

/**
 * Shows result in the result section
 * @param {string} message - The message to display
 * @param {string} type - The type of message (success, error)
 */
function showResult(message, type) {
    const resultDiv = document.getElementById("result");
    const resultContent = document.getElementById("result-content");
    
    resultDiv.style.display = "block";
    resultContent.innerHTML = message;
    resultContent.className = type;
    
    // Auto-hide after 5 seconds for success messages
    if (type === "success") {
        setTimeout(() => {
            resultDiv.style.display = "none";
        }, 5000);
    }
}

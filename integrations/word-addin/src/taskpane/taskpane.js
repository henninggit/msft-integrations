/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Original buttons - remove IDs since we're using onclick in HTML
        // document.getElementById("insert-text").onclick = insertText;
        // document.getElementById("get-selection").onclick = getSelectedText;
        // document.getElementById("format-text").onclick = formatSelectedText;
        // document.getElementById("insert-table").onclick = insertTable;
        // document.getElementById("get-document-info").onclick = getDocumentInfo;
        
        // Initialize character count
        updateCharacterCount();
    }
});

// Copilot-style functions
function updateCharacterCount() {
    const input = document.getElementById('promptInput');
    const charCount = document.getElementById('charCount');
    if (input && charCount) {
        charCount.textContent = input.value.length;
    }
}

function closeCopilot() {
    // In a real implementation, this might close the task pane
    // For now, just clear the input
    document.getElementById('promptInput').value = '';
    updateCharacterCount();
}

async function generateContent() {
    const prompt = document.getElementById('promptInput').value;
    if (!prompt.trim()) {
        alert('Please enter a prompt first');
        return;
    }
    
    const button = document.querySelector('.generate-button');
    button.disabled = true;
    button.textContent = 'Generating...';
    
    try {
        // Simulate AI generation - in real implementation, call your AI backend
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        const generatedContent = generateMockContent(prompt);
        
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.insertText(generatedContent, Word.InsertLocation.replace);
            
            await context.sync();
            showResult('Generated content inserted successfully!', 'success');
        });
        
    } catch (error) {
        showResult('Error generating content: ' + error.message, 'error');
    } finally {
        button.disabled = false;
        button.textContent = 'Generate';
    }
}

function applySuggestion(type) {
    const suggestions = {
        'blog-intro': 'Write a compelling introduction for a blog post that hooks readers and clearly outlines the main benefits discussed in the article.',
        'summary': 'Create a concise summary of the selected text, highlighting key points and main takeaways.',
        'improve': 'Enhance the writing style and clarity of the selected text, making it more engaging and professional.'
    };
    
    document.getElementById('promptInput').value = suggestions[type] || '';
    updateCharacterCount();
}

function generateMockContent(prompt) {
    // Mock AI-generated content based on prompt
    if (prompt.toLowerCase().includes('introduction') || prompt.toLowerCase().includes('blog')) {
        return `# The Future of Remote Work: Transforming How We Connect and Collaborate

In an era where digital connectivity has redefined the boundaries of traditional workspaces, remote work has emerged as more than just a temporary solution—it's become a fundamental shift in how we approach professional life. This transformation brings with it a wealth of opportunities that extend far beyond the simple convenience of working from home.

Remote work offers unprecedented flexibility, allowing professionals to design their work environment around their personal needs and peak productivity hours. This autonomy not only enhances work-life balance but also opens doors to global talent pools, enabling companies to access diverse skills and perspectives regardless of geographical limitations.`;
    }
    
    return `Generated content based on your prompt: "${prompt}"

This is a mock implementation that would normally connect to your AI backend to generate relevant content. The generated text can be customized based on the user's specific requirements and integrated seamlessly into the Word document.`;
}

function showResult(message, type = 'success') {
    const resultDiv = document.getElementById('result');
    const resultContent = document.getElementById('result-content');
    
    if (resultDiv && resultContent) {
        resultContent.innerHTML = `<div class="${type}">${message}</div>`;
        resultDiv.style.display = 'block';
        
        // Hide after 3 seconds
        setTimeout(() => {
            resultDiv.style.display = 'none';
        }, 3000);
    }
}

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



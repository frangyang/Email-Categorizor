// categorizer.js
let isDebugMode = false;
let verboseLogging = false;
let categoryMapping = {};

// Configuration defaults
const defaultConfig = {
    subjectWeight: 10,  // Keywords in subject are 10x more important
    bodyWeight: 1,      // Base weight for body
    recencyMultiplier: 1.5,  // More recent emails get 1.5x weight
    scoreThreshold: 10  // Minimum score needed to apply category
};

Office.onReady(() => {
    document.getElementById('runButton').onclick = processFolder;
    loadFolders();
    loadCategoriesFromStorage();
    debugLog('Add-in initialized', 'info');
});

// Storage Functions
function saveCategoriestoStorage() {
    localStorage.setItem('categoryMapping', JSON.stringify(categoryMapping));
    debugLog('Categories saved to storage', 'info', categoryMapping);
}

function loadCategoriesFromStorage() {
    const saved = localStorage.getItem('categoryMapping');
    if (saved) {
        categoryMapping = JSON.parse(saved);
        debugLog('Categories loaded from storage', 'info', categoryMapping);
        renderCategories();
    }
}

// Category Management
function addCategory() {
    const categoryName = document.getElementById('categoryName').value.trim();
    const keywords = document.getElementById('keywordInput').value
        .split(',')
        .map(k => k.trim())
        .filter(k => k);

    if (!categoryName || keywords.length === 0) {
        debugLog('Invalid category input', 'error');
        return;
    }

    categoryMapping[categoryName] = keywords;
    saveCategoriestoStorage();
    renderCategories();
    
    // Clear inputs
    document.getElementById('categoryName').value = '';
    document.getElementById('keywordInput').value = '';
}

function deleteCategory(categoryName) {
    delete categoryMapping[categoryName];
    saveCategoriestoStorage();
    renderCategories();
    debugLog(`Category deleted: ${categoryName}`, 'info');
}

function renderCategories() {
    const container = document.getElementById('categoryList');
    container.innerHTML = '';

    Object.entries(categoryMapping).forEach(([category, keywords]) => {
        const div = document.createElement('div');
        div.className = 'category-item';
        div.innerHTML = `
            <div style="display: flex; justify-content: space-between;">
                <strong>${category}</strong>
                <button onclick="deleteCategory('${category}')">Delete</button>
            </div>
            <div class="keyword-list">
                ${keywords.map(k => `<span class="keyword-tag">${k}</span>`).join('')}
            </div>
        `;
        container.appendChild(div);
    });
}

// Email Processing
async function parseEmailChain(emailBody) {
    const parts = [];
    const emailHeaderRegex = /From:[\s\S]*?Sent:.*?((?=From:)|$)/gi;
    const matches = emailBody.match(emailHeaderRegex) || [];
    
    matches.forEach((part, index) => {
        const dateMatch = part.match(/Sent:.*?([\d]{1,2}\/[\d]{1,2}\/[\d]{4}|[\d]{4}-[\d]{2}-[\d]{2})/i);
        const timestamp = dateMatch ? new Date(dateMatch[1]).getTime() : Date.now() - (index * 86400000);
        parts.push({
            text: part,
            timestamp: timestamp
        });
    });

    return parts.sort((a, b) => b.timestamp - a.timestamp);
}

async function calculateCategoryScore(email, category, keywords) {
    const weights = {
        subject: Number(document.getElementById('subjectWeight').value) || defaultConfig.subjectWeight,
        body: Number(document.getElementById('bodyWeight').value) || defaultConfig.bodyWeight
    };

    let score = 0;
    const subject = email.subject.toLowerCase();
    
    // Score subject matches
    keywords.forEach(keyword => {
        const keywordLower = keyword.toLowerCase();
        const subjectMatches = (subject.match(new RegExp(keywordLower, 'g')) || []).length;
        score += subjectMatches * weights.subject;
        debugLog(`Subject match for "${keyword}": ${subjectMatches} occurrences`, 'verbose');
    });

    // Score body matches with recency weighting
    try {
        const bodyContent = (await Office.context.mailbox.item.body.getAsync('text')).value;
        const emailParts = await parseEmailChain(bodyContent);
        
        emailParts.forEach((part, index) => {
            const recencyWeight = defaultConfig.recencyMultiplier ** (emailParts.length - index - 1);
            
            keywords.forEach(keyword => {
                const keywordLower = keyword.toLowerCase();
                const matches = (part.text.toLowerCase().match(new RegExp(keywordLower, 'g')) || []).length;
                score += matches * weights.body * recencyWeight;
                debugLog(`Body match for "${keyword}" in part ${index + 1}: ${matches} occurrences (weight: ${recencyWeight})`, 'verbose');
            });
        });
    } catch (error) {
        debugLog('Error processing email body', 'error', error);
    }

    debugLog(`Final score for category "${category}": ${score}`, 'verbose');
    return score;
}

async function processEmail(email) {
    debugLog(`Processing email: ${email.subject}`, 'verbose');
    
    const scoreThreshold = Number(document.getElementById('scoreThreshold').value) || defaultConfig.scoreThreshold;
    let highestScore = 0;
    let bestCategory = null;
    const scores = {};

    // Calculate scores for each category
    for (const [category, keywords] of Object.entries(categoryMapping)) {
        const score = await calculateCategoryScore(email, category, keywords);
        scores[category] = score;
        
        if (score > highestScore) {
            highestScore = score;
            bestCategory = category;
        }
    }

    debugLog('Category scores', 'verbose', scores);

    // Apply category only if it meets the threshold
    if (bestCategory && highestScore >= scoreThreshold) {
        try {
            await email.categories.addAsync([bestCategory]);
            debugLog(`Category added: ${bestCategory} (score: ${highestScore})`, 'success');
            showStatus(`Added category ${bestCategory} to email: ${email.subject}`);
            return true;
        } catch (error) {
            debugLog(`Error adding category: ${error.message}`, 'error', error);
            showStatus(`Error categorizing email: ${error.message}`);
            return false;
        }
    } else {
        debugLog(`No category applied. Highest score (${highestScore}) below threshold (${scoreThreshold})`, 'info');
        return false;
    }
}

async function processFolder() {
    const folderId = document.getElementById('folderSelect').value;
    const folderName = document.getElementById('folderSelect').selectedOptions[0].text;
    showStatus('Processing folder...');
    debugLog(`Starting to process folder: ${folderName}`, 'info');
    
    try {
        const folder = Office.context.mailbox.folders.getItem(folderId);
        const items = await folder.getItems();
        
        debugLog(`Found ${items.value.length} emails to process`, 'info');
        
        let processedCount = 0;
        let categorizedCount = 0;
        
        for (const item of items.value) {
            const wasCategorized = await processEmail(item);
            processedCount++;
            if (wasCategorized) categorizedCount++;
            
            debugLog(`Progress: ${processedCount}/${items.value.length} emails processed`, 'verbose');
        }
        
        const summary = `Completed: ${processedCount} emails processed, ${categorizedCount} categorized`;
        debugLog(summary, 'success');
        showStatus(summary);
    } catch (error) {
        debugLog(`Error processing folder: ${error.message}`, 'error', error);
        showStatus('Error processing folder: ' + error.message);
    }
}

// Folder Loading
async function loadFolders() {
    try {
        debugLog('Loading folders...', 'info');
        const folders = await Office.context.mailbox.folders.getAsync();
        const select = document.getElementById('folderSelect');
        
        folders.value.forEach(folder => {
            const option = document.createElement('option');
            option.value = folder.id;
            option.text = folder.displayName;
            select.appendChild(option);
            debugLog(`Folder loaded: ${folder.displayName}`, 'verbose');
        });
        
        debugLog(`Loaded ${folders.value.length} folders`, 'success');
    } catch (error) {
        debugLog('Error loading folders', 'error', error);
        showStatus('Error loading folders: ' + error.message);
    }
}

// Debug Functions
function debugLog(message, type = 'info', data = null) {
    const timestamp = new Date().toLocaleTimeString();
    const debugLogs = document.getElementById('debugLogs');
    
    if (!debugLogs || (!isDebugMode && type !== 'error')) return;
    
    if (!verboseLogging && type === 'verbose') return;
    
    const logMessage = `[${timestamp}] ${message}`;
    const logElement = document.createElement('div');
    logElement.className = type;
    logElement.textContent = logMessage;
    
    if (data && verboseLogging) {
        const dataElement = document.createElement('pre');
        dataElement.textContent = JSON.stringify(data, null, 2);
        logElement.appendChild(dataElement);
    }
    
    debugLogs.insertBefore(logElement, debugLogs.firstChild);
}

function toggleDebug() {
    isDebugMode = !isDebugMode;
    const debugPanel = document.getElementById('debugPanel');
    const toggleButton = document.getElementById('debugToggle');
    debugPanel.style.display = isDebugMode ? 'block' : 'none';
    toggleButton.textContent = isDebugMode ? 'Hide Debug Panel' : 'Show Debug Panel';
    debugLog('Debug mode ' + (isDebugMode ? 'enabled' : 'disabled'), 'info');
}

function clearDebugLogs() {
    document.getElementById('debugLogs').innerHTML = '';
    debugLog('Logs cleared', 'info');
}

// Event listener for verbose logging checkbox
document.getElementById('verboseLogging').addEventListener('change', (e) => {
    verboseLogging = e.checked;
    debugLog(`Verbose logging ${verboseLogging ? 'enabled' : 'disabled'}`, 'info');
});

function showStatus(message) {
    const statusElement = document.getElementById('status');
    statusElement.textContent = message;
    statusElement.className = message.toLowerCase().includes('error') ? 'error' : 'info';
}
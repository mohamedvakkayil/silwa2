// Global variables
let treeData = null;
let root = null;
let svg = null;
let g = null;
let zoom = null;
let nodeMap = new Map();

// Performance optimization: Production mode flag
const PRODUCTION_MODE = true; // Set to false for debugging

// Optimized console logging (disabled in production)
const logger = {
    log: PRODUCTION_MODE ? () => {} : console.log.bind(console),
    warn: PRODUCTION_MODE ? () => {} : console.warn.bind(console),
    error: console.error.bind(console), // Always show errors
    info: PRODUCTION_MODE ? () => {} : console.info.bind(console)
};

// Cache for resolved formulas
const formulaCache = new Map();

// Cache for loaded sheets
const sheetCache = new Map();

// Debounce function for expensive operations
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// Throttle function for frequent operations
function throttle(func, limit) {
    let inThrottle;
    return function(...args) {
        if (!inThrottle) {
            func.apply(this, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
}

// Password Protection
const SITE_PASSWORD = 'Technoplus2024'; // Change this to your desired password

// Check authentication status
function checkAuthentication() {
    const isAuthenticated = sessionStorage.getItem('authenticated') === 'true';
    return isAuthenticated;
}

// Show login modal
function showLoginModal() {
    const loginModal = document.getElementById('loginModal');
    const mainContent = document.getElementById('mainSiteContent');
    
    if (loginModal) {
        loginModal.classList.remove('hidden');
    }
    if (mainContent) {
        mainContent.classList.add('hidden');
    }
}

// Hide login modal and show main content
function hideLoginModal() {
    const loginModal = document.getElementById('loginModal');
    const mainContent = document.getElementById('mainSiteContent');
    
    if (loginModal) {
        loginModal.classList.add('hidden');
    }
    if (mainContent) {
        mainContent.classList.remove('hidden');
    }
}

// Handle login form submission
function handleLogin(event) {
    event.preventDefault();
    
    const passwordInput = document.getElementById('passwordInput');
    const errorMessage = document.getElementById('errorMessage');
    const enteredPassword = passwordInput.value.trim();
    
    // Clear previous error
    if (errorMessage) {
        errorMessage.classList.remove('show');
        errorMessage.textContent = '';
    }
    
    // Check password
    if (enteredPassword === SITE_PASSWORD) {
        // Set authentication
        sessionStorage.setItem('authenticated', 'true');
        
        // Hide login modal and show main content
        hideLoginModal();
        
            // Load data after authentication
        loadData();
        // Start automatic syncing
        startAutoSync();
    } else {
        // Show error message
        if (errorMessage) {
            errorMessage.textContent = 'Incorrect password. Please try again.';
            errorMessage.classList.add('show');
        }
        
        // Clear password input
        passwordInput.value = '';
        passwordInput.focus();
        
        // Shake animation
        const loginBox = document.querySelector('.login-box');
        if (loginBox) {
            loginBox.style.animation = 'shake 0.5s';
            setTimeout(() => {
                loginBox.style.animation = '';
            }, 500);
        }
    }
}

// Initialize authentication on page load
document.addEventListener('DOMContentLoaded', function() {
    // Check if user is already authenticated
    if (checkAuthentication()) {
        hideLoginModal();
        // Load data after a short delay to ensure DOM is ready
        setTimeout(() => {
            loadData();
            // Start automatic syncing
            startAutoSync();
        }, 100);
    } else {
        showLoginModal();
    }
    
    // Set up login form handler
    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.addEventListener('submit', handleLogin);
    }
    
    // Allow Enter key to submit
    const passwordInput = document.getElementById('passwordInput');
    if (passwordInput) {
        passwordInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                handleLogin(e);
            }
        });
    }
});

// Excel file configuration
const EXCEL_FILE_NAME = 'e2.xlsx';

// Helper function to normalize MDB values (MDB -> MDB1, MDB.GF.04 -> MDB4, except for KIND field)
function normalizeMDB(mdbValue) {
    if (!mdbValue || typeof mdbValue !== 'string') {
        return mdbValue || '';
    }
    const trimmed = mdbValue.trim().toUpperCase();
    // Normalize "MDB" to "MDB1" (case-insensitive)
    if (trimmed === 'MDB') {
        return 'MDB1';
    }
    // Normalize "MDB.GF.04" to "MDB4"
    if (trimmed === 'MDB.GF.04' || trimmed === 'MDB GF 04') {
        return 'MDB4';
    }
    return mdbValue;
}

// Helper function to parse load value and extract numeric kW (exclude KVAR)
function parseLoadValue(loadStr) {
    if (!loadStr || typeof loadStr !== 'string') {
        return 0;
    }
    
    // Check if it contains KVAR - exclude it
    const upperStr = loadStr.toUpperCase();
    if (upperStr.includes('KVAR')) {
        return 0; // KVAR is not load, exclude it
    }
    
    // Extract numeric value (handles formats like "1893.42 kW", "1893.42kW", etc.)
    const match = loadStr.match(/(\d+\.?\d*)/);
    if (match) {
        return parseFloat(match[1]) || 0;
    }
    
    return 0;
}

// Extract NO OF UNITS from a detailed sheet (row immediately after SUM AFTER LABOUR)
async function extractNoOfUnitsFromSheet(sheetName) {
    try {
        const response = await fetch(EXCEL_FILE_NAME);
        if (!response.ok) {
            logger.warn(`Could not load Excel file to extract NO OF UNITS for ${sheetName}`);
            return null;
        }
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        // Normalize sheet name (MDB -> MDB1, MDB.GF.04 -> MDB4)
        const normalizedSheetName = normalizeMDB(sheetName);
        let worksheet = workbook.Sheets[normalizedSheetName] || workbook.Sheets[sheetName];
        
        if (!worksheet) {
            logger.warn(`Sheet not found: ${sheetName} or ${normalizedSheetName}`);
            return null;
        }
        
        // Convert sheet to JSON
        const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        
        if (!sheetData || sheetData.length === 0) {
            return null;
        }
        
        // Find ITEM and AMOUNT columns
        let itemColumnIndex = -1;
        let amountColumnIndex = -1;
        
        // Search headers in first few rows
        for (let rowIdx = 0; rowIdx < Math.min(5, sheetData.length); rowIdx++) {
            const row = sheetData[rowIdx];
            if (!row || row.length === 0) continue;
            
            for (let colIdx = 0; colIdx < row.length; colIdx++) {
                const cellValue = (row[colIdx] || '').toString().toLowerCase().trim();
                if (cellValue === 'item' || cellValue === 'items') {
                    itemColumnIndex = colIdx;
                }
                if (cellValue === 'amount') {
                    amountColumnIndex = colIdx;
                }
            }
            
            if (itemColumnIndex !== -1 && amountColumnIndex !== -1) break;
        }
        
        if (itemColumnIndex === -1 || amountColumnIndex === -1) {
            logger.warn(`Could not find ITEM or AMOUNT columns in sheet ${sheetName}`);
            return null;
        }
        
        // Find SUM AFTER LABOUR row
        let sumAfterLabourIndex = -1;
        for (let rowIdx = 0; rowIdx < sheetData.length; rowIdx++) {
            const row = sheetData[rowIdx];
            if (!row || row.length === 0) continue;
            
            const itemValue = row[itemColumnIndex];
            if (itemValue) {
                const itemStr = itemValue.toString().toLowerCase().trim();
                if (itemStr.includes('sum after labour') || itemStr.includes('sum after labor')) {
                    sumAfterLabourIndex = rowIdx;
                    break;
                }
            }
        }
        
        if (sumAfterLabourIndex === -1) {
            logger.warn(`Could not find SUM AFTER LABOUR in sheet ${sheetName}`);
            return null;
        }
        
        // Get NO OF UNITS from row immediately after SUM AFTER LABOUR
        const noOfUnitsRow = sheetData[sumAfterLabourIndex + 1];
        if (!noOfUnitsRow || noOfUnitsRow.length === 0) {
            return null;
        }
        
        const amountValue = noOfUnitsRow[amountColumnIndex];
        if (amountValue !== null && amountValue !== undefined && amountValue !== '') {
            const parsed = parseFloat(amountValue);
            if (!isNaN(parsed) && parsed > 0) {
                logger.log(`✓ Extracted NO OF UNITS from ${sheetName}: ${parsed}`);
                return parsed;
            }
        }
        
        return null;
    } catch (error) {
        console.error(`Error extracting NO OF UNITS from sheet ${sheetName}:`, error);
        return null;
    }
}

// Calculate sum of immediate children loads for a board (MDB, SMDB, ESMDB, MCC, etc.) excluding KVAR
async function calculateChildrenLoadSum(boardName, boardNode = null) {
    // If boardNode is provided, use it directly (for recursive calls)
    let targetNode = boardNode;
    
    if (!targetNode) {
        if (!treeData || !treeData.children) {
            logger.warn(`Tree data not available for load calculation: ${boardName}`);
            return 0;
        }
        
        // Search recursively through the tree to find the board node
        function findNode(node, name) {
            if (!node) return null;
            
            const nodeName = node.name || node.id || '';
            const normalizedNodeName = normalizeMDB(nodeName);
            const normalizedName = normalizeMDB(name);
            
            if (nodeName === name || normalizedNodeName === name ||
                nodeName === normalizedName || normalizedNodeName === normalizedName) {
                return node;
            }
            
            // Search in children
            const children = node.children || node._children || [];
            for (const child of children) {
                const found = findNode(child, name);
                if (found) return found;
            }
            
            return null;
        }
        
        // First try to find in MDB level
        targetNode = treeData.children.find(child => {
            const childName = child.name || child.id || '';
            const normalizedChildName = normalizeMDB(childName);
            return childName === boardName || normalizedChildName === boardName ||
                   childName === normalizeMDB(boardName) || normalizedChildName === normalizeMDB(boardName);
        });
        
        // If not found at MDB level, search recursively
        if (!targetNode) {
            for (const mdb of treeData.children) {
                targetNode = findNode(mdb, boardName);
                if (targetNode) break;
            }
        }
    }
    
    if (!targetNode) {
        logger.warn(`Board node not found for load calculation: ${boardName}`);
        return 0;
    }
    
    // Get children (check both children and _children for collapsed nodes)
    const children = targetNode.children || targetNode._children || [];
    
    if (children.length === 0) {
        return 0;
    }
    
    let sum = 0;
    
    // Process children - for DBs, load detailed sheet to get NO OF UNITS
    for (const child of children) {
        // Get load from child node - check multiple possible locations
        let childLoad = '0 kW';
        
        // Try different ways to get load value
        if (child.load) {
            childLoad = child.load;
        } else if (child.data) {
            if (child.data.Load) {
                childLoad = child.data.Load;
            } else if (child.data.LOAD) {
                childLoad = child.data.LOAD;
            } else if (child.data.load) {
                childLoad = child.data.load;
            } else if (child.data.data) {
                childLoad = child.data.data.Load || child.data.data.LOAD || child.data.data.load || '0 kW';
            }
        }
        
        // Get number of units/items
        let noOfUnits = 1; // Default to 1 if not found
        
        // Check child kind - if it's a DB or ESMDB, handle special cases
        const childKind = (child.kind || child.data?.kind || child.data?.KIND || '').toString().toUpperCase();
        const childName = child.name || child.id || '';
        
        // Check if this is a special DB (grouped with copy count)
        const specialDBs = ['DB.TN.LXX.1B1.01', 'DB.TN.LXX.2B1.01', 'DB.TN.LXX.3B1.01', 'DB.TH.GF.01'];
        const isSpecialDB = specialDBs.some(db => childName.includes(db));
        
        // Check if this is a special ESMDB (appears under multiple parents)
        const specialESMDBs = ['ESMDB.LL.RF.01(LIFT)', 'ESMDB.LL.RF.02(LIFT)'];
        const isSpecialESMDB = specialESMDBs.some(esmdb => childName.includes(esmdb));
        
        const isGroupNode = childName.includes('copies') || child.data?.isGroup || child.data?.copyCount;
        
        if (childKind === 'DB') {
            // Special handling for grouped special DBs
            if (isGroupNode || isSpecialDB) {
                // Try to get copy count from group node data
                if (child.data?.copyCount !== undefined && child.data?.copyCount !== null) {
                    noOfUnits = parseFloat(child.data.copyCount) || 1;
                } else {
                    // Extract copy count from name like "DB.TN.LXX.1B1.01 (4 copies)"
                    const match = childName.match(/\((\d+)\s*copies?\)/i);
                    if (match) {
                        noOfUnits = parseInt(match[1]) || 1;
                    } else {
                        // For special DBs, try to get copy count from parent SMDB/MDB
                        // Find parent name from fedFromSMDB or parent node
                        const parentName = child.data?.fedFromSMDB || 
                                         (targetNode && targetNode.name ? targetNode.name : null);
                        if (parentName && isSpecialDB) {
                            const copyCounts = getDBCopyCounts(parentName);
                            if (copyCounts) {
                                // Extract base DB name (remove " (X copies)" if present)
                                const baseDBName = childName.replace(/\s*\(\d+\s*copies?\)/i, '').trim();
                                if (copyCounts[baseDBName]) {
                                    noOfUnits = copyCounts[baseDBName];
                                }
                            }
                        }
                    }
                }
            } else {
                // For regular DBs (not grouped), extract NO OF UNITS from detailed sheet
                const extractedUnits = await extractNoOfUnitsFromSheet(childName);
                if (extractedUnits !== null && extractedUnits > 0) {
                    noOfUnits = extractedUnits;
                } else {
                    // Fallback to data properties if sheet extraction fails
                    if (child.data) {
                        if (child.data['NO OF ITEMS'] !== undefined && child.data['NO OF ITEMS'] !== null && child.data['NO OF ITEMS'] !== '') {
                            noOfUnits = parseFloat(child.data['NO OF ITEMS']) || 1;
                        } else if (child.data['NO OF UNITS'] !== undefined && child.data['NO OF UNITS'] !== null && child.data['NO OF UNITS'] !== '') {
                            noOfUnits = parseFloat(child.data['NO OF UNITS']) || 1;
                        } else if (child.data.data) {
                            if (child.data.data['NO OF ITEMS'] !== undefined && child.data.data['NO OF ITEMS'] !== null && child.data.data['NO OF ITEMS'] !== '') {
                                noOfUnits = parseFloat(child.data.data['NO OF ITEMS']) || 1;
                            } else if (child.data.data['NO OF UNITS'] !== undefined && child.data.data['NO OF UNITS'] !== null && child.data.data['NO OF UNITS'] !== '') {
                                noOfUnits = parseFloat(child.data.data['NO OF UNITS']) || 1;
                            }
                        }
                    }
                }
            }
        } else if (childKind === 'ESMDB' && (isGroupNode || isSpecialESMDB)) {
            // Special handling for grouped special ESMDBs
            // Try to get copy count from group node data
            if (child.data?.copyCount !== undefined && child.data?.copyCount !== null) {
                noOfUnits = parseFloat(child.data.copyCount) || 1;
            } else {
                // Extract copy count from name like "ESMDB.LL.RF.01(LIFT) (1 copies)"
                const match = childName.match(/\((\d+)\s*copies?\)/i);
                if (match) {
                    noOfUnits = parseInt(match[1]) || 1;
                } else {
                    // For special ESMDBs, try to get copy count from parent
                    const parentName = child.data?.fedFromParent || 
                                     (targetNode && targetNode.name ? targetNode.name : null);
                    if (parentName && isSpecialESMDB) {
                        const copyCounts = getESMDBCopyCounts(parentName);
                        if (copyCounts) {
                            // Extract base ESMDB name (remove " (X copies)" if present)
                            const baseESMDBName = childName.replace(/\s*\(\d+\s*copies?\)/i, '').trim();
                            if (copyCounts[baseESMDBName]) {
                                noOfUnits = copyCounts[baseESMDBName];
                            }
                        }
                    }
                }
            }
        } else {
            // For non-DB/non-special-ESMDB children, use data properties
            if (child.data) {
                if (child.data['NO OF ITEMS'] !== undefined && child.data['NO OF ITEMS'] !== null && child.data['NO OF ITEMS'] !== '') {
                    noOfUnits = parseFloat(child.data['NO OF ITEMS']) || 1;
                } else if (child.data['NO OF UNITS'] !== undefined && child.data['NO OF UNITS'] !== null && child.data['NO OF UNITS'] !== '') {
                    noOfUnits = parseFloat(child.data['NO OF UNITS']) || 1;
                } else if (child.data.data) {
                    if (child.data.data['NO OF ITEMS'] !== undefined && child.data.data['NO OF ITEMS'] !== null && child.data.data['NO OF ITEMS'] !== '') {
                        noOfUnits = parseFloat(child.data.data['NO OF ITEMS']) || 1;
                    } else if (child.data.data['NO OF UNITS'] !== undefined && child.data.data['NO OF UNITS'] !== null && child.data.data['NO OF UNITS'] !== '') {
                        noOfUnits = parseFloat(child.data.data['NO OF UNITS']) || 1;
                    }
                }
            }
        }
        
        const loadValue = parseLoadValue(childLoad);
        // Multiply load by number of units
        const effectiveLoad = loadValue * noOfUnits;
        sum += effectiveLoad;
        
        // Debug log
        logger.log(`  Child ${childName} (${childKind}): Load=${loadValue.toFixed(2)} kW × Units=${noOfUnits} = ${effectiveLoad.toFixed(2)} kW`);
    }
    
    logger.log(`✓ Calculated children load sum for ${boardName}: ${sum.toFixed(2)} kW (load × units)`);
    return sum;
}

// Store last modified time for change detection
let lastModifiedTime = null;
let pollingInterval = null;
const POLL_INTERVAL = 60000; // Check every 60 seconds (reduced frequency for better performance)

// Load data dynamically from Excel file
function loadData(showLoading = true, silent = false) {
    // Show loading indicator only on initial load
    if (showLoading && !silent) {
        const mindmapContainer = document.getElementById('mindmap');
        if (mindmapContainer) {
            mindmapContainer.innerHTML = `<div style="padding: 40px; text-align: center; color: #666;"><div style="font-size: 18px; margin-bottom: 10px;">Loading data from Excel...</div><div style="font-size: 14px; color: #999;">Syncing with ${EXCEL_FILE_NAME}</div></div>`;
        }
    }
    
    // Fetch Excel file (cache: 'no-store' prevents browser caching)
    fetch(EXCEL_FILE_NAME, { cache: 'no-store' })
        .then(response => {
            if (!response.ok) {
                throw new Error(`Failed to fetch Excel file: ${response.status} ${response.statusText}`);
            }
            
            // Check if file has changed using headers
            const currentModified = response.headers.get('Last-Modified');
            const etag = response.headers.get('ETag');
            const checkValue = currentModified || etag;
            
            if (silent) {
                logger.log('Silent check - Last-Modified:', currentModified, 'ETag:', etag, 'Stored:', lastModifiedTime);
            }
            
            // For silent checks, skip if file hasn't changed
            if (silent && lastModifiedTime && checkValue && checkValue === lastModifiedTime) {
                logger.log('✓ Excel file unchanged, skipping reload');
                return null; // Signal no change
            }
            
            // Update last modified time
            if (checkValue) {
                const wasUpdate = lastModifiedTime !== checkValue;
                lastModifiedTime = checkValue;
                if (wasUpdate && silent) {
                    logger.log('✓ File changed detected, reloading data...');
                }
            } else if (silent) {
                logger.warn('⚠ No Last-Modified or ETag header available, will reload anyway');
            }
            
            return response.arrayBuffer();
        })
        .then(data => {
            if (!data) return; // No change detected, skip processing
            processExcelData(data, silent);
        })
        .catch(error => {
            console.error('Error loading Excel file:', error);
            if (!silent) {
                const mindmapContainer = document.getElementById('mindmap');
                if (mindmapContainer) {
                    mindmapContainer.innerHTML = 
                        '<div style="padding: 40px; text-align: center; color: #d32f2f;">' +
                        '<div style="font-size: 18px; margin-bottom: 10px;">Error loading data</div>' +
                        '<div style="font-size: 14px; color: #666; margin-bottom: 20px;">' + error.message + '</div>' +
                        '<button onclick="loadData()" style="padding: 10px 20px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer;">Retry</button>' +
                        '</div>';
                }
            }
        });
}

// Process Excel data
function processExcelData(data, silent = false) {
    if (!data) return;
    
    // Parse Excel file
    const workbook = XLSX.read(data, { type: 'array' });
    
    // Read TOTALLIST sheet
    if (!workbook.SheetNames.includes('TOTALLIST')) {
        throw new Error('TOTALLIST sheet not found in Excel file');
    }
    
    const worksheet = workbook.Sheets['TOTALLIST'];
    
    // Helper function to resolve Excel formulas by reading referenced cells (with caching)
    function resolveFormula(formula, workbook) {
        if (!formula || typeof formula !== 'string' || !formula.startsWith('=')) {
            return formula; // Not a formula, return as-is
        }
        
        // Check cache first
        if (formulaCache.has(formula)) {
            return formulaCache.get(formula);
        }
        
        // Parse formula like ='SheetName'!F163 or =SheetName!F163
        // Handle both quoted and unquoted sheet names
        let sheetName, col, row;
        
        // Try pattern with quotes: ='Sheet Name'!F163
        const quotedMatch = formula.match(/^='([^']+)'!([A-Z]+)(\d+)$/);
        if (quotedMatch) {
            sheetName = quotedMatch[1];
            col = quotedMatch[2];
            row = parseInt(quotedMatch[3]);
        } else {
            // Try pattern without quotes: =SheetName!F163
            const unquotedMatch = formula.match(/^=([^!]+)!([A-Z]+)(\d+)$/);
            if (unquotedMatch) {
                sheetName = unquotedMatch[1];
                col = unquotedMatch[2];
                row = parseInt(unquotedMatch[3]);
            } else {
                logger.warn(`Could not parse formula: ${formula}`);
                formulaCache.set(formula, 0);
                return 0;
            }
        }
        
        try {
            // Find matching sheet (handle normalization: MDB -> MDB1, MDB.GF.04 -> MDB4)
            let targetSheetName = sheetName;
            if (sheetName === 'MDB') {
                targetSheetName = 'MDB1';
            } else if (sheetName === 'MDB.GF.04' || sheetName === 'MDB GF 04') {
                targetSheetName = 'MDB4';
            }
            
            if (workbook.SheetNames.includes(targetSheetName)) {
                const refSheet = workbook.Sheets[targetSheetName];
                // Convert column letter to index (F = 5, A = 0)
                const colIndex = XLSX.utils.decode_col(col);
                // Convert row number to index (163 -> 162, since it's 0-based)
                const rowIndex = row - 1;
                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
                const cell = refSheet[cellAddress];
                
                if (cell) {
                    // Return the value (v property contains the calculated value)
                    const value = cell.v;
                    if (value !== undefined && value !== null && value !== '') {
                        const result = typeof value === 'number' ? value : parseFloat(value) || 0;
                        formulaCache.set(formula, result);
                        return result;
                    }
                }
            } else {
                // Try to find sheet with similar name
                const matchingSheet = workbook.SheetNames.find(name => 
                    name.toUpperCase() === targetSheetName.toUpperCase() ||
                    name.toUpperCase() === sheetName.toUpperCase()
                );
                if (matchingSheet) {
                    const refSheet = workbook.Sheets[matchingSheet];
                    const colIndex = XLSX.utils.decode_col(col);
                    const rowIndex = row - 1;
                    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
                    const cell = refSheet[cellAddress];
                    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
                        const result = typeof cell.v === 'number' ? cell.v : parseFloat(cell.v) || 0;
                        formulaCache.set(formula, result);
                        return result;
                    }
                } else {
                    logger.warn(`Sheet "${targetSheetName}" (from formula ${formula}) not found.`);
                }
            }
        } catch (e) {
            logger.warn(`Error resolving formula ${formula}:`, e);
        }
        
        // Cache the result (0) to avoid repeated failures
        formulaCache.set(formula, 0);
        return 0; // Return 0 if formula can't be resolved
    }
    
    // First, find the Estimate column index and header row
    let estimateColIndex = -1;
    let headerRowIndex = 0;
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Find header row and Estimate column
    for (let row = 0; row <= Math.min(10, range.e.r); row++) {
        for (let col = 0; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            if (cell && cell.v) {
                const cellValue = String(cell.v).toLowerCase().trim();
                if (cellValue === 'estimate') {
                    estimateColIndex = col;
                    headerRowIndex = row;
                    break;
                }
            }
        }
        if (estimateColIndex !== -1) break;
    }
    
    // Convert sheet to JSON array
    // Use raw: true to preserve numeric types (consistent with data.js format)
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        defval: '', // Default value for empty cells
        raw: true  // Preserve numeric types for calculations
    });
    
    // Convert to array of objects (same format as data.js)
    const processedData = jsonData.map((row, rowIndex) => {
        const processedRow = {};
        for (const [key, value] of Object.entries(row)) {
            // Handle NaN and null values
            if (value === null || value === undefined || value === '') {
                processedRow[key] = '';
            } else if (typeof value === 'number' && isNaN(value)) {
                processedRow[key] = '';
            } else {
                // Preserve the value as-is (numbers stay numbers, strings stay strings)
                let processedValue = value;
                
                // Check if this is the Estimate column and if the cell contains a formula
                if (key === 'Estimate' || key === 'ESTIMATE') {
                    // Read the cell directly to check for formulas
                    if (estimateColIndex !== -1) {
                        // Calculate Excel row: headerRowIndex + 1 (for header) + rowIndex (data row)
                        const excelRowIndex = headerRowIndex + 1 + rowIndex;
                        const cellAddress = XLSX.utils.encode_cell({ r: excelRowIndex, c: estimateColIndex });
                        const cell = worksheet[cellAddress];
                        
                        if (cell) {
                            // Check if cell has a formula (f property)
                            if (cell.f) {
                                // Cell contains a formula - resolve it
                                processedValue = resolveFormula(cell.f, workbook);
                            } else if (cell.v !== undefined && cell.v !== null && cell.v !== '') {
                                // Cell has a calculated value (Excel already calculated it)
                                processedValue = typeof cell.v === 'number' ? cell.v : parseFloat(cell.v) || 0;
                            } else if (typeof processedValue === 'string' && processedValue.startsWith('=')) {
                                // Value is a formula string - resolve it
                                processedValue = resolveFormula(processedValue, workbook);
                            }
                        } else {
                            // Cell not found - might be empty or out of range
                            // Keep the value from sheet_to_json
                        }
                    } else if (typeof processedValue === 'string' && processedValue.startsWith('=')) {
                        // Fallback: if value is a formula string, resolve it
                        processedValue = resolveFormula(processedValue, workbook);
                    }
                } else {
                    // Check if this is a formula (starts with =) and resolve it
                    if (typeof processedValue === 'string' && processedValue.startsWith('=')) {
                        // This is likely a formula - try to resolve it
                        processedValue = resolveFormula(processedValue, workbook);
                    }
                }
                
                // Fix SMDB.LL.04.01 -> SMDB.LL.L04.01 (correct missing L in level number)
                if (typeof processedValue === 'string' && (key === 'Itemdrop' || key === 'MDB')) {
                    // Fix pattern: SMDB.LL.##.01 -> SMDB.LL.L##.01
                    processedValue = processedValue.replace(/^SMDB\.LL\.(\d{2})\.01$/i, 'SMDB.LL.L$1.01');
                }
                
                processedRow[key] = processedValue;
            }
        }
        return processedRow;
    });
    
    const wasInitialLoad = !window.allData;
    const dataChanged = wasInitialLoad || JSON.stringify(window.allData) !== JSON.stringify(processedData);
    
    logger.log('Data loaded from Excel:', processedData.length, 'items');
    
    // Verify ESTIMATE and LOAD columns are present
    if (processedData.length > 0) {
        const firstItem = processedData[0];
        const hasEstimateColumn = 'Estimate' in firstItem || 'ESTIMATE' in firstItem;
        const hasLoadColumn = 'Load' in firstItem || 'LOAD' in firstItem;
        logger.log('✓ ESTIMATE column present:', hasEstimateColumn);
        logger.log('✓ LOAD column present:', hasLoadColumn);
        if (hasEstimateColumn) {
            const estimateKey = 'Estimate' in firstItem ? 'Estimate' : 'ESTIMATE';
            logger.log('✓ ESTIMATE column key:', estimateKey);
            // Sample a few items to verify estimates
            const sampleItems = processedData.slice(0, 5).filter(item => item[estimateKey]);
            logger.log('✓ Sample ESTIMATE values:', sampleItems.map(item => ({
                name: item.Itemdrop || item.MDB,
                estimate: item[estimateKey]
            })));
        }
        if (hasLoadColumn) {
            const loadKey = 'Load' in firstItem ? 'Load' : 'LOAD';
            logger.log('✓ LOAD column key:', loadKey);
            // Sample a few items to verify loads
            const sampleItems = processedData.slice(0, 5).filter(item => item[loadKey]);
            logger.log('✓ Sample LOAD values:', sampleItems.map(item => ({
                name: item.Itemdrop || item.MDB,
                load: item[loadKey]
            })));
        }
    }
    
    window.allData = processedData;
    
    // Process and display data (use requestAnimationFrame for smooth updates)
    requestAnimationFrame(() => {
        processData();
        requestAnimationFrame(() => {
            initializeDiagram();
            updateStats();
            populateFilters();
        });
    });
    
    // Show notification if data was updated (not initial load)
    if (!wasInitialLoad && dataChanged && !silent) {
        showSyncNotification('Data synchronized with Excel file');
    }
}

// Show sync notification
function showSyncNotification(message) {
    // Remove existing notification if any
    const existingNotification = document.getElementById('sync-notification');
    if (existingNotification) {
        existingNotification.remove();
    }
    
    // Create notification element
    const notification = document.createElement('div');
    notification.id = 'sync-notification';
    notification.textContent = message;
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #4CAF50;
        color: white;
        padding: 12px 20px;
        border-radius: 5px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        z-index: 10000;
        font-size: 14px;
        animation: slideIn 0.3s ease-out;
    `;
    
    document.body.appendChild(notification);
    
    // Remove after 3 seconds
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease-out';
        setTimeout(() => {
            if (notification.parentNode) {
                notification.remove();
            }
        }, 300);
    }, 3000);
}

// Start automatic polling for Excel file changes
function startAutoSync() {
    // Clear any existing interval
    if (pollingInterval) {
        clearInterval(pollingInterval);
        logger.log('Cleared existing auto-sync interval');
    }
    
    // Poll every POLL_INTERVAL milliseconds
    pollingInterval = setInterval(() => {
        logger.log(`Auto-sync: Checking ${EXCEL_FILE_NAME} for changes...`);
        loadData(false, true); // Silent check (no loading indicator)
    }, POLL_INTERVAL);
    
    logger.log(`✓ Auto-sync started: Monitoring ${EXCEL_FILE_NAME} for changes every ${POLL_INTERVAL / 1000} seconds`);
    logger.log(`✓ Next check will be in ${POLL_INTERVAL / 1000} seconds`);
    
    // Update sync status indicator
    updateSyncStatus(true);
}

// Stop automatic polling
function stopAutoSync() {
    if (pollingInterval) {
        clearInterval(pollingInterval);
        pollingInterval = null;
        logger.log('Auto-sync stopped');
        updateSyncStatus(false);
    }
}

// Update sync status indicator
function updateSyncStatus(isActive) {
    const syncIndicator = document.getElementById('sync-indicator');
    const syncText = document.getElementById('sync-text');
    
    if (syncIndicator && syncText) {
        if (isActive) {
            syncIndicator.textContent = '●';
            syncIndicator.style.color = '#4CAF50';
            syncIndicator.style.animation = 'pulse 2s infinite';
            syncText.textContent = `Syncing with ${EXCEL_FILE_NAME} (every ${POLL_INTERVAL / 1000}s)`;
            syncText.style.color = '#4CAF50';
        } else {
            syncIndicator.textContent = '●';
            syncIndicator.style.color = '#999';
            syncIndicator.style.animation = 'none';
            syncText.textContent = 'Sync stopped';
            syncText.style.color = '#999';
        }
    }
}

// Helper function to get DB copy counts for special SMDBs/MDBs
function getDBCopyCounts(parentName) {
    const normalized = parentName.replace(/\.01$/, '');
    
    // Special case for SMDB.TN.P2
    if (normalized === 'SMDB.TN.P2') {
        return {
            'DB.TN.LXX.1B1.01': 3,
            'DB.TN.LXX.2B1.01': 1,
            'DB.TN.LXX.3B1.01': 1
        };
    }
    
    // Special case for DB.TH.GF.01 - appears under SMDB.TH.B1.01 and MDB4
    if (normalized === 'SMDB.TH.B1' || parentName === 'SMDB.TH.B1.01') {
        return {
            'DB.TH.GF.01': 1
        };
    }
    
    // Handle MDB4 (MDB.GF.04) as parent for DB.TH.GF.01
    if (normalized === 'MDB.GF.04' || normalized === 'MDB4' || parentName === 'MDB4' || parentName === 'MDB.GF.04') {
        return {
            'DB.TH.GF.01': 1
        };
    }
    
    // Handle SMDB.TN.L## pattern
    const match = normalized.match(/^SMDB\.TN\.L(\d+)$/);
    
    if (!match) return null;
    
    const level = parseInt(match[1]);
    
    if (level === 1) {
        return {
            'DB.TN.LXX.1B1.01': 5,
            'DB.TN.LXX.2B1.01': 3,
            'DB.TN.LXX.3B1.01': 1
        };
    }
    
    if (level === 23) {
        return {
            'DB.TN.LXX.2B1.01': 2,
            'DB.TN.LXX.1B1.01': 1,
            'DB.TN.LXX.3B1.01': 1
        };
    }
    
    if ((level >= 2 && level <= 22) || (level >= 24 && level <= 47)) {
        return {
            'DB.TN.LXX.2B1.01': 4,
            'DB.TN.LXX.1B1.01': 4,
            'DB.TN.LXX.3B1.01': 1
        };
    }
    
    return null;
}

// Helper function to get ESMDB copy counts for special ESMDBs
// ESMDB.LL.RF.01(LIFT) and ESMDB.LL.RF.02(LIFT) appear under both BB.05 and EMDB.GF.01
function getESMDBCopyCounts(parentName) {
    const normalized = parentName.replace(/\.01$/, '');
    
    // Special ESMDBs that appear under multiple parents
    const specialESMDBs = ['ESMDB.LL.RF.01(LIFT)', 'ESMDB.LL.RF.02(LIFT)'];
    
    // Check if parent is BB.05 or EMDB.GF.01
    if (normalized === 'BB.05' || parentName === 'BB.05' || 
        normalized === 'EMDB.GF' || parentName === 'EMDB.GF.01') {
        return {
            'ESMDB.LL.RF.01(LIFT)': 1,
            'ESMDB.LL.RF.02(LIFT)': 1
        };
    }
    
    return null;
}

function processData() {
    const allData = window.allData || (typeof allData !== 'undefined' ? allData : []);
    if (!allData || allData.length === 0) {
        console.error('No data available');
        return;
    }
    nodeMap.clear();
    
    // Create RMU node as root
    const rmuNode = {
        name: 'RMU',
        id: 'RMU',
        kind: 'RMU',
        children: []
    };
    
    // Create MDB nodes
    const mdbList = ['MDB1', 'MDB2', 'MDB3', 'MDB4'];
    const mdbNodes = {};
    
    mdbList.forEach(mdbName => {
        // Find corresponding data item for this MDB
        // For MDB4, also check for MDB.GF.04
        let mdbData = allData.find(d => {
            const itemdrop = d.Itemdrop || '';
            const mdb = d.MDB || '';
            return itemdrop === mdbName || mdb === mdbName || 
                   (mdbName === 'MDB4' && (itemdrop === 'MDB.GF.04' || mdb === 'MDB.GF.04'));
        });
        
        // If not found, try to find by sheet name pattern (MDB1 sheet for MDB1, etc.)
        if (!mdbData && mdbName === 'MDB1') {
            // Also check for sheets named "MDB" (without number)
            mdbData = allData.find(d => {
                const itemdrop = (d.Itemdrop || '').toString().trim();
                const mdb = (d.MDB || '').toString().trim();
                return (itemdrop === 'MDB' || mdb === 'MDB') && 
                       (d.KIND || '').toString().toUpperCase() === 'MDB';
            });
        }
        
        // Ensure mdbData has all required fields, even if empty
        // Get load from LOAD column (try both cases for column name)
        const mdbLoad = mdbData ? ((mdbData.Load !== undefined && mdbData.Load !== null && mdbData.Load !== '') 
            ? mdbData.Load 
            : ((mdbData.LOAD !== undefined && mdbData.LOAD !== null && mdbData.LOAD !== '') ? mdbData.LOAD : '0 kW')) : '0 kW';
        
        // Normalize Itemdrop and MDB to remove "MDB" from MDB1 and "MDB.GF.04" from MDB4
        const normalizedItemdrop = mdbData ? (normalizeMDB(mdbData.Itemdrop) || mdbName) : mdbName;
        const normalizedMDB = mdbData ? (normalizeMDB(mdbData.MDB) || mdbName) : mdbName;
        
        const completeMdbData = mdbData ? {
            ...mdbData, // Spread to include any other fields first
            // Override Itemdrop and MDB with normalized values (removes "MDB" from MDB1, "MDB.GF.04" from MDB4)
            Itemdrop: normalizedItemdrop,
            MDB: normalizedMDB,
            KIND: mdbData.KIND || 'MDB',
            Load: mdbLoad,
            Estimate: mdbData.Estimate || 0,
            'NO OF ITEMS': mdbData['NO OF ITEMS'] || 0,
            'FED FROM': mdbData['FED FROM'] || 'RMU'
        } : {
            Itemdrop: mdbName,
            MDB: mdbName,
            KIND: 'MDB',
            Load: '0 kW',
            Estimate: 0,
            'NO OF ITEMS': 0,
            'FED FROM': 'RMU'
        };
        
        const mdbNode = {
            name: mdbName,
            id: mdbName,
            kind: 'MDB',
            load: mdbLoad,
            estimate: parseFloat(completeMdbData.Estimate) || 0,
            mdb: mdbName,
            data: completeMdbData, // Use complete data with all fields
            children: []
        };
        rmuNode.children.push(mdbNode);
        mdbNodes[mdbName] = mdbNode;
    });
    
    // Process all items - build complete tree structure
    const specialDBs = ['DB.TN.LXX.1B1.01', 'DB.TN.LXX.2B1.01', 'DB.TN.LXX.3B1.01', 'DB.TH.GF.01'];
    const itemMap = new Map();
    const processedItems = new Set();
    
    // First pass: create all node structures
    const allDataArray = window.allData || [];
    allDataArray.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        
        // Skip if itemName matches an MDB node (these are reference entries, not actual tree nodes)
        if (mdbList.includes(itemName)) return;
        
        // Skip if no itemName or if it's a special/reserved name
        if (!itemName || itemName === 'RMU' || specialDBs.includes(itemName)) return;
        
        const kind = item.KIND || 'Unknown';
        // Get load from LOAD column (try both cases for column name)
        const load = (item.Load !== undefined && item.Load !== null && item.Load !== '') 
            ? item.Load 
            : ((item.LOAD !== undefined && item.LOAD !== null && item.LOAD !== '') ? item.LOAD : '0 kW');
        // Get estimate from ESTIMATE column (try both cases for column name)
        const estimate = (item.Estimate !== undefined && item.Estimate !== null && item.Estimate !== '') 
            ? item.Estimate 
            : ((item.ESTIMATE !== undefined && item.ESTIMATE !== null && item.ESTIMATE !== '') ? item.ESTIMATE : 0);
        const mdb = item.MDB || '';
        
        const nodeData = {
            name: itemName,
            id: itemName,
            kind: kind,
            load: load,
            estimate: estimate,
            mdb: normalizeMDB(mdb),
            data: item,
            children: []
        };
        
        itemMap.set(itemName, nodeData);
    });
    
    // Second pass: establish relationships
    allDataArray.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        
        // Skip if itemName matches an MDB node (these are reference entries, not actual tree nodes)
        if (mdbList.includes(itemName)) return;
        
        // Skip if no itemName or if it's a special/reserved name
        if (!itemName || itemName === 'RMU' || specialDBs.includes(itemName)) return;
        if (processedItems.has(itemName)) return;
        
        const fedFrom = item['FED FROM'] || '';
        const mdb = item.MDB || '';
        const nodeData = itemMap.get(itemName);
        
        if (!nodeData) return;
        
        // Handle FED FROM relationships
        if (fedFrom && fedFrom.trim() !== '') {
            const parents = fedFrom.split('\n').map(p => p.trim()).filter(p => p && p !== 'RMU');
            
            if (parents.length > 0) {
                // Use first parent for tree structure
                const parent = parents[0];
                
                if (parent.startsWith('MDB') && mdbNodes[parent]) {
                    // Direct connection to MDB
                    if (!mdbNodes[parent].children.find(c => c.id === itemName)) {
                        mdbNodes[parent].children.push(nodeData);
                        processedItems.add(itemName);
                    }
                } else {
                    // Connect to parent item
                    let parentNode = itemMap.get(parent);
                    if (!parentNode) {
                        // Create placeholder parent
                        parentNode = {
                            name: parent,
                            id: parent,
                            kind: parent.startsWith('BB') ? 'BUS BAR RAISER' : 
                                  parent.startsWith('SMDB') ? 'SMDB' : 
                                  parent.startsWith('DB') ? 'DB' : 'Parent',
                            children: []
                        };
                        itemMap.set(parent, parentNode);
                        
                        // Try to connect parent to its MDB
                        const parentItem = allDataArray.find(d => (d.Itemdrop || d.MDB) === parent);
                        if (parentItem) {
                            const parentMDB = parentItem.MDB || '';
                            const parentFedFrom = parentItem['FED FROM'] || '';
                            if (parentMDB && mdbNodes[parentMDB]) {
                                if (!mdbNodes[parentMDB].children.find(c => c.id === parent)) {
                                    mdbNodes[parentMDB].children.push(parentNode);
                                }
                            } else if (parentFedFrom.includes('MDB')) {
                                const mdbMatch = parentFedFrom.match(/MDB\d/);
                                if (mdbMatch && mdbNodes[mdbMatch[0]]) {
                                    if (!mdbNodes[mdbMatch[0]].children.find(c => c.id === parent)) {
                                        mdbNodes[mdbMatch[0]].children.push(parentNode);
                                    }
                                }
                            }
                        }
                    }
                    
                    if (!parentNode.children.find(c => c.id === itemName)) {
                        parentNode.children.push(nodeData);
                        processedItems.add(itemName);
                    }
                }
            } else if (mdb && mdbNodes[mdb]) {
                // No parent specified, connect to MDB
                if (!mdbNodes[mdb].children.find(c => c.id === itemName)) {
                    mdbNodes[mdb].children.push(nodeData);
                    processedItems.add(itemName);
                }
            }
        } else if (mdb && mdbNodes[mdb]) {
            // No FED FROM, connect to MDB
            if (!mdbNodes[mdb].children.find(c => c.id === itemName)) {
                mdbNodes[mdb].children.push(nodeData);
                processedItems.add(itemName);
            }
        }
    });
    
    // Handle special DBs with grouping
    // First, process special DBs from their FED FROM fields
    allDataArray.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        if (!specialDBs.includes(itemName)) return;
        
        const fedFrom = item['FED FROM'] || '';
        const kind = item.KIND || 'Unknown';
        // Get load from LOAD column (try both cases for column name)
        const load = (item.Load !== undefined && item.Load !== null && item.Load !== '') 
            ? item.Load 
            : ((item.LOAD !== undefined && item.LOAD !== null && item.LOAD !== '') ? item.LOAD : '0 kW');
        // Get estimate from ESTIMATE column (try both cases for column name)
        const estimate = (item.Estimate !== undefined && item.Estimate !== null && item.Estimate !== '') 
            ? item.Estimate 
            : ((item.ESTIMATE !== undefined && item.ESTIMATE !== null && item.ESTIMATE !== '') ? item.ESTIMATE : 0);
        const mdb = item.MDB || '';
        
        const parentList = fedFrom.split('\n').map(p => p.trim()).filter(p => p && p !== 'RMU');
        
        parentList.forEach(parentName => {
            const copyCounts = getDBCopyCounts(parentName);
            if (copyCounts && copyCounts[itemName]) {
                const count = copyCounts[itemName];
                
                // Check if parent is MDB4 or an SMDB
                const isMDB4 = parentName === 'MDB4' || parentName === 'MDB.GF.04' || parentName === 'MDB GF 04';
                
                // Find or create parent node (SMDB or MDB4)
                let parentNode = itemMap.get(parentName);
                if (!parentNode) {
                    if (isMDB4) {
                        // For MDB4, use the existing MDB node
                        const mdb4Name = 'MDB4';
                        parentNode = mdbNodes[mdb4Name];
                        if (!parentNode) {
                            // Create MDB4 node if it doesn't exist
                            parentNode = {
                                name: mdb4Name,
                                id: mdb4Name,
                                kind: 'MDB',
                                children: []
                            };
                            mdbNodes[mdb4Name] = parentNode;
                            if (rmuNode && !rmuNode.children.find(c => c.id === mdb4Name)) {
                                rmuNode.children.push(parentNode);
                            }
                        }
                    } else {
                        // For SMDB, create new node
                        parentNode = {
                            name: parentName,
                            id: parentName,
                            kind: 'SMDB',
                            children: []
                        };
                        itemMap.set(parentName, parentNode);
                        
                        // Connect SMDB to its parent (find from data)
                        const smdbItem = allDataArray.find(d => (d.Itemdrop || d.MDB) === parentName);
                        if (smdbItem) {
                            const smdbFedFrom = smdbItem['FED FROM'] || '';
                            if (smdbFedFrom.includes('MDB')) {
                                const mdbMatch = smdbFedFrom.match(/MDB\d/);
                                if (mdbMatch && mdbNodes[mdbMatch[0]]) {
                                    if (!mdbNodes[mdbMatch[0]].children.find(c => c.id === parentName)) {
                                        mdbNodes[mdbMatch[0]].children.push(parentNode);
                                    }
                                }
                            }
                        }
                    }
                }
                
                // Create group node
                const groupNode = {
                    name: `${itemName} (${count} copies)`,
                    id: `${parentName}_${itemName}_group`,
                    kind: kind,
                    load: load,
                    estimate: estimate,
                    mdb: normalizeMDB(mdb),
                    data: { ...item, copyCount: count, fedFromSMDB: parentName, isGroup: true },
                    groupCount: count,
                    collapsed: true,
                    children: []
                };
                
                parentNode.children.push(groupNode);
            }
        });
    });
    
    // Second, process SMDBs/MDBs that have special copy counts but may not be in FED FROM
    // This handles cases like SMDB.TN.P2.01, SMDB.TH.B1.01, and MDB4 with DB.TH.GF.01
    allDataArray.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        
        // Check if this is SMDB.TN.*, SMDB.TH.*, or MDB4/MDB.GF.04
        const isSMDBTN = itemName.startsWith('SMDB.TN.');
        const isSMDBTH = itemName.startsWith('SMDB.TH.');
        const isMDB4 = itemName === 'MDB4' || itemName === 'MDB.GF.04' || itemName === 'MDB GF 04';
        
        if (!isSMDBTN && !isSMDBTH && !isMDB4) return;
        
        const copyCounts = getDBCopyCounts(itemName);
        if (!copyCounts) return;
        
        // Find or create SMDB/MDB node
        let parentNode = itemMap.get(itemName);
        if (!parentNode) {
            if (isMDB4) {
                // For MDB4, use the existing MDB node
                const mdb4Name = 'MDB4';
                parentNode = mdbNodes[mdb4Name];
                if (!parentNode) {
                    // Create MDB4 node if it doesn't exist
                    parentNode = {
                        name: mdb4Name,
                        id: mdb4Name,
                        kind: 'MDB',
                        children: []
                    };
                    mdbNodes[mdb4Name] = parentNode;
                    if (rmuNode && !rmuNode.children.find(c => c.id === mdb4Name)) {
                        rmuNode.children.push(parentNode);
                    }
                }
            } else {
                // For SMDB, create new node
                parentNode = {
                    name: itemName,
                    id: itemName,
                    kind: 'SMDB',
                    children: []
                };
                itemMap.set(itemName, parentNode);
                
                // Connect SMDB to its parent (find from data)
                const smdbFedFrom = item['FED FROM'] || '';
                if (smdbFedFrom.includes('MDB')) {
                    const mdbMatch = smdbFedFrom.match(/MDB\d/);
                    if (mdbMatch && mdbNodes[mdbMatch[0]]) {
                        if (!mdbNodes[mdbMatch[0]].children.find(c => c.id === itemName)) {
                            mdbNodes[mdbMatch[0]].children.push(parentNode);
                        }
                    }
                }
            }
        }
        
        // Create group nodes for each special DB type
        specialDBs.forEach(dbType => {
            if (copyCounts[dbType]) {
                const count = copyCounts[dbType];
                
                // Check if group node already exists
                const groupId = `${itemName}_${dbType}_group`;
                if (parentNode.children.find(c => c.id === groupId)) return;
                
                // Find the DB item data
                const dbItem = allDataArray.find(d => (d.Itemdrop || d.MDB) === dbType);
                const kind = dbItem ? (dbItem.KIND || 'Unknown') : 'DB';
                // Get load from LOAD column (try both cases for column name)
                const load = dbItem ? ((dbItem.Load !== undefined && dbItem.Load !== null && dbItem.Load !== '') 
                    ? dbItem.Load 
                    : ((dbItem.LOAD !== undefined && dbItem.LOAD !== null && dbItem.LOAD !== '') ? dbItem.LOAD : '0 kW')) : '0 kW';
                const estimate = dbItem ? (dbItem.Estimate || 0) : 0;
                const mdb = dbItem ? (dbItem.MDB || '') : '';
                
                // Create group node
                const groupNode = {
                    name: `${dbType} (${count} copies)`,
                    id: groupId,
                    kind: kind,
                    load: load,
                    estimate: estimate,
                    mdb: normalizeMDB(mdb),
                    data: dbItem ? { ...dbItem, copyCount: count, fedFromSMDB: itemName, isGroup: true } : { copyCount: count, fedFromSMDB: itemName, isGroup: true },
                    groupCount: count,
                    collapsed: true,
                    children: []
                };
                
                parentNode.children.push(groupNode);
            }
        });
    });
    
    // Third, process special ESMDBs that appear under multiple parents (BB.05 and EMDB.GF.01)
    const specialESMDBs = ['ESMDB.LL.RF.01(LIFT)', 'ESMDB.LL.RF.02(LIFT)'];
    const esmdbParents = ['BB.05', 'EMDB.GF.01'];
    
    esmdbParents.forEach(parentName => {
        const copyCounts = getESMDBCopyCounts(parentName);
        if (!copyCounts) return;
        
        // Find or create parent node (BB.05 or EMDB.GF.01)
        let parentNode = itemMap.get(parentName);
        
        // Try to find parent in existing tree structure
        if (!parentNode) {
            // Search for BB.05 or EMDB.GF.01 in the tree
            function findParentInTree(node, name) {
                if (!node) return null;
                if (node.name === name || node.id === name) return node;
                
                const children = node.children || [];
                for (const child of children) {
                    const found = findParentInTree(child, name);
                    if (found) return found;
                }
                return null;
            }
            
            // Search from RMU node
            if (rmuNode) {
                parentNode = findParentInTree(rmuNode, parentName);
            }
            
            // If still not found, try to find from allDataArray
            if (!parentNode) {
                const parentItem = allDataArray.find(d => (d.Itemdrop || d.MDB) === parentName);
                if (parentItem) {
                    const kind = parentItem.KIND || 'BUS BAR RAISER';
                    parentNode = {
                        name: parentName,
                        id: parentName,
                        kind: kind,
                        children: []
                    };
                    itemMap.set(parentName, parentNode);
                    
                    // Try to connect to its parent (MDB)
                    const parentFedFrom = parentItem['FED FROM'] || '';
                    if (parentFedFrom.includes('MDB')) {
                        const mdbMatch = parentFedFrom.match(/MDB\d/);
                        if (mdbMatch && mdbNodes[mdbMatch[0]]) {
                            if (!mdbNodes[mdbMatch[0]].children.find(c => c.id === parentName)) {
                                mdbNodes[mdbMatch[0]].children.push(parentNode);
                            }
                        }
                    }
                }
            }
        }
        
        if (!parentNode) {
            console.warn(`Parent node not found for special ESMDBs: ${parentName}`);
            return;
        }
        
        // Create group nodes for each special ESMDB type
        specialESMDBs.forEach(esmdbType => {
            if (copyCounts[esmdbType]) {
                const count = copyCounts[esmdbType];
                
                // Check if group node already exists
                const groupId = `${parentName}_${esmdbType}_group`;
                if (parentNode.children.find(c => c.id === groupId)) return;
                
                // Find the ESMDB item data
                const esmdbItem = allDataArray.find(d => (d.Itemdrop || d.MDB) === esmdbType);
                const kind = esmdbItem ? (esmdbItem.KIND || 'ESMDB') : 'ESMDB';
                // Get load from LOAD column (try both cases for column name)
                const load = esmdbItem ? ((esmdbItem.Load !== undefined && esmdbItem.Load !== null && esmdbItem.Load !== '') 
                    ? esmdbItem.Load 
                    : ((esmdbItem.LOAD !== undefined && esmdbItem.LOAD !== null && esmdbItem.LOAD !== '') ? esmdbItem.LOAD : '0 kW')) : '0 kW';
                const estimate = esmdbItem ? (esmdbItem.Estimate || 0) : 0;
                const mdb = esmdbItem ? (esmdbItem.MDB || '') : '';
                
                // Create group node
                const groupNode = {
                    name: `${esmdbType} (${count} copies)`,
                    id: groupId,
                    kind: kind,
                    load: load,
                    estimate: estimate,
                    mdb: normalizeMDB(mdb),
                    data: esmdbItem ? { ...esmdbItem, copyCount: count, fedFromParent: parentName, isGroup: true } : { copyCount: count, fedFromParent: parentName, isGroup: true },
                    groupCount: count,
                    collapsed: true,
                    children: []
                };
                
                parentNode.children.push(groupNode);
            }
        });
    });
    
    // Filter out duplicate MDB nodes: remove "MDB" from MDB1 children, "MDB.GF.04" from MDB4 children
    function filterDuplicateMDBChildren(node) {
        if (!node.children || node.children.length === 0) return;
        
        // Filter children based on parent node
        if (node.name === 'MDB1' || node.id === 'MDB1') {
            // Remove "MDB" or "MDB1" children from MDB1
            node.children = node.children.filter(child => {
                const childName = (child.name || '').toString().trim();
                const childId = (child.id || '').toString().trim();
                return childName !== 'MDB' && childName !== 'MDB1' && 
                       childId !== 'MDB' && childId !== 'MDB1';
            });
        } else if (node.name === 'MDB4' || node.id === 'MDB4') {
            // Remove "MDB.GF.04" or "MDB4" children from MDB4
            node.children = node.children.filter(child => {
                const childName = (child.name || '').toString().trim();
                const childId = (child.id || '').toString().trim();
                return childName !== 'MDB.GF.04' && childName !== 'MDB GF 04' && 
                       childName !== 'MDB4' &&
                       childId !== 'MDB.GF.04' && childId !== 'MDB GF 04' && 
                       childId !== 'MDB4';
            });
        }
        
        // Recursively filter children
        node.children.forEach(child => filterDuplicateMDBChildren(child));
    }
    
    // Apply filter to remove duplicate MDB nodes
    filterDuplicateMDBChildren(rmuNode);
    
    treeData = rmuNode;
    root = d3.hierarchy(treeData);
    root.x0 = 0;
    root.y0 = 0;
}

function initializeDiagram() {
    const container = d3.select('#mindmap');
    container.selectAll('*').remove();
    
    const width = container.node().offsetWidth || 1200;
    const height = Math.max(600, window.innerHeight * 0.6);
    
    // Create SVG
    svg = container
        .append('svg')
        .attr('width', width)
        .attr('height', height)
        .attr('viewBox', `0 0 ${width} ${height}`);
    
    // Create zoom behavior
    zoom = d3.zoom()
        .scaleExtent([0.1, 3])
        .on('zoom', (event) => {
            g.attr('transform', event.transform);
        });
    
    svg.call(zoom);
    
    // Create main group
    g = svg.append('g');
    
    // Create tree layout
    const tree = d3.tree()
        .size([height - 100, width - 200])
        .separation((a, b) => (a.parent === b.parent ? 1 : 1.5) / a.depth);
    
    // Update function
    function update(source) {
        const treeData = tree(root);
        const nodes = treeData.descendants();
        const links = treeData.links();
        
        // Normalize for fixed-depth
        nodes.forEach(d => {
            d.y = d.depth * 180;
        });
        
        // Update nodes
        const node = g.selectAll('g.node')
            .data(nodes, d => d.data.id);
        
        const nodeEnter = node.enter()
            .append('g')
            .attr('class', 'node')
            .attr('transform', d => `translate(${source.y0},${source.x0})`)
            .on('click', click);
        
        // Add circles for nodes
        nodeEnter.append('circle')
            .attr('r', d => getNodeRadius(d))
            .attr('fill', d => getNodeColor(d.data.kind))
            .attr('stroke', d => getNodeBorderColor(d.data.kind))
            .attr('stroke-width', d => d.data.id === 'RMU' ? 3 : 2)
            .style('cursor', 'pointer');
        
        // Add labels (clickable to open detailed page)
        const labels = nodeEnter.append('text')
            .attr('dy', d => d.data.id === 'RMU' ? 5 : 4)
            .attr('x', d => {
                // Group nodes need more space for longer labels
                if (d.data.isGroup || (d.data.name && d.data.name.includes('copies'))) {
                    return d.children || d._children ? -15 : 15;
                }
                return d.children || d._children ? -13 : 13;
            })
            .attr('text-anchor', d => d.children || d._children ? 'end' : 'start')
            .attr('font-size', d => getNodeFontSize(d))
            .attr('font-weight', d => {
                if (d.data.id === 'RMU' || d.data.kind === 'MDB') return 'bold';
                // Make group nodes (with copy counts) slightly bolder for readability
                if (d.data.isGroup || (d.data.name && d.data.name.includes('copies'))) return '600';
                return 'normal';
            })
            .text(d => {
                let name = d.data.name;
                // Normalize MDB names for display - remove "MDB" from MDB1, "MDB.GF.04" from MDB4
                // Check if this is a child MDB node that should be normalized
                if (name === 'MDB' && d.parent && d.parent.data && (d.parent.data.name === 'MDB1' || d.parent.data.id === 'MDB1')) {
                    name = 'MDB1';
                } else if ((name === 'MDB.GF.04' || name === 'MDB GF 04') && d.parent && d.parent.data && (d.parent.data.name === 'MDB4' || d.parent.data.id === 'MDB4')) {
                    name = 'MDB4';
                } else {
                    // Apply general normalization
                    name = normalizeMDB(name);
                }
                // Don't truncate group nodes (they contain copy counts)
                if (d.data.isGroup || name.includes('copies')) {
                    return name;
                }
                // Truncate other long names
                return name.length > 20 ? name.substring(0, 18) + '...' : name;
            })
            .style('pointer-events', 'all')
            .style('fill', '#333')
            .style('cursor', d => d.data.data ? 'pointer' : 'default')
            .on('click', function(event, d) {
                event.stopPropagation();
                // Handle MDB nodes - they should be clickable
                if (d.data.data) {
                    openDetailedPage(d.data.data, d.data.name);
                } else if (d.data.kind === 'MDB' && d.data.name) {
                    // MDB nodes should be clickable - use data from node
                    const mdbData = d.data.data || {
                        Itemdrop: d.data.name,
                        MDB: d.data.name,
                        KIND: 'MDB',
                        Load: d.data.load || '0 kW',
                        Estimate: d.data.estimate || 0,
                        'NO OF ITEMS': 0,
                        'FED FROM': 'RMU'
                    };
                    openDetailedPage(mdbData, d.data.name);
                }
            });
        
        // Add expand/collapse indicators
        nodeEnter.filter(d => d.children || d._children)
            .append('text')
            .attr('dy', 5)
            .attr('x', d => d.children ? -8 : -8)
            .attr('text-anchor', 'middle')
            .attr('font-size', '12px')
            .text(d => d.children ? '▼' : '▶')
            .style('fill', '#666')
            .style('pointer-events', 'none');
        
        // Update node positions
        const nodeUpdate = nodeEnter.merge(node);
        nodeUpdate.transition()
            .duration(300)
            .attr('transform', d => `translate(${d.y},${d.x})`);
        
        // Remove exiting nodes
        const nodeExit = node.exit()
            .transition()
            .duration(300)
            .attr('transform', d => `translate(${source.y},${source.x})`)
            .remove();
        
        // Update links
        const link = g.selectAll('path.link')
            .data(links, d => d.target.data.id);
        
        const linkEnter = link.enter()
            .insert('path', 'g')
            .attr('class', 'link')
            .attr('d', d => {
                const o = { x: source.x0, y: source.y0 };
                return diagonal({ source: o, target: o });
            })
            .attr('fill', 'none')
            .attr('stroke', '#999')
            .attr('stroke-width', 2);
        
        const linkUpdate = linkEnter.merge(link);
        linkUpdate.transition()
            .duration(300)
            .attr('d', diagonal);
        
        link.exit()
            .transition()
            .duration(300)
            .attr('d', d => {
                const o = { x: source.x, y: source.y };
                return diagonal({ source: o, target: o });
            })
            .remove();
        
        // Store positions for transitions
        nodes.forEach(d => {
            d.x0 = d.x;
            d.y0 = d.y;
        });
    }
    
    function diagonal(d) {
        return `M ${d.source.y},${d.source.x}
                C ${(d.source.y + d.target.y) / 2},${d.source.x}
                  ${(d.source.y + d.target.y) / 2},${d.target.x}
                  ${d.target.y},${d.target.x}`;
    }
    
    function click(event, d) {
        // Don't handle clicks on text labels (they have their own handler)
        if (event.target.tagName === 'text') {
            return;
        }
        
        // Regular click for expand/collapse
        if (d.children) {
            d._children = d.children;
            d.children = null;
        } else {
            d.children = d._children;
            d._children = null;
        }
        
        // Handle group nodes
        if (d.data.isGroup && d.data.groupCount) {
            if (d.children) {
                // Expand group
                expandGroupNode(d);
            } else {
                // Collapse group
                collapseGroupNode(d);
            }
        }
        
        // Double click on node circle opens detailed page (fallback)
        // Also handle MDB nodes (MDB1, MDB2, MDB3, MDB4) - they should be clickable
        if (event.detail === 2) {
            if (d.data.data) {
                openDetailedPage(d.data.data, d.data.name);
            } else if (d.data.kind === 'MDB' && d.data.name) {
                // MDB nodes should be clickable - use data from node
                // Get load from node data or use LOAD column from TOTALLIST
                const nodeLoad = d.data.load || '0 kW';
                const mdbData = d.data.data || {
                    Itemdrop: d.data.name,
                    MDB: d.data.name,
                    KIND: 'MDB',
                    Load: nodeLoad,
                    LOAD: nodeLoad, // Also set LOAD for consistency
                    Estimate: d.data.estimate || 0,
                    'NO OF ITEMS': 0,
                    'FED FROM': 'RMU'
                };
                openDetailedPage(mdbData, d.data.name);
            }
        }
        
        update(d);
    }
    
    function expandGroupNode(d) {
        if (!d.data.groupCount || d.children) return;
        
        d.children = [];
        for (let i = 1; i <= d.data.groupCount; i++) {
            d.children.push({
                name: `Copy #${i}`,
                id: `${d.data.id}_copy${i}`,
                kind: d.data.kind,
                load: d.data.load,
                estimate: d.data.estimate,
                mdb: normalizeMDB(d.data.mdb),
                data: { ...d.data.data, copyNumber: i },
                children: null
            });
        }
    }
    
    function collapseGroupNode(d) {
        if (d.children) {
            d._children = d.children;
            d.children = null;
        }
    }
    
    // Collapse all nodes initially
    root.children.forEach(collapse);
    root.children.forEach(collapse);
    root.children.forEach(collapse);
    
    update(root);
    
    // Center the diagram
    const bounds = g.node().getBBox();
    const fullWidth = width;
    const fullHeight = height;
    const widthScale = (fullWidth - 40) / bounds.width;
    const heightScale = (fullHeight - 40) / bounds.height;
    const scale = Math.min(widthScale, heightScale, 1);
    
    const translate = [
        (fullWidth - bounds.width * scale) / 2 - bounds.x * scale,
        (fullHeight - bounds.height * scale) / 2 - bounds.y * scale
    ];
    
    g.attr('transform', `translate(${translate}) scale(${scale})`);
    
    function collapse(d) {
        if (d.children) {
            d._children = d.children;
            d._children.forEach(collapse);
            d.children = null;
        }
    }
}

function getNodeRadius(d) {
    if (d.data.id === 'RMU') return 25;
    if (d.data.kind === 'MDB' || d.data.id === 'RMU') return 20;
    if (d.data.kind === 'BUS BAR RAISER' || d.data.kind === 'SMDB') return 15;
    return 12;
}

function getNodeColor(kind) {
    const colorMap = {
        'RMU': '#667eea',
        'MDB': '#FF6B6B',
        'BUS BAR RAISER': '#4ECDC4',
        'SMDB': '#95E1D3',
        'DB': '#A8E6CF',
        'EDB': '#FFD93D',
        'EMDB': '#C7CEEA',
        'EMCC': '#FFB6C1',
        'ESMDB': '#DDA0DD',
        'MCC': '#98D8C8',
        'POWR FACTOR CORRECTOR': '#FFE66D',
        'EFP': '#F7DC6F',
        'CP': '#AED6F1',
        'EVCS': '#F8C471',
        'CHWP': '#85C1E2',
        'AHU': '#F1948A',
        'CU': '#F9E79F',
        'GSM': '#BB8FCE',
        'CBS': '#85C1E9',
        'UDB': '#F7DC6F',
        'SPF': '#A3E4D7',
        'JEF': '#FAD7A0'
    };
    return colorMap[kind] || '#D3D3D3';
}

function getNodeBorderColor(kind) {
    const colorMap = {
        'RMU': '#7D3C98',
        'MDB': '#C92A2A',
        'BUS BAR RAISER': '#0E7C7B',
        'SMDB': '#2D8659',
        'DB': '#2D8659'
    };
    return colorMap[kind] || '#808080';
}

function getNodeFontSize(d) {
    if (d.data.id === 'RMU') return '14px';
    // Group nodes (with copy counts) should be more readable
    if (d.data.isGroup || (d.data.name && d.data.name.includes('copies'))) return '11px';
    if (d.data.kind === 'MDB' || d.data.id === 'RMU') return '12px';
    if (d.data.kind === 'BUS BAR RAISER' || d.data.kind === 'SMDB') return '11px';
    return '10px';
}

function updateStats() {
    const data = window.allData || [];
    if (!data || data.length === 0) return;
    
    const totalItems = data.length;
    let totalLoad = 0;
    let totalEstimate = 0;
    
    // Calculate Total Load: Sum of loads from MDB1, MDB2, MDB3, MDB4 only
    const mdbNames = ['MDB1', 'MDB2', 'MDB3', 'MDB4'];
    mdbNames.forEach(mdbName => {
        const mdbItem = data.find(item => {
            const itemName = normalizeMDB(item.Itemdrop || item.MDB || '');
            return itemName === mdbName || item.Itemdrop === mdbName || item.MDB === mdbName;
        });
        
        if (mdbItem) {
            // Get load from LOAD column (try both cases for column name)
            const loadValue = (mdbItem.Load !== undefined && mdbItem.Load !== null && mdbItem.Load !== '') 
                ? mdbItem.Load 
                : ((mdbItem.LOAD !== undefined && mdbItem.LOAD !== null && mdbItem.LOAD !== '') ? mdbItem.LOAD : '0 kW');
            const loadNum = parseLoadValue(loadValue);
            totalLoad += loadNum;
            logger.log(`MDB Load: ${mdbName} = ${loadNum.toFixed(2)} kW`);
        } else {
            logger.warn(`MDB not found: ${mdbName}`);
        }
    });
    
    // Calculate Total Estimate: Sum of all Estimate column values in TOTALLIST sheet
    // Exclude total rows and empty rows
    data.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        
        // Skip empty rows
        if (!itemName || itemName.trim() === '') {
            return;
        }
        
        // Skip total rows
        const itemdropLower = itemName.toLowerCase().trim();
        if (itemdropLower === 'total' || itemdropLower === 'net total' || 
            itemdropLower.includes('sum after')) {
            return;
        }
        
        // Get estimate directly from ESTIMATE column
        const estimateValue = item.Estimate || item.ESTIMATE; // Try both cases
        if (estimateValue !== null && estimateValue !== undefined && estimateValue !== '') {
            const parsed = parseFloat(estimateValue);
            if (!isNaN(parsed)) {
                totalEstimate += parsed;
            }
        }
    });
    
    logger.log(`Total Estimate: Sum of all Estimate column values = ${totalEstimate.toFixed(2)}`);
    logger.log(`Total Load: Sum of MDB1, MDB2, MDB3, MDB4 loads = ${totalLoad.toFixed(2)} kW`);
    
    document.getElementById('totalItems').textContent = totalItems;
    document.getElementById('totalLoad').textContent = totalLoad.toFixed(2) + ' kW';
    document.getElementById('totalEstimate').textContent = formatNumber(totalEstimate);
}

function populateFilters() {
    const kindFilter = document.getElementById('kindFilter');
    const data = window.allData || [];
    if (!data || data.length === 0) return;
    const kinds = [...new Set(data.map(item => item.KIND).filter(Boolean))].sort();
    
    kinds.forEach(kind => {
        const option = document.createElement('option');
        option.value = kind;
        option.textContent = kind;
        kindFilter.appendChild(option);
    });
}

function formatNumber(num) {
    // Handle null, undefined, empty string, or NaN
    if (num === null || num === undefined || num === '' || (typeof num === 'number' && isNaN(num))) {
        return 'N/A';
    }
    // Handle 0 as a valid number
    if (num === 0 || num === '0') {
        return '0.00';
    }
    // Convert to number if it's a string
    const numValue = typeof num === 'string' ? parseFloat(num) : num;
    if (isNaN(numValue)) {
        return 'N/A';
    }
    return new Intl.NumberFormat('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(numValue);
}

function showDetails(item) {
    const panel = document.getElementById('details-panel');
    const content = document.getElementById('detailsContent');
    
    const isCopy = item.copyNumber && item.fedFromSMDB;
    const itemName = item.Itemdrop || item.MDB || 'N/A';
    const fedFrom = isCopy ? item.fedFromSMDB : (item['FED FROM'] || 'N/A');
    
    content.innerHTML = `
        <div class="detail-item">
            <div class="detail-label">Item Name:</div>
            <div class="detail-value">${itemName}${isCopy ? ` (Copy #${item.copyNumber})` : ''}</div>
        </div>
        ${isCopy ? `
        <div class="detail-item">
            <div class="detail-label">Copy Number:</div>
            <div class="detail-value">${item.copyNumber}</div>
        </div>
        ` : ''}
        <div class="detail-item">
            <div class="detail-label">Kind:</div>
            <div class="detail-value">${item.KIND || 'N/A'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">MDB:</div>
            <div class="detail-value">${normalizeMDB(item.MDB) || 'N/A'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Fed From:</div>
            <div class="detail-value">${fedFrom}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Load:</div>
            <div class="detail-value">${(item.Load !== undefined && item.Load !== null && item.Load !== '') 
                ? item.Load 
                : ((item.LOAD !== undefined && item.LOAD !== null && item.LOAD !== '') ? item.LOAD : 'N/A')}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Number of Items:</div>
            <div class="detail-value">${item['NO OF ITEMS'] || 'N/A'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Estimate:</div>
            <div class="detail-value">${formatNumber(item.Estimate)}</div>
        </div>
    `;
    
    panel.classList.remove('hidden');
}

// Open detailed page with metadata and Excel sheet data
function openDetailedPage(item, nodeName, sourceView = 'main') {
    const detailedPage = document.getElementById('detailed-page');
    const detailedTitle = document.getElementById('detailed-title');
    const metadataContent = document.getElementById('metadata-content');
    const detailedDataContent = document.getElementById('detailed-data-content');
    
    // Store the source view so we know where to return
    detailedPage.setAttribute('data-source-view', sourceView);
    
    // Hide main content or MDB view based on source
    if (sourceView === 'mdb') {
        document.getElementById('mdb-view-page').classList.add('hidden');
    } else {
        document.querySelector('main > .controls').style.display = 'none';
        document.querySelector('main > .legend-panel').style.display = 'none';
        document.querySelector('main > .info-panel').style.display = 'none';
        document.getElementById('main-content').style.display = 'none';
    }
    document.getElementById('details-panel').classList.add('hidden');
    
    // Show detailed page
    detailedPage.classList.remove('hidden');
    
    // Set title - normalize MDB names
    let itemName = item ? (item.Itemdrop || item.MDB || nodeName) : nodeName || 'N/A';
    // Normalize MDB names for display (MDB -> MDB1, MDB.GF.04 -> MDB4)
    if (itemName === 'MDB' || itemName === 'MDB.GF.04') {
        itemName = normalizeMDB(itemName);
    }
    detailedTitle.textContent = itemName;
    
    // Display metadata (use item if available, otherwise create minimal metadata)
    if (item) {
        displayMetadata(item, metadataContent);
    } else {
        // Create minimal metadata for nodes without data
        metadataContent.innerHTML = `
            <table class="metadata-table">
                <tr>
                    <td>Item Name</td>
                    <td>${itemName}</td>
                </tr>
            </table>
        `;
    }
    
    // Load and display detailed data from Excel sheet
    detailedDataContent.innerHTML = '<div class="loading">Loading detailed data...</div>';
    
    // Normalize sheet name for MDB nodes
    let sheetName = itemName;
    if (itemName === 'MDB') {
        sheetName = 'MDB1'; // MDB sheet should be treated as MDB1
    } else if (itemName === 'MDB.GF.04') {
        sheetName = 'MDB4'; // MDB.GF.04 should be treated as MDB4
    }
    
    // Check if this is a special DB that needs recalculation
    const specialDBs = ['DB.TN.LXX.1B1.01', 'DB.TN.LXX.2B1.01', 'DB.TN.LXX.3B1.01', 'DB.TH.GF.01'];
    const isSpecialDB = specialDBs.includes(itemName);
    
    if (isSpecialDB && item) {
        // Use NO OF ITEMS from TOTALLIST sheet (same as NO OF UNITS)
        // Try both 'NO OF ITEMS' and 'NO OF UNITS' column names, and also check copy counts as fallback
        let noOfItems = item['NO OF ITEMS'] || item['NO OF UNITS'];
        
        // If not found in item, try to get from copy counts (fallback)
        if (!noOfItems || noOfItems === 0) {
            const smdbName = findParentSMDB(itemName, item);
            if (smdbName) {
                const copyCounts = getDBCopyCounts(smdbName);
                noOfItems = copyCounts && copyCounts[itemName] ? copyCounts[itemName] : 1;
            } else {
                noOfItems = 1; // Default fallback
            }
        }
        
        // Convert to number if it's a string
        const noOfUnits = typeof noOfItems === 'string' ? parseFloat(noOfItems) || 1 : (noOfItems || 1);
        
        // Load with recalculation using NO OF ITEMS value
        loadDetailedDataWithRecalculation(itemName, detailedDataContent, noOfUnits);
    } else {
        // Normal loading (SMDBs, MDBs, and other items show as-is)
        loadDetailedData(sheetName, detailedDataContent);
    }
}

// Find parent SMDB for a special DB
function findParentSMDB(dbName, item) {
    // Check if item has fedFromSMDB (for copy nodes)
    if (item.fedFromSMDB) {
        return item.fedFromSMDB;
    }
    
    // Check FED FROM field
    const fedFrom = item['FED FROM'] || '';
    if (fedFrom) {
        // FED FROM can contain multiple SMDBs separated by newlines
        const smdbList = fedFrom.split('\n').map(p => p.trim()).filter(p => p && p !== 'RMU' && !p.startsWith('MDB'));
        if (smdbList.length > 0) {
            // Return the first SMDB (or we could check which one has this DB)
            return smdbList[0];
        }
    }
    
    // Search the tree structure to find parent SMDB
    if (treeData && treeData.children) {
        for (const mdb of treeData.children) {
            if (mdb.children) {
                for (const child of mdb.children) {
                    if (child.children) {
                        for (const grandchild of child.children) {
                            // Check if this is a group node containing our DB
                            if (grandchild.data && grandchild.data.isGroup && grandchild.data.name && grandchild.data.name.includes(dbName)) {
                                return child.name; // Return the SMDB name
                            }
                            // Check if this is the DB node itself
                            if (grandchild.name === dbName || grandchild.id === dbName) {
                                return child.name; // Return the SMDB name
                            }
                        }
                    }
                }
            }
        }
    }
    
    return null;
}

// Display metadata in table format
function displayMetadata(item, container) {
    const isCopy = item.copyNumber && item.fedFromSMDB;
    const itemName = item.Itemdrop || item.MDB || 'N/A';
    const fedFrom = isCopy ? item.fedFromSMDB : (item['FED FROM'] || 'N/A');
    
    // Get load value
    const itemLoad = (item.Load !== undefined && item.Load !== null && item.Load !== '') 
        ? item.Load 
        : ((item.LOAD !== undefined && item.LOAD !== null && item.LOAD !== '') ? item.LOAD : 'N/A');
    
    // Check if this is a board (MDB, SMDB, ESMDB, MCC, etc.) and validate load against children
    // Apply validation to all boards until DB level
    let loadValidationHtml = '';
    const kind = (item.KIND || '').toString().toUpperCase();
    const normalizedMDB = normalizeMDB(itemName);
    
    // Check if this is a board type that should have load validation (not DB level)
    const isBoardType = kind === 'MDB' || 
                        kind === 'SMDB' || 
                        kind === 'ESMDB' || 
                        kind === 'MCC' ||
                        kind === 'BUS BAR RAISER' ||
                        kind === 'EMDB' ||
                        kind === 'EMCC';
    
    if (isBoardType) {
        // Get board's base load and number of units
        const boardBaseLoad = parseLoadValue(itemLoad);
        const boardNoOfUnits = parseFloat(item['NO OF ITEMS'] || item['NO OF UNITS'] || 1);
        // Board's effective load = base load × number of units
        const boardLoad = boardBaseLoad * boardNoOfUnits;
        
        // Calculate children load sum asynchronously (will load detailed sheets for DBs)
        calculateChildrenLoadSum(itemName).then(childrenLoadSum => {
            if (boardLoad > 0 && childrenLoadSum > 0) {
                const difference = Math.abs(boardLoad - childrenLoadSum);
                const percentageDiff = (difference / boardLoad) * 100;
                const tolerancePercent = 5; // ±5% tolerance
                
                let validationRow = '';
                if (percentageDiff > tolerancePercent) {
                    // Load mismatch detected (>5% difference)
                    const statusColor = '#f44336'; // Red for >5% difference
                    validationRow = `
                        <tr style="background-color: ${statusColor}20;">
                            <td colspan="2" style="padding: 8px; color: ${statusColor}; font-weight: 600;">
                                ⚠ Load Validation: ${kind} Load (${boardBaseLoad.toFixed(2)} kW × ${boardNoOfUnits} = ${boardLoad.toFixed(2)} kW) vs Sum of Children (${childrenLoadSum.toFixed(2)} kW)
                                <br>Difference: ${difference.toFixed(2)} kW (${percentageDiff.toFixed(2)}%) - Exceeds ±5% tolerance
                            </td>
                        </tr>
                    `;
                    console.warn(`Load mismatch for ${itemName} (${kind}): Board=${boardBaseLoad.toFixed(2)} kW × ${boardNoOfUnits} = ${boardLoad.toFixed(2)} kW, Children Sum=${childrenLoadSum.toFixed(2)} kW, Diff=${difference.toFixed(2)} kW (${percentageDiff.toFixed(2)}%)`);
                } else {
                    // Load matches (within ±5% tolerance)
                    validationRow = `
                        <tr style="background-color: #4CAF5020;">
                            <td colspan="2" style="padding: 8px; color: #4CAF50; font-weight: 600;">
                                ✓ Load Validation: ${kind} Load (${boardBaseLoad.toFixed(2)} kW × ${boardNoOfUnits} = ${boardLoad.toFixed(2)} kW) ≈ Sum of Children (${childrenLoadSum.toFixed(2)} kW)
                                <br>Difference: ${difference.toFixed(2)} kW (${percentageDiff.toFixed(2)}%) - Within ±5% tolerance
                            </td>
                        </tr>
                    `;
                    console.log(`Load validated for ${itemName} (${kind}): Board=${boardBaseLoad.toFixed(2)} kW × ${boardNoOfUnits} = ${boardLoad.toFixed(2)} kW, Children Sum=${childrenLoadSum.toFixed(2)} kW, Diff=${difference.toFixed(2)} kW (${percentageDiff.toFixed(2)}%)`);
                }
                
                // Update the metadata table with validation result
                const metadataTable = container.querySelector('.metadata-table');
                if (metadataTable) {
                    const tbody = metadataTable.querySelector('tbody') || metadataTable;
                    tbody.insertAdjacentHTML('beforeend', validationRow);
                }
            } else if (boardLoad > 0 && childrenLoadSum === 0) {
                // Board has load but no children to compare
                const validationRow = `
                    <tr style="background-color: #99920;">
                        <td colspan="2" style="padding: 8px; color: #999; font-weight: 600;">
                            ⓘ Load Validation: ${kind} Load (${boardBaseLoad.toFixed(2)} kW × ${boardNoOfUnits} = ${boardLoad.toFixed(2)} kW) - No children to compare
                        </td>
                    </tr>
                `;
                const metadataTable = container.querySelector('.metadata-table');
                if (metadataTable) {
                    const tbody = metadataTable.querySelector('tbody') || metadataTable;
                    tbody.insertAdjacentHTML('beforeend', validationRow);
                }
            }
        }).catch(error => {
            console.error(`Error calculating children load sum for ${itemName}:`, error);
        });
    }
    
    container.innerHTML = `
        <table class="metadata-table">
            <tr>
                <td>Item Name</td>
                <td>${itemName}${isCopy ? ` (Copy #${item.copyNumber})` : ''}</td>
            </tr>
            ${isCopy ? `
            <tr>
                <td>Copy Number</td>
                <td>${item.copyNumber}</td>
            </tr>
            ` : ''}
            <tr>
                <td>Kind</td>
                <td>${item.KIND || 'N/A'}</td>
            </tr>
            <tr>
                <td>MDB</td>
                <td>${normalizeMDB(item.MDB) || 'N/A'}</td>
            </tr>
            <tr>
                <td>Fed From</td>
                <td>${fedFrom}</td>
            </tr>
            <tr>
                <td>Load</td>
                <td>${itemLoad}</td>
            </tr>
            ${loadValidationHtml}
            <tr>
                <td>Number of Items</td>
                <td>${item['NO OF ITEMS'] || 'N/A'}</td>
            </tr>
            <tr>
                <td>Estimate</td>
                <td>${formatNumber(item.Estimate)}</td>
            </tr>
        </table>
    `;
}

// Load detailed data from Excel sheet
function loadDetailedData(itemName, container) {
    // Try to fetch the Excel file
    fetch(EXCEL_FILE_NAME)
        .then(response => {
            if (!response.ok) {
                throw new Error('Excel file not found');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Normal loading - show sheet as-is (no recalculation)
            loadNormalSheet(itemName, container, workbook);
        })
        .catch(error => {
            console.error('Error loading Excel file:', error);
            container.innerHTML = `
                <div class="error-message">
                    Error loading detailed data: ${error.message}
                    <br><small>Make sure e2.xlsx is in the same directory</small>
                </div>
            `;
        });
}

// Load detailed data with recalculation for special DBs
function loadDetailedDataWithRecalculation(dbName, container, noOfUnits) {
    // Try to fetch the Excel file
    fetch(EXCEL_FILE_NAME)
        .then(response => {
            if (!response.ok) {
                throw new Error('Excel file not found');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetNames = workbook.SheetNames;
            
            // Find DB sheet
            let dbSheetName = sheetNames.find(name => 
                name === dbName || 
                name.toLowerCase() === dbName.toLowerCase()
            );
            
            if (!dbSheetName) {
                dbSheetName = sheetNames.find(name => 
                    name.toLowerCase().includes(dbName.toLowerCase()) ||
                    dbName.toLowerCase().includes(name.toLowerCase())
                );
            }
            
            if (!dbSheetName) {
                container.innerHTML = `
                    <div class="error-message">
                        No matching sheet found for "${dbName}". Available sheets: ${sheetNames.join(', ')}
                    </div>
                `;
                return;
            }
            
            const dbWorksheet = workbook.Sheets[dbSheetName];
            const dbData = XLSX.utils.sheet_to_json(dbWorksheet, { header: 1 });
            
            // Calculate estimate for this DB
            const calculation = calculateDBEstimate(dbData, noOfUnits);
            
            if (!calculation) {
                // Fallback to normal display if calculation fails
                displayDetailedTable(dbData, container, dbSheetName);
                return;
            }
            
            // Display with recalculated values
            displayDetailedTableWithRecalculation(dbData, container, dbSheetName, calculation);
        })
        .catch(error => {
            console.error('Error loading Excel file:', error);
            container.innerHTML = `
                <div class="error-message">
                    Error loading detailed data: ${error.message}
                    <br><small>Make sure e2.xlsx is in the same directory</small>
                </div>
            `;
        });
}

// Load normal sheet (non-SMDB or SMDB without special DBs)
function loadNormalSheet(itemName, container, workbook) {
    const sheetNames = workbook.SheetNames;
    let sheetName = null;
    
    // Normalize MDB sheet names
    let searchName = itemName;
    if (itemName === 'MDB') {
        searchName = 'MDB1'; // MDB sheet should be treated as MDB1
    } else if (itemName === 'MDB.GF.04') {
        searchName = 'MDB4'; // MDB.GF.04 should be treated as MDB4
    }
    
    // Try exact match first (both original and normalized)
    if (sheetNames.includes(searchName)) {
        sheetName = searchName;
    } else if (sheetNames.includes(itemName)) {
        sheetName = itemName;
    } else {
        // Try to find sheet that contains the item name
        sheetName = sheetNames.find(name => 
            name.toLowerCase().includes(searchName.toLowerCase()) ||
            searchName.toLowerCase().includes(name.toLowerCase()) ||
            name.toLowerCase().includes(itemName.toLowerCase()) ||
            itemName.toLowerCase().includes(name.toLowerCase())
        );
    }
    
    // If no match, try to find a sheet with similar pattern
    if (!sheetName) {
        const baseName = searchName.split('.')[0] || searchName;
        sheetName = sheetNames.find(name => 
            name.toLowerCase().includes(baseName.toLowerCase())
        );
    }
    
    // Default to first sheet if no match
    if (!sheetName && sheetNames.length > 0) {
        sheetName = sheetNames[0];
    }
    
    if (sheetName) {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        displayDetailedTable(jsonData, container, sheetName);
    } else {
        container.innerHTML = `
            <div class="error-message">
                No matching sheet found for "${itemName}". Available sheets: ${sheetNames.join(', ')}
            </div>
        `;
    }
}


// Calculate DB estimate based on formula
// Note: NO OF UNITS parameter is the same as NO OF ITEMS from TOTALLIST sheet
function calculateDBEstimate(dbData, noOfUnits) {
    if (!dbData || dbData.length === 0) return null;
    
    const headers = dbData[0] || [];
    const rows = dbData.slice(1);
    
    // Find column indices - use same logic as display function
    const itemColumnIndex = headers.findIndex(h => 
        h && (h.toString().toLowerCase().includes('item') || 
              h.toString().toLowerCase().includes('description'))
    );
    
    // Find AMOUNT column specifically (prioritize exact match, then partial)
    let amountColumnIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().trim() === 'amount'
    );
    
    // If exact match not found, try partial match
    if (amountColumnIndex === -1) {
        amountColumnIndex = headers.findIndex(h => 
            h && (h.toString().toLowerCase().includes('amount') && 
                  !h.toString().toLowerCase().includes('price'))
        );
    }
    
    // Fallback to cost/estimate if amount not found
    if (amountColumnIndex === -1) {
        amountColumnIndex = headers.findIndex(h => 
            h && (h.toString().toLowerCase().includes('cost') ||
                  h.toString().toLowerCase().includes('estimate'))
        );
    }
    
    console.log('calculateDBEstimate - Column indices:', { itemColumnIndex, amountColumnIndex });
    console.log('calculateDBEstimate - Headers:', headers);
    
    if (itemColumnIndex === -1 || amountColumnIndex === -1) {
        console.warn('Could not find required columns in DB sheet');
        return null;
    }
    
    let total = 0;
    let labour = 0;
    let sumAfterLabour = 0;
    
    // Find TOTAL and LABOUR rows - check all columns for the label, then use amountColumnIndex for value
    rows.forEach((row, idx) => {
        // Check all columns for TOTAL/LABOUR labels
        let foundLabel = '';
        let labelColumnIndex = -1;
        
        for (let colIdx = 0; colIdx < row.length; colIdx++) {
            const cellValue = row[colIdx];
            if (!cellValue) continue;
            
            const cellStr = cellValue.toString().toLowerCase().trim();
            
            // Check for TOTAL row - exact match, not "NET TOTAL" or "SUM AFTER TOTAL"
            if (cellStr === 'total' && !cellStr.includes('sum after') && !cellStr.includes('net')) {
                foundLabel = 'TOTAL';
                labelColumnIndex = colIdx;
                break;
            } 
            // Check for LABOUR row
            else if (cellStr === 'labour' || cellStr === 'labor') {
                foundLabel = 'LABOUR';
                labelColumnIndex = colIdx;
                break;
            } 
            // Check for SUM AFTER LABOUR
            else if (cellStr.includes('sum after labour') || cellStr.includes('sum after labor')) {
                foundLabel = 'SUM AFTER LABOUR';
                labelColumnIndex = colIdx;
                break;
            }
        }
        
        // If we found a label, get the amount from the amount column
        if (foundLabel && amountColumnIndex !== -1 && row[amountColumnIndex] !== undefined) {
            const amountValue = parseFloat(row[amountColumnIndex]) || 0;
            
            if (foundLabel === 'TOTAL') {
                total = amountValue;
                console.log(`Found TOTAL at row ${idx}:`, row[labelColumnIndex], 'Amount column value =', amountValue, 'Full row:', row);
            } else if (foundLabel === 'LABOUR') {
                labour = amountValue;
                console.log(`Found LABOUR at row ${idx}:`, row[labelColumnIndex], 'Amount column value =', amountValue, 'Full row:', row);
            } else if (foundLabel === 'SUM AFTER LABOUR') {
                sumAfterLabour = amountValue;
                console.log(`Found SUM AFTER LABOUR at row ${idx}:`, row[labelColumnIndex], 'Amount column value =', amountValue, 'Full row:', row);
            }
        }
    });
    
    console.log('Extracted values:', { total, labour, sumAfterLabour, noOfUnits });
    
    // Validate that we have required values
    if (isNaN(total) || total === 0) {
        console.error('TOTAL is missing or zero after search. Sample rows:', rows.slice(140, 160).map((r, i) => ({
            index: 140 + i,
            item: r[itemColumnIndex],
            amount: r[amountColumnIndex],
            fullRow: r
        })));
        return null;
    }
    
    if (isNaN(labour) || labour === 0) {
        console.error('LABOUR is missing or zero after search');
        return null;
    }
    
    // Ensure noOfUnits is valid
    if (!noOfUnits || noOfUnits <= 0 || isNaN(noOfUnits)) {
        console.warn('Invalid NO OF UNITS:', noOfUnits);
        return null;
    }
    
    // If sumAfterLabour not found, calculate it
    if (sumAfterLabour === 0 || isNaN(sumAfterLabour)) {
        sumAfterLabour = total + labour;
    }
    
    // Apply formula as per user requirements
    // SUM AFTER NUMBER OF UNITS = (TOTAL + LABOUR) * NO OF UNITS
    const sumAfterNumberOfUnits = (total + labour) * noOfUnits;
    
    // Validate calculation
    if (isNaN(sumAfterNumberOfUnits) || !isFinite(sumAfterNumberOfUnits)) {
        console.error('Error calculating SUM AFTER NUMBER OF UNITS:', { total, labour, noOfUnits });
        return null;
    }
    
    // OVER HEAD = SUM AFTER NUMBER OF UNITS * 50%
    const overhead = sumAfterNumberOfUnits * 0.50;
    
    // SUM AFTER OVERHEAD = OVER HEAD + SUM AFTER NUMBER OF UNITS
    const sumAfterOverhead = overhead + sumAfterNumberOfUnits;
    
    // TAX 5% = SUM AFTER OVERHEAD * 5%
    const tax5 = sumAfterOverhead * 0.05;
    
    // SUM AFTER TAX = SUM AFTER OVERHEAD + TAX 5%
    const sumAfterTax = sumAfterOverhead + tax5;
    
    // PROVISIONAL SUM 10% = SUM AFTER TAX * 10%
    const provisionalSum10 = sumAfterTax * 0.10;
    
    // NET TOTAL = SUM AFTER TAX + PROVISIONAL SUM 10%
    const netTotal = sumAfterTax + provisionalSum10;
    
    // Debug logging
    console.log('DB Calculation:', {
        total,
        labour,
        sumAfterLabour,
        noOfUnits,
        sumAfterNumberOfUnits,
        overhead,
        sumAfterOverhead,
        tax5,
        sumAfterTax,
        provisionalSum10,
        netTotal
    });
    
    return {
        total,
        labour,
        sumAfterLabour,
        noOfUnits,
        sumAfterNumberOfUnits,
        overhead,
        sumAfterOverhead,
        tax5,
        sumAfterTax,
        provisionalSum10,
        netTotal
    };
}

// Display detailed table with recalculated values for special DBs
function displayDetailedTableWithRecalculation(data, container, sheetName, calculation) {
    if (!data || data.length === 0) {
        container.innerHTML = '<div class="error-message">No data found in sheet "' + sheetName + '"</div>';
        return;
    }
    
    // First row as headers
    const headers = data[0] || [];
    const rows = data.slice(1);
    
    // Find column indices
    const itemColumnIndex = headers.findIndex(h => 
        h && (h.toString().toLowerCase().includes('item') || 
              h.toString().toLowerCase().includes('description'))
    );
    
    // Find AMOUNT column specifically (prioritize exact match, then partial)
    let amountColumnIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().trim() === 'amount'
    );
    
    // If exact match not found, try partial match
    if (amountColumnIndex === -1) {
        amountColumnIndex = headers.findIndex(h => 
            h && (h.toString().toLowerCase().includes('amount') && 
                  !h.toString().toLowerCase().includes('price'))
        );
    }
    
    // Fallback to cost/estimate if amount not found
    if (amountColumnIndex === -1) {
        amountColumnIndex = headers.findIndex(h => 
            h && (h.toString().toLowerCase().includes('cost') ||
                  h.toString().toLowerCase().includes('estimate'))
        );
    }
    
    console.log('Column indices:', { itemColumnIndex, amountColumnIndex });
    console.log('Headers:', headers);
    console.log('Calculation values:', calculation);
    
    if (itemColumnIndex === -1 || amountColumnIndex === -1) {
        container.innerHTML = '<div class="error-message">Could not find required columns in sheet</div>';
        return;
    }
    
    // List of calculation-related row names that should always be shown
    const calculationRows = [
        'no of units',
        'number of units',
        'sum after number of units',
        'sum after no of units',
        'over head',
        'overhead',
        'sum after overhead',
        'tax',
        'tax 5%',
        'sum after tax',
        'provisional sum',
        'provisional sum 10%',
        'net total'
    ];
    
    // Debug: Log all rows to see what we're working with
    console.log('Sample rows for matching:', rows.slice(0, 30).map((row, idx) => ({
        index: idx,
        item: row[itemColumnIndex],
        amount: row[amountColumnIndex],
        fullRow: row
    })));
    
    // Find rows to modify and ensure calculation rows are included
    const modifiedRows = rows.map((row, rowIndex) => {
        const itemValue = row[itemColumnIndex];
        if (!itemValue) return row;
        
        const itemStr = itemValue.toString().toLowerCase().trim();
        const newRow = [...row];
        let updated = false;
        let matchedLabel = '';
        
        // Update specific rows based on formula - use more flexible matching
        // Check for "NO OF UNITS" or "NUMBER OF UNITS" (but not "SUM AFTER NUMBER OF UNITS")
        if ((itemStr === 'no of units' || itemStr === 'number of units') || 
            (itemStr.includes('no') && itemStr.includes('units') && !itemStr.includes('sum after') && !itemStr.includes('number'))) {
            newRow[amountColumnIndex] = calculation.noOfUnits;
            updated = true;
            matchedLabel = 'NO OF UNITS';
            console.log(`✓ Updated NO OF UNITS row ${rowIndex}: "${itemValue}" -> ${calculation.noOfUnits}`);
        } 
        // Check for "SUM AFTER NUMBER OF UNITS"
        else if (itemStr.includes('sum after number of units') || 
                 itemStr.includes('sum after no of units') ||
                 (itemStr.includes('sum after') && itemStr.includes('number') && itemStr.includes('units'))) {
            newRow[amountColumnIndex] = calculation.sumAfterNumberOfUnits;
            updated = true;
            matchedLabel = 'SUM AFTER NUMBER OF UNITS';
            console.log(`✓ Updated SUM AFTER NUMBER OF UNITS row ${rowIndex}: "${itemValue}" -> ${calculation.sumAfterNumberOfUnits}`);
        } 
        // Check for "OVER HEAD" or "OVERHEAD"
        else if (itemStr === 'over head' || itemStr === 'overhead' || 
                 (itemStr.includes('over') && itemStr.includes('head') && !itemStr.includes('sum after'))) {
            newRow[amountColumnIndex] = calculation.overhead;
            updated = true;
            matchedLabel = 'OVER HEAD';
            console.log(`✓ Updated OVER HEAD row ${rowIndex}: "${itemValue}" -> ${calculation.overhead}`);
        } 
        // Check for "SUM AFTER OVERHEAD"
        else if (itemStr.includes('sum after overhead')) {
            newRow[amountColumnIndex] = calculation.sumAfterOverhead;
            updated = true;
            matchedLabel = 'SUM AFTER OVERHEAD';
            console.log(`✓ Updated SUM AFTER OVERHEAD row ${rowIndex}: "${itemValue}" -> ${calculation.sumAfterOverhead}`);
        } 
        // Check for "TAX 5%" (but not "SUM AFTER TAX")
        else if ((itemStr.includes('tax') && itemStr.includes('5%')) || 
                 (itemStr === 'tax 5%') || 
                 (itemStr === 'tax' && !itemStr.includes('sum after'))) {
            newRow[amountColumnIndex] = calculation.tax5;
            updated = true;
            matchedLabel = 'TAX 5%';
            console.log(`✓ Updated TAX 5% row ${rowIndex}: "${itemValue}" -> ${calculation.tax5}`);
        } 
        // Check for "SUM AFTER TAX"
        else if (itemStr.includes('sum after tax')) {
            newRow[amountColumnIndex] = calculation.sumAfterTax;
            updated = true;
            matchedLabel = 'SUM AFTER TAX';
            console.log(`✓ Updated SUM AFTER TAX row ${rowIndex}: "${itemValue}" -> ${calculation.sumAfterTax}`);
        } 
        // Check for "PROVISIONAL SUM 10%"
        else if ((itemStr.includes('provisional sum') && itemStr.includes('10%')) || 
                 itemStr.includes('provisional sum 10%') ||
                 (itemStr.includes('provisional') && itemStr.includes('sum') && !itemStr.includes('after'))) {
            newRow[amountColumnIndex] = calculation.provisionalSum10;
            updated = true;
            matchedLabel = 'PROVISIONAL SUM 10%';
            console.log(`✓ Updated PROVISIONAL SUM 10% row ${rowIndex}: "${itemValue}" -> ${calculation.provisionalSum10}`);
        } 
        // Check for "NET TOTAL"
        else if (itemStr.includes('net total')) {
            newRow[amountColumnIndex] = calculation.netTotal;
            updated = true;
            matchedLabel = 'NET TOTAL';
            console.log(`✓ Updated NET TOTAL row ${rowIndex}: "${itemValue}" -> ${calculation.netTotal}`);
        }
        
        if (updated) {
            console.log(`Row ${rowIndex} before:`, row[amountColumnIndex], `after:`, newRow[amountColumnIndex]);
        }
        
        return newRow;
    });
    
    // Count how many rows were updated
    const updatedCount = modifiedRows.filter((row, idx) => {
        const itemValue = row[itemColumnIndex];
        if (!itemValue) return false;
        const itemStr = itemValue.toString().toLowerCase().trim();
        return itemStr.includes('no of units') || 
               itemStr.includes('sum after number') ||
               itemStr.includes('over head') ||
               itemStr.includes('sum after overhead') ||
               (itemStr.includes('tax') && !itemStr.includes('sum after')) ||
               itemStr.includes('sum after tax') ||
               itemStr.includes('provisional') ||
               itemStr.includes('net total');
    }).length;
    
    console.log(`Total calculation rows found and updated: ${updatedCount}`);
    
    // Display with modified rows, but use a flag to preserve calculation rows
    displayDetailedTable(modifiedRows, container, sheetName, headers, true);
    
    // Update metadata with NET TOTAL and NO OF UNITS from calculation
    if (calculation && (calculation.netTotal !== undefined || calculation.noOfUnits !== undefined)) {
        const metadataContent = document.getElementById('metadata-content');
        if (metadataContent) {
            const metadataTable = metadataContent.querySelector('.metadata-table');
            if (metadataTable) {
                // Update Estimate with NET TOTAL
                if (calculation.netTotal !== undefined && calculation.netTotal !== null) {
                    const estimateRow = Array.from(metadataTable.querySelectorAll('tr')).find(tr => {
                        const firstCell = tr.querySelector('td');
                        return firstCell && firstCell.textContent.trim() === 'Estimate';
                    });
                    if (estimateRow) {
                        const estimateCell = estimateRow.querySelectorAll('td')[1];
                        if (estimateCell) {
                            estimateCell.textContent = formatNumber(calculation.netTotal);
                            console.log('Updated metadata Estimate from calculation to:', calculation.netTotal);
                        }
                    }
                }
                
                // Update Number of Items with NO OF UNITS
                if (calculation.noOfUnits !== undefined && calculation.noOfUnits !== null) {
                    const noOfItemsRow = Array.from(metadataTable.querySelectorAll('tr')).find(tr => {
                        const firstCell = tr.querySelector('td');
                        return firstCell && (firstCell.textContent.trim() === 'Number of Items' || 
                                           firstCell.textContent.trim() === 'NO OF ITEMS');
                    });
                    if (noOfItemsRow) {
                        const noOfItemsCell = noOfItemsRow.querySelectorAll('td')[1];
                        if (noOfItemsCell) {
                            noOfItemsCell.textContent = calculation.noOfUnits.toString();
                            console.log('Updated metadata Number of Items from calculation to:', calculation.noOfUnits);
                        }
                    }
                }
            }
        }
    }
}

// Display detailed data as table
function displayDetailedTable(data, container, sheetName, providedHeaders = null, preserveCalculationRows = false) {
    if (!data || data.length === 0) {
        container.innerHTML = '<div class="error-message">No data found in sheet "' + sheetName + '"</div>';
        return;
    }
    
    // First row as headers (or use provided headers)
    const headers = providedHeaders || data[0] || [];
    const rows = providedHeaders ? data : data.slice(1);
    
    // Find the "Amount" column index (case-insensitive)
    const amountColumnIndex = headers.findIndex(h => 
        h && (h.toString().toLowerCase().includes('amount') || 
              h.toString().toLowerCase().includes('estimate') ||
              h.toString().toLowerCase().includes('cost'))
    );
    
    // Find item/description column index for checking calculation rows
    const itemColumnIndex = headers.findIndex(h => 
        h && (h.toString().toLowerCase().includes('item') || 
              h.toString().toLowerCase().includes('description'))
    );
    
    // Find BRAND column index
    const brandColumnIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().includes('brand')
    );
    
    // Find PRICE column index
    const priceColumnIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().includes('price')
    );
    
    // Find price and amount column indices for formatting
    const priceAmountColumnIndices = headers.map((h, index) => {
        if (!h) return null;
        const headerLower = h.toString().toLowerCase();
        if (headerLower.includes('price') || 
            headerLower.includes('amount') || 
            headerLower.includes('cost') ||
            headerLower.includes('estimate') ||
            headerLower.includes('total')) {
            return index;
        }
        return null;
    }).filter(index => index !== null);
    
    // List of calculation-related row names that should always be shown
    const calculationRowKeywords = [
        'no of units',
        'number of units',
        'sum after number of units',
        'sum after no of units',
        'over head',
        'overhead',
        'sum after overhead',
        'tax 5%',
        'tax',
        'sum after tax',
        'provisional sum',
        'provisional sum 10%',
        'net total'
    ];
    
    // Filter rows: only show rows where amount column is not zero
    // BUT preserve calculation rows if preserveCalculationRows flag is set
    const filteredRows = rows.filter(row => {
        if (amountColumnIndex === -1) {
            // If no amount column found, show all rows
            return true;
        }
        
        // Check if this is a calculation row that should be preserved
        if (preserveCalculationRows && itemColumnIndex !== -1) {
            const itemValue = row[itemColumnIndex];
            if (itemValue) {
                const itemStr = itemValue.toString().toLowerCase().trim();
                const isCalculationRow = calculationRowKeywords.some(keyword => 
                    itemStr.includes(keyword)
                );
                if (isCalculationRow) {
                    return true; // Always show calculation rows
                }
            }
        }
        
        const amountValue = row[amountColumnIndex];
        // Check if amount is not zero (handle various formats: 0, "0", 0.0, null, undefined, empty string)
        if (amountValue === null || amountValue === undefined || amountValue === '') {
            return false;
        }
        const numValue = parseFloat(amountValue);
        return !isNaN(numValue) && numValue !== 0;
    });
    
    if (filteredRows.length === 0) {
        container.innerHTML = '<div class="error-message">No rows with non-zero amounts found in sheet "' + sheetName + '"</div>';
        return;
    }
    
    // Helper function to format numeric values
    const formatNumericValue = (value, columnIndex) => {
        if (priceAmountColumnIndices.includes(columnIndex)) {
            // Try to parse as number
            const numValue = parseFloat(value);
            if (!isNaN(numValue) && value !== null && value !== undefined && value !== '') {
                return numValue.toFixed(2);
            }
        }
        return value;
    };
    
    // Find indices of rows that should be highlighted (between "Total" and "Net Total")
    // Check all columns, not just the first one
    let highlightStartIndex = -1;
    let highlightEndIndex = -1;
    
    filteredRows.forEach((row, rowIndex) => {
        // Check all cells in the row for "TOTAL" or "NET TOTAL"
        for (let colIndex = 0; colIndex < row.length; colIndex++) {
            const cellValue = row[colIndex];
            if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
                const cellValueStr = cellValue.toString().toLowerCase().trim();
                
                // Find "TOTAL" row - exact match or starts with "total" but not "net total" or "sum after"
                if (cellValueStr === 'total' || 
                    (cellValueStr.startsWith('total') && 
                     !cellValueStr.includes('net') && 
                     !cellValueStr.includes('sum after') &&
                     !cellValueStr.includes('after'))) {
                    if (highlightStartIndex === -1) {
                        highlightStartIndex = rowIndex;
                        console.log(`Found TOTAL at row ${rowIndex}:`, cellValue);
                        break; // Found it, move to next row
                    }
                }
                
                // Find "NET TOTAL" row
                if ((cellValueStr.includes('net') && cellValueStr.includes('total')) ||
                    cellValueStr === 'net total') {
                    highlightEndIndex = rowIndex;
                    console.log(`Found NET TOTAL at row ${rowIndex}:`, cellValue);
                    break; // Found it, move to next row
                }
            }
        }
    });
    
    // Debug logging
    console.log('=== HIGHLIGHT DEBUG ===');
    console.log('Highlight range:', highlightStartIndex, 'to', highlightEndIndex);
    console.log('Total filtered rows:', filteredRows.length);
    if (highlightStartIndex !== -1 && highlightEndIndex !== -1) {
        console.log('Will highlight rows:', highlightStartIndex, 'through', highlightEndIndex);
        filteredRows.slice(highlightStartIndex, highlightEndIndex + 1).forEach((row, idx) => {
            console.log(`Row ${highlightStartIndex + idx}:`, row);
        });
    } else {
        console.log('WARNING: Could not find TOTAL or NET TOTAL rows!');
        console.log('First 5 rows:', filteredRows.slice(0, 5).map(r => r[0]));
    }
    
    let html = `<p><strong>Sheet:</strong> ${sheetName}</p>`;
    html += '<table class="detailed-data-table"><thead><tr>';
    
    headers.forEach(header => {
        html += `<th>${header || ''}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    filteredRows.forEach((row, rowIndex) => {
        // Determine if this row should be highlighted
        const shouldHighlight = highlightStartIndex !== -1 && 
                               highlightEndIndex !== -1 && 
                               rowIndex >= highlightStartIndex && 
                               rowIndex <= highlightEndIndex;
        
        if (shouldHighlight) {
            console.log(`Highlighting row ${rowIndex}:`, row[0]);
        }
        
        const rowClass = shouldHighlight ? 'highlight-row' : '';
        const bgColor = shouldHighlight ? '#ffeb3b' : '';
        const rowStyle = shouldHighlight ? ` style="background-color: ${bgColor} !important; font-weight: 600 !important;"` : '';
        html += `<tr class="${rowClass}"${rowStyle}>`;
        headers.forEach((_, index) => {
            let value = row[index] !== undefined ? row[index] : '';
            
            // Correct typo: "SUM AFTER LABOUT" -> "SUM AFTER LABOUR"
            if (typeof value === 'string') {
                value = value.replace(/SUM AFTER LABOUT/gi, 'SUM AFTER LABOUR');
            }
            
            // Hide '0' values in BRAND column
            if (index === brandColumnIndex) {
                if (value === '0' || value === 0 || value === '0.00' || value === null || value === undefined || value === '') {
                    value = ''; // Hide the '0' value
                }
            }
            
            // Hide '0' values in PRICE column
            if (index === priceColumnIndex) {
                const priceNum = value !== null && value !== undefined && value !== '' ? parseFloat(value) : 0;
                if (isNaN(priceNum) || priceNum === 0) {
                    value = ''; // Hide the '0' value
                }
            }
            
            const formattedValue = formatNumericValue(value, index);
            const cellBgColor = shouldHighlight ? '#ffeb3b' : '';
            const cellBorder = shouldHighlight ? '2px solid #ff9800' : '';
            const cellStyle = shouldHighlight ? ` style="background-color: ${cellBgColor} !important; border: ${cellBorder} !important; font-weight: 600 !important;"` : '';
            html += `<td${cellStyle}>${formattedValue}</td>`;
        });
        html += '</tr>';
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
    
    // Extract NET TOTAL and NO OF UNITS from the table and update metadata
    let netTotal = null;
    let noOfUnits = null;
    let sumAfterLabourIndex = -1;
    
    // First pass: find SUM AFTER LABOUR row index
    filteredRows.forEach((row, rowIndex) => {
        const itemValue = itemColumnIndex !== -1 ? row[itemColumnIndex] : '';
        if (itemValue) {
            const itemStr = itemValue.toString().toLowerCase().trim();
            if (itemStr.includes('sum after labour') || itemStr.includes('sum after labor')) {
                sumAfterLabourIndex = rowIndex;
            }
        }
    });
    
    // Second pass: extract NET TOTAL and NO OF UNITS
    filteredRows.forEach((row, rowIndex) => {
        const itemValue = itemColumnIndex !== -1 ? row[itemColumnIndex] : '';
        const amountValue = amountColumnIndex !== -1 ? row[amountColumnIndex] : null;
        
        if (itemValue) {
            const itemStr = itemValue.toString().toLowerCase().trim();
            
            // Find NET TOTAL
            if ((itemStr.includes('net') && itemStr.includes('total')) || itemStr === 'net total') {
                if (amountValue !== null && amountValue !== undefined && amountValue !== '') {
                    const parsed = parseFloat(amountValue);
                    if (!isNaN(parsed)) {
                        netTotal = parsed;
                        console.log('Extracted NET TOTAL from detailed sheet:', netTotal);
                    }
                }
            }
            
            // Find NO OF UNITS - should be the row immediately after SUM AFTER LABOUR
            if (sumAfterLabourIndex !== -1 && rowIndex === sumAfterLabourIndex + 1) {
                // Check if this row contains NO OF UNITS
                if (itemStr === 'no of units' || itemStr === 'number of units' || 
                    (itemStr.includes('no') && itemStr.includes('units') && !itemStr.includes('sum after'))) {
                    if (amountValue !== null && amountValue !== undefined && amountValue !== '') {
                        const parsed = parseFloat(amountValue);
                        if (!isNaN(parsed)) {
                            noOfUnits = parsed;
                            console.log('Extracted NO OF UNITS from detailed sheet (after SUM AFTER LABOUR):', noOfUnits);
                        }
                    }
                }
            }
        }
    });
    
    // Update metadata if values were found
    if (netTotal !== null || noOfUnits !== null) {
        const metadataContent = document.getElementById('metadata-content');
        if (metadataContent) {
            const metadataTable = metadataContent.querySelector('.metadata-table');
            if (metadataTable) {
                // Update Estimate if NET TOTAL was found
                if (netTotal !== null) {
                    const estimateRow = Array.from(metadataTable.querySelectorAll('tr')).find(tr => {
                        const firstCell = tr.querySelector('td');
                        return firstCell && firstCell.textContent.trim() === 'Estimate';
                    });
                    if (estimateRow) {
                        const estimateCell = estimateRow.querySelectorAll('td')[1];
                        if (estimateCell) {
                            estimateCell.textContent = formatNumber(netTotal);
                            console.log('Updated metadata Estimate to:', netTotal);
                        }
                    }
                }
                
                // Update Number of Items if NO OF UNITS was found
                if (noOfUnits !== null) {
                    const noOfItemsRow = Array.from(metadataTable.querySelectorAll('tr')).find(tr => {
                        const firstCell = tr.querySelector('td');
                        return firstCell && (firstCell.textContent.trim() === 'Number of Items' || 
                                           firstCell.textContent.trim() === 'NO OF ITEMS');
                    });
                    if (noOfItemsRow) {
                        const noOfItemsCell = noOfItemsRow.querySelectorAll('td')[1];
                        if (noOfItemsCell) {
                            noOfItemsCell.textContent = noOfUnits.toString();
                            console.log('Updated metadata Number of Items to:', noOfUnits);
                        }
                    }
                }
            }
        }
    }
}

// Close detailed page and return to diagram
function closeDetailedPage() {
    const detailedPage = document.getElementById('detailed-page');
    const sourceView = detailedPage.getAttribute('data-source-view') || 'main';
    
    detailedPage.classList.add('hidden');
    
    // Return to the appropriate view based on source
    if (sourceView === 'mdb') {
        // Return to MDB view
        document.getElementById('mdb-view-page').classList.remove('hidden');
    } else {
        // Return to main content
        document.querySelector('main > .controls').style.display = 'block';
        document.querySelector('main > .legend-panel').style.display = 'block';
        document.querySelector('main > .info-panel').style.display = 'block';
        document.getElementById('main-content').style.display = 'flex';
    }
}

// Event listeners
document.getElementById('closeDetails').addEventListener('click', () => {
    document.getElementById('details-panel').classList.add('hidden');
});

document.getElementById('backToHome').addEventListener('click', () => {
    closeDetailedPage();
});

document.getElementById('searchBtn').addEventListener('click', searchItems);
document.getElementById('searchInput').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        searchItems();
    }
});

document.getElementById('kindFilter').addEventListener('change', filterByKind);

document.getElementById('resetView').addEventListener('click', () => {
    if (svg && zoom) {
        svg.transition().duration(750).call(
            zoom.transform,
            d3.zoomIdentity
        );
    }
});

document.getElementById('fitView').addEventListener('click', () => {
    if (svg && g) {
        const bounds = g.node().getBBox();
        const fullWidth = svg.node().getBoundingClientRect().width;
        const fullHeight = svg.node().getBoundingClientRect().height;
        const widthScale = (fullWidth - 40) / bounds.width;
        const heightScale = (fullHeight - 40) / bounds.height;
        const scale = Math.min(widthScale, heightScale, 1);
        
        const translate = [
            (fullWidth - bounds.width * scale) / 2 - bounds.x * scale,
            (fullHeight - bounds.height * scale) / 2 - bounds.y * scale
        ];
        
        svg.transition().duration(750).call(
            zoom.transform,
            d3.zoomIdentity.translate(translate[0], translate[1]).scale(scale)
        );
    }
});

document.getElementById('expandAll').addEventListener('click', () => {
    if (root) {
        root.each(d => {
            if (d._children) {
                d.children = d._children;
                d._children = null;
            }
        });
        initializeDiagram();
    }
});

document.getElementById('collapseAll').addEventListener('click', () => {
    if (root) {
        root.children.forEach(collapse);
        root.children.forEach(collapse);
        root.children.forEach(collapse);
        initializeDiagram();
    }
});

function collapse(d) {
    if (d.children) {
        d._children = d.children;
        d._children.forEach(collapse);
        d.children = null;
    }
}

document.getElementById('zoomIn').addEventListener('click', () => {
    if (svg && zoom) {
        svg.transition().call(zoom.scaleBy, 1.3);
    }
});

document.getElementById('zoomOut').addEventListener('click', () => {
    if (svg && zoom) {
        svg.transition().call(zoom.scaleBy, 0.77);
    }
});

document.getElementById('zoomReset').addEventListener('click', () => {
    if (svg && zoom) {
        svg.transition().duration(750).call(
            zoom.transform,
            d3.zoomIdentity
        );
    }
});

document.getElementById('toggleTreeView').addEventListener('click', () => {
    const sidebar = document.getElementById('tree-sidebar');
    sidebar.classList.toggle('hidden');
});

document.getElementById('closeTreeView').addEventListener('click', () => {
    document.getElementById('tree-sidebar').classList.add('hidden');
});

// Auto-sync is handled automatically, no refresh button needed

// MDB View button
document.getElementById('showMDBView').addEventListener('click', () => {
    openMDBView();
});

// Back button from MDB view
document.getElementById('backToHomeFromMDB').addEventListener('click', () => {
    closeMDBView();
});

// MDB Tab buttons
document.querySelectorAll('.mdb-tab').forEach(tab => {
    tab.addEventListener('click', () => {
        // Remove active class from all tabs
        document.querySelectorAll('.mdb-tab').forEach(t => t.classList.remove('active'));
        // Add active class to clicked tab
        tab.classList.add('active');
        // Load the corresponding MDB table
        const mdbName = tab.getAttribute('data-mdb');
        loadMDBTable(mdbName);
    });
});

function searchItems() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();
    if (!searchTerm) {
        resetFilter();
        return;
    }
    
    // Highlight matching nodes
    if (svg && g) {
        g.selectAll('g.node').each(function(d) {
            const node = d3.select(this);
            const name = (d.data.name || '').toLowerCase();
            if (name.includes(searchTerm)) {
                node.select('circle').attr('stroke-width', 4).attr('stroke', '#FFD700');
            } else {
                node.select('circle').attr('stroke-width', d.data.id === 'RMU' ? 3 : 2)
                    .attr('stroke', getNodeBorderColor(d.data.kind));
            }
        });
    }
}

function filterByKind() {
    const selectedKind = document.getElementById('kindFilter').value;
    if (!selectedKind) {
        resetFilter();
        return;
    }
    
    if (svg && g) {
        g.selectAll('g.node').each(function(d) {
            const node = d3.select(this);
            if (d.data.kind === selectedKind) {
                node.style('opacity', 1);
            } else {
                node.style('opacity', 0.2);
            }
        });
    }
}

function resetFilter() {
    if (svg && g) {
        g.selectAll('g.node').each(function(d) {
            const node = d3.select(this);
            node.style('opacity', 1);
            node.select('circle').attr('stroke-width', d.data.id === 'ROOT' ? 3 : 2)
                .attr('stroke', getNodeBorderColor(d.data.kind));
        });
    }
}

// Handle window resize
window.addEventListener('resize', () => {
    if (svg) {
        const container = d3.select('#mindmap');
        const width = container.node().offsetWidth || 1200;
        const height = Math.max(600, window.innerHeight * 0.6);
        svg.attr('width', width).attr('height', height);
    }
});

// MDB View Functions
function openMDBView() {
    document.getElementById('main-content').style.display = 'none';
    document.getElementById('detailed-page').classList.add('hidden');
    document.getElementById('mdb-view-page').classList.remove('hidden');
    
    // Load MDB1 by default
    loadMDBTable('MDB1');
}

function closeMDBView() {
    document.getElementById('mdb-view-page').classList.add('hidden');
    document.getElementById('main-content').style.display = 'block';
}

function loadMDBTable(mdbName) {
    if (!treeData) {
        console.error('Tree data not available');
        return;
    }
    
    // Find the MDB node
    const mdbNode = treeData.children.find(child => child.name === mdbName);
    if (!mdbNode) {
        document.getElementById('mdb-table-container').innerHTML = 
            `<div class="error-message">No data found for ${mdbName}</div>`;
        return;
    }
    
    // Show loading message
    document.getElementById('mdb-table-container').innerHTML = 
        '<div class="loading">Loading and calculating estimates...</div>';
    
    // Flatten the tree structure into a table
    const tableData = flattenTreeForTable(mdbNode);
    
    // Recalculate estimates for special DBs
    recalculateMDBTableEstimates(tableData).then(() => {
        // Generate HTML table with recalculated estimates
        const html = generateMDBTable(tableData, mdbName);
        document.getElementById('mdb-table-container').innerHTML = html;
        
        // Setup collapse functionality after HTML is inserted
        setTimeout(() => {
            setupMDBTableCollapse(mdbName);
        }, 100);
    }).catch(error => {
        console.error('Error recalculating estimates:', error);
        // Fallback to table without recalculation
        const html = generateMDBTable(tableData, mdbName);
        document.getElementById('mdb-table-container').innerHTML = html;
        
        // Setup collapse functionality after HTML is inserted
        setTimeout(() => {
            setupMDBTableCollapse(mdbName);
        }, 100);
    });
}

// Recalculate estimates for special DBs in MDB table
async function recalculateMDBTableEstimates(tableData) {
    const specialDBs = ['DB.TN.LXX.1B1.01', 'DB.TN.LXX.2B1.01', 'DB.TN.LXX.3B1.01', 'DB.TH.GF.01'];
    
    // Load Excel file once
    const response = await fetch(EXCEL_FILE_NAME);
    if (!response.ok) {
        console.warn('Could not load Excel file for recalculation');
        return;
    }
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    // Process each row to find special DBs and recalculate
    for (let i = 0; i < tableData.length; i++) {
        const row = tableData[i];
        const itemName = row.name || row.itemdrop || '';
        
        // Check if this is a special DB (exact match or group node)
        let dbName = null;
        let noOfUnits = 1;
        
        // Check for exact match
        if (specialDBs.includes(itemName)) {
            dbName = itemName;
        } 
        // Check for group node pattern: "DB.TN.LXX.1B1.01 (5 copies)"
        else {
            for (const dbType of specialDBs) {
                if (itemName.includes(dbType) && itemName.includes('copies')) {
                    dbName = dbType;
                    // Extract copy count from name like "DB.TN.LXX.1B1.01 (5 copies)"
                    const match = itemName.match(/\((\d+)\s*copies?\)/i);
                    if (match) {
                        noOfUnits = parseInt(match[1]) || 1;
                    }
                    break;
                }
            }
        }
        
        if (dbName) {
            // Use NO OF ITEMS from TOTALLIST sheet (same as NO OF UNITS)
            // First try to get from row data (from TOTALLIST)
            let noOfItems = row.noOfItems || row['NO OF ITEMS'] || row['NO OF UNITS'];
            
            // If not found in row, try to get from window.allData (TOTALLIST)
            if (!noOfItems || noOfItems === 0) {
                if (window.allData) {
                    const totallistItem = window.allData.find(d => 
                        (d.Itemdrop || d.MDB) === dbName || 
                        (d.Itemdrop || d.MDB) === itemName
                    );
                    if (totallistItem) {
                        noOfItems = totallistItem['NO OF ITEMS'] || totallistItem['NO OF UNITS'];
                    }
                }
            }
            
            // If still not found, try copy counts as fallback
            if (!noOfItems || noOfItems === 0) {
                const fedFrom = row.fedFrom || '';
                const smdbList = fedFrom.split('\n').map(p => p.trim()).filter(p => p && p !== 'RMU' && !p.startsWith('MDB'));
                if (smdbList.length > 0) {
                    const smdbName = smdbList[0];
                    const copyCounts = getDBCopyCounts(smdbName);
                    noOfItems = copyCounts && copyCounts[dbName] ? copyCounts[dbName] : 1;
                } else {
                    noOfItems = noOfUnits; // Use the value extracted from group name
                }
            }
            
            // Convert to number if it's a string
            const finalNoOfUnits = typeof noOfItems === 'string' ? parseFloat(noOfItems) || 1 : (noOfItems || 1);
            
            // Use the extracted noOfUnits from group name if it's greater than 1, otherwise use noOfItems
            noOfUnits = (noOfUnits > 1) ? noOfUnits : finalNoOfUnits;
            
            if (noOfUnits > 0) {
                // Calculate estimate for this DB using NO OF ITEMS value
                const calculation = await calculateDBEstimateFromSheet(workbook, dbName, noOfUnits);
                
                if (calculation && calculation.netTotal) {
                    // Update the estimate with recalculated NET TOTAL
                    row.estimate = calculation.netTotal;
                    row.recalculated = true;
                    console.log(`Recalculated ${itemName} (${dbName}, ${noOfUnits} units from NO OF ITEMS) estimate: ${calculation.netTotal}`);
                }
            }
        }
    }
}

// Calculate DB estimate from Excel workbook
async function calculateDBEstimateFromSheet(workbook, dbName, noOfUnits) {
    const sheetNames = workbook.SheetNames;
    
    // Find DB sheet
    let dbSheetName = sheetNames.find(name => 
        name === dbName || 
        name.toLowerCase() === dbName.toLowerCase()
    );
    
    if (!dbSheetName) {
        dbSheetName = sheetNames.find(name => 
            name.toLowerCase().includes(dbName.toLowerCase()) ||
            dbName.toLowerCase().includes(name.toLowerCase())
        );
    }
    
    if (!dbSheetName) {
        console.warn(`Sheet not found for ${dbName}`);
        return null;
    }
    
    const dbWorksheet = workbook.Sheets[dbSheetName];
    const dbData = XLSX.utils.sheet_to_json(dbWorksheet, { header: 1 });
    
    // Calculate estimate for this DB
    return calculateDBEstimate(dbData, noOfUnits);
}

function flattenTreeForTable(node, level = 0, parentName = '') {
    const rows = [];
    
    // Add current node - ensure all properties are populated from data if available
    const nodeData = node.data || {};
    
    // Normalize the name - remove "MDB" from MDB1, "MDB.GF.04" from MDB4
    let normalizedName = node.name;
    if (normalizedName === 'MDB' && (parentName === 'MDB1' || parentName.startsWith('MDB1'))) {
        normalizedName = 'MDB1';
    } else if ((normalizedName === 'MDB.GF.04' || normalizedName === 'MDB GF 04') && (parentName === 'MDB4' || parentName.startsWith('MDB4'))) {
        normalizedName = 'MDB4';
    } else {
        // Apply general normalization
        normalizedName = normalizeMDB(normalizedName);
    }
    
    // Normalize itemdrop as well
    let normalizedItemdrop = nodeData.Itemdrop || node.name || '';
    normalizedItemdrop = normalizeMDB(normalizedItemdrop);
    
    const row = {
        level: level,
        name: normalizedName,
        kind: node.kind || nodeData.KIND || 'Unknown',
        load: node.load || nodeData.Load || '0 kW',
        estimate: node.estimate !== undefined ? node.estimate : (nodeData.Estimate || 0),
        fedFrom: parentName || nodeData['FED FROM'] || '',
        noOfItems: nodeData['NO OF ITEMS'] !== undefined ? nodeData['NO OF ITEMS'] : '',
        itemdrop: normalizedItemdrop,
        mdb: normalizeMDB(node.mdb || nodeData.MDB || ''),
        data: nodeData
    };
    
    // Filter out duplicate entries: if normalized name matches parent MDB name, skip it
    // (e.g., if parent is MDB1 and this node is also MDB1 after normalization, skip to avoid duplicate)
    if (level > 0 && normalizedName === parentName && node.kind === 'MDB') {
        // Skip this duplicate MDB entry
    } else {
        rows.push(row);
    }
    
    // Process children - use normalized name as parent name
    if (node.children && node.children.length > 0) {
        node.children.forEach(child => {
            const childRows = flattenTreeForTable(child, level + 1, normalizedName);
            rows.push(...childRows);
        });
    }
    
    return rows;
}

function generateMDBTable(data, mdbName) {
    if (!data || data.length === 0) {
        return '<div class="error-message">No data available</div>';
    }
    
    // Group data by parent to enable collapsing
    const groupedData = groupDataByParent(data);
    
    let html = `<h3 style="margin-bottom: 20px; color: #667eea;">${mdbName} - ${data.length} items</h3>`;
    html += '<table class="mdb-data-table" id="mdb-table-' + mdbName + '">';
    html += '<thead><tr>';
    html += '<th>Name</th>';
    html += '<th>Kind</th>';
    html += '<th>Load</th>';
    html += '<th>Load Validation</th>';
    html += '<th>No. of Items</th>';
    html += '<th>Fed From</th>';
    html += '<th>Estimate (AED)</th>';
    html += '</tr></thead><tbody>';
    
    // Generate rows with collapsible functionality
    groupedData.forEach((group, groupIndex) => {
        const parentRow = group.parent;
        const children = group.children;
        const hasChildren = children.length > 0;
        const indent = '&nbsp;'.repeat(parentRow.level * 4);
        const kindClass = `kind-${(parentRow.kind || '').toLowerCase().replace(/\s+/g, '-')}`;
        const estimate = parseFloat(parentRow.estimate) || 0;
        const isRecalculated = parentRow.recalculated || false;
        const isParentType = parentRow.kind && (
            parentRow.kind.toLowerCase().includes('smdb') || 
            parentRow.kind.toLowerCase().includes('esmdb') || 
            parentRow.kind.toLowerCase().includes('mcc')
        );
        const rowId = `mdb-row-${mdbName}-${groupIndex}`;
        const collapseId = `collapse-${mdbName}-${groupIndex}`;
        
        // Parent row (SMDB, ESMDB, MCC or other parent)
        // Ensure data is properly serialized - use itemdrop or name to find original data if node.data is empty
        let rowData = parentRow.data || {};
        if (!rowData || Object.keys(rowData).length === 0) {
            // Try to find data from allData if available
            const itemName = parentRow.name || parentRow.itemdrop || '';
            if (itemName && window.allData) {
                // Try exact match first
                let foundData = window.allData.find(d => {
                    const itemdrop = (d.Itemdrop || '').toString().trim();
                    const mdb = (d.MDB || '').toString().trim();
                    return itemdrop === itemName || mdb === itemName;
                });
                
                // If not found and itemName is MDB1, try finding "MDB" entry
                if (!foundData && itemName === 'MDB1') {
                    foundData = window.allData.find(d => {
                        const itemdrop = (d.Itemdrop || '').toString().trim();
                        const mdb = (d.MDB || '').toString().trim();
                        return (itemdrop === 'MDB' || mdb === 'MDB') && 
                               (d.KIND || '').toString().toUpperCase() === 'MDB';
                    });
                }
                
                // If not found and itemName is MDB4, try finding "MDB.GF.04" entry
                if (!foundData && itemName === 'MDB4') {
                    foundData = window.allData.find(d => {
                        const itemdrop = (d.Itemdrop || '').toString().trim();
                        const mdb = (d.MDB || '').toString().trim();
                        return (itemdrop === 'MDB.GF.04' || mdb === 'MDB.GF.04' || 
                                itemdrop === 'MDB GF 04' || mdb === 'MDB GF 04') &&
                               (d.KIND || '').toString().toUpperCase() === 'MDB';
                    });
                }
                
                if (foundData) {
                    rowData = foundData;
                }
            }
        }
        
        // For MDB nodes (MDB1, MDB2, MDB3, MDB4), make them clickable even if estimate is 0
        // They should be clickable because they have detailed sheets
        const isMDBNode = parentRow.kind === 'MDB' && ['MDB1', 'MDB2', 'MDB3', 'MDB4'].includes(parentRow.name);
        const isClickable = estimate > 0 || isMDBNode;
        const clickableClass = isClickable ? 'clickable-row' : 'non-clickable-row';
        
        // Calculate load validation for board nodes (MDB, SMDB, ESMDB, MCC, etc.) - not for DB level
        let loadValidationHtml = '<td>-</td>';
        const kind = (parentRow.kind || '').toString().toUpperCase();
        const isBoardType = kind === 'MDB' || 
                           kind === 'SMDB' || 
                           kind === 'ESMDB' || 
                           kind === 'MCC' ||
                           kind === 'BUS BAR RAISER' ||
                           kind === 'EMDB' ||
                           kind === 'EMCC';
        
        if (isBoardType) {
            // Get board's base load and number of units
            const boardBaseLoad = parseLoadValue(parentRow.load || '0 kW');
            const boardNoOfUnits = parseFloat(parentRow.noOfItems || parentRow['NO OF ITEMS'] || parentRow['NO OF UNITS'] || 1);
            // Board's effective load = base load × number of units
            const boardLoad = boardBaseLoad * boardNoOfUnits;
            
            // Placeholder for async load validation - will be updated after calculation
            const validationCellId = `load-validation-${mdbName}-${groupIndex}`;
            loadValidationHtml = `<td id="${validationCellId}" style="color: #999; font-size: 0.9em;">Calculating...</td>`;
            
            // Calculate children load sum asynchronously (will load detailed sheets for DBs)
            calculateChildrenLoadSum(parentRow.name).then(childrenLoadSum => {
                const validationCell = document.getElementById(validationCellId);
                if (!validationCell) return;
                
                if (boardLoad > 0 && childrenLoadSum > 0) {
                    const difference = Math.abs(boardLoad - childrenLoadSum);
                    const percentageDiff = (difference / boardLoad) * 100;
                    const tolerancePercent = 5; // ±5% tolerance
                    
                    if (percentageDiff > tolerancePercent) {
                        // Load mismatch detected (>5% difference)
                        const statusColor = '#f44336'; // Red for >5% difference
                        validationCell.innerHTML = `⚠ ${boardLoad.toFixed(2)} vs ${childrenLoadSum.toFixed(2)} kW<br>
                            <span style="font-size: 0.85em;">(${boardBaseLoad.toFixed(2)}×${boardNoOfUnits}) Diff: ${difference.toFixed(2)} kW (${percentageDiff.toFixed(2)}%)</span>`;
                        validationCell.style.backgroundColor = `${statusColor}20`;
                        validationCell.style.color = statusColor;
                        validationCell.style.fontWeight = '600';
                    } else {
                        // Load matches (within ±5% tolerance)
                        validationCell.innerHTML = `✓ ${boardLoad.toFixed(2)} ≈ ${childrenLoadSum.toFixed(2)} kW<br>
                            <span style="font-size: 0.85em;">(${boardBaseLoad.toFixed(2)}×${boardNoOfUnits}, ${percentageDiff.toFixed(2)}% diff)</span>`;
                        validationCell.style.backgroundColor = '#4CAF5020';
                        validationCell.style.color = '#4CAF50';
                        validationCell.style.fontWeight = '600';
                    }
                } else if (boardLoad > 0) {
                    // Board has load but no children to compare
                    validationCell.innerHTML = 'No children';
                    validationCell.style.color = '#999';
                } else {
                    validationCell.innerHTML = '-';
                }
            }).catch(error => {
                console.error(`Error calculating children load sum for ${parentRow.name}:`, error);
                const validationCell = document.getElementById(validationCellId);
                if (validationCell) {
                    validationCell.innerHTML = 'Error';
                    validationCell.style.color = '#999';
                }
            });
        }
        
        html += `<tr class="mdb-parent-row ${hasChildren && isParentType ? 'has-children' : ''} ${clickableClass}" data-row-id="${rowId}" data-collapse-id="${collapseId}" data-item-name="${parentRow.name || ''}" data-item-data='${JSON.stringify(rowData)}' data-estimate="${estimate}">`;
        html += `<td>${indent}`;
        if (hasChildren && isParentType) {
            html += `<span class="collapse-toggle" style="cursor: pointer; margin-right: 5px; user-select: none;">▶</span>`;
        }
        html += `${parentRow.name || ''}${isRecalculated ? ' <span style="color: #4CAF50; font-size: 0.85em;">(recalculated)</span>' : ''}</td>`;
        html += `<td><span class="${kindClass}">${parentRow.kind || ''}</span></td>`;
        html += `<td>${parentRow.load || '0 kW'}</td>`;
        html += loadValidationHtml;
        html += `<td>${parentRow.noOfItems || ''}</td>`;
        html += `<td>${parentRow.fedFrom || ''}</td>`;
        html += `<td class="numeric"${isRecalculated ? ' style="font-weight: 600; color: #4CAF50;"' : ''}>${formatNumber(estimate)}</td>`;
        html += '</tr>';
        
        // Child rows - initially hidden if parent is SMDB, ESMDB, or MCC
        if (hasChildren && isParentType) {
            children.forEach((childRow, childIndex) => {
                const childIndent = '&nbsp;'.repeat((childRow.level - parentRow.level) * 4 + 8); // Extra indent for children
                const childKindClass = `kind-${(childRow.kind || '').toLowerCase().replace(/\s+/g, '-')}`;
                const childEstimate = parseFloat(childRow.estimate) || 0;
                const childIsRecalculated = childRow.recalculated || false;
                const childRowClass = `mdb-child-row ${collapseId}`;
                const childIsClickable = childEstimate > 0;
                const childClickableClass = childIsClickable ? 'clickable-row' : 'non-clickable-row';
                // Ensure data is properly serialized - use itemdrop or name to find original data if node.data is empty
                let childRowData = childRow.data || {};
                if (!childRowData || Object.keys(childRowData).length === 0) {
                    // Try to find data from allData if available
                    const childItemName = childRow.name || childRow.itemdrop || '';
                    if (childItemName && window.allData) {
                        const foundData = window.allData.find(d => (d.Itemdrop || d.MDB) === childItemName);
                        if (foundData) {
                            childRowData = foundData;
                        }
                    }
                }
                html += `<tr class="${childRowClass} ${childClickableClass}" style="display: none;" data-item-name="${childRow.name || ''}" data-item-data='${JSON.stringify(childRowData)}' data-estimate="${childEstimate}">`;
                html += `<td>${childIndent}${childRow.name || ''}${childIsRecalculated ? ' <span style="color: #4CAF50; font-size: 0.85em;">(recalculated)</span>' : ''}</td>`;
                html += `<td><span class="${childKindClass}">${childRow.kind || ''}</span></td>`;
                html += `<td>${childRow.load || '0 kW'}</td>`;
                html += `<td>-</td>`; // Load validation not applicable for child rows
                html += `<td>${childRow.noOfItems || ''}</td>`;
                html += `<td>${childRow.fedFrom || ''}</td>`;
                html += `<td class="numeric"${childIsRecalculated ? ' style="font-weight: 600; color: #4CAF50;"' : ''}>${formatNumber(childEstimate)}</td>`;
                html += '</tr>';
            });
        }
    });
    
    html += '</tbody></table>';
    
    // Calculate totals - sum of MDB and all its descendants
    // The data array already includes the MDB node (first row) and all its children
    const totalEstimate = data.reduce((sum, row) => {
        const estimate = parseFloat(row.estimate) || 0;
        return sum + estimate;
    }, 0);
    
    const totalLoad = data.reduce((sum, row) => {
        const loadStr = row.load || '0 kW';
        const loadValue = parseFloat(loadStr.replace(/[^\d.]/g, '')) || 0;
        return sum + loadValue;
    }, 0);
    
    // Display totals prominently
    html += '<div style="margin-top: 20px; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 8px; color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">';
    html += `<div style="font-size: 1.1em; margin-bottom: 10px;"><strong>${mdbName} Summary</strong></div>`;
    html += `<div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px;">`;
    html += `<div><strong>Total Load:</strong> ${totalLoad.toFixed(2)} kW</div>`;
    html += `<div style="font-size: 1.2em; font-weight: bold;"><strong>Total Estimate:</strong> ${formatNumber(totalEstimate)} AED</div>`;
    html += `</div>`;
    html += `<div style="margin-top: 10px; font-size: 0.9em; opacity: 0.9;">Includes ${mdbName} and all boards/equipment under it</div>`;
    html += '</div>';
    
    return html;
}

// Setup collapse functionality for MDB table
function setupMDBTableCollapse(mdbName) {
    const table = document.getElementById('mdb-table-' + mdbName);
    if (!table) {
        console.warn('Table not found:', 'mdb-table-' + mdbName);
        return;
    }
    
    const parentRows = table.querySelectorAll('.mdb-parent-row.has-children');
    console.log('Found parent rows with children:', parentRows.length);
    
    parentRows.forEach((parentRow, index) => {
        const collapseId = parentRow.getAttribute('data-collapse-id');
        const toggle = parentRow.querySelector('.collapse-toggle');
        
        if (!collapseId || !toggle) {
            console.warn('Missing collapseId or toggle for row', index);
            return;
        }
        
        // Query child rows using the collapseId class
        const childRows = Array.from(table.querySelectorAll('tr')).filter(row => 
            row.classList.contains('mdb-child-row') && row.classList.contains(collapseId)
        );
        
        console.log(`Row ${index}: collapseId=${collapseId}, childRows=${childRows.length}`);
        
        if (childRows.length > 0) {
            // Collapse handler - only for toggle icon
            const collapseHandler = function(e) {
                e.preventDefault();
                e.stopPropagation();
                
                const isCollapsed = childRows[0].style.display === 'none' || 
                                   childRows[0].style.display === '' ||
                                   window.getComputedStyle(childRows[0]).display === 'none';
                
                childRows.forEach(row => {
                    row.style.display = isCollapsed ? 'table-row' : 'none';
                });
                
                toggle.textContent = isCollapsed ? '▼' : '▶';
                
                console.log(`Toggled ${collapseId}: ${isCollapsed ? 'expanded' : 'collapsed'}`);
            };
            
            // Only attach collapse handler to toggle icon
            toggle.addEventListener('click', collapseHandler);
        }
        
        // Add click handler for opening detailed page (for all parent rows with non-zero estimates or MDB nodes)
        const itemName = parentRow.getAttribute('data-item-name');
        const itemDataStr = parentRow.getAttribute('data-item-data');
        const rowEstimate = parseFloat(parentRow.getAttribute('data-estimate')) || 0;
        const isMDBNode = itemName && ['MDB1', 'MDB2', 'MDB3', 'MDB4'].includes(itemName);
        
        if (itemName && itemDataStr && (rowEstimate > 0 || isMDBNode)) {
            parentRow.style.cursor = 'pointer';
            parentRow.setAttribute('data-click-handler', 'true');
            parentRow.addEventListener('click', function(e) {
                // Don't trigger if clicking on collapse toggle
                if (e.target.classList.contains('collapse-toggle') || 
                    e.target.closest('.collapse-toggle')) {
                    return;
                }
                
                try {
                    const itemData = JSON.parse(itemDataStr);
                    openDetailedPage(itemData, itemName, 'mdb');
                } catch (err) {
                    console.error('Error parsing item data:', err);
                    openDetailedPage(null, itemName, 'mdb');
                }
            });
        } else if (rowEstimate === 0) {
            // Make non-clickable rows visually distinct
            parentRow.style.cursor = 'not-allowed';
            parentRow.style.opacity = '0.6';
        }
    });
    
    // Add click handlers for ALL rows with non-zero estimates (including MDB, PFC, BusBarRaiser, etc.)
    const allRowsWithData = table.querySelectorAll('tr[data-item-name]');
    allRowsWithData.forEach(row => {
        const itemName = row.getAttribute('data-item-name');
        const itemDataStr = row.getAttribute('data-item-data');
        const rowEstimate = parseFloat(row.getAttribute('data-estimate')) || 0;
        
        // Skip if already has click handler (parent rows processed above)
        if (row.hasAttribute('data-click-handler')) {
            return;
        }
        
        // Only make clickable if estimate > 0 and has valid data, or if it's an MDB node
        const isMDBNode = itemName && ['MDB1', 'MDB2', 'MDB3', 'MDB4'].includes(itemName);
        if (itemName && itemDataStr && (rowEstimate > 0 || isMDBNode)) {
            row.style.cursor = 'pointer';
            row.setAttribute('data-click-handler', 'true');
            row.addEventListener('click', function(e) {
                // Don't trigger if clicking on collapse toggle
                if (e.target.classList.contains('collapse-toggle') || 
                    e.target.closest('.collapse-toggle')) {
                    return;
                }
                
                try {
                    const itemData = JSON.parse(itemDataStr);
                    console.log('Opening detailed page for:', itemName, 'Data:', itemData);
                    openDetailedPage(itemData, itemName, 'mdb');
                } catch (err) {
                    console.error('Error parsing item data:', err, 'Item:', itemName, 'Data string:', itemDataStr);
                    openDetailedPage(null, itemName, 'mdb');
                }
            });
        } else if (rowEstimate === 0) {
            // Make non-clickable rows visually distinct
            row.style.cursor = 'not-allowed';
            row.style.opacity = '0.6';
        } else {
            // Debug: log rows that should be clickable but aren't
            console.warn('Row not clickable:', {
                itemName,
                hasData: !!itemDataStr,
                estimate: rowEstimate,
                row: row
            });
        }
    });
}

// Group data by parent for collapsible display
function groupDataByParent(data) {
    const groups = [];
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const isParentType = row.kind && (
            row.kind.toLowerCase().includes('smdb') || 
            row.kind.toLowerCase().includes('esmdb') || 
            row.kind.toLowerCase().includes('mcc')
        );
        
        // Check if this is an SMDB, ESMDB, or MCC that might have children
        if (isParentType) {
            // Look ahead to find children that belong to this parent
            const children = [];
            for (let j = i + 1; j < data.length; j++) {
                const nextRow = data[j];
                
                // If we hit another parent type or item at same or lower level, stop
                const nextIsParentType = nextRow.kind && (
                    nextRow.kind.toLowerCase().includes('smdb') || 
                    nextRow.kind.toLowerCase().includes('esmdb') || 
                    nextRow.kind.toLowerCase().includes('mcc')
                );
                
                if (nextRow.level <= row.level || nextIsParentType) {
                    break;
                }
                
                // If it's a direct child (one level deeper), add it
                if (nextRow.level === row.level + 1) {
                    children.push(nextRow);
                }
            }
            
            groups.push({
                parent: row,
                children: children
            });
            
            // Skip the children rows in the main loop (they're already added)
            i += children.length;
        } 
        // Standalone row (not a parent type with children)
        else {
            groups.push({
                parent: row,
                children: []
            });
        }
    }
    
    return groups;
}

// Note: loadData() is now called from the authentication handler after successful login
// The original DOMContentLoaded listener has been moved to the authentication section

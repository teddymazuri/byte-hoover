// Byte Hoover - Professional Data Cleaning Tool
// Core Application State
const AppState = {
    selectedFiles: [],
    currentFileIndex: 0,
    originalData: null,
    cleanedData: null,
    history: [],
    currentHistoryIndex: -1,
    processingMetrics: {
        startTime: null,
        endTime: null,
        rowsProcessed: 0,
        cellsProcessed: 0,
        issuesFixed: 0,
        columnsRemoved: 0,
        duplicatesRemoved: 0
    },
    configurations: [],
    currentConfig: null,
    isProcessing: false,
    isDarkMode: window.matchMedia('(prefers-color-scheme: dark)').matches,
    exportSettings: {
        format: 'xlsx',
        fileNamePrefix: 'cleaned_',
        compressOutput: false,
        preserveOriginalStructure: true
    }
};

// DOM Elements
const elements = {
    fileInput: document.getElementById('fileInput'),
    dropZone: document.getElementById('dropZone'),
    fileNameDisplay: document.getElementById('fileNameDisplay'),
    processBtn: document.getElementById('processBtn'),
    previewBtn: document.getElementById('previewBtn'),
    logArea: document.getElementById('logArea'),
    logBody: document.getElementById('logBody'),
    progressWrapper: document.getElementById('progressWrapper'),
    progressBar: document.getElementById('progressBar'),
    progressText: document.getElementById('progressText'),
    progressPercent: document.getElementById('progressPercent'),
    qualityScore: document.getElementById('qualityScore'),
    qualityScoreValue: document.getElementById('qualityScoreValue'),
    issuesFixed: document.getElementById('issuesFixed'),
    processTime: document.getElementById('processTime'),
    processingSpeed: document.getElementById('processingSpeed'),
    themeToggle: document.getElementById('themeToggle'),
    themeIcon: document.getElementById('themeIcon'),
    undoBtn: document.getElementById('undoBtn'),
    redoBtn: document.getElementById('redoBtn'),
    fileInfo: document.getElementById('fileInfo'),
    fileNameInfo: document.getElementById('fileNameInfo'),
    fileSizeInfo: document.getElementById('fileSizeInfo'),
    batchFiles: document.getElementById('batchFiles'),
    fileList: document.getElementById('fileList'),
    fileCount: document.getElementById('fileCount'),
    saveConfigBtn: document.getElementById('saveConfigBtn'),
    // Advanced Options Elements
    exportFormat: document.getElementById('exportFormat'),
    fileNamePrefix: document.getElementById('fileNamePrefix'),
    zipOutput: document.getElementById('zipOutput'),
    preserveStructure: document.getElementById('preserveStructure')
};

// Enhanced Data Patterns for Validation
const DataPatterns = {
    EMAIL: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,
    PHONE_US: /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/,
    PHONE_INTERNATIONAL: /^\+?[1-9]\d{1,14}$/,
    ZIP_US: /^\d{5}(-\d{4})?$/,
    DATE_ISO: /^\d{4}-\d{2}-\d{2}$/,
    DATE_US: /^\d{1,2}\/\d{1,2}\/\d{4}$/,
    DATE_EU: /^\d{1,2}\.\d{1,2}\.\d{4}$/,
    URL: /^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*\/?$/,
    SSN: /^\d{3}-\d{2}-\d{4}$/,
    CREDIT_CARD: /^\d{4}[\s\-]?\d{4}[\s\-]?\d{4}[\s\-]?\d{4}$/
};

// Preset Configurations
const Presets = {
    basic: {
        name: "Basic Clean",
        settings: {
            capitalize: true,
            punctuation: true,
            spaces: true,
            phone: false,
            dates: false,
            emails: false,
            empty: true,
            emptyCols: false,
            dupes: false,
            markInvalid: false,
            preserveNames: true,
            strictDates: false
        }
    },
    contact: {
        name: "Contact List",
        settings: {
            capitalize: true,
            punctuation: true,
            spaces: true,
            phone: true,
            dates: false,
            emails: true,
            empty: true,
            emptyCols: true,
            dupes: true,
            markInvalid: true,
            preserveNames: true,
            strictDates: false
        }
    },
    export: {
        name: "Export Ready",
        settings: {
            capitalize: true,
            punctuation: true,
            spaces: true,
            phone: true,
            dates: true,
            emails: true,
            empty: true,
            emptyCols: true,
            dupes: true,
            markInvalid: false,
            preserveNames: true,
            strictDates: true
        }
    },
    sensitive: {
        name: "Sensitive Data",
        settings: {
            capitalize: true,
            punctuation: true,
            spaces: true,
            phone: false,
            dates: false,
            emails: false,
            empty: true,
            emptyCols: true,
            dupes: false,
            markInvalid: false,
            preserveNames: true,
            strictDates: false,
            anonymize: true
        }
    },
    custom: {
        name: "Custom",
        settings: {}
    }
};

// Initialize Application
function init() {
    console.log('Byte Hoover v2.0.0 - Initializing...');
    
    // Load saved configurations
    loadConfigurations();
    
    // Load export settings
    loadExportSettings();
    
    // Check JSZip availability
    checkJSZipAvailability();
    
    // Apply dark mode if enabled
    if (AppState.isDarkMode) {
        document.body.setAttribute('data-bs-theme', 'dark');
        elements.themeIcon.className = 'bi bi-sun';
    }
    
    // Set up event listeners
    setupEventListeners();
    
    // Apply default preset
    applyPreset('basic');
    
    // Initialize tooltips
    initializeTooltips();
    
    // Add welcome log
    addLog('System', 'Byte Hoover v2.0 initialized with enhanced cleaning logic', 'info');
    
    console.log('Byte Hoover ready!');
}

// Initialize Bootstrap tooltips
function initializeTooltips() {
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.forEach(function (tooltipTriggerEl) {
        new bootstrap.Tooltip(tooltipTriggerEl);
    });
}

// Check JSZip Availability
function checkJSZipAvailability() {
    if (typeof JSZip === 'undefined') {
        console.warn('JSZip library not loaded - compression features will be disabled');
        
        // Disable compression checkbox if it exists
        if (elements.zipOutput) {
            elements.zipOutput.disabled = true;
            elements.zipOutput.checked = false;
            elements.zipOutput.parentElement.setAttribute('title', 'JSZip library required for compression');
            
            // Add tooltip for disabled state
            new bootstrap.Tooltip(elements.zipOutput.parentElement, {
                title: 'JSZip library not loaded. Add &lt;script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"&gt;&lt;/script&gt; to enable compression.',
                html: true
            });
        }
        
        addLog('System', 'JSZip library not loaded - compression disabled', 'warning');
    } else {
        console.log('JSZip library loaded - compression enabled');
        if (elements.zipOutput) {
            elements.zipOutput.disabled = false;
            elements.zipOutput.parentElement.removeAttribute('title');
        }
    }
}

// Event Listeners Setup
function setupEventListeners() {
    // File input change
    elements.fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelection(Array.from(e.target.files));
        }
    });
    
    // Drag and drop
    elements.dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        elements.dropZone.classList.add('drag-over');
    });
    
    elements.dropZone.addEventListener('dragleave', () => {
        elements.dropZone.classList.remove('drag-over');
    });
    
    elements.dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        elements.dropZone.classList.remove('drag-over');
        
        if (e.dataTransfer.files.length > 0) {
            handleFileSelection(Array.from(e.dataTransfer.files));
        }
    });
    
    // Process button
    elements.processBtn.addEventListener('click', processFiles);
    
    // Theme toggle
    elements.themeToggle.addEventListener('click', toggleTheme);
    
    // Preset buttons
    document.querySelectorAll('.preset-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const preset = e.target.dataset.preset;
            applyPreset(preset);
            
            // Update active state
            document.querySelectorAll('.preset-btn').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
        });
    });
    
    // Undo/Redo buttons
    elements.undoBtn.addEventListener('click', undo);
    elements.redoBtn.addEventListener('click', redo);
    
    // Save config button
    elements.saveConfigBtn.addEventListener('click', saveConfiguration);
    
    // Advanced options event listeners
    if (elements.exportFormat) {
        elements.exportFormat.addEventListener('change', saveExportSettings);
    }
    
    if (elements.fileNamePrefix) {
        elements.fileNamePrefix.addEventListener('change', saveExportSettings);
        elements.fileNamePrefix.addEventListener('input', saveExportSettings);
    }
    
    if (elements.zipOutput) {
        elements.zipOutput.addEventListener('change', (e) => {
            if (e.target.checked && typeof JSZip === 'undefined') {
                alert('JSZip library is not loaded. Compression requires JSZip. Please add the JSZip library to enable this feature.');
                e.target.checked = false;
            }
            saveExportSettings();
        });
    }
    
    if (elements.preserveStructure) {
        elements.preserveStructure.addEventListener('change', saveExportSettings);
    }
    
    // Keyboard shortcuts
    document.addEventListener('keydown', handleKeyboardShortcuts);
    
    // Handle URL parameters
    parseURLParameters();
}

// Handle Keyboard Shortcuts
function handleKeyboardShortcuts(e) {
    // Ctrl/Cmd + O to open file
    if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
        e.preventDefault();
        elements.fileInput.click();
    }
    
    // Ctrl/Cmd + Enter to process
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && !elements.processBtn.disabled) {
        e.preventDefault();
        processFiles();
    }
    
    // Ctrl/Cmd + Z to undo
    if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
        e.preventDefault();
        undo();
    }
    
    // Ctrl/Cmd + Shift + Z to redo
    if ((e.ctrlKey || e.metaKey) && e.key === 'z' && e.shiftKey) {
        e.preventDefault();
        redo();
    }
    
    // Ctrl/Cmd + P to preview
    if ((e.ctrlKey || e.metaKey) && e.key === 'p') {
        e.preventDefault();
        showPreview();
    }
    
    // Ctrl/Cmd + D to toggle dark mode
    if ((e.ctrlKey || e.metaKey) && e.key === 'd') {
        e.preventDefault();
        toggleTheme();
    }
    
    // Ctrl/Cmd + S to save config
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        saveConfiguration();
    }
    
    // Ctrl/Cmd + A to toggle advanced options
    if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
        e.preventDefault();
        const advancedButton = document.querySelector('[data-bs-target="#collapseAdvanced"]');
        if (advancedButton) {
            const collapse = new bootstrap.Collapse(document.getElementById('collapseAdvanced'));
            collapse.toggle();
        }
    }
}

// Handle File Selection
function handleFileSelection(files) {
    const validFiles = files.filter(file => {
        const extension = '.' + file.name.split('.').pop().toLowerCase();
        const validExtensions = ['.xlsx', '.xls', '.csv', '.tsv', '.txt'];
        
        if (!validExtensions.includes(extension)) {
            addLog('Validation', `Skipped ${file.name}: Invalid format`, 'warning');
            return false;
        }
        
        if (file.size > 100 * 1024 * 1024) {
            addLog('Validation', `Skipped ${file.name}: Exceeds 100MB limit`, 'warning');
            return false;
        }
        
        return true;
    });
    
    if (validFiles.length === 0) {
        alert('No valid files selected. Please upload Excel (.xlsx, .xls), CSV, TSV, or text files under 100MB.');
        return;
    }
    
    AppState.selectedFiles = validFiles;
    
    if (validFiles.length === 1) {
        // Single file mode
        const file = validFiles[0];
        elements.fileNameDisplay.textContent = file.name;
        elements.fileNameDisplay.style.color = 'var(--primary-color)';
        
        // Show file info
        elements.fileNameInfo.textContent = file.name;
        elements.fileSizeInfo.textContent = `(${formatFileSize(file.size)})`;
        elements.fileInfo.classList.remove('d-none');
        
        addLog('File Selected', `${file.name} (${formatFileSize(file.size)})`, 'info');
    } else {
        // Batch mode
        elements.fileNameDisplay.textContent = `${validFiles.length} files selected`;
        showBatchFileList(validFiles);
        addLog('Batch Selected', `${validFiles.length} files ready for processing`, 'info');
    }
    
    elements.processBtn.disabled = false;
    elements.previewBtn.disabled = false;
}

// Show Batch File List
function showBatchFileList(files) {
    elements.batchFiles.style.display = 'block';
    elements.fileCount.textContent = `${files.length} files`;
    
    const fileListHTML = files.map((file, index) => `
        <div class="file-item fade-in">
            <div class="file-item-info">
                <i class="bi bi-file-earmark-spreadsheet text-primary"></i>
                <div>
                    <div class="fw-bold small">${file.name}</div>
                    <div class="text-muted extra-small">${formatFileSize(file.size)}</div>
                </div>
            </div>
            <div class="file-item-actions">
                <button class="btn btn-sm btn-outline-danger" onclick="removeFile(${index})">
                    <i class="bi bi-x"></i>
                </button>
            </div>
        </div>
    `).join('');
    
    elements.fileList.innerHTML = fileListHTML;
}

// Remove File from Batch
function removeFile(index) {
    AppState.selectedFiles.splice(index, 1);
    
    if (AppState.selectedFiles.length === 0) {
        elements.batchFiles.style.display = 'none';
        elements.fileNameDisplay.textContent = 'Drop your file here or click to browse';
        elements.processBtn.disabled = true;
        elements.previewBtn.disabled = true;
    } else {
        showBatchFileList(AppState.selectedFiles);
        elements.fileNameDisplay.textContent = `${AppState.selectedFiles.length} files selected`;
    }
}

// Clear Current File
function clearFile() {
    AppState.selectedFiles = [];
    AppState.originalData = null;
    AppState.cleanedData = null;
    AppState.history = [];
    AppState.currentHistoryIndex = -1;
    
    elements.fileNameDisplay.textContent = 'Drop your file here or click to browse';
    elements.fileNameDisplay.style.color = '';
    elements.fileInfo.classList.add('d-none');
    elements.batchFiles.style.display = 'none';
    elements.processBtn.disabled = true;
    elements.previewBtn.disabled = true;
    elements.qualityScore.style.display = 'none';
    
    // Reset processing metrics
    AppState.processingMetrics = {
        startTime: null,
        endTime: null,
        rowsProcessed: 0,
        cellsProcessed: 0,
        issuesFixed: 0,
        columnsRemoved: 0,
        duplicatesRemoved: 0
    };
    
    addLog('System', 'File cleared', 'info');
}

// Format File Size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Add Log Entry
function addLog(action, details, type = 'info') {
    const now = new Date();
    const timeString = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    
    let icon = 'bi-info-circle';
    let colorClass = 'text-info';
    
    switch(type) {
        case 'success':
            icon = 'bi-check-circle';
            colorClass = 'text-success';
            break;
        case 'warning':
            icon = 'bi-exclamation-triangle';
            colorClass = 'text-warning';
            break;
        case 'error':
            icon = 'bi-x-circle';
            colorClass = 'text-danger';
            break;
        case 'debug':
            icon = 'bi-bug';
            colorClass = 'text-secondary';
            break;
    }
    
    const row = `
        <tr class="fade-in">
            <td class="text-muted" style="width: 150px;">
                <i class="bi ${icon} ${colorClass} me-2"></i>${timeString}
            </td>
            <td class="fw-semibold" style="width: 200px;">${action}</td>
            <td>${details}</td>
        </tr>
    `;
    
    elements.logBody.insertAdjacentHTML('afterbegin', row);
    elements.logArea.style.display = 'block';
    
    // Update status badge
    const statusBadge = elements.logArea.querySelector('.status-badge');
    if (statusBadge) {
        statusBadge.textContent = type.toUpperCase();
        statusBadge.className = `status-badge ${type}`;
    }
}

// Update Progress
function updateProgress(percent, message) {
    elements.progressBar.style.width = `${percent}%`;
    elements.progressPercent.textContent = `${Math.round(percent)}%`;
    elements.progressText.textContent = message;
    
    // Add pulsing animation for intermediate states
    if (percent > 0 && percent < 100) {
        elements.progressBar.classList.add('progress-bar-animated');
    } else {
        elements.progressBar.classList.remove('progress-bar-animated');
    }
}

// Toggle Theme
function toggleTheme() {
    if (document.body.getAttribute('data-bs-theme') === 'dark') {
        document.body.setAttribute('data-bs-theme', 'light');
        elements.themeIcon.className = 'bi bi-moon';
        addLog('System', 'Switched to light mode', 'info');
    } else {
        document.body.setAttribute('data-bs-theme', 'dark');
        elements.themeIcon.className = 'bi bi-sun';
        addLog('System', 'Switched to dark mode', 'info');
    }
    
    // Save preference
    localStorage.setItem('byteHooverTheme', document.body.getAttribute('data-bs-theme'));
}

// Apply Preset
function applyPreset(presetName) {
    const preset = Presets[presetName];
    if (!preset) return;
    
    // Update UI checkboxes
    Object.keys(preset.settings).forEach(key => {
        const element = document.getElementById(`check${key.charAt(0).toUpperCase() + key.slice(1)}`);
        if (element) {
            element.checked = preset.settings[key];
        }
    });
    
    addLog('Preset Applied', preset.name, 'success');
}

// Get Current Settings
function getCurrentSettings() {
    const settings = {
        capitalize: document.getElementById('checkCapitalize').checked,
        punctuation: document.getElementById('checkPunctuation').checked,
        spaces: document.getElementById('checkSpaces').checked,
        phone: document.getElementById('checkPhone').checked,
        dates: document.getElementById('checkDates').checked,
        emails: document.getElementById('checkEmails').checked,
        empty: document.getElementById('checkEmpty').checked,
        emptyCols: document.getElementById('checkEmptyCols').checked,
        dupes: document.getElementById('checkDupes').checked,
        markInvalid: document.getElementById('checkMarkInvalid') ? document.getElementById('checkMarkInvalid').checked : false,
        preserveNames: document.getElementById('checkPreserveNames') ? document.getElementById('checkPreserveNames').checked : true,
        strictDates: document.getElementById('checkStrictDates') ? document.getElementById('checkStrictDates').checked : false,
        anonymize: document.getElementById('checkAnonymize') ? document.getElementById('checkAnonymize').checked : false
    };
    
    // Ensure backward compatibility
    if (settings.markInvalid === undefined) settings.markInvalid = false;
    if (settings.preserveNames === undefined) settings.preserveNames = true;
    if (settings.strictDates === undefined) settings.strictDates = false;
    if (settings.anonymize === undefined) settings.anonymize = false;
    
    return settings;
}

// Load Export Settings from localStorage
function loadExportSettings() {
    const savedSettings = localStorage.getItem('byteHooverExportSettings');
    if (savedSettings) {
        try {
            const loadedSettings = JSON.parse(savedSettings);
            
            // Merge with default settings for backward compatibility
            AppState.exportSettings = {
                ...AppState.exportSettings,
                ...loadedSettings
            };
            
            // Update UI elements if they exist
            if (elements.exportFormat) {
                elements.exportFormat.value = AppState.exportSettings.format || 'xlsx';
            }
            
            if (elements.fileNamePrefix) {
                elements.fileNamePrefix.value = AppState.exportSettings.fileNamePrefix || 'cleaned_';
            }
            
            if (elements.zipOutput) {
                // Only check if JSZip is available
                if (AppState.exportSettings.compressOutput && typeof JSZip === 'undefined') {
                    AppState.exportSettings.compressOutput = false;
                    addLog('Settings', 'Compression disabled - JSZip not loaded', 'warning');
                }
                elements.zipOutput.checked = AppState.exportSettings.compressOutput || false;
            }
            
            if (elements.preserveStructure) {
                elements.preserveStructure.checked = AppState.exportSettings.preserveOriginalStructure !== false;
            }
            
            addLog('Settings Loaded', 'Export preferences restored', 'info');
        } catch (error) {
            console.warn('Failed to load export settings:', error);
            // Use default settings
        }
    }
}

// Save Export Settings to localStorage
function saveExportSettings() {
    // Update AppState with current values
    AppState.exportSettings = {
        format: elements.exportFormat ? elements.exportFormat.value : 'xlsx',
        fileNamePrefix: elements.fileNamePrefix ? elements.fileNamePrefix.value : 'cleaned_',
        compressOutput: elements.zipOutput ? elements.zipOutput.checked : false,
        preserveOriginalStructure: elements.preserveStructure ? elements.preserveStructure.checked : true
    };
    
    // Save to localStorage
    localStorage.setItem('byteHooverExportSettings', JSON.stringify(AppState.exportSettings));
}

// Process Files
async function processFiles() {
    if (AppState.selectedFiles.length === 0 || AppState.isProcessing) return;
    
    AppState.isProcessing = true;
    AppState.processingMetrics.startTime = performance.now();
    
    // Reset processing metrics
    AppState.processingMetrics = {
        startTime: performance.now(),
        endTime: null,
        rowsProcessed: 0,
        cellsProcessed: 0,
        issuesFixed: 0,
        columnsRemoved: 0,
        duplicatesRemoved: 0
    };
    
    // Reset UI
    elements.logBody.innerHTML = '';
    elements.logArea.style.display = 'none';
    elements.progressWrapper.style.display = 'block';
    elements.processBtn.disabled = true;
    elements.processBtn.innerHTML = '<span class="spinner"></span> PROCESSING...';
    elements.processBtn.classList.add('processing');
    elements.qualityScore.style.display = 'none';
    
    updateProgress(0, 'Initializing...');
    
    try {
        if (AppState.selectedFiles.length === 1) {
            // Single file processing
            await processSingleFile(AppState.selectedFiles[0]);
        } else {
            // Batch processing
            await processBatchFiles();
        }
        
        AppState.processingMetrics.endTime = performance.now();
        showQualityReport();
        
    } catch (error) {
        console.error('Processing error:', error);
        addLog('Error', error.message, 'error');
        updateProgress(0, 'Processing failed');
        
        // Reset UI on error
        setTimeout(() => {
            elements.progressWrapper.style.display = 'none';
            resetProcessButton();
        }, 2000);
    } finally {
        AppState.isProcessing = false;
    }
}

// Process Single File
async function processSingleFile(file) {
    addLog('Processing', `Starting: ${file.name}`, 'info');
    
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                updateProgress(10, 'Reading file...');
                
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { 
                    type: 'array', 
                    cellDates: true,
                    cellNF: false,
                    cellText: false,
                    raw: false,
                    dateNF: 'yyyy-mm-dd'
                });
                
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                let rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                if (rows.length === 0) {
                    throw new Error('The file appears to be empty');
                }
                
                // Preserve original structure if setting is enabled
                const originalStructure = {
                    headerRow: rows[0] || [],
                    columnCount: rows[0] ? rows[0].length : 0,
                    hasHeaders: false
                };
                
                // Check if first row looks like headers (mostly text)
                if (rows.length > 0) {
                    const headerRow = rows[0];
                    const textCount = headerRow.filter(cell => 
                        cell && typeof cell === 'string' && cell.trim().length > 0
                    ).length;
                    originalStructure.hasHeaders = textCount > (headerRow.length / 2);
                }
                
                AppState.originalData = rows;
                const settings = getCurrentSettings();
                
                // Save to history
                saveToHistory(rows, 'Original data loaded');
                
                // Process data
                updateProgress(25, 'Cleaning data...');
                const cleanedRows = await cleanData(rows, settings, originalStructure);
                
                // Remove duplicates if enabled
                if (settings.dupes) {
                    updateProgress(70, 'Removing duplicates...');
                    const { dedupedRows, removedCount } = removeDuplicates(cleanedRows, originalStructure.hasHeaders);
                    if (removedCount > 0) {
                        addLog('Duplicates Removed', `${removedCount} duplicate rows removed`, 'success');
                        AppState.processingMetrics.duplicatesRemoved = removedCount;
                        AppState.processingMetrics.issuesFixed += removedCount;
                    }
                    AppState.cleanedData = dedupedRows;
                } else {
                    AppState.cleanedData = cleanedRows;
                }
                
                // Export file
                updateProgress(90, 'Preparing download...');
                await exportFile(AppState.cleanedData, file.name, originalStructure);
                
                updateProgress(100, 'Processing complete!');
                
                // Update metrics
                AppState.processingMetrics.rowsProcessed = AppState.cleanedData.length;
                AppState.processingMetrics.cellsProcessed = AppState.cleanedData.reduce((sum, row) => sum + (row?.length || 0), 0);
                
                // Final UI updates
                setTimeout(() => {
                    elements.progressWrapper.style.display = 'none';
                    elements.processBtn.innerHTML = '<i class="bi bi-check-circle me-2"></i>PROCESSING COMPLETE';
                    elements.processBtn.classList.remove('processing');
                    elements.processBtn.classList.add('success');
                    
                    // Show log area
                    elements.logArea.style.display = 'block';
                    
                    // Reset button after 3 seconds
                    setTimeout(resetProcessButton, 3000);
                }, 500);
                
                resolve();
                
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => reject(new Error('Error reading file'));
        reader.readAsArrayBuffer(file);
    });
}

// Enhanced Data Cleaning Function
async function cleanData(rows, settings, originalStructure = null) {
    let cleanedRows = [...rows];
    
    // Remove empty rows (only if entire row is empty)
    if (settings.empty) {
        const initialCount = cleanedRows.length;
        cleanedRows = cleanedRows.filter((row, rowIndex) => {
            // For header row (first row), keep it even if empty
            if (rowIndex === 0 && originalStructure?.hasHeaders) return true;
            
            // Check if row has any meaningful content
            const hasContent = row && row.some(cell => {
                if (cell === null || cell === undefined) return false;
                const strValue = String(cell).trim();
                return strValue !== '';
            });
            return hasContent;
        });
        
        const removedCount = initialCount - cleanedRows.length;
        if (removedCount > 0) {
            addLog('Empty Rows Removed', `${removedCount} rows removed`, 'success');
            AppState.processingMetrics.issuesFixed += removedCount;
        }
    }
    
    // Clean each cell
    cleanedRows = cleanedRows.map((row, rowIndex) => {
        return row.map((cell, colIndex) => {
            if (cell === null || cell === undefined || cell === '') {
                return '';
            }
            
            let value = cell;
            const isHeaderCell = (rowIndex === 0 && originalStructure?.hasHeaders);
            
            // Anonymize sensitive data if enabled
            if (settings.anonymize) {
                value = anonymizeData(value, colIndex, isHeaderCell);
            }
            
            // Handle dates
            if (settings.dates) {
                value = validateAndFormatDate(value, settings);
            } else if (value instanceof Date) {
                // Keep as string representation but don't reformat
                value = value.toString();
            }
            
            if (typeof value === 'string' || typeof value === 'number') {
                value = String(value);
                
                // Don't process headers if preserveNames is enabled
                if (isHeaderCell && settings.preserveNames) {
                    // Only minimal cleaning for headers
                    value = value.trim().replace(/\s+/g, ' ');
                } else {
                    // Trim whitespace - but be careful with names!
                    if (settings.spaces) {
                        // Preserve spaces in names but clean extra spaces
                        if (isLikelyName(value) && settings.preserveNames) {
                            value = cleanNameSpacing(value);
                        } else {
                            value = value.trim().replace(/\s+/g, ' ');
                        }
                    } else {
                        // Still do minimal trimming
                        value = value.trim();
                    }
                    
                    // Remove excessive punctuation but preserve important characters
                    if (settings.punctuation) {
                        value = cleanPunctuation(value);
                    }
                    
                    // Capitalization - be smart about it
                    if (settings.capitalize) {
                        value = smartCapitalize(value, isHeaderCell);
                    }
                    
                    // Format phone numbers - validate and format
                    if (settings.phone && isPhoneNumber(value)) {
                        value = formatPhoneNumber(value);
                    }
                    
                    // Validate emails
                    if (settings.emails && value.includes('@')) {
                        if (!DataPatterns.EMAIL.test(value)) {
                            addLog('Invalid Email', `Row ${rowIndex + 1}, Col ${colIndex + 1}: "${value}"`, 'warning');
                            // Optionally mark invalid emails
                            if (settings.markInvalid) {
                                value = `[INVALID] ${value}`;
                            }
                        } else {
                            // Standardize email format (lowercase)
                            value = value.toLowerCase();
                        }
                    }
                }
            }
            
            return value;
        });
    });
    
    // Remove empty columns - but be careful not to shift data
    if (settings.emptyCols && cleanedRows.length > 0) {
        const colCount = cleanedRows[0].length;
        const emptyColIndices = [];
        
        for (let c = 0; c < colCount; c++) {
            // Check if column is completely empty
            const isColEmpty = cleanedRows.every((row, rowIndex) => {
                // For header row, check if header name is meaningful
                if (rowIndex === 0 && originalStructure?.hasHeaders) {
                    const headerValue = row[c];
                    return !headerValue || String(headerValue).trim() === '' || 
                           headerValue.toString().toLowerCase().includes('unnamed');
                }
                // For data rows, check if cell is empty
                return !row[c] || String(row[c]).trim() === '';
            });
            
            if (isColEmpty) emptyColIndices.push(c);
        }
        
        // Only remove if we found empty columns AND preserveStructure is not strictly enforced
        if (emptyColIndices.length > 0 && !AppState.exportSettings.preserveOriginalStructure) {
            // Remove from end to beginning to maintain indices
            for (let i = emptyColIndices.length - 1; i >= 0; i--) {
                const colIndex = emptyColIndices[i];
                cleanedRows = cleanedRows.map(row => {
                    // Create new array without the empty column
                    const newRow = [...row];
                    newRow.splice(colIndex, 1);
                    return newRow;
                });
            }
            
            addLog('Empty Columns Removed', `${emptyColIndices.length} columns removed`, 'success');
            AppState.processingMetrics.columnsRemoved = emptyColIndices.length;
            AppState.processingMetrics.issuesFixed += emptyColIndices.length;
        } else if (emptyColIndices.length > 0) {
            addLog('Structure Preserved', `${emptyColIndices.length} empty columns kept for structure`, 'info');
        }
    }
    
    return cleanedRows;
}

// Enhanced Helper Functions
function isLikelyName(value) {
    if (typeof value !== 'string') return false;
    
    const namePattern = /^[A-Za-z\s\.\-\',]+$/;
    const commonTitles = /\b(Mr|Mrs|Ms|Miss|Dr|Prof|Rev|Hon|Sir|Madam|Lady|Lord)\.?\b/i;
    const nameParts = value.trim().split(/\s+/);
    
    // Check if it has typical name structure
    return (namePattern.test(value) && nameParts.length <= 4) || 
           commonTitles.test(value) ||
           (nameParts.length >= 2 && nameParts.length <= 4 && 
            nameParts.every(part => part.length > 1 && /^[A-Z]/.test(part)));
}

function cleanNameSpacing(name) {
    return name
        .trim()
        .replace(/\s+/g, ' ') // Multiple spaces to single
        .replace(/\s*-\s*/g, '-') // Clean spaces around hyphens
        .replace(/\s*\.\s*/g, '.') // Clean spaces around dots
        .replace(/\s*,\s*/g, ', ') // Clean spaces around commas
        .replace(/\s{2,}/g, ' ') // Final cleanup
        .replace(/^\s+|\s+$/g, ''); // Trim again
}

function cleanPunctuation(value) {
    // Remove excessive punctuation but preserve meaningful characters
    return value
        .replace(/[^\w\s@\+\.\,\-\:\/\(\)\&\%\$\#\*\?\!'"<>]/gi, '') // Allow common punctuation
        .replace(/([\.\,\?\!])\1+/g, '$1') // Remove duplicate punctuation
        .replace(/\s*([\.\,\?\!])\s*/g, '$1 ') // Standardize spacing around punctuation
        .replace(/\s+/g, ' ') // Clean extra spaces
        .trim();
}

function smartCapitalize(value, isHeader = false) {
    // Don't capitalize if it looks like an email, URL, or acronym
    if (value.includes('@') || value.startsWith('http') || /^[A-Z0-9\.\-_]+$/.test(value)) {
        return value;
    }
    
    // Handle names specially
    if (isLikelyName(value)) {
        return capitalizeNames(value);
    }
    
    // For headers, use title case
    if (isHeader) {
        return value.split(' ').map(word => {
            if (word.length === 0) return word;
            return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
        }).join(' ');
    }
    
    // General capitalization
    return value.split(' ').map((word, index, words) => {
        if (word.length === 0) return word;
        
        // Don't capitalize small words unless they're the first/last word
        const smallWords = ['a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'in', 'nor', 'of', 'on', 'or', 'the', 'to', 'up', 'yet', 'vs', 'via'];
        
        if (smallWords.includes(word.toLowerCase()) && 
            index > 0 && index < words.length - 1) {
            return word.toLowerCase();
        }
        
        // Preserve ALL CAPS acronyms
        if (word === word.toUpperCase() && word.length <= 5 && /^[A-Z]+$/.test(word)) {
            return word;
        }
        
        // Preserve mixed case words (like iPhone, eBay)
        if (/^[a-z][A-Z]/.test(word) || /^[A-Z][a-z][A-Z]/.test(word)) {
            return word;
        }
        
        // Standard capitalization
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    }).join(' ');
}

function capitalizeNames(name) {
    const suffixes = ['jr', 'sr', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x', 'phd', 'md', 'dds', 'dvm', 'esq', 'cpa'];
    const prefixes = ['mc', 'mac', 'van', 'von', 'de', 'la', 'di', 'del', 'della', 'el', 'al'];
    const compoundNames = ['van der', 'van den', 'de la'];
    
    // Handle compound names first
    for (const compound of compoundNames) {
        if (name.toLowerCase().includes(compound)) {
            const parts = name.split(new RegExp(`(${compound})`, 'i'));
            return parts.map(part => 
                compoundNames.includes(part.toLowerCase()) ? 
                    part.toLowerCase() : 
                    capitalizeNamePart(part, suffixes, prefixes)
            ).join('');
        }
    }
    
    return name.split(' ').map((part, index, parts) => 
        capitalizeNamePart(part, suffixes, prefixes, index, parts)
    ).join(' ');
}

function capitalizeNamePart(part, suffixes, prefixes, index = 0, parts = []) {
    const lowerPart = part.toLowerCase();
    
    // Handle suffixes (last part)
    if (index === parts.length - 1 && suffixes.includes(lowerPart.replace(/\./g, ''))) {
        return part.toUpperCase();
    }
    
    // Handle prefixes (not first part)
    if (index > 0 && prefixes.includes(lowerPart)) {
        return part.charAt(0).toUpperCase() + part.slice(1);
    }
    
    // Handle hyphenated names
    if (part.includes('-')) {
        return part.split('-').map(subPart => 
            subPart.charAt(0).toUpperCase() + subPart.slice(1)
        ).join('-');
    }
    
    // Handle apostrophe names (O'Connor, D'Angelo)
    if (part.includes("'")) {
        const apostropheParts = part.split("'");
        return apostropheParts.map((subPart, i) => 
            i === 0 ? subPart.charAt(0).toUpperCase() + subPart.slice(1) : subPart
        ).join("'");
    }
    
    // Handle "Mc" and "Mac" names
    if (part.toLowerCase().startsWith('mc') && part.length > 2) {
        return 'Mc' + part.charAt(2).toUpperCase() + part.slice(3);
    }
    
    if (part.toLowerCase().startsWith('mac') && part.length > 3) {
        return 'Mac' + part.charAt(3).toUpperCase() + part.slice(4);
    }
    
    // Standard capitalization
    return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
}

function isPhoneNumber(value) {
    if (typeof value !== 'string') return false;
    
    const cleaned = value.replace(/[\s\-\(\)\.\+]/g, '');
    
    // Check for various phone number formats
    const patterns = [
        /^\d{10}$/, // 1234567890
        /^1\d{10}$/, // 11234567890
        /^\+\d{11,15}$/, // +11234567890
        /^\d{3}-\d{3}-\d{4}$/, // 123-456-7890
        /^\(\d{3}\)\s*\d{3}-\d{4}$/, // (123) 456-7890
        /^\d{3}\.\d{3}\.\d{4}$/, // 123.456.7890
        /^\d{3}\s\d{3}\s\d{4}$/ // 123 456 7890
    ];
    
    return patterns.some(pattern => pattern.test(value)) || 
           DataPatterns.PHONE_US.test(value) ||
           DataPatterns.PHONE_INTERNATIONAL.test(cleaned);
}

function formatPhoneNumber(phone) {
    const digits = phone.replace(/\D/g, '');
    
    if (digits.length === 10) {
        // US format: (123) 456-7890
        return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
    } else if (digits.length === 11 && digits.startsWith('1')) {
        // US with country code: +1 (123) 456-7890
        return `+1 (${digits.slice(1, 4)}) ${digits.slice(4, 7)}-${digits.slice(7)}`;
    } else if (digits.length >= 8 && digits.length <= 15) {
        // International format
        const countryCode = digits.length > 10 ? `+${digits.slice(0, digits.length - 10)} ` : '';
        const number = digits.slice(-10);
        return `${countryCode}(${number.slice(0, 3)}) ${number.slice(3, 6)}-${number.slice(6)}`;
    }
    
    // Return cleaned but not reformatted if we can't format it properly
    return phone.replace(/\D/g, '');
}

function validateAndFormatDate(value, settings) {
    // If it's already a date object
    if (value instanceof Date) {
        // Check if it's a valid date
        if (isNaN(value.getTime())) {
            return value.toString(); // Return as string if invalid
        }
        if (settings.strictDates) {
            return value.toISOString().split('T')[0]; // YYYY-MM-DD
        }
        return value.toLocaleDateString();
    }
    
    // If it's a string, try to parse it
    if (typeof value === 'string') {
        const trimmed = value.trim();
        
        // Try ISO format first (YYYY-MM-DD)
        const isoMatch = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (isoMatch) {
            const [_, year, month, day] = isoMatch;
            const date = new Date(year, month - 1, day);
            if (!isNaN(date.getTime())) {
                return settings.strictDates ? `${year.padStart(4, '0')}-${month.padStart(2, '0')}-${day.padStart(2, '0')}` : trimmed;
            }
        }
        
        // Try US format (MM/DD/YYYY or M/D/YYYY)
        const usMatch = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
        if (usMatch) {
            const [_, month, day, year] = usMatch;
            const date = new Date(year, month - 1, day);
            if (!isNaN(date.getTime())) {
                return settings.strictDates ? 
                    `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}` :
                    `${month}/${day}/${year}`;
            }
        }
        
        // Try EU format (DD.MM.YYYY or DD-MM-YYYY)
        const euMatch = trimmed.match(/^(\d{1,2})[\.\-](\d{1,2})[\.\-](\d{4})$/);
        if (euMatch) {
            const [_, day, month, year] = euMatch;
            const date = new Date(year, month - 1, day);
            if (!isNaN(date.getTime())) {
                return settings.strictDates ? 
                    `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}` :
                    `${day}.${month}.${year}`;
            }
        }
        
        // Try year first format (YYYY/MM/DD)
        const yearFirstMatch = trimmed.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
        if (yearFirstMatch) {
            const [_, year, month, day] = yearFirstMatch;
            const date = new Date(year, month - 1, day);
            if (!isNaN(date.getTime())) {
                return settings.strictDates ? 
                    `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}` :
                    `${year}/${month}/${day}`;
            }
        }
    }
    
    // If we can't parse it as a date, return original
    return value;
}

function anonymizeData(value, columnIndex, isHeader) {
    if (isHeader || typeof value !== 'string') return value;
    
    const strValue = String(value).toLowerCase();
    
    // Anonymize based on content patterns
    if (strValue.includes('@')) {
        // Email: replace with generic pattern
        const [local, domain] = strValue.split('@');
        return `user${Math.floor(Math.random() * 10000)}@${domain}`;
    }
    
    if (isPhoneNumber(strValue)) {
        // Phone: replace with generic pattern
        const digits = strValue.replace(/\D/g, '');
        if (digits.length === 10) {
            return `(XXX) XXX-${digits.slice(-4)}`;
        }
        return `XXX-XXX-${digits.slice(-4)}`;
    }
    
    if (DataPatterns.SSN.test(strValue)) {
        // SSN: show only last 4 digits
        return `XXX-XX-${strValue.slice(-4)}`;
    }
    
    if (DataPatterns.CREDIT_CARD.test(strValue.replace(/[\s\-]/g, ''))) {
        // Credit card: show only last 4 digits
        const digits = strValue.replace(/\D/g, '');
        return `XXXX-XXXX-XXXX-${digits.slice(-4)}`;
    }
    
    // For names, replace with generic names
    if (isLikelyName(strValue)) {
        const firstNames = ['John', 'Jane', 'Alex', 'Sam', 'Chris', 'Taylor', 'Jordan', 'Morgan'];
        const lastNames = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Miller', 'Davis'];
        const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
        const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
        return `${firstName} ${lastName}`;
    }
    
    return value;
}

// Enhanced Remove Duplicates with header handling
function removeDuplicates(rows, hasHeaders = false) {
    if (rows.length < 2) return { dedupedRows: rows, removedCount: 0 };
    
    const seen = new Set();
    const uniqueRows = [];
    let startIndex = hasHeaders ? 1 : 0;
    let removedCount = 0;
    
    // Always include header row if present
    if (hasHeaders && rows.length > 0) {
        uniqueRows.push(rows[0]);
    }
    
    for (let i = startIndex; i < rows.length; i++) {
        const row = rows[i];
        const rowString = JSON.stringify(row);
        
        if (!seen.has(rowString)) {
            seen.add(rowString);
            uniqueRows.push(row);
        } else {
            removedCount++;
        }
    }
    
    return { dedupedRows: uniqueRows, removedCount };
}

// Enhanced Export File with structure preservation
async function exportFile(data, originalName, originalStructure = null) {
    const { format, fileNamePrefix, compressOutput } = AppState.exportSettings;
    const timestamp = new Date().toLocaleString('en-CA', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    }).replace(/[\/,: ]/g, '-').replace('--', '-');
    
    const baseName = originalName.split('.')[0];
    const safeBaseName = baseName.replace(/[^\w\s-]/g, '_').replace(/\s+/g, '_');
    
    try {
        if (compressOutput && typeof JSZip !== 'undefined') {
            const zip = new JSZip();
            
            switch(format) {
                case 'xlsx':
                    const worksheet = XLSX.utils.aoa_to_sheet(data);
                    const workbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(workbook, worksheet, 'Cleaned_Data');
                    
                    const xlsxData = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
                    const xlsxFileName = `${safeBaseName}_${timestamp}.xlsx`;
                    zip.file(xlsxFileName, xlsxData);
                    break;
                    
                case 'csv':
                    const csvContent = data.map(row => 
                        row.map(cell => {
                            const strValue = cell === null || cell === undefined ? '' : String(cell);
                            // Escape quotes and wrap in quotes if contains comma, quote, or newline
                            if (strValue.includes('"') || strValue.includes(',') || strValue.includes('\n')) {
                                return `"${strValue.replace(/"/g, '""')}"`;
                            }
                            return strValue;
                        }).join(',')
                    ).join('\n');
                    
                    const csvFileName = `${safeBaseName}_${timestamp}.csv`;
                    zip.file(csvFileName, csvContent);
                    break;
                    
                case 'json':
                    let jsonData;
                    if (originalStructure?.hasHeaders && data.length > 0) {
                        const headers = data[0];
                        jsonData = data.slice(1).map(row => {
                            const obj = {};
                            headers.forEach((header, index) => {
                                obj[header || `Column${index + 1}`] = row[index];
                            });
                            return obj;
                        });
                    } else {
                        jsonData = data.map((row, index) => ({
                            row: index + 1,
                            data: row
                        }));
                    }
                    
                    const jsonFileName = `${safeBaseName}_${timestamp}.json`;
                    zip.file(jsonFileName, JSON.stringify(jsonData, null, 2));
                    break;
            }
            
            // Generate and download the ZIP file
            const zipContent = await zip.generateAsync({ type: 'blob' });
            const zipFileName = `${fileNamePrefix}${safeBaseName}_${timestamp}.zip`;
            saveAsBlob(zipContent, zipFileName);
            addLog('File Exported', `Compressed ${zipFileName}`, 'success');
            
        } else {
            let fileName, blob;
            
            switch(format) {
                case 'xlsx':
                    const worksheet = XLSX.utils.aoa_to_sheet(data);
                    const workbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(workbook, worksheet, 'Cleaned_Data');
                    
                    fileName = `${fileNamePrefix}${safeBaseName}_${timestamp}.xlsx`;
                    XLSX.writeFile(workbook, fileName);
                    break;
                    
                case 'csv':
                    const csvContent = data.map(row => 
                        row.map(cell => {
                            const strValue = cell === null || cell === undefined ? '' : String(cell);
                            if (strValue.includes('"') || strValue.includes(',') || strValue.includes('\n')) {
                                return `"${strValue.replace(/"/g, '""')}"`;
                            }
                            return strValue;
                        }).join(',')
                    ).join('\n');
                    
                    fileName = `${fileNamePrefix}${safeBaseName}_${timestamp}.csv`;
                    blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                    saveAsBlob(blob, fileName);
                    break;
                    
                case 'json':
                    let jsonData;
                    if (originalStructure?.hasHeaders && data.length > 0) {
                        const headers = data[0];
                        jsonData = data.slice(1).map(row => {
                            const obj = {};
                            headers.forEach((header, index) => {
                                obj[header || `Column${index + 1}`] = row[index];
                            });
                            return obj;
                        });
                    } else {
                        jsonData = data.map((row, index) => ({
                            row: index + 1,
                            data: row
                        }));
                    }
                    
                    fileName = `${fileNamePrefix}${safeBaseName}_${timestamp}.json`;
                    blob = new Blob([JSON.stringify(jsonData, null, 2)], { type: 'application/json' });
                    saveAsBlob(blob, fileName);
                    break;
            }
            
            if (compressOutput && typeof JSZip === 'undefined') {
                addLog('Compression Warning', 'JSZip library not loaded - exported without compression', 'warning');
            }
            
            addLog('File Exported', fileName, 'success');
        }
    } catch (error) {
        console.error('Export error:', error);
        addLog('Export Error', error.message, 'error');
        throw error;
    }
}

// Save Blob as File
function saveAsBlob(blob, fileName) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 100);
}

// Enhanced Show Quality Report
function showQualityReport() {
    const duration = (AppState.processingMetrics.endTime - AppState.processingMetrics.startTime) / 1000;
    const rowsPerSecond = AppState.processingMetrics.rowsProcessed > 0 ? 
        Math.round(AppState.processingMetrics.rowsProcessed / duration) : 0;
    
    // Calculate quality score
    let qualityScore = 100;
    const issuePenalty = Math.min(AppState.processingMetrics.issuesFixed * 0.2, 40); // Max 40% penalty
    const duplicatePenalty = Math.min(AppState.processingMetrics.duplicatesRemoved * 0.3, 20); // Max 20% penalty
    
    qualityScore = Math.max(50, 100 - issuePenalty - duplicatePenalty);
    
    elements.qualityScoreValue.textContent = `${Math.round(qualityScore)}%`;
    elements.issuesFixed.textContent = AppState.processingMetrics.issuesFixed;
    elements.processTime.textContent = `${duration.toFixed(2)}s`;
    elements.processingSpeed.textContent = `${rowsPerSecond}`;
    
    // Add quality rating
    let qualityRating = 'Excellent';
    if (qualityScore < 70) qualityRating = 'Good';
    if (qualityScore < 60) qualityRating = 'Fair';
    if (qualityScore < 50) qualityRating = 'Poor';
    
    // Update quality score display with rating
    const scoreDisplay = elements.qualityScoreValue.parentElement;
    if (scoreDisplay.querySelector('.quality-rating')) {
        scoreDisplay.querySelector('.quality-rating').textContent = qualityRating;
    } else {
        const ratingSpan = document.createElement('span');
        ratingSpan.className = 'quality-rating small ms-2';
        ratingSpan.textContent = qualityRating;
        scoreDisplay.appendChild(ratingSpan);
    }
    
    elements.qualityScore.style.display = 'block';
}

// Reset Process Button
function resetProcessButton() {
    elements.processBtn.disabled = false;
    elements.processBtn.innerHTML = '<i class="bi bi-play-fill me-2"></i>Inhale & Cleanse Data';
    elements.processBtn.classList.remove('processing', 'success');
}

// Show Preview
function showPreview() {
    if (!AppState.originalData && AppState.selectedFiles.length === 0) {
        alert('Please select a file first!');
        return;
    }
    
    // If we have data, show it
    if (AppState.originalData && AppState.cleanedData) {
        showDataPreview(AppState.originalData, AppState.cleanedData);
    } else if (AppState.selectedFiles.length === 1) {
        // Load and preview the file
        previewSelectedFile();
    }
}

// Show Data Preview
function showDataPreview(originalData, cleanedData) {
    const previewCount = Math.min(50, originalData.length);
    document.getElementById('previewRowCount').textContent = `${previewCount} rows`;
    
    // Show original data
    const originalPreview = document.getElementById('originalPreview');
    originalPreview.innerHTML = generatePreviewTable(originalData.slice(0, previewCount), 'Original');
    
    // Show cleaned data
    const cleanedPreview = document.getElementById('cleanedPreview');
    cleanedPreview.innerHTML = generatePreviewTable(cleanedData.slice(0, previewCount), 'Cleaned');
    
    // Update comparison stats
    const originalRows = originalData.length;
    const cleanedRows = cleanedData.length;
    const diffRows = originalRows - cleanedRows;
    
    const statsDiv = document.getElementById('previewStats');
    if (statsDiv) {
        statsDiv.innerHTML = `
            <div class="row">
                <div class="col-6">
                    <div class="card">
                        <div class="card-body text-center">
                            <h6 class="card-subtitle mb-1 text-muted">Original Rows</h6>
                            <h4 class="card-title">${originalRows}</h4>
                        </div>
                    </div>
                </div>
                <div class="col-6">
                    <div class="card">
                        <div class="card-body text-center">
                            <h6 class="card-subtitle mb-1 text-muted">Cleaned Rows</h6>
                            <h4 class="card-title ${diffRows > 0 ? 'text-success' : ''}">${cleanedRows}</h4>
                        </div>
                    </div>
                </div>
            </div>
            ${diffRows > 0 ? `<div class="mt-2 text-center"><small class="text-success"><i class="bi bi-check-circle me-1"></i>${diffRows} rows cleaned</small></div>` : ''}
        `;
    }
    
    // Show modal
    new bootstrap.Modal(document.getElementById('previewModal')).show();
}

// Generate Preview Table
function generatePreviewTable(data, title) {
    if (!data || data.length === 0) {
        return `<div class="alert alert-info">No data to preview</div>`;
    }
    
    let html = `<div class="preview-table">
        <table class="table table-sm table-hover mb-0">
            <thead>
                <tr>`;
    
    // Headers (use first row or generate column numbers)
    if (data[0]) {
        data[0].forEach((cell, index) => {
            const headerText = cell && String(cell).trim() !== '' ? 
                String(cell).slice(0, 30) + (String(cell).length > 30 ? '...' : '') : 
                `Column ${index + 1}`;
            html += `<th title="${cell || `Column ${index + 1}`}">${headerText}</th>`;
        });
    }
    
    html += `</tr></thead><tbody>`;
    
    // Data rows
    const startRow = data[0] && data[0].some(cell => cell && String(cell).trim() !== '') ? 1 : 0;
    for (let rowIndex = startRow; rowIndex < Math.min(data.length, 11); rowIndex++) {
        const row = data[rowIndex];
        html += `<tr>`;
        (row || []).forEach((cell, cellIndex) => {
            const displayValue = cell === null || cell === undefined ? '' : String(cell);
            const titleAttr = displayValue.length > 30 ? `title="${displayValue.replace(/"/g, '&quot;')}"` : '';
            const truncatedValue = displayValue.slice(0, 30);
            html += `<td ${titleAttr}>${truncatedValue}${displayValue.length > 30 ? '...' : ''}</td>`;
        });
        html += `</tr>`;
    }
    
    if (data.length > 11) {
        html += `<tr><td colspan="${data[0]?.length || 1}" class="text-center text-muted">... and ${data.length - 11} more rows</td></tr>`;
    }
    
    html += `</tbody></table></div>`;
    return html;
}

// Preview Selected File
function previewSelectedFile() {
    const file = AppState.selectedFiles[0];
    const reader = new FileReader();
    
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            const previewCount = Math.min(50, rows.length);
            document.getElementById('previewRowCount').textContent = `${previewCount} rows`;
            
            const previewDiv = document.getElementById('originalPreview');
            previewDiv.innerHTML = generatePreviewTable(rows.slice(0, previewCount), 'Original');
            
            // Update stats for single preview
            const statsDiv = document.getElementById('previewStats');
            if (statsDiv) {
                statsDiv.innerHTML = `
                    <div class="alert alert-info">
                        <i class="bi bi-info-circle me-2"></i>
                        Select cleaning options and click "Inhale & Cleanse Data" to see cleaned version
                    </div>
                `;
            }
            
            new bootstrap.Modal(document.getElementById('previewModal')).show();
            
        } catch (error) {
            alert('Error previewing file: ' + error.message);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// History Management
function saveToHistory(data, action) {
    AppState.history = AppState.history.slice(0, AppState.currentHistoryIndex + 1);
    AppState.history.push({
        data: JSON.parse(JSON.stringify(data)),
        action: action,
        timestamp: new Date()
    });
    AppState.currentHistoryIndex++;
    updateUndoRedoButtons();
}

function undo() {
    if (AppState.currentHistoryIndex > 0) {
        AppState.currentHistoryIndex--;
        const historyItem = AppState.history[AppState.currentHistoryIndex];
        AppState.cleanedData = historyItem.data;
        addLog('Undo', historyItem.action, 'info');
        updateUndoRedoButtons();
    }
}

function redo() {
    if (AppState.currentHistoryIndex < AppState.history.length - 1) {
        AppState.currentHistoryIndex++;
        const historyItem = AppState.history[AppState.currentHistoryIndex];
        AppState.cleanedData = historyItem.data;
        addLog('Redo', historyItem.action, 'info');
        updateUndoRedoButtons();
    }
}

function updateUndoRedoButtons() {
    elements.undoBtn.disabled = AppState.currentHistoryIndex <= 0;
    elements.redoBtn.disabled = AppState.currentHistoryIndex >= AppState.history.length - 1;
}

// Configuration Management
function saveConfiguration() {
    const name = prompt('Enter a name for this configuration:');
    if (!name) return;
    
    const config = {
        id: Date.now(),
        name: name,
        timestamp: new Date().toISOString(),
        settings: getCurrentSettings(),
        exportSettings: { ...AppState.exportSettings }
    };
    
    AppState.configurations.push(config);
    localStorage.setItem('byteHooverConfigs', JSON.stringify(AppState.configurations));
    
    addLog('Configuration Saved', name, 'success');
    
    // Show success notification
    const notification = document.createElement('div');
    notification.className = 'alert alert-success alert-dismissible fade show position-fixed top-0 end-0 m-3';
    notification.style.zIndex = '1060';
    notification.innerHTML = `
        <i class="bi bi-check-circle me-2"></i>
        <strong>Configuration saved!</strong> "${name}" has been saved.
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    document.body.appendChild(notification);
    
    setTimeout(() => {
        if (notification.parentNode) {
            notification.remove();
        }
    }, 3000);
}

function loadConfigurations() {
    const saved = localStorage.getItem('byteHooverConfigs');
    if (saved) {
        try {
            AppState.configurations = JSON.parse(saved);
        } catch (error) {
            console.warn('Failed to load configurations:', error);
            AppState.configurations = [];
        }
    }
}

// Parse URL Parameters
function parseURLParameters() {
    const params = new URLSearchParams(window.location.search);
    
    if (params.has('preset')) {
        const preset = params.get('preset');
        if (Presets[preset]) {
            applyPreset(preset);
            addLog('URL Preset', `Applied "${preset}" preset from URL`, 'info');
        }
    }
    
    if (params.has('file')) {
        // Could implement file loading from URL here
        addLog('URL Parameter', 'File parameter detected - manual upload required', 'info');
    }
}

// Batch Processing
async function processBatchFiles() {
    addLog('Batch Processing', `Starting batch of ${AppState.selectedFiles.length} files`, 'info');
    
    const totalFiles = AppState.selectedFiles.length;
    const batchResults = {
        success: 0,
        failed: 0,
        files: []
    };
    
    for (let i = 0; i < totalFiles; i++) {
        const file = AppState.selectedFiles[i];
        const progressPercent = ((i / totalFiles) * 90).toFixed(1);
        updateProgress(progressPercent, `Processing ${i + 1}/${totalFiles}: ${file.name}`);
        
        try {
            // Process each file individually
            const reader = new FileReader();
            await new Promise((resolve, reject) => {
                reader.onload = async (e) => {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { 
                            type: 'array', 
                            cellDates: true,
                            dateNF: 'yyyy-mm-dd'
                        });
                        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        const settings = getCurrentSettings();
                        const cleanedRows = await cleanData(rows, settings);
                        
                        if (settings.dupes) {
                            const { dedupedRows } = removeDuplicates(cleanedRows);
                            AppState.cleanedData = dedupedRows;
                        } else {
                            AppState.cleanedData = cleanedRows;
                        }
                        
                        await exportFile(AppState.cleanedData, file.name);
                        
                        batchResults.success++;
                        batchResults.files.push({
                            name: file.name,
                            status: 'success',
                            rows: cleanedRows.length
                        });
                        
                        addLog('Batch Complete', `Processed: ${file.name}`, 'success');
                        resolve();
                    } catch (error) {
                        reject(error);
                    }
                };
                
                reader.onerror = () => reject(new Error(`Error reading ${file.name}`));
                reader.readAsArrayBuffer(file);
            });
        } catch (error) {
            batchResults.failed++;
            batchResults.files.push({
                name: file.name,
                status: 'failed',
                error: error.message
            });
            addLog('Batch Error', `Failed to process ${file.name}: ${error.message}`, 'error');
        }
    }
    
    updateProgress(100, 'Batch processing complete!');
    
    // Show batch summary
    addLog('Batch Summary', `${batchResults.success} files processed, ${batchResults.failed} failed`, 
           batchResults.failed === 0 ? 'success' : 'warning');
    
    if (batchResults.failed > 0) {
        const failedList = batchResults.files.filter(f => f.status === 'failed')
            .map(f => f.name).join(', ');
        addLog('Failed Files', failedList, 'warning');
    }
}

// Contact Form Handler
document.addEventListener('DOMContentLoaded', function() {
    const contactForm = document.getElementById('contactForm');
    const formStatus = document.getElementById('formStatus');
    const formMessage = document.getElementById('formMessage');
    
    if (contactForm) {
        contactForm.addEventListener('submit', function(e) {
            e.preventDefault();
            
            // Get form values
            const name = document.getElementById('contactName').value.trim();
            const email = document.getElementById('contactEmail').value.trim();
            const message = document.getElementById('contactMessage').value.trim();
            
            // Simple validation
            if (!name || !email || !message) {
                showFormStatus('Please fill in all required fields.', 'error');
                return;
            }
            
            if (!DataPatterns.EMAIL.test(email)) {
                showFormStatus('Please enter a valid email address.', 'error');
                return;
            }
            
            if (!document.getElementById('contactConsent').checked) {
                showFormStatus('Please check the consent box.', 'error');
                return;
            }
            
            // Show loading state
            const submitBtn = contactForm.querySelector('button[type="submit"]');
            const originalText = submitBtn.innerHTML;
            submitBtn.innerHTML = '<span class="spinner"></span> Sending...';
            submitBtn.disabled = true;
            
            // Simulate form submission
            setTimeout(() => {
                console.log('Contact Form Submission:', { name, email, message });
                
                // Show success message
                showFormStatus('Thank you for your message! We\'ll get back to you soon.', 'success');
                
                // Reset form
                contactForm.reset();
                
                // Restore button
                submitBtn.innerHTML = originalText;
                submitBtn.disabled = false;
                
                // Auto-close modal after 3 seconds
                setTimeout(() => {
                    const modal = bootstrap.Modal.getInstance(document.getElementById('contactModal'));
                    if (modal) modal.hide();
                }, 3000);
                
            }, 1500);
        });
    }
    
    function showFormStatus(message, type) {
        if (!formStatus || !formMessage) return;
        
        formMessage.textContent = message;
        formStatus.className = 'alert';
        
        switch(type) {
            case 'success':
                formStatus.classList.add('alert-success');
                break;
            case 'error':
                formStatus.classList.add('alert-danger');
                break;
            case 'warning':
                formStatus.classList.add('alert-warning');
                break;
            default:
                formStatus.classList.add('alert-info');
        }
        
        formStatus.classList.remove('d-none');
        
        // Auto-hide after 5 seconds
        setTimeout(() => {
            formStatus.classList.add('d-none');
        }, 5000);
    }
    
    // Clear form status when modal is hidden
    const contactModal = document.getElementById('contactModal');
    if (contactModal) {
        contactModal.addEventListener('hidden.bs.modal', function() {
            if (formStatus) {
                formStatus.classList.add('d-none');
            }
            if (contactForm) {
                contactForm.reset();
            }
        });
    }
});

// Initialize when DOM is loaded
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}

// Make some functions globally available for onclick handlers
window.clearFile = clearFile;
window.removeFile = removeFile;
window.showPreview = showPreview;
window.undo = undo;
window.redo = redo;
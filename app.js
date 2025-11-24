'use strict';

/**
 * Excel File Comparison Tool
 * Version: 4.0.0
 * Author: Tarang Srivastava
 * Description: Compare two Excel files and generate detailed difference reports
 */

// Configuration Constants
const CONFIG = {
    APP_VERSION: '4.0.0',
    MAX_FILE_SIZE: 52428800, // 50MB in bytes
    NUMERIC_EPSILON: 0.000001,
    VALID_FILE_TYPES: [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'application/octet-stream'
    ],
    VALID_EXTENSIONS: ['.xlsx', '.xls'],
    DEBUG_MODE: false
};

// Utility Functions
const Logger = {
    log: (...args) => {
        if (CONFIG.DEBUG_MODE) {
            console.log(...args);
        }
    },
    error: (...args) => {
        console.error(...args);
    },
    warn: (...args) => {
        if (CONFIG.DEBUG_MODE) {
            console.warn(...args);
        }
    },
    info: (...args) => {
        if (CONFIG.DEBUG_MODE) {
            console.info(...args);
        }
    }
};

const Utils = {
    sanitizeHTML: (str) => {
        const temp = document.createElement('div');
        temp.textContent = str;
        return temp.innerHTML;
    },
    
    sanitizeFilename: (filename) => {
        return filename.replace(/[^a-z0-9_\-]/gi, '_').substring(0, 255);
    },
    
    formatFileSize: (bytes) => {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    },
    
    formatNumber: (num) => {
        const rounded = Number(num.toFixed(2));
        const parts = rounded.toString().split('.');
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        return parts.join('.');
    },
    
    getCurrentDateString: () => {
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }
};

class ExcelComparator {
    constructor() {
        this.originalFile = null;
        this.modifiedFile = null;
        this.originalData = null;
        this.modifiedData = null;
        this.comparisonResult = null;
        this.differenceMap = new Map();
        this.reportFormat = 'single';
        
        this.initializeEventListeners();
        this.setDefaultFilename();
        this.announceToScreenReader('Excel File Comparison Tool loaded and ready');
        this.logAppVersion();
    }

    logAppVersion() {
        Logger.log(`%c Excel File Comparison Tool v${CONFIG.APP_VERSION} `, 
            'background: #217346; color: white; font-weight: bold; padding: 5px 10px;');
        Logger.log('Ready to compare Excel files!');
    }

    announceToScreenReader(message) {
        const announcement = document.createElement('div');
        announcement.setAttribute('role', 'status');
        announcement.setAttribute('aria-live', 'polite');
        announcement.className = 'sr-only';
        announcement.textContent = message;
        document.body.appendChild(announcement);
        setTimeout(() => announcement.remove(), 1000);
    }

    initializeEventListeners() {
        const originalFileInput = document.getElementById('originalFile');
        const modifiedFileInput = document.getElementById('modifiedFile');
        const compareBtn = document.getElementById('compareBtn');
        const downloadBtn = document.getElementById('downloadBtn');
        
        if (originalFileInput) {
            originalFileInput.addEventListener('change', (e) => {
                this.handleFileUpload(e, 'original');
            });
        }
        
        if (modifiedFileInput) {
            modifiedFileInput.addEventListener('change', (e) => {
                this.handleFileUpload(e, 'modified');
            });
        }

        if (compareBtn) {
            compareBtn.addEventListener('click', () => {
                this.performComparison();
            });
        }

        if (downloadBtn) {
            downloadBtn.addEventListener('click', () => {
                this.downloadReport();
            });
        }

        document.querySelectorAll('input[name="reportFormat"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                this.reportFormat = e.target.value;
                Logger.log(`Report format changed to: ${this.reportFormat}`);
            });
        });

        this.setupDragAndDrop();
        this.setupKeyboardNavigation();
    }

    setupKeyboardNavigation() {
        document.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                const compareBtn = document.getElementById('compareBtn');
                if (compareBtn && !compareBtn.disabled) {
                    e.preventDefault();
                    compareBtn.click();
                }
            }
            
            if ((e.ctrlKey || e.metaKey) && e.key === 'd') {
                const downloadBtn = document.getElementById('downloadBtn');
                if (downloadBtn && !downloadBtn.classList.contains('hidden')) {
                    e.preventDefault();
                    downloadBtn.click();
                }
            }
        });
    }

    setDefaultFilename() {
        const filenameInput = document.getElementById('outputFilename');
        if (filenameInput) {
            const dateStr = Utils.getCurrentDateString();
            filenameInput.value = `comparison_report_${dateStr}`;
        }
    }

    setupDragAndDrop() {
        const uploadSection = document.querySelector('.form-section');
        if (!uploadSection) return;
        
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            uploadSection.addEventListener(eventName, this.preventDefaults, false);
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            uploadSection.addEventListener(eventName, () => {
                uploadSection.classList.add('dragover');
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            uploadSection.addEventListener(eventName, () => {
                uploadSection.classList.remove('dragover');
            }, false);
        });

        uploadSection.addEventListener('drop', this.handleDrop.bind(this), false);
    }

    preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    handleDrop(e) {
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const originalFileInput = document.getElementById('originalFile');
            const modifiedFileInput = document.getElementById('modifiedFile');
            
            if (!this.originalFile && originalFileInput) {
                const dt = new DataTransfer();
                dt.items.add(files[0]);
                originalFileInput.files = dt.files;
                this.handleFileUpload({target: {files: [files[0]]}}, 'original');
            } else if (!this.modifiedFile && files.length > 0 && modifiedFileInput) {
                const dt = new DataTransfer();
                dt.items.add(files[0]);
                modifiedFileInput.files = dt.files;
                this.handleFileUpload({target: {files: [files[0]]}}, 'modified');
            }
        }
    }

    async handleFileUpload(event, fileType) {
        const file = event.target.files[0];
        
        if (!file) {
            this.clearFileStatus(fileType);
            return;
        }

        const validation = this.validateFile(file);
        if (!validation.isValid) {
            this.showFileStatus(fileType, validation.message, 'error');
            this.announceToScreenReader(`Error: ${validation.message}`);
            return;
        }

        try {
            this.showFileStatus(fileType, 'Loading file...', 'warning');
            
            const data = await this.readExcelFile(file);
            
            if (fileType === 'original') {
                this.originalFile = file;
                this.originalData = data;
            } else {
                this.modifiedFile = file;
                this.modifiedData = data;
            }

            const successMessage = `Loaded: ${file.name} (${Utils.formatFileSize(file.size)}, ${data.SheetNames.length} sheets)`;
            this.showFileStatus(fileType, successMessage, 'success');
            this.announceToScreenReader(`${fileType} file loaded successfully`);

            this.updateCompareButton();

        } catch (error) {
            this.showFileStatus(fileType, `Error loading file: ${error.message}`, 'error');
            this.announceToScreenReader(`Error loading ${fileType} file`);
            Logger.error('File loading error:', error);
        }
    }

    validateFile(file) {
        const hasValidExtension = CONFIG.VALID_EXTENSIONS.some(ext => 
            file.name.toLowerCase().endsWith(ext)
        );

        if (!CONFIG.VALID_FILE_TYPES.includes(file.type) && !hasValidExtension) {
            return {
                isValid: false,
                message: 'Invalid file type. Please select an Excel file (.xlsx or .xls)'
            };
        }

        if (file.size > CONFIG.MAX_FILE_SIZE) {
            return {
                isValid: false,
                message: `File too large. Maximum size is 50MB (current: ${Utils.formatFileSize(file.size)})`
            };
        }

        return { isValid: true };
    }

    async readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    
                    if (this.isEncryptedFile(data)) {
                        reject(new Error('Encrypted files are not supported for comparison.\nPlease copy the content into a new blank file, save it, and try again.'));
                        return;
                    }
                    
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                        reject(new Error('Encrypted files are not supported for comparison.\nPlease copy the content into a new blank file, save it, and try again.'));
                        return;
                    }
                    
                    resolve(workbook);
                } catch (error) {
                    if (error.message && (
                        error.message.toLowerCase().includes('password') ||
                        error.message.toLowerCase().includes('encrypted') ||
                        error.message.toLowerCase().includes('protected') ||
                        error.message.toLowerCase().includes('corrupt') ||
                        error.message.includes('Unsupported file')
                    )) {
                        reject(new Error('Encrypted files are not supported for comparison.\nPlease copy the content into a new blank file, save it, and try again.'));
                    } else {
                        reject(new Error(`Failed to parse Excel file: ${error.message}`));
                    }
                }
            };
            
            reader.onerror = () => {
                reject(new Error('Failed to read file'));
            };
            
            reader.readAsArrayBuffer(file);
        });
    }

    isEncryptedFile(data) {
        try {
            if (data.length >= 8) {
                const oleSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
                let isOLE = true;
                for (let i = 0; i < 8; i++) {
                    if (data[i] !== oleSignature[i]) {
                        isOLE = false;
                        break;
                    }
                }
                
                if (isOLE) {
                    const dataStr = Array.from(data.slice(0, Math.min(1024, data.length)))
                        .map(byte => String.fromCharCode(byte))
                        .join('');
                    
                    if (dataStr.includes('EncryptionInfo') || 
                        dataStr.includes('EncryptedPackage') ||
                        dataStr.includes('Microsoft.Container.EncryptionTransform')) {
                        return true;
                    }
                }
            }
            
            if (data.length >= 4) {
                const zipSignature = [0x50, 0x4B, 0x03, 0x04];
                let isZip = true;
                for (let i = 0; i < 4; i++) {
                    if (data[i] !== zipSignature[i]) {
                        isZip = false;
                        break;
                    }
                }
                
                if (isZip) {
                    const dataStr = Array.from(data.slice(0, Math.min(2048, data.length)))
                        .map(byte => String.fromCharCode(byte))
                        .join('');
                    
                    if (dataStr.includes('EncryptionInfo') || 
                        dataStr.includes('EncryptedPackage') ||
                        !dataStr.includes('xl/workbook.xml')) {
                        return true;
                    }
                }
            }
            
            return false;
        } catch (error) {
            return false;
        }
    }

    showFileStatus(fileType, message, statusType) {
        const statusElement = document.getElementById(`${fileType === 'original' ? 'original' : 'modified'}-status`);
        if (!statusElement) return;
        
        statusElement.textContent = message;
        statusElement.className = `file-status status-${statusType}`;
        statusElement.classList.remove('hidden');
    }

    clearFileStatus(fileType) {
        const statusElement = document.getElementById(`${fileType === 'original' ? 'original' : 'modified'}-status`);
        if (statusElement) {
            statusElement.classList.add('hidden');
        }
        
        if (fileType === 'original') {
            this.originalFile = null;
            this.originalData = null;
        } else {
            this.modifiedFile = null;
            this.modifiedData = null;
        }
        
        this.updateCompareButton();
    }

    updateCompareButton() {
        const compareBtn = document.getElementById('compareBtn');
        if (!compareBtn) return;
        
        const isEnabled = !!(this.originalFile && this.modifiedFile);
        compareBtn.disabled = !isEnabled;
        compareBtn.setAttribute('aria-disabled', !isEnabled);
        
        if (isEnabled) {
            this.announceToScreenReader('Both files loaded. Ready to compare.');
        }
    }

    async performComparison() {
        try {
            this.showProgress(true);
            this.updateProgress(10, 'Validating file structures...');
            this.announceToScreenReader('Starting comparison');
            
            this.hideElement('resultsSection');
            this.hideElement('errorSection');
            this.hideElement('downloadBtn');
            this.differenceMap.clear();

            const structureValidation = this.validateStructures();
            if (!structureValidation.isValid) {
                throw new Error(structureValidation.message);
            }

            this.updateProgress(30, 'Comparing file contents...');
            
            await this.compareFiles();
            
            this.updateProgress(80, 'Generating report...');
            
            this.generateReport();
            
            this.updateProgress(100, 'Comparison completed!');
            
            setTimeout(() => {
                this.showProgress(false);
                this.showResults();
                this.announceToScreenReader('Comparison completed successfully');
            }, 1000);

        } catch (error) {
            this.showProgress(false);
            this.showError(error.message);
            this.announceToScreenReader('Comparison failed. Please check the error message.');
            Logger.error('Comparison error:', error);
        }
    }

    validateStructures() {
        const originalSheets = this.originalData.SheetNames;
        const modifiedSheets = this.modifiedData.SheetNames;

        if (originalSheets.length !== modifiedSheets.length) {
            return {
                isValid: false,
                message: `Different number of sheets: Original has ${originalSheets.length}, Modified has ${modifiedSheets.length}`
            };
        }

        for (let i = 0; i < originalSheets.length; i++) {
            if (originalSheets[i] !== modifiedSheets[i]) {
                return {
                    isValid: false,
                    message: `Sheet names don't match: "${originalSheets[i]}" vs "${modifiedSheets[i]}"`
                };
            }
        }

        return { isValid: true };
    }

    async compareFiles() {
        const differences = [];
        const sheetNames = this.originalData.SheetNames;

        for (const sheetName of sheetNames) {
            const originalSheet = XLSX.utils.sheet_to_json(
                this.originalData.Sheets[sheetName], 
                { header: 1, defval: '' }
            );
            
            const modifiedSheet = XLSX.utils.sheet_to_json(
                this.modifiedData.Sheets[sheetName], 
                { header: 1, defval: '' }
            );

            const maxRows = Math.max(originalSheet.length, modifiedSheet.length);
            const maxCols = Math.max(
                Math.max(...originalSheet.map(row => row.length)),
                Math.max(...modifiedSheet.map(row => row.length))
            );

            const sheetDifferences = new Set();

            for (let row = 0; row < maxRows; row++) {
                for (let col = 0; col < maxCols; col++) {
                    const originalValue = (originalSheet[row] && originalSheet[row][col]) || '';
                    const modifiedValue = (modifiedSheet[row] && modifiedSheet[row][col]) || '';

                    if (this.valuesAreDifferent(originalValue, modifiedValue)) {
                        const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
                        sheetDifferences.add(`${row}-${col}`);
                        
                        differences.push({
                            sheet: sheetName,
                            row: row + 1,
                            col: col + 1,
                            cellRef: cellRef,
                            originalValue: originalValue,
                            modifiedValue: modifiedValue
                        });
                    }
                }
            }

            this.differenceMap.set(sheetName, sheetDifferences);
        }

        this.comparisonResult = {
            differences: differences,
            totalDifferences: differences.length,
            sheetsCompared: sheetNames.length,
            originalFile: this.originalFile.name,
            modifiedFile: this.modifiedFile.name,
            comparisonDate: new Date().toISOString()
        };
    }

    valuesAreDifferent(originalValue, modifiedValue) {
        if (originalValue === modifiedValue) {
            return false;
        }

        const originalNumeric = this.getNumericValue(originalValue, false);
        const modifiedNumeric = this.getNumericValue(modifiedValue, false);

        if (originalNumeric !== null && modifiedNumeric !== null) {
            return Math.abs(originalNumeric - modifiedNumeric) > CONFIG.NUMERIC_EPSILON;
        }

        const originalPercent = this.extractPercentageValue(originalValue);
        const modifiedPercent = this.extractPercentageValue(modifiedValue);

        if (originalPercent !== null && modifiedPercent !== null) {
            return Math.abs(originalPercent - modifiedPercent) > CONFIG.NUMERIC_EPSILON;
        }

        return String(originalValue).trim() !== String(modifiedValue).trim();
    }

    extractPercentageValue(value) {
        if (typeof value === 'number' && value >= 0 && value <= 1) {
            return value * 100;
        }

        if (typeof value === 'string' && value.trim().endsWith('%')) {
            const percentStr = value.trim().slice(0, -1);
            if (!isNaN(percentStr)) {
                return parseFloat(percentStr);
            }
        }

        return null;
    }

    generateReport() {
        const reportWorkbook = XLSX.utils.book_new();
        this.createComparisonSheets(reportWorkbook);
        this.reportWorkbook = reportWorkbook;
    }

    createComparisonSheets(workbook) {
        this.originalData.SheetNames.forEach(sheetName => {
            const originalSheet = this.originalData.Sheets[sheetName];
            const modifiedSheet = this.modifiedData.Sheets[sheetName];
            
            const originalData = XLSX.utils.sheet_to_json(originalSheet, { header: 1, defval: '' });
            const modifiedData = XLSX.utils.sheet_to_json(modifiedSheet, { header: 1, defval: '' });
            
            const comparisonData = this.reportFormat === 'single' 
                ? this.createComparisonDataSingleCell(originalData, modifiedData, sheetName)
                : this.createComparisonDataMultiCell(originalData, modifiedData, sheetName);
            
            const comparisonSheet = XLSX.utils.aoa_to_sheet(comparisonData);
            
            XLSX.utils.book_append_sheet(workbook, comparisonSheet, `${sheetName}_Comparison`);
        });
    }

    createComparisonDataSingleCell(originalData, modifiedData, sheetName) {
        const maxRows = Math.max(originalData.length, modifiedData.length);
        const maxCols = Math.max(
            Math.max(...originalData.map(row => row.length)),
            Math.max(...modifiedData.map(row => row.length))
        );

        const comparisonData = [];
        const sheetDifferences = this.differenceMap.get(sheetName) || new Set();

        comparisonData.push(['COMPARISON LEGEND']);
        comparisonData.push(['Red Circle = Field name where values have changed']);
        comparisonData.push(['No symbol = Field name where values are identical']);
        comparisonData.push(['Format: Single Cell (Original Value, Changed Value, Difference)']);
        comparisonData.push(['']);
        comparisonData.push(['Field Name', 'Value Comparison']);
        comparisonData.push(['']);

        for (let row = 0; row < maxRows; row++) {
            const comparisonRow = [];
            
            let hasAnyDifference = false;
            for (let col = 1; col < maxCols; col++) {
                const originalValue = (originalData[row] && originalData[row][col]) || '';
                const modifiedValue = (modifiedData[row] && modifiedData[row][col]) || '';
                
                if (this.valuesAreDifferent(originalValue, modifiedValue)) {
                    hasAnyDifference = true;
                    break;
                }
            }
            
            for (let col = 0; col < maxCols; col++) {
                const originalValue = (originalData[row] && originalData[row][col]) || '';
                const modifiedValue = (modifiedData[row] && modifiedData[row][col]) || '';

                if (col === 0) {
                    if (hasAnyDifference) {
                        comparisonRow.push(`[CHANGED] ${originalValue}`);
                    } else {
                        comparisonRow.push(originalValue);
                    }
                } else {
                    const originalCellRef = XLSX.utils.encode_cell({ r: row, c: col });
                    const modifiedCellRef = XLSX.utils.encode_cell({ r: row, c: col });
                    const originalCellInfo = this.originalData.Sheets[sheetName][originalCellRef];
                    const modifiedCellInfo = this.modifiedData.Sheets[sheetName][modifiedCellRef];
                    
                    const formattedComparison = this.createValueComparisonSingleCell(
                        originalValue, modifiedValue, 
                        originalCellInfo, modifiedCellInfo, 
                        sheetName, originalCellRef
                    );
                    comparisonRow.push(formattedComparison);
                }
            }
            
            comparisonData.push(comparisonRow);
        }

        return comparisonData;
    }

    createComparisonDataMultiCell(originalData, modifiedData, sheetName) {
        const maxRows = Math.max(originalData.length, modifiedData.length);
        const maxCols = Math.max(
            Math.max(...originalData.map(row => row.length)),
            Math.max(...modifiedData.map(row => row.length))
        );

        const comparisonData = [];
        const sheetDifferences = this.differenceMap.get(sheetName) || new Set();

        comparisonData.push(['COMPARISON LEGEND']);
        comparisonData.push(['Red Circle = Field name where values have changed']);
        comparisonData.push(['No symbol = Field name where values are identical']);
        comparisonData.push(['Format: Multi Cell (Separate columns for Original, Changed, Difference)']);
        comparisonData.push(['']);

        const headerRow = ['Field Name'];
        for (let col = 1; col < maxCols; col++) {
            headerRow.push('Original Value', 'Changed Value', 'Difference');
        }
        comparisonData.push(headerRow);
        comparisonData.push(['']);

        for (let row = 0; row < maxRows; row++) {
            const comparisonRow = [];
            
            let hasAnyDifference = false;
            for (let col = 1; col < maxCols; col++) {
                const originalValue = (originalData[row] && originalData[row][col]) || '';
                const modifiedValue = (modifiedData[row] && modifiedData[row][col]) || '';
                
                if (this.valuesAreDifferent(originalValue, modifiedValue)) {
                    hasAnyDifference = true;
                    break;
                }
            }
            
            for (let col = 0; col < maxCols; col++) {
                const originalValue = (originalData[row] && originalData[row][col]) || '';
                const modifiedValue = (modifiedData[row] && modifiedData[row][col]) || '';

                if (col === 0) {
                    if (hasAnyDifference) {
                        comparisonRow.push(`[CHANGED] ${originalValue}`);
                    } else {
                        comparisonRow.push(originalValue);
                    }
                } else {
                    const originalCellRef = XLSX.utils.encode_cell({ r: row, c: col });
                    const modifiedCellRef = XLSX.utils.encode_cell({ r: row, c: col });
                    const originalCellInfo = this.originalData.Sheets[sheetName][originalCellRef];
                    const modifiedCellInfo = this.modifiedData.Sheets[sheetName][modifiedCellRef];
                    
                    const { original, changed, difference } = this.createValueComparisonMultiCell(
                        originalValue, modifiedValue, 
                        originalCellInfo, modifiedCellInfo, 
                        sheetName, originalCellRef
                    );
                    
                    comparisonRow.push(original, changed, difference);
                }
            }
            
            comparisonData.push(comparisonRow);
        }

        return comparisonData;
    }

    createValueComparisonSingleCell(originalValue, modifiedValue, originalCellInfo, modifiedCellInfo, sheetName, cellRef) {
        const isOriginalBlank = originalValue === '' || originalValue === null || originalValue === undefined;
        const isModifiedBlank = modifiedValue === '' || modifiedValue === null || modifiedValue === undefined;
        
        if (isOriginalBlank && isModifiedBlank) {
            return '';
        }
        
        const formattedOriginal = this.formatValue(originalValue, originalCellInfo, sheetName, cellRef);
        const formattedModified = this.formatValue(modifiedValue, modifiedCellInfo, sheetName, cellRef);
        
        const originalIsPercentage = this.isPercentageCell(originalCellInfo, sheetName, cellRef, originalValue) || 
                                   (typeof formattedOriginal === 'string' && formattedOriginal.includes('%'));
        const modifiedIsPercentage = this.isPercentageCell(modifiedCellInfo, sheetName, cellRef, modifiedValue) || 
                                   (typeof formattedModified === 'string' && formattedModified.includes('%'));
        
        const originalNumeric = this.getNumericValue(originalValue, originalIsPercentage);
        const modifiedNumeric = this.getNumericValue(modifiedValue, modifiedIsPercentage);
        
        let result = `Original Value: ${formattedOriginal}, Changed Value: ${formattedModified}`;
        
        if (originalNumeric !== null && modifiedNumeric !== null) {
            if (originalIsPercentage && modifiedIsPercentage) {
                const difference = modifiedNumeric - originalNumeric;
                const sign = difference > 0 ? '+' : (difference < 0 ? '' : '');
                const formattedDiff = Utils.formatNumber(Math.abs(difference));
                
                if (Math.abs(difference) > CONFIG.NUMERIC_EPSILON) {
                    result += `, Difference: ${sign}${formattedDiff}%`;
                } else {
                    result += `, Difference: 0%`;
                }
            } else if (!originalIsPercentage && !modifiedIsPercentage) {
                const difference = modifiedNumeric - originalNumeric;
                const sign = difference > 0 ? '+' : (difference < 0 ? '' : '');
                const formattedDiff = Utils.formatNumber(Math.abs(difference));
                
                if (Math.abs(difference) > CONFIG.NUMERIC_EPSILON) {
                    result += `, Difference: ${sign}${formattedDiff}`;
                } else {
                    result += `, Difference: 0`;
                }
            } else {
                result += `, Difference: N/A (type mismatch)`;
            }
        } else {
            result += `, Difference: N/A (non-numeric)`;
        }
        
        return result;
    }

    createValueComparisonMultiCell(originalValue, modifiedValue, originalCellInfo, modifiedCellInfo, sheetName, cellRef) {
        const isOriginalBlank = originalValue === '' || originalValue === null || originalValue === undefined;
        const isModifiedBlank = modifiedValue === '' || modifiedValue === null || modifiedValue === undefined;
        
        if (isOriginalBlank && isModifiedBlank) {
            return { original: '', changed: '', difference: '' };
        }
        
        const formattedOriginal = this.formatValue(originalValue, originalCellInfo, sheetName, cellRef);
        const formattedModified = this.formatValue(modifiedValue, modifiedCellInfo, sheetName, cellRef);
        
        const originalIsPercentage = this.isPercentageCell(originalCellInfo, sheetName, cellRef, originalValue) || 
                                   (typeof formattedOriginal === 'string' && formattedOriginal.includes('%'));
        const modifiedIsPercentage = this.isPercentageCell(modifiedCellInfo, sheetName, cellRef, modifiedValue) || 
                                   (typeof formattedModified === 'string' && formattedModified.includes('%'));
        
        const originalNumeric = this.getNumericValue(originalValue, originalIsPercentage);
        const modifiedNumeric = this.getNumericValue(modifiedValue, modifiedIsPercentage);
        
        let differenceText = '';
        
        if (originalNumeric !== null && modifiedNumeric !== null) {
            if (originalIsPercentage && modifiedIsPercentage) {
                const difference = modifiedNumeric - originalNumeric;
                const sign = difference > 0 ? '+' : (difference < 0 ? '' : '');
                const formattedDiff = Utils.formatNumber(Math.abs(difference));
                
                if (Math.abs(difference) > CONFIG.NUMERIC_EPSILON) {
                    differenceText = `${sign}${formattedDiff}%`;
                } else {
                    differenceText = '0%';
                }
            } else if (!originalIsPercentage && !modifiedIsPercentage) {
                const difference = modifiedNumeric - originalNumeric;
                const sign = difference > 0 ? '+' : (difference < 0 ? '' : '');
                const formattedDiff = Utils.formatNumber(Math.abs(difference));
                
                if (Math.abs(difference) > CONFIG.NUMERIC_EPSILON) {
                    differenceText = `${sign}${formattedDiff}`;
                } else {
                    differenceText = '0';
                }
            } else {
                differenceText = 'N/A (type mismatch)';
            }
        } else {
            differenceText = 'N/A (non-numeric)';
        }
        
        return {
            original: formattedOriginal,
            changed: formattedModified,
            difference: differenceText
        };
    }

    getNumericValue(value, isPercentage) {
        if (typeof value === 'number') {
            return isPercentage ? value * 100 : value;
        }
        
        if (typeof value === 'string' && value.trim() !== '') {
            const trimmedValue = value.trim();
            
            if (trimmedValue.endsWith('%')) {
                const percentValue = trimmedValue.slice(0, -1);
                if (!isNaN(percentValue)) {
                    return parseFloat(percentValue);
                }
            }
            
            const cleanValue = trimmedValue.replace(/,/g, '');
            if (!isNaN(cleanValue)) {
                return parseFloat(cleanValue);
            }
        }
        
        return null;
    }

    formatValue(value, originalCell = null, sheetName = null, cellRef = null) {
        if (this.isPercentageCell(originalCell, sheetName, cellRef, value)) {
            const percentValue = Math.abs(Number((value * 100).toFixed(2)));
            const sign = value < 0 ? '-' : '';
            return sign + Utils.formatNumber(percentValue) + '%';
        }
        
        if (typeof value === 'number') {
            return Utils.formatNumber(value);
        }
        
        if (typeof value === 'string' && value.trim() !== '') {
            const trimmedValue = value.trim();
            
            if (trimmedValue.endsWith('%')) {
                const percentValue = trimmedValue.slice(0, -1);
                if (!isNaN(percentValue)) {
                    const numValue = parseFloat(percentValue);
                    return Utils.formatNumber(Math.abs(numValue)) + '%';
                }
            }
            
            if (!isNaN(trimmedValue)) {
                const numValue = parseFloat(trimmedValue);
                return Utils.formatNumber(numValue);
            }
        }
        
        return value;
    }

    isPercentageCell(originalCell, sheetName, cellRef, value) {
        if (originalCell && originalCell.z) {
            const formatCode = originalCell.z.toLowerCase();
            if (formatCode.includes('%') || formatCode.includes('percent')) {
                return true;
            }
        }
        
        if (sheetName && cellRef && this.originalData) {
            try {
                const cellAddress = XLSX.utils.decode_cell(cellRef);
                const headerRow = 0;
                const headerCellRef = XLSX.utils.encode_cell({ r: headerRow, c: cellAddress.c });
                const headerCell = this.originalData.Sheets[sheetName][headerCellRef];
                
                if (headerCell && headerCell.v) {
                    const headerText = headerCell.v.toString().toLowerCase();
                    if (headerText.includes('%') || headerText.includes('percent') || 
                        headerText.includes('rate') || headerText.includes('ratio')) {
                        return true;
                    }
                }
            } catch (error) {
                Logger.warn('Error checking percentage cell:', error);
            }
        }
        
        if (typeof value === 'number' && Math.abs(value) <= 1 && Math.abs(value) > 0) {
            const valueStr = value.toString();
            const decimalPart = valueStr.split('.')[1];
            if (decimalPart && decimalPart.length >= 2) {
                const absValue = Math.abs(value);
                if (absValue >= 0.001 && absValue <= 1) {
                    return true;
                }
            }
        }
        
        return false;
    }

    getFormattedFilename() {
        const filenameInput = document.getElementById('outputFilename');
        const customFilename = filenameInput ? filenameInput.value.trim() : '';
        
        if (customFilename) {
            return Utils.sanitizeFilename(customFilename);
        }
        
        const dateStr = Utils.getCurrentDateString();
        return `comparison_report_${dateStr}`;
    }

    downloadReport() {
        try {
            const filename = this.getFormattedFilename();
            const wbout = XLSX.write(this.reportWorkbook, { bookType: 'xlsx', type: 'binary' });
            
            const blob = new Blob([this.s2ab(wbout)], { type: 'application/octet-stream' });
            
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${filename}.xlsx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            this.showStatus('Report downloaded successfully!', 'success');
            this.announceToScreenReader('Report downloaded successfully');
        } catch (error) {
            this.showError(`Download failed: ${error.message}`);
            this.announceToScreenReader('Download failed');
            Logger.error('Download error:', error);
        }
    }

    s2ab(s) {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) {
            view[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf;
    }

    showResults() {
        const resultsSection = document.getElementById('resultsSection');
        if (!resultsSection) return;
        
        const sanitizedOriginalFile = Utils.sanitizeHTML(this.originalFile.name);
        const sanitizedModifiedFile = Utils.sanitizeHTML(this.modifiedFile.name);
        
        resultsSection.innerHTML = `
            <h3 id="results-heading">
                <span aria-hidden="true">&#128202;</span>
                Comparison Results
            </h3>
            <div class="results-grid">
                <div class="stat-card">
                    <div class="stat-value">${Utils.formatNumber(this.comparisonResult.totalDifferences)}</div>
                    <div class="stat-label">Total Differences</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">${Utils.formatNumber(this.comparisonResult.sheetsCompared)}</div>
                    <div class="stat-label">Sheets Compared</div>
                </div>
                <div class="stat-card filename-card">
                    <div class="stat-value" title="${sanitizedOriginalFile}">${sanitizedOriginalFile}</div>
                    <div class="stat-label">Original File</div>
                </div>
                <div class="stat-card filename-card">
                    <div class="stat-value" title="${sanitizedModifiedFile}">${sanitizedModifiedFile}</div>
                    <div class="stat-label">Modified File</div>
                </div>
            </div>
        `;
        
        resultsSection.classList.remove('hidden');
        
        const downloadBtn = document.getElementById('downloadBtn');
        if (downloadBtn) {
            downloadBtn.classList.remove('hidden');
        }
    }

    showProgress(show) {
        const progressSection = document.getElementById('progressSection');
        if (!progressSection) return;
        
        if (show) {
            progressSection.classList.remove('hidden');
        } else {
            progressSection.classList.add('hidden');
        }
    }

    updateProgress(percentage, text) {
        const progressFill = document.getElementById('progressFill');
        const progressText = document.getElementById('progressText');
        
        if (progressFill) {
            const progressBar = progressFill.parentElement;
            progressFill.style.width = `${percentage}%`;
            if (progressBar) {
                progressBar.setAttribute('aria-valuenow', percentage);
            }
        }
        
        if (progressText) {
            progressText.textContent = text;
        }
    }

    showStatus(message, type) {
        const statusElement = document.getElementById('statusMessage');
        if (!statusElement) return;
        
        statusElement.textContent = message;
        statusElement.className = `status-message status-${type}`;
        statusElement.classList.remove('hidden');
        
        setTimeout(() => {
            statusElement.classList.add('hidden');
        }, 5000);
    }

    showError(message) {
        const errorSection = document.getElementById('errorSection');
        if (!errorSection) return;
        
        const sanitizedMessage = Utils.sanitizeHTML(message);
        const formattedMessage = sanitizedMessage.replace(/\n/g, '<br>');
        const isEncryptionError = message.includes('Encrypted files are not supported');
        
        if (isEncryptionError) {
            errorSection.innerHTML = `
                <h4><span aria-hidden="true">&#128274;</span> Encrypted File Detected</h4>
                <p>${formattedMessage}</p>
                <p><strong>How to resolve this:</strong></p>
                <ol>
                    <li>Open the encrypted file in Excel</li>
                    <li>Enter your password when prompted</li>
                    <li>Select all content (Ctrl+A)</li>
                    <li>Copy the content (Ctrl+C)</li>
                    <li>Create a new blank Excel file</li>
                    <li>Paste the content (Ctrl+V)</li>
                    <li>Save the new file without encryption</li>
                    <li>Use the new unencrypted file for comparison</li>
                </ol>
            `;
        } else {
            errorSection.innerHTML = `
                <h4><span aria-hidden="true">&#10060;</span> Error</h4>
                <p>${formattedMessage}</p>
                <p><strong>Please check:</strong></p>
                <ul>
                    <li>Both files are valid Excel files (.xlsx or .xls)</li>
                    <li>Files have identical structure (same sheets, columns, and row count)</li>
                    <li>Files are not corrupted or password protected</li>
                    <li>File sizes are under 50MB</li>
                </ul>
            `;
        }
        
        errorSection.classList.remove('hidden');
    }

    hideElement(elementId) {
        const element = document.getElementById(elementId);
        if (element) {
            element.classList.add('hidden');
        }
    }
}

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    if (typeof XLSX === 'undefined') {
        showCriticalError('library');
        return;
    }

    if (!window.FileReader) {
        showCriticalError('browser');
        return;
    }

    try {
        new ExcelComparator();
        Logger.log(`Excel File Comparison Tool v${CONFIG.APP_VERSION} initialized successfully`);
    } catch (error) {
        Logger.error('Initialization error:', error);
        showCriticalError('general');
    }
});
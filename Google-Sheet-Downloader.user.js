// ==UserScript==
// @name         Google Sheets Export Bypass
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Extract data from Google Sheets using export API endpoints for sheets which have download/copy/share disabled and download them client side
// @author       ank
// @match        https://docs.google.com/spreadsheets/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';
    
    // Wait for page to load
    function waitForPageLoad() {
        return new Promise((resolve) => {
            if (document.readyState === 'complete') {
                resolve();
            } else {
                window.addEventListener('load', resolve);
            }
        });
    }
    
    // Extract sheet ID and GID from URL
    function extractSheetInfo() {
        const url = window.location.href;
        const sheetIdMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        const gidMatch = url.match(/gid=([0-9]+)/) || url.match(/#gid=([0-9]+)/);
        
        if (!sheetIdMatch) {
            console.error('Could not extract sheet ID from URL');
            return null;
        }
        
        const sheetId = sheetIdMatch[1];
        const gid = gidMatch ? gidMatch[1] : '0'; // Default to first sheet if no GID
        
        console.log(`Extracted - Sheet ID: ${sheetId}, GID: ${gid}`);
        return { sheetId, gid };
    }
    
    // Create export URLs
    function createExportUrls(sheetId, gid) {
        return {
            csv: `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`,
            tsv: `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=tsv&gid=${gid}`,
            xlsx: `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx&gid=${gid}`,
            ods: `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=ods&gid=${gid}`,
            pdf: `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=pdf&gid=${gid}`,
            html: `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=html&gid=${gid}`
        };
    }
    
    // Download data from URL
    async function downloadFromUrl(url, filename, format) {
        try {
            console.log(`Attempting to download ${format} from: ${url}`);
            
            const response = await fetch(url, {
                method: 'GET',
                credentials: 'include', // Include your cookies for authentication - works only with sheets you can view
                headers: {
                    'Accept': '*/*',
                    'Cache-Control': 'no-cache'
                }
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const contentType = response.headers.get('content-type');
            console.log(`Response content type: ${contentType}`);
            
            // Check if we got an error page instead of data
            const text = await response.text();
            
            // Google returns HTML error pages for unauthorized access
            if (text.includes('<html') && text.includes('error')) {
                throw new Error('Access denied - received error page');
            }
            
            // Check if we got actual data
            if (text.length < 50) {
                throw new Error('Response too short - likely not valid data');
            }
            
            // Create and download the file
            const blob = new Blob([text], { 
                type: contentType || 'text/plain' 
            });
            
            const downloadUrl = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = downloadUrl;
            link.download = filename;
            link.style.display = 'none';
            
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            URL.revokeObjectURL(downloadUrl);
            
            console.log(`Successfully downloaded: ${filename}`);
            return { success: true, data: text };
            
        } catch (error) {
            console.error(`Failed to download ${format}:`, error);
            return { success: false, error: error.message };
        }
    }
    
    // Create UI
    function createUI() {
        // Remove existing UI if present
        const existingUI = document.getElementById('sheets-export-ui');
        if (existingUI) {
            existingUI.remove();
        }
        
        const ui = document.createElement('div');
        ui.id = 'sheets-export-ui';
        ui.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 10000;
            background: white;
            border: 2px solid #4285f4;
            border-radius: 8px;
            padding: 15px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            font-family: 'Google Sans', Arial, sans-serif;
            font-size: 14px;
            min-width: 280px;
            max-width: 350px;
        `;
        
        ui.innerHTML = `
            <div style="font-weight: bold; margin-bottom: 12px; color: #4285f4; display: flex; align-items: center;">
                <span style="margin-right: 8px;"></span>
                Sheet Export Tool
            </div>
            
            <div style="margin-bottom: 15px; font-size: 12px; color: #666;">
                Extract data using Google's export API
            </div>
            
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 15px;">
                <button id="export-csv" style="background: #34a853; color: white; border: none; padding: 8px 12px; border-radius: 4px; cursor: pointer; font-size: 12px;">
                     CSV
                </button>
                <button id="export-xlsx" style="background: #1a73e8; color: white; border: none; padding: 8px 12px; border-radius: 4px; cursor: pointer; font-size: 12px;">
                    Excel
                </button>
                <button id="export-tsv" style="background: #ea4335; color: white; border: none; padding: 8px 12px; border-radius: 4px; cursor: pointer; font-size: 12px;">
                    TSV
                </button>
                <button id="export-html" style="background: #ff6d01; color: white; border: none; padding: 8px 12px; border-radius: 4px; cursor: pointer; font-size: 12px;">
                    HTML
                </button>
            </div>
            
            <button id="export-all" style="background: #9334e6; color: white; border: none; padding: 10px; border-radius: 4px; cursor: pointer; width: 100%; font-weight: bold;">
                ⬇️ Download All Formats
            </button>
            
            <div id="status" style="margin-top: 12px; font-size: 12px; color: #666; min-height: 20px;"></div>
            
            <div style="margin-top: 10px; font-size: 11px; color: #999; border-top: 1px solid #eee; padding-top: 8px;">
                Works by using Google's export API endpoints
            </div>
        `;
        
        document.body.appendChild(ui);
        
        // Add event listeners
        const sheetInfo = extractSheetInfo();
        if (!sheetInfo) {
            document.getElementById('status').textContent = 'Could not extract sheet info';
            return;
        }
        
        const urls = createExportUrls(sheetInfo.sheetId, sheetInfo.gid);
        const statusDiv = document.getElementById('status');
        
        // Individual format buttons
        document.getElementById('export-csv').addEventListener('click', () => {
            exportFormat('csv', urls.csv, statusDiv);
        });
        
        document.getElementById('export-xlsx').addEventListener('click', () => {
            exportFormat('xlsx', urls.xlsx, statusDiv);
        });
        
        document.getElementById('export-tsv').addEventListener('click', () => {
            exportFormat('tsv', urls.tsv, statusDiv);
        });
        
        document.getElementById('export-html').addEventListener('click', () => {
            exportFormat('html', urls.html, statusDiv);
        });
        
        // Export all button
        document.getElementById('export-all').addEventListener('click', () => {
            exportAllFormats(urls, statusDiv);
        });
        
        // Show initial status
        statusDiv.textContent = `Ready! Sheet ID: ${sheetInfo.sheetId.substring(0, 8)}...`;
    }
    
    // Export single format
    async function exportFormat(format, url, statusDiv) {
        statusDiv.textContent = `⏳ Downloading ${format.toUpperCase()}...`;
        
        const filename = `sheet_export_${Date.now()}.${format}`;
        const result = await downloadFromUrl(url, filename, format);
        
        if (result.success) {
            statusDiv.textContent = `Downloaded ${format.toUpperCase()} successfully!`;
            setTimeout(() => {
                statusDiv.textContent = 'Ready for next download';
            }, 3000);
        } else {
            statusDiv.textContent = `Failed: ${result.error}`;
        }
    }
    
    // Export all formats
    async function exportAllFormats(urls, statusDiv) {
        statusDiv.textContent = '⏳ Downloading all formats...';
        
        const formats = [
            { name: 'csv', url: urls.csv },
            { name: 'xlsx', url: urls.xlsx },
            { name: 'tsv', url: urls.tsv },
            { name: 'html', url: urls.html }
        ];
        
        let successCount = 0;
        const timestamp = Date.now();
        
        for (const format of formats) {
            const filename = `sheet_export_${timestamp}.${format.name}`;
            const result = await downloadFromUrl(format.url, filename, format.name);
            
            if (result.success) {
                successCount++;
            }
            
            // Small delay between downloads
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        
        statusDiv.textContent = `Downloaded ${successCount}/${formats.length} formats successfully!`;
    }
    
    // Add CSS for button hover effects
    function addStyles() {
        const style = document.createElement('style');
        style.textContent = `
            #sheets-export-ui button:hover {
                opacity: 0.9;
                transform: translateY(-1px);
                transition: all 0.2s ease;
            }
            
            #sheets-export-ui button:active {
                transform: translateY(0);
            }
        `;
        document.head.appendChild(style);
    }
    
    // Initialize the script
    async function init() {
        console.log('Google Sheets Export Bypass - Initializing...');
        
        // Wait for page to load
        await waitForPageLoad();
        
        // Add a small delay to ensure the page is fully rendered
        setTimeout(() => {
            addStyles();
            createUI();
            console.log('Google Sheets Export Bypass - Ready!');
        }, 1000);
    }
    
    // Start the script
    init();
    
})();

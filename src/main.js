import Storehouse from 'storehouse-js';
import * as monaco from 'https://cdn.jsdelivr.net/npm/monaco-editor@0.52.2/+esm';
import { marked } from 'marked';
import DOMPurify from 'dompurify';
import 'github-markdown-css/github-markdown-light.css';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';

const init = () => {
    let hasEdited = false;
    let scrollBarSync = false;

    const localStorageNamespace = 'com.markdownlivepreview';
    const localStorageKey = 'last_state';
    const localStorageScrollBarKey = 'scroll_bar_settings';
    const confirmationMessage = 'Are you sure you want to reset? Your changes will be lost.';
    // default template
    const defaultInput = `# Markdown syntax guide

## Headers

# This is a Heading h1
## This is a Heading h2
###### This is a Heading h6

## Emphasis

*This text will be italic*  
_This will also be italic_

**This text will be bold**  
__This will also be bold__

_You **can** combine them_

## Lists

### Unordered

* Item 1
* Item 2
* Item 2a
* Item 2b
    * Item 3a
    * Item 3b

### Ordered

1. Item 1
2. Item 2
3. Item 3
    1. Item 3a
    2. Item 3b

## Images

![This is an alt text.](/image/sample.webp "This is a sample image.")

## Links

You may be using [Markdown Live Preview](https://markdownlivepreview.com/).

## Blockquotes

> Markdown is a lightweight markup language with plain-text-formatting syntax, created in 2004 by John Gruber with Aaron Swartz.
>
>> Markdown is often used to format readme files, for writing messages in online discussion forums, and to create rich text using a plain text editor.

## Tables

| Left columns  | Right columns |
| ------------- |:-------------:|
| left foo      | right foo     |
| left bar      | right bar     |
| left baz      | right baz     |

## Blocks of code

${"`"}${"`"}${"`"}
let message = 'Hello world';
alert(message);
${"`"}${"`"}${"`"}

## Inline code

This web site is using ${"`"}markedjs/marked${"`"}.
`;

    self.MonacoEnvironment = {
        getWorker(_, label) {
            return new Proxy({}, { get: () => () => { } });
        }
    }

    let setupEditor = () => {
        let editor = monaco.editor.create(document.querySelector('#editor'), {
            fontSize: 14,
            language: 'markdown',
            minimap: { enabled: false },
            scrollBeyondLastLine: false,
            automaticLayout: true,
            scrollbar: {
                vertical: 'visible',
                horizontal: 'visible'
            },
            wordWrap: 'on',
            hover: { enabled: false },
            quickSuggestions: false,
            suggestOnTriggerCharacters: false,
            folding: false
        });

        editor.onDidChangeModelContent(() => {
            let changed = editor.getValue() != defaultInput;
            if (changed) {
                hasEdited = true;
            }
            let value = editor.getValue();
            convert(value);
            saveLastContent(value);
        });

        editor.onDidScrollChange((e) => {
            if (!scrollBarSync) {
                return;
            }

            const scrollTop = e.scrollTop;
            const scrollHeight = e.scrollHeight;
            const height = editor.getLayoutInfo().height;

            const maxScrollTop = scrollHeight - height;
            const scrollRatio = scrollTop / maxScrollTop;

            let previewElement = document.querySelector('#preview');
            let targetY = (previewElement.scrollHeight - previewElement.clientHeight) * scrollRatio;
            previewElement.scrollTo(0, targetY);
        });

        return editor;
    };

    // Render markdown text as html
    let convert = (markdown) => {
        let options = {
            headerIds: false,
            mangle: false
        };
        let html = marked.parse(markdown, options);
        let sanitized = DOMPurify.sanitize(html);
        document.querySelector('#output').innerHTML = sanitized;
    };

    // Reset input text
    let reset = () => {
        let changed = editor.getValue() != defaultInput;
        if (hasEdited || changed) {
            var confirmed = window.confirm(confirmationMessage);
            if (!confirmed) {
                return;
            }
        }
        presetValue(defaultInput);
        document.querySelectorAll('.column').forEach((element) => {
            element.scrollTo({ top: 0 });
        });
    };

    let presetValue = (value) => {
        editor.setValue(value);
        editor.revealPosition({ lineNumber: 1, column: 1 });
        editor.focus();
        hasEdited = false;
    };

    // ----- sync scroll position -----

    let initScrollBarSync = (settings) => {
        let checkbox = document.querySelector('#sync-scroll-checkbox');
        checkbox.checked = settings;
        scrollBarSync = settings;

        checkbox.addEventListener('change', (event) => {
            let checked = event.currentTarget.checked;
            scrollBarSync = checked;
            saveScrollBarSettings(checked);
        });
    };

    let enableScrollBarSync = () => {
        scrollBarSync = true;
    };

    let disableScrollBarSync = () => {
        scrollBarSync = false;
    };

    // ----- clipboard utils -----

    let copyToClipboard = (text, successHandler, errorHandler) => {
        navigator.clipboard.writeText(text).then(
            () => {
                successHandler();
            },

            () => {
                errorHandler();
            }
        );
    };

    let notifyCopied = () => {
        let labelElement = document.querySelector("#copy-button a");
        labelElement.innerHTML = "Copied!";
        setTimeout(() => {
            labelElement.innerHTML = "Copy";
        }, 1000)
    };

    // ----- theme management -----

    let currentTheme = 'light';
    const themeStorageKey = 'markdown-preview-theme';

    let initTheme = () => {
        const savedTheme = localStorage.getItem(themeStorageKey);
        if (savedTheme) {
            currentTheme = savedTheme;
        }
        applyTheme(currentTheme);
    };

    let applyTheme = (theme) => {
        document.documentElement.setAttribute('data-theme', theme);
        currentTheme = theme;
        localStorage.setItem(themeStorageKey, theme);
        
        // Update theme icon
        const themeIcon = document.getElementById('theme-icon');
        if (themeIcon) {
            themeIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
        }
    };

    let toggleTheme = () => {
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        applyTheme(newTheme);
    };

    // ----- page setup settings -----

    let pageSetupSettings = {
        pdf: {
            pageSize: 'a4',
            orientation: 'portrait',
            marginTop: 20,
            marginBottom: 20,
            marginLeft: 20,
            marginRight: 20,
            fontSize: 12,
            lineHeight: 1.5,
            pageNumbers: true
        },
        word: {
            pageSize: 'a4',
            orientation: 'portrait',
            marginTop: 25.4,
            marginBottom: 25.4,
            marginLeft: 25.4,
            marginRight: 25.4,
            fontFamily: 'Arial',
            fontSize: 12,
            lineHeight: 1.15,
            pageNumbers: true,
            tableOfContents: false
        }
    };

    // Page size mappings
    const pageSizeMap = {
        a4: { width: 210, height: 297 },
        a3: { width: 297, height: 420 },
        a5: { width: 148, height: 210 },
        letter: { width: 215.9, height: 279.4 },
        legal: { width: 215.9, height: 355.6 },
        tabloid: { width: 279.4, height: 431.8 }
    };

    // ----- settings and styling utils -----

    let currentSettings = {
        h1Size: 1.8,
        h2Size: 1.5,
        h3Size: 1.25,
        headingColor: '#333333',
        bodyColor: '#333333',
        bodySize: 15,
        headingFont: 'PingFang SC',
        bodyFont: 'PingFang SC',
        background1: '#f8f9fa',
        background2: '#ffffff',
        codeSize: 14,
        codeBg: '#f6f8fa',
        codeText: '#24292e',
        codeFont: 'Monaco',
        tableSize: 14,
        borderColor: '#d0d7de',
        headerBg: '#f6f8fa',
        headerText: '#24292e',
        cellBg: '#ffffff',
        cellText: '#24292e'
    };

    let themePresets = {
        light: {
            h1Size: 1.8, h2Size: 1.5, h3Size: 1.25,
            headingColor: '#333333', bodyColor: '#333333', bodySize: 15,
            headingFont: 'Arial', bodyFont: 'Arial',
            background1: '#ffffff', background2: '#f8f9fa',
            codeSize: 14, codeBg: '#f6f8fa', codeText: '#24292e', codeFont: 'Monaco',
            tableSize: 14, borderColor: '#d0d7de', headerBg: '#f6f8fa', headerText: '#24292e',
            cellBg: '#ffffff', cellText: '#24292e'
        },
        warm: {
            h1Size: 1.8, h2Size: 1.5, h3Size: 1.25,
            headingColor: '#8B4513', bodyColor: '#5D4037', bodySize: 15,
            headingFont: 'Georgia', bodyFont: 'Georgia',
            background1: '#FFF8DC', background2: '#F5F5DC',
            codeSize: 14, codeBg: '#F0E68C', codeText: '#8B4513', codeFont: 'Courier New',
            tableSize: 14, borderColor: '#D2B48C', headerBg: '#F0E68C', headerText: '#8B4513',
            cellBg: '#FFF8DC', cellText: '#5D4037'
        },
        elegant: {
            h1Size: 2.0, h2Size: 1.6, h3Size: 1.3,
            headingColor: '#2C3E50', bodyColor: '#34495E', bodySize: 16,
            headingFont: 'Times New Roman', bodyFont: 'Times New Roman',
            background1: '#FDFDFD', background2: '#F8F9FA',
            codeSize: 14, codeBg: '#ECF0F1', codeText: '#2C3E50', codeFont: 'Consolas',
            tableSize: 14, borderColor: '#BDC3C7', headerBg: '#ECF0F1', headerText: '#2C3E50',
            cellBg: '#FDFDFD', cellText: '#34495E'
        },
        dark: {
            h1Size: 1.8, h2Size: 1.5, h3Size: 1.25,
            headingColor: '#E8E8E8', bodyColor: '#D0D0D0', bodySize: 15,
            headingFont: 'Arial', bodyFont: 'Arial',
            background1: '#2D2D2D', background2: '#1E1E1E',
            codeSize: 14, codeBg: '#3C3C3C', codeText: '#E8E8E8', codeFont: 'Monaco',
            tableSize: 14, borderColor: '#555555', headerBg: '#3C3C3C', headerText: '#E8E8E8',
            cellBg: '#2D2D2D', cellText: '#D0D0D0'
        },
        gradient: {
            h1Size: 1.8, h2Size: 1.5, h3Size: 1.25,
            headingColor: '#4A90E2', bodyColor: '#333333', bodySize: 15,
            headingFont: 'Arial', bodyFont: 'Arial',
            background1: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)', background2: '#ffffff',
            codeSize: 14, codeBg: '#F0F4F8', codeText: '#2D3748', codeFont: 'Monaco',
            tableSize: 14, borderColor: '#E2E8F0', headerBg: '#F0F4F8', headerText: '#2D3748',
            cellBg: '#ffffff', cellText: '#333333'
        },
        nature: {
            h1Size: 1.8, h2Size: 1.5, h3Size: 1.25,
            headingColor: '#2D5016', bodyColor: '#3E5C2A', bodySize: 15,
            headingFont: 'Georgia', bodyFont: 'Georgia',
            background1: '#F0F8F0', background2: '#E8F5E8',
            codeSize: 14, codeBg: '#D4E6D4', codeText: '#2D5016', codeFont: 'Courier New',
            tableSize: 14, borderColor: '#A8C8A8', headerBg: '#D4E6D4', headerText: '#2D5016',
            cellBg: '#F0F8F0', cellText: '#3E5C2A'
        },
        sunset: {
            h1Size: 1.8, h2Size: 1.5, h3Size: 1.25,
            headingColor: '#D2691E', bodyColor: '#8B4513', bodySize: 15,
            headingFont: 'Georgia', bodyFont: 'Georgia',
            background1: '#FFF5EE', background2: '#FFE4B5',
            codeSize: 14, codeBg: '#FFE4B5', codeText: '#8B4513', codeFont: 'Courier New',
            tableSize: 14, borderColor: '#DEB887', headerBg: '#FFE4B5', headerText: '#8B4513',
            cellBg: '#FFF5EE', cellText: '#8B4513'
        },
        ocean: {
            h1Size: 1.8, h2Size: 1.5, h3Size: 1.25,
            headingColor: '#0066CC', bodyColor: '#003366', bodySize: 15,
            headingFont: 'Arial', bodyFont: 'Arial',
            background1: '#F0F8FF', background2: '#E6F3FF',
            codeSize: 14, codeBg: '#CCE6FF', codeText: '#003366', codeFont: 'Monaco',
            tableSize: 14, borderColor: '#99CCFF', headerBg: '#CCE6FF', headerText: '#003366',
            cellBg: '#F0F8FF', cellText: '#003366'
        },
        mint: {
            h1Size: 1.8, h2Size: 1.5, h3Size: 1.25,
            headingColor: '#006B6B', bodyColor: '#004D4D', bodySize: 15,
            headingFont: 'Arial', bodyFont: 'Arial',
            background1: '#F0FFFF', background2: '#E0FFFF',
            codeSize: 14, codeBg: '#B0E0E6', codeText: '#004D4D', codeFont: 'Monaco',
            tableSize: 14, borderColor: '#87CEEB', headerBg: '#B0E0E6', headerText: '#004D4D',
            cellBg: '#F0FFFF', cellText: '#004D4D'
        },
        tiffany: {
            h1Size: 1.8, h2Size: 1.5, h3Size: 1.25,
            headingColor: '#0ABAB5', bodyColor: '#333333', bodySize: 15,
            headingFont: 'PingFang SC', bodyFont: 'PingFang SC',
            background1: '#F0FDFC', background2: '#E6FFFE',
            codeSize: 14, codeBg: '#CCFBF1', codeText: '#0F766E', codeFont: 'Monaco',
            tableSize: 14, borderColor: '#5EEAD4', headerBg: '#CCFBF1', headerText: '#0F766E',
            cellBg: '#F0FDFC', cellText: '#333333'
        }
    };

    let showSettingsModal = () => {
        const modal = document.getElementById('settings-modal');
        modal.style.display = 'block';
        loadSettingsToForm();
        updatePreview();
    };

    let hideSettingsModal = () => {
        const modal = document.getElementById('settings-modal');
        modal.style.display = 'none';
    };

    let loadSettingsToForm = () => {
        Object.keys(currentSettings).forEach(key => {
            const element = document.getElementById(key.replace(/([A-Z])/g, '-$1').toLowerCase());
            if (element) {
                element.value = currentSettings[key];
            }
        });
    };

    let saveSettingsFromForm = () => {
        Object.keys(currentSettings).forEach(key => {
            const element = document.getElementById(key.replace(/([A-Z])/g, '-$1').toLowerCase());
            if (element) {
                currentSettings[key] = element.value;
            }
        });
    };

    let applyCustomStyles = (element, settings) => {
        const originalStyles = {};
        const style = element.style;
        
        // Store original styles
        const properties = ['fontSize', 'fontFamily', 'color', 'backgroundColor'];
        properties.forEach(prop => {
            originalStyles[prop] = style[prop];
        });

        // Apply custom styles
        style.fontSize = settings.bodySize + 'px';
        style.fontFamily = settings.bodyFont;
        style.color = settings.bodyColor;
        style.backgroundColor = settings.background1;

        // Apply heading styles
        const headings = element.querySelectorAll('h1, h2, h3, h4, h5, h6');
        headings.forEach(heading => {
            const level = parseInt(heading.tagName.charAt(1));
            const sizeMap = {
                1: settings.h1Size,
                2: settings.h2Size,
                3: settings.h3Size,
                4: settings.h3Size * 0.9,
                5: settings.h3Size * 0.8,
                6: settings.h3Size * 0.7
            };
            
            heading.style.fontSize = (sizeMap[level] || settings.h3Size) + 'em';
            heading.style.fontFamily = settings.headingFont;
            heading.style.color = settings.headingColor;
        });

        // Apply code styles
        const codeElements = element.querySelectorAll('code, pre');
        codeElements.forEach(code => {
            code.style.fontSize = settings.codeSize + 'px';
            code.style.fontFamily = settings.codeFont;
            code.style.backgroundColor = settings.codeBg;
            code.style.color = settings.codeText;
        });

        // Apply table styles
        const tables = element.querySelectorAll('table');
        tables.forEach(table => {
            table.style.fontSize = settings.tableSize + 'px';
            table.style.borderColor = settings.borderColor;
            
            const headers = table.querySelectorAll('th');
            headers.forEach(header => {
                header.style.backgroundColor = settings.headerBg;
                header.style.color = settings.headerText;
            });
            
            const cells = table.querySelectorAll('td');
            cells.forEach(cell => {
                cell.style.backgroundColor = settings.cellBg;
                cell.style.color = settings.cellText;
            });
        });

        return originalStyles;
    };

    let restoreOriginalStyles = (element, originalStyles) => {
        const style = element.style;
        Object.keys(originalStyles).forEach(prop => {
            style[prop] = originalStyles[prop];
        });

        // Reset all custom styles
        const allElements = element.querySelectorAll('*');
        allElements.forEach(el => {
            el.style.fontSize = '';
            el.style.fontFamily = '';
            el.style.color = '';
            el.style.backgroundColor = '';
        });
    };

    let applyThemePreset = (themeName) => {
        if (themePresets[themeName]) {
            currentSettings = { ...themePresets[themeName] };
            loadSettingsToForm();
            updatePreview();
        }
    };

    let updatePreview = () => {
        const previewContent = document.getElementById('preview-content');
        const originalContent = document.querySelector('#output').innerHTML;
        
        // Copy the original content to preview
        previewContent.innerHTML = originalContent;
        
        // Apply current settings to preview
        const tempSettings = { ...currentSettings };
        saveSettingsFromForm();
        applyCustomStyles(previewContent, currentSettings);
        
        // Restore settings if they were changed
        currentSettings = tempSettings;
    };

    let resetPreview = () => {
        const previewContent = document.getElementById('preview-content');
        const originalContent = document.querySelector('#output').innerHTML;
        previewContent.innerHTML = originalContent;
        
        // Remove all custom styles
        const allElements = previewContent.querySelectorAll('*');
        allElements.forEach(el => {
            el.style.fontSize = '';
            el.style.fontFamily = '';
            el.style.color = '';
            el.style.backgroundColor = '';
        });
        previewContent.style.fontSize = '';
        previewContent.style.fontFamily = '';
        previewContent.style.color = '';
        previewContent.style.backgroundColor = '';
    };

    let setupPreviewListeners = () => {
        // Auto-update preview when settings change
        const autoPreviewCheckbox = document.getElementById('auto-preview');
        const updatePreviewBtn = document.getElementById('update-preview');
        const resetPreviewBtn = document.getElementById('reset-preview');
        
        // Get all input elements in the settings
        const allInputs = document.querySelectorAll('#settings-modal input, #settings-modal select');
        
        allInputs.forEach(input => {
            input.addEventListener('input', () => {
                if (autoPreviewCheckbox.checked) {
                    saveSettingsFromForm();
                    updatePreview();
                }
            });
            
            input.addEventListener('change', () => {
                if (autoPreviewCheckbox.checked) {
                    saveSettingsFromForm();
                    updatePreview();
                }
            });
        });
        
        // Manual update button
        updatePreviewBtn.addEventListener('click', () => {
            saveSettingsFromForm();
            updatePreview();
        });
        
        // Reset preview button
        resetPreviewBtn.addEventListener('click', resetPreview);
    };

    // ----- page setup modal management -----

    let showPageSetupModal = () => {
        const modal = document.getElementById('page-setup-modal');
        modal.style.display = 'block';
        loadPageSetupToForm();
    };

    let hidePageSetupModal = () => {
        const modal = document.getElementById('page-setup-modal');
        modal.style.display = 'none';
    };

    let loadPageSetupToForm = () => {
        // Load PDF settings
        Object.keys(pageSetupSettings.pdf).forEach(key => {
            const element = document.getElementById('pdf-' + key.replace(/([A-Z])/g, '-$1').toLowerCase());
            if (element) {
                if (element.type === 'checkbox') {
                    element.checked = pageSetupSettings.pdf[key];
                } else {
                    element.value = pageSetupSettings.pdf[key];
                }
            }
        });

        // Load Word settings
        Object.keys(pageSetupSettings.word).forEach(key => {
            const element = document.getElementById('word-' + key.replace(/([A-Z])/g, '-$1').toLowerCase());
            if (element) {
                if (element.type === 'checkbox') {
                    element.checked = pageSetupSettings.word[key];
                } else {
                    element.value = pageSetupSettings.word[key];
                }
            }
        });
    };

    let savePageSetupFromForm = () => {
        // Save PDF settings
        Object.keys(pageSetupSettings.pdf).forEach(key => {
            const element = document.getElementById('pdf-' + key.replace(/([A-Z])/g, '-$1').toLowerCase());
            if (element) {
                if (element.type === 'checkbox') {
                    pageSetupSettings.pdf[key] = element.checked;
                } else if (element.type === 'number') {
                    pageSetupSettings.pdf[key] = parseFloat(element.value);
                } else {
                    pageSetupSettings.pdf[key] = element.value;
                }
            }
        });

        // Save Word settings
        Object.keys(pageSetupSettings.word).forEach(key => {
            const element = document.getElementById('word-' + key.replace(/([A-Z])/g, '-$1').toLowerCase());
            if (element) {
                if (element.type === 'checkbox') {
                    pageSetupSettings.word[key] = element.checked;
                } else if (element.type === 'number') {
                    pageSetupSettings.word[key] = parseFloat(element.value);
                } else {
                    pageSetupSettings.word[key] = element.value;
                }
            }
        });
    };

    let resetPageSetupToDefault = () => {
        pageSetupSettings = {
            pdf: {
                pageSize: 'a4',
                orientation: 'portrait',
                marginTop: 20,
                marginBottom: 20,
                marginLeft: 20,
                marginRight: 20,
                fontSize: 12,
                lineHeight: 1.5,
                pageNumbers: true
            },
            word: {
                pageSize: 'a4',
                orientation: 'portrait',
                marginTop: 25.4,
                marginBottom: 25.4,
                marginLeft: 25.4,
                marginRight: 25.4,
                fontFamily: 'Arial',
                fontSize: 12,
                lineHeight: 1.15,
                pageNumbers: true,
                tableOfContents: false
            }
        };
        loadPageSetupToForm();
    };

    // ----- download utils -----

    let downloadAsImage = async () => {
        // Show settings modal for image download
        showSettingsModal();
    };

    let downloadAsImageWithSettings = async (settings) => {
        try {
            // Apply custom styles to preview element
            const previewElement = document.querySelector('#output');
            const originalStyles = applyCustomStyles(previewElement, settings);
            
            const canvas = await html2canvas(previewElement, {
                backgroundColor: settings.background1 || '#ffffff',
                scale: 2,
                useCORS: true,
                allowTaint: true
            });
            
            // Restore original styles
            restoreOriginalStyles(previewElement, originalStyles);
            
            const link = document.createElement('a');
            link.download = 'markdown-preview.png';
            link.href = canvas.toDataURL('image/png');
            link.click();
        } catch (error) {
            console.error('Error generating image:', error);
            alert('Error generating image. Please try again.');
        }
    };

    let downloadAsPDF = async () => {
        try {
            const previewElement = document.querySelector('#output');
            const canvas = await html2canvas(previewElement, {
                backgroundColor: '#ffffff',
                scale: 2,
                useCORS: true,
                allowTaint: true
            });
            
            const imgData = canvas.toDataURL('image/png');
            
            // Get page setup settings
            const settings = pageSetupSettings.pdf;
            const pageSize = pageSizeMap[settings.pageSize];
            const orientation = settings.orientation;
            
            // Calculate dimensions based on orientation
            let pageWidth, pageHeight;
            if (orientation === 'landscape') {
                pageWidth = pageSize.height;
                pageHeight = pageSize.width;
            } else {
                pageWidth = pageSize.width;
                pageHeight = pageSize.height;
            }
            
            // Create PDF with custom page size
            const pdf = new jsPDF(orientation, 'mm', [pageWidth, pageHeight]);
            
            // Calculate content area (page size minus margins)
            const contentWidth = pageWidth - settings.marginLeft - settings.marginRight;
            const contentHeight = pageHeight - settings.marginTop - settings.marginBottom;
            
            // Calculate image dimensions to fit content area
            const imgWidth = contentWidth;
            const imgHeight = (canvas.height * imgWidth) / canvas.width;
            
            let heightLeft = imgHeight;
            let position = 0;
            let pageNumber = 1;

            // Add first page
            pdf.addImage(imgData, 'PNG', settings.marginLeft, settings.marginTop + position, imgWidth, imgHeight);
            
            // Add page numbers if enabled
            if (settings.pageNumbers) {
                pdf.setFontSize(10);
                pdf.text(`Page ${pageNumber}`, pageWidth - 20, pageHeight - 10);
            }
            
            heightLeft -= contentHeight;

            // Add additional pages if needed
            while (heightLeft >= 0) {
                position = heightLeft - imgHeight;
                pdf.addPage();
                pageNumber++;
                
                pdf.addImage(imgData, 'PNG', settings.marginLeft, settings.marginTop + position, imgWidth, imgHeight);
                
                // Add page numbers if enabled
                if (settings.pageNumbers) {
                    pdf.setFontSize(10);
                    pdf.text(`Page ${pageNumber}`, pageWidth - 20, pageHeight - 10);
                }
                
                heightLeft -= contentHeight;
            }

            pdf.save('markdown-preview.pdf');
        } catch (error) {
            console.error('Error generating PDF:', error);
            alert('Error generating PDF. Please try again.');
        }
    };

    let downloadAsWord = async () => {
        try {
            const markdownText = editor.getValue();
            const html = marked.parse(markdownText, { headerIds: false, mangle: false });
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = html;
            
            // Get page setup settings
            const settings = pageSetupSettings.word;
            const pageSize = pageSizeMap[settings.pageSize];
            const orientation = settings.orientation;
            
            // Calculate dimensions based on orientation
            let pageWidth, pageHeight;
            if (orientation === 'landscape') {
                pageWidth = pageSize.height;
                pageHeight = pageSize.width;
            } else {
                pageWidth = pageSize.width;
                pageHeight = pageSize.height;
            }
            
            const paragraphs = [];
            
            // Add table of contents if enabled
            if (settings.tableOfContents) {
                paragraphs.push(new Paragraph({
                    text: "Table of Contents",
                    heading: HeadingLevel.HEADING_1
                }));
                paragraphs.push(new Paragraph({
                    text: "This is a placeholder for table of contents. In a full implementation, this would be automatically generated based on the headings in the document.",
                    italics: true
                }));
                paragraphs.push(new Paragraph({ text: "" })); // Empty line
            }
            
            // Process each element in the HTML
            const processElement = (element) => {
                if (element.nodeType === Node.TEXT_NODE) {
                    const text = element.textContent.trim();
                    if (text) {
                        return new TextRun({
                            text: text,
                            font: settings.fontFamily,
                            size: settings.fontSize * 2 // docx uses half-points
                        });
                    }
                    return null;
                }
                
                if (element.nodeType === Node.ELEMENT_NODE) {
                    const tagName = element.tagName.toLowerCase();
                    const children = Array.from(element.childNodes).map(processElement).filter(Boolean);
                    
                    if (tagName === 'h1') {
                        return new Paragraph({
                            children: children,
                            heading: HeadingLevel.HEADING_1,
                            spacing: { after: 200, before: 200 }
                        });
                    } else if (tagName === 'h2') {
                        return new Paragraph({
                            children: children,
                            heading: HeadingLevel.HEADING_2,
                            spacing: { after: 200, before: 200 }
                        });
                    } else if (tagName === 'h3') {
                        return new Paragraph({
                            children: children,
                            heading: HeadingLevel.HEADING_3,
                            spacing: { after: 200, before: 200 }
                        });
                    } else if (tagName === 'h4') {
                        return new Paragraph({
                            children: children,
                            heading: HeadingLevel.HEADING_4,
                            spacing: { after: 200, before: 200 }
                        });
                    } else if (tagName === 'h5') {
                        return new Paragraph({
                            children: children,
                            heading: HeadingLevel.HEADING_5,
                            spacing: { after: 200, before: 200 }
                        });
                    } else if (tagName === 'h6') {
                        return new Paragraph({
                            children: children,
                            heading: HeadingLevel.HEADING_6,
                            spacing: { after: 200, before: 200 }
                        });
                    } else if (tagName === 'p') {
                        return new Paragraph({
                            children: children,
                            spacing: { after: 200, before: 0 }
                        });
                    } else if (tagName === 'strong' || tagName === 'b') {
                        return new TextRun({
                            text: element.textContent,
                            bold: true,
                            font: settings.fontFamily,
                            size: settings.fontSize * 2
                        });
                    } else if (tagName === 'em' || tagName === 'i') {
                        return new TextRun({
                            text: element.textContent,
                            italics: true,
                            font: settings.fontFamily,
                            size: settings.fontSize * 2
                        });
                    } else if (tagName === 'code') {
                        return new TextRun({
                            text: element.textContent,
                            font: 'Courier New',
                            size: settings.fontSize * 2
                        });
                    } else if (tagName === 'br') {
                        return new TextRun({
                            text: '\n',
                            font: settings.fontFamily,
                            size: settings.fontSize * 2
                        });
                    } else {
                        // For other elements, just process their children
                        return children.length > 0 ? children : null;
                    }
                }
                return null;
            };
            
            const processedElements = Array.from(tempDiv.childNodes).map(processElement).filter(Boolean);
            paragraphs.push(...processedElements);
            
            // Add page numbers if enabled
            if (settings.pageNumbers) {
                paragraphs.push(new Paragraph({
                    text: "Page numbers would be added here in a full implementation",
                    italics: true,
                    alignment: "center"
                }));
            }
            
            const doc = new Document({
                sections: [{
                    properties: {
                        page: {
                            size: {
                                width: pageWidth * 28.35, // Convert mm to points
                                height: pageHeight * 28.35
                            },
                            margin: {
                                top: settings.marginTop * 28.35,
                                bottom: settings.marginBottom * 28.35,
                                left: settings.marginLeft * 28.35,
                                right: settings.marginRight * 28.35
                            }
                        }
                    },
                    children: paragraphs
                }]
            });
            
            const buffer = await Packer.toBuffer(doc);
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'markdown-preview.docx';
            link.click();
            URL.revokeObjectURL(link.href);
        } catch (error) {
            console.error('Error generating Word document:', error);
            alert('Error generating Word document. Please try again.');
        }
    };

    // ----- setup -----

    // setup navigation actions
    let setupResetButton = () => {
        document.querySelector("#reset-button").addEventListener('click', (event) => {
            event.preventDefault();
            reset();
        });
    };

    let setupCopyButton = (editor) => {
        document.querySelector("#copy-button").addEventListener('click', (event) => {
            event.preventDefault();
            let value = editor.getValue();
            copyToClipboard(value, () => {
                notifyCopied();
            },
                () => {
                    // nothing to do
                });
        });
    };

    let setupDownloadButtons = () => {
        document.querySelector("#download-image").addEventListener('click', (event) => {
            event.preventDefault();
            downloadAsImage();
        });

        document.querySelector("#download-pdf").addEventListener('click', (event) => {
            event.preventDefault();
            downloadAsPDF();
        });

        document.querySelector("#download-word").addEventListener('click', (event) => {
            event.preventDefault();
            downloadAsWord();
        });
    };

    let setupPageSetupButton = () => {
        document.querySelector("#page-setup-button").addEventListener('click', (event) => {
            event.preventDefault();
            showPageSetupModal();
        });
    };

    let setupThemeToggle = () => {
        const themeToggleBtn = document.getElementById('theme-toggle-btn');
        if (themeToggleBtn) {
            themeToggleBtn.addEventListener('click', toggleTheme);
        }
    };

    let setupSettingsModal = () => {
        const modal = document.getElementById('settings-modal');
        const closeBtn = document.querySelector('.close');
        const cancelBtn = document.getElementById('cancel-settings');
        const applyBtn = document.getElementById('apply-settings');
        const resetBtn = document.getElementById('reset-settings');

        // Close modal events
        closeBtn.addEventListener('click', hideSettingsModal);
        cancelBtn.addEventListener('click', hideSettingsModal);
        window.addEventListener('click', (event) => {
            if (event.target === modal) {
                hideSettingsModal();
            }
        });

        // Tab switching
        const tabButtons = document.querySelectorAll('.tab-button');
        const tabContents = document.querySelectorAll('.tab-content');

        tabButtons.forEach(button => {
            button.addEventListener('click', () => {
                const tabName = button.getAttribute('data-tab');
                
                // Remove active class from all buttons and contents
                tabButtons.forEach(btn => btn.classList.remove('active'));
                tabContents.forEach(content => content.classList.remove('active'));
                
                // Add active class to clicked button and corresponding content
                button.classList.add('active');
                document.getElementById(tabName + '-tab').classList.add('active');
            });
        });

        // Collapsible sections
        const collapsibles = document.querySelectorAll('.collapsible');
        collapsibles.forEach(collapsible => {
            const header = collapsible.querySelector('.collapsible-header');
            header.addEventListener('click', () => {
                collapsible.classList.toggle('active');
            });
        });

        // Theme preset buttons
        const presetButtons = document.querySelectorAll('.preset-btn');
        presetButtons.forEach(button => {
            button.addEventListener('click', () => {
                // Remove active class from all buttons
                presetButtons.forEach(btn => btn.classList.remove('active'));
                // Add active class to clicked button
                button.classList.add('active');
                
                const themeName = button.getAttribute('data-theme');
                applyThemePreset(themeName);
            });
        });

        // Apply settings
        applyBtn.addEventListener('click', () => {
            saveSettingsFromForm();
            downloadAsImageWithSettings(currentSettings);
            hideSettingsModal();
        });

        // Setup preview functionality
        setupPreviewListeners();

        // Reset settings
        resetBtn.addEventListener('click', () => {
            currentSettings = {
                h1Size: 1.8,
                h2Size: 1.5,
                h3Size: 1.25,
                headingColor: '#333333',
                bodyColor: '#333333',
                bodySize: 15,
                headingFont: 'PingFang SC',
                bodyFont: 'PingFang SC',
                background1: '#f8f9fa',
                background2: '#ffffff',
                codeSize: 14,
                codeBg: '#f6f8fa',
                codeText: '#24292e',
                codeFont: 'Monaco',
                tableSize: 14,
                borderColor: '#d0d7de',
                headerBg: '#f6f8fa',
                headerText: '#24292e',
                cellBg: '#ffffff',
                cellText: '#24292e'
            };
            loadSettingsToForm();
            updatePreview();
        });
    };

    let setupPageSetupModal = () => {
        const modal = document.getElementById('page-setup-modal');
        const closeBtn = document.getElementById('close-page-setup');
        const cancelBtn = document.getElementById('cancel-page-setup');
        const applyBtn = document.getElementById('apply-page-setup');
        const resetBtn = document.getElementById('reset-page-setup');

        // Close modal events
        closeBtn.addEventListener('click', hidePageSetupModal);
        cancelBtn.addEventListener('click', hidePageSetupModal);
        window.addEventListener('click', (event) => {
            if (event.target === modal) {
                hidePageSetupModal();
            }
        });

        // Tab switching
        const tabButtons = document.querySelectorAll('#page-setup-modal .tab-button');
        const tabContents = document.querySelectorAll('#page-setup-modal .tab-content');

        tabButtons.forEach(button => {
            button.addEventListener('click', () => {
                const tabName = button.getAttribute('data-tab');
                
                // Remove active class from all buttons and contents
                tabButtons.forEach(btn => btn.classList.remove('active'));
                tabContents.forEach(content => content.classList.remove('active'));
                
                // Add active class to clicked button and corresponding content
                button.classList.add('active');
                document.getElementById(tabName + '-tab').classList.add('active');
            });
        });

        // Apply settings
        applyBtn.addEventListener('click', () => {
            savePageSetupFromForm();
            hidePageSetupModal();
        });

        // Reset settings
        resetBtn.addEventListener('click', resetPageSetupToDefault);
    };

    // ----- local state -----

    let loadLastContent = () => {
        let lastContent = Storehouse.getItem(localStorageNamespace, localStorageKey);
        return lastContent;
    };

    let saveLastContent = (content) => {
        let expiredAt = new Date(2099, 1, 1);
        Storehouse.setItem(localStorageNamespace, localStorageKey, content, expiredAt);
    };

    let loadScrollBarSettings = () => {
        let lastContent = Storehouse.getItem(localStorageNamespace, localStorageScrollBarKey);
        return lastContent;
    };

    let saveScrollBarSettings = (settings) => {
        let expiredAt = new Date(2099, 1, 1);
        Storehouse.setItem(localStorageNamespace, localStorageScrollBarKey, settings, expiredAt);
    };

    let setupDivider = () => {
        let lastLeftRatio = 0.5;
        const divider = document.getElementById('split-divider');
        const leftPane = document.getElementById('edit');
        const rightPane = document.getElementById('preview');
        const container = document.getElementById('container');

        let isDragging = false;

        divider.addEventListener('mouseenter', () => {
            divider.classList.add('hover');
        });

        divider.addEventListener('mouseleave', () => {
            if (!isDragging) {
                divider.classList.remove('hover');
            }
        });

        divider.addEventListener('mousedown', () => {
            isDragging = true;
            divider.classList.add('active');
            document.body.style.cursor = 'col-resize';
        });

        divider.addEventListener('dblclick', () => {
            const containerRect = container.getBoundingClientRect();
            const totalWidth = containerRect.width;
            const dividerWidth = divider.offsetWidth;
            const halfWidth = (totalWidth - dividerWidth) / 2;

            leftPane.style.width = halfWidth + 'px';
            rightPane.style.width = halfWidth + 'px';
        });

        document.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            document.body.style.userSelect = 'none';
            const containerRect = container.getBoundingClientRect();
            const totalWidth = containerRect.width;
            const offsetX = e.clientX - containerRect.left;
            const dividerWidth = divider.offsetWidth;

            // Prevent overlap or out-of-bounds
            const minWidth = 100;
            const maxWidth = totalWidth - minWidth - dividerWidth;
            const leftWidth = Math.max(minWidth, Math.min(offsetX, maxWidth));
            leftPane.style.width = leftWidth + 'px';
            rightPane.style.width = (totalWidth - leftWidth - dividerWidth) + 'px';
            lastLeftRatio = leftWidth / (totalWidth - dividerWidth);
        });

        document.addEventListener('mouseup', () => {
            if (isDragging) {
                isDragging = false;
                divider.classList.remove('active');
                divider.classList.remove('hover');
                document.body.style.cursor = 'default';
                document.body.style.userSelect = '';
            }
        });

        window.addEventListener('resize', () => {
            const containerRect = container.getBoundingClientRect();
            const totalWidth = containerRect.width;
            const dividerWidth = divider.offsetWidth;
            const availableWidth = totalWidth - dividerWidth;

            const newLeft = availableWidth * lastLeftRatio;
            const newRight = availableWidth * (1 - lastLeftRatio);

            leftPane.style.width = newLeft + 'px';
            rightPane.style.width = newRight + 'px';
        });
    };

    // ----- entry point -----
    let lastContent = loadLastContent();
    let editor = setupEditor();
    if (lastContent) {
        presetValue(lastContent);
    } else {
        presetValue(defaultInput);
    }
    setupResetButton();
    setupCopyButton(editor);
    setupDownloadButtons();
    setupPageSetupButton();
    setupSettingsModal();
    setupPageSetupModal();
    setupThemeToggle();

    let scrollBarSettings = loadScrollBarSettings() || false;
    initScrollBarSync(scrollBarSettings);

    setupDivider();
    initTheme();
};

window.addEventListener("load", () => {
    init();
});

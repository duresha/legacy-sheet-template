// main.js
// Import libraries for PDF generation with download animation
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { saveAs } from 'file-saver';
import * as pdfjsLib from 'pdfjs-dist';

// Configure the PDF.js worker to use a local file
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./pdf.worker.min.mjs', import.meta.url).toString();

document.addEventListener('DOMContentLoaded', function() {
  console.log('Legacy Sheet Template loaded successfully');
  
  // Get the Save PDF button
  const savePdfBtn = document.getElementById('save-pdf-btn');
  
  // Add click event listener to the button
  savePdfBtn.addEventListener('click', function() {
    // Update button state
    savePdfBtn.textContent = 'Generating PDF...';
    savePdfBtn.disabled = true;
    
    // Hide the button temporarily
    savePdfBtn.style.display = 'none';
    
    // Get the element to convert to PDF
    const element = document.querySelector('.legacy-sheet');
    
    // Use html2canvas to capture the legacy sheet
    html2canvas(element, {
      scale: 2,  // Higher scale for better quality
      useCORS: true,
      scrollY: -window.scrollY,
      windowWidth: document.documentElement.offsetWidth,
      windowHeight: document.documentElement.offsetHeight
    }).then(canvas => {
      // Convert canvas to image
      const imgData = canvas.toDataURL('image/png');
      
      // Initialize PDF with letter size in portrait orientation
      const pdf = new jsPDF({
        unit: 'in',
        format: 'letter',
        orientation: 'portrait'
      });
      
      // Calculate dimensions to fit the page correctly
      const imgProps = pdf.getImageProperties(imgData);
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = pdf.internal.pageSize.getHeight();
      
      // Add image to PDF, sizing to fit the page
      pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
      
      // Generate the PDF as blob
      const pdfBlob = pdf.output('blob');
      
      // Use FileSaver to trigger browser download animation
      saveAs(pdfBlob, 'Legacy-Sheet.pdf');
      
      // Show button again after a short delay
      setTimeout(() => {
        savePdfBtn.style.display = 'block';
        savePdfBtn.textContent = 'Save PDF';
        savePdfBtn.disabled = false;
      }, 1500);
    });
  });

  // PDF Upload and Processing Functionality
  const uploadArea = document.getElementById('upload-area');
  const fileInput = document.getElementById('image-upload');
  const uploadPreview = document.getElementById('upload-preview');
  const pdfFilename = document.getElementById('image-filename');
  const extractionProgress = document.getElementById('extraction-progress');
  const extractionStatus = document.getElementById('extraction-status');
  const dataPreview = document.getElementById('data-preview');
  const applyDataBtn = document.getElementById('apply-data-btn');
  const removeImageBtn = document.getElementById('remove-image-btn');
  const extractDataBtn = document.getElementById('extract-data-btn');
  const imagePreviewThumb = document.getElementById('image-preview-thumb');

  // Track if we've already cleared the static template
  let staticTemplateCleared = false;
  
  // Track if the current data has already been applied
  let currentDataApplied = false;

  // Create Next Person Entry button
  const nextPersonBtn = document.createElement('button');
  nextPersonBtn.id = 'next-person-btn';
  nextPersonBtn.className = 'action-button primary-button';
  nextPersonBtn.innerHTML = 'Next Person Entry <span class="kbd-shortcut">Enter ↵</span>'; // Use the new CSS class
  nextPersonBtn.style.marginLeft = '0'; // Remove left margin to align with the button above
  nextPersonBtn.style.display = 'none'; // Initially hidden
  nextPersonBtn.style.width = '100%'; // Make it full width like the Apply button
  nextPersonBtn.style.marginTop = '20px'; // Add top margin to prevent overlap with toggle buttons
  nextPersonBtn.title = 'Press Enter key as a shortcut'; // Add tooltip to show keyboard shortcut
  
  // Insert Next Person button after the Apply button
  applyDataBtn.parentNode.insertBefore(nextPersonBtn, applyDataBtn.nextSibling);
  
  // Next Person Entry button click handler
  nextPersonBtn.addEventListener('click', function() {
    // Reset the UI for next person entry
    resetSidePanelUI();
  });

  // Add keyboard shortcut (Enter key) for Next Person Entry
  document.addEventListener('keydown', function(e) {
    // Check if Enter key is pressed and the Next Person button is visible
    if (e.key === 'Enter' && nextPersonBtn.style.display === 'inline-block') {
      e.preventDefault(); // Prevent default Enter key behavior
      resetSidePanelUI(); // Call the same function as button click
    }
  });

  // Function to reset ONLY the side panel UI for next person entry
  function resetSidePanelUI() {
    // Clear the file input
    fileInput.value = '';
    
    // Hide the upload preview and clear preview thumb
    uploadPreview.style.display = 'none';
    if (imagePreviewThumb.src && imagePreviewThumb.src !== '#') {
      URL.revokeObjectURL(imagePreviewThumb.src);
    }
    imagePreviewThumb.style.display = 'none';
    imagePreviewThumb.src = '#';
    document.getElementById('image-filename').style.display = 'block'; // Show filename field again
    pdfFilename.textContent = 'document.pdf'; // Reset placeholder text
    
    // Reset progress bar
    extractionProgress.style.width = '0%';
    
    // Reset extraction status
    extractionStatus.textContent = 'Waiting for Image...';
    
    // Clear extracted data
    dataPreview.innerHTML = '<p class="no-data-message">No data extracted yet</p>';
    
    // Reset button states (Apply button is now hidden)
    extractDataBtn.style.display = 'none';
    nextPersonBtn.style.display = 'none'; // Hide Next Person button
    
    // Clear any stored data for the side panel
    window.parsedGenealogyData = null;
    currentImageFile = null; // Clear stored file
    
    // Reset the applied flag
    currentDataApplied = false;
    
    // Show the upload area again
    uploadArea.style.display = 'flex';
    
    // Hide the remove image button
    removeImageBtn.style.display = 'none';
    
    console.log('Side panel UI reset for next person entry');
  }

  let currentImageFile = null; // To store the currently loaded image file

  // Handle remove Image button click
  removeImageBtn.addEventListener('click', function() {
    // Clear the file input
    fileInput.value = '';
    
    // Hide the upload preview and clear preview thumb
    uploadPreview.style.display = 'none';
    if (imagePreviewThumb.src && imagePreviewThumb.src !== '#') {
      URL.revokeObjectURL(imagePreviewThumb.src);
    }
    imagePreviewThumb.style.display = 'none';
    imagePreviewThumb.src = '#';
    document.getElementById('image-filename').style.display = 'block'; // Show filename field again
    pdfFilename.textContent = 'document.pdf'; // Reset placeholder text
    
    // Reset progress bar
    extractionProgress.style.width = '0%';
    
    // Reset extraction status
    extractionStatus.textContent = 'Waiting for Image...';
    
    // Clear extracted data
    dataPreview.innerHTML = '<p class="no-data-message">No data extracted yet</p>';
    
    // Disable apply button and hide extract button
    applyDataBtn.disabled = true;
    extractDataBtn.style.display = 'none';
    
    // Clear any stored data
    window.parsedGenealogyData = null;
    currentImageFile = null; // Clear stored file
    
    // Show the upload area again
    uploadArea.style.display = 'flex';
    
    console.log('Image removed and interface reset');
  });

  // Handle drag and drop events
  ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    uploadArea.addEventListener(eventName, preventDefaults, false);
    document.addEventListener(eventName, preventDefaults, false);
  });

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  ['dragenter', 'dragover'].forEach(eventName => {
    uploadArea.addEventListener(eventName, highlight, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    uploadArea.addEventListener(eventName, unhighlight, false);
  });

  function highlight() {
    uploadArea.classList.add('dragover');
  }

  function unhighlight() {
    uploadArea.classList.remove('dragover');
  }

  // Handle file drop
  uploadArea.addEventListener('drop', handleDrop, false);
  
  function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files && files.length > 0) {
      handleFiles(files);
    }
  }

  // Handle paste event for images globally
  document.addEventListener('paste', handlePaste, false);

  function handlePaste(e) {
    // Only process paste if the upload area is visible
    if (uploadArea.style.display === 'none') {
      console.log('Paste event ignored: upload area is not visible.');
      return;
    }

    const items = (e.clipboardData || e.originalEvent.clipboardData).items;
    let imageFile = null;
    for (let i = 0; i < items.length; i++) {
      if (items[i].type.indexOf('image') !== -1) {
        imageFile = items[i].getAsFile();
        break;
      }
    }

    if (imageFile) {
      // Create a FileList-like object to pass to handleFiles
      // as handleFiles expects a FileList
      const dataTransfer = new DataTransfer();
      dataTransfer.items.add(imageFile);
      handleFiles(dataTransfer.files);
      e.preventDefault(); // Prevent default paste action (e.g., displaying image in browser)
    } else {
      console.log('Pasted content was not an image.');
      // Optionally, inform the user that only images can be pasted
      // alert('Only images can be pasted.');
    }
  }

  // Handle file input change
  fileInput.addEventListener('change', function(e) {
    console.log('File input change event triggered');
    console.log('Files selected:', this.files ? this.files.length : 0);
    
    if (this.files && this.files.length > 0) {
      handleFiles(this.files);
    }
  });

  // Fix click handler to properly trigger file selection
  uploadArea.addEventListener('click', function(e) {
    // Only trigger if we're clicking on the "browse image" text
    if (e.target.classList.contains('browse-text')) {
      console.log('Browse image text clicked, triggering file input click');
      // Reset the file input to ensure change event fires even if selecting the same file
      fileInput.value = '';
      fileInput.click();
    }
  });

  // Handle the selected files
  function handleFiles(files) {
    console.log('Handling files:', files);
    if (!files || files.length === 0) return;
    
    const file = files[0];
    console.log('Processing file:', file.name, file.type);
    
    if (!file.type.startsWith('image/')) {
      alert('Please upload an image file.');
      return;
    }

    currentImageFile = file; // Store the file
    
    // Show image preview
    if (imagePreviewThumb.src && imagePreviewThumb.src !== '#') {
      URL.revokeObjectURL(imagePreviewThumb.src); // Revoke old object URL if any
    }
    imagePreviewThumb.src = URL.createObjectURL(file);
    imagePreviewThumb.style.display = 'block';
    
    // Hide upload area & show preview section
    uploadArea.style.display = 'none';
    uploadPreview.style.display = 'block';
    
    document.getElementById('image-filename').style.display = 'none'; // Hide the filename text
    
    extractionProgress.style.width = '0%'; // Reset progress bar
    extractionStatus.textContent = 'Starting OCR process...'; // Update status
    
    applyDataBtn.disabled = true; // Ensure apply button is disabled until data is parsed
    extractDataBtn.style.display = 'none'; // Hide Extract Data button since we're auto-extracting
    removeImageBtn.style.display = 'inline-block'; // Ensure remove button is visible
    
    // Start OCR extraction automatically
    extractTextFromImageWithOCR(file).catch(err => {
      console.error('Error during auto OCR extraction:', err);
      extractionStatus.textContent = 'Error during OCR. Please try another image.';
      removeImageBtn.disabled = false;
    });
  }

  // We don't need this event listener anymore since extraction happens automatically
  // But keep the button defined for backward compatibility
  extractDataBtn.addEventListener('click', function() {
    if (currentImageFile) {
      extractionStatus.textContent = 'Starting OCR process...';
      extractDataBtn.disabled = true;
      extractDataBtn.textContent = 'Extracting...';
      removeImageBtn.disabled = true; // Disable remove button during extraction

      extractTextFromImageWithOCR(currentImageFile).finally(() => {
        extractDataBtn.disabled = false;
        extractDataBtn.textContent = 'Extract Data';
        removeImageBtn.disabled = false;
      });
    }
  });

  // Update progress bar
  function updateProgressBar(percent) {
    extractionProgress.style.width = percent + '%';
  }

  // Extract text from PDF (No longer used, will be removed or commented)
  /*
  async function extractTextFromPdf(pdfFile) {
    try {
      extractionStatus.textContent = 'Extracting text...';
      updateProgressBar(10);
      
      // Read the file as ArrayBuffer
      const arrayBuffer = await pdfFile.arrayBuffer();
      updateProgressBar(20);
      
      // Load the PDF document
      const pdfDocument = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      updateProgressBar(30);
      
      extractionStatus.textContent = `PDF loaded successfully. Pages: ${pdfDocument.numPages}`;
      
      // Extract text from each page with layout preservation
      let fullText = '';
      let rawTextData = []; // Store raw text data for display
      
      for (let pageNum = 1; pageNum <= pdfDocument.numPages; pageNum++) {
        extractionStatus.textContent = `Extracting text from page ${pageNum}/${pdfDocument.numPages}...`;
        
        const page = await pdfDocument.getPage(pageNum);
        const textContent = await page.getTextContent();
        
        // Sort text items by vertical position (top to bottom) and then horizontal position (left to right)
        const sortedItems = textContent.items.sort((a, b) => {
          const verticalDiff = Math.abs(a.transform[5] - b.transform[5]);
          if (verticalDiff > 5) { // If items are on different lines (more than 5 units apart)
            return b.transform[5] - a.transform[5]; // Sort by vertical position (top to bottom)
          }
          return a.transform[4] - b.transform[4]; // Sort by horizontal position (left to right)
        });
        
        // Group items by line (items that are close vertically)
        let currentLine = [];
        let currentLineY = null;
        let pageText = '';
        
        sortedItems.forEach(item => {
          if (currentLineY === null) {
            currentLineY = item.transform[5];
          }
          
          // If this item is on a new line (more than 5 units different in Y position)
          if (Math.abs(item.transform[5] - currentLineY) > 5) {
            // Add the previous line to the page text
            if (currentLine.length > 0) {
              pageText += currentLine.join(' ') + '\n';
            }
            currentLine = [];
            currentLineY = item.transform[5];
          }
          
          currentLine.push(item.str);
        });
        
        // Add the last line
        if (currentLine.length > 0) {
          pageText += currentLine.join(' ') + '\n';
        }
        
        fullText += pageText + '\n\n';
        
        // Store raw text data for display
        rawTextData.push({
          pageNumber: pageNum,
          text: pageText,
          items: textContent.items
        });
        
        // Update progress based on pages (from 30% to 70%)
        const progressPercent = 30 + (pageNum / pdfDocument.numPages * 40);
        updateProgressBar(progressPercent);
      }
      
      // Progress to analysis phase
      updateProgressBar(70);
      extractionStatus.textContent = 'Analyzing extracted text...';
      
      // Store raw text data for display
      window.rawExtractedText = rawTextData;
      
      // Process the extracted text
      processExtractedText(fullText);
    } catch (error) {
      console.error('Error extracting text from PDF:', error);
      extractionStatus.textContent = 'Error extracting text from PDF. Please try another file.';
      updateProgressBar(0);
    }
  }
  */

  // Process the extracted text to identify genealogy data
  function processExtractedText(text) {
    extractionStatus.textContent = 'Analyzing genealogy data...';
    updateProgressBar(75);

    // Try to identify generation title
    let generationTitle = '';
    const generationMatch = text.match(/^(.*?Generation)/im);
    if (generationMatch) {
      generationTitle = generationMatch[1].trim();
    }

    // Split text into person entries by numbers at the start of a line (e.g., 119., 120., etc.)
    const personEntryRegex = /^\s*(\d{1,4})\.\s/mg;
    let match;
    let indices = [];
    while ((match = personEntryRegex.exec(text)) !== null) {
      indices.push({
        index: match.index,
        number: match[1]
      });
    }

    let persons = [];
    for (let i = 0; i < indices.length; i++) {
      const start = indices[i].index;
      const end = i + 1 < indices.length ? indices[i + 1].index : text.length;
      const entryText = text.slice(start, end).trim();
      // Extract number
      const numberMatch = entryText.match(/^(\d{1,4})\.\s*/);
      let number = numberMatch ? numberMatch[1] : '';
      // Remove the number prefix
      let restText = entryText.replace(/^(\d{1,4})\.\s*/, '');
      // Split into lines
      const lines = restText.split(/\r?\n/).filter(line => line.trim() !== '');
      let name = '';
      let rest = '';
      if (lines.length > 0) {
        // If the first line is bolded (e.g., **Name**), use it as name
        // Otherwise, just use the first line
        name = lines[0].replace(/\*\*(.+)\*\*/, '$1').trim();
        rest = lines.slice(1).join('\n');
      } else {
        rest = restText;
      }
      persons.push({
        number,
        name,
        raw: entryText,
        rest
      });
    }

    // Build parsedData structure
    const parsedData = {
      generation: generationTitle,
      persons
    };

    updateProgressBar(95);
    displayParsedData(parsedData);
  }

  // Display parsed data without triggering application (for tab switching)
  function displayParsedDataForTabSwitch(parsedData) {
    if (!parsedData || !parsedData.persons || parsedData.persons.length === 0) {
      dataPreview.innerHTML = '<p class="no-data-message">No genealogy data could be extracted. Please try another file or format.</p>';
      return;
    }

    let html = '';

    // Generation title
    if (parsedData.generation) {
      html += `<div class="parsed-item"><strong>Generation:</strong> ${parsedData.generation}</div>`;
    }

    // Persons
    parsedData.persons.forEach(person => {
      let detailsHtml = '';
      let nameForHeader = preserveHyperlinks(person.name || '');

      if (person.raw) {
        const rawLines = person.raw.split(/\\r?\\n/);
        let firstMeaningfulLineText = '';
        let firstMeaningfulLineIndex = -1;

        for (let i = 0; i < rawLines.length; i++) {
          if (rawLines[i].trim() !== '') {
            firstMeaningfulLineText = rawLines[i];
            firstMeaningfulLineIndex = i;
            break;
          }
        }

        if (firstMeaningfulLineText) {
          // Try to extract name and bio from the first meaningful line for preview
          const nameAndBioMatch = firstMeaningfulLineText.match(/^(?:\\d{1,4}\\.\\s*)?([A-ZÅÄÖÜ][a-zåäöü'-]+(?:\\s+[A-ZÅÄÖÜ][a-zåäöü'-]+)*)([\s\\S]*)/);
          let currentParagraphPreview = '';

          if (nameAndBioMatch) {
            const extractedName = nameAndBioMatch[1].trim();
            const afterNameText = nameAndBioMatch[2] ? nameAndBioMatch[2].trim() : '';
            currentParagraphPreview = `<b>${preserveHyperlinks(extractedName)}</b>${afterNameText ? ' ' + preserveHyperlinks(afterNameText) : ''}`;
            if (!person.name) nameForHeader = preserveHyperlinks(extractedName); // Update header name if not already set
          } else {
            if (person.name && firstMeaningfulLineText.startsWith(person.name)) {
              currentParagraphPreview = `<b>${preserveHyperlinks(person.name)}</b>${preserveHyperlinks(firstMeaningfulLineText.substring(person.name.length))}`;
            } else {
              currentParagraphPreview = preserveHyperlinks(firstMeaningfulLineText);
            }
          }

          // Add a few more lines of this first paragraph for context, joined by <br>
          let linesInCurrentPara = 0;
          for (let i = firstMeaningfulLineIndex + 1; i < rawLines.length && linesInCurrentPara < 2; i++) {
            if (rawLines[i].trim() === '') break; // Stop if empty line (new paragraph signal)
            currentParagraphPreview += '<br>' + preserveHyperlinks(rawLines[i]);
            linesInCurrentPara++;
          }
          detailsHtml += `<div>${currentParagraphPreview}</div>`;

          // Check for a subsequent distinct paragraph for a small preview
          let nextParaStartIndex = firstMeaningfulLineIndex + linesInCurrentPara + 1;
          while (nextParaStartIndex < rawLines.length && rawLines[nextParaStartIndex].trim() === '') {
            nextParaStartIndex++; // Skip empty lines
          }

          if (nextParaStartIndex < rawLines.length) {
            // Show a snippet of the next paragraph (e.g., first few words)
            const nextParaSnippet = preserveHyperlinks(rawLines[nextParaStartIndex].split(' ').slice(0, 10).join(' ')) + (rawLines[nextParaStartIndex].split(' ').length > 10 || rawLines.length > nextParaStartIndex + 1 ? '...' : '');
            detailsHtml += `<div style="margin-left: 15px; margin-top: 0.5em; font-style: italic; color: #555;">${nextParaSnippet}</div>`;
          }

        } else if (person.name) { // Fallback if raw line processing fails but name exists
          detailsHtml = `<div><b>${nameForHeader}</b></div>`;
        } else {
          detailsHtml = '<div>No displayable content.</div>';
        }
      } else if (person.name) { // Fallback if no person.raw but name exists
        detailsHtml = `<div><b>${nameForHeader}</b></div>`;
      } else {
        detailsHtml = '<div>No data available for this person.</div>';
      }

      html += `
        <div class="parsed-person">
          <div class="parsed-person-header">
            <span class="parsed-person-number">${person.number}</span>
            <span class="parsed-person-name" style="margin-left: 8px; font-weight:bold;">${nameForHeader}</span>
          </div>
          <div class="parsed-person-details">
            ${detailsHtml}
          </div>
        </div>
      `;
    });

    // Update the data preview
    dataPreview.innerHTML = html;
  }

  // Add view toggle functionality
  const rawViewBtn = document.getElementById('raw-view-btn');
  const parsedViewBtn = document.getElementById('parsed-view-btn');

  rawViewBtn.addEventListener('click', function() {
    rawViewBtn.classList.add('active');
    parsedViewBtn.classList.remove('active');
    displayRawExtractedText();
  });

  parsedViewBtn.addEventListener('click', function() {
    parsedViewBtn.classList.add('active');
    rawViewBtn.classList.remove('active');
    if (window.parsedGenealogyData) {
      // Use the tab-switch version that doesn't trigger application
      displayParsedDataForTabSwitch(window.parsedGenealogyData);
    }
  });

  // Display raw extracted text
  function displayRawExtractedText() {
    if (!window.rawExtractedText) {
      dataPreview.innerHTML = '<p class="no-data-message">No data extracted yet</p>';
      return;
    }
    
    let html = '<div class="raw-text-view">';
    
    // Check if rawExtractedText is an array (for old PDF logic) or a single object/string (for new OCR logic)
    if (Array.isArray(window.rawExtractedText)) {
      window.rawExtractedText.forEach(page => {
        html += `
          <div class="raw-text-page">
            <div class="raw-text-page-header">Page ${page.pageNumber}</div>
            <div class="raw-text-content">${page.text}</div>
          </div>
        `;
      });
    } else if (window.rawExtractedText && typeof window.rawExtractedText.text === 'string') { // Assuming OCR might put text in rawExtractedText.text
      html += `
        <div class="raw-text-page">
          <div class="raw-text-page-header">Extracted Text (OCR)</div>
          <div class="raw-text-content">${window.rawExtractedText.text}</div>
        </div>
      `;
    } else if (typeof window.rawExtractedText === 'string') { // If rawExtractedText is just the string
       html += `
        <div class="raw-text-page">
          <div class="raw-text-page-header">Extracted Text (OCR)</div>
          <div class="raw-text-content">${window.rawExtractedText}</div>
        </div>
      `;
    }
    
    html += '</div>';
    dataPreview.innerHTML = html;
  }

  // Hide the Apply to Template button since we've automated the process
  if (applyDataBtn) {
    applyDataBtn.style.display = 'none';
  }

  // Modify the existing displayParsedData function to handle view state
  function displayParsedData(parsedData) {
    if (!parsedData.persons || parsedData.persons.length === 0) {
      dataPreview.innerHTML = '<p class="no-data-message">No genealogy data could be extracted. Please try another file or format.</p>';
      extractionStatus.textContent = 'No genealogy data found';
      updateProgressBar(0);
      return;
    }

    // Extraction complete
    updateProgressBar(100);
    extractionStatus.textContent = 'Data extraction complete!';

    // Only update the display if we're in parsed view
    if (parsedViewBtn.classList.contains('active')) {
      displayParsedDataForTabSwitch(parsedData);
    }

    // Store the parsed data for later use
    window.parsedGenealogyData = parsedData;
    
    // Don't apply data again if it's already been applied for this extraction
    if (!currentDataApplied) {
      // Automatically apply data to template after extraction completes
      extractionStatus.textContent = 'Automatically applying to template...';
      
      // Small delay to allow the status update to render
      setTimeout(() => {
        // Call the applyDataToTemplate function directly
        applyDataToTemplate();
        currentDataApplied = true; // Mark as applied
      }, 300);
    }
  }
  
  // Extract the applyDataBtn logic into a function that can be called directly
  function applyDataToTemplate() {
    if (!window.parsedGenealogyData) {
      alert('No data available to apply to the template.');
      return;
    }

    const data = window.parsedGenealogyData;
    const legacySheet = document.querySelector('.legacy-sheet');

    // Update the generation title only if found
    if (data.generation) {
      const generationTitle = document.querySelector('.generation-title');
      if (generationTitle) {
        generationTitle.textContent = data.generation;
      }
    }

    // Clear existing person entries ONLY if this is the first time we're applying data
    // This removes the static template entries but keeps our dynamically added entries
    if (!staticTemplateCleared) {
      const existingEntries = legacySheet.querySelectorAll('.person-entry');
      const horizontalRule = document.querySelector('.horizontal-rule');

      // Remove all person entries and their dynamically added dividers
      existingEntries.forEach(entry => {
        const nextSibling = entry.nextElementSibling;
        // Only remove the next sibling if it's NOT part of the main footer structure
        if (nextSibling && !nextSibling.classList.contains('footer-line') && !nextSibling.classList.contains('footer')) {
          nextSibling.remove();
        }
        entry.remove(); // Removes the .person-entry
      });

      // Mark that we've cleared the static template
      staticTemplateCleared = true;
    }

    // Process persons one at a time - for our one-person-at-a-time workflow
    // We'll always process just the first person in the array
    if (data.persons && data.persons.length > 0) {
      const person = data.persons[0]; // Get the first person in the array
      
      // Create person entry
      const personEntry = document.createElement('div');
      personEntry.className = 'person-entry';

      // Person number
      const personNumber = document.createElement('div');
      personNumber.className = 'person-number';
      personNumber.textContent = person.number;

      // Person content
      const personContent = document.createElement('div');
      personContent.className = 'person-content';

      // Process person.raw for main and subsequent paragraphs
      if (person.raw) {
        const rawLines = person.raw.split(/\r?\n/); // Keep empty lines for now
        let lineIndex = 0;

        // Find the first non-empty line for the main biographical entry
        while (lineIndex < rawLines.length && rawLines[lineIndex].trim() === '') {
          lineIndex++;
        }

        if (lineIndex < rawLines.length) {
          // The first line of actual content, which includes the number, name, and start of bio
          // The number itself is handled by personNumber div, so we focus on name and bio from person.raw

          const firstContentLineText = rawLines[lineIndex];
          const nameAndBioMatch = firstContentLineText.match(/^(?:\d{1,4}\.\s*)?([A-ZÅÄÖÜ][a-zåäöü'-]+(?:\s+[A-ZÅÄÖÜ][a-zåäöü'-]+)*)([\s\S]*)/);
          
          let mainParagraphHTML = '';
          if (nameAndBioMatch) {
            const name = nameAndBioMatch[1].trim();
            const afterName = nameAndBioMatch[2] ? nameAndBioMatch[2].trim() : '';
            mainParagraphHTML = `<b>${preserveHyperlinks(name)}</b>${afterName ? ' ' + preserveHyperlinks(afterName) : ''}`;
          } else {
            // Fallback if regex doesn't match (e.g. no clear name pattern on the first line)
            mainParagraphHTML = preserveHyperlinks(firstContentLineText);
          }
          lineIndex++; // Move to the line AFTER the first content line

          // Continue adding subsequent lines to the main paragraph
          // until an empty line, end of lines, or a marriage indicator line.
          while (lineIndex < rawLines.length) {
            const currentRawLine = rawLines[lineIndex];
            const currentLineTrimmed = currentRawLine.trim();

            if (currentLineTrimmed === '') {
              // Empty line signifies the end of the main paragraph.
              break;
            }

            // Define a regex to check for marriage indicators at the start of a trimmed line.
            const firstName = person.name ? person.name.split(' ')[0] : '[A-ZÅÄÖÜ][a-zåäöü\'-]+'; // Use person's first name or a general pattern
            const marriageIndicatorRegex = new RegExp(
              `^(${firstName}\\s+married\\b|Married\\b|(He|She)\\s+married\\b)`, 'i'
            );

            if (marriageIndicatorRegex.test(currentLineTrimmed)) {
              // This line appears to start marriage details. Stop the main paragraph here.
              break;
            }

            mainParagraphHTML += ' ' + preserveHyperlinks(currentRawLine.trim());
            lineIndex++;
          }

          const mainParagraphDiv = document.createElement('div');
          mainParagraphDiv.innerHTML = mainParagraphHTML;
          personContent.appendChild(mainParagraphDiv);

          // Process subsequent distinct paragraphs (indented)
          while (lineIndex < rawLines.length) {
            // Skip any further empty lines between paragraphs
            while (lineIndex < rawLines.length && rawLines[lineIndex].trim() === '') {
              lineIndex++;
            }

            if (lineIndex < rawLines.length) { // Found start of a new potential sub-paragraph
              // IMPROVED APPROACH: Collect all text for this sub-paragraph in a single string
              let allSubParagraphText = '';
              
              // Gather all lines for this sub-paragraph first
              while (lineIndex < rawLines.length && rawLines[lineIndex].trim() !== '') {
                // Handle explicit user-inserted breaks
                const currentLine = rawLines[lineIndex].replace(/%%USER_BREAK%%/g, '<br>');
                
                // If we already have content and this line doesn't start with <br>, add space
                if (allSubParagraphText && !currentLine.trim().startsWith('<br>')) {
                  allSubParagraphText += ' ';
                }
                
                // Add the line content
                allSubParagraphText += currentLine.trim();
                lineIndex++;
              }
              
              // Only create the paragraph if we have content
              if (allSubParagraphText) {
                // Use our new function to create the paragraph with controls
                const paraDiv = createEditableSubparagraph(allSubParagraphText, person, personContent);
                personContent.appendChild(paraDiv);
              }
            }
          }
        }
      }

      // Assemble the person entry
      personEntry.appendChild(personNumber);
      personEntry.appendChild(personContent);

      // Add to the document, inserting before the main footer text container
      legacySheet.insertBefore(personEntry, document.querySelector('.footer'));

      // Add divider
      const divider = document.createElement('hr');
      divider.className = 'entry-divider';
    
      // Make divider removable with a click
      divider.style.cursor = 'pointer'; // Change cursor to show it's clickable
      
      // Add a subtle visual indicator on hover
      divider.addEventListener('mouseover', function() {
        this.style.opacity = '0.5';
        this.title = 'Click to remove this divider';
      });
      
      divider.addEventListener('mouseout', function() {
        this.style.opacity = '1';
      });
      
      // Add click event to remove the divider
      divider.addEventListener('click', function() {
        // Remove the divider with a fade-out effect
        this.style.transition = 'opacity 0.3s ease';
        this.style.opacity = '0';
        
        // After fade-out, remove the element and adjust spacing
        setTimeout(() => {
          this.remove();
        }, 300);
      });
      
      legacySheet.insertBefore(divider, document.querySelector('.footer'));
    }

    // No need to update the Apply button status since it's hidden now
    
    // Show Next Person button - adjust its styling now that Apply button is hidden
    nextPersonBtn.style.display = 'inline-block';
    nextPersonBtn.style.marginTop = '20px'; // Ensure proper spacing above the button
    
    // Update extraction status
    extractionStatus.textContent = 'Entry applied! Ready for next person.';
  }

  // Helper to preserve hyperlinks if present in the text
  function preserveHyperlinks(text) {
    // If the text already contains <a href=...>, return as is
    if (/<a\s+href=/.test(text)) return text;
    // If the text contains [Name](url) style, convert to <a href="url">Name</a>
    return text.replace(/\[([^\]]+)\]\(([^\)]+)\)/g, '<a href="$2" target="_blank">$1</a>');
  }

  // Store original template data when the page loads
  const originalTemplateData = {
    generationTitle: document.querySelector('.generation-title').textContent,
    personEntries: Array.from(document.querySelectorAll('.person-entry')).map(entry => entry.outerHTML),
    dividers: Array.from(document.querySelectorAll('.entry-divider')).map(divider => divider.outerHTML)
  };

  // Apply the extracted data to the template - keep this event handler for backwards compatibility
  applyDataBtn.addEventListener('click', function() {
    applyDataToTemplate();
  });

  // Add event listener for the unapply button
  const unapplyBtn = document.getElementById('unapply-data-btn');
  unapplyBtn.addEventListener('click', function() {
    // Restore original template data
    const legacySheet = document.querySelector('.legacy-sheet');
    const generationTitle = document.querySelector('.generation-title');
    const horizontalRule = document.querySelector('.horizontal-rule');
    const footerLine = document.querySelector('.footer-line');
    
    // Restore generation title
    generationTitle.textContent = originalTemplateData.generationTitle;
    
    // Remove all current person entries and dividers
    const currentEntries = legacySheet.querySelectorAll('.person-entry');
    const currentDividers = legacySheet.querySelectorAll('.entry-divider');
    
    currentEntries.forEach(entry => entry.remove());
    currentDividers.forEach(divider => divider.remove());
    
    // Insert original person entries and dividers
    const tempDiv = document.createElement('div');
    
    for (let i = 0; i < originalTemplateData.personEntries.length; i++) {
      // Insert person entry
      tempDiv.innerHTML = originalTemplateData.personEntries[i];
      const personEntry = tempDiv.firstChild;
      legacySheet.insertBefore(personEntry, footerLine);
      
      // Insert divider if not the last entry
      if (i < originalTemplateData.personEntries.length - 1) {
        tempDiv.innerHTML = originalTemplateData.dividers[i];
        const divider = tempDiv.firstChild;
        legacySheet.insertBefore(divider, footerLine);
      }
    }
    
    // Reset button states
    applyDataBtn.textContent = 'Apply to Template';
    applyDataBtn.classList.remove('applied-button');
    applyDataBtn.disabled = false;
    
    // Hide unapply button
    unapplyBtn.style.display = 'none';
    
    // Reset staticTemplateCleared flag
    staticTemplateCleared = false;
    
    extractionStatus.textContent = 'Reverted to original template data';
  });

  // Helper function to update person.raw with user-inserted breaks
  function updatePersonRawWithUserBreaks(personObject, personContentElement) {
    // Find all indented paragraphs in this person's content
    const indentedParagraphs = personContentElement.querySelectorAll('.indented-paragraph');
    
    // Placeholder for a more detailed implementation
    console.log(`Updating breaks for person ${personObject.number}`);
    console.log(`Found ${indentedParagraphs.length} indented paragraphs`);
    
    // In a real implementation, we would:
    // 1. Get the current HTML content of each paragraph
    // 2. Convert <br> tags to %%USER_BREAK%% markers
    // 3. Update the person.raw with these breaks
  }

  async function extractTextFromImageWithOCR(imageFile) {
    try {
      extractionStatus.textContent = 'Preparing image for OCR...';
      updateProgressBar(10);
      // Disable buttons during OCR - already handled by caller, but good for safety
      extractDataBtn.disabled = true; 
      removeImageBtn.disabled = true;

      const { data: { text } } = await Tesseract.recognize(
        imageFile,
        'eng', // Language - English. Add more or make configurable if needed.
        {
          logger: m => {
            console.log(m); // Log Tesseract progress
            if (m.status === 'recognizing text') {
              // Tesseract progress is 0 to 1, scale it for our progress bar (e.g., from 10% to 70%)
              const tesseractProgress = m.progress * 100; // Convert to percentage
              updateProgressBar(10 + (tesseractProgress * 0.6)); // Scale to fit 10-70% range
            } else if (m.status === 'loading language model') {
              updateProgressBar(15);
              extractionStatus.textContent = 'Loading language model...';
            } else if (m.status === 'initializing api') {
              updateProgressBar(20);
              extractionStatus.textContent = 'Initializing OCR API...';
            } else if (m.status === 'recognizing text') {
               extractionStatus.textContent = 'Recognizing text in image...';
            }
          }
        }
      );

      updateProgressBar(70); // OCR part done
      extractionStatus.textContent = 'Analyzing extracted text...';
      
      // Store raw text for display
      window.rawExtractedText = text; 
      
      processExtractedText(text);

    } catch (error) {
      console.error('Error during OCR processing:', error);
      extractionStatus.textContent = 'Error during OCR. Please try another image or check console.';
      updateProgressBar(0);
      // No specific data to parse on error, so apply button remains disabled.
    } finally {
      // Ensure buttons are re-enabled even if an unexpected error occurs within Tesseract.recognize not caught by its own try/catch
      // This is also handled by the caller's finally block, providing a fallback.
      extractDataBtn.disabled = false;
      extractDataBtn.textContent = 'Extract Data';
      removeImageBtn.disabled = false;
    }
  }

  // Only add the specific transition property needed for fade effect
  const fadeTransitionStyle = document.createElement('style');
  fadeTransitionStyle.textContent = `
    .entry-divider {
      transition: opacity 0.3s ease;
      cursor: pointer;
    }
    
    /* Styles for subparagraph hover controls - Apple UI Style */
    .indented-paragraph {
      position: relative;
      transition: background-color 0.2s ease;
      padding: 4px 6px;
      border-radius: 3px;
    }
    
    .indented-paragraph:hover {
      background-color: rgba(240, 240, 240, 0.5);
    }
    
    .paragraph-controls {
      position: absolute;
      right: 10px;
      top: -45px;
      display: none;
      background-color: rgba(248, 248, 248, 0.95);
      border-radius: 14px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.2);
      padding: 0px;
      z-index: 100;
      flex-direction: row;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    /* Add the downward-pointing arrow */
    .paragraph-controls:after {
      content: '';
      position: absolute;
      bottom: -8px;
      left: 50%;
      width: 16px;
      height: 16px;
      background-color: rgba(248, 248, 248, 0.95);
      transform: translateX(-50%) rotate(45deg);
      box-shadow: 2px 2px 3px rgba(0,0,0,0.05);
      z-index: -1;
    }
    
    .indented-paragraph:hover .paragraph-controls {
      display: flex;
    }
    
    .paragraph-controls button {
      border: none;
      background: none;
      cursor: pointer;
      padding: 8px 12px;
      margin: 0;
      color: #000;
      font-size: 16px;
      display: flex;
      align-items: center;
      justify-content: center;
      min-width: 40px;
      text-align: center;
    }
    
    .paragraph-controls button:not(:last-child) {
      border-right: 1px solid rgba(0,0,0,0.1);
    }
    
    .paragraph-controls button:active {
      background-color: rgba(0,0,0,0.05);
    }
    
    .paragraph-controls button:first-child {
      border-top-left-radius: 14px;
      border-bottom-left-radius: 14px;
    }
    
    .paragraph-controls button:last-child {
      border-top-right-radius: 14px;
      border-bottom-right-radius: 14px;
    }
  `;
  document.head.appendChild(fadeTransitionStyle);

  // Update the paragraph controls functions to match Apple UI style
  function addParagraphControls(paraDiv, person, personContent) {
    // Create controls container
    const controls = document.createElement('div');
    controls.className = 'paragraph-controls';
    
    // Create edit button
    const editBtn = document.createElement('button');
    editBtn.className = 'edit-btn';
    editBtn.title = 'Edit paragraph'; // Add tooltip
    editBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M17 3a2.85 2.85 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"></path>
    </svg>`;
    editBtn.addEventListener('click', function(e) {
      e.stopPropagation(); // Prevent event bubbling
      paraDiv.focus(); // Focus the paragraph for editing
    });
    
    // Create bold button
    const boldBtn = document.createElement('button');
    boldBtn.className = 'bold-btn';
    boldBtn.title = 'Bold text';
    boldBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M6 4h8a4 4 0 0 1 4 4 4 4 0 0 1-4 4H6z"></path>
      <path d="M6 12h9a4 4 0 0 1 4 4 4 4 0 0 1-4 4H6z"></path>
    </svg>`;
    boldBtn.addEventListener('click', function(e) {
      e.stopPropagation(); // Prevent event bubbling
      e.preventDefault(); // Prevent default
      
      // Make sure the paragraph has focus
      paraDiv.focus();
      
      // Check if there's a selection within this paragraph
      const selection = window.getSelection();
      
      // If no selection or collapsed (just a cursor), select all paragraph text
      if (!selection.rangeCount || selection.isCollapsed) {
        console.log("No selection, selecting all text");
        const range = document.createRange();
        range.selectNodeContents(paraDiv);
        selection.removeAllRanges();
        selection.addRange(range);
      } else {
        // There is a selection, check if it's within our paragraph
        const range = selection.getRangeAt(0);
        const selectionParent = range.commonAncestorContainer;
        
        if (!paraDiv.contains(selectionParent)) {
          console.log("Selection outside paragraph, selecting all text");
          // Selection is outside our paragraph, select all paragraph content instead
          const newRange = document.createRange();
          newRange.selectNodeContents(paraDiv);
          selection.removeAllRanges();
          selection.addRange(newRange);
        } else {
          console.log("Selection within paragraph, using it");
          // Selection is already within our paragraph, use it
        }
      }
      
      // Now apply bold command to the current selection
      document.execCommand('bold', false, null);
      
      // Don't clear selection if user selected specific text
      // Only clear if we auto-selected the whole paragraph
      if (!selection.rangeCount || selection.getRangeAt(0).toString().trim() === paraDiv.textContent.trim()) {
        selection.removeAllRanges();
      }
      
      // Update the raw data with the formatting changes
      if (person && personContent) {
        updatePersonRawWithUserBreaks(person, personContent);
      }
    });
    
    // Create delete button
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'delete-btn';
    deleteBtn.title = 'Delete paragraph'; // Add tooltip
    deleteBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="3 6 5 6 21 6"></polyline>
      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
    </svg>`;
    deleteBtn.addEventListener('click', function(e) {
      e.stopPropagation(); // Prevent event bubbling
      
      // Fade out the paragraph
      paraDiv.style.transition = 'opacity 0.3s ease, margin 0.3s ease, padding 0.3s ease, height 0.3s ease';
      paraDiv.style.opacity = '0';
      paraDiv.style.height = '0';
      paraDiv.style.margin = '0';
      paraDiv.style.padding = '0';
      paraDiv.style.overflow = 'hidden';
      
      // Remove the paragraph after the animation
      setTimeout(() => {
        paraDiv.remove();
        
        // Update the raw data if needed
        if (person && personContent) {
          updatePersonRawWithRemovedParagraph(person, personContent, paraDiv);
        }
      }, 300);
    });
    
    // Add buttons to controls
    controls.appendChild(editBtn);
    controls.appendChild(boldBtn);
    controls.appendChild(deleteBtn);
    
    // Add controls to paragraph
    paraDiv.appendChild(controls);
  }

  // Update the paragraph creation code to add controls
  function createEditableSubparagraph(allSubParagraphText, person, personContent) {
    const paraDiv = document.createElement('div');
    paraDiv.className = 'indented-paragraph';
    
    // Process and set the content
    paraDiv.innerHTML = preserveHyperlinks(allSubParagraphText);
    
    // Make it editable
    paraDiv.setAttribute('contenteditable', 'true');
    paraDiv.classList.add('editable-subparagraph');
    
    // Add event listener for 'Enter' key to handle user breaks
    paraDiv.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        document.execCommand('insertLineBreak');
        // Update the raw data with the breaks
        updatePersonRawWithUserBreaks(person, personContent);
      }
    });
    
    // Add event listeners to check if selection is bold
    paraDiv.addEventListener('mouseup', checkSelectionFormatting);
    paraDiv.addEventListener('keyup', checkSelectionFormatting);
    paraDiv.addEventListener('click', checkSelectionFormatting);
    paraDiv.addEventListener('focus', checkSelectionFormatting);
    
    // Also check formatting on mousedown to handle selection by double/triple clicking
    paraDiv.addEventListener('mousedown', function(e) {
      // Slight delay to allow browser to complete the selection process
      setTimeout(checkSelectionFormatting, 10);
    });
    
    // Add listener for selection changes within the document
    document.addEventListener('selectionchange', function(e) {
      // Only process if the paragraph has focus
      if (document.activeElement === paraDiv) {
        checkSelectionFormatting();
      }
    });
    
    // Function to check if current selection is within bold text
    function checkSelectionFormatting() {
      const controls = paraDiv.querySelector('.paragraph-controls');
      if (!controls) return;
      
      const boldBtn = controls.querySelector('.bold-btn');
      if (!boldBtn) return;
      
      try {
        // Only check if this paragraph is active
        if (document.activeElement !== paraDiv) {
          boldBtn.classList.remove('active-format');
          return;
        }
        
        // Simplest check - does the current selection or cursor position have bold formatting
        const isBold = document.queryCommandState('bold');
        
        if (isBold) {
          boldBtn.classList.add('active-format');
        } else {
          boldBtn.classList.remove('active-format');
        }
      } catch (e) {
        console.error('Error checking bold state:', e);
      }
    }
    
    // Add hover controls for editing/deleting
    addParagraphControls(paraDiv, person, personContent);
    
    return paraDiv;
  }

  // Helper function to update raw data when paragraph is removed
  function updatePersonRawWithRemovedParagraph(personObject, personContentElement, removedParagraph) {
    console.log(`Paragraph removed from person ${personObject.number}`);
    // In a real implementation, we would update personObject.raw to remove the paragraph
    // For now, this is just a placeholder
  }
});

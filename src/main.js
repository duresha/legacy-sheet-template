// main.js
// Import libraries for PDF generation with download animation
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { saveAs } from 'file-saver';
import * as pdfjsLib from 'pdfjs-dist';

// Configure the PDF.js worker to use a local file
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./pdf.worker.min.mjs', import.meta.url).href;

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
    extractionStatus.textContent = 'Image loaded. Ready for extraction.'; // Update status
    
    applyDataBtn.disabled = true; // Ensure apply button is disabled until data is parsed
    extractDataBtn.style.display = 'block'; // Show Extract Data button
    removeImageBtn.style.display = 'inline-block'; // Ensure remove button is visible
    
    // DO NOT automatically extract: extractTextFromImageWithOCR(file);
  }

  // Add event listener for the new Extract Data button
  extractDataBtn.addEventListener('click', function() {
    if (currentImageFile) {
      console.log('Extract Data button clicked for:', currentImageFile.name);
      extractionStatus.textContent = 'Starting OCR process...';
      extractDataBtn.disabled = true;
      extractDataBtn.textContent = 'Extracting...';
      removeImageBtn.disabled = true; // Disable remove button during extraction

      extractTextFromImageWithOCR(currentImageFile).finally(() => {
        // Re-enable buttons or update text after OCR is done (success or fail)
        extractDataBtn.disabled = false;
        extractDataBtn.textContent = 'Extract Data';
        removeImageBtn.disabled = false;
        // applyDataBtn will be enabled in displayParsedData if successful
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
      displayParsedData(window.parsedGenealogyData);
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
      let html = '';

      // Generation title
      if (parsedData.generation) {
        html += `<div class="parsed-item"><strong>Generation:</strong> ${parsedData.generation}</div>`;
      }

      // Persons
      parsedData.persons.forEach(person => {
        // Split the first line into bold name and normal text
        let firstLine = '';
        let restLines = '';
        if (person.raw) {
          const lines = person.raw.split(/\r?\n/).filter(line => line.trim() !== '');
          if (lines.length > 0) {
            // Try to extract the name and the rest of the first line
            const firstLineMatch = lines[0].match(/^(\d{1,4})\.\s*([^.,]+)([\s\S]*)/);
            if (firstLineMatch) {
              // number, name, rest of first line
              const name = firstLineMatch[2].trim();
              const afterName = firstLineMatch[3] ? firstLineMatch[3] : '';
              firstLine = `<b>${preserveHyperlinks(name)}</b>${preserveHyperlinks(afterName)}`;
            } else {
              firstLine = preserveHyperlinks(lines[0]);
            }
            restLines = lines.slice(1).map(l => `<div class='indented-paragraph'>${preserveHyperlinks(l)}</div>`).join('');
          }
        }
        html += `
          <div class="parsed-person">
            <div class="parsed-person-header">
              <span class="parsed-person-number">${person.number}</span>
            </div>
            <div class="parsed-person-details">
              <div>${firstLine}</div>
              ${restLines}
            </div>
          </div>
        `;
      });

      // Update the data preview
      dataPreview.innerHTML = html;
    }

    // Enable the apply data button
    applyDataBtn.disabled = false;

    // Store the parsed data for later use
    window.parsedGenealogyData = parsedData;
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

  // Apply the extracted data to the template
  applyDataBtn.addEventListener('click', function() {
    if (!window.parsedGenealogyData) {
      alert('No data available to apply to the template.');
      return;
    }

    const data = window.parsedGenealogyData;

    // Update the generation title only if found
    if (data.generation) {
      const generationTitle = document.querySelector('.generation-title');
      if (generationTitle) {
        generationTitle.textContent = data.generation;
      }
    }

    // Clear existing person entries
    const legacySheet = document.querySelector('.legacy-sheet');
    const existingEntries = legacySheet.querySelectorAll('.person-entry');
    const horizontalRule = document.querySelector('.horizontal-rule');

    // Remove all person entries but keep the first horizontal rule
    existingEntries.forEach(entry => {
      entry.nextElementSibling?.remove(); // Remove divider if present
      entry.remove();
    });

    // Create new person entries
    data.persons.forEach((person, index) => {
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

      // Split the first line into bold name and normal text
      if (person.raw) {
        const lines = person.raw.split(/\r?\n/).filter(line => line.trim() !== '');
        if (lines.length > 0) {
          const firstLineMatch = lines[0].match(/^(\d{1,4})\.\s*([^.,]+)([\s\S]*)/);
          let firstLine = '';
          if (firstLineMatch) {
            const name = firstLineMatch[2].trim();
            const afterName = firstLineMatch[3] ? firstLineMatch[3] : '';
            firstLine = `<b>${preserveHyperlinks(name)}</b>${preserveHyperlinks(afterName)}`;
          } else {
            firstLine = preserveHyperlinks(lines[0]);
          }
          const firstLineDiv = document.createElement('div');
          firstLineDiv.innerHTML = firstLine;
          personContent.appendChild(firstLineDiv);
          // Indent subsequent paragraphs
          for (let i = 1; i < lines.length; i++) {
            const paraDiv = document.createElement('div');
            paraDiv.className = 'indented-paragraph';
            paraDiv.innerHTML = preserveHyperlinks(lines[i]);
            personContent.appendChild(paraDiv);
          }
        }
      }

      // Assemble the person entry
      personEntry.appendChild(personNumber);
      personEntry.appendChild(personContent);

      // Add to the document
      legacySheet.insertBefore(personEntry, document.querySelector('.footer-line'));

      // Add divider after each person (except the last one)
      if (index < data.persons.length - 1) {
        const divider = document.createElement('hr');
        divider.className = 'entry-divider';
        legacySheet.insertBefore(divider, document.querySelector('.footer-line'));
      }
    });

    // Update button states
    applyDataBtn.textContent = 'Applied to Template';
    applyDataBtn.classList.add('applied-button');
    applyDataBtn.disabled = true;

    // Show unapply button
    const unapplyBtn = document.getElementById('unapply-data-btn');
    unapplyBtn.style.display = 'block';

    // Update extraction status
    extractionStatus.textContent = `Applied ${data.persons.length} person entries to the template`;
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
    
    extractionStatus.textContent = 'Reverted to original template data';
  });

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
});

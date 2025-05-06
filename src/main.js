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
  const fileInput = document.getElementById('pdf-upload');
  const uploadPreview = document.getElementById('upload-preview');
  const pdfFilename = document.getElementById('pdf-filename');
  const extractionProgress = document.getElementById('extraction-progress');
  const extractionStatus = document.getElementById('extraction-status');
  const dataPreview = document.getElementById('data-preview');
  const applyDataBtn = document.getElementById('apply-data-btn');
  const removePdfBtn = document.getElementById('remove-pdf-btn');

  // Handle remove PDF button click
  removePdfBtn.addEventListener('click', function() {
    // Clear the file input
    fileInput.value = '';
    
    // Hide the upload preview
    uploadPreview.style.display = 'none';
    
    // Reset progress bar
    extractionProgress.style.width = '0%';
    
    // Reset extraction status
    extractionStatus.textContent = 'Waiting for PDF...';
    
    // Clear extracted data
    dataPreview.innerHTML = '<p class="no-data-message">No data extracted yet</p>';
    
    // Disable apply button
    applyDataBtn.disabled = true;
    
    // Clear any stored data
    window.parsedGenealogyData = null;
    
    // Show the upload area again
    uploadArea.style.display = 'flex';
    
    console.log('PDF removed and interface reset');
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
    // Only trigger if we're clicking on the upload area itself or its direct children
    if (e.target === uploadArea || e.target.parentNode === uploadArea ||
        e.target.classList.contains('upload-icon') || 
        e.target.classList.contains('browse-text')) {
      
      console.log('Upload area clicked, triggering file input click');
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
    
    // Check if the file is a PDF
    if (file.type !== 'application/pdf') {
      alert('Please upload a PDF file.');
      return;
    }
    
    // Hide upload area
    uploadArea.style.display = 'none';
    
    // Show upload preview
    uploadPreview.style.display = 'block';
    
    // Set filename
    pdfFilename.textContent = file.name;
    
    // Reset progress bar
    extractionProgress.style.width = '0%';
    
    // Update extraction status
    extractionStatus.textContent = 'Processing PDF...';
    
    // Start with initial progress
    updateProgressBar(5);
    
    // Extract text from PDF
    extractTextFromPdf(file);
  }

  // Update progress bar
  function updateProgressBar(percent) {
    extractionProgress.style.width = percent + '%';
  }

  // Extract text from PDF
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

  // Process the extracted text to identify genealogy data
  function processExtractedText(text) {
    extractionStatus.textContent = 'Analyzing genealogy data...';
    updateProgressBar(75);
    
    // Initialize our parsed data structure
    const parsedData = {
      generation: '',
      persons: []
    };
    
    // Try to identify generation title
    const generationMatch = text.match(/(\w+)\s+Generation/i);
    if (generationMatch) {
      parsedData.generation = generationMatch[0];
    }
    
    updateProgressBar(80);
    
    // Regular expressions for identifying person entries
    const personRegex = /(\d+)\.\s+([A-Za-z\s]+),\s+(son|daughter)\s+of\s+([A-Za-z\s]+)\s+and\s+([A-Za-z\s]+),\s+was\s+born\s+(on|in)\s+(.+?)(\s+in\s+(.+?))?(\s+and\s+died\s+(on|in)\s+(.+?)(\s+in\s+(.+?))?)?/g;
    const marriageRegex = /married\s+([A-Za-z\s]+)\s+(on|in)\s+(.+?)(\s+in\s+(.+?))?/g;
    const childrenRegex = /Children\s+from\s+this\s+marriage\s+were:/i;
    const childRegex = /(i|ii|iii|iv|v|vi|vii|viii|ix|x|xi|xii)\.\s+([A-Za-z\s"]+),\s+born\s+(on|in)\s+(.+?)(\s+in\s+(.+?))?/g;
    
    // Look for person entries
    let match;
    let currentText = text;
    let lastIndex = 0;
    
    updateProgressBar(85);
    
    // Extract person entries
    while ((match = personRegex.exec(currentText)) !== null) {
      const personNumber = match[1];
      const personName = match[2].trim();
      const relationship = match[3]; // son or daughter
      const father = match[4].trim();
      const mother = match[5].trim();
      const birthDatePrefix = match[6]; // on or in
      const birthDateInfo = match[7].trim();
      const birthPlace = match[9] ? match[9].trim() : '';
      const hasDeath = match[10] ? true : false;
      const deathDateInfo = hasDeath ? match[12].trim() : '';
      const deathPlace = hasDeath && match[14] ? match[14].trim() : '';
      
      // Create a person object
      const person = {
        number: personNumber,
        name: personName,
        relationship: relationship,
        father: father,
        mother: mother,
        birthInfo: `${birthDatePrefix} ${birthDateInfo}${birthPlace ? ' in ' + birthPlace : ''}`,
        deathInfo: hasDeath ? `${match[11]} ${deathDateInfo}${deathPlace ? ' in ' + deathPlace : ''}` : '',
        marriages: [],
        children: []
      };
      
      // Look for marriage info
      const personTextEnd = currentText.indexOf(personNumber + 1 + '.', match.index);
      const personText = personTextEnd !== -1 ? 
        currentText.substring(match.index, personTextEnd) : 
        currentText.substring(match.index);
      
      let marriageMatch;
      while ((marriageMatch = marriageRegex.exec(personText)) !== null) {
        const spouseName = marriageMatch[1].trim();
        const marriageDatePrefix = marriageMatch[2]; // on or in
        const marriageDateInfo = marriageMatch[3].trim();
        const marriagePlace = marriageMatch[5] ? marriageMatch[5].trim() : '';
        
        person.marriages.push({
          spouse: spouseName,
          info: `${marriageDatePrefix} ${marriageDateInfo}${marriagePlace ? ' in ' + marriagePlace : ''}`
        });
      }
      
      // Look for children
      if (childrenRegex.test(personText)) {
        let childMatch;
        while ((childMatch = childRegex.exec(personText)) !== null) {
          const childMarker = childMatch[1];
          const childName = childMatch[2].trim();
          const birthDatePrefix = childMatch[3]; // on or in
          const birthDateInfo = childMatch[4].trim();
          const birthPlace = childMatch[6] ? childMatch[6].trim() : '';
          
          person.children.push({
            marker: childMarker,
            name: childName,
            birthInfo: `${birthDatePrefix} ${birthDateInfo}${birthPlace ? ' in ' + birthPlace : ''}`
          });
        }
      }
      
      // Add the person to our parsed data
      parsedData.persons.push(person);
      
      lastIndex = match.index + match[0].length;
    }
    
    updateProgressBar(95);
    
    // Display the parsed data
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
    
    window.rawExtractedText.forEach(page => {
      html += `
        <div class="raw-text-page">
          <div class="raw-text-page-header">Page ${page.pageNumber}</div>
          <div class="raw-text-content">${page.text}</div>
        </div>
      `;
    });
    
    html += '</div>';
    dataPreview.innerHTML = html;
  }

  // Modify the existing displayParsedData function to handle view state
  function displayParsedData(parsedData) {
    if (parsedData.persons.length === 0) {
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
        html += `
          <div class="parsed-person">
            <div class="parsed-person-header">
              <span class="parsed-person-number">${person.number}</span>
              <span class="parsed-person-name">${person.name}</span>
            </div>
            <div class="parsed-person-details">
              <div><strong>Relationship:</strong> ${person.relationship} of ${person.father} and ${person.mother}</div>
              <div><strong>Birth:</strong> ${person.birthInfo}</div>
              ${person.deathInfo ? `<div><strong>Death:</strong> ${person.deathInfo}</div>` : ''}
              
              ${person.marriages.length > 0 ? 
                `<div class="parsed-marriages">
                  <strong>Marriages:</strong>
                  ${person.marriages.map(m => `<div>Married ${m.spouse} ${m.info}</div>`).join('')}
                </div>` : ''}
              
              ${person.children.length > 0 ? 
                `<div class="parsed-children">
                  <strong>Children:</strong>
                  <ul>
                    ${person.children.map(c => `<li>${c.marker}. ${c.name}, born ${c.birthInfo}</li>`).join('')}
                  </ul>
                </div>` : ''}
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
    
    // Update the generation title
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
    let lastElement = horizontalRule;
    
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
      
      // Person name and basic info
      const personName = document.createElement('span');
      personName.className = 'person-name';
      personName.textContent = person.name;
      
      const relationshipText = document.createElement('span');
      relationshipText.className = 'relationship-text';
      
      // Create father and mother links
      const fatherLink = document.createElement('a');
      fatherLink.href = '#';
      fatherLink.textContent = person.father;
      
      const motherLink = document.createElement('a');
      motherLink.href = '#';
      motherLink.textContent = person.mother;
      
      // Build the relationship text
      let relationshipHTML = `, ${person.relationship} of `;
      relationshipText.innerHTML = `${relationshipHTML}${fatherLink.outerHTML} and ${motherLink.outerHTML}, was <br> born ${person.birthInfo}`;
      
      if (person.deathInfo) {
        relationshipText.innerHTML += ` and died ${person.deathInfo}`;
      }
      
      personContent.appendChild(personName);
      personContent.appendChild(relationshipText);
      
      // Add marriages if present
      if (person.marriages.length > 0) {
        person.marriages.forEach(marriage => {
          const marriageDetails = document.createElement('div');
          marriageDetails.className = 'marriage-details';
          
          const spouseLink = document.createElement('a');
          spouseLink.href = '#';
          spouseLink.textContent = marriage.spouse;
          
          marriageDetails.innerHTML = `${person.name.split(' ')[0]} married ${spouseLink.outerHTML} ${marriage.info}.`;
          personContent.appendChild(marriageDetails);
        });
      }
      
      // Add children if present
      if (person.children.length > 0) {
        const childrenHeader = document.createElement('div');
        childrenHeader.className = 'children-header';
        childrenHeader.textContent = 'Children from this marriage were:';
        personContent.appendChild(childrenHeader);
        
        const childrenList = document.createElement('div');
        childrenList.className = 'children-list';
        
        person.children.forEach((child, childIndex) => {
          const childEntry = document.createElement('div');
          childEntry.className = 'child-entry';
          
          const childMarker = document.createElement('span');
          childMarker.className = 'child-marker';
          childMarker.textContent = child.marker + '.';
          
          const childName = document.createElement('span');
          childName.className = 'child-name';
          childName.textContent = child.name;
          
          childEntry.appendChild(childMarker);
          childEntry.appendChild(childName);
          childEntry.innerHTML += `, born ${child.birthInfo}.`;
          
          childrenList.appendChild(childEntry);
        });
        
        personContent.appendChild(childrenList);
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
});

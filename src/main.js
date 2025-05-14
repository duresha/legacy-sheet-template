// main.js
// Import libraries for PDF generation with download animation
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { saveAs } from 'file-saver';
import * as pdfjsLib from 'pdfjs-dist';

// Configure the PDF.js worker to use a local file
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./pdf.worker.min.mjs', import.meta.url).toString();

/**
 * PageManager class to handle multi-page functionality
 */
class PageManager {
  constructor() {
    this.pages = [];
    this.currentPageIndex = 0;
    this.isDataApplied = false;
    this.fileHandle = null;
    this.startPageNumber = 1; // Add custom start page number property
    
    // Initialize navigation control flags
    this.isManuallyNavigating = false;
    this.tempDisableScrollEvents = false;
    
    // DOM elements
    this.pagesContainer = document.getElementById('pages-container');
    this.prevPageBtn = document.getElementById('prev-page-btn');
    this.nextPageBtn = document.getElementById('next-page-btn');
    this.pageIndicator = document.getElementById('page-indicator');
    this.addPageBtn = document.getElementById('add-page-btn');
    this.insertAfterBtn = document.getElementById('insert-after-btn');
    this.deletePageBtn = document.getElementById('delete-page-btn');
    this.saveStateBtn = document.getElementById('save-state-btn');
    
    // Initialize event listeners
    this.initEventListeners();
    
    // Load saved state if available
    this.loadState();
    
    // Replace Save PDF button with Print button
    this.replaceSavePDFWithPrint();
  }
  
  /**
   * Replace the Save PDF button with a Print button and add a Load Document button
   */
  replaceSavePDFWithPrint() {
    // Replace PDF button with Print button
    const savePdfBtn = document.getElementById('save-pdf-btn');
    if (savePdfBtn) {
      savePdfBtn.textContent = 'Print Document';
      savePdfBtn.innerHTML = `
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" width="1.2em" height="1.2em" style="margin-right: 10px; vertical-align: middle;">
          <path d="M19 8h-1V3H6v5H5c-1.66 0-3 1.34-3 3v6h4v4h12v-4h4v-6c0-1.66-1.34-3-3-3zM8 5h8v3H8V5zm8 14H8v-6h8v6zm2-8H6v-1c0-.55.45-1 1-1h10c.55 0 1 .45 1 1v1z"/>
        </svg>
        Print Document
      `;
      
      // Replace with Print functionality
      savePdfBtn.onclick = () => {
        // First, save the current page state
        this.saveCurrentPageState();
        
        // Show a notification to inform the user
        this.showSaveLoadNotification('Opening print dialog...', 'info');
        
        // Ensure hyperlinks are properly initialized before printing
        if (typeof attachAllHyperlinkHandlers === 'function') {
          attachAllHyperlinkHandlers();
        }
        
        // Prepare document for printing by ensuring all hyperlinks are properly styled
        const allLinks = document.querySelectorAll('.page-link');
        allLinks.forEach(link => {
          // Make sure all hyperlinks have href for PDF compatibility
          if (!link.href || !link.href.includes('#page=')) {
            const targetPage = link.getAttribute('data-target-page');
            if (targetPage) {
              link.href = '#page=' + targetPage;
            }
          }
        });
        
        // Use a small timeout to ensure the notification is visible before print dialog opens
        setTimeout(() => {
          // Use the browser's built-in print functionality
          window.print();
        }, 100);
      };
    }
    
    // Check if we need to add a Load Document button
    const saveStateBtn = document.getElementById('save-state-btn');
    if (saveStateBtn) {
      let loadStateBtn = document.getElementById('load-state-btn');
      
      // Create Load Document button if it doesn't exist
      if (!loadStateBtn) {
        loadStateBtn = document.createElement('button');
        loadStateBtn.id = 'load-state-btn';
        loadStateBtn.className = 'pagination-control load-btn';
        loadStateBtn.innerHTML = `
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" width="1.2em" height="1.2em" style="margin-right: 10px; vertical-align: middle;">
            <path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm-2 17c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H6V4h7v5h2v2z"/>
          </svg>
          Load Document
        `;
        
        // Add click handler to trigger file loading
        loadStateBtn.addEventListener('click', () => {
          this.loadStateFromUserAction();
        });
        
        // Insert the button next to the Save button
        if (saveStateBtn.parentNode) {
          saveStateBtn.parentNode.insertBefore(loadStateBtn, saveStateBtn.nextSibling);
        }
      }
    }
  }
  
  /**
   * Initialize event listeners for page navigation
   */
  initEventListeners() {
    // Add page button
    if (this.addPageBtn) {
      this.addPageBtn.addEventListener('click', () => this.addNewPage());
    }
    
    // Insert page after current button
    if (this.insertAfterBtn) {
      this.insertAfterBtn.addEventListener('click', () => this.insertPageAfterCurrent());
    }
    
    // Navigation buttons
    if (this.prevPageBtn) {
      this.prevPageBtn.addEventListener('click', () => this.navigateToPrevPage());
    }
    
    if (this.nextPageBtn) {
      this.nextPageBtn.addEventListener('click', () => this.navigateToNextPage());
    }
    
    // Delete page button
    if (this.deletePageBtn) {
      this.deletePageBtn.addEventListener('click', () => this.deleteCurrentPage());
    }
    
    // Save state button
    if (this.saveStateBtn) {
      this.saveStateBtn.addEventListener('click', () => this.manualSaveState());
    }
    
    // Page jump functionality
    const pageJumpForm = document.getElementById('page-jump-form');
    const pageInput = document.getElementById('page-input');
    const pageTotal = document.getElementById('page-total');

    if (pageJumpForm && pageInput) {
      // Handle form submission
      pageJumpForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const pageNumber = parseInt(pageInput.value);
        this.jumpToPage(pageNumber);
      });

      // Update when input loses focus
      pageInput.addEventListener('blur', () => {
        const pageNumber = parseInt(pageInput.value);
        this.jumpToPage(pageNumber);
      });
      
      // Update when Enter key is pressed
      pageInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          const pageNumber = parseInt(pageInput.value);
          this.jumpToPage(pageNumber);
        }
      });
    }
    
    // Start page number functionality
    const startPageNumberInput = document.getElementById('start-page-number');
    const applyStartPageBtn = document.getElementById('apply-start-page');
    
    if (startPageNumberInput && applyStartPageBtn) {
      applyStartPageBtn.addEventListener('click', () => {
        const startNumber = parseInt(startPageNumberInput.value);
        if (!isNaN(startNumber) && startNumber >= 1) {
          this.startPageNumber = startNumber;
          this.updateAllPageNumbers();
          this.showSaveLoadNotification(`Current and following pages now start from ${startNumber}`, 'info');
        } else {
          // If input is invalid or empty, reset to sequential numbering
          this.startPageNumber = 1;
          this.updateAllPageNumbers();
          this.showSaveLoadNotification('Current and following pages reset to sequential numbering', 'info');
        }
      });
      
      // Also apply when Enter key is pressed
      startPageNumberInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          applyStartPageBtn.click();
        }
      });
    }
    
    // Keyboard navigation
    document.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowLeft') {
        this.navigateToPrevPage();
      } else if (e.key === 'ArrowRight') {
        this.navigateToNextPage();
      } else if (e.key === 'Delete' && e.ctrlKey) {
        // Add Ctrl+Delete as keyboard shortcut for deleting pages
        this.deleteCurrentPage();
      }
    });
    
    // Initialize flags for scroll management
    this.isManuallyNavigating = false;
    this.tempDisableScrollEvents = false;
    
    // Handle horizontal scroll in pages container
    this.pagesContainer.addEventListener('wheel', (e) => {
      if (e.shiftKey || Math.abs(e.deltaX) > Math.abs(e.deltaY)) {
        e.preventDefault();
        this.pagesContainer.scrollLeft += e.deltaY;
        
        // Only update active page if not currently navigating manually
        if (!this.tempDisableScrollEvents && !this.isManuallyNavigating) {
          this.updateActivePageFromScroll();
        }
      }
    });
    
    // Handle scroll end to update active page, with a much longer debounce
    this.pagesContainer.addEventListener('scroll', this.debounce(() => {
      // Only update if not currently navigating manually
      if (!this.tempDisableScrollEvents && !this.isManuallyNavigating) {
        this.updateActivePageFromScroll();
      }
    }, 300)); // Increased debounce time to prevent frequent updates
    
    // Add reload warning and save prompt before unload (closing tab/refreshing)
    window.addEventListener('beforeunload', (e) => {
      // Create a flag to check if we've explicitly saved first
      const lastManualSave = localStorage.getItem('lastManualSaveTime') || '0';
      const timeSinceLastManualSave = Date.now() - parseInt(lastManualSave);
      
      // If it's been less than 2 minutes since manual save, don't show warning
      if (timeSinceLastManualSave < 2 * 60 * 1000) {
        return;
      }
      
      // Save state automatically
      this.saveState('auto');
      
      // Standard way to show confirmation dialog before closing/reloading
      e.preventDefault();
      e.returnValue = 'You have unsaved changes. Are you sure you want to leave? Your work might be lost if you haven\'t used the Save button.';
      return e.returnValue;
    });
    
    // Disabled automatic periodic saves as they were causing issues
    // Instead we'll save only on important events and manual saves
    
    // Add custom reload handler to ensure state is properly restored
    window.addEventListener('load', () => {
      console.log('Page loaded - ensuring proper state restoration');
      // Set a flag indicating this was a manual reload/navigation
      localStorage.setItem('wasManualNavigation', 'true');
    });
  }
  
  /**
   * Update active page based on scroll position
   */
  updateActivePageFromScroll() {
    // Skip automatic scroll detection when user is manually navigating
    if (this.isManuallyNavigating) {
      return;
    }
    
    const containerLeft = this.pagesContainer.scrollLeft;
    const containerWidth = this.pagesContainer.clientWidth;
    const centerPoint = containerLeft + (containerWidth / 2);
    
    let closestPageIndex = 0;
    let closestDistance = Number.MAX_SAFE_INTEGER;
    
    // Find the page closest to the center of the viewport
    const pageElements = this.pagesContainer.querySelectorAll('.legacy-sheet');
    pageElements.forEach((page, index) => {
      const rect = page.getBoundingClientRect();
      const pageCenter = rect.left + (rect.width / 2);
      const distance = Math.abs(pageCenter - centerPoint);
      
      if (distance < closestDistance) {
        closestDistance = distance;
        closestPageIndex = index;
      }
    });
    
    // Only update if the page is significantly in view (more than 75% visible)
    if (closestDistance < containerWidth * 0.25 && this.currentPageIndex !== closestPageIndex) {
      // Update the page indicator and buttons without scrolling
      this.currentPageIndex = closestPageIndex;
      this.updatePageIndicator();
      this.updateNavigationButtons();
      
      // Just update the active class without scrolling
      const pages = this.pagesContainer.querySelectorAll('.legacy-sheet');
      pages.forEach(page => page.classList.remove('active-page'));
      pages[closestPageIndex].classList.add('active-page');
    }
  }
  
  /**
   * Add a new page to the document
   */
  addNewPage() {
    // First save current page state
    this.saveCurrentPageState();
    
    // Create a new blank page instead of cloning the current page
    const newPage = document.createElement('div');
    newPage.className = 'legacy-sheet';
    
    // Add basic page structure with only one generation title
    newPage.innerHTML = `
      <div class="generation-title-container">
        <div class="generation-title" contenteditable="true">New Generation</div>
        <button class="generation-title-delete-btn">×</button>
        <hr class="horizontal-rule">
      </div>
      
      <!-- Footer -->
      <div class="footer-line"></div>
      <div class="footer">
        <div class="report-name">Sapling Platinum Report</div>
        <div class="page-number">1</div>
      </div>
    `;
    
    // Set page index and ensure it's not active
    const newIndex = this.pagesContainer.querySelectorAll('.legacy-sheet').length;
    newPage.dataset.pageIndex = newIndex;
    
    // Update page number in footer based on custom start page number
    const pageNumberElement = newPage.querySelector('.page-number');
    if (pageNumberElement) {
      // For a new page, calculate its number based on its position relative to currentPageIndex
      const pageNumber = (newIndex - this.currentPageIndex) + this.startPageNumber;
      pageNumberElement.textContent = pageNumber.toString();
    }
    
    // Add deletion functionality to the generation title
    const deleteButton = newPage.querySelector('.generation-title-delete-btn');
    const titleContainer = newPage.querySelector('.generation-title-container');
    const horizontalRule = newPage.querySelector('.horizontal-rule');
    
    if (deleteButton && titleContainer) {
      // Add delete functionality to the delete button
      deleteButton.addEventListener('click', function(e) {
        e.stopPropagation(); // Prevent triggering other click events
        
        // Add fade-out effect
        titleContainer.style.transition = 'opacity 0.3s ease, transform 0.3s ease, height 0.3s ease, margin 0.3s ease';
        titleContainer.style.opacity = '0';
        titleContainer.style.transform = 'translateY(-10px)';
        titleContainer.style.height = '0';
        titleContainer.style.marginTop = '0';
        titleContainer.style.marginBottom = '0';
        titleContainer.style.overflow = 'hidden';
        
        // Remove the container after the animation completes
        setTimeout(() => {
          titleContainer.remove();
          
          // Save state after deletion
          if (window.pageManager) {
            window.pageManager.saveCurrentPageState();
            window.pageManager.saveState();
          }
        }, 300);
      });
    }
    
    if (horizontalRule) {
      // Add click event to the horizontal rule for deletion
      horizontalRule.addEventListener('click', function(e) {
        if (deleteButton) deleteButton.click();
      });
    }
    
    const generationTitle = newPage.querySelector('.generation-title');
    if (generationTitle) {
      // Make it editable with proper styling
      generationTitle.style.outline = 'none';
      generationTitle.style.transition = 'background-color 0.2s ease';
      
      // Add hover effect
      generationTitle.addEventListener('mouseover', function() {
        this.style.backgroundColor = 'rgba(240, 240, 240, 0.5)';
        this.title = 'Click to edit generation title';
      });
      
      generationTitle.addEventListener('mouseout', function() {
        this.style.backgroundColor = 'transparent';
      });
      
      // Add focus/blur effects
      generationTitle.addEventListener('focus', function() {
        this.style.backgroundColor = 'rgba(240, 240, 240, 0.8)';
      });
      
      generationTitle.addEventListener('blur', function() {
        this.style.backgroundColor = 'transparent';
        
        // Save state when generation title is edited
        if (window.pageManager) {
          window.pageManager.saveCurrentPageState();
          window.pageManager.saveState();
        }
      });
      
      // Prevent Enter key from creating new lines
      generationTitle.addEventListener('keydown', function(e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          this.blur(); // Remove focus when Enter is pressed
        }
      });
    }
    
    // Add to container
    this.pagesContainer.appendChild(newPage);
    
    // Use setTimeout to ensure DOM is updated before setting the page active
    setTimeout(() => {
      // Ensure all existing pages are inactive
      const allPages = this.pagesContainer.querySelectorAll('.legacy-sheet');
      allPages.forEach(page => page.classList.remove('active-page'));
      
      // Directly add the active class to the new page
      newPage.classList.add('active-page');
      
      // Force scroll to the new page immediately
      newPage.scrollIntoView({
        behavior: 'auto', // Use immediate scrolling instead of smooth
        block: 'nearest',
        inline: 'center'
      });
      
      // Update current page index
      this.currentPageIndex = newIndex;
      
      // Update page indicator and navigation buttons
      this.updatePageIndicator();
      this.updateNavigationButtons();
      
      // Save state after all updates
      this.saveState();
    }, 50); // Short timeout to ensure DOM updates
  }
  
  /**
   * Reset page content but maintain structure
   * @param {HTMLElement} page - The page element to reset
   * @deprecated No longer used in addNewPage as we create a new blank page instead
   */
  resetPageContent(page) {
    // Clear person entries but keep header and footer
    const personEntries = page.querySelectorAll('.person-entry');
    const dividers = page.querySelectorAll('.entry-divider');
    
    // Remove person entries and dividers
    personEntries.forEach(entry => entry.remove());
    dividers.forEach(divider => divider.remove());
    
    // Keep only the first generation title container, remove the rest
    const generationTitleContainers = page.querySelectorAll('.generation-title-container');
    if (generationTitleContainers.length > 1) {
      for (let i = 1; i < generationTitleContainers.length; i++) {
        generationTitleContainers[i].remove();
      }
    }
    
    // Update the first generation title to default text
    const generationTitle = page.querySelector('.generation-title');
    if (generationTitle) {
      generationTitle.textContent = 'Generation';
    }
  }
  
  /**
   * Navigate to the previous page
   */
  navigateToPrevPage() {
    if (this.currentPageIndex > 0) {
      // Save current page state
      this.saveCurrentPageState();
      
      // Explicitly set manual navigation mode to true
      this.isManuallyNavigating = true;
      
      // Navigate to previous page
      this.setActivePage(this.currentPageIndex - 1);
    }
  }
  
  /**
   * Navigate to the next page
   */
  navigateToNextPage() {
    const pageCount = this.pagesContainer.querySelectorAll('.legacy-sheet').length;
    if (this.currentPageIndex < pageCount - 1) {
      // Save current page state
      this.saveCurrentPageState();
      
      // Explicitly set manual navigation mode to true
      this.isManuallyNavigating = true;
      
      // Navigate to next page
      this.setActivePage(this.currentPageIndex + 1);
    }
  }
  
  /**
   * Set the active page and update UI
   * @param {number} pageIndex - The index of the page to activate
   * @param {boolean} preventScroll - Whether to prevent scrolling to the page
   */
  setActivePage(pageIndex, preventScroll = false) {
    // Get all pages
    const pages = this.pagesContainer.querySelectorAll('.legacy-sheet');
    if (pageIndex < 0 || pageIndex >= pages.length) return;
    
    // Set manual navigation flag to prevent scroll detection from changing the page
    this.isManuallyNavigating = true;
    
    // Store the target page
    const targetPage = pages[pageIndex];
    
    // Remove active class from all pages
    pages.forEach(page => page.classList.remove('active-page'));
    
    // Add active class to target page
    targetPage.classList.add('active-page');
    
    if (!preventScroll) {
      // Disable scroll event handling temporarily
      this.tempDisableScrollEvents = true;
      
      // Scroll to the page immediately without animation
      targetPage.scrollIntoView({
        behavior: 'auto', // Use immediate scrolling for reliable positioning
        block: 'nearest',
        inline: 'center'
      });
      
      // Re-enable scroll events after scrolling is complete
      setTimeout(() => {
        this.tempDisableScrollEvents = false;
        this.isManuallyNavigating = false;
      }, 500); // Allow enough time for any scroll animations to complete
    }
    
    // Update current page index
    this.currentPageIndex = pageIndex;
    
    // Update page indicator
    this.updatePageIndicator();
    
    // Update navigation buttons
    this.updateNavigationButtons();
    
    // Make sure the page is visible by forcing focus to an element on the page
    // This helps ensure it's properly displayed
    const generationTitle = targetPage.querySelector('.generation-title');
    if (generationTitle) {
      // Set focus then immediately remove to avoid keyboard input
      generationTitle.setAttribute('tabindex', '-1');
      generationTitle.focus();
      generationTitle.blur();
      generationTitle.removeAttribute('tabindex');
    }
  }
  
  /**
   * Update the page indicator text
   */
  updatePageIndicator() {
    const pageCount = this.pagesContainer.querySelectorAll('.legacy-sheet').length;
    const pageInput = document.getElementById('page-input');
    const pageTotal = document.getElementById('page-total');
    
    if (pageInput && pageTotal) {
      // Update input value to current page (1-based)
      pageInput.value = this.currentPageIndex + 1;
      
      // Update max attribute to match total pages
      pageInput.max = pageCount;
      
      // Update total pages display
      pageTotal.textContent = pageCount;
    } else if (this.pageIndicator) {
      // Fallback to old text-only display if input not found
      this.pageIndicator.textContent = `Page ${this.currentPageIndex + 1} of ${pageCount}`;
    }
  }
  
  /**
   * Update page numbers based on the start page number for current and subsequent pages only
   */
  updateAllPageNumbers() {
    const pages = this.pagesContainer.querySelectorAll('.legacy-sheet');
    
    pages.forEach((page, index) => {
      // Only update current and subsequent pages
      if (index >= this.currentPageIndex) {
        const pageNumberElement = page.querySelector('.page-number');
        if (pageNumberElement) {
          // Calculate page number based on index and start page number
          const pageNumber = index - this.currentPageIndex + this.startPageNumber;
          pageNumberElement.textContent = pageNumber.toString();
        }
      }
    });
    
    // Save state to preserve new page numbering
    this.saveState();
  }
  
  /**
   * Update navigation button states
   */
  updateNavigationButtons() {
    const pageCount = this.pagesContainer.querySelectorAll('.legacy-sheet').length;
    
    // Update prev button
    if (this.prevPageBtn) {
      this.prevPageBtn.disabled = this.currentPageIndex === 0;
    }
    
    // Update next button
    if (this.nextPageBtn) {
      this.nextPageBtn.disabled = this.currentPageIndex === pageCount - 1;
    }
  }
  
  /**
   * Save the current page state
   */
  saveCurrentPageState() {
    const currentPage = this.pagesContainer.querySelector(`.legacy-sheet.active-page`);
    if (!currentPage) return;
    
    // Create a serialized version of the page
    const pageState = this.serializePage(currentPage);
    
    // Store in the pages array
    this.pages[this.currentPageIndex] = pageState;
  }
  
  /**
   * Serialize page content for storage
   * @param {HTMLElement} page - The page element to serialize
   * @returns {Object} - Serialized page data
   */
  serializePage(page) {
    return {
      pageIndex: parseInt(page.dataset.pageIndex),
      html: page.innerHTML,
      generationTitle: page.querySelector('.generation-title')?.textContent || ''
    };
  }
  
  /**
   * Save all page states with visual feedback - this is the explicit user-triggered save
   */
  async manualSaveState() {
    try {
    // Show saving indicator
    if (this.saveStateBtn) {
      this.saveStateBtn.classList.add('button-saving');
      this.saveStateBtn.disabled = true;
    }
    
    // Show initial toast notification
    this.showSaveLoadNotification('Saving document...');
    
      // Save current page first
      this.saveCurrentPageState();
      
      // Create state object
      const state = {
        pages: this.pages,
        currentPageIndex: this.currentPageIndex,
        isDataApplied: this.isDataApplied,
        startPageNumber: this.startPageNumber,
        savedAt: new Date().toISOString(),
        saveType: 'manual'
      };
      
      // Convert to JSON
      const fileContent = JSON.stringify(state);
      
      // Check if we have the File System Access API
      if ('showSaveFilePicker' in window) {
        // Either use existing file handle or get a new one
        if (!this.fileHandle) {
          // Configure save dialog options
          const options = {
            suggestedName: 'legacy-sheet-document.gene',
            types: [{
              description: 'Legacy Sheet Document',
              accept: {
                'application/json': ['.gene'],
              },
            }],
          };
          
          try {
            // Show the file picker
            this.fileHandle = await window.showSaveFilePicker(options);
          } catch (err) {
            // User cancelled the save dialog
            if (err.name !== 'AbortError') {
              console.error('Error showing file picker:', err);
              this.showSaveLoadNotification('Error saving file: ' + err.message, 'error');
            }
            
            // Reset UI
      if (this.saveStateBtn) {
        this.saveStateBtn.classList.remove('button-saving');
        this.saveStateBtn.disabled = false;
            }
            
            return;
          }
        }
        
        // Create a FileSystemWritableFileStream to write to
        const writable = await this.fileHandle.createWritable();
        
        // Write the contents of the file to the stream
        await writable.write(fileContent);
        
        // Close the file and write the contents to disk
        await writable.close();
        
        // Show success notification
      const formattedDate = new Date().toLocaleString(undefined, {
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
      
        this.showSaveLoadNotification(`Document saved to file at ${formattedDate}`, 'permanent-save');
        console.log('State saved successfully to file at ' + new Date().toISOString());
      } else {
        // Fallback for browsers without File System Access API
        const blob = new Blob([fileContent], { type: 'application/json' });
        saveAs(blob, 'legacy-sheet-document.gene');
        
        // Show success notification
        const formattedDate = new Date().toLocaleString(undefined, {
          month: 'short',
          day: 'numeric',
          hour: '2-digit',
          minute: '2-digit'
        });
        
        this.showSaveLoadNotification(`Document downloaded at ${formattedDate}`, 'permanent-save');
      }
      
      // Reset UI
      if (this.saveStateBtn) {
        this.saveStateBtn.classList.remove('button-saving');
        this.saveStateBtn.classList.add('button-success');
        this.saveStateBtn.disabled = false;
        
        setTimeout(() => {
          this.saveStateBtn.classList.remove('button-success');
        }, 1000);
      }
    } catch (e) {
      console.error('Error saving state to file:', e);
      this.showSaveLoadNotification('Error saving document: ' + e.message, 'error');
      
      // Reset UI
      if (this.saveStateBtn) {
        this.saveStateBtn.classList.remove('button-saving');
        this.saveStateBtn.disabled = false;
      }
    }
  }
  
  /**
   * Save state method - now just calls manualSaveState when needed
   * @param {string} saveType - The type of save operation
   */
  saveState(saveType = 'auto') {
    // For auto saves, we do nothing now - only manual saves are supported
        if (saveType === 'manual') {
      this.manualSaveState();
    } else {
      // Just save the current page state in memory, no persistence
      this.saveCurrentPageState();
      console.log('Auto-save skipped - document only saved on manual save now');
    }
    
    // Save the startPageNumber to local storage so it persists
    try {
      const startPageNumberSetting = { startPageNumber: this.startPageNumber };
      localStorage.setItem('legacySheetPageNumbering', JSON.stringify(startPageNumberSetting));
    } catch (e) {
      console.error('Error saving page numbering settings:', e);
    }
  }
  
  /**
   * Initial state loading - only sets up empty template, doesn't actually load files
   */
  loadState() {
    // Skip loading state if we just restored the template
    if (window.preventAutoDataApplyAfterReset) {
      console.log('Skipping state load after template reset');
      return;
    }
    
    // Just initialize with empty template
    console.log('Starting with empty template - use Load Document button to load saved files');
    
    // Load custom page numbering settings if they exist
    try {
      const pageNumberingSettings = localStorage.getItem('legacySheetPageNumbering');
      if (pageNumberingSettings) {
        const settings = JSON.parse(pageNumberingSettings);
        if (settings.startPageNumber) {
          this.startPageNumber = settings.startPageNumber;
          
          // Update the start page number input field if it exists
          const startPageNumberInput = document.getElementById('start-page-number');
          if (startPageNumberInput) {
            startPageNumberInput.value = this.startPageNumber;
          }
          
          // Update all page numbers to use the custom start number
          this.updateAllPageNumbers();
        }
      }
    } catch (e) {
      console.error('Error loading page numbering settings:', e);
    }
  }
  
  /**
   * Load state from a local file selected by the user - triggered by the Load button
   */
  async loadStateFromUserAction() {
    try {
      // Show initial loading notification
      this.showSaveLoadNotification('Please select a document file to open...', 'info');
      
      // Check if File System Access API is available
      if ('showOpenFilePicker' in window) {
        // Configure open dialog options
        const options = {
          types: [{
            description: 'Legacy Sheet Document',
            accept: {
              'application/json': ['.gene'],
            },
          }],
          multiple: false
        };
        
        // Show open file picker (now triggered by user interaction)
        const [fileHandle] = await window.showOpenFilePicker(options);
        this.fileHandle = fileHandle;
        
        // Get the file
        const file = await fileHandle.getFile();
        
        // Read the file
        const text = await file.text();
        
        // Parse the state
        const state = JSON.parse(text);
        
        // Apply the state
          if (this.applyLoadedState(state, true)) {
          console.log('Successfully loaded state from file');
          this.showSaveLoadNotification('Document loaded successfully', 'success');
        }
      } else {
        // Fallback for browsers without File System Access API
        // Create a file input element to allow file selection
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.gene,application/json';
        
        // Set up event handler for file selection
        fileInput.addEventListener('change', (event) => {
          if (event.target.files && event.target.files[0]) {
            this.loadStateFromFile(event.target.files[0]);
          }
        });
        
        // Trigger file dialog (this is ok in direct response to user interaction)
        fileInput.click();
      }
    } catch (err) {
      // User might have cancelled the dialog
      if (err.name !== 'AbortError') {
        console.error('Error loading file:', err);
        this.showSaveLoadNotification('Error loading file: ' + err.message, 'error');
      }
      
      console.log('No file selected or error occurred');
    }
  }
  
  /**
   * Load state from a file input element
   * @param {File} file - The file to load
   */
    async loadStateFromFile(file) {
    try {
      // Show loading notification
      this.showSaveLoadNotification('Loading document...', 'info');
      
      // Read the file
      const text = await file.text();
      
      // Parse the state
      const state = JSON.parse(text);
      
      // Apply the state
      if (this.applyLoadedState(state, true)) {
        this.showSaveLoadNotification('Document loaded successfully', 'success');
        console.log('Successfully loaded state from file');
      }
    } catch (err) {
      console.error('Error loading file:', err);
      this.showSaveLoadNotification('Error loading document: ' + err.message, 'error');
    }
  }
  
  /**
   * Load page states from the server API - DEPRECATED, now using file system
   * Kept for reference only, not used anymore
   */
  /*async loadStateFromServer() {
    try {
      // Skip if we're showing static template after reset
      if (window.preventAutoDataApplyAfterReset) {
        console.log('Skipping server state load after template reset');
        return false;
      }
      
      // Show loading indicator
      this.showSaveLoadNotification('Loading document...');
      
      // Fetch state from the server
      const response = await fetch('/api/load-state');
      
      // Check if the request was successful
      if (!response.ok) {
        // If server returns 404 (no state found), fall back to localStorage
        if (response.status === 404) {
          console.log('No saved state found on server');
          return this.loadStateFromLocalStorage();
        }
        throw new Error(`Server returned ${response.status}: ${response.statusText}`);
      }
      
      // Parse the state data
      const state = await response.json();
      
      // Before applying server state, check if we have more recent permanent local state
      try {
        const permanentSavedState = localStorage.getItem('legacySheetState_permanent');
        if (permanentSavedState) {
          const localState = JSON.parse(permanentSavedState);
          
          // Compare timestamps
          const serverTime = new Date(state.savedAt).getTime();
          const localTime = new Date(localState.savedAt).getTime();
          
          // If local permanent save is newer, use it instead
          if (localTime > serverTime) {
            console.log('Found newer permanent local save, using that instead of server state');
            
            this.showSaveLoadNotification('Restoring from your latest saved document', 'info');
            
            // Apply the local state
            this.applyLoadedState(localState, true);
            return true;
          }
        }
      } catch (e) {
        console.error('Error comparing saved states:', e);
      }
      
      // Apply the server state
      this.applyLoadedState(state);
      
      // Log success
      console.log('State loaded successfully from server');
      
      // Return true to indicate success
      return true;
    } catch (error) {
      console.error('Error loading state from server:', error);
      // Re-throw the error so the caller can fall back to localStorage
      throw error;
    }
  }*/
  
  /**
   * Load page states from localStorage as a fallback - DEPRECATED, now using file system
   * Kept for reference only, not used anymore
   */
  /*loadStateFromLocalStorage() {
    try {
      // Skip if we're showing static template after reset
      if (window.preventAutoDataApplyAfterReset) {
        console.log('Skipping localStorage state load after template reset');
        return false;
      }
      
      // First check if we have a permanent saved state
      const permanentSavedState = localStorage.getItem('legacySheetState_permanent');
      if (permanentSavedState) {
        try {
          console.log('Found permanent saved state, trying this first');
          const permanentState = JSON.parse(permanentSavedState);
          
          // Apply the permanent state
          const success = this.applyLoadedState(permanentState, true);
          if (success) {
            console.log('Successfully loaded permanent saved state');
            this.showSaveLoadNotification('Restored from your permanently saved document', 'success');
            return true;
          }
        } catch (e) {
          console.error('Error loading permanent saved state, falling back to regular state:', e);
        }
      }
      
      // Check for backup saved states
      const backupState = localStorage.getItem('legacySheetBackup');
      if (backupState) {
        try {
          console.log('Found backup saved state, trying this second');
          const parsedBackupState = JSON.parse(backupState);
          
          // Apply the backup state
          const success = this.applyLoadedState(parsedBackupState, true);
          if (success) {
            console.log('Successfully loaded from backup saved state');
            this.showSaveLoadNotification('Restored from your backup saved document', 'success');
            return true;
          }
        } catch (e) {
          console.error('Error loading backup saved state, falling back to regular state:', e);
        }
      }
      
      // Finally, try regular state
      const savedState = localStorage.getItem('legacySheetState');
      if (savedState) {
        const state = JSON.parse(savedState);
        
        // Apply the state
        return this.applyLoadedState(state);
      } else {
        console.log('No saved state found in localStorage');
        
        // Show a warning if we got here after a manual navigation
        const wasManualNavigation = localStorage.getItem('wasManualNavigation') === 'true';
        if (wasManualNavigation) {
          this.showSaveLoadNotification('No saved document found! Starting with a new template.', 'warning');
          localStorage.removeItem('wasManualNavigation');
        }
        
        return false;
      }
    } catch (e) {
      console.error('Error loading state from localStorage:', e);
      return false;
    }
  }*/
  
  /**
   * Apply a loaded state object to the application
   * @param {Object} state - The state object to apply
   * @param {boolean} isPermanent - Whether this is a permanent/manual save being restored
   * @returns {boolean} - Whether the state was applied successfully
   */
  applyLoadedState(state, isPermanent = false) {
    try {
      // Check if this is a permanent state or has the permanent flag
      const isPermanentState = isPermanent || state.isPermanent || state.saveType === 'manual';
      
      // Restore pages if we have saved pages
      if (state.pages && state.pages.length > 0) {
        // Clear existing pages except the first one
        const pages = this.pagesContainer.querySelectorAll('.legacy-sheet');
        for (let i = 1; i < pages.length; i++) {
          pages[i].remove();
        }
        
        // Restore first page
        const firstPage = this.pagesContainer.querySelector('.legacy-sheet');
        if (firstPage && state.pages[0]) {
          firstPage.innerHTML = state.pages[0].html;
          firstPage.dataset.pageIndex = state.pages[0].pageIndex;
        }
        
        // Add additional pages
        for (let i = 1; i < state.pages.length; i++) {
          const pageData = state.pages[i];
          const newPage = firstPage.cloneNode(false); // clone without children
          newPage.innerHTML = pageData.html;
          newPage.dataset.pageIndex = pageData.pageIndex;
          newPage.classList.remove('active-page');
          this.pagesContainer.appendChild(newPage);
        }
        
        // Set the active page
        this.currentPageIndex = state.currentPageIndex || 0;
        this.setActivePage(this.currentPageIndex);
        
        // Set data applied flag
        this.isDataApplied = state.isDataApplied || false;
        
        // Reinitialize hyperlink handlers for loaded content
        if (typeof attachAllHyperlinkHandlers === 'function') {
          setTimeout(attachAllHyperlinkHandlers, 200);
        }
        
        // Restore startPageNumber if it exists in the state
        if (state.startPageNumber && !isNaN(state.startPageNumber)) {
          this.startPageNumber = state.startPageNumber;
          
          // Update the start page number input field if it exists
          const startPageNumberInput = document.getElementById('start-page-number');
          if (startPageNumberInput) {
            startPageNumberInput.value = this.startPageNumber;
          }
          
          // Apply page numbering based on start number
          this.updateAllPageNumbers();
        }
        
        // Update UI elements based on loaded state
        this.updateNavigationButtons();
        this.updatePageIndicator();
        this.updateAddPageButtonVisibility();
        
        // Check for missing images after restoration and add re-upload buttons
        this.handleMissingImagesAfterRestore();
        
        // No need to store in localStorage anymore as we're using File System Access API
        
        // Display save timestamp if available
        if (state.savedAt) {
          const saveDate = new Date(state.savedAt);
          const formattedDate = saveDate.toLocaleDateString(undefined, { 
            year: 'numeric', 
            month: 'short', 
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
          });
          
          if (isPermanentState) {
            console.log('Permanent state restored from save point: ' + formattedDate);
            // Show a temporary toast notification about the loaded permanent state
            this.showSaveLoadNotification(`Restored your permanently saved document from ${formattedDate}`, 'success');
          } else {
            console.log('State restored from save point: ' + formattedDate);
            // Show a temporary toast notification about the loaded state
            this.showSaveLoadNotification(`Loaded document from ${formattedDate}`);
          }
        } else {
          console.log('State restored successfully');
          this.showSaveLoadNotification('Document loaded successfully');
        }
        
        return true;
      }
      return false;
    } catch (error) {
      console.error('Error applying loaded state:', error);
      this.showSaveLoadNotification('Error loading document', 'error');
      return false;
    }
  }
  
  /**
   * Create a debounced function
   * @param {Function} func - The function to debounce
   * @param {number} wait - The debounce wait time in ms
   * @returns {Function} - The debounced function
   */
  debounce(func, wait) {
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
  
  /**
   * Show a temporary notification toast about save/load operations
   * @param {string} message - The message to display
   * @param {string} type - The type of notification ('success', 'warning', 'error', 'info', 'permanent-save')
   */
  /**
   * Handle missing images after state restoration
   * Checks for image containers with broken images and adds re-upload functionality
   */
  handleMissingImagesAfterRestore() {
    // Find all pages
    const pages = this.pagesContainer.querySelectorAll('.legacy-sheet');
    
    pages.forEach(page => {
      // Find all image containers
      const imageContainers = page.querySelectorAll('.person-image-container');
      
      imageContainers.forEach(container => {
        const img = container.querySelector('.person-image');
        const personContent = container.closest('.person-content');
        
        if (img) {
          // Check if image is broken or empty src
          const checkImage = () => {
            if (img.naturalWidth === 0 || img.src === '' || img.src === '#' || (img.src.startsWith('blob:') && !img.complete)) {
              console.log('Found missing image, replacing with re-upload button');
              this.replaceWithImageUploadPlaceholder(container, personContent);
            }
          };
          
          // Check immediately and also after load attempt fails
          checkImage();
          img.addEventListener('error', () => {
            checkImage();
          });
        } else {
          // No image element found in container, replace with upload placeholder
          this.replaceWithImageUploadPlaceholder(container, personContent);
        }
      });
    });
  }
  
  /**
   * Replace missing image with upload placeholder
   * @param {HTMLElement} container - The image container element
   * @param {HTMLElement} personContent - The person content element
   */
  replaceWithImageUploadPlaceholder(container, personContent) {
    // Preserve position and styling of the original container
    const containerStyle = window.getComputedStyle(container);
    const marginTop = containerStyle.marginTop;
    const marginLeft = containerStyle.marginLeft;
    const float = containerStyle.float;
    
    // Create a new placeholder with the same positioning
    const personImagePlaceholder = document.createElement('button'); // Changed from div to button for better accessibility
    personImagePlaceholder.className = 'person-image-placeholder';
    personImagePlaceholder.style.marginTop = marginTop;
    personImagePlaceholder.style.marginLeft = marginLeft;
    personImagePlaceholder.style.float = float;
    personImagePlaceholder.style.background = 'none';
    personImagePlaceholder.style.border = 'none';
    personImagePlaceholder.style.cursor = 'pointer';
    personImagePlaceholder.style.padding = '0';
    personImagePlaceholder.style.width = '180px';
    personImagePlaceholder.style.height = '220px';
    personImagePlaceholder.style.outline = 'none';
    personImagePlaceholder.innerHTML = `
      <svg width="180" height="220" xmlns="http://www.w3.org/2000/svg">
        <rect width="180" height="220" fill="#F0F0F0" stroke="#CCC" stroke-width="2" stroke-dasharray="5,5"/>
        <text x="90" y="100" font-family="Arial" font-size="16" text-anchor="middle" fill="#666">Click to re-upload</text>
        <text x="90" y="120" font-family="Arial" font-size="16" text-anchor="middle" fill="#666">person image</text>
        <text x="90" y="140" font-family="Arial" font-size="14" text-anchor="middle" fill="#666">📷</text>
      </svg>
    `;
    
        // Add click event to open file picker
    personImagePlaceholder.onclick = function(e) {
      e.preventDefault(); // Prevent any default button behavior
      e.stopPropagation(); // Stop event bubbling
      
      // Create a temporary file input element
      const fileInput = document.createElement('input');
      fileInput.type = 'file';
      fileInput.accept = 'image/*';
      
      // Listen for file selection
      fileInput.onchange = function() {
        if (fileInput.files && fileInput.files[0]) {
          const selectedFile = fileInput.files[0];
          
          // Create image element to replace placeholder
          const personImage = document.createElement('img');
          personImage.className = 'person-image';
          personImage.alt = 'Person image';
          
          // Create an object URL for the selected file
          const objectURL = URL.createObjectURL(selectedFile);
          personImage.src = objectURL;
          
          // Create a container for the image and delete button
          const imageContainer = document.createElement('div');
          imageContainer.className = 'person-image-container';
          
          // Apply the same positioning as the original container
          imageContainer.style.marginTop = marginTop;
          imageContainer.style.marginLeft = marginLeft;
          imageContainer.style.float = float;
          
          // Create delete button
          const deleteButton = document.createElement('button');
          deleteButton.className = 'person-image-delete-btn';
          deleteButton.innerHTML = '×';
          deleteButton.title = 'Remove image';
          
          // Add click event to delete button
          deleteButton.onclick = function(e) {
            e.stopPropagation(); // Prevent triggering image click
            
            // Revoke object URL to prevent memory leaks
            URL.revokeObjectURL(personImage.src);
            
            // Remove the container
            imageContainer.remove();
            
            // Save state after removing the image
            if (window.pageManager) {
              window.pageManager.saveCurrentPageState();
              window.pageManager.saveState();
            }
          };
          
          // Add image and delete button to container
          imageContainer.appendChild(personImage);
          imageContainer.appendChild(deleteButton);
          
          // Replace placeholder with container
          personImagePlaceholder.parentNode.replaceChild(imageContainer, personImagePlaceholder);
          
          // Make the image container vertically draggable
          makeImageDraggable(imageContainer, personContent);
          
          // Add click event to the image to allow replacing it
          personImage.onclick = function(e) {
            e.stopPropagation(); // Prevent event bubbling
            
            const newFileInput = document.createElement('input');
            newFileInput.type = 'file';
            newFileInput.accept = 'image/*';
            
            newFileInput.onchange = function() {
              if (newFileInput.files && newFileInput.files[0]) {
                // Revoke the old object URL to prevent memory leaks
                URL.revokeObjectURL(personImage.src);
                
                // Create a new object URL for the new file
                const newObjectURL = URL.createObjectURL(newFileInput.files[0]);
                personImage.src = newObjectURL;
                
                // Reflow the main paragraph text to accommodate the image
                const mainParagraph = personContent.querySelector('.main-person-paragraph');
                if (mainParagraph && typeof reflowImageIfNeeded === 'function') {
                  reflowImageIfNeeded(mainParagraph);
                }
                
                // Save state after changing the image
                if (window.pageManager) {
                  window.pageManager.saveCurrentPageState();
                  window.pageManager.saveState();
                }
              }
            };
            
            // Use a timeout to ensure the file input dialog appears properly
            setTimeout(() => {
              newFileInput.click();
            }, 100);
          };
          
          // Reflow the main paragraph text to accommodate the image
          const mainParagraph = personContent.querySelector('.main-person-paragraph');
          if (mainParagraph && typeof reflowImageIfNeeded === 'function') {
            reflowImageIfNeeded(mainParagraph);
          }
          
          // Save state after adding the image
          if (window.pageManager) {
            window.pageManager.saveCurrentPageState();
            window.pageManager.saveState();
          }
        }
      };
      
      // Trigger file selection dialog with a slight delay to ensure it always appears
      setTimeout(() => {
        fileInput.click();
      }, 100);
    };
    
    // Replace the original container with our placeholder
    container.parentNode.replaceChild(personImagePlaceholder, container);
  }
  
  showSaveLoadNotification(message, type = 'info') {
    // Check if notification element already exists
    let notification = document.querySelector('.save-load-notification');
    
    // Create notification if it doesn't exist
    if (!notification) {
      notification = document.createElement('div');
      notification.className = 'save-load-notification';
      
      // Add icon
      const icon = document.createElement('div');
      icon.className = 'save-load-notification-icon';
      
      // Default icon is checkmark
      let iconSvg = `
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" width="1.2em" height="1.2em">
          <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
        </svg>
      `;
      
      // Add message element
      const messageEl = document.createElement('span');
      messageEl.className = 'save-load-notification-message';
      
      // Assemble notification
      notification.appendChild(icon);
      notification.appendChild(messageEl);
      
      // Add to DOM
      document.body.appendChild(notification);
    }
    
    // Update message
    const messageEl = notification.querySelector('.save-load-notification-message');
    if (messageEl) {
      messageEl.textContent = message;
    }
    
    // Update icon based on type
    const iconEl = notification.querySelector('.save-load-notification-icon');
    if (iconEl) {
      let iconSvg = '';
      
      switch (type) {
        case 'success':
          iconSvg = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#28a745" stroke-width="1.5" width="1.2em" height="1.2em">
              <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
            </svg>
          `;
          break;
        case 'warning':
          iconSvg = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#ffc107" stroke-width="1.5" width="1.2em" height="1.2em">
              <path d="M12 2L1 21h22L12 2z"/>
              <path d="M12 16v-5"/>
              <path d="M12 19v.01"/>
            </svg>
          `;
          break;
        case 'error':
          iconSvg = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#dc3545" stroke-width="1.5" width="1.2em" height="1.2em">
              <circle cx="12" cy="12" r="10"/>
              <path d="M15 9l-6 6M9 9l6 6"/>
            </svg>
          `;
          break;
        case 'permanent-save':
          iconSvg = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#28a745" stroke-width="1.5" width="1.2em" height="1.2em">
              <path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/>
            </svg>
          `;
          break;
        default: // info
          iconSvg = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#17a2b8" stroke-width="1.5" width="1.2em" height="1.2em">
              <circle cx="12" cy="12" r="10"/>
              <path d="M12 8v4M12 16h.01"/>
            </svg>
          `;
      }
      
      iconEl.innerHTML = iconSvg;
    }
    
    // Update notification class based on type
    notification.className = 'save-load-notification';
    notification.classList.add(`notification-${type}`);
    
    // For permanent saves, make the notification more prominent
    if (type === 'permanent-save') {
      notification.classList.add('permanent-save-notification');
    }
    
    // Show notification
    notification.classList.add('visible');
    
    // Determine display duration based on type
    const duration = type === 'permanent-save' ? 5000 : 3000;
    
    // Hide after a delay
    setTimeout(() => {
      notification.classList.remove('visible');
    }, duration);
  }
  
  /**
   * Check if data has been applied to the template
   * @returns {boolean} - True if data has been applied
   */
  hasDataBeenApplied() {
    // Check if static template was cleared
    const currentPage = this.pagesContainer.querySelector('.legacy-sheet.active-page');
    if (!currentPage) return false;
    
    // Check if there are any person entries
    const personEntries = currentPage.querySelectorAll('.person-entry');
    return personEntries.length > 0;
  }
  
  /**
   * Set data applied flag
   * @param {boolean} value - The new value
   */
  setDataApplied(value) {
    this.isDataApplied = value;
    this.updateAddPageButtonVisibility();
  }
  
  /**
   * Update the visibility of the Add Page button
   */
  updateAddPageButtonVisibility() {
    if (!this.addPageBtn) return;
    
    // Only show if data has been applied to the current page
    if (this.hasDataBeenApplied()) {
      this.addPageBtn.style.display = 'flex';
    } else {
      this.addPageBtn.style.display = 'none';
    }
  }

  /**
   * Insert a new page after the current active page
   */
  insertPageAfterCurrent() {
    // First save current page state
    this.saveCurrentPageState();
    
    // Create a new blank page similar to addNewPage method
    const newPage = document.createElement('div');
    newPage.className = 'legacy-sheet';
    
    // Add basic page structure with only one generation title
    newPage.innerHTML = `
      <div class="generation-title-container">
        <div class="generation-title" contenteditable="true">New Generation</div>
        <button class="generation-title-delete-btn">×</button>
        <hr class="horizontal-rule">
      </div>
      
      <!-- Footer -->
      <div class="footer-line"></div>
      <div class="footer">
        <div class="report-name">Sapling Platinum Report</div>
        <div class="page-number">1</div>
      </div>
    `;
    
    // Calculate the insertion index (current page index + 1)
    const insertIndex = this.currentPageIndex + 1;
    
    // Get all pages and the reference page (the page after which we insert)
    const pages = this.pagesContainer.querySelectorAll('.legacy-sheet');
    const refPage = pages[this.currentPageIndex];
    
    // Insert the new page after the current page in the DOM
    if (insertIndex < pages.length) {
      this.pagesContainer.insertBefore(newPage, pages[insertIndex]);
    } else {
      // If we're inserting after the last page, just append
      this.pagesContainer.appendChild(newPage);
    }
    
    // Set page index and ensure it's not active
    newPage.dataset.pageIndex = insertIndex;
    
    // Update page number in footer based on custom start page number
    const pageNumberElement = newPage.querySelector('.page-number');
    if (pageNumberElement) {
      // Calculate page number for the new page - it's 1 more than current page's number
      const currentPageElement = pages[this.currentPageIndex].querySelector('.page-number');
      const currentPageNumber = parseInt(currentPageElement?.textContent || this.startPageNumber);
      const pageNumber = currentPageNumber + 1;
      pageNumberElement.textContent = pageNumber.toString();
    }
    
    // Add deletion functionality to the generation title
    const deleteButton = newPage.querySelector('.generation-title-delete-btn');
    const titleContainer = newPage.querySelector('.generation-title-container');
    const horizontalRule = newPage.querySelector('.horizontal-rule');
    
    if (deleteButton && titleContainer) {
      // Add delete functionality to the delete button
      deleteButton.addEventListener('click', function(e) {
        e.stopPropagation(); // Prevent triggering other click events
        
        // Add fade-out effect
        titleContainer.style.transition = 'opacity 0.3s ease, transform 0.3s ease, height 0.3s ease, margin 0.3s ease';
        titleContainer.style.opacity = '0';
        titleContainer.style.transform = 'translateY(-10px)';
        titleContainer.style.height = '0';
        titleContainer.style.marginTop = '0';
        titleContainer.style.marginBottom = '0';
        titleContainer.style.overflow = 'hidden';
        
        // Remove the container after the animation completes
        setTimeout(() => {
          titleContainer.remove();
          
          // Save state after deletion
          if (window.pageManager) {
            window.pageManager.saveCurrentPageState();
            window.pageManager.saveState();
          }
        }, 300);
      });
    }
    
    if (horizontalRule) {
      // Add click event to the horizontal rule for deletion
      horizontalRule.addEventListener('click', function(e) {
        if (deleteButton) deleteButton.click();
      });
    }
    
    const generationTitle = newPage.querySelector('.generation-title');
    if (generationTitle) {
      // Make it editable with proper styling
      generationTitle.style.outline = 'none';
      generationTitle.style.transition = 'background-color 0.2s ease';
      
      // Add hover effect
      generationTitle.addEventListener('mouseover', function() {
        this.style.backgroundColor = 'rgba(240, 240, 240, 0.5)';
        this.title = 'Click to edit generation title';
      });
      
      generationTitle.addEventListener('mouseout', function() {
        this.style.backgroundColor = 'transparent';
      });
      
      // Add focus/blur effects
      generationTitle.addEventListener('focus', function() {
        this.style.backgroundColor = 'rgba(240, 240, 240, 0.8)';
      });
      
      generationTitle.addEventListener('blur', function() {
        this.style.backgroundColor = 'transparent';
        
        // Save state when generation title is edited
        if (window.pageManager) {
          window.pageManager.saveCurrentPageState();
          window.pageManager.saveState();
        }
      });
      
      // Prevent Enter key from creating new lines
      generationTitle.addEventListener('keydown', function(e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          this.blur(); // Remove focus when Enter is pressed
        }
      });
    }
    
    // Update page indices for all pages after the inserted page
    for (let i = insertIndex + 1; i < pages.length + 1; i++) {
      const page = pages[i - 1]; // -1 because we've shifted all indices
      if (page) {
        // Update page index data attribute
        page.dataset.pageIndex = i;
        
        // Update page number in footer - we need to increment all subsequent page numbers
        const pageNum = page.querySelector('.page-number');
        if (pageNum) {
          // Get the current number from the new page that was just inserted
          const newPageNumber = parseInt(newPage.querySelector('.page-number')?.textContent || this.startPageNumber);
          
          // For all subsequent pages, add 1 to their current value
          const currentNumber = parseInt(pageNum.textContent);
          // Make sure we have a valid number
          if (!isNaN(currentNumber)) {
            pageNum.textContent = (currentNumber + 1).toString();
          }
        }
      }
    }
    
    // Update the array to include our new page at the right position
    this.pages.splice(insertIndex, 0, this.serializePage(newPage));
    
    // Set the active page to our new page
    setTimeout(() => {
      // Update current page index
      this.currentPageIndex = insertIndex;
      
      // Navigate to the new page
      this.setActivePage(insertIndex);
      
      // Update page indicator and navigation buttons
      this.updatePageIndicator();
      this.updateNavigationButtons();
      
      // Save state after all updates
      this.saveState();
      
      // Show a notification
      this.showSaveLoadNotification('New page inserted after current page', 'success');
    }, 50); // Short timeout to ensure DOM updates
  }

  /**
   * Delete the current active page
   */
  deleteCurrentPage() {
    // Get all pages
    const pages = this.pagesContainer.querySelectorAll('.legacy-sheet');
    
    // Don't delete if there's only one page left
    if (pages.length <= 1) {
      alert('Cannot delete the only page. At least one page must remain.');
      return;
    }
    
    // Ask for confirmation
    if (!confirm('Are you sure you want to delete this page? This action cannot be undone.')) {
      return;
    }
    
    // Get the current active page
    const currentPage = this.pagesContainer.querySelector('.legacy-sheet.active-page');
    if (!currentPage) return;
    
    // Determine the new active page index (previous page or first page if deleting first)
    let newActiveIndex = this.currentPageIndex > 0 ? this.currentPageIndex - 1 : 1;
    
    // Remove the page from the DOM
    currentPage.remove();
    
    // Remove the page data from our pages array
    if (this.pages[this.currentPageIndex]) {
      this.pages.splice(this.currentPageIndex, 1);
    }
    
    // Update page indices for all remaining pages
    pages.forEach((page, index) => {
      if (index !== this.currentPageIndex) { // Skip the one we just removed
        const pageIndex = parseInt(page.dataset.pageIndex);
        if (pageIndex > this.currentPageIndex) {
          // Decrement indices for pages that came after the deleted one
          page.dataset.pageIndex = (pageIndex - 1).toString();
          
          // Update the page number in the footer based on custom start page number
          const pageNumberElement = page.querySelector('.page-number');
          if (pageNumberElement) {
            // Only update if page is after the current page index
            if (pageIndex > this.currentPageIndex) {
              const newPageNumber = (pageIndex - 1 - this.currentPageIndex) + this.startPageNumber;
              pageNumberElement.textContent = newPageNumber.toString();
            }
          }
        }
      }
    });
    
    // Navigate to the new active page
    this.setActivePage(newActiveIndex);
    
    // Update navigation controls
    this.updateNavigationButtons();
    this.updatePageIndicator();
    
    // Save state
    this.saveState();
    
    console.log(`Deleted page at index ${this.currentPageIndex}`);
  }

  /**
   * Jump to a specific page by page number (1-based index)
   * @param {number} pageNumber - The page number to jump to (1-based)
   */
  jumpToPage(pageNumber) {
    const pageCount = this.pagesContainer.querySelectorAll('.legacy-sheet').length;
    
    // Validate page number
    if (isNaN(pageNumber) || pageNumber < 1 || pageNumber > pageCount) {
      // Reset to current page if invalid
      this.updatePageIndicator();
      return;
    }
    
    // Convert from 1-based page number to 0-based index
    const pageIndex = pageNumber - 1;
    
    // Save current page state before jumping
    this.saveCurrentPageState();
    
    // Explicitly set manual navigation mode to true
    this.isManuallyNavigating = true;
    
    // Set the active page
    this.setActivePage(pageIndex);
  }
}

document.addEventListener('DOMContentLoaded', function() {
  console.log('Legacy Sheet Template loaded successfully');
  
  // Check if we need to show the static template (coming from a reset)
  const showStaticTemplate = localStorage.getItem('showStaticTemplate');
  if (showStaticTemplate === 'true') {
    // Clear the flag so it doesn't persist on future page loads
    localStorage.removeItem('showStaticTemplate');
    
    // Force clearing of any previously saved state
    localStorage.removeItem('legacySheetState');
    
    // Set staticTemplateCleared to false to ensure we see the original template
    window.staticTemplateCleared = false;
    
    // Set this flag to skip auto-application of data
    window.preventAutoDataApplyAfterReset = true;
    
    // Show a success message
    const successMessage = document.createElement('div');
    successMessage.textContent = 'Original template restored';
    successMessage.style.position = 'fixed';
    successMessage.style.bottom = '80px';
    successMessage.style.left = '50%';
    successMessage.style.transform = 'translateX(-50%)';
    successMessage.style.padding = '10px 20px';
    successMessage.style.backgroundColor = 'rgba(40, 167, 69, 0.8)';
    successMessage.style.color = 'white';
    successMessage.style.borderRadius = '5px';
    successMessage.style.fontFamily = "'Public Sans', sans-serif";
    successMessage.style.fontSize = '14px';
    successMessage.style.zIndex = '9999';
    
    document.body.appendChild(successMessage);
    
    // Remove the message after 3 seconds
    setTimeout(() => {
      successMessage.style.opacity = '0';
      successMessage.style.transition = 'opacity 0.5s ease';
      setTimeout(() => {
        successMessage.remove();
      }, 500);
    }, 3000);
    
    console.log('Restored to static template');
  }
  
  // Initialize the page manager
  const pageManager = new PageManager();
  window.pageManager = pageManager; // Make it globally accessible
  
  // Add the "Next person has an image" checkbox to the upload area
  const uploadContainer = document.querySelector('.upload-container');
  
  // We skip dynamically creating the checkbox since it already exists in the HTML
  // Just get a reference to the existing checkbox
  const nextPersonHasImageCheckbox = document.getElementById('next-person-has-image');
  
  // Update the structure of existing generation titles to add delete functionality
  const updateExistingGenerationTitles = function() {
    // Find all generation titles that don't have proper container structure
    const allPages = document.querySelectorAll('.legacy-sheet');
    
    allPages.forEach(page => {
      // Find direct generation titles that aren't in a generation-title-container
      const generationTitle = page.querySelector('.generation-title:not(.generation-title-container .generation-title)');
      const horizontalRule = generationTitle ? generationTitle.nextElementSibling : null;
      
      if (generationTitle && horizontalRule && horizontalRule.classList.contains('horizontal-rule')) {
        // Create a new container
        const titleContainer = document.createElement('div');
        titleContainer.className = 'generation-title-container';
        
        // Create delete button
        const deleteButton = document.createElement('button');
        deleteButton.className = 'generation-title-delete-btn';
        deleteButton.innerHTML = '×';
        deleteButton.title = 'Delete generation title';
        
        // Add delete functionality
        deleteButton.addEventListener('click', function(e) {
          e.stopPropagation();
          
          // Add fade-out effect
          titleContainer.style.transition = 'opacity 0.3s ease, transform 0.3s ease, height 0.3s ease, margin 0.3s ease';
          titleContainer.style.opacity = '0';
          titleContainer.style.transform = 'translateY(-10px)';
          titleContainer.style.height = '0';
          titleContainer.style.marginTop = '0';
          titleContainer.style.marginBottom = '0';
          titleContainer.style.overflow = 'hidden';
          
          // Remove after animation
          setTimeout(() => {
            titleContainer.remove();
            
            // Save state
            if (window.pageManager) {
              window.pageManager.saveCurrentPageState();
              window.pageManager.saveState();
            }
          }, 300);
        });
        
        // Add click event to the horizontal rule
        horizontalRule.addEventListener('click', function(e) {
          deleteButton.click();
        });
        
        // Replace the original structure
        const parentNode = generationTitle.parentNode;
        parentNode.insertBefore(titleContainer, generationTitle);
        
        // Move existing elements into container
        titleContainer.appendChild(generationTitle);
        titleContainer.appendChild(deleteButton);
        titleContainer.appendChild(horizontalRule);
      }
    });
  };
  
  // Call the function to update existing titles
  updateExistingGenerationTitles();
  
  // Make generation title editable when document loads
  const generationTitles = document.querySelectorAll('.generation-title');
  generationTitles.forEach(generationTitle => {
    if (generationTitle) {
      // Make it editable
      generationTitle.setAttribute('contenteditable', 'true');
      
      // Add styling for better UX when editing
      generationTitle.style.outline = 'none';
      generationTitle.style.transition = 'background-color 0.2s ease';
      
      // Add hover effect
      generationTitle.addEventListener('mouseover', function() {
        this.style.backgroundColor = 'rgba(240, 240, 240, 0.5)';
        this.title = 'Click to edit generation title';
      });
      
      generationTitle.addEventListener('mouseout', function() {
        this.style.backgroundColor = 'transparent';
      });
      
      // Add focus/blur effects
      generationTitle.addEventListener('focus', function() {
        this.style.backgroundColor = 'rgba(240, 240, 240, 0.8)';
      });
      
      generationTitle.addEventListener('blur', function() {
        this.style.backgroundColor = 'transparent';
        
        // Save state when generation title is edited
        if (window.pageManager) {
          window.pageManager.saveCurrentPageState();
          window.pageManager.saveState();
        }
      });
      
      // Prevent Enter key from creating new lines
      generationTitle.addEventListener('keydown', function(e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          this.blur(); // Remove focus when Enter is pressed
        }
      });
    }
  });
  
  // Initialize insertion controls for adding generation titles
  addInsertionControls();
  setupInsertionControlsObserver();
  
  // Create Add Generation Title button in the UI
  createAddGenerationTitleButton();
  
  // Create Add Child Number button in the UI
  createAddChildNumberButton();
  
  // Add event listener for the reset template button
  const resetTemplateBtn = document.getElementById('reset-template-btn');
  if (resetTemplateBtn) {
    resetTemplateBtn.addEventListener('click', function() {
      // Show confirmation dialog
      if (confirm('Are you sure you want to reset all pages to the initial template state? This will remove all your data and cannot be undone.')) {
        resetToInitialTemplate();
      }
    });
  }
  
  // Function to reset all pages to initial template state
  function resetToInitialTemplate() {
    try {
      // Show a loading message
      const loadingMessage = document.createElement('div');
      loadingMessage.textContent = 'Resetting to original template...';
      loadingMessage.style.position = 'fixed';
      loadingMessage.style.top = '50%';
      loadingMessage.style.left = '50%';
      loadingMessage.style.transform = 'translate(-50%, -50%)';
      loadingMessage.style.padding = '15px 25px';
      loadingMessage.style.backgroundColor = 'rgba(0, 0, 0, 0.7)';
      loadingMessage.style.color = 'white';
      loadingMessage.style.borderRadius = '5px';
      loadingMessage.style.fontFamily = "'Public Sans', sans-serif";
      loadingMessage.style.fontSize = '16px';
      loadingMessage.style.zIndex = '10000';
      
      document.body.appendChild(loadingMessage);
      
      // Create and store a special flag to ensure the original template is shown
      localStorage.setItem('showStaticTemplate', 'true');
      
      // Force a complete reset by clearing all state from localStorage
      localStorage.removeItem('legacySheetState');
      
      // Also clear the server-side state by posting an empty state
      fetch('/api/save-state', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ reset: true, resetAt: new Date().toISOString() })
      })
      .then(response => response.json())
      .then(data => {
        console.log('Server state reset:', data);
      })
      .catch(error => {
        console.error('Failed to reset server state:', error);
      });
      
      // Reset the staticTemplateCleared flag
      staticTemplateCleared = false;
      
      // Reset global variables that might be causing issues
      window.parsedGenealogyData = null;
      
      // Reload the page to get a fresh start
      setTimeout(() => {
        window.location.reload();
      }, 500);
    } catch (error) {
      console.error('Error resetting template:', error);
      alert('There was an error resetting the template. Please try again.');
    }
  }
  
  // Set up a mutation observer to make any new generation titles editable
  const pagesContainer = document.getElementById('pages-container');
  if (pagesContainer) {
    const observer = new MutationObserver(mutations => {
      mutations.forEach(mutation => {
        if (mutation.addedNodes.length) {
          mutation.addedNodes.forEach(node => {
            if (node.classList && node.classList.contains('legacy-sheet')) {
              // Call the function to update generation titles for new pages
              updateExistingGenerationTitles();
              
              // Process generation title that might not have been caught by updateExistingGenerationTitles
              const newTitle = node.querySelector('.generation-title');
              if (newTitle) {
                // Make it editable
                newTitle.setAttribute('contenteditable', 'true');
                
                // Add styling for better UX when editing
                newTitle.style.outline = 'none';
                newTitle.style.transition = 'background-color 0.2s ease';
                
                // Add hover effect
                newTitle.addEventListener('mouseover', function() {
                  this.style.backgroundColor = 'rgba(240, 240, 240, 0.5)';
                  this.title = 'Click to edit generation title';
                });
                
                newTitle.addEventListener('mouseout', function() {
                  this.style.backgroundColor = 'transparent';
                });
                
                // Add focus/blur effects
                newTitle.addEventListener('focus', function() {
                  this.style.backgroundColor = 'rgba(240, 240, 240, 0.8)';
                });
                
                newTitle.addEventListener('blur', function() {
                  this.style.backgroundColor = 'transparent';
                  
                  // Save state when generation title is edited
                  if (window.pageManager) {
                    window.pageManager.saveCurrentPageState();
                    window.pageManager.saveState();
                  }
                });
                
                // Prevent Enter key from creating new lines
                newTitle.addEventListener('keydown', function(e) {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    this.blur(); // Remove focus when Enter is pressed
                  }
                });
              }
              
              // Process insertion controls for new elements
              const personEntries = node.querySelectorAll('.person-entry');
              personEntries.forEach(entry => {
                addInsertionControl(entry);
              });
              
              const entryDividers = node.querySelectorAll('.entry-divider');
              entryDividers.forEach(divider => {
                addInsertionControl(divider);
              });
            }
          });
        }
      });
    });
    
    observer.observe(pagesContainer, { childList: true, subtree: true });
  }
  
  // Get the Save PDF button
  const savePdfBtn = document.getElementById('save-pdf-btn');
  
  // Add click event listener to the button
  savePdfBtn.addEventListener('click', function() {
    // Update button state
    savePdfBtn.textContent = 'Generating PDF...';
    savePdfBtn.disabled = true;
    
    // Hide the button temporarily
    savePdfBtn.style.display = 'none';
    
    // Save the current page state before generating PDF
    window.pageManager.saveCurrentPageState();
    
    // Get all pages
    const pages = document.querySelectorAll('.legacy-sheet');
    if (!pages.length) {
      console.error('No pages found');
      return;
    }
    
    // Initialize PDF with letter size in portrait orientation
    const pdf = new jsPDF({
      unit: 'in',
      format: 'letter',
      orientation: 'portrait'
    });
    
    // Track current page for processing
    let currentPageIndex = 0;
    
    // Function to process each page sequentially
    function processNextPage() {
      if (currentPageIndex >= pages.length) {
        // All pages processed, save the PDF
        const pdfBlob = pdf.output('blob');
        
        // Use FileSaver to trigger browser download animation
        saveAs(pdfBlob, 'Legacy-Sheet.pdf');
        
        // Show button again after a short delay
        setTimeout(() => {
          savePdfBtn.style.display = 'block';
          savePdfBtn.textContent = 'Save PDF';
          savePdfBtn.disabled = false;
        }, 1500);
        
        return;
      }
      
      const page = pages[currentPageIndex];
      
      // Store current opacity
      const originalOpacity = page.style.opacity;
      const originalTransform = page.style.transform;
      
      // Ensure page is fully visible for capture
      page.style.opacity = '1';
      page.style.transform = 'none';
      
      // Use html2canvas to capture the page
      html2canvas(page, {
        scale: 2,  // Higher scale for better quality
        useCORS: true,
        scrollY: -window.scrollY,
        windowWidth: document.documentElement.offsetWidth,
        windowHeight: document.documentElement.offsetHeight
      }).then(canvas => {
        // Convert canvas to image
        const imgData = canvas.toDataURL('image/png');
        
        // Add a new page to the PDF for pages after the first one
        if (currentPageIndex > 0) {
          pdf.addPage();
        }
        
        // Calculate dimensions to fit the page correctly
        const imgProps = pdf.getImageProperties(imgData);
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = pdf.internal.pageSize.getHeight();
        
        // Add image to PDF, sizing to fit the page
        pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
        
        // Restore original styles
        page.style.opacity = originalOpacity;
        page.style.transform = originalTransform;
        
        // Process next page
        currentPageIndex++;
        processNextPage();
      });
    }
    
    // Start processing pages
    processNextPage();
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
  // We're using the global nextPersonHasImageCheckbox reference instead

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
    
    // Clear extracted data and reset "Has Image" checkbox
    dataPreview.innerHTML = '<p class="no-data-message">No data extracted yet</p>';
    
    // Reset the "Next person has an image" checkbox
    if (nextPersonHasImageCheckbox) {
      nextPersonHasImageCheckbox.checked = false;
    }
    
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
            
            // Check if the afterNameText starts with a comma or if we need to add one
            if (afterNameText) {
              if (afterNameText.startsWith(',')) {
                // If afterNameText already has a comma, keep it as is
                currentParagraphPreview = `<b>${preserveHyperlinks(extractedName)}</b>${preserveHyperlinks(afterNameText)}`;
              } else {
                // If no comma, add one with a space after it
                currentParagraphPreview = `<b>${preserveHyperlinks(extractedName)}</b>, ${preserveHyperlinks(afterNameText)}`;
              }
            } else {
              // If no text after name, just output the name
              currentParagraphPreview = `<b>${preserveHyperlinks(extractedName)}</b>`;
            }
            
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

    // If we detect we're showing the static template after a reset, don't auto-apply
    if (localStorage.getItem('showStaticTemplate') === 'true' || window.preventAutoDataApplyAfterReset) {
      console.log('Static template mode detected, not auto-applying data');
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
    
    // Get the active page from the page manager
    const legacySheet = document.querySelector('.legacy-sheet.active-page');
    if (!legacySheet) {
      console.error('No active page found');
      return;
    }

    // Update the generation title only if found
    if (data.generation) {
      const generationTitle = legacySheet.querySelector('.generation-title');
      if (generationTitle) {
        generationTitle.textContent = data.generation;
      }
    }

    // Clear existing person entries ONLY if this is the first time we're applying data
    // This removes the static template entries but keeps our dynamically added entries
    if (!staticTemplateCleared && !window.staticTemplateCleared) {
      const existingEntries = legacySheet.querySelectorAll('.person-entry');
      const horizontalRule = legacySheet.querySelector('.horizontal-rule');

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
      window.staticTemplateCleared = true;
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
      
      // Make person number deletable with hover effect
      personNumber.style.cursor = 'pointer';
      personNumber.title = 'Click to remove this number';
      
      // Add hover effects
      personNumber.addEventListener('mouseover', function() {
        this.style.opacity = '0.6';
        this.style.boxShadow = '0 0 3px rgba(0,0,0,0.2)';
      });
      
      personNumber.addEventListener('mouseout', function() {
        this.style.opacity = '1';
        this.style.boxShadow = 'none';
      });
      
      // Add click event to delete the number
      personNumber.addEventListener('click', function() {
        // Add transition for smooth fade out
        this.style.transition = 'opacity 0.3s ease, width 0.3s ease, margin 0.3s ease';
        this.style.opacity = '0';
        this.style.width = '0';
        this.style.margin = '0';
        this.style.border = 'none';
        this.style.padding = '0';
        
        // Remove the element after transition
        setTimeout(() => {
          this.remove();
          
          // Adjust the entry padding/margin to maintain alignment
          personEntry.style.paddingLeft = '0';
          personContent.style.marginLeft = '0';
          
          // Add a subtle transition for content reflow
          personContent.style.transition = 'margin-left 0.3s ease';
          personContent.style.marginLeft = '0';
        }, 300);
      });

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
            
            // Check if the afterName starts with a comma or if we need to add one
            if (afterName) {
              if (afterName.startsWith(',')) {
                // If afterName already has a comma, keep it as is
                mainParagraphHTML = `<b>${preserveHyperlinks(name)}</b>${preserveHyperlinks(afterName)}`;
              } else {
                // If no comma, add one with a space after it
                mainParagraphHTML = `<b>${preserveHyperlinks(name)}</b>, ${preserveHyperlinks(afterName)}`;
              }
            } else {
              // If no text after name, just output the name
              mainParagraphHTML = `<b>${preserveHyperlinks(name)}</b>`;
            }
          } else {
            // Fallback if regex doesn't match (e.g. no clear name pattern on the first line)
            mainParagraphHTML = preserveHyperlinks(firstContentLineText);
          }
          lineIndex++;

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

          // Create main paragraph div with hover controls
          const mainParagraphDiv = document.createElement('div');
          mainParagraphDiv.className = 'main-person-paragraph';
          mainParagraphDiv.innerHTML = mainParagraphHTML;
          
          // Make it editable
          mainParagraphDiv.setAttribute('contenteditable', 'true');
          
          // Add event listener for 'Enter' key to handle user breaks
          mainParagraphDiv.addEventListener('keydown', function(e) {
            if (e.key === 'Enter') {
              e.preventDefault();
              document.execCommand('insertLineBreak');
              // Update the raw data with the breaks
              updatePersonRawWithUserBreaks(person, personContent);
            }
          });
          
          // Add event listeners to check if selection is bold
          mainParagraphDiv.addEventListener('mouseup', checkSelectionFormatting);
          mainParagraphDiv.addEventListener('keyup', checkSelectionFormatting);
          mainParagraphDiv.addEventListener('click', checkSelectionFormatting);
          mainParagraphDiv.addEventListener('focus', checkSelectionFormatting);
          
                      // Add hover controls for editing/deleting
            addParagraphControls(mainParagraphDiv, person, personContent);
           
            // Make main paragraph resizable
            makeResizable(mainParagraphDiv);
          
          personContent.appendChild(mainParagraphDiv);
          
          // Check if the next person has an image using our HTML checkbox
          if (nextPersonHasImageCheckbox && nextPersonHasImageCheckbox.checked) {
            // Create image placeholder
            const personImagePlaceholder = document.createElement('div');
            personImagePlaceholder.className = 'person-image-placeholder';
            personImagePlaceholder.innerHTML = `
              <svg width="180" height="220" xmlns="http://www.w3.org/2000/svg">
                <rect width="180" height="220" fill="#F0F0F0" stroke="#CCC" stroke-width="2" stroke-dasharray="5,5"/>
                <text x="90" y="100" font-family="Arial" font-size="16" text-anchor="middle" fill="#666">Click to add</text>
                <text x="90" y="120" font-family="Arial" font-size="16" text-anchor="middle" fill="#666">person image</text>
                <text x="90" y="140" font-family="Arial" font-size="14" text-anchor="middle" fill="#666">📷</text>
              </svg>
            `;
            
            // Add click event to open file picker
            personImagePlaceholder.addEventListener('click', function() {
              // Create a temporary file input element
              const fileInput = document.createElement('input');
              fileInput.type = 'file';
              fileInput.accept = 'image/*';
              
              // Listen for file selection
              fileInput.addEventListener('change', function() {
                if (fileInput.files && fileInput.files[0]) {
                  const selectedFile = fileInput.files[0];
                  
                  // Create image element to replace placeholder
                  const personImage = document.createElement('img');
                  personImage.className = 'person-image';
                  personImage.alt = person.name || 'Person image';
                  
                  // Create an object URL for the selected file
                  const objectURL = URL.createObjectURL(selectedFile);
                  personImage.src = objectURL;
                  
                  // Create a container for the image and delete button
                  const imageContainer = document.createElement('div');
                  imageContainer.className = 'person-image-container';
                  
                  // Create delete button
                  const deleteButton = document.createElement('button');
                  deleteButton.className = 'person-image-delete-btn';
                  deleteButton.innerHTML = '×';
                  deleteButton.title = 'Remove image';
                  
                  // Add click event to delete button
                  deleteButton.addEventListener('click', function(e) {
                    e.stopPropagation(); // Prevent triggering image click
                    
                    // Revoke object URL to prevent memory leaks
                    URL.revokeObjectURL(personImage.src);
                    
                    // Remove the container
                    imageContainer.remove();
                    
                    // Save state after removing the image
                    if (window.pageManager) {
                      window.pageManager.saveCurrentPageState();
                      window.pageManager.saveState();
                    }
                  });
                  
                  // Add image and delete button to container
                  imageContainer.appendChild(personImage);
                  imageContainer.appendChild(deleteButton);
                  
                  // Replace placeholder with container
                  personImagePlaceholder.parentNode.replaceChild(imageContainer, personImagePlaceholder);
                  
                  // Make the image container vertically draggable
                  makeImageDraggable(imageContainer, personContent);
                  
                  // Add click event to the image to allow replacing it
                  personImage.addEventListener('click', function() {
                    const newFileInput = document.createElement('input');
                    newFileInput.type = 'file';
                    newFileInput.accept = 'image/*';
                    
                    newFileInput.addEventListener('change', function() {
                      if (newFileInput.files && newFileInput.files[0]) {
                        // Revoke the old object URL to prevent memory leaks
                        URL.revokeObjectURL(personImage.src);
                        
                        // Create a new object URL for the new file
                        const newObjectURL = URL.createObjectURL(newFileInput.files[0]);
                        personImage.src = newObjectURL;
                        
                        // Reflow the main paragraph text to accommodate the image
                        const mainParagraph = personContent.querySelector('.main-person-paragraph');
                        if (mainParagraph && typeof reflowImageIfNeeded === 'function') {
                          reflowImageIfNeeded(mainParagraph);
                        }
                        
                        // Save state after changing the image
                        if (window.pageManager) {
                          window.pageManager.saveCurrentPageState();
                          window.pageManager.saveState();
                        }
                      }
                    });
                    
                    newFileInput.click();
                  });
                  
                  // Reflow the main paragraph text to accommodate the image
                  const mainParagraph = personContent.querySelector('.main-person-paragraph');
                  if (mainParagraph && typeof reflowImageIfNeeded === 'function') {
                    reflowImageIfNeeded(mainParagraph);
                  }
                  
                  // Save state after adding the image
                  if (window.pageManager) {
                    window.pageManager.saveCurrentPageState();
                    window.pageManager.saveState();
                  }
                }
              });
              
              // Trigger file selection dialog
              fileInput.click();
            });
            
            // Add the placeholder to person content before the subparagraphs
            personContent.insertBefore(personImagePlaceholder, mainParagraphDiv.nextSibling);
          }
          
          // Also check formatting on mousedown to handle selection by double/triple clicking
          mainParagraphDiv.addEventListener('mousedown', function(e) {
            // Slight delay to allow browser to complete the selection process
            setTimeout(checkSelectionFormatting, 10);
          });
          
          // Add listener for selection changes within the document
          document.addEventListener('selectionchange', function(e) {
            // Only process if the paragraph has focus
            if (document.activeElement === mainParagraphDiv) {
              checkSelectionFormatting();
            }
          });
          
          // Function to check if current selection is within bold text
          function checkSelectionFormatting() {
            const controls = mainParagraphDiv.querySelector('.paragraph-controls');
            if (!controls) return;
            
            const boldBtn = controls.querySelector('.bold-btn');
            if (!boldBtn) return;
            
            try {
              // Only check if this paragraph is active
              if (document.activeElement !== mainParagraphDiv) {
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
      legacySheet.insertBefore(personEntry, legacySheet.querySelector('.footer'));

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
      
      legacySheet.insertBefore(divider, legacySheet.querySelector('.footer'));
    }

    // No need to update the Apply button status since it's hidden now
    
    // Show Next Person button - adjust its styling now that Apply button is hidden
    nextPersonBtn.style.display = 'inline-block';
    nextPersonBtn.style.marginTop = '20px'; // Ensure proper spacing above the button
    
    // Update extraction status
    extractionStatus.textContent = 'Entry applied! Ready for next person.';

    // Update the PageManager's dataApplied flag and show "Add New Page" button if needed
    if (window.pageManager) {
      window.pageManager.setDataApplied(true);
      window.pageManager.updateAddPageButtonVisibility();
      
      // Save state after applying data
      window.pageManager.saveState();
    }
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
    
    /* Enhanced notification styles */
    .save-load-notification {
      position: fixed;
      bottom: 20px;
      left: 50%;
      transform: translateX(-50%);
      background-color: white;
      color: #333;
      padding: 10px 20px;
      border-radius: 5px;
      display: flex;
      align-items: center;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.2);
      opacity: 0;
      visibility: hidden;
      transition: opacity 0.3s ease, visibility 0.3s ease, transform 0.3s ease;
      z-index: 9999;
      font-family: 'Public Sans', sans-serif;
    }
    
    .save-load-notification.visible {
      opacity: 1;
      visibility: visible;
    }
    
    /* Notification types */
    .save-load-notification.notification-success {
      border-left: 4px solid #28a745;
    }
    
    .save-load-notification.notification-warning {
      border-left: 4px solid #ffc107;
    }
    
    .save-load-notification.notification-error {
      border-left: 4px solid #dc3545;
    }
    
    .save-load-notification.notification-info {
      border-left: 4px solid #17a2b8;
    }
    
    .save-load-notification.notification-permanent-save {
      border-left: 4px solid #28a745;
      background-color: #f0fff4;
      box-shadow: 0 3px 15px rgba(40, 167, 69, 0.3);
      transform: translateX(-50%) scale(1.05);
    }
    
    /* Animation for permanent save notification */
    .save-load-notification.permanent-save-notification {
      animation: pulse-border 2s ease-in-out;
    }
    
    @keyframes pulse-border {
      0% { border-left-width: 4px; }
      50% { border-left-width: 10px; }
      100% { border-left-width: 4px; }
    }
    
    /* Add styles for person-number hover states */
    .person-number {
      transition: opacity 0.2s ease, box-shadow 0.2s ease;
      position: relative;
    }
    
    .person-number:before {
      content: '';
      position: absolute;
      top: -5px;
      left: -5px;
      right: -5px;
      bottom: -5px;
      border-radius: 50%;
      pointer-events: none;
      transition: background-color 0.2s ease;
    }
    
    .person-number:hover:before {
      background-color: rgba(0,0,0,0.03);
    }
    
    /* Make sure horizontal rule stays fixed regardless of generation title size */
    .generation-title {
      display: block;
      width: auto;
      text-align: center;
      min-height: 40px;
    }
    
    /* Theme color for specially formatted text */
    .theme-color {
      color: #a85733 !important;
    }
    
    /* Hyperlink styled text */
    .hyperlink-style {
      color: #a85733;
      font-weight: bold;
      text-decoration: underline;
    }
    
    /* Styling for active formatting buttons */
    .paragraph-controls button.active-format {
      background-color: rgba(0, 0, 0, 0.1);
      color: #a85733;
    }
    
    /* Style the color button with theme color */
    .paragraph-controls .color-btn svg {
      stroke: #a85733;
    }
    
    /* Style the hyperlink button with theme color */
    .paragraph-controls .hyperlink-btn svg {
      stroke: #a85733;
    }
    
    /* Styles for main person paragraph */
    .main-person-paragraph {
      position: relative;
      transition: background-color 0.2s ease;
      padding: 4px 6px;
      border-radius: 3px;
      margin-bottom: 1em;
    }
    
    .main-person-paragraph:hover {
      background-color: rgba(240, 240, 240, 0.5);
    }
    
    /* Show paragraph controls on hover for main paragraph too */
    .main-person-paragraph:hover .paragraph-controls {
      display: flex;
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
    
    /* Person image placeholder styling */
    .person-image-placeholder {
      float: right;
      width: 180px;
      height: 220px;
      margin-left: 0.145in;
      margin-bottom: 10px;
      margin-top: -32px;
      cursor: pointer;
      border-radius: 10%;
      transition: background-color 0.2s ease, transform 0.2s ease;
    }
    
    .person-image-placeholder:hover {
      background-color: #f5f5f5;
      transform: scale(1.02);
    }
    
    /* Container for person image and delete button */
    .person-image-container {
      float: right;
      position: relative;
      margin-left: 0.145in;
      margin-bottom: 10px;
      margin-top: -32px;
      border-radius: 10%;
      cursor: move;
      z-index: 10;
    }
    
    .person-image-container.dragging {
      opacity: 0.8;
      pointer-events: none;
      z-index: 1000;
      box-shadow: 0 5px 15px rgba(0, 0, 0, 0.3);
    }
    
    /* Drag handle for vertical movement */
    .image-drag-handle {
      position: absolute;
      left: 50%;
      transform: translateX(-50%);
      width: 30px;
      height: 12px;
      background-color: rgba(0, 120, 215, 0.7);
      border-radius: 6px;
      cursor: ns-resize;
      opacity: 0;
      transition: opacity 0.2s ease;
      z-index: 100;
    }
    
    .image-drag-handle::after {
      content: '';
      position: absolute;
      top: 4px;
      left: 7px;
      width: 16px;
      height: 4px;
      background-color: white;
      border-radius: 2px;
    }
    
    .image-drag-handle.top {
      top: -6px;
    }
    
    .image-drag-handle.bottom {
      bottom: -6px;
    }
    
    .person-image-container:hover .image-drag-handle {
      opacity: 0.8;
    }
    
    .image-drag-handle:hover {
      opacity: 1 !important;
    }
    
    /* Delete button for person image */
    .person-image-delete-btn {
      position: absolute;
      top: -10px;
      right: -10px;
      width: 24px;
      height: 24px;
      border-radius: 50%;
      background-color: rgba(220, 53, 69, 0.9);
      color: white;
      border: none;
      font-size: 18px;
      line-height: 1;
      padding: 0;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      opacity: 0;
      transform: scale(0.8);
      transition: opacity 0.2s ease, transform 0.2s ease;
    }
    
    .person-image-container:hover .person-image-delete-btn {
      opacity: 1;
      transform: scale(1);
    }
    
    .person-image-delete-btn:hover {
      background-color: rgba(220, 53, 69, 1);
    }
    
    /* Style for person image */
    .person-image {
      max-width: 180px;
      max-height: 220px;
      width: auto;
      height: auto;
      border-radius: 10%;
      object-fit: contain;
      cursor: pointer;
    }
    
    .person-image:hover {
      opacity: 0.9;
      transform: scale(1.02);
      transition: opacity 0.2s ease, transform 0.2s ease;
    }
    
    /* Add style for child number circles - improved to handle large numbers */
    .child-number {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 20px; /* Base size that will grow */
      height: 20px;
      border-radius: 10px; /* Half of height to ensure it stays circular */
      background-color: transparent;
      border: 1px solid #a85733;
      font-size: 11px; /* Slightly smaller font for better fit */
      margin-right: 4px;
      text-align: center;
      line-height: 1;
      font-weight: normal;
      color: #333;
      box-sizing: content-box; /* Ensures border doesn't affect dimensions */
      padding: 0 4px; /* Horizontal padding to accommodate larger numbers */
    }
    
    /* For extremely large child numbers */
    .child-number.large-number {
      padding: 0 6px; /* More horizontal padding for very large numbers */
      font-size: 9px; /* Smaller font for very large numbers */
      border-radius: 12px; /* Adjusted for the shape */
    }
    
    /* Paragraph controls - Apple UI Style */
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
    
    /* "Has image" checkbox styling */
    .has-image-checkbox {
      accent-color: #a85733;
      width: 16px;
      height: 16px;
    }
    
    /* Add Child Number button styling */
    .add-child-number-button {
      background-color: #fff;
      color: #555;
      transition: all 0.2s ease;
    }
    
    .add-child-number-button:hover {
      background-color: #f0f0f0;
      color: #a85733;
    }
    
    .add-child-number-button svg {
      display: block;
    }
    
    /* Notification toast styling */
    .notification-toast {
      position: fixed;
      bottom: 30px;
      left: 50%;
      transform: translateX(-50%);
      background-color: white;
      color: #333;
      padding: 12px 16px;
      border-radius: 8px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      display: flex;
      align-items: center;
      z-index: 10000;
      opacity: 0;
      visibility: hidden;
      transition: opacity 0.3s ease, visibility 0.3s ease;
      font-family: 'Public Sans', sans-serif;
      font-size: 14px;
      max-width: 400px;
    }
    
    .notification-toast.visible {
      opacity: 1;
      visibility: visible;
    }
    
    .notification-icon {
      margin-right: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    
    .notification-success .notification-icon {
      color: #28a745;
    }
    
    .notification-warning .notification-icon {
      color: #ffc107;
    }
    
    .notification-error .notification-icon {
      color: #dc3545;
    }
    
    .notification-info .notification-icon {
      color: #17a2b8;
    }
    
    /* Drag handle for vertical movement */
    .image-drag-handle {
      position: absolute;
      left: 50%;
      transform: translateX(-50%);
      width: 30px;
      height: 12px;
      background-color: rgba(0, 120, 215, 0.7);
      border-radius: 6px;
      cursor: ns-resize;
      opacity: 0;
      transition: opacity 0.2s ease;
      z-index: 100;
    }
    
    .image-drag-handle::after {
      content: '';
      position: absolute;
      top: 4px;
      left: 7px;
      width: 16px;
      height: 4px;
      background-color: white;
      border-radius: 2px;
    }
    
    .image-drag-handle.top {
      top: -6px;
    }
    
    .image-drag-handle.bottom {
      bottom: -6px;
    }
    
    .person-image-container:hover .image-drag-handle {
      opacity: 0.8;
    }
    
    /* Styles for draggable image container already defined above */
    .image-drag-handle:hover {
      opacity: 1 !important;
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
      
      // Save state after edit
      saveStateAfterEdit();
    });
    
    // Create bold button
    const boldBtn = document.createElement('button');
    boldBtn.className = 'bold-btn';
    boldBtn.title = 'Bold text (Ctrl+B)';
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
      
      // Save state after edit
      saveStateAfterEdit();
    });
    
    // Create brown text color button
    const colorBtn = document.createElement('button');
    colorBtn.className = 'color-btn';
    colorBtn.title = 'Apply theme color (Ctrl+J)';
    colorBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#a85733" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8"/>
    </svg>`;
    colorBtn.addEventListener('click', function(e) {
      e.stopPropagation(); // Prevent event bubbling
      e.preventDefault(); // Prevent default
      
      // Make sure the paragraph has focus
      paraDiv.focus();
      
      // Check if there's a selection
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
      
      // Apply theme color to the selection
      document.execCommand('foreColor', false, '#a85733');
      
      // Don't clear selection if user selected specific text
      // Only clear if we auto-selected the whole paragraph
      if (!selection.rangeCount || selection.getRangeAt(0).toString().trim() === paraDiv.textContent.trim()) {
        selection.removeAllRanges();
      }
      
      // Update the raw data with the formatting changes
      if (person && personContent) {
        updatePersonRawWithUserBreaks(person, personContent);
      }
      
      // Save state after edit
      saveStateAfterEdit();
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
        
        // Save state after edit
        saveStateAfterEdit();
      }, 300);
    });
    
    // Create underline button
    const underlineBtn = document.createElement('button');
    underlineBtn.className = 'underline-btn';
    underlineBtn.title = 'Underline text (Ctrl+U)';
    underlineBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M6 3v7a6 6 0 0 0 12 0V3"/>
      <line x1="4" y1="21" x2="20" y2="21"/>
    </svg>`;
    underlineBtn.addEventListener('click', function(e) {
      e.stopPropagation(); // Prevent event bubbling
      e.preventDefault(); // Prevent default
      
      // Make sure the paragraph has focus
      paraDiv.focus();
      
      // Check if there's a selection
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
      
      // Apply underline to the selection
      document.execCommand('underline', false, null);
      
      // Don't clear selection if user selected specific text
      // Only clear if we auto-selected the whole paragraph
      if (!selection.rangeCount || selection.getRangeAt(0).toString().trim() === paraDiv.textContent.trim()) {
        selection.removeAllRanges();
      }
      
      // Update the raw data with the formatting changes
      if (person && personContent) {
        updatePersonRawWithUserBreaks(person, personContent);
      }
      
      // Save state after edit
      saveStateAfterEdit();
    });
    
    // Create hyperlink style button (combines bold, theme color, and underline)
    const hyperlinkBtn = document.createElement('button');
    hyperlinkBtn.className = 'hyperlink-btn';
    hyperlinkBtn.title = 'Apply hyperlink style (Ctrl+D)';
    hyperlinkBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#a85733" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>
      <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>
    </svg>`;
    hyperlinkBtn.addEventListener('click', function(e) {
      e.stopPropagation(); // Prevent event bubbling
      e.preventDefault(); // Prevent default
      
      // Make sure the paragraph has focus
      paraDiv.focus();
      
      // Check if there's a selection
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
      
      // Apply all three formatting styles at once (hyperlink style)
      document.execCommand('bold', false, null);
      document.execCommand('foreColor', false, '#a85733');
      document.execCommand('underline', false, null);
      
      // Don't clear selection if user selected specific text
      // Only clear if we auto-selected the whole paragraph
      if (!selection.rangeCount || selection.getRangeAt(0).toString().trim() === paraDiv.textContent.trim()) {
        selection.removeAllRanges();
      }
      
      // Update the raw data with the formatting changes
      if (person && personContent) {
        updatePersonRawWithUserBreaks(person, personContent);
      }
      
      // Save state after edit
      saveStateAfterEdit();
    });
    
    // Add buttons to controls
    controls.appendChild(editBtn);
    controls.appendChild(boldBtn);
    controls.appendChild(colorBtn);
    controls.appendChild(underlineBtn);
    controls.appendChild(hyperlinkBtn);
    controls.appendChild(deleteBtn);
    
    // Add controls to paragraph
    paraDiv.appendChild(controls);
    
    // Add blur event to save state after editing text
    paraDiv.addEventListener('blur', function() {
      saveStateAfterEdit();
    });
  }
  
  // Helper function to debounce page state saving
  const saveStateAfterEdit = debounce(function() {
    if (window.pageManager) {
      window.pageManager.saveCurrentPageState();
      window.pageManager.saveState();
    }
  }, 500);
  
  // Debounce utility function
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

  // Update the paragraph creation code to add controls
  function createEditableSubparagraph(allSubParagraphText, person, personContent) {
    const paraDiv = document.createElement('div');
    paraDiv.className = 'indented-paragraph';
    
    // Process for children entries - updated regex to match both formats:
    // 1. "60 i. Christoph Jakob Muller" (number without parentheses)
    // 2. "(2) v. John Walter Rogers" (number with parentheses)
    const childEntryPattern = /^(?:\((\d+)\)|(\d+))\s+([ivxlcdm]+)\.\s+([A-ZÅÄÖÜ][a-zåäöü'-]+(?:\s+[A-ZÅÄÖÜ][a-zåäöü'-]+)*)([\s\S]*)/i;
    let processedText = allSubParagraphText;
    const childMatch = allSubParagraphText.match(childEntryPattern);

    if (childMatch) {
      // We found a child entry pattern
      const childNumber = childMatch[1] || childMatch[2]; // The number (either with or without parentheses)
      const romanNumeral = childMatch[3]; // The roman numeral
      const childName = childMatch[4]; // The child's name
      const restOfText = childMatch[5]; // The rest of the text (details)
      
      // Create the HTML with proper formatting
      let childEntryHtml = '';
      
      // Add the circular number regardless of original format
      // Check if number is large (more than 3 digits) to add special class
      const isLargeNumber = childNumber.length > 3;
      childEntryHtml += `<span class="child-number${isLargeNumber ? ' large-number' : ''}">${childNumber}</span> `;
      
      // Add the roman numeral
      childEntryHtml += `${romanNumeral}. `;
      
      // Add the child name in bold
      childEntryHtml += `<b>${preserveHyperlinks(childName)}</b>`;
      
      // Add the rest of the text
      childEntryHtml += preserveHyperlinks(restOfText);
      
      processedText = childEntryHtml;
    } else {
      // Not a child entry, just preserve any links
      processedText = preserveHyperlinks(allSubParagraphText);
    }
    
    // Set the processed content
    paraDiv.innerHTML = processedText;
    
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
    
    // Function to check if current selection has formatting applied
    function checkSelectionFormatting() {
      const controls = paraDiv.querySelector('.paragraph-controls');
      if (!controls) return;
      
      const boldBtn = controls.querySelector('.bold-btn');
      const colorBtn = controls.querySelector('.color-btn');
      const underlineBtn = controls.querySelector('.underline-btn');
      const hyperlinkBtn = controls.querySelector('.hyperlink-btn');
      
      if (!boldBtn) return;
      
      try {
        // Only check if this paragraph is active
        if (document.activeElement !== paraDiv) {
          boldBtn.classList.remove('active-format');
          if (colorBtn) colorBtn.classList.remove('active-format');
          if (underlineBtn) underlineBtn.classList.remove('active-format');
          if (hyperlinkBtn) hyperlinkBtn.classList.remove('active-format');
          return;
        }
        
        // Check formatting states
        const isBold = document.queryCommandState('bold');
        const isUnderlined = document.queryCommandState('underline');
        
        // Update button states
        if (isBold) {
          boldBtn.classList.add('active-format');
        } else {
          boldBtn.classList.remove('active-format');
        }
        
        if (underlineBtn) {
          if (isUnderlined) {
            underlineBtn.classList.add('active-format');
          } else {
            underlineBtn.classList.remove('active-format');
          }
        }
        
        // Check if current selection has the theme color
        const hasThemeColor = document.queryCommandValue('foreColor').toLowerCase().includes('a85733');
        
        if (colorBtn) {
          if (hasThemeColor) {
            colorBtn.classList.add('active-format');
          } else {
            colorBtn.classList.remove('active-format');
          }
        }
        
        // If all three formats are applied, highlight the hyperlink button
        if (hyperlinkBtn) {
          if (isBold && isUnderlined && hasThemeColor) {
            hyperlinkBtn.classList.add('active-format');
          } else {
            hyperlinkBtn.classList.remove('active-format');
          }
        }
      } catch (e) {
        console.error('Error checking formatting state:', e);
      }
    }
    
    // Add hover controls for editing/deleting
    addParagraphControls(paraDiv, person, personContent);
    
    // Make the paragraph resizable with the handles
    makeResizable(paraDiv);
    
    return paraDiv;
  }

  // Helper function to update raw data when paragraph is removed
  function updatePersonRawWithRemovedParagraph(personObject, personContentElement, removedParagraph) {
    console.log(`Paragraph removed from person ${personObject.number}`);
    // In a real implementation, we would update personObject.raw to remove the paragraph
    // For now, this is just a placeholder
  }

  // Add global keyboard shortcuts for formatting
  document.addEventListener('keydown', function(e) {
    // Get the active element (focused element)
    const activeElement = document.activeElement;
    
    // Only proceed with formatting shortcuts if we're in an editable paragraph
    const isEditableParagraph = activeElement && (
      activeElement.classList.contains('main-person-paragraph') || 
      activeElement.classList.contains('indented-paragraph') ||
      activeElement.classList.contains('editable-subparagraph')
    );
    
    // Check if Ctrl+J is pressed (for theme color)
    if (e.ctrlKey && e.key === 'j' && isEditableParagraph) {
      e.preventDefault(); // Prevent default browser action
      
      // Get current selection
      const selection = window.getSelection();
      
      // If there's a valid selection range
      if (selection.rangeCount > 0) {
        // Apply theme color
        document.execCommand('foreColor', false, '#a85733');
        console.log('Applied theme color with Ctrl+J');
      }
    }
    
    // Check if Ctrl+U is pressed (for underline)
    if (e.ctrlKey && e.key === 'u' && isEditableParagraph) {
      e.preventDefault(); // Prevent default browser action
      
      // Get current selection
      const selection = window.getSelection();
      
      // If there's a valid selection range
      if (selection.rangeCount > 0) {
        // Apply underline
        document.execCommand('underline', false, null);
        console.log('Applied underline with Ctrl+U');
      }
    }
    
    // Check if Ctrl+D is pressed (for combined hyperlink styling: bold + theme color + underline)
    if (e.ctrlKey && e.key === 'd' && isEditableParagraph) {
      e.preventDefault(); // Prevent default browser action
      
      // Get current selection
      const selection = window.getSelection();
      
      // If there's a valid selection range
      if (selection.rangeCount > 0) {
        // Apply all three formatting styles at once
        document.execCommand('bold', false, null);
        document.execCommand('foreColor', false, '#a85733');
        document.execCommand('underline', false, null);
        console.log('Applied hyperlink styling (bold, theme color, underline) with Ctrl+D');
        
        // Show a subtle notification
        showNotification('Hyperlink style applied', 'success');
      }
    }
    
    // Check if Ctrl+Q is pressed (for child number)
    if (e.ctrlKey && e.key === 'q') {
      e.preventDefault(); // Prevent default browser action
      
      // Insert child number at cursor position
      insertChildNumberAtCursor();
      console.log('Child number insertion triggered with Ctrl+Q');
    }
    
    // Check if Ctrl+Space is pressed (for hyperlink functionality)
    if (e.ctrlKey && e.key === ' ' && isEditableParagraph) {
      e.preventDefault(); // Prevent default browser action
      
      // Get current selection
      const selection = window.getSelection();
      
      // If there's a valid selection range
      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        
        // Only proceed if text is selected
        if (!range.collapsed) {
          // Show hyperlink popup
          showHyperlinkPopup(selection, range);
          console.log('Hyperlink popup shown with Ctrl+Space');
        } else {
          // Show notification if no text is selected
          showNotification('Select text first to create a hyperlink', 'warning');
        }
      }
    }
  });

  // Add CSS for paragraph resizing functionality
  const resizeHandlesStyles = document.createElement('style');
  resizeHandlesStyles.textContent = `
    .resizable-paragraph {
      position: relative;
    }
    
    .resize-handle {
      position: absolute;
      width: 8px;
      height: 40px;
      background-color: rgba(0, 120, 215, 0.7);
      border-radius: 4px;
      cursor: col-resize;
      opacity: 0;
      transition: opacity 0.2s ease;
      z-index: 100;
    }
    
    .resize-handle.left {
      left: -4px;
      top: 50%;
      transform: translateY(-50%);
    }
    
    .resize-handle.right {
      right: -4px;
      top: 50%;
      transform: translateY(-50%);
    }
    
    .resizable-paragraph:hover .resize-handle,
    .resizable-paragraph.resizing .resize-handle {
      opacity: 1;
    }
    
    .resizable-paragraph.resizing {
      cursor: col-resize;
      user-select: none;
    }
    
    .resize-overlay {
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      border: 2px dashed rgba(0, 120, 215, 0.7);
      pointer-events: none;
      display: none;
      z-index: 99;
    }
    
    .resizable-paragraph.resizing .resize-overlay {
      display: block;
    }
  `;
  document.head.appendChild(resizeHandlesStyles);

  // Debounce utility function
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

  // Make paragraph resizable with handles on sides
  function makeResizable(element) {
    // Make sure we don't add duplicate handlers
    if (element.classList.contains('resizable-paragraph')) {
      return;
    }
    
    // Add the resizable class
    element.classList.add('resizable-paragraph');
    
    // Create resize handles and overlay
    const leftHandle = document.createElement('div');
    leftHandle.className = 'resize-handle left';
    
    const rightHandle = document.createElement('div');
    rightHandle.className = 'resize-handle right';
    
    const resizeOverlay = document.createElement('div');
    resizeOverlay.className = 'resize-overlay';
    
    // Add handles and overlay to the element
    element.appendChild(leftHandle);
    element.appendChild(rightHandle);
    element.appendChild(resizeOverlay);
    
    // Variables to track resize state
    let isResizing = false;
    let startX = 0;
    let startWidth = 0;
    let handle = null;
    
    // Handle mousedown on resize handles
    function startResize(e) {
      e.preventDefault();
      e.stopPropagation();
      
      isResizing = true;
      handle = e.target;
      startX = e.clientX;
      startWidth = element.offsetWidth;
      
      // Add resizing class to element
      element.classList.add('resizing');
      
      // Add event listeners for resize
      document.addEventListener('mousemove', resize);
      document.addEventListener('mouseup', stopResize);
      
      // Prevent text selection during resize
      document.body.style.userSelect = 'none';
    }
    
    // Handle mousemove during resize
    function resize(e) {
      if (!isResizing) return;
      
      let newWidth;
      
      if (handle.classList.contains('right')) {
        // Resize from right handle
        newWidth = startWidth + (e.clientX - startX);
      } else {
        // Resize from left handle
        const deltaX = e.clientX - startX;
        newWidth = startWidth - deltaX;
        
        // If resizing from left, adjust the min width
        if (newWidth < 200) {
          newWidth = 200;
        }
      }
      
      // Apply max width constraint if needed
      const maxWidth = element.parentElement.offsetWidth - 40; // 40px margin
      if (newWidth > maxWidth) {
        newWidth = maxWidth;
      }
      
      // Update element width
      element.style.width = `${newWidth}px`;
      
      // If this is a main paragraph with an image, reflow the image
      reflowImageIfNeeded(element);
    }
    
    // Handle mouseup after resize
    function stopResize() {
      isResizing = false;
      
      // Remove resizing class
      element.classList.remove('resizing');
      
      // Remove event listeners
      document.removeEventListener('mousemove', resize);
      document.removeEventListener('mouseup', stopResize);
      
      // Re-enable text selection
      document.body.style.userSelect = '';
      
      // Save state after resize
      saveStateAfterEdit();
    }
    
    // Call the global reflowImageIfNeeded function
    reflowImageIfNeeded(element);
    
    // Add event listeners to handles
    leftHandle.addEventListener('mousedown', startResize);
    rightHandle.addEventListener('mousedown', startResize);
    
    // Initial reflow if needed
    reflowImageIfNeeded(element);
  }

  // Helper function to handle image reflow
  function reflowImageIfNeeded(element) {
    // Check if this is a main paragraph with an image
    if (element.classList.contains('main-person-paragraph')) {
      const personContent = element.parentElement;
      if (!personContent) return;
      
      const personImageContainer = personContent.querySelector('.person-image-container');
      if (personImageContainer) {
        // If image exists and paragraph width is less than 80% of content width,
        // position the image appropriately
        const contentWidth = personContent.offsetWidth;
        const paragraphWidth = element.offsetWidth;
        
        if (paragraphWidth < contentWidth * 0.8) {
          // There's space for the image to float right
          personImageContainer.style.float = 'right';
          personImageContainer.style.clear = 'right';
          personImageContainer.style.marginLeft = '15px';
        } else {
          // Full width paragraph, center the image
          personImageContainer.style.float = 'none';
          personImageContainer.style.clear = 'both';
          personImageContainer.style.margin = '10px auto';
          personImageContainer.style.display = 'block';
        }
      }
    }
    
    // For indented paragraphs, just apply the width
    if (element.classList.contains('indented-paragraph')) {
      // The width is already being set directly
    }
  }

  // Make an image container vertically draggable
  function makeImageDraggable(imageContainer, personContent) {
    // Make sure we don't add duplicate handlers
    if (imageContainer.classList.contains('vertically-draggable')) {
      return;
    }
    
    // Mark as draggable
    imageContainer.classList.add('vertically-draggable');
    
    // Create drag handles
    const topHandle = document.createElement('div');
    topHandle.className = 'image-drag-handle top';
    
    const bottomHandle = document.createElement('div');
    bottomHandle.className = 'image-drag-handle bottom';
    
    // Add handles to container
    imageContainer.appendChild(topHandle);
    imageContainer.appendChild(bottomHandle);
    
    // Variables to track dragging state
    let isDragging = false;
    let startY = 0;
    let startTop = 0;
    let currentHandle = null;
    
    // Handler for starting the drag
    function startDrag(e) {
      e.preventDefault();
      e.stopPropagation();
      
      // Set dragging state
      isDragging = true;
      currentHandle = e.target;
      startY = e.clientY;
      
      // Get the current computed top value
      const computedStyle = window.getComputedStyle(imageContainer);
      startTop = parseInt(computedStyle.marginTop) || 0;
      
      // Add dragging class
      imageContainer.classList.add('dragging');
      
      // Add event listeners for drag movement
      document.addEventListener('mousemove', drag);
      document.addEventListener('mouseup', stopDrag);
      
      // Prevent text selection during drag
      document.body.style.userSelect = 'none';
    }
    
    // Handler for dragging
    function drag(e) {
      if (!isDragging) return;
      
      // Calculate new top position based on the drag delta
      const deltaY = e.clientY - startY;
      let newMarginTop = startTop + deltaY;
      
      // Set a minimum margin-top value to keep the image visible
      const minMarginTop = -imageContainer.offsetHeight + 30;
      if (newMarginTop < minMarginTop) {
        newMarginTop = minMarginTop;
      }
      
      // Set a maximum margin-top to prevent image from going too far down
      const maxMarginTop = personContent.offsetHeight - 100;
      if (newMarginTop > maxMarginTop) {
        newMarginTop = maxMarginTop;
      }
      
      // Apply the new position
      imageContainer.style.marginTop = `${newMarginTop}px`;
      
      // Reflow text around the image as it moves
      reflowTextAroundMovedImage(imageContainer, personContent);
    }
    
    // Handler for stopping the drag
    function stopDrag() {
      if (!isDragging) return;
      
      // Reset dragging state
      isDragging = false;
      currentHandle = null;
      
      // Remove dragging class
      imageContainer.classList.remove('dragging');
      
      // Remove event listeners
      document.removeEventListener('mousemove', drag);
      document.removeEventListener('mouseup', stopDrag);
      
      // Re-enable text selection
      document.body.style.userSelect = '';
      
      // Final reflow of text
      reflowTextAroundMovedImage(imageContainer, personContent);
      
      // Save state after drag is complete
      saveStateAfterEdit();
    }
    
    // Also make the container itself draggable (not just the handles)
    imageContainer.addEventListener('mousedown', function(e) {
      // Only handle dragging from the container itself, not from children (like the image or delete button)
      if (e.target === imageContainer) {
        startDrag(e);
      }
    });
    
    // Add event listeners to handles
    topHandle.addEventListener('mousedown', startDrag);
    bottomHandle.addEventListener('mousedown', startDrag);
  }
  
  // Helper function to reflow text around a moved image
  function reflowTextAroundMovedImage(imageContainer, personContent) {
    // Find all paragraphs that might need reflowing
    const paragraphs = personContent.querySelectorAll('.main-person-paragraph, .indented-paragraph');
    
    // Get image position info
    const imageRect = imageContainer.getBoundingClientRect();
    const containerRect = personContent.getBoundingClientRect();
    
    // Calculate relative positions
    const imageTop = imageRect.top - containerRect.top;
    const imageBottom = imageRect.bottom - containerRect.top;
    
    // Check each paragraph to see if it needs adjustment based on image position
    paragraphs.forEach(paragraph => {
      const paraRect = paragraph.getBoundingClientRect();
      const paraTop = paraRect.top - containerRect.top;
      const paraBottom = paraRect.bottom - containerRect.top;
      
      // Determine if this paragraph overlaps with the image vertically
      const overlapsImage = (paraBottom > imageTop && paraTop < imageBottom);
      
      if (overlapsImage) {
        // If paragraph overlaps image vertically, we need to adjust its width
        // if it's also not already resized by the user
        if (!paragraph.style.width) {
          const mainParagraphWidth = personContent.offsetWidth - imageRect.width - 30; // 30px for spacing
          paragraph.style.width = `${mainParagraphWidth}px`;
        }
      } else {
        // If no overlap, the paragraph can use the full width
        if (!paragraph.style.width || paragraph.style.width === 'auto') {
          paragraph.style.width = 'auto';
        }
      }
    });
  }

  /**
 * Creates and inserts a new generation title at the specified position
 * @param {HTMLElement} container - The legacy sheet container
 * @param {HTMLElement} insertAfter - The element to insert after (optional)
 */
function insertGenerationTitle(container, insertAfter = null) {
  // Create the generation title elements
  const titleContainer = document.createElement('div');
  titleContainer.className = 'generation-title-container';
  
  const generationTitle = document.createElement('div');
  generationTitle.className = 'generation-title';
  generationTitle.textContent = 'New Generation';
  generationTitle.setAttribute('contenteditable', 'true');
  
  const horizontalRule = document.createElement('hr');
  horizontalRule.className = 'horizontal-rule';
  
  // Create delete button for the generation title
  const deleteButton = document.createElement('button');
  deleteButton.className = 'generation-title-delete-btn';
  deleteButton.innerHTML = '×';
  deleteButton.title = 'Delete generation title';
  
  // Add delete functionality to the delete button
  deleteButton.addEventListener('click', function(e) {
    e.stopPropagation(); // Prevent triggering other click events
    
    // Add fade-out effect
    titleContainer.style.transition = 'opacity 0.3s ease, transform 0.3s ease, height 0.3s ease, margin 0.3s ease';
    titleContainer.style.opacity = '0';
    titleContainer.style.transform = 'translateY(-10px)';
    titleContainer.style.height = '0';
    titleContainer.style.marginTop = '0';
    titleContainer.style.marginBottom = '0';
    titleContainer.style.overflow = 'hidden';
    
    // Remove the container after the animation completes
    setTimeout(() => {
      titleContainer.remove();
      
      // Save state after deletion
      if (window.pageManager) {
        window.pageManager.saveCurrentPageState();
        window.pageManager.saveState();
      }
    }, 300);
  });
  
  // Add click event to the horizontal rule for deletion
  horizontalRule.addEventListener('click', function(e) {
    // Since the horizontal rule is part of the title container, just trigger the delete button click
    deleteButton.click();
  });
  
  // Add elements to the container
  titleContainer.appendChild(generationTitle);
  titleContainer.appendChild(deleteButton);
  titleContainer.appendChild(horizontalRule);
  
  // Create a fade-in animation effect
  titleContainer.style.opacity = '0';
  titleContainer.style.transform = 'translateY(-10px)';
  titleContainer.style.transition = 'opacity 0.3s ease, transform 0.3s ease';
  
  // Position the new title elements
  if (insertAfter) {
    // Insert after the specified element
    if (insertAfter.nextElementSibling) {
      container.insertBefore(titleContainer, insertAfter.nextElementSibling);
    } else {
      container.appendChild(titleContainer);
    }
  } else {
    // If no position specified, add at the beginning (after existing title if any)
    const existingTitle = container.querySelector('.generation-title');
    if (existingTitle) {
      const existingTitleContainer = existingTitle.closest('.generation-title-container');
      if (existingTitleContainer && existingTitleContainer.nextElementSibling) {
        container.insertBefore(titleContainer, existingTitleContainer.nextElementSibling);
      } else {
        container.appendChild(titleContainer);
      }
    } else {
      container.prepend(titleContainer);
    }
  }
    
    // Make the generation title editable
    // Add styling for better UX when editing
    generationTitle.style.outline = 'none';
    generationTitle.style.transition = 'background-color 0.2s ease';
    
    // Add hover effect
    generationTitle.addEventListener('mouseover', function() {
      this.style.backgroundColor = 'rgba(240, 240, 240, 0.5)';
      this.title = 'Click to edit generation title';
    });
    
    generationTitle.addEventListener('mouseout', function() {
      this.style.backgroundColor = 'transparent';
    });
    
    // Add focus/blur effects
    generationTitle.addEventListener('focus', function() {
      this.style.backgroundColor = 'rgba(240, 240, 240, 0.8)';
    });
    
    generationTitle.addEventListener('blur', function() {
      this.style.backgroundColor = 'transparent';
      
      // Save state when generation title is edited
      if (window.pageManager) {
        window.pageManager.saveCurrentPageState();
        window.pageManager.saveState();
      }
    });
    
    // Prevent Enter key from creating new lines
    generationTitle.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.blur(); // Remove focus when Enter is pressed
      }
    });
    
    // Trigger the animation
    setTimeout(() => {
      titleContainer.style.opacity = '1';
      titleContainer.style.transform = 'translateY(0)';
    }, 10);
    
    // Focus on the title for immediate editing
    setTimeout(() => {
      generationTitle.focus();
      
      // Select all text for easy replacement
      const range = document.createRange();
      range.selectNodeContents(generationTitle);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    }, 300);
    
    // Save state after adding the new title
    if (window.pageManager) {
      setTimeout(() => {
        window.pageManager.saveCurrentPageState();
        window.pageManager.saveState();
      }, 500);
    }
    
    return titleContainer;
  }

  // Add an insertion control to the person entries to allow adding generation titles
  function addInsertionControls() {
    // Select all person entries and entry dividers
    const personEntries = document.querySelectorAll('.person-entry');
    const entryDividers = document.querySelectorAll('.entry-divider');
    
    // Add insertion controls to each person entry
    personEntries.forEach(entry => {
      addInsertionControl(entry);
    });
    
    // Add insertion controls to each entry divider
    entryDividers.forEach(divider => {
      addInsertionControl(divider);
    });
  }

  // Add insertion control to a specific element
  function addInsertionControl(element) {
    // Create the insertion control
    const insertControl = document.createElement('div');
    insertControl.className = 'insert-control';
    insertControl.innerHTML = `
      <button class="insert-title-btn" title="Insert generation title">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
          <line x1="12" y1="5" x2="12" y2="19"></line>
          <line x1="5" y1="12" x2="19" y2="12"></line>
        </svg>
        <span>Add Generation Title</span>
      </button>
    `;
    
    // Add click event to the insert button
    const insertButton = insertControl.querySelector('.insert-title-btn');
    insertButton.addEventListener('click', function(e) {
      e.stopPropagation();
      
      // Get the container and insert the new generation title
      const legacySheet = element.closest('.legacy-sheet');
      insertGenerationTitle(legacySheet, element);
      
      // Hide the control after use
      insertControl.style.opacity = '0';
      setTimeout(() => {
        insertControl.style.display = 'none';
      }, 300);
    });
    
    // Add hover behavior to show the control
    element.addEventListener('mouseover', function() {
      // Position the control at the center top of the element
      const rect = element.getBoundingClientRect();
      insertControl.style.top = `${-30}px`;
      insertControl.style.left = '50%';
      insertControl.style.transform = 'translateX(-50%)';
      
      // Show the control
      insertControl.style.display = 'block';
      setTimeout(() => {
        insertControl.style.opacity = '1';
      }, 10);
    });
    
    element.addEventListener('mouseout', function(e) {
      // Check if we're not moving to the control itself
      if (!e.relatedTarget || !insertControl.contains(e.relatedTarget)) {
        insertControl.style.opacity = '0';
        setTimeout(() => {
          insertControl.style.display = 'none';
        }, 300);
      }
    });
    
    // Style the element to be a position reference
    if (!element.style.position || element.style.position === 'static') {
      element.style.position = 'relative';
    }
    
    // Add the control to the element
    element.appendChild(insertControl);
  }

  // Setup mutation observer to add insertion controls to new elements and handle generation titles
  function setupInsertionControlsObserver() {
    const observer = new MutationObserver(mutations => {
      mutations.forEach(mutation => {
        if (mutation.addedNodes.length) {
          mutation.addedNodes.forEach(node => {
            // Check if node is an Element
            if (node.nodeType === 1) {
              // Add insertion controls to person entries and dividers
              if (node.classList.contains('person-entry') || node.classList.contains('entry-divider')) {
                addInsertionControl(node);
              }
              
              // Handle new pages
              if (node.classList.contains('legacy-sheet')) {
                // Add insertion controls to person entries and dividers in the new page
                const personEntries = node.querySelectorAll('.person-entry');
                const entryDividers = node.querySelectorAll('.entry-divider');
                
                personEntries.forEach(entry => {
                  addInsertionControl(entry);
                });
                
                entryDividers.forEach(divider => {
                  addInsertionControl(divider);
                });
                
                // Make sure all generation titles in the new page have delete buttons
                const generationTitles = node.querySelectorAll('.generation-title-container');
                generationTitles.forEach(titleContainer => {
                  // Check if this container already has a delete button
                  if (!titleContainer.querySelector('.generation-title-delete-btn')) {
                    // Create and add a delete button
                    const deleteButton = document.createElement('button');
                    deleteButton.className = 'generation-title-delete-btn';
                    deleteButton.innerHTML = '×';
                    deleteButton.title = 'Delete generation title';
                    
                    // Add delete functionality to the delete button
                    deleteButton.addEventListener('click', function(e) {
                      e.stopPropagation(); // Prevent triggering other click events
                      
                      // Add fade-out effect
                      titleContainer.style.transition = 'opacity 0.3s ease, transform 0.3s ease, height 0.3s ease, margin 0.3s ease';
                      titleContainer.style.opacity = '0';
                      titleContainer.style.transform = 'translateY(-10px)';
                      titleContainer.style.height = '0';
                      titleContainer.style.marginTop = '0';
                      titleContainer.style.marginBottom = '0';
                      titleContainer.style.overflow = 'hidden';
                      
                      // Remove the container after the animation completes
                      setTimeout(() => {
                        titleContainer.remove();
                        
                        // Save state after deletion
                        if (window.pageManager) {
                          window.pageManager.saveCurrentPageState();
                          window.pageManager.saveState();
                        }
                      }, 300);
                    });
                    
                    // Add the delete button to the title container
                    const generationTitle = titleContainer.querySelector('.generation-title');
                    if (generationTitle) {
                      titleContainer.insertBefore(deleteButton, generationTitle.nextSibling);
                    } else {
                      titleContainer.appendChild(deleteButton);
                    }
                    
                    // Add click event to the horizontal rule for deletion
                    const horizontalRule = titleContainer.querySelector('.horizontal-rule');
                    if (horizontalRule) {
                      horizontalRule.addEventListener('click', function(e) {
                        // Trigger the delete button click
                        deleteButton.click();
                      });
                    }
                  }
                });
              }
            }
          });
        }
      });
    });
    
    // Observe the pages container
    const pagesContainer = document.getElementById('pages-container');
    if (pagesContainer) {
      observer.observe(pagesContainer, { 
        childList: true,
        subtree: true
      });
    }
  }

  // Initialize insertion controls when the DOM is loaded
  document.addEventListener('DOMContentLoaded', function() {
    // ... existing code ...
    
    // Initialize insertion controls
    addInsertionControls();
    setupInsertionControlsObserver();
    
    // ... existing code ...
  });

  /**
 * Creates the Add Generation Title button in the UI
 */
function createAddGenerationTitleButton() {
  // Check if button already exists
  if (document.getElementById('add-generation-title-btn')) {
    return;
  }
  
  // Create the button
  const addGenerationTitleBtn = document.createElement('button');
  addGenerationTitleBtn.id = 'add-generation-title-btn';
  addGenerationTitleBtn.className = 'pagination-button add-generation-title-button';
  addGenerationTitleBtn.title = 'Add generation title';
  addGenerationTitleBtn.innerHTML = `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" width="1.2em" height="1.2em">
      <path d="M6 9h12M6 15h12M9 4.5v15M15 4.5v15"/>
    </svg>
  `;
  
  // Add click event to the button
  addGenerationTitleBtn.addEventListener('click', function() {
    // Find the current active page
    const activePage = document.querySelector('.legacy-sheet.active-page');
    if (!activePage) return;
    
    // Insert a new generation title at the top of the page
    insertGenerationTitle(activePage);
    
    // Show success feedback
    addGenerationTitleBtn.classList.add('button-success');
    setTimeout(() => {
      addGenerationTitleBtn.classList.remove('button-success');
    }, 1000);
  });
  
  // Add the button to the pagination controls after the page indicator
  const paginationControls = document.querySelector('.pagination-controls');
  const pageIndicator = document.getElementById('page-indicator');
  
  if (paginationControls && pageIndicator) {
    // Create a divider
    const divider = document.createElement('div');
    divider.className = 'pagination-divider';
    
    // Insert in the DOM
    paginationControls.insertBefore(divider, pageIndicator.nextSibling);
    paginationControls.insertBefore(addGenerationTitleBtn, divider.nextSibling);
  }
}

/**
 * Creates the Add Child Number button in the UI
 */
function createAddChildNumberButton() {
  // Check if button already exists
  if (document.getElementById('add-child-number-btn')) {
    return;
  }
  
  // Create the button
  const addChildNumberBtn = document.createElement('button');
  addChildNumberBtn.id = 'add-child-number-btn';
  addChildNumberBtn.className = 'pagination-button add-child-number-button';
  addChildNumberBtn.title = 'Add child number (Ctrl+Q)';
  addChildNumberBtn.innerHTML = `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#a85733" stroke-width="1.5" width="1.2em" height="1.2em">
      <circle cx="12" cy="12" r="8"/>
      <text x="12" y="16" font-size="10" text-anchor="middle" fill="#a85733" font-family="sans-serif">n</text>
    </svg>
  `;
  
  // Add click event to the button
  addChildNumberBtn.addEventListener('click', function() {
    // Insert child number at cursor position
    insertChildNumberAtCursor();
    
    // Show success feedback
    addChildNumberBtn.classList.add('button-success');
    setTimeout(() => {
      addChildNumberBtn.classList.remove('button-success');
    }, 1000);
  });
  
  // Add the button to the pagination controls after the Add Generation Title button
  const paginationControls = document.querySelector('.pagination-controls');
  const addGenerationTitleBtn = document.getElementById('add-generation-title-btn');
  
  if (paginationControls && addGenerationTitleBtn) {
    // Create a divider
    const divider = document.createElement('div');
    divider.className = 'pagination-divider';
    
    // Insert in the DOM
    paginationControls.insertBefore(divider, addGenerationTitleBtn.nextSibling);
    paginationControls.insertBefore(addChildNumberBtn, divider.nextSibling);
  } else if (paginationControls) {
    // Fallback if Add Generation Title button not found
    paginationControls.appendChild(addChildNumberBtn);
  }
}

/**
 * Function to insert a child number at the current cursor position
 */
function insertChildNumberAtCursor() {
  // Get current selection
  const selection = window.getSelection();
  if (!selection.rangeCount) {
    showNotification('Please position cursor in an editable paragraph first', 'warning');
    return;
  }
  
  // Get the range and its container
  const range = selection.getRangeAt(0);
  const container = range.startContainer.parentNode;
  
  // Only proceed if we're in an editable paragraph
  if (!container.closest('.editable-subparagraph') && 
      !container.closest('.main-person-paragraph') && 
      !container.closest('.indented-paragraph')) {
    // Show notification that we need to be in an editable area
    showNotification('Please position cursor in an editable paragraph', 'warning');
    return;
  }
  
  // Prompt for the child number
  const childNumber = prompt('Enter the child number:', '1');
  if (!childNumber) return; // User cancelled
  
  // Create the child number element
  const isLargeNumber = childNumber.length > 3;
  const childNumberElement = document.createElement('span');
  childNumberElement.className = `child-number${isLargeNumber ? ' large-number' : ''}`;
  childNumberElement.textContent = childNumber;
  
  // Insert at cursor position
  range.deleteContents();
  range.insertNode(childNumberElement);
  
  // Move cursor after the inserted element
  range.setStartAfter(childNumberElement);
  range.setEndAfter(childNumberElement);
  selection.removeAllRanges();
  selection.addRange(range);
  
  // Save state after insertion
  if (window.pageManager) {
    window.pageManager.saveCurrentPageState();
    window.pageManager.saveState();
  }
  
  // Show success notification
  showNotification('Child number inserted successfully', 'success');
}

/**
 * Shows a notification to the user
 * @param {string} message - The message to display
 * @param {string} type - The type of notification ('success', 'warning', 'error')
 */
function showNotification(message, type = 'info') {
  // Check if notification element already exists
  let notification = document.querySelector('.notification-toast');
  
  // Create notification if it doesn't exist
  if (!notification) {
    notification = document.createElement('div');
    notification.className = 'notification-toast';
    
    // Add icon based on type
    const icon = document.createElement('div');
    icon.className = 'notification-icon';
    
    let iconSvg = '';
    switch (type) {
      case 'success':
        iconSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" width="1.2em" height="1.2em">
          <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
        </svg>`;
        break;
      case 'warning':
        iconSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" width="1.2em" height="1.2em">
          <path d="M12 2L1 21h22L12 2z"/>
          <path d="M12 16v-5"/>
          <path d="M12 19v.01"/>
        </svg>`;
        break;
      case 'error':
        iconSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" width="1.2em" height="1.2em">
          <circle cx="12" cy="12" r="10"/>
          <path d="M15 9l-6 6M9 9l6 6"/>
        </svg>`;
        break;
      default:
        iconSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" width="1.2em" height="1.2em">
          <circle cx="12" cy="12" r="10"/>
          <path d="M12 8v4M12 16h.01"/>
        </svg>`;
    }
    
    icon.innerHTML = iconSvg;
    
    // Add message element
    const messageEl = document.createElement('span');
    messageEl.className = 'notification-message';
    
    // Assemble notification
    notification.appendChild(icon);
    notification.appendChild(messageEl);
    
    // Add to DOM
    document.body.appendChild(notification);
  }
  
  // Update message
  const messageEl = notification.querySelector('.notification-message');
  if (messageEl) {
    messageEl.textContent = message;
  }
  
  // Update class based on type
  notification.className = 'notification-toast';
  notification.classList.add(`notification-${type}`);
  
  // Show notification
  notification.classList.add('visible');
  
  // Hide after a delay
  setTimeout(() => {
    notification.classList.remove('visible');
  }, 3000);
}

/**
 * Sets up a keyboard shortcut (Ctrl+S) to replace "General" with "Sampling" when text is selected
 */
function setupGeneralSamplingReplacement() {
  document.addEventListener('keydown', function(e) {
    // Check if Ctrl+S is pressed
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
      e.preventDefault(); // Prevent browser's save function
      
      // Get the current selection
      const selection = window.getSelection();
      if (!selection.rangeCount) return;
      
      const range = selection.getRangeAt(0);
      const selectedText = selection.toString().trim();
      
      // Check if selected text contains "General"
      if (selectedText && selectedText.includes("General")) {
        // Create a new text node with "Sampling" replacing "General"
        const modifiedText = selectedText.replace(/General/g, "Sampling ");
        
        // Delete the selected content and insert the new content
        range.deleteContents();
        range.insertNode(document.createTextNode(modifiedText));
        
        // Show success notification
        showNotification("Replaced 'General' with 'Sampling'", "success");
        
        // Save state after replacement
        if (window.pageManager) {
          window.pageManager.saveCurrentPageState();
          window.pageManager.saveState();
        }
      }
    }
  });
}

// Set up General to Sampling replacement shortcut
setupGeneralSamplingReplacement();

  /**
   * Shows the hyperlink popup dialog for the current text selection
   * @param {Selection} selection - The current text selection
   * @param {Range} range - The selected range
   */
  function showHyperlinkPopup(selection, range) {
    // Create popup container
    const popup = document.createElement('div');
    popup.className = 'hyperlink-popup';
    
    // Create popup content
    popup.innerHTML = `
      <div class="hyperlink-popup-title">Create Page Link</div>
      <div class="hyperlink-input-container">
        <label class="hyperlink-input-label">Target page number:</label>
        <input type="number" class="hyperlink-page-input" min="1" value="1">
      </div>
      <div class="hyperlink-popup-buttons">
        <button class="hyperlink-popup-button hyperlink-cancel-button">Cancel</button>
        <button class="hyperlink-popup-button hyperlink-confirm-button">Confirm</button>
      </div>
    `;
    
    // Position the popup near the selection
    const rect = range.getBoundingClientRect();
    popup.style.top = `${rect.bottom + window.scrollY + 10}px`;
    popup.style.left = `${rect.left + window.scrollX}px`;
    
    // Add to document
    document.body.appendChild(popup);
    
    // Focus the input
    const pageInput = popup.querySelector('.hyperlink-page-input');
    pageInput.focus();
    
    // Close popup if clicked outside
    document.addEventListener('click', function closePopup(e) {
      if (!popup.contains(e.target)) {
        popup.remove();
        document.removeEventListener('click', closePopup);
      }
    });
    
    // Handle cancel button
    const cancelButton = popup.querySelector('.hyperlink-cancel-button');
    cancelButton.addEventListener('click', function() {
      popup.remove();
    });
    
    // Handle confirm button
    const confirmButton = popup.querySelector('.hyperlink-confirm-button');
    confirmButton.addEventListener('click', function() {
      const targetPage = parseInt(pageInput.value);
      if (!isNaN(targetPage) && targetPage >= 1) {
        createPageLink(range, targetPage);
        popup.remove();
        
        // Show confirmation notification
        showNotification(`Hyperlink to page ${targetPage} created`, 'success');
        
        // Save state after creating the hyperlink
        if (window.pageManager) {
          window.pageManager.saveCurrentPageState();
          window.pageManager.saveState();
        }
      } else {
        // Show error notification
        showNotification('Please enter a valid page number', 'error');
      }
    });
    
    // Also handle Enter key
    pageInput.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        confirmButton.click();
      }
    });
  }
  
  /**
   * Creates a page link from the selected text
   * @param {Range} range - The range of selected text
   * @param {number} targetPage - The target page number
   */
  function createPageLink(range, targetPage) {
    // Create an actual anchor element instead of span for better PDF support
    const link = document.createElement('a');
    link.className = 'page-link';
    link.setAttribute('data-target-page', targetPage);
    // Set href attribute for better PDF compatibility
    link.href = '#page=' + targetPage;
    
    // Extract the selected content
    const fragment = range.extractContents();
    link.appendChild(fragment);
    
    // Insert the link at the selection
    range.insertNode(link);
    
    // Clear the selection
    window.getSelection().removeAllRanges();
    
    // Add click event to the newly created link
    attachHyperlinkClickHandler(link);
  }
  
  /**
   * Attaches click handler to page links for navigation
   * @param {HTMLElement} link - The link element
   */
  function attachHyperlinkClickHandler(link) {
    link.addEventListener('click', function(e) {
      // Only prevent default in the app, not in PDF
      if (window.pageManager) {
        e.preventDefault();
        e.stopPropagation();
        
        // Get target page
        const targetPage = parseInt(this.getAttribute('data-target-page'));
        
        // Jump to the page if valid
        if (!isNaN(targetPage) && targetPage >= 1) {
          window.pageManager.jumpToPage(targetPage);
        }
      }
    });
  }
  
  // Function to attach click handlers to all hyperlinks in the document
  function attachAllHyperlinkHandlers() {
    const pageLinks = document.querySelectorAll('.page-link');
    pageLinks.forEach(link => {
      attachHyperlinkClickHandler(link);
    });
  }
  
  // Attach handlers to existing hyperlinks
  setTimeout(attachAllHyperlinkHandlers, 1000);
  
  // Add mutation observer to handle dynamically added hyperlinks
  const hyperlinkObserver = new MutationObserver(mutations => {
    mutations.forEach(mutation => {
      // Check for added nodes that might contain hyperlinks
      if (mutation.addedNodes.length) {
        mutation.addedNodes.forEach(node => {
          // If it's an element node
          if (node.nodeType === 1) {
            // Look for .page-link elements
            const links = node.querySelectorAll ? node.querySelectorAll('.page-link') : [];
            links.forEach(link => {
              attachHyperlinkClickHandler(link);
            });
          }
        });
      }
    });
  });
  
  // Start observing the document for hyperlinks
  hyperlinkObserver.observe(document.body, {
    childList: true,
    subtree: true
  });
});

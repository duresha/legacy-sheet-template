// main.js
// Import libraries for PDF generation with download animation
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { saveAs } from 'file-saver';

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
});

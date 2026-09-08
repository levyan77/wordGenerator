import { jsPDF } from "jspdf";

const fileToBase64 = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve(reader.result);
    reader.onerror = error => reject(error);
  });
};

const getImageDimensions = (file) => {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.width, height: img.height });
    img.onerror = reject;
    img.src = URL.createObjectURL(file);
  });
};

export const generatePdf = async (folders, config) => {
  const { globalLayout, docTitle, headerText, footerText } = config;
  const doc = new jsPDF();
  
  let yOffset = 20;
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 20;

  const checkPageBreak = (heightNeeded) => {
    if (yOffset + heightNeeded > pageHeight - margin) {
      doc.addPage();
      yOffset = 20;
      addHeaderFooter();
    }
  };

  const addHeaderFooter = () => {
    if (headerText) {
      doc.setFontSize(10);
      doc.setTextColor(150);
      doc.text(headerText, pageWidth - margin, 10, { align: "right" });
    }
    if (footerText) {
      doc.setFontSize(10);
      doc.setTextColor(150);
      doc.text(footerText, pageWidth / 2, pageHeight - 10, { align: "center" });
    }
    doc.setTextColor(0); // Reset
  };

  addHeaderFooter();

  if (docTitle) {
    doc.setFontSize(22);
    doc.text(docTitle, pageWidth / 2, yOffset, { align: "center" });
    yOffset += 15;
  }

  for (const folder of folders) {
    checkPageBreak(15);
    doc.setFontSize(16);
    doc.setFont(undefined, 'bold');
    doc.text(folder.name, margin, yOffset);
    yOffset += 8;

    if (folder.note) {
      checkPageBreak(15);
      doc.setFontSize(12);
      doc.setFont(undefined, 'italic');
      doc.setTextColor(100);
      
      const splitNote = doc.splitTextToSize(folder.note, pageWidth - margin * 2);
      doc.text(splitNote, margin, yOffset);
      yOffset += (splitNote.length * 6) + 5;
      doc.setTextColor(0);
    }

    doc.setFont(undefined, 'normal');

    for (const file of folder.files) {
      const base64Img = await fileToBase64(file);
      const { width, height } = await getImageDimensions(file);
      
      const MAX_WIDTH = globalLayout === 'Two Columns' ? (pageWidth - margin * 2) / 2 - 5 : pageWidth - margin * 2;
      let finalWidth = width;
      let finalHeight = height;

      if (width > MAX_WIDTH) {
        const ratio = MAX_WIDTH / width;
        finalWidth = MAX_WIDTH;
        finalHeight = height * ratio;
      }

      checkPageBreak(finalHeight + 10);
      
      // Calculate X to center if single column
      const xPos = globalLayout === 'Two Columns' ? margin : (pageWidth - finalWidth) / 2;
      
      // Determine format
      const isPng = file.type === 'image/png';
      const format = isPng ? 'PNG' : 'JPEG';

      doc.addImage(base64Img, format, xPos, yOffset, finalWidth, finalHeight);
      yOffset += finalHeight + 5; // Small gap for text

      // Print filename
      doc.setFontSize(9);
      doc.setTextColor(100);
      const textWidth = doc.getTextWidth(file.name);
      const textX = xPos + (finalWidth - textWidth) / 2; // Center under image
      doc.text(file.name, textX, yOffset);
      
      yOffset += 15; // Gap before next image or folder
      doc.setTextColor(0); // Reset color
    }
    
    yOffset += 10; // Space between folders
  }

  doc.save(docTitle ? `${docTitle}.pdf` : "Generated_Images_Document.pdf");
};

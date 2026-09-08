import { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, HeadingLevel, Header, Footer } from "docx";
import { saveAs } from "file-saver";

const fileToArrayBuffer = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

const getImageDimensions = (file) => {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      resolve({ width: img.width, height: img.height });
    };
    img.onerror = reject;
    img.src = URL.createObjectURL(file);
  });
};

export const generateDocx = async (folders, config) => {
  const { globalLayout, docTitle, headerText, footerText } = config;
  
  // Create Header
  const header = headerText ? new Header({
    children: [
      new Paragraph({
        children: [new TextRun({ text: headerText, color: "666666" })],
        alignment: AlignmentType.RIGHT,
      }),
    ],
  }) : undefined;

  // Create Footer
  const footer = footerText ? new Footer({
    children: [
      new Paragraph({
        children: [new TextRun({ text: footerText, color: "666666" })],
        alignment: AlignmentType.CENTER,
      }),
    ],
  }) : undefined;

  let children = [];

  // Add Cover/Title if provided
  if (docTitle) {
    children.push(
      new Paragraph({
        text: docTitle,
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { after: 400, before: 400 },
      })
    );
  }

  for (const folder of folders) {
    // Add Folder Name as Heading
    children.push(
      new Paragraph({
        text: folder.name,
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 400, after: 200 },
      })
    );

    // Add Note if exists
    if (folder.note) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: folder.note,
              italics: true,
            }),
          ],
          spacing: { after: 200 },
        })
      );
    }

    // Process Images
    for (const file of folder.files) {
      const arrayBuffer = await fileToArrayBuffer(file);
      const { width, height } = await getImageDimensions(file);
      
      const MAX_WIDTH = globalLayout === 'Two Columns' ? 300 : 500;
      let finalWidth = width;
      let finalHeight = height;

      if (width > MAX_WIDTH) {
        const ratio = MAX_WIDTH / width;
        finalWidth = MAX_WIDTH;
        finalHeight = height * ratio;
      }

      children.push(
        new Paragraph({
          children: [
            new ImageRun({
              data: arrayBuffer,
              transformation: {
                width: finalWidth,
                height: finalHeight,
              },
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
        })
      );
    }
  }

  const doc = new Document({
    sections: [
      {
        headers: { default: header },
        footers: { default: footer },
        children: children,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, docTitle ? `${docTitle}.docx` : "Generated_Images_Document.docx");
};

import React, { useState } from 'react';
import { Document, Packer, Paragraph, ImageRun, TextRun } from 'docx';
import { saveAs } from 'file-saver';
import { FolderUp, FileDown, Settings, Trash2, HelpCircle, Download, Monitor } from 'lucide-react';
import './App.css';

function App() {
  const [folders, setFolders] = useState([]);
  const [globalLayout, setGlobalLayout] = useState('Single Column');
  const [isGenerating, setIsGenerating] = useState(false);
  const [showTutorial, setShowTutorial] = useState(false);

  // Handle folder selection
  const handleFolderSelect = async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    // Group files by their parent folder path
    const folderGroups = {};
    for (const file of files) {
      if (!file.type.startsWith('image/')) continue;
      
      const folderPath = file.webkitRelativePath.split('/')[0];
      if (!folderGroups[folderPath]) {
        folderGroups[folderPath] = [];
      }
      folderGroups[folderPath].push(file);
    }

    const newFolders = Object.entries(folderGroups).map(([name, files]) => ({
      id: Date.now() + Math.random(),
      name,
      files,
      note: ''
    }));

    setFolders(prev => [...prev, ...newFolders]);
  };

  const removeFolder = (id) => {
    setFolders(folders.filter(f => f.id !== id));
  };

  const updateNote = (id, note) => {
    setFolders(folders.map(f => f.id === id ? { ...f, note } : f));
  };

  const fileToArrayBuffer = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  const getImageDimensions = (file) => {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => resolve({ width: img.width, height: img.height });
      img.src = URL.createObjectURL(file);
    });
  };

  const generateDocument = async () => {
    if (folders.length === 0) return alert('Please add at least one folder of images.');
    setIsGenerating(true);

    try {
      const doc = new Document({
        sections: [],
      });

      for (const folder of folders) {
        const children = [];

        // Add Folder Title
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: \Folder: \\,
                bold: true,
                size: 32,
              }),
            ],
            spacing: { after: 200 },
          })
        );

        // Add Note if exists
        if (folder.note) {
          children.push(
            new Paragraph({
              children: [
                new TextRun({ text: "Note: ", bold: true }),
                new TextRun({ text: folder.note }),
              ],
              spacing: { after: 400 },
            })
          );
        }

        // Process Images
        for (const file of folder.files) {
          const arrayBuffer = await fileToArrayBuffer(file);
          const { width, height } = await getImageDimensions(file);
          
          // Calculate max width (e.g. 600px for docx)
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
              spacing: { after: 200 },
            })
          );
        }

        doc.addSection({ children });
      }

      const blob = await Packer.toBlob(doc);
      saveAs(blob, "Generated_Images_Document.docx");
    } catch (error) {
      console.error(error);
      alert('Error generating document: ' + error.message);
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 p-8 font-sans">
      <header className="max-w-4xl mx-auto bg-white p-6 rounded-lg shadow-sm mb-6 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div className="flex flex-col gap-2">
          <div className="flex items-center gap-3 text-blue-600">
            <Settings size={28} />
            <h1 className="text-2xl font-bold text-gray-800">Image to Word Generator (Web Edition)</h1>
          </div>
          <p className="text-gray-500">Select folders of images to compile them into an organized Word Document (.docx)</p>
        </div>
        
        <div className="flex flex-col sm:flex-row gap-2">
          <button 
            onClick={() => setShowTutorial(!showTutorial)}
            className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md transition cursor-pointer"
          >
            <HelpCircle size={16} /> How to Use
          </button>
          <a 
            href="/sample-images.zip" 
            download
            className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-md transition"
          >
            <Download size={16} /> Sample Files
          </a>
          <a 
            href="https://github.com/levyan77/wordGenerator" 
            target="_blank"
            rel="noreferrer"
            className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-purple-700 bg-purple-50 hover:bg-purple-100 rounded-md transition"
          >
            <Monitor size={16} /> Desktop App
          </a>
        </div>
      </header>

      {showTutorial && (
        <div className="max-w-4xl mx-auto bg-blue-50 border border-blue-200 p-6 rounded-lg shadow-sm mb-6 text-blue-900">
          <h2 className="font-bold text-lg mb-3 flex items-center gap-2"><HelpCircle size={20}/> Quick Tutorial</h2>
          <ol className="list-decimal list-inside space-y-2">
            <li><strong>Download Sample Files:</strong> Click the <span className="font-semibold text-blue-700">Sample Files</span> button above and extract the ZIP file to your computer.</li>
            <li><strong>Select Image Folders:</strong> Click the <span className="font-semibold text-blue-600">Select Image Folders</span> button below. Browse to the extracted sample folder (or your own folder) and select it. <em>Note: The browser will ask for permission to view files; click Allow.</em></li>
            <li><strong>Organize & Annotate:</strong> Your selected folders will appear on the right. You can type notes or descriptions in the text box below each folder's preview.</li>
            <li><strong>Choose Layout:</strong> Select "Single Column" (large images) or "Two Columns" (smaller images) from the layout options.</li>
            <li><strong>Generate:</strong> Click <span className="font-semibold text-green-700">Generate Word Doc</span>. The app will compile all images and notes into a neat <code>.docx</code> file and download it automatically!</li>
          </ol>
        </div>
      )}

      <main className="max-w-4xl mx-auto flex flex-col md:flex-row gap-6">
        <div className="w-full md:w-1/3 bg-white p-6 rounded-lg shadow-sm flex flex-col gap-6 h-fit">
          <label className="flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white p-3 rounded-lg cursor-pointer transition font-medium">
            <FolderUp size={20} /> Select Image Folders
            <input 
              type="file" 
              webkitdirectory="true" 
              directory="true" 
              multiple 
              onChange={handleFolderSelect} 
              style={{ display: 'none' }} 
            />
          </label>

          <div className="flex flex-col gap-2">
            <label className="font-semibold text-gray-700">Layout Option:</label>
            <select 
              className="p-2 border rounded-md"
              value={globalLayout} 
              onChange={(e) => setGlobalLayout(e.target.value)}
            >
              <option>Single Column</option>
              <option>Two Columns</option>
            </select>
          </div>

          <button 
            className="flex items-center justify-center gap-2 bg-green-600 hover:bg-green-700 text-white p-3 rounded-lg transition font-medium cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
            onClick={generateDocument} 
            disabled={isGenerating || folders.length === 0}
          >
            <FileDown size={20} /> {isGenerating ? 'Generating...' : 'Generate Word Doc'}
          </button>
        </div>

        <div className="w-full md:w-2/3 flex flex-col gap-4">
          {folders.length === 0 ? (
            <div className="bg-white p-12 rounded-lg shadow-sm text-center flex flex-col items-center gap-4 text-gray-400">
              <FolderUp size={48} />
              <p>No folders selected. Click the button to select folders containing images.</p>
            </div>
          ) : (
            folders.map(folder => (
              <div key={folder.id} className="bg-white p-4 rounded-lg shadow-sm border border-gray-100 flex flex-col gap-4">
                <div className="flex justify-between items-center border-b pb-2">
                  <h3 className="font-bold text-gray-700">{folder.name} <span className="text-sm font-normal text-gray-500">({folder.files.length} images)</span></h3>
                  <button className="text-red-500 hover:text-red-700 transition cursor-pointer" onClick={() => removeFolder(folder.id)}>
                    <Trash2 size={18} />
                  </button>
                </div>
                <div className="flex gap-2 overflow-x-auto py-2">
                  {folder.files.slice(0, 4).map((f, i) => (
                    <img key={i} src={URL.createObjectURL(f)} alt="preview" className="h-16 w-16 object-cover rounded-md border" />
                  ))}
                  {folder.files.length > 4 && <div className="h-16 w-16 bg-gray-100 rounded-md border flex items-center justify-center text-sm text-gray-500">+{folder.files.length - 4}</div>}
                </div>
                <textarea 
                  className="w-full p-2 border rounded-md text-sm"
                  placeholder="Add an optional note or description for this folder..."
                  value={folder.note}
                  onChange={(e) => updateNote(folder.id, e.target.value)}
                  rows={2}
                />
              </div>
            ))
          )}
        </div>
      </main>
    </div>
  );
}

export default App;

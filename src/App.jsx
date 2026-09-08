import React, { useState, useEffect } from 'react';
import { Settings, HelpCircle, Download, Monitor, FolderUp, FileDown, FileText, CheckCircle2 } from 'lucide-react';
import { DndContext, closestCenter, KeyboardSensor, PointerSensor, useSensor, useSensors } from '@dnd-kit/core';
import { arrayMove, SortableContext, sortableKeyboardCoordinates, verticalListSortingStrategy } from '@dnd-kit/sortable';

import SortableFolder from './components/SortableFolder';
import ImageEditorModal from './components/ImageEditorModal';
import { generateDocx } from './utils/docxExport';
import { generatePdf } from './utils/pdfExport';
import { saveDraft, loadDraft, clearDraft } from './utils/storage';

function App() {
  const [folders, setFolders] = useState([]);
  const [config, setConfig] = useState({
    globalLayout: 'Two Columns',
    docTitle: '',
    headerText: '',
    footerText: '',
  });
  
  const [showTutorial, setShowTutorial] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [editingImage, setEditingImage] = useState(null); // { folderId, fileIndex, file }
  const [saveStatus, setSaveStatus] = useState('');

  // Load Draft
  useEffect(() => {
    loadDraft().then((data) => {
      if (data.folders.length > 0) {
        setFolders(data.folders);
      }
      if (data.config && Object.keys(data.config).length > 0) {
        setConfig(data.config);
      }
    });
  }, []);

  // Save Draft automatically
  useEffect(() => {
    if (folders.length > 0) {
      saveDraft(folders, config);
      setSaveStatus('Draft saved automatically.');
      const t = setTimeout(() => setSaveStatus(''), 3000);
      return () => clearTimeout(t);
    } else {
      clearDraft();
    }
  }, [folders, config]);

  const updateConfig = (key, value) => {
    setConfig(prev => ({ ...prev, [key]: value }));
  };

  const handleFolderSelect = (event) => {
    const files = Array.from(event.target.files);
    if (files.length === 0) return;

    const grouped = files.reduce((acc, file) => {
      if (!file.type.startsWith('image/')) return acc;
      const parts = file.webkitRelativePath.split('/');
      const folderName = parts.length > 1 ? parts[0] : 'Root Images';
      if (!acc[folderName]) acc[folderName] = [];
      acc[folderName].push(file);
      return acc;
    }, {});

    const newFolders = Object.entries(grouped).map(([name, folderFiles]) => ({
      id: crypto.randomUUID(),
      name,
      files: folderFiles,
      note: ''
    }));

    setFolders(prev => [...prev, ...newFolders]);
    event.target.value = null;
  };

  const removeFolder = (id) => {
    setFolders(prev => prev.filter(f => f.id !== id));
  };

  const updateNote = (id, text) => {
    setFolders(prev => prev.map(f => f.id === id ? { ...f, note: text } : f));
  };

  const clearAll = () => {
    if(confirm('Are you sure you want to clear all folders and drafts?')) {
      setFolders([]);
      clearDraft();
    }
  };

  // --- DnD Handlers ---
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 5 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates })
  );

  const handleDragEnd = (event) => {
    const { active, over } = event;
    if (active.id !== over.id) {
      setFolders((items) => {
        const oldIndex = items.findIndex(i => i.id === active.id);
        const newIndex = items.findIndex(i => i.id === over.id);
        return arrayMove(items, oldIndex, newIndex);
      });
    }
  };

  // --- Image Editor Handlers ---
  const openImageEditor = (folderId, fileIndex) => {
    const folder = folders.find(f => f.id === folderId);
    setEditingImage({ folderId, fileIndex, file: folder.files[fileIndex] });
  };

  const saveEditedImage = (newFile) => {
    setFolders(prev => prev.map(f => {
      if (f.id === editingImage.folderId) {
        const newFiles = [...f.files];
        newFiles[editingImage.fileIndex] = newFile;
        return { ...f, files: newFiles };
      }
      return f;
    }));
    setEditingImage(null);
  };

  // --- Export Handlers ---
  const handleExport = async (type) => {
    setIsGenerating(true);
    try {
      if (type === 'docx') {
        await generateDocx(folders, config);
      } else {
        await generatePdf(folders, config);
      }
    } catch (error) {
      console.error(error);
      alert(`Error generating ${type.toUpperCase()}: ` + error.message);
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 p-4 md:p-8 font-sans pb-24">
      {/* HEADER */}
      <header className="max-w-6xl mx-auto bg-white p-6 rounded-xl shadow-sm border border-gray-100 mb-6 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div className="flex flex-col gap-2">
          <div className="flex items-center gap-3 text-blue-600">
            <Settings size={32} strokeWidth={2.5} />
            <h1 className="text-2xl font-black text-gray-800 tracking-tight">DocCompiler Pro</h1>
          </div>
          <p className="text-gray-500 font-medium">Professional Image-to-Document Generator</p>
        </div>
        
        <div className="flex flex-wrap gap-2">
          <button 
            onClick={() => setShowTutorial(!showTutorial)}
            className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition cursor-pointer"
          >
            <HelpCircle size={16} /> Tutorial
          </button>
          <a 
            href="/sample-images.zip" 
            download
            className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition"
          >
            <Download size={16} /> Samples
          </a>
          <a 
            href="https://github.com/levyan77/wordGenerator" 
            target="_blank"
            rel="noreferrer"
            className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-purple-700 bg-purple-50 hover:bg-purple-100 rounded-lg transition"
          >
            <Monitor size={16} /> Desktop App
          </a>
        </div>
      </header>

      {/* TUTORIAL */}
      {showTutorial && (
        <div className="max-w-6xl mx-auto bg-blue-50 border-2 border-blue-100 p-6 rounded-xl shadow-sm mb-6 text-blue-900">
          <h2 className="font-bold text-lg mb-3 flex items-center gap-2"><HelpCircle size={20}/> Quick Tutorial</h2>
          <ol className="list-decimal list-inside space-y-2 font-medium">
            <li><strong>Import:</strong> Click <span className="text-blue-700">Select Image Folders</span>.</li>
            <li><strong>Reorder:</strong> Drag and drop folders using the grip icon to change their order.</li>
            <li><strong>Edit:</strong> Click any image to open the <span className="text-blue-700">Cropping & Rotation Tool</span>.</li>
            <li><strong>Customize:</strong> Add Cover Title, Header, and Footer text in the sidebar.</li>
            <li><strong>Export:</strong> Click <span className="text-green-700">Generate Word Doc</span> or <span className="text-red-600">Generate PDF</span>!</li>
          </ol>
        </div>
      )}

      {/* MAIN CONTENT */}
      <main className="max-w-6xl mx-auto flex flex-col md:flex-row gap-8">
        
        {/* SIDEBAR */}
        <div className="w-full md:w-80 shrink-0 flex flex-col gap-6">
          <label className="flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white p-4 rounded-xl cursor-pointer transition shadow-md font-bold text-lg">
            <FolderUp size={24} /> Import Folders
            <input 
              type="file" 
              webkitdirectory="true" 
              directory="true" 
              multiple 
              onChange={handleFolderSelect} 
              style={{ display: 'none' }} 
            />
          </label>

          <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 flex flex-col gap-4">
            <h3 className="font-bold text-gray-800 border-b pb-2">Document Settings</h3>
            
            <div className="flex flex-col gap-1">
              <label className="font-semibold text-gray-600 text-sm">Layout Strategy:</label>
              <select 
                className="p-2 border rounded-md font-medium text-gray-700 bg-gray-50 focus:ring-2 outline-none"
                value={config.globalLayout} 
                onChange={(e) => updateConfig('globalLayout', e.target.value)}
              >
                <option>Single Column (Large)</option>
                <option>Two Columns (Compact)</option>
              </select>
            </div>

            <div className="flex flex-col gap-1">
              <label className="font-semibold text-gray-600 text-sm">Cover Title:</label>
              <input 
                type="text"
                placeholder="e.g. Q3 Field Report"
                className="p-2 border rounded-md text-sm outline-none focus:ring-2"
                value={config.docTitle}
                onChange={(e) => updateConfig('docTitle', e.target.value)}
              />
            </div>

            <div className="flex flex-col gap-1">
              <label className="font-semibold text-gray-600 text-sm">Header Text:</label>
              <input 
                type="text"
                placeholder="Top right text..."
                className="p-2 border rounded-md text-sm outline-none focus:ring-2"
                value={config.headerText}
                onChange={(e) => updateConfig('headerText', e.target.value)}
              />
            </div>

            <div className="flex flex-col gap-1">
              <label className="font-semibold text-gray-600 text-sm">Footer Text:</label>
              <input 
                type="text"
                placeholder="Bottom center text..."
                className="p-2 border rounded-md text-sm outline-none focus:ring-2"
                value={config.footerText}
                onChange={(e) => updateConfig('footerText', e.target.value)}
              />
            </div>
          </div>

          <div className="flex flex-col gap-3">
             <button 
              className="flex items-center justify-center gap-2 bg-[#2b579a] hover:bg-[#1e3f70] text-white p-4 rounded-xl transition shadow-md font-bold cursor-pointer disabled:opacity-50"
              onClick={() => handleExport('docx')} 
              disabled={isGenerating || folders.length === 0}
            >
              <FileText size={20} /> {isGenerating ? 'Processing...' : 'Export as .DOCX'}
            </button>
            <button 
              className="flex items-center justify-center gap-2 bg-[#df2020] hover:bg-[#b01616] text-white p-4 rounded-xl transition shadow-md font-bold cursor-pointer disabled:opacity-50"
              onClick={() => handleExport('pdf')} 
              disabled={isGenerating || folders.length === 0}
            >
              <FileDown size={20} /> {isGenerating ? 'Processing...' : 'Export as .PDF'}
            </button>
          </div>
          
          {folders.length > 0 && (
            <button onClick={clearAll} className="text-red-500 font-bold text-sm hover:underline text-center cursor-pointer">
              Clear All Data & Drafts
            </button>
          )}

          {saveStatus && (
             <div className="flex items-center justify-center gap-2 text-green-600 text-sm font-bold animate-pulse">
               <CheckCircle2 size={16} /> {saveStatus}
             </div>
          )}
        </div>

        {/* FOLDERS LIST (DND) */}
        <div className="flex-1 flex flex-col gap-4">
          {folders.length === 0 ? (
            <div className="bg-white p-16 rounded-xl shadow-sm border border-dashed border-gray-300 text-center flex flex-col items-center gap-4 text-gray-400">
              <FolderUp size={64} className="text-gray-300" />
              <h2 className="text-xl font-bold text-gray-500">No Image Folders Yet</h2>
              <p>Import folders from the sidebar to start compiling your document.</p>
            </div>
          ) : (
            <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
              <SortableContext items={folders.map(f => f.id)} strategy={verticalListSortingStrategy}>
                {folders.map(folder => (
                  <SortableFolder 
                    key={folder.id} 
                    folder={folder} 
                    removeFolder={removeFolder} 
                    updateNote={updateNote} 
                    onEditImage={openImageEditor}
                  />
                ))}
              </SortableContext>
            </DndContext>
          )}
        </div>

      </main>

      {/* IMAGE EDITOR MODAL */}
      {editingImage && (
        <ImageEditorModal 
          file={editingImage.file} 
          onClose={() => setEditingImage(null)} 
          onSave={saveEditedImage} 
        />
      )}
    </div>
  );
}

export default App;

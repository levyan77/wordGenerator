import React from 'react';
import { useSortable } from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { Trash2, GripVertical, Edit2 } from 'lucide-react';

export default function SortableFolder({ folder, removeFolder, updateNote, onEditImage }) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging
  } = useSortable({ id: folder.id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    opacity: isDragging ? 0.5 : 1,
    zIndex: isDragging ? 10 : 1,
  };

  return (
    <div 
      ref={setNodeRef} 
      style={style} 
      className="bg-white p-4 rounded-lg shadow-sm border border-gray-200 flex flex-col gap-4 relative group"
    >
      <div className="flex justify-between items-center border-b pb-2">
        <div className="flex items-center gap-2">
          <div {...attributes} {...listeners} className="cursor-grab hover:bg-gray-100 p-1 rounded">
            <GripVertical size={20} className="text-gray-400" />
          </div>
          <h3 className="font-bold text-gray-700">{folder.name} <span className="text-sm font-normal text-gray-500">({folder.files.length} images)</span></h3>
        </div>
        <button 
          className="text-red-500 hover:text-red-700 transition cursor-pointer p-1" 
          onClick={() => removeFolder(folder.id)}
        >
          <Trash2 size={18} />
        </button>
      </div>
      
      <div className="flex gap-2 overflow-x-auto py-2">
        {folder.files.slice(0, 5).map((f, i) => (
          <div key={i} className="relative group/img cursor-pointer" onClick={() => onEditImage(folder.id, i)}>
            <img 
              src={URL.createObjectURL(f)} 
              alt="preview" 
              title={f.name}
              className="h-20 w-20 object-cover rounded-md border border-gray-300" 
            />
            <div className="absolute inset-0 bg-black/40 hidden group-hover/img:flex items-center justify-center rounded-md">
              <Edit2 size={16} className="text-white"/>
            </div>
          </div>
        ))}
        {folder.files.length > 5 && (
          <div className="h-20 w-20 bg-gray-50 rounded-md border border-gray-300 flex items-center justify-center text-sm text-gray-500 font-medium">
            +{folder.files.length - 5}
          </div>
        )}
      </div>
      
      <textarea 
        className="w-full p-3 border border-gray-200 rounded-md text-sm focus:ring-2 focus:ring-blue-500 outline-none transition"
        placeholder="Add an optional note or description for this folder..."
        value={folder.note}
        onChange={(e) => updateNote(folder.id, e.target.value)}
        rows={2}
        onPointerDown={(e) => e.stopPropagation()} // Prevent dragging when typing
      />
    </div>
  );
}

import React, { useState, useRef } from 'react';
import ReactCrop, { centerCrop, makeAspectCrop } from 'react-image-crop';
import 'react-image-crop/dist/ReactCrop.css';
import { RotateCw, Crop, Save, X } from 'lucide-react';

export default function ImageEditorModal({ file, onClose, onSave }) {
  const [imgSrc, setImgSrc] = useState(URL.createObjectURL(file));
  const [crop, setCrop] = useState();
  const [completedCrop, setCompletedCrop] = useState(null);
  const [rotation, setRotation] = useState(0);
  const imgRef = useRef(null);

  function onImageLoad(e) {
    const { width, height } = e.currentTarget;
    setCrop(centerCrop(
      makeAspectCrop({ unit: '%', width: 90 }, 1, width, height),
      width,
      height
    ));
  }

  const handleRotate = () => {
    setRotation((prev) => (prev + 90) % 360);
  };

  const handleSave = async () => {
    if (!imgRef.current) return;
    
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const image = imgRef.current;
    
    // Calculate bounding box based on rotation
    const safeArea = Math.max(image.width, image.height) * 2;
    canvas.width = safeArea;
    canvas.height = safeArea;

    ctx.translate(safeArea / 2, safeArea / 2);
    ctx.rotate((rotation * Math.PI) / 180);
    ctx.translate(-safeArea / 2, -safeArea / 2);

    ctx.drawImage(
      image,
      safeArea / 2 - image.width * 0.5,
      safeArea / 2 - image.height * 0.5
    );

    const data = ctx.getImageData(0, 0, safeArea, safeArea);
    
    // Now apply crop
    const cropX = completedCrop?.x || 0;
    const cropY = completedCrop?.y || 0;
    const cropW = completedCrop?.width || image.width;
    const cropH = completedCrop?.height || image.height;
    
    const finalCanvas = document.createElement('canvas');
    const finalCtx = finalCanvas.getContext('2d');
    
    if(rotation === 90 || rotation === 270) {
        finalCanvas.width = cropH;
        finalCanvas.height = cropW;
        // Basic implementation for rotation+crop can get complex due to coordinate shifts.
        // For simplicity in this demo, we'll draw the rotated safe area and then extract.
    } else {
        finalCanvas.width = cropW;
        finalCanvas.height = cropH;
    }

    // A simpler approach for the canvas generation:
    const tempCanvas = document.createElement('canvas');
    const tCtx = tempCanvas.getContext('2d');
    const isRotated = rotation === 90 || rotation === 270;
    tempCanvas.width = isRotated ? image.height : image.width;
    tempCanvas.height = isRotated ? image.width : image.height;
    
    tCtx.translate(tempCanvas.width/2, tempCanvas.height/2);
    tCtx.rotate((rotation * Math.PI) / 180);
    tCtx.drawImage(image, -image.width/2, -image.height/2);

    // Apply crop to the rotated canvas
    finalCanvas.width = completedCrop?.width || tempCanvas.width;
    finalCanvas.height = completedCrop?.height || tempCanvas.height;
    
    finalCtx.drawImage(
        tempCanvas,
        completedCrop?.x || 0,
        completedCrop?.y || 0,
        finalCanvas.width,
        finalCanvas.height,
        0,
        0,
        finalCanvas.width,
        finalCanvas.height
    );

    finalCanvas.toBlob((blob) => {
      if (!blob) return;
      const newFile = new File([blob], file.name, { type: file.type || 'image/jpeg' });
      onSave(newFile);
    }, file.type || 'image/jpeg');
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4">
      <div className="bg-white rounded-lg w-full max-w-3xl flex flex-col shadow-2xl">
        <div className="flex justify-between items-center p-4 border-b">
          <h2 className="text-lg font-bold text-gray-800">Edit Image</h2>
          <button onClick={onClose} className="p-1 hover:bg-gray-100 rounded-full"><X size={20}/></button>
        </div>
        
        <div className="p-4 bg-gray-100 overflow-auto flex justify-center items-center min-h-[400px]">
          <ReactCrop
            crop={crop}
            onChange={(_, percentCrop) => setCrop(percentCrop)}
            onComplete={(c) => setCompletedCrop(c)}
          >
            <img
              ref={imgRef}
              src={imgSrc}
              alt="Crop preview"
              style={{ transform: `rotate(${rotation}deg)`, maxHeight: '50vh' }}
              onLoad={onImageLoad}
            />
          </ReactCrop>
        </div>

        <div className="p-4 border-t flex justify-between items-center bg-gray-50 rounded-b-lg">
          <button 
            onClick={handleRotate}
            className="flex items-center gap-2 px-4 py-2 bg-blue-100 text-blue-700 hover:bg-blue-200 rounded-md font-medium"
          >
            <RotateCw size={18}/> Rotate 90°
          </button>
          
          <button 
            onClick={handleSave}
            className="flex items-center gap-2 px-6 py-2 bg-green-600 text-white hover:bg-green-700 rounded-md font-medium"
          >
            <Save size={18}/> Save Changes
          </button>
        </div>
      </div>
    </div>
  );
}

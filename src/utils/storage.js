import localforage from 'localforage';

// Configure localforage to use IndexedDB
localforage.config({
  name: 'WordGeneratorDrafts',
  storeName: 'draft_folders',
});

// Since File objects can sometimes lose their metadata in IndexedDB across browsers,
// we convert them to Blobs with metadata before saving, and restore them when loading.

export const saveDraft = async (folders, config) => {
  try {
    const serializedFolders = folders.map(folder => ({
      ...folder,
      files: folder.files.map(file => ({
        data: file, // localforage handles Blobs automatically in IndexedDB
        name: file.name,
        type: file.type,
        lastModified: file.lastModified
      }))
    }));
    await localforage.setItem('folders', serializedFolders);
    if(config) await localforage.setItem('config', config);
  } catch (error) {
    console.error('Error saving draft:', error);
  }
};

export const loadDraft = async () => {
  try {
    const serializedFolders = await localforage.getItem('folders');
    const config = await localforage.getItem('config');
    
    let folders = [];
    if (serializedFolders) {
      folders = serializedFolders.map(folder => ({
        ...folder,
        files: folder.files.map(f => {
          // Reconstruct File object from Blob
          return new File([f.data], f.name, {
            type: f.type,
            lastModified: f.lastModified
          });
        })
      }));
    }
    
    return { folders, config: config || {} };
  } catch (error) {
    console.error('Error loading draft:', error);
    return { folders: [], config: {} };
  }
};

export const clearDraft = async () => {
  await localforage.clear();
};

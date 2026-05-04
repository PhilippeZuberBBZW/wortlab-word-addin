function dataUrlToBase64(dataUrl: string): string {
  const marker = 'base64,';
  const index = dataUrl.indexOf(marker);
  return index >= 0 ? dataUrl.substring(index + marker.length) : dataUrl;
}

async function blobToPngBase64(blob: Blob): Promise<string> {
  if (typeof createImageBitmap === 'function') {
    try {
      const bitmap = await createImageBitmap(blob);
      const canvas = document.createElement('canvas');
      canvas.width = bitmap.width;
      canvas.height = bitmap.height;
      const context = canvas.getContext('2d');
      if (!context) {
        throw new Error('Canvas konnte nicht initialisiert werden.');
      }
      context.drawImage(bitmap, 0, 0);
      bitmap.close();
      return dataUrlToBase64(canvas.toDataURL('image/png'));
    } catch {
    }
  }

  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const result = reader.result;
      if (typeof result !== 'string') {
        reject(new Error('Bild konnte nicht verarbeitet werden.'));
        return;
      }
      resolve(dataUrlToBase64(result));
    };
    reader.onerror = () => reject(new Error('Bild konnte nicht gelesen werden.'));
    reader.readAsDataURL(blob);
  });
}

export async function insertWordText(text: string): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(text, Word.InsertLocation.replace);
    await context.sync();
  });
}

export async function insertWordImage(blob: Blob): Promise<void> {
  const base64 = await blobToPngBase64(blob);
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
    await context.sync();
  });
}

function toBase64(dataUrl: string): string {
  const marker = 'base64,';
  const index = dataUrl.indexOf(marker);
  return index >= 0 ? dataUrl.substring(index + marker.length) : dataUrl;
}

async function fetchImageAsBase64(url: string): Promise<string> {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Bild konnte nicht geladen werden: ${response.status}`);
  }

  const blob = await response.blob();
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const result = reader.result;
      if (typeof result !== 'string') {
        reject(new Error('Bild konnte nicht verarbeitet werden.'));
        return;
      }
      resolve(toBase64(result));
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

export async function insertWordImage(url: string): Promise<void> {
  const base64 = await fetchImageAsBase64(url);
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
    await context.sync();
  });
}

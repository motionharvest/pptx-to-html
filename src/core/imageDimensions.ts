/** Read pixel dimensions from common image formats embedded in PPTX media. */
export function readImageDimensions(bytes: Uint8Array): { width: number; height: number } | undefined {
  if (bytes.length < 10) return undefined;

  // PNG: IHDR at byte 16
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47 && bytes.length >= 24) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const width = view.getUint32(16, false);
    const height = view.getUint32(20, false);
    return width > 0 && height > 0 ? { width, height } : undefined;
  }

  // GIF
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes.length >= 10) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const width = view.getUint16(6, true);
    const height = view.getUint16(8, true);
    return width > 0 && height > 0 ? { width, height } : undefined;
  }

  // JPEG: scan for SOF0/SOF2 marker
  if (bytes[0] === 0xff && bytes[1] === 0xd8) {
    let i = 2;
    while (i + 8 < bytes.length) {
      if (bytes[i] !== 0xff) {
        i++;
        continue;
      }
      const marker = bytes[i + 1];
      if (marker === 0xc0 || marker === 0xc2) {
        const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        const height = view.getUint16(i + 5, false);
        const width = view.getUint16(i + 7, false);
        return width > 0 && height > 0 ? { width, height } : undefined;
      }
      if (marker === 0xd9 || marker === 0xda) break;
      const len = (bytes[i + 2] << 8) | bytes[i + 3];
      if (len < 2) break;
      i += 2 + len;
    }
  }

  return undefined;
}

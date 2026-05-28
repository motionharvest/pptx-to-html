import { ImageCrop } from "../models/SlideElement";

export interface CropLayoutInput {
  frameW: number;
  frameH: number;
  crop: ImageCrop;
  fill?: ImageCrop;
  naturalSize?: { width: number; height: number };
  /** a:picLocks @noChangeAspect — scale uniformly (cover) instead of stretching. */
  preserveAspectRatio?: boolean;
}

export interface CropLayoutResult {
  imgW: number;
  imgH: number;
  imgLeft: number;
  imgTop: number;
}

const EMPTY: ImageCrop = { left: 0, top: 0, right: 0, bottom: 0 };

/**
 * Map a:srcRect crop onto a:fillRect within the shape frame (a:stretch).
 * Follows Apache POI DrawTexturePaint: crop source sub-rectangle, then scale into fill area.
 */
export function computeCroppedImageLayout(input: CropLayoutInput): CropLayoutResult | null {
  const { frameW, frameH, crop, naturalSize, preserveAspectRatio } = input;
  const fill = input.fill ?? EMPTY;

  const visibleW = 1 - crop.left - crop.right;
  const visibleH = 1 - crop.top - crop.bottom;
  if (visibleW <= 0 || visibleH <= 0 || frameW <= 0 || frameH <= 0) return null;

  const fillLeft = fill.left * frameW;
  const fillTop = fill.top * frameH;
  const fillW = frameW * (1 - fill.left - fill.right);
  const fillH = frameH * (1 - fill.top - fill.bottom);
  if (fillW <= 0 || fillH <= 0) return null;

  if (naturalSize) {
    const cropWpx = naturalSize.width * visibleW;
    const cropHpx = naturalSize.height * visibleH;
    let scaleX = fillW / cropWpx;
    let scaleY = fillH / cropHpx;

    if (preserveAspectRatio) {
      const scale = Math.max(scaleX, scaleY);
      scaleX = scale;
      scaleY = scale;
    }

    const imgW = naturalSize.width * scaleX;
    const imgH = naturalSize.height * scaleY;

    if (preserveAspectRatio) {
      return {
        imgW,
        imgH,
        imgLeft: fillLeft + (fillW - cropWpx * scaleX) / 2 - crop.left * imgW,
        imgTop: fillTop + (fillH - cropHpx * scaleY) / 2 - crop.top * imgH,
      };
    }

    return {
      imgW,
      imgH,
      imgLeft: fillLeft - crop.left * imgW,
      imgTop: fillTop - crop.top * imgH,
    };
  }

  // Fallback when pixel dimensions unavailable: scale relative to frame.
  const imgW = fillW / visibleW;
  const imgH = fillH / visibleH;
  return {
    imgW,
    imgH,
    imgLeft: fillLeft - crop.left * imgW,
    imgTop: fillTop - crop.top * imgH,
  };
}

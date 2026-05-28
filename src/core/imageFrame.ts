import { ImageElement, Position, ShapeElement, ShapeFrameStyle, Size } from "../models/SlideElement";

function centerOf(position: Position, size: Size): { x: number; y: number } {
  return {
    x: position.x + size.width / 2,
    y: position.y + size.height / 2,
  };
}

function containsPoint(position: Position, size: Size, x: number, y: number): boolean {
  return (
    x >= position.x
    && x <= position.x + size.width
    && y >= position.y
    && y <= position.y + size.height
  );
}

function frameArea(size: Size): number {
  return size.width * size.height;
}

function shapeToFrame(shape: ShapeElement, clipImage: boolean): ShapeFrameStyle {
  return {
    shapeType: shape.shapeType,
    fillColor: shape.fillColor,
    borderColor: shape.borderColor,
    strokeWidth: shape.strokeWidth,
    position: { ...shape.position },
    size: { ...shape.size },
    shadow: shape.shadow,
    clipImage,
  };
}

/** Photo fills the circle when image bounds cover most of the backing oval. */
function shouldClipImageToFrame(image: ImageElement, frameSize: Size): boolean {
  const coverageW = image.size.width / frameSize.width;
  const coverageH = image.size.height / frameSize.height;
  return coverageW >= 0.85 && coverageH >= 0.85;
}

function isFrameCandidate(shape: ShapeElement): boolean {
  if (shape.type !== "shape") return false;
  if (shape.shapeType !== "ellipse" && shape.shapeType !== "roundRect") return false;
  return (
    shape.fillColor !== "transparent"
    || shape.borderColor !== undefined
    || shape.shadow !== undefined
  );
}

/**
 * Attach ellipse (or roundRect) backing shapes to images whose center falls inside the frame.
 * Slide 7 step icons use a separate oval shape under each picture.
 * Team headshots use a shadow oval the same size as the photo.
 */
export function attachBackingShapeFrames(
  shapes: ShapeElement[],
  images: ImageElement[],
): Set<ShapeElement> {
  const consumed = new Set<ShapeElement>();
  const frameShapes = shapes.filter(isFrameCandidate);

  if (!frameShapes.length) return consumed;

  for (const image of images) {
    if (image.frame) continue;

    const { x: cx, y: cy } = centerOf(image.position, image.size);
    let best: ShapeElement | null = null;
    let bestArea = Infinity;

    for (const shape of frameShapes) {
      if (consumed.has(shape)) continue;
      if (!containsPoint(shape.position, shape.size, cx, cy)) continue;
      const area = frameArea(shape.size);
      if (area < bestArea) {
        best = shape;
        bestArea = area;
      }
    }

    if (best) {
      if (image.frame) {
        if (!image.frame.shadow && best.shadow) {
          image.frame = { ...image.frame, shadow: best.shadow };
        }
        consumed.add(best);
        continue;
      }

      const clipImage = shouldClipImageToFrame(image, best.size);
      image.frame = shapeToFrame(best, clipImage);
      consumed.add(best);
    }
  }

  return consumed;
}

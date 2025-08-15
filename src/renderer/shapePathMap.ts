/**
 * Maps known PowerPoint shape types to SVG path definitions.
 * Supports arrows, connectors, and geometric shapes.
 * @param type The shape type from the pptx.
 * @returns SVG path string to use in <path> or <polyline>, etc.
 */
export function getSvgPathForShape(type: string): string {
  switch (type) {
    // ▸ Basic Arrows
    case "rightArrow":
    case "arrow":
      return "POLYGON 0,25 70,25 70,0 100,50 70,100 70,75 0,75";

    case "leftArrow":
      return "POLYGON 100,25 30,25 30,0 0,50 30,100 30,75 100,75";

    case "leftRightArrow":
      return "POLYGON 0,50 30,0 30,25 70,25 70,0 100,50 70,100 70,75 30,75 30,100";

    case "triangle":
      return "POLYGON 50,0 100,100 0,100";

    case "star5":
      return "POLYGON 50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35";

    case "cloud":
      return "PATH M20,60 C10,60 10,40 25,40 C30,20 50,20 55,35 C70,30 80,40 80,50 C90,50 90,70 75,70 H25 Z";

    // ▸ Straight connector (line)
    case "straightConnector1":
      return "LINE_ARROW 0,50 100,50";

    // ▸ Bent connectors
    case "bentConnector2":
      return "POLYLINE 0,50 50,50 50,100";

    case "bentConnector3":
      return "POLYLINE_ARROW 0,50 40,50 40,70 100,70";

    case "bentConnector4":
      return "POLYLINE 0,30 30,30 30,70 70,70 70,100";

    case "bentConnector5":
      return "POLYLINE 0,20 30,20 30,50 60,50 60,80 100,80";

    // ▸ Curved connectors
    case "curvedConnector2":
      return "PATH M0,50 Q50,0 100,50";

    case "curvedConnector3":
      return "PATH M0,50 Q25,0 50,50 Q75,100 100,50";

    case "curvedConnector4":
      return "PATH M0,40 Q20,0 40,40 Q60,80 80,40 Q90,20 100,40";

    case "curvedConnector5":
      return "PATH M0,50 Q20,20 40,50 Q60,80 80,50 Q90,40 100,50";

    // ▸ Notched, bent, and curved arrows
    case "bentArrow":
      return "POLYGON 0,0 70,0 70,30 100,30 50,100 50,30 0,30";

    case "notchedRightArrow":
      return "POLYGON 0,20 60,20 60,0 100,50 60,100 60,80 0,80";

    case "curvedRightArrow":
      return "PATH M0,50 Q50,0 100,50 Q50,100 0,50 Z";

    // Default fallback (rectangular shape)
    default:
      return "POLYGON 0,0 100,0 100,100 0,100";
  }
}
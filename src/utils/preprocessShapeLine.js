import JSZip from "jszip";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const V_NS = "urn:schemas-microsoft-com:vml";

const DEFAULT_COLOR = "FF0000";
const DEFAULT_SIZE_EIGHTHS = 14;

export async function preprocessShapeLines (buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/document.xml");
  if (!file) return buffer;

  const xml = await file.async("string");
  const newXml = transformDocumentXml(xml);
  if (newXml === xml) return buffer;

  zip.file("word/document.xml", newXml);
  return zip.generateAsync({ type: "arraybuffer" });
}

function transformDocumentXml (xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) return xml;

  const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
  let changed = false;

  paragraphs.forEach(function (p) {
    const shapes = findShapeLines(p);
    if (shapes.length === 0) return;

    const first = shapes[0];
    const color = readColor(first) || DEFAULT_COLOR;
    const size = readSizeEighths(first) || DEFAULT_SIZE_EIGHTHS;

    shapes.forEach(function (shapeEl) {
      removeShapeWrapper(p, shapeEl);
    });
    addBottomBorder(doc, p, color, size);
    changed = true;
  });

  if (!changed) return xml;
  return new XMLSerializer().serializeToString(doc);
}

function findShapeLines (p) {
  const out = [];

  Array.from(p.getElementsByTagNameNS(W_NS, "drawing")).forEach(function (d) {
    const prsts = d.getElementsByTagNameNS(A_NS, "prstGeom");
    for (let i = 0; i < prsts.length; i++) {
      const prst = prsts[i].getAttribute("prst");
      if (prst === "line" || prst === "straightConnector1") {
        out.push(d);
        break;
      }
    }
  });

  Array.from(p.getElementsByTagNameNS(W_NS, "pict")).forEach(function (pict) {
    if (pict.getElementsByTagNameNS(V_NS, "line").length > 0) {
      out.push(pict);
      return;
    }
    const vShapes = pict.getElementsByTagNameNS(V_NS, "shape");
    for (let i = 0; i < vShapes.length; i++) {
      const t = vShapes[i].getAttribute("type") || "";
      if (t.includes("_x0000_t32") || t.includes("_x0000_t20")) {
        out.push(pict);
        break;
      }
    }
  });

  return out;
}

function readColor (shapeEl) {
  const srgb = shapeEl.getElementsByTagNameNS(A_NS, "srgbClr");
  if (srgb.length > 0) {
    const v = srgb[0].getAttribute("val");
    if (v) return v.toUpperCase();
  }
  const vLines = shapeEl.getElementsByTagNameNS(V_NS, "line");
  if (vLines.length > 0) {
    const c = vLines[0].getAttribute("strokecolor");
    if (c && c.startsWith("#")) return c.slice(1).toUpperCase();
  }
  const vShapes = shapeEl.getElementsByTagNameNS(V_NS, "shape");
  if (vShapes.length > 0) {
    const c = vShapes[0].getAttribute("strokecolor");
    if (c && c.startsWith("#")) return c.slice(1).toUpperCase();
  }
  return undefined;
}

function readSizeEighths (shapeEl) {
  const lns = shapeEl.getElementsByTagNameNS(A_NS, "ln");
  for (let i = 0; i < lns.length; i++) {
    const w = lns[i].getAttribute("w");
    if (!w) continue;
    const emu = parseInt(w, 10);
    if (!isNaN(emu)) {
      const pt = emu / 12700;
      return Math.max(4, Math.min(96, Math.round(pt * 8)));
    }
  }
  const vLines = shapeEl.getElementsByTagNameNS(V_NS, "line");
  if (vLines.length > 0) {
    const weight = vLines[0].getAttribute("strokeweight");
    const pt = parseStrokeWeightPt(weight);
    if (pt != null) return Math.max(4, Math.min(96, Math.round(pt * 8)));
  }
  return undefined;
}

function parseStrokeWeightPt (weight) {
  if (!weight) return null;
  const m = weight.match(/^([\d.]+)\s*(pt|px)?$/i);
  if (!m) return null;
  const n = parseFloat(m[1]);
  if (isNaN(n)) return null;
  const unit = (m[2] || "pt").toLowerCase();
  return unit === "px" ? n * 0.75 : n;
}

function removeShapeWrapper (p, shapeEl) {
  let topToRemove = shapeEl;
  let walker = shapeEl.parentNode;
  while (walker && walker !== p) {
    const ns = walker.namespaceURI;
    const ln = walker.localName;
    if (ns === W_NS && ln === "r") {
      topToRemove = walker;
      break;
    }
    if (ns === MC_NS && (ln === "Choice" || ln === "Fallback" || ln === "AlternateContent")) {
      topToRemove = walker;
      walker = walker.parentNode;
      continue;
    }
    break;
  }
  if (topToRemove.parentNode) topToRemove.parentNode.removeChild(topToRemove);
}

function addBottomBorder (doc, p, color, sizeEighths) {
  let pPr = Array.from(p.children).find(
    (c) => c.namespaceURI === W_NS && c.localName === "pPr"
  );
  if (!pPr) {
    pPr = doc.createElementNS(W_NS, "w:pPr");
    p.insertBefore(pPr, p.firstChild);
  }

  const existing = Array.from(pPr.children).find(
    (c) => c.namespaceURI === W_NS && c.localName === "pBdr"
  );
  if (existing) pPr.removeChild(existing);

  const pBdr = doc.createElementNS(W_NS, "w:pBdr");
  const bottom = doc.createElementNS(W_NS, "w:bottom");
  bottom.setAttributeNS(W_NS, "w:val", "single");
  bottom.setAttributeNS(W_NS, "w:sz", String(sizeEighths));
  bottom.setAttributeNS(W_NS, "w:space", "1");
  bottom.setAttributeNS(W_NS, "w:color", color);
  pBdr.appendChild(bottom);
  pPr.appendChild(pBdr);
}
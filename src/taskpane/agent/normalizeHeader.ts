function collapseWhitespace(str: string) {
  let out = "";
  let inSpace = false;
  for (const ch of str || "") {
    const code = ch.charCodeAt(0);
    const isSpace = code === 32 || code === 9 || code === 10 || code === 13 || code === 12;
    if (isSpace) {
      if (!inSpace) {
        out += " ";
        inSpace = true;
      }
    } else {
      out += ch;
      inSpace = false;
    }
  }
  return out.trim();
}

function removeDiacritics(str: string) {
  const decomposed = (str || "").normalize("NFD");
  let out = "";
  for (const ch of decomposed) {
    const code = ch.charCodeAt(0);
    if (code < 0x0300 || code > 0x036f) out += ch;
  }
  return out;
}

export function normalizeHeader(text?: string | null) {
  return removeDiacritics(collapseWhitespace((text || "").toString().trim()).toLowerCase());
}


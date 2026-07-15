export const escXml = (s) =>
  String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

export function buildNumberRuns(numberText) {
  const t = (numberText || "").trim();
  if (t) return `<w:r><w:t xml:space="preserve">${escXml(t)}</w:t></w:r>`;
  return (
    `<w:r><w:t xml:space="preserve">(</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> SEQ equation \\* ARABIC </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>1</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
    `<w:r><w:t xml:space="preserve">)</w:t></w:r>`
  );
}

export function buildFlatOpc(ommlMath, layout, numberText) {
  const oMath = ommlMath.replace(/^\s*<\?xml[^>]*\?>\s*/i, "").trim();
  let mathBlock;
  if (layout === "display") {
    mathBlock = `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${oMath}</m:oMathPara>`;
  } else if (layout === "numbered") {
    mathBlock =
      `<w:r><w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/></w:r>` +
      oMath +
      `<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>` +
      buildNumberRuns(numberText);
  } else {
    mathBlock = oMath;
  }
  const documentXml =
    `<w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
    `<w:body><w:p>${mathBlock}</w:p></w:body>` +
    `</w:document>`;
  const relsXml =
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ` +
    `Target="word/document.xml"/>` +
    `</Relationships>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<?mso-application progid="Word.Document"?>` +
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" ` +
    `pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">` +
    `<pkg:xmlData>${relsXml}</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" ` +
    `pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData>${documentXml}</pkg:xmlData></pkg:part>` +
    `</pkg:package>`
  );
}

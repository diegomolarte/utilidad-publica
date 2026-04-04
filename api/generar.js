import Anthropic from '@anthropic-ai/sdk';
import { createHash } from 'crypto';
import { gzipSync } from 'zlib';

// Genera un .docx (ZIP con XML) sin dependencias externas
function crearDocx(campos, secciones) {
  const p = v => v || '[PENDIENTE]';

  const estilos = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
</w:styles>`;

  const numeracion = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  function parrafo(txt, opts = {}) {
    const { bold = false, size = 22, before = 0, after = 100, justify = true, bullet = false } = opts;
    const pPr = `<w:pPr>${justify && !bullet ? '<w:jc w:val="both"/>' : ''}${bullet ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : ''}<w:spacing w:before="${before}" w:after="${after}"/></w:pPr>`;
    const rPr = `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${size}"/>${bold ? '<w:b/>' : ''}</w:rPr>`;
    return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${esc(txt)}</w:t></w:r></w:p>`;
  }

  function seccion(texto) {
    if (!texto) return parrafo('[PENDIENTE]');
    let out = '';
    for (const bloque of texto.split(/\n\n+/)) {
      const lineas = bloque.split('\n').map(l => l.trim()).filter(Boolean);
      if (!lineas.length) continue;
      if (lineas.every(l => /^[•\-\*]/.test(l))) {
        lineas.forEach(l => out += parrafo(l.replace(/^[•\-\*]\s*/,''), {bullet:true,after:60}));
      } else {
        out += parrafo(lineas.join(' '), {justify:true});
      }
    }
    return out;
  }

  const hoy = campos.fecha_entrevista || new Date().toLocaleDateString('es-CO',{day:'numeric',month:'long',year:'numeric'});

  const cuerpo = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:sectPr><w:pgMar w:top="1134" w:right="1418" w:bottom="1418" w:left="1701" w:header="567" w:footer="567"/></w:sectPr>
${parrafo('REPORTE DE ENTREVISTA', {bold:true,size:24,after:40,justify:false})}
${parrafo('LEY 2292 DE 2023', {bold:true,size:24,after:40,justify:false})}
${parrafo('SERVICIOS DE UTILIDAD PÚBLICA', {bold:true,size:24,after:40,justify:false})}
${parrafo('')}
${parrafo('1.  Datos generales', {bold:true,before:200,after:100,justify:false})}
${parrafo(`Nombre completo: ${p(campos.nombre)}`, {bullet:true,after:60})}
${parrafo(`No. documento: ${p(campos.cedula)}`, {bullet:true,after:60})}
${parrafo(`Fecha de nacimiento: ${p(campos.fecha_nacimiento)}`, {bullet:true,after:60})}
${parrafo(`Lugar de reclusión: ${p(campos.lugar_reclusion)}`, {bullet:true,after:60})}
${parrafo(`Delito: ${p(campos.delito)}`, {bullet:true,after:60})}
${parrafo(`Fecha de la entrevista: ${p(campos.fecha_entrevista)}`, {bullet:true,after:60})}
${parrafo(`Nombre del entrevistador(a): ${p(campos.entrevistador)}`, {bullet:true,after:60})}
${parrafo('')}
${parrafo('2.  Condiciones asociadas a la marginalidad', {bold:true,before:200,after:100,justify:false})}
${seccion(secciones.marginalidad)}
${parrafo('')}
${parrafo('3.  Rol de jefatura de hogar', {bold:true,before:200,after:100,justify:false})}
${seccion(secciones.jefatura_hogar)}
${parrafo('')}
${parrafo('4.  Hechos que dieron lugar a la captura en flagrancia.', {bold:true,before:200,after:100,justify:false})}
${seccion(secciones.hechos_captura)}
${parrafo('')}
${parrafo('Conclusión.', {bold:true,before:200,after:100,justify:false})}
${seccion(secciones.conclusion)}
${parrafo('')}
${parrafo('')}
${parrafo('Entrevistó:')}
${parrafo('')}
${parrafo('')}
${parrafo('')}
${parrafo(p(campos.entrevistador))}
${parrafo('Equipo de atención jurídica')}
${parrafo('Dirección Nacional de Defensoría Pública')}
${parrafo(hoy)}
</w:body>
</w:document>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`;

  const packageRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  // Construir ZIP manualmente
  function strToBytes(s) { return Buffer.from(s, 'utf8'); }

  function crc32(buf) {
    let crc = 0xFFFFFFFF;
    const table = [];
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) c = (c & 1) ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      table[i] = c;
    }
    for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function zipEntry(name, data) {
    const nameBuf = Buffer.from(name, 'utf8');
    const dataBuf = Buffer.isBuffer(data) ? data : strToBytes(data);
    const compressed = gzipSync(dataBuf, { level: 6 });
    // Usar stored (no compresión) para simplificar
    const crc = crc32(dataBuf);
    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8); // stored
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(dataBuf.length, 18);
    local.writeUInt32LE(dataBuf.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    nameBuf.copy(local, 30);
    return { local: Buffer.concat([local, dataBuf]), name: nameBuf, crc, size: dataBuf.length, compressed: dataBuf.length };
  }

  const archivos = [
    ['[Content_Types].xml', contentTypes],
    ['_rels/.rels', packageRels],
    ['word/document.xml', cuerpo],
    ['word/styles.xml', estilos],
    ['word/numbering.xml', numeracion],
    ['word/_rels/document.xml.rels', rels],
  ];

  const entries = [];
  let offset = 0;
  const localParts = [];

  for (const [name, content] of archivos) {
    const entry = zipEntry(name, content);
    entries.push({ ...entry, offset });
    localParts.push(entry.local);
    offset += entry.local.length;
  }

  const centralDir = [];
  for (const e of entries) {
    const cd = Buffer.alloc(46 + e.name.length);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0, 8);
    cd.writeUInt16LE(0, 10);
    cd.writeUInt16LE(0, 12);
    cd.writeUInt16LE(0, 14);
    cd.writeUInt32LE(e.crc, 16);
    cd.writeUInt32LE(e.size, 20);
    cd.writeUInt32LE(e.compressed, 24);
    cd.writeUInt16LE(e.name.length, 28);
    cd.writeUInt16LE(0, 30);
    cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34);
    cd.writeUInt16LE(0, 36);
    cd.writeUInt32LE(0, 38);
    cd.writeUInt32LE(e.offset, 42);
    e.name.copy(cd, 46);
    centralDir.push(cd);
  }

  const cdBuf = Buffer.concat(centralDir);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cdBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localParts, cdBuf, eocd]);
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-api-key');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  const apiKey = req.headers['x-api-key'];
  if (!apiKey || !apiKey.startsWith('sk-ant-')) return res.status(401).json({ error: 'API key inválida' });

  const { campos } = req.body;
  if (!campos?.notas) return res.status(400).json({ error: 'Faltan notas del caso' });

  try {
    const client = new Anthropic({ apiKey });
    const message = await client.messages.create({
      model: 'claude-opus-4-6',
      max_tokens: 4000,
      messages: [{ role: 'user', content: `Eres asistente jurídico de la Defensoría del Pueblo de Colombia. Redacta el Reporte de Entrevista formal bajo la Ley 2292 de 2023.

DATOS:
- Nombre: ${campos.nombre || '[no proporcionado]'}
- Cédula: ${campos.cedula || '[no proporcionado]'}
- Fecha nacimiento: ${campos.fecha_nacimiento || '[no proporcionado]'}
- Lugar reclusión: ${campos.lugar_reclusion || '[no proporcionado]'}
- Delito: ${campos.delito || '[no proporcionado]'}
- Fecha entrevista: ${campos.fecha_entrevista || '[no proporcionado]'}
- Entrevistador: ${campos.entrevistador || '[no proporcionado]'}

NOTAS:
${campos.notas}

Registro formal jurídico tercera persona. Usa "manifestó que", "señaló que", "refirió que". Transforma lenguaje informal. Hijos con nombre completo y edad. Cifras en formato "quinientos mil pesos ($500.000)". Si falta dato usa [PENDIENTE].

Responde SOLO JSON sin backticks:
{"marginalidad":"3-5 párrafos separados por doble salto","jefatura_hogar":"4-6 párrafos","hechos_captura":"5-6 párrafos cronológicos","conclusion":"3-4 párrafos mencionando Ley 2292 de 2023"}` }]
    });

    const secciones = JSON.parse(message.content[0].text.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim());
    const buffer = crearDocx(campos, secciones);
    const filename = `Entrevista a usuaria - ${campos.nombre || 'usuaria'}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.setHeader('Content-Length', buffer.length);
    return res.status(200).send(buffer);

  } catch (err) {
    console.error('Error:', err.message);
    return res.status(500).json({ error: 'Error interno: ' + err.message });
  }
}

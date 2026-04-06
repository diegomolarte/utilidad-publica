import Anthropic from '@anthropic-ai/sdk';
import { gzipSync } from 'zlib';

function crearDocx(campos, secciones) {
  const p = v => v || '[PENDIENTE]';

  const estilos = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
</w:styles>`;

  const numeracion = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  function parrafo(txt, opts = {}) {
    const { bold = false, size = 24, before = 0, after = 0, justify = true, bullet = false } = opts;
    const pPr = `<w:pPr>${justify && !bullet ? '<w:jc w:val="both"/>' : ''}${bullet ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : ''}<w:spacing w:before="${before}" w:after="${after}"/></w:pPr>`;
    const rPr = `<w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:sz w:val="${size}"/>${bold ? '<w:b/>' : ''}</w:rPr>`;
    if (!txt && txt !== 0) return `<w:p>${pPr}</w:p>`;
    return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${esc(String(txt))}</w:t></w:r></w:p>`;
  }

  function seccion(texto) {
    if (!texto) return parrafo('[PENDIENTE]', {justify:true, before:120, after:120});
    let out = '';
    for (const bloque of texto.split(/\n\n+/)) {
      const lineas = bloque.split('\n').map(l => l.trim()).filter(Boolean);
      if (!lineas.length) continue;
      if (lineas.every(l => /^[•\-\*]/.test(l))) {
        lineas.forEach(l => out += parrafo(l.replace(/^[•\-\*]\s*/,''), {bullet:true, before:0, after:0}));
      } else {
        out += parrafo(lineas.join(' '), {justify:true, before:120, after:120});
      }
    }
    return out;
  }

  const hoy = campos.fecha_entrevista || new Date().toLocaleDateString('es-CO', {day:'numeric',month:'long',year:'numeric'});

  const cuerpo = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:sectPr><w:pgMar w:top="1134" w:right="1418" w:bottom="1418" w:left="1701" w:header="567" w:footer="567"/></w:sectPr>
${parrafo('REPORTE DE ENTREVISTA A FAMILIAR O TERCERO', {bold:true,size:24,after:0,justify:false})}
${parrafo('LEY 2292 DE 2023', {bold:true,size:24,after:0,justify:false})}
${parrafo('SERVICIOS DE UTILIDAD PÚBLICA', {bold:true,size:24,after:120,justify:false})}
${parrafo('1.  Datos de la usuaria', {bold:true,before:200,after:120,justify:false})}
${parrafo(`Nombre completo: ${p(campos.nombre_usuaria)}`, {bullet:true,before:0,after:0})}
${parrafo(`No. documento: ${p(campos.cedula_usuaria)}`, {bullet:true,before:0,after:0})}
${parrafo(`Fecha de nacimiento: ${p(campos.fecha_nacimiento_usuaria)}`, {bullet:true,before:0,after:0})}
${parrafo(`Lugar de reclusión: ${p(campos.lugar_reclusion)}`, {bullet:true,before:0,after:0})}
${parrafo(`Delito: ${p(campos.delito)}`, {bullet:true,before:0,after:0})}
${parrafo('2.  Datos del familiar o tercero entrevistado', {bold:true,before:200,after:120,justify:false})}
${parrafo(`Nombre completo: ${p(campos.nombre_familiar)}`, {bullet:true,before:0,after:0})}
${parrafo(`No. documento: ${p(campos.cedula_familiar)}`, {bullet:true,before:0,after:0})}
${parrafo(`Vínculo con la usuaria: ${p(campos.vinculo_familiar)}`, {bullet:true,before:0,after:0})}
${parrafo(`Fecha de la entrevista: ${p(campos.fecha_entrevista)}`, {bullet:true,before:0,after:0})}
${parrafo(`Nombre del entrevistador(a): ${p(campos.entrevistador)}`, {bullet:true,before:0,after:0})}
${parrafo('3.  Corroboración de condiciones asociadas a la marginalidad', {bold:true,before:200,after:120,justify:false})}
${seccion(secciones.marginalidad)}
${parrafo('4.  Corroboración del rol de jefatura de hogar', {bold:true,before:200,after:120,justify:false})}
${seccion(secciones.jefatura_hogar)}
${parrafo('5.  Información adicional aportada', {bold:true,before:200,after:120,justify:false})}
${seccion(secciones.informacion_adicional)}
${parrafo('Conclusión.', {bold:true,before:200,after:120,justify:false})}
${seccion(secciones.conclusion)}
${parrafo('Entrevistó:', {before:240, after:0})}
${parrafo(p(campos.entrevistador))}
${parrafo('Equipo de atención jurídica de la Dirección Nacional de Defensoría Pública')}
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

  function crc32(buf) {
    const table = [];
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) c = (c & 1) ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      table[i] = c;
    }
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function zipEntry(name, data) {
    const nameBuf = Buffer.from(name, 'utf8');
    const dataBuf = Buffer.isBuffer(data) ? data : Buffer.from(data, 'utf8');
    const crc = crc32(dataBuf);
    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6); local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10); local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(dataBuf.length, 18);
    local.writeUInt32LE(dataBuf.length, 22); local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28); nameBuf.copy(local, 30);
    return { local: Buffer.concat([local, dataBuf]), name: nameBuf, crc, size: dataBuf.length };
  }

  const archivos = [
    ['[Content_Types].xml', contentTypes], ['_rels/.rels', packageRels],
    ['word/document.xml', cuerpo], ['word/styles.xml', estilos],
    ['word/numbering.xml', numeracion], ['word/_rels/document.xml.rels', rels],
  ];

  const entries = []; let offset = 0; const localParts = [];
  for (const [name, content] of archivos) {
    const entry = zipEntry(name, content);
    entries.push({ ...entry, offset }); localParts.push(entry.local);
    offset += entry.local.length;
  }

  const centralDir = [];
  for (const e of entries) {
    const cd = Buffer.alloc(46 + e.name.length);
    cd.writeUInt32LE(0x02014b50, 0); cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0, 8); cd.writeUInt16LE(0, 10); cd.writeUInt16LE(0, 12);
    cd.writeUInt16LE(0, 14); cd.writeUInt32LE(e.crc, 16); cd.writeUInt32LE(e.size, 20);
    cd.writeUInt32LE(e.size, 24); cd.writeUInt16LE(e.name.length, 28);
    cd.writeUInt16LE(0, 30); cd.writeUInt16LE(0, 32); cd.writeUInt16LE(0, 34);
    cd.writeUInt16LE(0, 36); cd.writeUInt32LE(0, 38); cd.writeUInt32LE(e.offset, 42);
    e.name.copy(cd, 46); centralDir.push(cd);
  }

  const cdBuf = Buffer.concat(centralDir);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(0, 4); eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8); eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cdBuf.length, 12); eocd.writeUInt32LE(offset, 16); eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localParts, cdBuf, eocd]);
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-api-key');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  const apiKey = req.headers['x-api-key'] || process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(401).json({ error: 'API key no configurada' });

  const { campos } = req.body;
  if (!campos?.notas) return res.status(400).json({ error: 'Faltan notas de la entrevista' });

  try {
    const client = new Anthropic({ apiKey });
    const message = await client.messages.create({
      model: 'claude-opus-4-6',
      max_tokens: 4000,
      messages: [{ role: 'user', content: `Eres asistente jurídico de la Defensoría del Pueblo de Colombia. Redacta el Reporte de Entrevista a familiar o tercero, bajo la Ley 2292 de 2023.

Este reporte recoge el testimonio de un familiar o tercero que conoce a la usuaria y corrobora sus condiciones de marginalidad y jefatura de hogar.

DATOS DE LA USUARIA:
- Nombre: ${campos.nombre_usuaria || '[no proporcionado]'}
- Cédula: ${campos.cedula_usuaria || '[no proporcionado]'}
- Delito: ${campos.delito || '[no proporcionado]'}

DATOS DEL FAMILIAR O TERCERO ENTREVISTADO:
- Nombre: ${campos.nombre_familiar || '[no proporcionado]'}
- Cédula: ${campos.cedula_familiar || '[no proporcionado]'}
- Vínculo con la usuaria: ${campos.vinculo_familiar || '[no proporcionado]'}
- Fecha de entrevista: ${campos.fecha_entrevista || '[no proporcionado]'}
- Entrevistador(a): ${campos.entrevistador || '[no proporcionado]'}

NOTAS DE LA ENTREVISTA:
${campos.notas}

INSTRUCCIONES:
- Registro formal jurídico en tercera persona.
- El entrevistado/a se refiere a la usuaria siempre por su nombre completo o como "la señora [apellido]".
- Para lo que dijo el entrevistado: "manifestó que", "señaló que", "refirió que", "indicó que", "corroboró que", "confirmó que", "añadió que".
- Transforma el lenguaje informal a formal jurídico.
- Los hijos siempre con nombre completo y edad entre paréntesis.
- Cifras en formato "quinientos mil pesos ($500.000)".
- Si falta un dato, usa [PENDIENTE] y nada más.
- Cuando menciones la ley, usa SIEMPRE "Ley 2292 de 2023" sin citar su nombre completo.
- Cuando te refieras a quien realizó la entrevista, usa siempre una expresión como "la entrevista fue realizada por [nombre], del equipo de atención jurídica de la Dirección Nacional de Defensoría Pública". NUNCA uses "defensor público" ni "defensora pública" al referirte al entrevistador.
- NUNCA incluyas párrafos de recomendaciones, sugerencias ni notas del tipo "se recomienda", "se sugiere", "los datos señalados como [PENDIENTE] quedan sujetos a". Nada al final del documento salvo la firma.

Responde SOLO JSON sin backticks:
{"marginalidad":"3-5 párrafos sobre lo que el familiar/tercero corroboró respecto a las condiciones de marginalidad de la usuaria, separados por doble salto","jefatura_hogar":"3-5 párrafos sobre lo que corroboró del rol de jefatura de hogar","informacion_adicional":"2-4 párrafos con información adicional relevante que aportó el entrevistado, o [PENDIENTE] si no hay nada adicional relevante","conclusion":"2-3 párrafos de conclusión mencionando Ley 2292 de 2023 y el valor del testimonio para sustentar la solicitud"}` }]
    });

    const secciones = JSON.parse(message.content[0].text.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim());
    const buffer = crearDocx(campos, secciones);
    const nombre = campos.nombre_familiar || 'familiar';
    const filename = `Entrevista a familiar - ${nombre}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.setHeader('Content-Length', buffer.length);
    return res.status(200).send(buffer);

  } catch (err) {
    console.error('Error:', err.message);
    return res.status(500).json({ error: 'Error interno: ' + err.message });
  }
}

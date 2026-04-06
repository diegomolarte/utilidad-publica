import Anthropic from '@anthropic-ai/sdk';
import { createHash } from 'crypto';
import { gzipSync } from 'zlib';

// ── ZIP/DOCX builder (igual que generar.js) ─────────────────
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



function crearDocx(secciones) {
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

  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const Pt = n => n * 20;
  const Cm = n => Math.round(n * 567);

  function p(txt, opts = {}) {
    const { bold=false, size=24, before=0, after=0, justify=true, bullet=false, center=false } = opts;
    const align = center ? '<w:jc w:val="center"/>' : (justify && !bullet ? '<w:jc w:val="both"/>' : '');
    const pPr = `<w:pPr>${align}${bullet ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : ''}<w:spacing w:before="${before}" w:after="${after}"/></w:pPr>`;
    const rPr = `<w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:sz w:val="${size}"/>${bold ? '<w:b/>' : ''}</w:rPr>`;
    if (!txt && txt !== 0) return `<w:p>${pPr}</w:p>`;
    return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${esc(String(txt))}</w:t></w:r></w:p>`;
  }

  function seccionTexto(texto) {
    if (!texto) return p('[PENDIENTE]');
    let out = '';
    for (const bloque of texto.split(/\n\n+/)) {
      const lineas = bloque.split('\n').map(l => l.trim()).filter(Boolean);
      if (!lineas.length) continue;
      if (lineas.every(l => /^[•\-\*]/.test(l))) {
        lineas.forEach(l => out += p(l.replace(/^[•\-\*]\s*/,''), {bullet:true, after:60}));
      } else {
        out += p(lineas.join(' '), {before:120, after:120});
      }
    }
    return out;
  }

  const cuerpo = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:sectPr><w:pgMar w:top="1134" w:right="1418" w:bottom="1418" w:left="1701" w:header="567" w:footer="567"/></w:sectPr>
${p(secciones.encabezado_ciudad_fecha || '', {justify:false})}
${p(secciones.encabezado_cargo_juez || '', {justify:false})}
${p(secciones.encabezado_nombre_juez || '', {justify:false, bold:true})}
${p(secciones.encabezado_juzgado || '', {justify:false})}
${p(secciones.encabezado_ciudad || '', {justify:false})}
${p('E.S.D.', {justify:false})}
${p('Honorable Señor(a) Juez:', {justify:false})}
${seccionTexto(secciones.parrafo_intro)}
${p('I. CONTEXTO DE JEFATURA DE HOGAR Y MARGINALIDAD', {bold:true, before:200, after:120, justify:false})}
${seccionTexto(secciones.seccion1_contexto)}
${p('II. FUNDAMENTOS JURÍDICOS', {bold:true, before:200, after:120, justify:false})}
${seccionTexto(TEXTO_FIJO_SECCION2.replace('[NOMBRE COMPLETO DE LA CONDENADA]', secciones.nombre_condenada || '[PENDIENTE]'))}
${seccionTexto(secciones.seccion2_requisitos)}
${p('III. PROPUESTA FRENTE AL PLAN DE SERVICIOS DE UTILIDAD PÚBLICA', {bold:true, before:200, after:120, justify:false})}
${seccionTexto(secciones.seccion3_plaza)}
${p('IV. PETICIÓN', {bold:true, before:200, after:120, justify:false})}
${seccionTexto(secciones.seccion4_peticion)}
${p('ANEXOS', {bold:true, before:200, after:120, justify:false})}
${p('Se aportan como pruebas, entre otros:')}
${seccionTexto(secciones.lista_anexos)}
${p('NOTIFICACIONES', {bold:true, before:200, after:120, justify:false})}
${seccionTexto(secciones.notificaciones)}
${p('Cordialmente,')}
${p(secciones.firma_nombre || '[NOMBRE DEFENSORA]', {bold:true, justify:false})}
${p('Defensora Pública', {justify:false})}
${p(secciones.firma_tp ? 'T.P. No. ' + secciones.firma_tp : 'T.P. No. [PENDIENTE]', {justify:false})}
${p('Defensoría del Pueblo', {justify:false})}
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

// ── Texto fijo obligatorio — arranque Sección 2 ────────────
const TEXTO_FIJO_SECCION2 = `FUNDAMENTOS JURÍDICOS LEY 2292 DE 2023 QUE SE CUMPLEN EN EL CASO CONCRETO

La Ley 2292 del 8 de marzo de 2023 surge como una respuesta del Estado frente a una realidad ampliamente documentada en el sistema penitenciario: un número significativo de mujeres privadas de la libertad son madres que han debido asumir solas la responsabilidad de sus hogares y que terminaron involucradas en conductas delictivas en contextos de precariedad económica, exclusión social y ausencia de oportunidades reales. Por ello, esta ley incorpora acciones afirmativas en materia de política criminal y penitenciaria para mujeres cabeza de familia, reconociendo que la respuesta punitiva tradicional (basada exclusivamente en el encarcelamiento) puede generar consecuencias desproporcionadas no solo para la mujer condenada, sino también para los hijos menores que dependen de ella.

En este mismo sentido, la Corte Constitucional, mediante la Sentencia C-256 de 2022, advirtió que los Estados deben implementar medidas alternativas al encarcelamiento cuando se trate de mujeres que, debido a sus condiciones de vida, se ven involucradas en delitos asociados, entre otros, al narcotráfico. Esta consideración no constituye una simple afirmación retórica. Por el contrario, representa un llamado a que la política criminal tenga en cuenta los datos empíricos y las realidades sociales que rodean la vida de muchas mujeres procesadas penalmente.

En desarrollo de esta orientación, la Ley 2292 incorporó al ordenamiento jurídico la pena sustitutiva de servicios de utilidad pública para mujeres cabeza de familia, consistente en la prestación de servicios no remunerados, en libertad, a favor de instituciones públicas u organizaciones sin ánimo de lucro en el lugar de domicilio de la condenada. Se trata de una medida que permite el cumplimiento de la sanción penal sin desarraigar completamente a la mujer de su entorno familiar, especialmente cuando existen hijos menores de edad que dependen de su cuidado y acompañamiento cotidiano.

La Ley 2292 de 2023, en su artículo 7 —que adicionó el artículo 38-I al Código Penal— establece los requisitos para acceder a esta pena sustitutiva. Dichos requisitos fueron sintetizados por la Honorable Corte Constitucional en los siguientes términos: "Solo opera para mujeres cabeza de familia, condenadas por los delitos establecidos en los artículos 239, 240, 241, 375, 376 y 377 del Código Penal, o condenadas a otros delitos cuya pena impuesta sea igual o inferior a ocho (8) años de prisión (art. 2 y art. 7.1). Se debe demostrar que la comisión del delito está asociada a condiciones de marginalidad que afecten la manutención del hogar (art. 2 y art. 7.6). La medida [...] no se aplicará cuando haya condena en firme por otro delito doloso dentro de los (5) años anteriores a la comisión del nuevo ilícito o exista concurso con conductas punibles distintas a las señaladas con anterioridad (art. 2), salvo que se trate de delitos culposos, que tengan como pena principal la multa o que sea por los punibles anteriormente enunciados (art. 7.2). Se debe demostrar que la beneficiaria es madre cabeza de familia. Lo cual significa, para los efectos del proyecto de ley, tener vínculos familiares, demostrando que la condenada ejerce la jefatura del hogar y tiene bajo su cargo afectiva, económica y socialmente de manera permanente hijos menores o personas en condición de discapacidad permanente (art. 7.4)." (Sentencia C-256 de 2022, párrafo 272).

A continuación se acredita el cumplimiento de cada uno de estos requisitos en el caso concreto de la señora [NOMBRE COMPLETO DE LA CONDENADA]:`;

// ── Handler principal ───────────────────────────────────────
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-api-key');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  const apiKey = req.headers['x-api-key'] || process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(401).json({ error: 'API key no configurada' });

  const { campos, archivos } = req.body;
  if (!campos) return res.status(400).json({ error: 'Faltan datos del caso' });

  try {
    const client = new Anthropic({ apiKey });

    // Construir el contenido para Claude — texto + PDFs/imágenes
    const contenidoUsuario = [];

    // Agregar cada archivo como documento
    if (archivos && archivos.length > 0) {
      for (const archivo of archivos) {
        const { nombre, etiqueta, tipo, base64 } = archivo;
        if (tipo === 'application/pdf') {
          contenidoUsuario.push({
            type: 'document',
            source: { type: 'base64', media_type: 'application/pdf', data: base64 },
            title: etiqueta || nombre,
            context: `Prueba aportada al caso: "${etiqueta || nombre}"`
          });
        } else if (tipo.startsWith('image/')) {
          contenidoUsuario.push({
            type: 'image',
            source: { type: 'base64', media_type: tipo, data: base64 }
          });
          contenidoUsuario.push({
            type: 'text',
            text: `La imagen anterior es: "${etiqueta || nombre}"`
          });
        }
      }
    }

    // El prompt principal
    contenidoUsuario.push({
      type: 'text',
      text: `Eres un abogado de la Defensoría del Pueblo de Colombia especializado en la Ley 2292 de 2023 (pena sustitutiva de servicios de utilidad pública). Debes redactar el borrador completo de una solicitud de sustitución de pena para el siguiente caso.

DATOS DEL CASO:
- Nombre completo: ${campos.nombre || '[PENDIENTE]'}
- Cédula: ${campos.cedula || '[PENDIENTE]'}
- Delito: ${campos.delito || '[PENDIENTE]'}
- Artículo Código Penal: ${campos.articulo_cp || '[PENDIENTE]'}
- Fecha de nacimiento: ${campos.fecha_nacimiento || '[PENDIENTE]'}
- Centro de reclusión: ${campos.centro_reclusion || '[PENDIENTE]'}
- Juzgado: ${campos.juzgado || '[PENDIENTE]'}
- Juez(a): ${campos.juez || '[PENDIENTE]'}
- Ciudad del juzgado: ${campos.ciudad_juzgado || '[PENDIENTE]'}
- Defensor(a): ${campos.defensor || '[PENDIENTE]'}
- Email defensor(a): ${campos.email_defensor || '[PENDIENTE]'}
- T.P. defensor(a): ${campos.tp_defensor || '[PENDIENTE]'}
- Plaza SIUP ID: ${campos.plaza_id || '[PENDIENTE]'}
- Entidad SIUP: ${campos.plaza_entidad || '[PENDIENTE]'}
- Ciudad plaza: ${campos.plaza_ciudad || '[PENDIENTE]'}
- NIT entidad: ${campos.plaza_nit || '[PENDIENTE]'}
- Fecha de la solicitud: ${campos.fecha_solicitud || new Date().toLocaleDateString('es-CO', {day:'numeric',month:'long',year:'numeric'})}

PRUEBAS APORTADAS (${archivos?.length || 0} archivo(s)):
${archivos?.map((a,i) => `${i+1}. "${a.etiqueta || a.nombre}" (${a.tipo})`).join('\n') || 'Ninguna'}

INSTRUCCIONES:

1. Lee cuidadosamente TODOS los documentos aportados. Extrae de ellos los datos relevantes: nombres de hijos, edades, clasificación SISBEN, datos ADRES, entrevistas psicosociales, cualquier prueba de marginalidad o jefatura de hogar.

2. Construye la solicitud completa con las 6 secciones del documento. El texto del cuerpo debe adaptarse a las pruebas disponibles — argumenta marginalidad y jefatura de hogar CON LAS PRUEBAS QUE HAY, no con las que idealmente deberían existir.

3. Numera los anexos dinámicamente según los archivos aportados, en este orden preferido:
   - Primero: reporte de entrevista de la Defensoría (si hay)
   - Luego: entrevistas psicosociales / informes forenses (si hay)
   - Luego: bases de datos (SISBEN, ADRES, RUES, Supernotariado — los que haya)
   - Luego: registros civiles (si hay)
   - Luego: registro fotográfico / audiovisual (si hay)
   - Luego: plaza SIUP (si hay)
   - Luego: otros documentos en el orden en que se aportaron

4. En el cuerpo del documento, inserta las referencias cruzadas (ver Anexo No. X) cada vez que menciones una prueba.

5. AJUSTE ADRES — PERSONAS PRIVADAS DE LA LIBERTAD: Si la consulta ADRES muestra a la condenada como "retirada" o sin afiliación vigente, NO interpretes esto como ausencia de cobertura de salud. Esto ocurre porque al ingresar al sistema penitenciario, la atención en salud pasa al Fondo de Atención en Salud para la Población Privada de la Libertad (FASPPL). En ese caso, redacta así: "La señora [NOMBRE] figuraba afiliada al momento de su privación de la libertad como [régimen y calidad que aparezca en la consulta]. En la actualidad aparece retirada del Sistema General de Seguridad Social en Salud, lo cual obedece a que, una vez privada de la libertad, su atención en salud quedó a cargo del Fondo de Atención en Salud para la Población Privada de la Libertad, conforme a la normativa penitenciaria vigente (ver Anexo No. X)."

6. DECRETO REGLAMENTARIO: La Ley 2292 de 2023 está reglamentada por el Decreto 1451 de 2023. Si citas el decreto reglamentario, usa siempre el Decreto 1451 de 2023. NUNCA cites el Decreto 1069 de 2024.

7. Registro formal jurídico. Tercera persona. Los hijos siempre con nombre completo y edad. Cifras en formato "quinientos mil pesos ($500.000)".

8. Si algún dato no está disponible, usa [PENDIENTE].

Responde SOLO con un JSON válido sin backticks, con esta estructura exacta:
{
  "nombre_condenada": "nombre completo de la condenada extraído de los documentos o datos del caso",
  "encabezado_ciudad_fecha": "ciudad y fecha",
  "encabezado_cargo_juez": "cargo del juez",
  "encabezado_nombre_juez": "nombre del juez",
  "encabezado_juzgado": "nombre del juzgado",
  "encabezado_ciudad": "ciudad, departamento",
  "parrafo_intro": "párrafo introductorio donde el defensor se presenta y formula la solicitud",
  "seccion1_contexto": "sección completa de contexto de jefatura y marginalidad — múltiples párrafos separados por doble salto de línea, con referencias cruzadas a los anexos",
  "seccion2_requisitos": "SOLO la verificación de los 4 requisitos de la Ley 2292 con los datos del caso concreto — NO incluyas el texto introductorio de la ley, ese ya se agrega automáticamente",
  "seccion3_plaza": "sección sobre la plaza SIUP y el plan de servicios, citando el Decreto 1451 de 2023",
  "seccion4_peticion": "petición con los tres numerales",
  "lista_anexos": "lista de anexos numerada, un anexo por línea comenzando con •",
  "notificaciones": "notificaciones con dirección de la condenada y del defensor",
  "firma_nombre": "nombre completo del defensor",
  "firma_tp": "número de tarjeta profesional"
}`
    });

    const message = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 8000,
      messages: [{ role: 'user', content: contenidoUsuario }]
    });

    const rawText = message.content[0].text;
    const secciones = JSON.parse(rawText.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim());

    const buffer = crearDocx(secciones);
    const apellido = (campos.nombre || 'usuaria').split(' ').slice(-2).join('_').toUpperCase();
    const filename = `SOLICITUD_${apellido}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.setHeader('Content-Length', buffer.length);
    return res.status(200).send(buffer);

  } catch (err) {
    console.error('Error solicitud:', err.message);
    return res.status(500).json({ error: 'Error interno: ' + err.message });
  }
}

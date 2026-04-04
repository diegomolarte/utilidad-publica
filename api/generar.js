import Anthropic from '@anthropic-ai/sdk';
import {
  Document, Paragraph, TextRun, AlignmentType, Packer,
  BorderStyle, Table, TableRow, TableCell, WidthType,
  Header, Footer, PageNumber
} from 'docx';
 
const Pt = (n) => n * 20;
const Cm = (n) => Math.round(n * 567);
 
function titulo(txt) {
  return new Paragraph({ children: [new TextRun({ text: txt, bold: true, size: Pt(12), font: 'Arial' })], spacing: { before: 0, after: 40 } });
}
function seccionH(txt) {
  return new Paragraph({ children: [new TextRun({ text: txt, bold: true, size: Pt(11), font: 'Arial' })], spacing: { before: 200, after: 100 } });
}
function cuerpo(txt) {
  return new Paragraph({ children: [new TextRun({ text: txt, size: Pt(11), font: 'Arial' })], alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 100 } });
}
function bullet(txt) {
  return new Paragraph({ children: [new TextRun({ text: txt, size: Pt(11), font: 'Arial' })], bullet: { level: 0 }, spacing: { before: 0, after: 60 } });
}
function vacio() {
  return new Paragraph({ children: [new TextRun({ text: '' })], spacing: { after: 60 } });
}
function seccionAParrafos(texto) {
  if (!texto) return [cuerpo('[PENDIENTE]')];
  const parrafos = [];
  for (const bloque of texto.split(/\n\n+/)) {
    const lineas = bloque.split('\n').map(l => l.trim()).filter(Boolean);
    if (!lineas.length) continue;
    if (lineas.every(l => /^[•\-\*]/.test(l))) {
      lineas.forEach(l => parrafos.push(bullet(l.replace(/^[•\-\*]\s*/, ''))));
    } else {
      parrafos.push(cuerpo(lineas.join(' ')));
    }
  }
  return parrafos;
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
 
Registro formal jurídico tercera persona. "manifestó que", "señaló que", "refirió que". Transforma lenguaje informal. Hijos con nombre completo y edad. Cifras en formato "quinientos mil pesos ($500.000)". Si falta dato usa [PENDIENTE].
 
Responde SOLO JSON válido sin backticks:
{"marginalidad":"3-5 párrafos doble salto de línea","jefatura_hogar":"4-6 párrafos","hechos_captura":"5-6 párrafos cronológicos","conclusion":"3-4 párrafos con referencia Ley 2292 de 2023"}` }]
    });
 
    const secciones = JSON.parse(message.content[0].text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim());
    const p = v => v || '[PENDIENTE]';
    const hoy = campos.fecha_entrevista || new Date().toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' });
 
    const doc = new Document({
      sections: [{
        properties: { page: { margin: { top: Cm(2.0), bottom: Cm(2.5), left: Cm(3.0), right: Cm(2.5) } } },
        headers: {
          default: new Header({
            children: [
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: { top: {style:BorderStyle.NONE}, bottom: {style:BorderStyle.NONE}, left: {style:BorderStyle.NONE}, right: {style:BorderStyle.NONE}, insideH: {style:BorderStyle.NONE}, insideV: {style:BorderStyle.NONE} },
                rows: [new TableRow({ children: [
                  new TableCell({ width:{size:40,type:WidthType.PERCENTAGE}, borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, children:[new Paragraph({children:[new TextRun({text:'Defensoría del Pueblo',bold:true,size:Pt(10),font:'Arial',color:'1a4f3a'})],spacing:{after:0}})] }),
                  new TableCell({ width:{size:60,type:WidthType.PERCENTAGE}, borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, children:[new Paragraph({children:[new TextRun({text:'#BuenFuturoHoy',bold:true,size:Pt(14),font:'Arial',color:'c9a84c'})],alignment:AlignmentType.RIGHT,spacing:{after:0}})] })
                ]})]
              }),
              new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'auto' } }, spacing: { before: 80, after: 0 }, children: [] })
            ]
          })
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({ children:[new TextRun({text:'Calle 55 # 10-32 · Sede Nacional · Bogotá, D.C.  |  PBX: (57) (601) 3144000 · Línea Nacional: 01 8000 914814  |  www.defensoria.gov.co',size:Pt(8),font:'Arial',color:'5a5a5a'})], alignment:AlignmentType.CENTER, border:{top:{style:BorderStyle.SINGLE,size:4,color:'auto'}}, spacing:{before:80,after:0} }),
              new Paragraph({ children:[new TextRun({children:[PageNumber.CURRENT],size:Pt(8),font:'Arial',color:'5a5a5a'})], alignment:AlignmentType.RIGHT, spacing:{before:0,after:0} })
            ]
          })
        },
        children: [
          titulo('REPORTE DE ENTREVISTA'), titulo('LEY 2292 DE 2023'), titulo('SERVICIOS DE UTILIDAD PÚBLICA'), vacio(),
          seccionH('1.  Datos generales'),
          bullet(`Nombre completo: ${p(campos.nombre)}`), bullet(`No. documento: ${p(campos.cedula)}`),
          bullet(`Fecha de nacimiento: ${p(campos.fecha_nacimiento)}`), bullet(`Lugar de reclusión: ${p(campos.lugar_reclusion)}`),
          bullet(`Delito: ${p(campos.delito)}`), bullet(`Fecha de la entrevista: ${p(campos.fecha_entrevista)}`),
          bullet(`Nombre del entrevistador(a): ${p(campos.entrevistador)}`), vacio(),
          seccionH('2.  Condiciones asociadas a la marginalidad'), ...seccionAParrafos(secciones.marginalidad), vacio(),
          seccionH('3.  Rol de jefatura de hogar'), ...seccionAParrafos(secciones.jefatura_hogar), vacio(),
          seccionH('4.  Hechos que dieron lugar a la captura en flagrancia.'), ...seccionAParrafos(secciones.hechos_captura), vacio(),
          new Paragraph({children:[new TextRun({text:'Conclusión.',bold:true,size:Pt(11),font:'Arial'})],spacing:{before:200,after:100}}),
          ...seccionAParrafos(secciones.conclusion),
          vacio(), vacio(),
          cuerpo('Entrevistó:'),
          new Paragraph({ children:[new TextRun({text:''})], spacing:{after:600} }),
          cuerpo(p(campos.entrevistador)), cuerpo('Equipo de atención jurídica'),
          cuerpo('Dirección Nacional de Defensoría Pública'), cuerpo(hoy)
        ]
      }]
    });
 
    const buffer = await Packer.toBuffer(doc);
    const filename = `Entrevista a usuaria - ${campos.nombre || 'usuaria'}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.setHeader('Content-Length', buffer.length);
    return res.status(200).send(buffer);
 
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'Error interno: ' + err.message });
  }
}
 

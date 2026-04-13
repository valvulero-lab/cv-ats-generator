const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, LevelFormat, WidthType } = require('docx');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch(e) { body = {}; } }

  const cv = body.cv;
  if (!cv) return res.status(400).json({ error: 'Faltan datos del CV' });

  try {
    const GRAY = '888888';
    const BLACK = '1A1A1A';
    const DARK = '444444';
    const LINE_COLOR = 'DDDDDD';
    const FONT = 'Arial';

    // Helper: horizontal rule via paragraph border
    const rule = () => new Paragraph({
      spacing: { before: 80, after: 80 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE_COLOR, space: 1 } },
      children: []
    });

    // Helper: section title
    const sectionTitle = (text) => new Paragraph({
      spacing: { before: 160, after: 60 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: LINE_COLOR, space: 2 } },
      children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 17, color: GRAY, font: FONT })]
    });

    // Helper: empty line
    const spacer = (pts = 60) => new Paragraph({ spacing: { before: 0, after: pts }, children: [] });

    // Build experience descriptions as bullet-style paragraphs
    const expDescLines = (desc) => {
      if (!desc) return [];
      return desc.split('\n').filter(l => l.trim()).map(line =>
        new Paragraph({
          spacing: { before: 20, after: 20 },
          indent: { left: 200 },
          children: [
            new TextRun({ text: '• ', size: 18, color: DARK, font: FONT }),
            new TextRun({ text: line.replace(/^[•\-]\s*/, '').trim(), size: 18, color: DARK, font: FONT })
          ]
        })
      );
    };

    // Skills as comma-separated text
    const allSkills = [...new Set([...(cv.habilidades||[]), ...(cv.habilidades_sugeridas||[])].map(s=>s.trim()).filter(Boolean))];

    const children = [];

    // NAME - centered, large
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: cv.nombre || '', bold: true, size: 40, color: BLACK, font: FONT })]
    }));

    // CONTACT - centered, gray
    const contactParts = [cv.email, cv.telefono, cv.ubicacion, cv.linkedin].filter(Boolean);
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: contactParts.join('  |  '), size: 17, color: GRAY, font: FONT })]
    }));

    // PERFIL PROFESIONAL
    const perfil = cv.perfil_optimizado || cv.perfil;
    if (perfil) {
      children.push(sectionTitle('Perfil Profesional'));
      children.push(new Paragraph({
        spacing: { before: 60, after: 100 },
        children: [new TextRun({ text: perfil, size: 18, color: DARK, font: FONT, italics: false })]
      }));
    }

    // EXPERIENCIA PROFESIONAL
    const exps = cv.experiencias_optimizadas || cv.experiencias || [];
    if (exps.length) {
      children.push(sectionTitle('Experiencia Profesional'));
      exps.forEach((e, i) => {
        // Cargo — bold, and date on same line via tab
        children.push(new Paragraph({
          spacing: { before: i === 0 ? 60 : 120, after: 0 },
          tabStops: [{ type: 'right', position: 9026 }],
          children: [
            new TextRun({ text: e.cargo || '', bold: true, size: 20, color: BLACK, font: FONT }),
            new TextRun({ text: '\t', size: 18, font: FONT }),
            new TextRun({ text: `${e.desde || ''} – ${e.hasta || ''}`, size: 17, color: GRAY, font: FONT })
          ]
        }));
        // Empresa
        children.push(new Paragraph({
          spacing: { before: 0, after: 40 },
          children: [new TextRun({ text: e.empresa || '', size: 18, color: GRAY, font: FONT })]
        }));
        // Description lines
        children.push(...expDescLines(e.desc_optimizada || e.desc || ''));
      });
    }

    // EDUCACIÓN
    const edus = cv.educacion || [];
    if (edus.length) {
      children.push(sectionTitle('Educación'));
      edus.forEach(e => {
        children.push(new Paragraph({
          spacing: { before: 60, after: 0 },
          tabStops: [{ type: 'right', position: 9026 }],
          children: [
            new TextRun({ text: e.titulo || '', bold: true, size: 19, color: BLACK, font: FONT }),
            new TextRun({ text: '\t', size: 18, font: FONT }),
            new TextRun({ text: e.anio || '', size: 17, color: GRAY, font: FONT })
          ]
        }));
        children.push(new Paragraph({
          spacing: { before: 0, after: 60 },
          children: [new TextRun({ text: e.inst || '', size: 18, color: GRAY, font: FONT })]
        }));
      });
    }

    // HABILIDADES
    if (allSkills.length) {
      children.push(sectionTitle('Habilidades'));
      children.push(new Paragraph({
        spacing: { before: 60, after: 80 },
        children: [new TextRun({ text: allSkills.join(' • '), size: 18, color: DARK, font: FONT })]
      }));
    }

    // IDIOMAS
    const idiomas = cv.idiomas || [];
    if (idiomas.length) {
      children.push(sectionTitle('Idiomas'));
      idiomas.forEach(l => {
        children.push(new Paragraph({
          spacing: { before: 60, after: 40 },
          children: [
            new TextRun({ text: l.nombre || '', bold: true, size: 18, color: BLACK, font: FONT }),
            new TextRun({ text: `  —  ${l.nivel || ''}`, size: 18, color: DARK, font: FONT })
          ]
        }));
      });
    }

    // CERTIFICACIONES
    if (cv.certificaciones) {
      children.push(sectionTitle('Certificaciones'));
      cv.certificaciones.split('\n').filter(l => l.trim()).forEach(line => {
        children.push(new Paragraph({
          spacing: { before: 40, after: 40 },
          indent: { left: 200 },
          children: [
            new TextRun({ text: '• ', size: 18, color: DARK, font: FONT }),
            new TextRun({ text: line.trim(), size: 18, color: DARK, font: FONT })
          ]
        }));
      });
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 900, right: 900, bottom: 900, left: 900 }
          }
        },
        children
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = `CV_${(cv.nombre || 'CV').replace(/\s+/g, '_')}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    res.status(200).send(buffer);

  } catch (error) {
    console.error('DOCX error:', error);
    res.status(500).json({ error: 'Error generando el Word: ' + error.message });
  }
}

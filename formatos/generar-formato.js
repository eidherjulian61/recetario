/**
 * Genera el "Formato único de registro para cuenta de cobro mensual"
 * de la Dirección de Procesos de Selección — ESAP.
 *   node formatos/generar-formato.js
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, VerticalAlign, HeightRule,
  Header, Footer, PageNumber, TabStopType, TabStopPosition, PageOrientation,
} = require('docx');

// ───────────────────────── paleta institucional ─────────────────────────
const AZUL      = '25478C'; // azul ESAP (fondo del logotipo)
const AZUL_OSC  = '1A3567'; // azul profundo, numerales y titulares
const AZUL_TEN  = 'E7ECF6'; // azul tenue, fondo de etiquetas
const AZUL_MED  = 'CBD6EA'; // azul medio, bandas de tabla
const GRIS_LIN  = 'B9C2D4'; // gris azulado de líneas
const GRIS_TXT  = '5F6980'; // gris de textos auxiliares
const GRIS_FON  = 'F1F2F5'; // gris muy claro, uso interno de la entidad
const AMBAR     = 'B8860B'; // ámbar sobrio, sólo para el aviso inicial

const FUENTE = 'Arial';
const ANCHO  = 10080; // ancho útil: carta (12240) − márgenes (2 × 1080)

// ───────────────────────────── utilidades ───────────────────────────────
const linea = (sz = 4, color = GRIS_LIN) => ({ style: BorderStyle.SINGLE, size: sz, color });
const marco = (sz = 4, color = GRIS_LIN) => ({
  top: linea(sz, color), bottom: linea(sz, color),
  left: linea(sz, color), right: linea(sz, color),
});
const SIN_BORDE = {
  top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
  left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
};

const t = (text, o = {}) => new TextRun({
  text, font: FUENTE, size: o.size || 18, bold: o.bold, italics: o.italics,
  color: o.color || '1C1C22', characterSpacing: o.spacing, allCaps: o.caps,
});

const p = (runs, o = {}) => new Paragraph({
  children: Array.isArray(runs) ? runs : [runs],
  alignment: o.align, spacing: { before: o.before || 0, after: o.after || 0, line: o.line },
  border: o.border, indent: o.indent, shading: o.shading,
});

const vacio = (after = 0) => p([t('')], { after });

/** Celda de etiqueta: fondo azul tenue, texto versalita azul profundo. */
const etq = (texto, ancho, o = {}) => new TableCell({
  width: { size: ancho, type: WidthType.DXA },
  columnSpan: o.span,
  shading: { type: ShadingType.CLEAR, fill: o.fill || AZUL_TEN, color: 'auto' },
  margins: { top: 40, bottom: 40, left: 90, right: 90 },
  verticalAlign: VerticalAlign.CENTER,
  borders: marco(),
  children: [p(
    texto.endsWith(' (*)')
      ? [t(texto.slice(0, -4), { size: 15, bold: true, color: AZUL_OSC, spacing: 6 }),
         t(' (*)', { size: 15, bold: true, color: AMBAR })]
      : [t(texto, { size: 15, bold: true, color: AZUL_OSC, spacing: 6 })],
  )],
});

/** Celda de captura: fondo blanco, con guía opcional en gris. */
const cap = (ancho, o = {}) => {
  const hijos = [];
  if (o.texto) {
    hijos.push(p([t(o.texto, { size: o.tam || 17, bold: o.bold, italics: o.italics, color: o.color })], { align: o.align }));
  } else {
    hijos.push(p([t('')], {}));
  }
  if (o.guia) hijos.push(p([t(o.guia, { size: 13, italics: true, color: GRIS_TXT })]));
  return new TableCell({
    width: { size: ancho, type: WidthType.DXA },
    columnSpan: o.span,
    shading: { type: ShadingType.CLEAR, fill: o.fill || 'FFFFFF', color: 'auto' },
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    verticalAlign: VerticalAlign.CENTER,
    borders: marco(),
    children: hijos,
  });
};

/** Celda con casillas de verificación en línea. */
const casillas = (ancho, opciones, o = {}) => new TableCell({
  width: { size: ancho, type: WidthType.DXA },
  columnSpan: o.span,
  shading: { type: ShadingType.CLEAR, fill: 'FFFFFF', color: 'auto' },
  margins: { top: 50, bottom: 50, left: 90, right: 90 },
  verticalAlign: VerticalAlign.CENTER,
  borders: marco(),
  children: [p(opciones.flatMap((op, i) => [
    t(i ? '   ' : '', { size: 15 }),
    t('☐ ', { size: 19, color: AZUL }),
    t(op, { size: 15 }),
  ]))],
});

const fila = (celdas, alto = 470) => new TableRow({
  height: { value: alto, rule: HeightRule.ATLEAST },
  children: celdas,
});

const tabla = (anchos, filas) => new Table({
  columnWidths: anchos,
  width: { size: ANCHO, type: WidthType.DXA },
  borders: marco(6, AZUL_OSC),
  rows: filas,
});

/** Banda de sección: numeral en azul profundo + rótulo sobre azul institucional. */
const seccion = (num, titulo, nota) => new Table({
  columnWidths: [560, ANCHO - 560],
  width: { size: ANCHO, type: WidthType.DXA },
  borders: marco(0, AZUL),
  rows: [new TableRow({
    height: { value: 400, rule: HeightRule.ATLEAST },
    children: [
      new TableCell({
        width: { size: 560, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: AZUL_OSC, color: 'auto' },
        verticalAlign: VerticalAlign.CENTER, borders: marco(0, AZUL_OSC),
        children: [p([t(String(num), { size: 22, bold: true, color: 'FFFFFF' })], { align: AlignmentType.CENTER })],
      }),
      new TableCell({
        width: { size: ANCHO - 560, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: AZUL, color: 'auto' },
        verticalAlign: VerticalAlign.CENTER, borders: marco(0, AZUL),
        margins: { left: 140, right: 120, top: 40, bottom: 40 },
        children: [
          p([t(titulo, { size: 18, bold: true, color: 'FFFFFF', spacing: 24 })]),
          ...(nota ? [p([t(nota, { size: 13, color: 'D6DEEE', italics: true })])] : []),
        ],
      }),
    ],
  })],
});

const banda = (num, titulo, nota) => { cuerpo.push(vacio(240)); cuerpo.push(seccion(num, titulo, nota)); };

const aclaracion = (texto) => p(
  [t('Nota. ', { size: 14, bold: true, color: AZUL_OSC }), t(texto, { size: 14, italics: true, color: GRIS_TXT })],
  { before: 70, after: 200 },
);

// ─────────────────────────── encabezado de página ───────────────────────
const logo = fs.readFileSync(path.join(__dirname, 'recursos', 'logo-esap.png'));

const rotuloControl = (izq, der) => new TableRow({
  height: { value: 200, rule: HeightRule.ATLEAST },
  children: [
    new TableCell({
      width: { size: 1180, type: WidthType.DXA }, borders: marco(2),
      shading: { type: ShadingType.CLEAR, fill: AZUL_TEN, color: 'auto' },
      margins: { left: 70, right: 40, top: 20, bottom: 20 },
      verticalAlign: VerticalAlign.CENTER,
      children: [p([t(izq, { size: 13, bold: true, color: AZUL_OSC })])],
    }),
    new TableCell({
      width: { size: 1500, type: WidthType.DXA }, borders: marco(2),
      margins: { left: 70, right: 40, top: 20, bottom: 20 },
      verticalAlign: VerticalAlign.CENTER,
      children: [p(der)],
    }),
  ],
});

const encabezado = new Header({
  children: [
    new Table({
      columnWidths: [2760, 4640, 2680],
      width: { size: ANCHO, type: WidthType.DXA },
      borders: marco(6, AZUL_OSC),
      rows: [new TableRow({
        children: [
          new TableCell({
            width: { size: 2760, type: WidthType.DXA }, borders: marco(6, AZUL_OSC),
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 80, bottom: 80, left: 110, right: 110 },
            children: [p(new ImageRun({ data: logo, type: 'png', transformation: { width: 152, height: 52 } }))],
          }),
          new TableCell({
            width: { size: 4640, type: WidthType.DXA }, borders: marco(6, AZUL_OSC),
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              p([t('Proceso: Gestión Contractual', { size: 13, color: GRIS_TXT, spacing: 10 })], { align: AlignmentType.CENTER }),
              p([t('FORMATO ÚNICO DE REGISTRO PARA CUENTA DE COBRO MENSUAL', { size: 17, bold: true, color: AZUL_OSC, spacing: 8 })],
                { align: AlignmentType.CENTER, before: 30 }),
              p([t('Dirección de Procesos de Selección', { size: 14, color: '3B4457' })], { align: AlignmentType.CENTER, before: 20 }),
            ],
          }),
          new TableCell({
            width: { size: 2680, type: WidthType.DXA }, borders: marco(6, AZUL_OSC),
            verticalAlign: VerticalAlign.CENTER, margins: { top: 40, bottom: 40, left: 0, right: 0 },
            children: [new Table({
              columnWidths: [1180, 1500],
              width: { size: 2680, type: WidthType.DXA },
              borders: marco(2),
              rows: [
                rotuloControl('Código', [t('ESAP-GC-FO-032', { size: 13 })]),
                rotuloControl('Versión', [t('04', { size: 13 })]),
                rotuloControl('Vigente desde', [t('15-01-2026', { size: 13 })]),
                rotuloControl('Página', [
                  new TextRun({ children: [PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES], font: FUENTE, size: 13 }),
                ]),
              ],
            })],
          }),
        ],
      })],
    }),
    vacio(60),
  ],
});

const piePagina = new Footer({
  children: [
    p([t('')], { border: { top: linea(6, AZUL) }, after: 60 }),
    p([t('Escuela Superior de Administración Pública · Sede Nacional CAN, calle 44 No. 53-37, Bogotá D.C. · www.esap.edu.co',
      { size: 12, color: GRIS_TXT })], { align: AlignmentType.CENTER }),
    p([t('Formato de uso obligatorio para los contratistas de la Dirección de Procesos de Selección. Diligenciar en su totalidad; no se admiten tachones ni enmendaduras.',
      { size: 12, italics: true, color: '8B93A5' })], { align: AlignmentType.CENTER }),
  ],
});

// ─────────────────────────────── cuerpo ─────────────────────────────────
const cuerpo = [];

// Aviso preliminar
cuerpo.push(new Table({
  columnWidths: [ANCHO],
  width: { size: ANCHO, type: WidthType.DXA },
  borders: { top: linea(4, AZUL_MED), bottom: linea(4, AZUL_MED), right: linea(4, AZUL_MED), left: { style: BorderStyle.SINGLE, size: 24, color: AMBAR } },
  rows: [new TableRow({ children: [new TableCell({
    width: { size: ANCHO, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: 'FAFAF5', color: 'auto' },
    margins: { top: 130, bottom: 130, left: 180, right: 160 },
    borders: { top: linea(4, AZUL_MED), bottom: linea(4, AZUL_MED), right: linea(4, AZUL_MED), left: { style: BorderStyle.SINGLE, size: 24, color: AMBAR } },
    children: [
      p([t('Antes de diligenciar', { size: 16, bold: true, color: AZUL_OSC, spacing: 14 })], { after: 70 }),
      p([t('Este formato consolida la información que la Dirección de Procesos de Selección requiere para tramitar el pago mensual del contrato de prestación de servicios. Diligéncielo en su totalidad —a mano en tinta negra o en medio digital— y radíquelo dentro de los cinco (5) primeros días hábiles del mes siguiente al periodo cobrado, junto con los soportes de la sección 8.',
        { size: 15 })], { after: 70 }),
      p([t('Los campos marcados con ', { size: 15 }), t('(*)', { size: 15, bold: true, color: AMBAR }),
        t(' son indispensables: su omisión obliga a devolver la cuenta sin trámite. Si un campo no aplica a su situación, escriba ', { size: 15 }),
        t('«No aplica»', { size: 15, italics: true }), t(' en lugar de dejarlo en blanco.', { size: 15 })]),
    ],
  })] })],
}));

// 1. Identificación del contratista
banda(1, 'IDENTIFICACIÓN DEL CONTRATISTA', 'Registre los datos tal como figuran en el documento de identidad y en el RUT.');
cuerpo.push(tabla([2400, 7680], [
  fila([etq('Nombres y apellidos completos (*)', 2400), cap(7680)]),
]));
cuerpo.push(tabla([2400, 2500, 1500, 1500, 2180], [
  fila([etq('Documento de identidad (*)', 2400),
        casillas(2500, ['C.C.', 'C.E.', 'Pas.']),
        etq('Número', 1500), cap(1500), cap(2180, { guia: 'Lugar de expedición' })]),
  fila([etq('Dirección de correspondencia', 2400), cap(4000, { span: 2 }),
        etq('Ciudad', 1500), cap(2180)]),
  fila([etq('Teléfono de contacto (*)', 2400), cap(2500),
        etq('Correo electrónico (*)', 1500), cap(3680, { span: 2, guia: 'A esta dirección se enviarán las notificaciones del trámite' })]),
  fila([etq('Profesión u oficio', 2400), cap(4000, { span: 2 }),
        etq('Tarjeta profesional No.', 1500), cap(2180)]),
]));
cuerpo.push(aclaracion('Cualquier cambio en la cuenta bancaria, la dirección o el correo electrónico debe informarse por escrito al supervisor antes de radicar la cuenta de cobro.'));

// 2. Información para el pago
banda(2, 'INFORMACIÓN PARA EL PAGO');
cuerpo.push(tabla([2100, 2900, 1700, 3380], [
  fila([etq('Entidad bancaria (*)', 2100), cap(2900),
        etq('Tipo de cuenta (*)', 1700), casillas(3380, ['Ahorros', 'Corriente'])]),
  fila([etq('Número de cuenta (*)', 2100), cap(2900),
        etq('Titular de la cuenta', 1700), cap(3380, { guia: 'Debe ser el contratista' })]),
]));
cuerpo.push(aclaracion('Adjunte certificación bancaria con expedición no mayor a treinta (30) días. La entidad no realiza pagos a cuentas de terceros.'));

// 3. Información del contrato
banda(3, 'INFORMACIÓN DEL CONTRATO', 'Consulte estos datos en el contrato suscrito y en el acta de inicio.');
cuerpo.push(tabla([2100, 1900, 1900, 1900, 2280], [
  fila([etq('Contrato de prestación de servicios No. (*)', 2100), cap(1900),
        etq('Vigencia', 1900), cap(1900), cap(2280, { guia: 'Año' })]),
  fila([etq('Fecha de suscripción', 2100), cap(1900, { guia: 'DD/MM/AAAA' }),
        etq('Acta de inicio', 1900), cap(1900, { guia: 'DD/MM/AAAA' }),
        cap(2280, { guia: 'Fecha de terminación' })]),
]));
cuerpo.push(tabla([2100, 7980], [
  fila([etq('Objeto contractual (*)', 2100), cap(7980, { guia: 'Transcríbalo literalmente de la cláusula primera del contrato' })], 900),
]));
cuerpo.push(tabla([2100, 2500, 1900, 1300, 2280], [
  fila([etq('Valor total del contrato', 2100), cap(2500, { guia: '$' }),
        etq('Plazo de ejecución', 1900), cap(3580, { span: 2, guia: 'Meses y días' })]),
  fila([etq('CDP No.', 2100), cap(2500), etq('RP No.', 1900), cap(3580, { span: 2 })]),
  fila([etq('Supervisor del contrato (*)', 2100), cap(2500),
        etq('Cargo', 1900), cap(3580, { span: 2 })]),
  fila([etq('Dependencia', 2100), cap(7980, { span: 4, texto: 'Dirección de Procesos de Selección', color: AZUL_OSC, bold: true, fill: 'F7F9FD' })]),
]));

cuerpo.push(new Paragraph({ children: [t('')], pageBreakBefore: true, spacing: { after: 0 } }));

// 4. Periodo y valor
banda(4, 'PERIODO Y VALOR DE LA CUENTA DE COBRO');
cuerpo.push(tabla([2100, 1300, 1250, 1350, 1300, 2780], [
  fila([etq('Cuenta de cobro No. (*)', 2100), cap(1300, { guia: 'Consecutivo' }),
        etq('Corresponde al pago', 1250), cap(1350, { guia: 'No. __ de __' }),
        etq('Modalidad', 1300), casillas(2780, ['Mensual', 'Proporcional'])]),
  fila([etq('Periodo cobrado (*)', 2100), cap(1300, { guia: 'Desde' }), cap(1250, { guia: 'Hasta' }),
        etq('Días ejecutados', 1350), cap(4080, { span: 2 })]),
]));
cuerpo.push(vacio(120));
cuerpo.push(tabla([560, 5920, 3600], [
  new TableRow({ height: { value: 320, rule: HeightRule.ATLEAST }, children: [
    new TableCell({ width: { size: 560, type: WidthType.DXA }, borders: marco(),
      shading: { type: ShadingType.CLEAR, fill: AZUL_MED, color: 'auto' }, verticalAlign: VerticalAlign.CENTER,
      children: [p([t('#', { size: 14, bold: true, color: AZUL_OSC })], { align: AlignmentType.CENTER })] }),
    new TableCell({ width: { size: 5920, type: WidthType.DXA }, borders: marco(),
      shading: { type: ShadingType.CLEAR, fill: AZUL_MED, color: 'auto' }, verticalAlign: VerticalAlign.CENTER,
      margins: { left: 120 },
      children: [p([t('CONCEPTO DE LIQUIDACIÓN', { size: 14, bold: true, color: AZUL_OSC, spacing: 12 })])] }),
    new TableCell({ width: { size: 3600, type: WidthType.DXA }, borders: marco(),
      shading: { type: ShadingType.CLEAR, fill: AZUL_MED, color: 'auto' }, verticalAlign: VerticalAlign.CENTER,
      margins: { right: 120 },
      children: [p([t('VALOR (PESOS)', { size: 14, bold: true, color: AZUL_OSC, spacing: 12 })], { align: AlignmentType.RIGHT })] }),
  ] }),
  ...[
    ['1', 'Valor bruto del periodo cobrado'],
    ['2', 'Menos retención en la fuente'],
    ['3', 'Menos retención de ICA (si aplica en el municipio)'],
    ['4', 'Menos otros descuentos (indique cuál)'],
  ].map(([n, c]) => fila([
    new TableCell({ width: { size: 560, type: WidthType.DXA }, borders: marco(), verticalAlign: VerticalAlign.CENTER,
      children: [p([t(n, { size: 16, color: GRIS_TXT })], { align: AlignmentType.CENTER })] }),
    cap(5920, { texto: c }),
    cap(3600, { align: AlignmentType.RIGHT, texto: '$' }),
  ], 420)),
  new TableRow({ height: { value: 440, rule: HeightRule.ATLEAST }, children: [
    new TableCell({ width: { size: 6480, type: WidthType.DXA }, columnSpan: 2, borders: marco(),
      shading: { type: ShadingType.CLEAR, fill: AZUL, color: 'auto' }, verticalAlign: VerticalAlign.CENTER,
      margins: { left: 120 },
      children: [p([t('VALOR NETO A PAGAR', { size: 17, bold: true, color: 'FFFFFF', spacing: 16 })])] }),
    new TableCell({ width: { size: 3600, type: WidthType.DXA }, borders: marco(),
      shading: { type: ShadingType.CLEAR, fill: AZUL, color: 'auto' }, verticalAlign: VerticalAlign.CENTER,
      margins: { right: 120 },
      children: [p([t('$', { size: 17, bold: true, color: 'FFFFFF' })], { align: AlignmentType.RIGHT })] }),
  ] }),
]));
cuerpo.push(vacio(120));
cuerpo.push(tabla([2400, 7680], [
  fila([etq('Valor neto en letras (*)', 2400), cap(7680, { guia: 'Ejemplo: cuatro millones doscientos mil pesos M/CTE' })], 560),
]));
cuerpo.push(aclaracion('El valor bruto debe corresponder al pago pactado en la cláusula de forma de pago. Cuando el periodo sea incompleto, liquide proporcionalmente los días ejecutados y anexe el cálculo.'));

// 5. Tributaria
banda(5, 'INFORMACIÓN TRIBUTARIA', 'Base para la depuración de la retención en la fuente (artículos 383 y 387 del Estatuto Tributario).');
cuerpo.push(tabla([2400, 2100, 2100, 3480], [
  fila([etq('RUT actualizado', 2400), casillas(2100, ['Sí', 'No']),
        etq('Código actividad económica (CIIU)', 2100), cap(3480)]),
  fila([etq('Responsable de IVA', 2400), casillas(2100, ['Sí', 'No']),
        etq('Declarante de renta', 2100), casillas(3480, ['Sí', 'No'])]),
  fila([etq('¿Solicita depuración de la base de retención?', 2400), casillas(2100, ['Sí', 'No']),
        cap(5580, { span: 2, texto: 'Si marcó «Sí», relacione a continuación los valores mensuales y adjunte los certificados.', italics: true, tam: 15, color: GRIS_TXT })]),
  fila([etq('Municipio donde ejecuta las actividades', 2400), cap(2100),
        etq('¿Inscrito en el registro de industria y comercio?', 2100), casillas(3480, ['Sí', 'No'])]),
]));
cuerpo.push(vacio(100));
cuerpo.push(tabla([2520, 2520, 2520, 2520], [
  fila([etq('Dependientes (10 %)', 2520, { fill: AZUL_TEN }), etq('Medicina prepagada', 2520, { fill: AZUL_TEN }),
        etq('Intereses de vivienda', 2520, { fill: AZUL_TEN }), etq('Aportes AFC / voluntarios', 2520, { fill: AZUL_TEN })], 320),
  fila([cap(2520, { texto: '$' }), cap(2520, { texto: '$' }), cap(2520, { texto: '$' }), cap(2520, { texto: '$' })]),
  fila([casillas(2520, ['Soporte adjunto']), casillas(2520, ['Soporte adjunto']),
        casillas(2520, ['Soporte adjunto']), casillas(2520, ['Soporte adjunto'])], 340),
]));
cuerpo.push(aclaracion('Cada deducción requiere el soporte respectivo (certificación de dependientes bajo la gravedad de juramento, certificados de la entidad de medicina prepagada, del banco o del fondo). Sin soporte no se aplica la depuración.'));

cuerpo.push(new Paragraph({ children: [t('')], pageBreakBefore: true, spacing: { after: 0 } }));

// 6. Seguridad social
banda(6, 'APORTES AL SISTEMA DE SEGURIDAD SOCIAL (PILA)', 'El ingreso base de cotización corresponde al 40 % del valor bruto mensual del contrato.');
cuerpo.push(tabla([2400, 2200, 1900, 3580], [
  fila([etq('Ingreso base de cotización (IBC)', 2400), cap(2200, { texto: '$' }),
        etq('Planilla PILA No. (*)', 1900), cap(3580)]),
  fila([etq('Periodo cotizado (*)', 2400), cap(2200, { guia: 'Mes / año' }),
        etq('Fecha de pago de la planilla', 1900), cap(3580, { guia: 'DD/MM/AAAA' })]),
]));
cuerpo.push(vacio(100));
const encApo = (texto, ancho, align) => new TableCell({
  width: { size: ancho, type: WidthType.DXA }, borders: marco(),
  shading: { type: ShadingType.CLEAR, fill: AZUL_MED, color: 'auto' },
  margins: { top: 50, bottom: 50, left: 110, right: 110 }, verticalAlign: VerticalAlign.CENTER,
  children: [p([t(texto, { size: 14, bold: true, color: AZUL_OSC, spacing: 10 })], { align })],
});
cuerpo.push(tabla([1700, 3100, 1000, 1900, 2380], [
  fila([encApo('APORTE', 1700), encApo('ADMINISTRADORA', 3100), encApo('TARIFA', 1000, AlignmentType.CENTER),
        encApo('VALOR PAGADO', 1900, AlignmentType.RIGHT), encApo('OBSERVACIONES', 2380)], 340),
  fila([etq('Salud', 1700), cap(3100, { guia: 'EPS a la que se encuentra afiliado' }),
        cap(1000, { texto: '12,5 %', align: AlignmentType.CENTER, tam: 15, color: GRIS_TXT }),
        cap(1900, { texto: '$', align: AlignmentType.RIGHT }), cap(2380)]),
  fila([etq('Pensión', 1700), cap(3100, { guia: 'Fondo de pensiones' }),
        cap(1000, { texto: '16 %', align: AlignmentType.CENTER, tam: 15, color: GRIS_TXT }),
        cap(1900, { texto: '$', align: AlignmentType.RIGHT }), cap(2380)]),
  fila([etq('Riesgos laborales', 1700), cap(3100, { guia: 'ARL — sólo si la actividad es de riesgo IV o V' }),
        cap(1000, { align: AlignmentType.CENTER }),
        cap(1900, { texto: '$', align: AlignmentType.RIGHT }),
        casillas(2380, ['Nivel IV', 'Nivel V'])]),
]));

// 7. Actividades
banda(7, 'ACTIVIDADES EJECUTADAS EN EL PERIODO', 'Relacione cada obligación contractual atendida durante el mes y el producto que la respalda.');
const encAct = (texto, ancho) => new TableCell({
  width: { size: ancho, type: WidthType.DXA }, borders: marco(),
  shading: { type: ShadingType.CLEAR, fill: AZUL_MED, color: 'auto' },
  margins: { top: 50, bottom: 50, left: 100, right: 100 }, verticalAlign: VerticalAlign.CENTER,
  children: [p([t(texto, { size: 14, bold: true, color: AZUL_OSC, spacing: 10 })], { align: AlignmentType.CENTER })],
});
cuerpo.push(tabla([620, 2860, 3900, 2700], [
  fila([encAct('No.', 620), encAct('Obligación contractual', 2860), encAct('Actividades ejecutadas', 3900), encAct('Producto o evidencia', 2700)], 340),
  ...[1, 2, 3, 4, 5, 6].map((n) => fila([
    new TableCell({ width: { size: 620, type: WidthType.DXA }, borders: marco(), verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.CLEAR, fill: n % 2 ? 'FFFFFF' : 'FAFBFD', color: 'auto' },
      children: [p([t(String(n), { size: 16, color: GRIS_TXT })], { align: AlignmentType.CENTER })] }),
    cap(2860, { fill: n % 2 ? 'FFFFFF' : 'FAFBFD' }),
    cap(3900, { fill: n % 2 ? 'FFFFFF' : 'FAFBFD' }),
    cap(2700, { fill: n % 2 ? 'FFFFFF' : 'FAFBFD' }),
  ], 900)),
]));
cuerpo.push(aclaracion('Si requiere más espacio, continúe en hoja anexa con la misma estructura, numérela y fírmela. Indique aquí el número de anexos: ____.'));

cuerpo.push(new Paragraph({ children: [t('')], pageBreakBefore: true, spacing: { after: 0 } }));

// 8. Soportes
banda(8, 'RELACIÓN DE SOPORTES ADJUNTOS', 'Marque los documentos que radica con esta cuenta de cobro.');
const chk = (texto, ancho, fill) => new TableCell({
  width: { size: ancho, type: WidthType.DXA }, borders: marco(),
  shading: { type: ShadingType.CLEAR, fill: fill || 'FFFFFF', color: 'auto' },
  margins: { top: 60, bottom: 60, left: 120, right: 100 }, verticalAlign: VerticalAlign.CENTER,
  children: [p([t('☐  ', { size: 20, color: AZUL }), t(texto, { size: 16 })])],
});
cuerpo.push(tabla([5040, 5040], [
  fila([chk('Cuenta de cobro firmada por el contratista', 5040), chk('Informe mensual de actividades', 5040)], 360),
  fila([chk('Planilla de aportes al sistema de seguridad social', 5040, 'FAFBFD'), chk('Certificación bancaria vigente', 5040, 'FAFBFD')], 360),
  fila([chk('Copia del RUT actualizado', 5040), chk('Copia del documento de identidad', 5040)], 360),
  fila([chk('Soportes de deducciones tributarias', 5040, 'FAFBFD'), chk('Otros (especifique): ______________________', 5040, 'FAFBFD')], 360),
]));
cuerpo.push(tabla([2400, 1400, 2400, 3880], [
  fila([etq('Total de folios radicados', 2400), cap(1400),
        etq('Fecha de radicación', 2400), cap(3880, { guia: 'DD/MM/AAAA' })], 380),
]));

// 9. Declaración
banda(9, 'DECLARACIÓN DEL CONTRATISTA');
cuerpo.push(new Table({
  columnWidths: [ANCHO],
  width: { size: ANCHO, type: WidthType.DXA },
  borders: marco(4, GRIS_LIN),
  rows: [new TableRow({ cantSplit: true, children: [new TableCell({
    width: { size: ANCHO, type: WidthType.DXA }, borders: marco(4, GRIS_LIN),
    shading: { type: ShadingType.CLEAR, fill: 'FCFDFE', color: 'auto' },
    margins: { top: 140, bottom: 140, left: 180, right: 180 },
    children: [
      p([t('Bajo la gravedad de juramento declaro que la información consignada en este formato es veraz, que ejecuté las actividades relacionadas durante el periodo cobrado y que me encuentro al día en el pago de los aportes al sistema de seguridad social integral, conforme al artículo 50 de la Ley 789 de 2002.',
        { size: 15 })], { after: 90 }),
      p([t('Manifiesto igualmente que no me encuentro incurso en causal de inhabilidad o incompatibilidad para contratar con el Estado y que autorizo a la ESAP a verificar los soportes aportados y a tratar mis datos personales para los fines del trámite de pago, en los términos de la Ley 1581 de 2012.',
        { size: 15 })]),
    ],
  })] })],
}));
cuerpo.push(vacio(140));

// 10. Firmas
const firma = (rotulo, lineas) => new TableCell({
  width: { size: 4940, type: WidthType.DXA }, borders: SIN_BORDE,
  margins: { top: 60, bottom: 60, left: 60, right: 60 },
  children: [
    p([t('')], { after: 700 }),
    p([t('')], { border: { bottom: linea(6, '2B2B33') }, after: 90 }),
    p([t(rotulo, { size: 15, bold: true, color: AZUL_OSC, spacing: 10 })]),
    ...lineas.map((l) => p([t(l, { size: 14, color: GRIS_TXT })], { before: 60 })),
  ],
});
banda(10, 'FIRMAS');
cuerpo.push(new Table({
  columnWidths: [4940, 200, 4940],
  width: { size: ANCHO, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: [new TableRow({ cantSplit: true, children: [
    firma('CONTRATISTA', ['Nombre: ____________________________________', 'C.C. No. ___________________________________', 'Fecha: _____________________________________']),
    new TableCell({ width: { size: 200, type: WidthType.DXA }, borders: SIN_BORDE, children: [p([t('')])] }),
    firma('SUPERVISOR DEL CONTRATO — VISTO BUENO', ['Nombre: ____________________________________', 'Cargo: _____________________________________', 'Fecha: _____________________________________']),
  ] })],
}));
cuerpo.push(vacio(140));

// 11. Uso exclusivo de la entidad
banda(11, 'ESPACIO RESERVADO PARA LA DIRECCIÓN DE PROCESOS DE SELECCIÓN', 'No diligenciar por el contratista.');
cuerpo.push(tabla([2200, 2600, 2200, 3080], [
  fila([etq('Recibido por', 2200, { fill: GRIS_FON }), cap(2600, { fill: GRIS_FON }),
        etq('Fecha y hora de recepción', 2200, { fill: GRIS_FON }), cap(3080, { fill: GRIS_FON })]),
  fila([etq('Revisión de soportes', 2200, { fill: GRIS_FON }),
        casillas(2600, ['Completa', 'Devuelta']),
        etq('Radicado interno No.', 2200, { fill: GRIS_FON }), cap(3080, { fill: GRIS_FON })]),
  fila([etq('Observaciones', 2200, { fill: GRIS_FON }), cap(7880, { span: 3, fill: GRIS_FON })], 1100),
]));

// ─────────────────────────────── documento ──────────────────────────────
const doc = new Document({
  creator: 'Dirección de Procesos de Selección — ESAP',
  title: 'Formato único de registro para cuenta de cobro mensual',
  description: 'Formato de captura de información para la cuenta de cobro mensual de los contratistas de la Dirección de Procesos de Selección de la ESAP.',
  styles: {
    default: {
      document: { run: { font: FUENTE, size: 18, color: '1C1C22' } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840, orientation: PageOrientation.PORTRAIT },
        margin: { top: 1900, right: 1080, bottom: 1400, left: 1080, header: 560, footer: 460 },
      },
    },
    headers: { default: encabezado },
    footers: { default: piePagina },
    children: cuerpo,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const salida = path.join(__dirname, 'Formato-cuenta-de-cobro-mensual-ESAP.docx');
  fs.writeFileSync(salida, buf);
  console.log('OK →', salida, (buf.length / 1024).toFixed(1) + ' KB');
});

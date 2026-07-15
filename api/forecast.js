// Vercel Serverless Function
// POST /api/forecast
//
// Agente G (Predictive Growth Engine) — pronóstico de demanda estacional +
// radar de fuga (churn).
//
// Todo lo que es un número (uplift esperado por fecha, score de precisión,
// distancia en días entre fechas, conteos de clientes en riesgo) se calcula
// acá mismo en JS, de forma determinística, NUNCA por Claude. Claude entra
// al final solo para redactar accionables y recomendaciones sobre esos
// números ya calculados — igual que en los Agentes A y F.

const MAX_FILAS_TABLA = 8000;
const MAX_TEXTO_LARGO = 6000;
const MESES_ADELANTE = 6;
const MAX_FECHAS = 8;

/* ================= PARSERS (mismos que Agentes A/F) ================= */

function bufferDesdeBase64(base64) {
  const data = String(base64).split(',').pop();
  return Buffer.from(data, 'base64');
}

function parseTabular(buffer, nombreArchivo) {
  let XLSX;
  try { XLSX = require('xlsx'); } catch (e) { return { error: 'No se pudo cargar el parser de planillas (xlsx).' }; }
  try {
    const wb = XLSX.read(buffer, { type: 'buffer' });
    const hoja = wb.Sheets[wb.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(hoja, { defval: '' });
    return {
      headers: filas.length ? Object.keys(filas[0]) : [],
      filas: filas.slice(0, MAX_FILAS_TABLA),
      totalFilasOriginal: filas.length,
      truncado: filas.length > MAX_FILAS_TABLA
    };
  } catch (e) { return { error: `No se pudo leer "${nombreArchivo}" como planilla: ${e.message}` }; }
}

async function parsePDF(buffer, nombreArchivo) {
  let pdfParse;
  try { pdfParse = require('pdf-parse'); } catch (e) { return { error: 'No se pudo cargar el parser de PDF.' }; }
  try {
    const data = await pdfParse(buffer);
    return { texto: (data.text || '').slice(0, MAX_TEXTO_LARGO) };
  } catch (e) { return { error: `No se pudo leer "${nombreArchivo}" como PDF: ${e.message}` }; }
}

async function parseArchivo(archivo) {
  const nombre = archivo.nombre || 'archivo';
  const ext = (nombre.split('.').pop() || '').toLowerCase();
  const buffer = bufferDesdeBase64(archivo.base64);
  if (['csv', 'xlsx', 'xls'].includes(ext)) return { kind: 'tabular', nombre, ...parseTabular(buffer, nombre) };
  if (ext === 'pdf') return { kind: 'texto', nombre, ...(await parsePDF(buffer, nombre)) };
  if (['txt', 'md', 'json'].includes(ext)) return { kind: 'texto', nombre, texto: buffer.toString('utf-8').slice(0, MAX_TEXTO_LARGO) };
  return { kind: 'desconocido', nombre, error: `Formato de "${nombre}" no reconocido — probá CSV, XLSX, PDF, TXT o MD.` };
}

const ALIAS = {
  clienteId: ['cliente_id', 'customer_id', 'id_cliente', 'id_socio', 'member_id', 'socio', 'cliente', 'customer', 'email', 'telefono', 'phone', 'dni'],
  fecha: ['fecha', 'date', 'fecha_compra', 'order_date', 'purchase_date', 'fecha_venta', 'fecha_evento'],
  producto: ['producto', 'product', 'item', 'categoria', 'categoria_producto', 'sku', 'linea', 'descripcion'],
  canal: ['canal', 'channel', 'tienda', 'store', 'retailer', 'punto_venta', 'plataforma'],
  monto: ['monto', 'total', 'amount', 'importe', 'venta', 'revenue', 'precio', 'total_venta', 'valor'],
  nombreEvento: ['evento', 'campana', 'campaña', 'nombre_evento', 'actividad', 'nombre', 'accion', 'acción']
};

function normalizar(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

function detectarColumnas(headers) {
  const norm = headers.map(h => ({ original: h, norm: normalizar(h) }));
  const mapping = {};
  for (const campo of Object.keys(ALIAS)) {
    const candidatos = ALIAS[campo];
    const match = norm.find(h => candidatos.includes(h.norm)) || norm.find(h => candidatos.some(c => h.norm.includes(c)));
    if (match) mapping[campo] = match.original;
  }
  return mapping;
}

function numeroDesde(valor) {
  if (typeof valor === 'number') return valor;
  const limpio = String(valor || '').replace(/[^\d.,-]/g, '').replace(/\.(?=\d{3},)/g, '').replace(',', '.');
  const n = parseFloat(limpio);
  return isNaN(n) ? 0 : n;
}

function parseFechaValor(f) {
  if (f === null || f === undefined || f === '') return null;
  if (typeof f === 'number' && isFinite(f)) {
    const ms = Math.round((f - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(f);
  return isNaN(d.getTime()) ? null : d;
}

/* ================= CALENDARIO COMERCIAL POR PAÍS ================= */

// n-ésimo día de la semana de un mes. weekday: 0=domingo..6=sábado. n: 1,2,3...
function nthWeekdayOfMonth(year, month, weekday, n) {
  const first = new Date(Date.UTC(year, month - 1, 1));
  const offset = (weekday - first.getUTCDay() + 7) % 7;
  const day = 1 + offset + (n - 1) * 7;
  return new Date(Date.UTC(year, month - 1, day));
}
function fechaFija(year, month, day) { return new Date(Date.UTC(year, month - 1, day)); }

// Reglas verificadas contra fuentes reales para 2026 (Colombia, Perú, México,
// Chile). El resto son aproximaciones razonables — el agente permite subir
// un calendario propio para sumar o corregir fechas puntuales.
const CALENDARIO_BASE = {
  CO: [
    { nombre: 'Día de la Madre', regla: y => nthWeekdayOfMonth(y, 5, 0, 2) },
    { nombre: 'Día del Padre', regla: y => nthWeekdayOfMonth(y, 6, 0, 3) },
    { nombre: 'Día del Amor y la Amistad', regla: y => nthWeekdayOfMonth(y, 9, 6, 3) },
    { nombre: 'Halloween', regla: y => fechaFija(y, 10, 31) },
    { nombre: 'Black Friday', regla: y => nthWeekdayOfMonth(y, 11, 5, 4) },
    { nombre: 'Navidad', regla: y => fechaFija(y, 12, 25) },
    { nombre: 'Año Nuevo', regla: y => fechaFija(y, 1, 1) }
  ],
  PE: [
    { nombre: 'Día de la Madre', regla: y => nthWeekdayOfMonth(y, 5, 0, 2) },
    { nombre: 'Día del Padre', regla: y => nthWeekdayOfMonth(y, 6, 0, 3) },
    { nombre: 'Fiestas Patrias', regla: y => fechaFija(y, 7, 28) },
    { nombre: 'Halloween', regla: y => fechaFija(y, 10, 31) },
    { nombre: 'Black Friday', regla: y => nthWeekdayOfMonth(y, 11, 5, 4) },
    { nombre: 'Navidad', regla: y => fechaFija(y, 12, 25) }
  ],
  MX: [
    { nombre: 'Día del Amor y la Amistad', regla: y => fechaFija(y, 2, 14) },
    { nombre: 'Día de las Madres', regla: y => fechaFija(y, 5, 10) },
    { nombre: 'Día del Padre', regla: y => nthWeekdayOfMonth(y, 6, 0, 3) },
    { nombre: 'Día de Muertos', regla: y => fechaFija(y, 11, 2) },
    { nombre: 'Buen Fin (aprox.)', regla: y => nthWeekdayOfMonth(y, 11, 5, 3) },
    { nombre: 'Navidad', regla: y => fechaFija(y, 12, 25) }
  ],
  AR: [
    { nombre: 'Día de la Madre', regla: y => nthWeekdayOfMonth(y, 10, 0, 3) },
    { nombre: 'Día del Padre', regla: y => nthWeekdayOfMonth(y, 6, 0, 3) },
    { nombre: 'Black Friday', regla: y => nthWeekdayOfMonth(y, 11, 5, 4) },
    { nombre: 'Navidad', regla: y => fechaFija(y, 12, 25) }
  ],
  CL: [
    { nombre: 'Día de la Madre', regla: y => nthWeekdayOfMonth(y, 5, 0, 2) },
    { nombre: 'Día del Padre', regla: y => nthWeekdayOfMonth(y, 6, 0, 3) },
    { nombre: 'Black Friday', regla: y => nthWeekdayOfMonth(y, 11, 5, 4) },
    { nombre: 'Navidad', regla: y => fechaFija(y, 12, 25) }
  ],
  GENERIC: [
    { nombre: 'Día de la Madre (genérico)', regla: y => nthWeekdayOfMonth(y, 5, 0, 2) },
    { nombre: 'Black Friday', regla: y => nthWeekdayOfMonth(y, 11, 5, 4) },
    { nombre: 'Navidad', regla: y => fechaFija(y, 12, 25) }
  ]
};

function normalizarPais(pais) {
  const p = normalizar(pais);
  if (p.includes('colombia')) return 'CO';
  if (p.includes('peru') || p.includes('perú')) return 'PE';
  if (p.includes('mexico') || p.includes('méxico')) return 'MX';
  if (p.includes('argentina')) return 'AR';
  if (p.includes('chile')) return 'CL';
  return 'GENERIC';
}

function calendarioDelAnio(paisCode, year, extra = []) {
  const base = CALENDARIO_BASE[paisCode] || CALENDARIO_BASE.GENERIC;
  return [...base, ...extra].map(e => ({ nombre: e.nombre, regla: e.regla, fecha: e.regla(year) }));
}

function proximasFechasComerciales(paisCode, hoy, mesesAdelante, extra = []) {
  const limite = new Date(hoy.getTime() + mesesAdelante * 30.4 * 86400000);
  const candidatas = [...calendarioDelAnio(paisCode, hoy.getUTCFullYear(), extra), ...calendarioDelAnio(paisCode, hoy.getUTCFullYear() + 1, extra)];
  return candidatas
    .filter(f => f.fecha >= hoy && f.fecha <= limite)
    .sort((a, b) => a.fecha - b.fecha)
    .slice(0, MAX_FECHAS);
}

function todasLasFechasDelAnio(paisCode, hoy, extra = []) {
  return [...calendarioDelAnio(paisCode, hoy.getUTCFullYear(), extra), ...calendarioDelAnio(paisCode, hoy.getUTCFullYear() + 1, extra)];
}

/* ================= MOTOR 1: HISTORIAL DE VENTAS → UPLIFT REAL ================= */

function calcularUpliftParaFecha(ventasFilas, entrada, anioObjetivo) {
  const aniosDisponibles = [...new Set(ventasFilas.map(v => v.fecha.getUTCFullYear()))].filter(y => y < anioObjetivo);
  if (!aniosDisponibles.length) return { uplift: null, confianza: 'baja', aniosUsados: 0 };

  const upliftsPorAnio = [];
  for (const y of aniosDisponibles) {
    const fechaY = entrada.regla(y);
    const ventanaIni = fechaY.getTime() - 3 * 86400000;
    const ventanaFin = fechaY.getTime() + 3 * 86400000;
    const ventasAnio = ventasFilas.filter(v => v.fecha.getUTCFullYear() === y);
    if (ventasAnio.length < 5) continue;
    const enVentana = ventasAnio.filter(v => v.fecha.getTime() >= ventanaIni && v.fecha.getTime() <= ventanaFin);
    if (!enVentana.length) continue;
    const totalVentana = enVentana.reduce((s, v) => s + v.monto, 0);

    // Baseline: promedio de ventanas de 7 días repartidas en el resto del
    // año (una por mes, salteando el mes del evento) — no un promedio
    // diario plano. Con facturación no diaria (habitual en B2B), dividir
    // el total entre todos los días del calendario infla artificialmente
    // el % de pico porque diluye el promedio con días sin facturación.
    const ventanasBase = [];
    for (let mes = 0; mes < 12; mes++) {
      const centro = new Date(Date.UTC(y, mes, 15)).getTime();
      if (Math.abs(centro - fechaY.getTime()) < 20 * 86400000) continue;
      const ini = centro - 3 * 86400000, fin = centro + 3 * 86400000;
      const enVentanaBase = ventasAnio.filter(v => v.fecha.getTime() >= ini && v.fecha.getTime() <= fin);
      if (enVentanaBase.length) ventanasBase.push(enVentanaBase.reduce((s, v) => s + v.monto, 0));
    }
    if (ventanasBase.length < 3) continue;
    const baseline = ventanasBase.reduce((s, v) => s + v, 0) / ventanasBase.length;
    if (baseline <= 0) continue;
    upliftsPorAnio.push((totalVentana - baseline) / baseline);
  }
  if (!upliftsPorAnio.length) return { uplift: null, confianza: 'baja', aniosUsados: 0 };
  const prom = upliftsPorAnio.reduce((s, u) => s + u, 0) / upliftsPorAnio.length;
  // Tope de sanidad: una variación extrema (>500%) suele indicar muestra
  // demasiado chica o dispersa más que una señal real — se baja la
  // confianza en vez de mostrar un número poco creíble en el reporte.
  const extremo = Math.abs(prom) > 5;
  const upliftFinal = Math.max(-0.9, Math.min(5, prom));
  const confianza = extremo ? 'baja' : (upliftsPorAnio.length >= 2 ? 'alta' : 'media');
  return { uplift: Math.round(upliftFinal * 1000) / 10, confianza, aniosUsados: upliftsPorAnio.length };
}

/* ================= PRECISIÓN DEL PRONÓSTICO ================= */

function calcularPrecision({ aniosVentas, usoCalendarioMarca, usoDatosClienteFinal, usoReportesMercado }) {
  let score = 30 + Math.min(aniosVentas, 3) * 15;
  if (usoCalendarioMarca) score += 10;
  if (usoDatosClienteFinal) score += 10;
  if (usoReportesMercado) score += 5;
  return Math.min(score, 95);
}

function sugerenciaMejora({ aniosVentas, usoCalendarioMarca, usoDatosClienteFinal, usoReportesMercado }) {
  const sugerencias = [];
  if (aniosVentas < 2) {
    const sumar = aniosVentas === 0 ? 2 : 1;
    const scoreConMas = calcularPrecision({ aniosVentas: Math.min(aniosVentas + sumar, 3), usoCalendarioMarca, usoDatosClienteFinal, usoReportesMercado });
    sugerencias.push(`Sumar ${sumar} año${sumar > 1 ? 's' : ''} más de historial de ventas podría subir la precisión a ~${scoreConMas}%.`);
  }
  if (!usoDatosClienteFinal) sugerencias.push('Sumar datos de cliente final (los mismos de Agente F) activaría el radar de fuga.');
  if (!usoCalendarioMarca) sugerencias.push('Sumar tu calendario propio de campañas permitiría detectar cruces y oportunidades de marca.');
  if (!sugerencias.length) sugerencias.push('Ya estás usando todas las fuentes disponibles — la precisión sube principalmente con más años de historial.');
  return sugerencias.join(' ');
}

/* ================= CRUCES: CALENDARIO PROPIO × MERCADO ================= */

function calcularCruces(eventosMarca, fechasComercialesDelAnio) {
  return eventosMarca.map((ev, i) => {
    let mejor = null, mejorDist = Infinity;
    for (const fc of fechasComercialesDelAnio) {
      const dist = Math.abs((fc.fecha.getTime() - ev.fecha.getTime()) / 86400000);
      if (dist < mejorDist) { mejorDist = dist; mejor = fc; }
    }
    const base = { id: 'cruce_' + i, eventoMarca: ev.nombre, fechaMarca: ev.fecha.toISOString().slice(0, 10) };
    if (mejor && mejorDist <= 21) {
      return { ...base, tipo: 'coincidencia', nombreMercado: mejor.nombre, fechaMercado: mejor.fecha.toISOString().slice(0, 10), diasEntre: Math.round(mejorDist) };
    }
    return { ...base, tipo: 'propia' };
  });
}

/* ================= MOTOR 2: RADAR DE FUGA (CHURN) ================= */

function agregarClientesParaChurn(archivosClienteFinal) {
  const porPersona = {};
  for (const archivo of archivosClienteFinal) {
    if (!archivo.mapping || !archivo.mapping.clienteId || !archivo.mapping.fecha) continue;
    for (const fila of archivo.filas) {
      const id = normalizar(fila[archivo.mapping.clienteId]);
      const fecha = parseFechaValor(fila[archivo.mapping.fecha]);
      if (!id || !fecha) continue;
      const monto = archivo.mapping.monto ? numeroDesde(fila[archivo.mapping.monto]) : 0;
      if (!porPersona[id]) porPersona[id] = { id, compras: [] };
      porPersona[id].compras.push({ fecha: fecha.getTime(), monto });
    }
  }
  const personas = Object.values(porPersona);
  if (!personas.length) return null;

  const hoy = Date.now();
  const perfiles = personas.map(p => {
    const ordenadas = p.compras.slice().sort((a, b) => a.fecha - b.fecha);
    const ultima = ordenadas[ordenadas.length - 1].fecha;
    const diasDesdeUltima = Math.round((hoy - ultima) / 86400000);
    let intervaloPromedio = null;
    if (ordenadas.length >= 3) {
      const intervalos = [];
      for (let i = 1; i < ordenadas.length; i++) intervalos.push((ordenadas[i].fecha - ordenadas[i - 1].fecha) / 86400000);
      intervaloPromedio = Math.round(intervalos.reduce((s, x) => s + x, 0) / intervalos.length);
    }
    let ticketBajando = false;
    if (ordenadas.length >= 4) {
      const mitad = Math.floor(ordenadas.length / 2);
      const avg = arr => arr.reduce((s, c) => s + c.monto, 0) / (arr.length || 1);
      const avgPrimera = avg(ordenadas.slice(0, mitad));
      const avgSegunda = avg(ordenadas.slice(mitad));
      ticketBajando = avgPrimera > 0 && avgSegunda < avgPrimera * 0.8;
    }
    const enRiesgo = !!(intervaloPromedio && diasDesdeUltima > intervaloPromedio * 2.5);
    return { id: p.id, compras: ordenadas.length, diasDesdeUltima, intervaloPromedio, enRiesgo, ticketBajando };
  });

  const totalPersonas = perfiles.length;
  const enRiesgo = perfiles.filter(p => p.enRiesgo);
  const ticketBajando = perfiles.filter(p => p.ticketBajando && !p.enRiesgo);

  return {
    totalPersonas,
    enRiesgoCount: enRiesgo.length,
    ticketBajandoCount: ticketBajando.length,
    muestraEnRiesgo: enRiesgo.slice(0, 8),
    muestraTicketBajando: ticketBajando.slice(0, 8)
  };
}

/* ================= CALENDARIO DE PAÍS PERSONALIZADO (editable) ================= */

// Fechas puntuales que el cliente conoce (feriados regionales, promos
// internas fijas) y que el calendario base no trae. Se tratan como fecha
// fija recurrente cada año (mismo mes/día), sumadas a las del calendario
// base — nunca lo reemplazan.
function parseCalendarioPersonalizado(filas, mapping) {
  const eventos = [];
  if (!mapping.fecha) return eventos;
  for (const fila of filas) {
    const fecha = parseFechaValor(fila[mapping.fecha]);
    if (!fecha) continue;
    const nombre = mapping.nombreEvento ? String(fila[mapping.nombreEvento] || '').trim() : 'Fecha propia del país';
    const mes = fecha.getUTCMonth(), dia = fecha.getUTCDate();
    eventos.push({ nombre: (nombre || 'Fecha propia del país') + ' (personalizada)', regla: y => new Date(Date.UTC(y, mes, dia)) });
  }
  return eventos;
}

/* ================= PROMPT Y LLAMADA A CLAUDE ================= */

function construirPrompt(datosCliente, fechasConUplift, cruces, churn, documentosTexto) {
  const { cliente, pais, categoria, canalesVenta, canalesComunicacion } = datosCliente;
  const bloques = [];

  bloques.push(`CLIENTE: ${cliente || 'sin nombre'} · PAÍS: ${pais || 'no especificado'} · CATEGORÍA: ${categoria || 'no especificada'}\nCANALES DE VENTA: ${canalesVenta || 'no especificados'} · CANALES DE COMUNICACIÓN: ${canalesComunicacion || 'no especificados'}`);

  bloques.push(`\n=== PRÓXIMAS FECHAS COMERCIALES (ya calculadas — vos NO calculás el % de uplift ni la fecha) ===\n${JSON.stringify(fechasConUplift.map(f => ({ id: f.id, nombre: f.nombre, fecha: f.fechaISO, diasFaltantes: f.diasFaltantes, upliftPct: f.uplift, confianza: f.confianza, aniosDeHistorialUsados: f.aniosUsados })))}`);

  if (cruces.length) {
    bloques.push(`\n=== CRUCES ENTRE CALENDARIO PROPIO DE LA MARCA Y CALENDARIO COMERCIAL (ya calculados) ===\n${JSON.stringify(cruces)}`);
  }

  if (churn) {
    bloques.push(`\n=== RADAR DE FUGA — YA CALCULADO SOBRE DATOS REALES DE CLIENTE FINAL ===
Total de personas identificadas: ${churn.totalPersonas}
En riesgo de fuga (intervalo de recompra habitual superado por más de 2.5x): ${churn.enRiesgoCount}
Con ticket promedio bajando (sin señal de fuga aún): ${churn.ticketBajandoCount}
Muestra en riesgo: ${JSON.stringify(churn.muestraEnRiesgo)}
Muestra con ticket bajando: ${JSON.stringify(churn.muestraTicketBajando)}`);
  }

  if (documentosTexto.length) {
    bloques.push(`\n=== DOCUMENTOS DE CONTEXTO (reportes de mercado, texto/PDF) ===`);
    documentosTexto.forEach(d => bloques.push(`--- ${d.fuente} (${d.nombre}) ---\n${d.texto}`));
  }

  return bloques.join('\n');
}

const SYSTEM_PROMPT = `Sos el motor de redacción del Agente G (Predictive Growth Engine) de LiftyFive, una agencia de retail media AI-first.

Todos los números que recibís (fechas, % de uplift, días entre fechas, conteos de clientes en riesgo) YA fueron calculados de forma determinística — vos NUNCA calculás ni corregís un número. Tu única tarea es escribir, sobre esa evidencia real:

1. Para cada fecha comercial: accionables concretos agrupados en "stock_logistica", "comunicacion", "ventas" y "logistica" (máximo 2 ítems cada uno, 1 oración cada ítem). Si la fecha tiene confianza "baja" (sin historial propio), los accionables tienen que ser más generales y decir explícitamente que no hay historial propio todavía para esa fecha puntual — nunca inventes un % o una certeza que no existe.
2. Para cada cruce entre calendario propio y calendario comercial: una "oportunidad" (1-2 oraciones) que aproveche específicamente esa coincidencia o esa fecha propia — nunca genérica, tiene que nombrar los dos eventos si es una coincidencia.
3. Si hay radar de fuga: hasta 2 "grupos" de clientes en riesgo, cada uno con una "descripcion" (qué los distingue, basado en los datos reales) y una "accion" concreta de retención.
4. Un "resumen_ejecutivo" de 2-3 oraciones sobre el panorama general de los próximos meses.

Reglas estrictas:
- Nunca inventes un porcentaje, una fecha o un conteo — todo sale de los datos que te paso.
- Si una fecha no tiene historial propio (confianza "baja"), decilo explícitamente en sus accionables en vez de sugerir algo con falsa certeza.
- No devuelvas texto fuera del JSON. Respondé ÚNICAMENTE con un JSON válido, sin bloques de markdown, con este schema exacto:

{
  "resumen_ejecutivo": "string",
  "fechas": [
    { "id": "string (igual al id que te paso)", "accionables": { "stock_logistica": ["string"], "comunicacion": ["string"], "ventas": ["string"], "logistica": ["string"] } }
  ],
  "cruces": [
    { "id": "string (igual al id que te paso)", "oportunidad": "string" }
  ],
  "churn_grupos": [
    { "nombre": "string corto", "cantidad": number, "descripcion": "string", "accion": "string" }
  ]
}`;

async function llamarClaude(prompt, apiKey) {
  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
    body: JSON.stringify({ model: 'claude-sonnet-5', max_tokens: 12000, system: SYSTEM_PROMPT, messages: [{ role: 'user', content: prompt }] })
  });
  const data = await r.json();
  if (!r.ok) throw new Error(data?.error?.message || 'Error llamando a la API de Claude.');

  const bloqueTexto = (data?.content || []).find(b => b.type === 'text');
  const texto = bloqueTexto?.text || '';
  if (!texto) {
    console.error('Claude respondió sin bloque de texto. stop_reason:', data?.stop_reason);
    throw new Error(`Claude no devolvió texto en la respuesta (stop_reason: ${data?.stop_reason || 'desconocido'}).`);
  }
  const limpio = texto.trim().replace(/^```(json)?/i, '').replace(/```$/, '').trim();
  try {
    return JSON.parse(limpio);
  } catch (e) {
    console.error('JSON inválido de Claude. stop_reason:', data?.stop_reason, 'respuesta completa:', texto);
    throw new Error(`Claude no devolvió un JSON válido (stop_reason: ${data?.stop_reason || 'desconocido'}). Respuesta cruda: ` + (limpio.slice(0, 500) || '(vacía)'));
  }
}

/* ================= HANDLER ================= */

module.exports = async (req, res) => {
  if (req.method !== 'POST') { res.status(405).json({ error: 'Método no permitido. Usá POST.' }); return; }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) { res.status(500).json({ error: 'Falta configurar ANTHROPIC_API_KEY en las variables de entorno de Vercel.' }); return; }

  let body;
  try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
  catch (e) { res.status(400).json({ error: 'Body inválido.' }); return; }

  const { cliente, pais, categoria, canalesVenta, canalesComunicacion, archivos = [] } = body || {};

  if (!pais || !pais.trim()) {
    res.status(400).json({ error: 'Necesito al menos el país para poder anclar el calendario comercial correcto.' });
    return;
  }

  try {
    const paisCode = normalizarPais(pais);
    const hoy = new Date(Date.UTC(new Date().getUTCFullYear(), new Date().getUTCMonth(), new Date().getUTCDate()));

    const ventasFiles = archivos.filter(a => a.tipo === 'ventasPropio');
    const calendarioMarcaFiles = archivos.filter(a => a.tipo === 'calendarioMarca');
    const calendarioPaisFiles = archivos.filter(a => a.tipo === 'calendarioPais');
    const clienteFinalFiles = archivos.filter(a => a.tipo === 'datosClienteFinal');
    const reportesMercadoFiles = archivos.filter(a => a.tipo === 'reportesMercado');

    // --- Calendario propio del país: fechas fijas que se suman al base ---
    let fechasPersonalizadas = [];
    for (const archivo of calendarioPaisFiles) {
      const parsed = await parseArchivo(archivo);
      if (parsed.kind !== 'tabular' || parsed.error) continue;
      const mapping = detectarColumnas(parsed.headers);
      fechasPersonalizadas.push(...parseCalendarioPersonalizado(parsed.filas, mapping));
    }

    // --- Ventas propio: combinar todos los archivos en una sola lista de filas ---
    let ventasFilas = [];
    let ventasFuentesResumen = [];
    for (const archivo of ventasFiles) {
      const parsed = await parseArchivo(archivo);
      if (parsed.kind !== 'tabular' || parsed.error) { ventasFuentesResumen.push({ nombre: archivo.nombre, usada: false, detalle: parsed.error || 'Formato no soportado para ventas.' }); continue; }
      const mapping = detectarColumnas(parsed.headers);
      if (!mapping.fecha || !mapping.monto) { ventasFuentesResumen.push({ nombre: archivo.nombre, usada: false, detalle: 'No se detectaron columnas de fecha y monto.' }); continue; }
      let usadas = 0;
      for (const fila of parsed.filas) {
        const fecha = parseFechaValor(fila[mapping.fecha]);
        if (!fecha) continue;
        ventasFilas.push({ fecha, monto: numeroDesde(fila[mapping.monto]), producto: mapping.producto ? String(fila[mapping.producto] || '').trim() : '' });
        usadas++;
      }
      ventasFuentesResumen.push({ nombre: archivo.nombre, usada: usadas > 0, detalle: `${usadas} filas con fecha y monto válidos.` });
    }
    const aniosVentas = new Set(ventasFilas.map(v => v.fecha.getUTCFullYear())).size;

    // --- Fechas comerciales próximas + uplift real ---
    const proximas = proximasFechasComerciales(paisCode, hoy, MESES_ADELANTE, fechasPersonalizadas);
    const fechasConUplift = proximas.map((f, i) => {
      const { uplift, confianza, aniosUsados } = calcularUpliftParaFecha(ventasFilas, f, f.fecha.getUTCFullYear());
      return {
        id: 'fecha_' + i,
        nombre: f.nombre,
        fechaISO: f.fecha.toISOString().slice(0, 10),
        diasFaltantes: Math.round((f.fecha.getTime() - hoy.getTime()) / 86400000),
        uplift, confianza, aniosUsados
      };
    });

    // --- Calendario propio de marca + cruces ---
    let eventosMarca = [];
    for (const archivo of calendarioMarcaFiles) {
      const parsed = await parseArchivo(archivo);
      if (parsed.kind !== 'tabular' || parsed.error) continue;
      const mapping = detectarColumnas(parsed.headers);
      if (!mapping.fecha) continue;
      for (const fila of parsed.filas) {
        const fecha = parseFechaValor(fila[mapping.fecha]);
        if (!fecha) continue;
        const nombre = mapping.nombreEvento ? String(fila[mapping.nombreEvento] || '').trim() : (mapping.producto ? String(fila[mapping.producto] || '').trim() : 'Evento propio');
        eventosMarca.push({ nombre: nombre || 'Evento propio', fecha });
      }
    }
    const fechasComercialesDelAnio = todasLasFechasDelAnio(paisCode, hoy, fechasPersonalizadas);
    const cruces = calcularCruces(eventosMarca, fechasComercialesDelAnio);

    // --- Churn ---
    let archivosClienteFinalParseados = [];
    for (const archivo of clienteFinalFiles) {
      const parsed = await parseArchivo(archivo);
      if (parsed.kind !== 'tabular' || parsed.error) continue;
      const mapping = detectarColumnas(parsed.headers);
      archivosClienteFinalParseados.push({ filas: parsed.filas, mapping });
    }
    const churn = archivosClienteFinalParseados.length ? agregarClientesParaChurn(archivosClienteFinalParseados) : null;

    // --- Reportes de mercado (texto/PDF) ---
    let documentosTexto = [];
    for (const archivo of reportesMercadoFiles) {
      const parsed = await parseArchivo(archivo);
      if (parsed.kind === 'texto' && parsed.texto) documentosTexto.push({ fuente: 'Reporte de mercado', nombre: parsed.nombre, texto: parsed.texto });
    }

    if (!ventasFilas.length && !eventosMarca.length && !churn && !documentosTexto.length) {
      // Igual seguimos: el calendario base solo ya da valor, pero avisamos que es 100% genérico.
    }

    const prompt = construirPrompt({ cliente, pais, categoria, canalesVenta, canalesComunicacion }, fechasConUplift, cruces, churn, documentosTexto);
    const resultado = await llamarClaude(prompt, apiKey);

    // Merge determinístico + redacción de Claude
    const accionablesPorId = {}; (resultado.fechas || []).forEach(f => { accionablesPorId[f.id] = f.accionables; });
    const fechasFinal = fechasConUplift.map(f => ({ ...f, accionables: accionablesPorId[f.id] || { stock_logistica: [], comunicacion: [], ventas: [], logistica: [] } }));

    const oportunidadPorId = {}; (resultado.cruces || []).forEach(c => { oportunidadPorId[c.id] = c.oportunidad; });
    const crucesFinal = cruces.map(c => ({ ...c, oportunidad: oportunidadPorId[c.id] || '' }));

    const precisionScore = calcularPrecision({
      aniosVentas, usoCalendarioMarca: eventosMarca.length > 0, usoDatosClienteFinal: !!churn, usoReportesMercado: documentosTexto.length > 0
    });
    const precisionDetalle = `Se usó${aniosVentas ? ` ${aniosVentas} año(s) de historial de ventas` : ' calendario base sin historial propio'}${churn ? ', datos de cliente final' : ''}${eventosMarca.length ? ', calendario propio de marca' : ''}${documentosTexto.length ? ' y reportes de mercado' : ''}.`;
    const comoMejorar = sugerenciaMejora({ aniosVentas, usoCalendarioMarca: eventosMarca.length > 0, usoDatosClienteFinal: !!churn, usoReportesMercado: documentosTexto.length > 0 });

    res.status(200).json({
      resumen_ejecutivo: resultado.resumen_ejecutivo || '',
      pais: paisCode,
      precision: { score: precisionScore, detalle: precisionDetalle, comoMejorar },
      fechas: fechasFinal,
      cruces: crucesFinal,
      churn: churn ? { totalPersonas: churn.totalPersonas, enRiesgoCount: churn.enRiesgoCount, ticketBajandoCount: churn.ticketBajandoCount, grupos: resultado.churn_grupos || [] } : null
    });
  } catch (err) {
    console.error('Error en /api/forecast:', err);
    res.status(500).json({ error: 'Error inesperado construyendo el pronóstico: ' + err.message });
  }
};

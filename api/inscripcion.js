// api/inscripcion.js — Vercel Serverless Function

const TENANT_ID       = process.env.TENANT_ID;
const CLIENT_ID       = process.env.CLIENT_ID;
const CLIENT_SECRET   = process.env.CLIENT_SECRET;
const SHAREPOINT_HOST = 'miltonochoacol.sharepoint.com';
const SITE_PATH       = '/sites/Programacin';
const EMAIL_USER      = process.env.EMAIL_USER;
const EMAIL_PASS      = process.env.EMAIL_PASS;

async function getAppToken() {
  const res  = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    body: new URLSearchParams({
      client_id: CLIENT_ID, client_secret: CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials',
    }),
  });
  const data = await res.json();
  if (!data.access_token) throw new Error('Token error: ' + JSON.stringify(data));
  return data.access_token;
}

async function getSiteId(token) {
  const res  = await fetch(`https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SITE_PATH}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await res.json();
  if (!data.id) throw new Error('Sitio no encontrado');
  return data.id;
}

async function getDriveId(token, siteId) {
  const res  = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await res.json();
  if (!data.id) throw new Error('Drive no encontrado');
  return data.id;
}

async function findFileId(token, siteId, driveId) {
  const paths = [
    'Saber Delta/Inscripciones_Delta.xlsx',
    'Saber Delta/Inscripciones_Delta.xlsx.xlsx',
    'Inscripciones_Delta.xlsx',
  ];
  for (const path of paths) {
    const res  = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${path}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const data = await res.json();
    if (data.id) { console.log('Archivo encontrado en:', path); return data.id; }
  }
  throw new Error('Archivo Excel no encontrado');
}

async function appendRow(token, siteId, driveId, fileId, inscrito) {
  const base    = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${fileId}/workbook`;
  const HEADERS = ['Fecha', 'Nombre', 'Ciudad', 'Institución', 'WhatsApp', 'Correo'];
  const row     = [inscrito.fecha, inscrito.nombre, inscrito.ciudad, inscrito.institucion, inscrito.whatsapp, inscrito.correo];

  // Leer columna A para encontrar primera fila vacía
  const colRes  = await fetch(`${base}/worksheets('Hoja1')/range(address='A1:A500')`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const colData = await colRes.json();

  let nextRow = 1;
  if (colData.values) {
    for (let i = 0; i < colData.values.length; i++) {
      if (!colData.values[i][0] || colData.values[i][0] === '') {
        nextRow = i + 1;
        break;
      }
      nextRow = i + 2;
    }
  }

  // Si primera fila está vacía, crear encabezados
  if (nextRow === 1) {
    await fetch(`${base}/worksheets('Hoja1')/range(address='A1:F1')`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: [HEADERS] }),
    });
    nextRow = 2;
  }

  // Escribir datos en siguiente fila vacía
  const writeRes = await fetch(`${base}/worksheets('Hoja1')/range(address='A${nextRow}:F${nextRow}')`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [row] }),
  });

  if (!writeRes.ok) throw new Error('Error escribiendo fila: ' + await writeRes.text());
  console.log('Fila escrita en:', nextRow);
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  try {
    const { nombre, ciudad, institucion, whatsapp, correo, fecha } = req.body;
    if (!nombre || !ciudad || !institucion || !whatsapp || !correo)
      return res.status(400).json({ error: 'Faltan campos' });

    const inscrito = { nombre, ciudad, institucion, whatsapp, correo, fecha: fecha || new Date().toISOString() };

    const appToken = await getAppToken();
    const siteId   = await getSiteId(appToken);
    const driveId  = await getDriveId(appToken, siteId);
    const fileId   = await findFileId(appToken, siteId, driveId);

    await appendRow(appToken, siteId, driveId, fileId, inscrito);

    // Correo pendiente de configurar
    console.log('Inscripción guardada:', inscrito.nombre);

    return res.status(200).json({ ok: true, message: 'Inscripción guardada exitosamente' });

  } catch (err) {
    console.error('Error:', err.message);
    return res.status(500).json({ ok: false, error: err.message });
  }
}

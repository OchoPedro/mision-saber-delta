// api/inscripcion.js — Vercel Serverless Function
// Guarda inscripción en SharePoint/Excel y envía correo de confirmación

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SHAREPOINT_HOST = 'miltonochoacol.sharepoint.com';
const SITE_PATH       = '/sites/Programacin';
const FROM_EMAIL      = 'Programacion@aamocolombia.com';

// ── Obtener token ─────────────────────────────────────────────────────────────
async function getToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id:     CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
    grant_type:    'client_credentials',
  });
  const res  = await fetch(url, { method: 'POST', body });
  const data = await res.json();
  if (!data.access_token) throw new Error('Token error: ' + JSON.stringify(data));
  return data.access_token;
}

// ── Obtener Site ID ───────────────────────────────────────────────────────────
async function getSiteId(token) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SITE_PATH}`;
  const res  = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await res.json();
  if (!data.id) throw new Error('Sitio no encontrado: ' + JSON.stringify(data));
  return data.id;
}

// ── Obtener Drive principal ───────────────────────────────────────────────────
async function getDefaultDriveId(token, siteId) {
  const url  = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`;
  const res  = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await res.json();
  if (!data.id) throw new Error('Drive no encontrado: ' + JSON.stringify(data));
  return data.id;
}

// ── Buscar archivo Excel por nombre en la carpeta Saber Delta ─────────────────
async function findFileId(token, siteId, driveId) {
  // Intentar ruta exacta primero
  const paths = [
    'Saber Delta/Inscripciones_Delta.xlsx',
    'Saber Delta/Inscripciones_Delta.xlsx.xlsx',
    'General/Saber Delta/Inscripciones_Delta.xlsx',
    'Inscripciones_Delta.xlsx',
  ];

  for (const path of paths) {
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${path}`;
    const res  = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    const data = await res.json();
    if (data.id) {
      console.log('Archivo encontrado en:', path);
      return { fileId: data.id, driveId };
    }
  }

  // Si no encontró, listar carpetas para debug
  const rootRes  = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const rootData = await rootRes.json();
  const names    = (rootData.value || []).map(i => i.name).join(', ');
  throw new Error(`Archivo no encontrado. Carpetas raíz: [${names}]`);
}

// ── Escribir fila en Excel ────────────────────────────────────────────────────
async function appendRow(token, siteId, driveId, fileId, inscrito) {
  const base = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${fileId}/workbook`;

  // Verificar tablas
  const tRes  = await fetch(`${base}/tables`, { headers: { Authorization: `Bearer ${token}` } });
  const tData = await tRes.json();

  const row = [inscrito.fecha, inscrito.nombre, inscrito.ciudad, inscrito.institucion, inscrito.whatsapp, inscrito.correo];

  if (tData.value && tData.value.length > 0) {
    // Agregar a tabla existente
    const tid    = tData.value[0].id;
    const addRes = await fetch(`${base}/tables/${tid}/rows`, {
      method:  'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body:    JSON.stringify({ values: [row] }),
    });
    if (!addRes.ok) throw new Error('Error tabla: ' + await addRes.text());
  } else {
    // Sin tabla — escribir directamente en celdas
    // Obtener rango usado
    const uRes  = await fetch(`${base}/worksheets('Hoja1')/usedRange`, { headers: { Authorization: `Bearer ${token}` } });
    const uData = await uRes.json();
    let nextRow  = uData.rowCount ? uData.rowCount + 1 : 1;

    if (nextRow === 1) {
      // Crear encabezados
      await fetch(`${base}/worksheets('Hoja1')/range(address='A1:F1')`, {
        method:  'PATCH',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body:    JSON.stringify({ values: [['Fecha', 'Nombre', 'Ciudad', 'Institución', 'WhatsApp', 'Correo']] }),
      });
      nextRow = 2;
    }

    const writeRes = await fetch(`${base}/worksheets('Hoja1')/range(address='A${nextRow}:F${nextRow}')`, {
      method:  'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body:    JSON.stringify({ values: [row] }),
    });
    if (!writeRes.ok) throw new Error('Error escribiendo: ' + await writeRes.text());
  }
}

// ── Enviar correo ─────────────────────────────────────────────────────────────
async function sendEmail(token, inscrito) {
  const uRes  = await fetch(`https://graph.microsoft.com/v1.0/users/${FROM_EMAIL}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const uData = await uRes.json();
  if (!uData.id) throw new Error('Remitente no encontrado: ' + JSON.stringify(uData));

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#081623;color:#c3ccd4;padding:40px;border-radius:12px;">
      <div style="text-align:center;margin-bottom:30px;">
        <h1 style="color:#ffffff;font-size:28px;letter-spacing:4px;">🔺 MISIÓN SABER DELTA</h1>
        <p style="color:#8794a1;letter-spacing:2px;font-size:13px;">MILTON OCHOA</p>
      </div>
      <div style="background:rgba(255,255,255,0.05);border:1px solid rgba(150,180,205,0.22);border-radius:10px;padding:25px;margin-bottom:25px;">
        <h2 style="color:#e0b24c;font-size:22px;margin-bottom:10px;">✅ INSCRIPCIÓN CONFIRMADA</h2>
        <p style="font-size:16px;line-height:1.6;">Agente <strong style="color:#ffffff;">${inscrito.nombre}</strong>, tu inscripción al grupo Delta ha sido registrada exitosamente.</p>
      </div>
      <div style="background:rgba(255,255,255,0.03);border:1px solid rgba(150,180,205,0.15);border-radius:10px;padding:20px;margin-bottom:25px;">
        <h3 style="color:#4db3ff;margin-bottom:15px;font-size:14px;letter-spacing:2px;">TUS DATOS DE INSCRIPCIÓN</h3>
        <table style="width:100%;font-size:15px;">
          <tr><td style="color:#8794a1;padding:6px 0;width:120px;">Nombre</td><td style="color:#ffffff;">${inscrito.nombre}</td></tr>
          <tr><td style="color:#8794a1;padding:6px 0;">Ciudad</td><td style="color:#ffffff;">${inscrito.ciudad}</td></tr>
          <tr><td style="color:#8794a1;padding:6px 0;">Institución</td><td style="color:#ffffff;">${inscrito.institucion}</td></tr>
          <tr><td style="color:#8794a1;padding:6px 0;">WhatsApp</td><td style="color:#ffffff;">${inscrito.whatsapp}</td></tr>
          <tr><td style="color:#8794a1;padding:6px 0;">Correo</td><td style="color:#ffffff;">${inscrito.correo}</td></tr>
        </table>
      </div>
      <p style="text-align:center;color:#8794a1;font-size:14px;line-height:1.6;">
        Pronto recibirás más información sobre la misión.<br/>
        Bienvenido a la élite. El profe <strong style="color:#ffffff;">Milton Ochoa</strong> te espera. 🔺
      </p>
      <div style="text-align:center;margin-top:30px;padding-top:20px;border-top:1px solid rgba(150,180,205,0.15);">
        <p style="color:#4db3ff;font-size:12px;letter-spacing:3px;">ASESORÍAS ACADÉMICAS MILTON OCHOA</p>
        <p style="color:#8794a1;font-size:11px;">aamocolombia.com</p>
      </div>
    </div>`;

  const sendRes = await fetch(`https://graph.microsoft.com/v1.0/users/${uData.id}/sendMail`, {
    method:  'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body:    JSON.stringify({
      message: {
        subject: '🔺 MISIÓN SABER DELTA — Inscripción Confirmada',
        body:    { contentType: 'HTML', content: html },
        toRecipients: [{ emailAddress: { address: inscrito.correo, name: inscrito.nombre } }],
      },
      saveToSentItems: true,
    }),
  });

  if (!sendRes.ok && sendRes.status !== 202) {
    throw new Error('Error correo: ' + await sendRes.text());
  }
}

// ── Handler ───────────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')   return res.status(405).json({ error: 'Método no permitido' });

  try {
    const { nombre, ciudad, institucion, whatsapp, correo, fecha } = req.body;
    if (!nombre || !ciudad || !institucion || !whatsapp || !correo)
      return res.status(400).json({ error: 'Faltan campos' });

    const inscrito = { nombre, ciudad, institucion, whatsapp, correo, fecha: fecha || new Date().toISOString() };

    const token   = await getToken();
    const siteId  = await getSiteId(token);
    const driveId = await getDefaultDriveId(token, siteId);
    const { fileId } = await findFileId(token, siteId, driveId);

    await appendRow(token, siteId, driveId, fileId, inscrito);
    await sendEmail(token, inscrito);

    return res.status(200).json({ ok: true });

  } catch (err) {
    console.error('Error:', err.message);
    return res.status(500).json({ ok: false, error: err.message });
  }
}

// api/inscripcion.js — Vercel Serverless Function

const TENANT_ID       = process.env.TENANT_ID;
const CLIENT_ID       = process.env.CLIENT_ID;
const CLIENT_SECRET   = process.env.CLIENT_SECRET;
const SHAREPOINT_HOST = 'miltonochoacol.sharepoint.com';
const SITE_PATH       = '/sites/Programacin';
const FROM_EMAIL      = 'Saber.Delta@aamocolombia.com';
const SENDER_EMAIL    = 'Pedro.Ochoa@aamocolombia.com'; // cuenta real que envía

async function getToken() {
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

  const uRes  = await fetch(`${base}/worksheets('Hoja1')/range(address='A1:A500')`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const uData = await uRes.json();

  let nextRow = 1;
  if (uData.values) {
    for (let i = 0; i < uData.values.length; i++) {
      if (!uData.values[i][0] || uData.values[i][0] === '') { nextRow = i + 1; break; }
      nextRow = i + 2;
    }
  }

  if (nextRow === 1) {
    await fetch(`${base}/worksheets('Hoja1')/range(address='A1:F1')`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: [HEADERS] }),
    });
    nextRow = 2;
  }

  const writeRes = await fetch(`${base}/worksheets('Hoja1')/range(address='A${nextRow}:F${nextRow}')`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [row] }),
  });

  if (!writeRes.ok) throw new Error('Error escribiendo fila: ' + await writeRes.text());
  console.log('Fila escrita en:', nextRow);
}

async function sendEmail(token, inscrito) {
  // Obtener ID del usuario que envía (Pedro.Ochoa)
  const uRes  = await fetch(`https://graph.microsoft.com/v1.0/users/${SENDER_EMAIL}`, {
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

  // Enviar desde Pedro.Ochoa pero mostrar From como Saber.Delta (buzón compartido)
  const sendRes = await fetch(`https://graph.microsoft.com/v1.0/users/${uData.id}/sendMail`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: {
        subject: '🔺 MISIÓN SABER DELTA — Inscripción Confirmada',
        body: { contentType: 'HTML', content: html },
        toRecipients: [{ emailAddress: { address: inscrito.correo, name: inscrito.nombre } }],
        from: { emailAddress: { address: FROM_EMAIL, name: 'Misión Saber Delta · Milton Ochoa' } },
      },
      saveToSentItems: true,
    }),
  });

  if (!sendRes.ok && sendRes.status !== 202) {
    const err = await sendRes.text();
    throw new Error('Error correo: ' + err);
  }
  console.log('Correo enviado a:', inscrito.correo);
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

    const token   = await getToken();
    const siteId  = await getSiteId(token);
    const driveId = await getDriveId(token, siteId);
    const fileId  = await findFileId(token, siteId, driveId);

    await appendRow(token, siteId, driveId, fileId, inscrito);
    await sendEmail(token, inscrito);

    return res.status(200).json({ ok: true, message: 'Inscripción guardada y correo enviado' });

  } catch (err) {
    console.error('Error:', err.message);
    return res.status(500).json({ ok: false, error: err.message });
  }
}

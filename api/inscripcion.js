// api/inscripcion.js — Vercel Serverless Function
// Guarda inscripción en SharePoint/Excel y envía correo de confirmación

const TENANT_ID    = process.env.TENANT_ID;
const CLIENT_ID    = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_NAME    = 'Programacin';   // nombre del sitio SharePoint
const FILE_NAME    = 'Inscripciones_Delta.xlsx';
const FOLDER_PATH  = 'Shared Documents/Saber Delta';
const FROM_EMAIL   = 'Programacion@aamocolombia.com';

// ── Obtener token de Microsoft Graph ─────────────────────────────────────────
async function getToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id:     CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
    grant_type:    'client_credentials',
  });
  const res = await fetch(url, { method: 'POST', body });
  const data = await res.json();
  if (!data.access_token) throw new Error('No se pudo obtener token: ' + JSON.stringify(data));
  return data.access_token;
}

// ── Obtener Site ID de SharePoint ─────────────────────────────────────────────
async function getSiteId(token) {
  const url = `https://graph.microsoft.com/v1.0/sites/miltonochoacol.sharepoint.com:/sites/${SITE_NAME}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await res.json();
  if (!data.id) throw new Error('Site no encontrado: ' + JSON.stringify(data));
  return data.id;
}

// ── Obtener Drive ID ──────────────────────────────────────────────────────────
async function getDriveId(token, siteId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await res.json();
  const drive = data.value?.find(d => d.name === 'Documents' || d.name === 'Documentos compartidos' || d.driveType === 'documentLibrary');
  if (!drive) throw new Error('Drive no encontrado');
  return drive.id;
}

// ── Obtener Workbook Session ──────────────────────────────────────────────────
async function getWorkbookSession(token, siteId, driveId) {
  // Buscar el archivo
  const fileUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${FOLDER_PATH}/${FILE_NAME}`;
  const fileRes = await fetch(fileUrl, { headers: { Authorization: `Bearer ${token}` } });
  const fileData = await fileRes.json();
  if (!fileData.id) throw new Error('Archivo Excel no encontrado: ' + JSON.stringify(fileData));
  return fileData.id;
}

// ── Agregar fila al Excel ─────────────────────────────────────────────────────
async function appendToExcel(token, siteId, driveId, fileId, rowData) {
  // Primero verificar/crear la tabla
  const tableUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${fileId}/workbook/tables`;
  const tableRes = await fetch(tableUrl, { headers: { Authorization: `Bearer ${token}` } });
  const tableData = await tableRes.json();

  let tableEndpoint;

  if (!tableData.value || tableData.value.length === 0) {
    // Crear tabla con encabezados si no existe
    const createTable = await fetch(tableUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        address: 'Sheet1!A1:F1',
        hasHeaders: true,
      }),
    });
    const newTable = await createTable.json();

    // Agregar encabezados
    const headerUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${fileId}/workbook/worksheets('Sheet1')/range(address='A1:F1')`;
    await fetch(headerUrl, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        values: [['Fecha', 'Nombre', 'Ciudad', 'Institución', 'WhatsApp', 'Correo']],
      }),
    });

    tableEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${fileId}/workbook/tables/${newTable.id}/rows`;
  } else {
    tableEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${fileId}/workbook/tables/${tableData.value[0].id}/rows`;
  }

  // Agregar la fila
  const addRow = await fetch(tableEndpoint, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      values: [[
        rowData.fecha,
        rowData.nombre,
        rowData.ciudad,
        rowData.institucion,
        rowData.whatsapp,
        rowData.correo,
      ]],
    }),
  });

  if (!addRow.ok) {
    const err = await addRow.json();
    throw new Error('Error al agregar fila: ' + JSON.stringify(err));
  }
}

// ── Enviar correo de confirmación ─────────────────────────────────────────────
async function sendConfirmationEmail(token, inscrito) {
  // Obtener el userId del remitente
  const userRes = await fetch(
    `https://graph.microsoft.com/v1.0/users/${FROM_EMAIL}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const userData = await userRes.json();
  if (!userData.id) throw new Error('Usuario remitente no encontrado: ' + JSON.stringify(userData));

  const emailBody = {
    message: {
      subject: '🔺 MISIÓN SABER DELTA — Inscripción Confirmada',
      body: {
        contentType: 'HTML',
        content: `
          <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #081623; color: #c3ccd4; padding: 40px; border-radius: 12px;">
            <div style="text-align: center; margin-bottom: 30px;">
              <h1 style="color: #ffffff; font-size: 28px; letter-spacing: 4px;">🔺 MISIÓN SABER DELTA</h1>
              <p style="color: #8794a1; letter-spacing: 2px; font-size: 13px;">MILTON OCHOA</p>
            </div>

            <div style="background: rgba(255,255,255,0.05); border: 1px solid rgba(150,180,205,0.22); border-radius: 10px; padding: 25px; margin-bottom: 25px;">
              <h2 style="color: #e0b24c; font-size: 22px; margin-bottom: 5px;">✅ INSCRIPCIÓN CONFIRMADA</h2>
              <p style="font-size: 16px; line-height: 1.6;">Agente <strong style="color: #ffffff;">${inscrito.nombre}</strong>, tu inscripción al grupo Delta ha sido registrada exitosamente.</p>
            </div>

            <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(150,180,205,0.15); border-radius: 10px; padding: 20px; margin-bottom: 25px;">
              <h3 style="color: #4db3ff; margin-bottom: 15px; font-size: 14px; letter-spacing: 2px;">TUS DATOS DE INSCRIPCIÓN</h3>
              <table style="width: 100%; font-size: 15px;">
                <tr><td style="color: #8794a1; padding: 6px 0;">Nombre</td><td style="color: #ffffff;">${inscrito.nombre}</td></tr>
                <tr><td style="color: #8794a1; padding: 6px 0;">Ciudad</td><td style="color: #ffffff;">${inscrito.ciudad}</td></tr>
                <tr><td style="color: #8794a1; padding: 6px 0;">Institución</td><td style="color: #ffffff;">${inscrito.institucion}</td></tr>
                <tr><td style="color: #8794a1; padding: 6px 0;">WhatsApp</td><td style="color: #ffffff;">${inscrito.whatsapp}</td></tr>
                <tr><td style="color: #8794a1; padding: 6px 0;">Correo</td><td style="color: #ffffff;">${inscrito.correo}</td></tr>
              </table>
            </div>

            <p style="text-align: center; color: #8794a1; font-size: 14px; line-height: 1.6;">
              Pronto recibirás más información sobre la misión.<br/>
              Bienvenido a la élite. El profe <strong style="color: #ffffff;">Milton Ochoa</strong> te espera. 🔺
            </p>

            <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(150,180,205,0.15);">
              <p style="color: #4db3ff; font-size: 12px; letter-spacing: 3px;">ASESORÍAS ACADÉMICAS MILTON OCHOA</p>
              <p style="color: #8794a1; font-size: 11px;">aamocolombia.com</p>
            </div>
          </div>
        `,
      },
      toRecipients: [
        { emailAddress: { address: inscrito.correo, name: inscrito.nombre } },
      ],
    },
    saveToSentItems: true,
  };

  const sendRes = await fetch(
    `https://graph.microsoft.com/v1.0/users/${userData.id}/sendMail`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(emailBody),
    }
  );

  if (!sendRes.ok && sendRes.status !== 202) {
    const err = await sendRes.text();
    throw new Error('Error enviando correo: ' + err);
  }
}

// ── Handler principal ─────────────────────────────────────────────────────────
export default async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  try {
    const { nombre, ciudad, institucion, whatsapp, correo, fecha } = req.body;

    if (!nombre || !ciudad || !institucion || !whatsapp || !correo) {
      return res.status(400).json({ error: 'Faltan campos obligatorios' });
    }

    const inscrito = {
      nombre, ciudad, institucion, whatsapp, correo,
      fecha: fecha || new Date().toISOString(),
    };

    // 1. Obtener token
    const token = await getToken();

    // 2. Obtener IDs de SharePoint
    const siteId  = await getSiteId(token);
    const driveId = await getDriveId(token, siteId);
    const fileId  = await getWorkbookSession(token, siteId, driveId);

    // 3. Guardar en Excel
    await appendToExcel(token, siteId, driveId, fileId, inscrito);

    // 4. Enviar correo de confirmación
    await sendConfirmationEmail(token, inscrito);

    return res.status(200).json({ ok: true, message: 'Inscripción guardada y correo enviado' });

  } catch (error) {
    console.error('Error:', error);
    return res.status(500).json({ ok: false, error: error.message });
  }
}

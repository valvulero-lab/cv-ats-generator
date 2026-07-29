const { MercadoPagoConfig, Preference } = require('mercadopago');

const PRECIO_WORD = 3999; // ARS

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  if (!process.env.MP_ACCESS_TOKEN) {
    return res.status(500).json({ error: 'MercadoPago no está configurado (falta MP_ACCESS_TOKEN)' });
  }

  try {
    const client = new MercadoPagoConfig({ accessToken: process.env.MP_ACCESS_TOKEN });
    const preference = new Preference(client);

    // Origen del sitio (para que funcione igual en el dominio propio y en preview deployments de Vercel)
    const origin = req.headers.origin || `https://${req.headers.host}`;

    const result = await preference.create({
      body: {
        items: [
          {
            id: 'cv-word',
            title: 'CV Listo — Descarga en Word (.docx)',
            description: 'Descarga de tu CV optimizado en formato Word editable',
            quantity: 1,
            unit_price: PRECIO_WORD,
            currency_id: 'ARS',
          },
        ],
        back_urls: {
          success: `${origin}/`,
          failure: `${origin}/`,
          pending: `${origin}/`,
        },
        auto_return: 'approved',
        statement_descriptor: 'CV LISTO',
      },
    });

    res.status(200).json({ init_point: result.init_point, preference_id: result.id });
  } catch (error) {
    console.error('MercadoPago create-payment error:', error);
    res.status(500).json({ error: 'No se pudo iniciar el pago: ' + error.message });
  }
};

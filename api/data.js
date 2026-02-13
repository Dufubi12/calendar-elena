// Vercel Serverless Function for data storage
// Uses Vercel KV for persistent storage (optional - requires setup)
// For now, just returns a success response - client will use localStorage

export default function handler(req, res) {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method === 'GET') {
    // Return empty data - client will use localStorage
    return res.status(200).json({
      students: [],
      lessons: [],
      attendance: {}
    });
  }

  if (req.method === 'POST') {
    // Accept data but don't store it server-side
    // Client will use localStorage for persistence
    return res.status(200).json({ ok: true });
  }

  return res.status(405).json({ error: 'Method not allowed' });
}

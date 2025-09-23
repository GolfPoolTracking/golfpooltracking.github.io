const fetch = require('node-fetch');

const TENANT_ID    = process.env.TENANT_ID;
const CLIENT_ID    = process.env.CLIENT_ID;
const CLIENT_SECRET= process.env.CLIENT_SECRET;
const FILE_ID      = process.env.FILE_ID;      // Excel file ID
const NAMED_RANGE  = 'WEBSITE_RESULTS';
const USER_PRINCIPAL = process.env.USER_PRINCIPAL; // e.g. 'user@yourdomain.com'

module.exports = async (req, res) => {
  try {
    // 1️⃣ Get access token (application / client credentials)
    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: CLIENT_ID,
          scope: 'https://graph.microsoft.com/.default',
          client_secret: CLIENT_SECRET,
          grant_type: 'client_credentials'
        })
      }
    );

    const tokenData = await tokenResponse.json();
    if (!tokenData.access_token) throw new Error('Failed to get access token');
    const accessToken = tokenData.access_token;

    // 2️⃣ Use a user or drive endpoint (❌ /me → ✅ /users/{user})
    const rangeUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      USER_PRINCIPAL
    )}/drive/items/${FILE_ID}/workbook/names('${NAMED_RANGE}')/range`;

    console.log('Fetching URL:', rangeUrl);

    const rangeResponse = await fetch(rangeUrl, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!rangeResponse.ok) {
      const text = await rangeResponse.text();
      throw new Error(`Graph API request failed: ${text}`);
    }

    const rangeData = await rangeResponse.json();
    console.log('Range data:', JSON.stringify(rangeData, null, 2));

    // 3️⃣ Return values
    if (!rangeData.values || rangeData.values.length === 0) {
      res.status(200).json({ values: [], message: 'Named range is empty' });
      return;
    }

    res.status(200).json(rangeData);
  } catch (err) {
    console.error('Function error:', err);
    res.status(500).json({ error: 'Failed to fetch leaderboard', details: err.message });
  }
};

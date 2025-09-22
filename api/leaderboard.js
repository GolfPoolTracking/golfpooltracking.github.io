import fetch from 'node-fetch';

const TENANT_ID = process.env.a5635f25-6614-42fc-a5aa-51d5a7545c22;
const CLIENT_ID = process.env.bf792ce4-6a94-42a0-9c59-0672fdfe650;
const CLIENT_SECRET = process.env.924b3ddc-09d3-4918-9ae3-6aa1cadfc484;
const FILE_ID = process.env.29A99D65AA1785D4!s4fc07b975524497da651a46e4;          
const NAMED_RANGE = 'WEBSITE_RESULTS';

export default async function handler(req, res) {
    try {
        // 1️⃣ Get MS Graph token
        const tokenResponse = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                client_id: CLIENT_ID,
                scope: 'https://graph.microsoft.com/.default',
                client_secret: CLIENT_SECRET,
                grant_type: 'client_credentials'
            })
        });
        const tokenData = await tokenResponse.json();
        const accessToken = tokenData.access_token;

        // 2️⃣ Fetch the named range by file ID
        const rangeUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${FILE_ID}/workbook/names('${NAMED_RANGE}')/range`;

        const rangeResponse = await fetch(rangeUrl, {
            headers: { Authorization: `Bearer ${accessToken}` }
        });
        const rangeData = await rangeResponse.json();

        // Debug: log data
        console.log(JSON.stringify(rangeData, null, 2));

        res.status(200).json(rangeData);

    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Failed to fetch leaderboard', details: err.message });
    }
}


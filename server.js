const express = require('express');
const ace = require('atlassian-connect-express');
const axios = require('axios');
require('dotenv').config();

const app = express();
const addon = ace(app);
app.use(addon.middleware());

app.get('/excel-options', addon.authenticate(), async (req, res) => {
  try {
    const token = await getAccessToken();
    const items = await getExcelOptions(token);
    const options = items.map((item) => ({ value: item }));
    res.json({ items: options });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Unable to fetch options' });
  }
});

async function getAccessToken() {
  const tenant = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;

  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', clientId);
  params.append('client_secret', clientSecret);
  params.append('scope', 'https://graph.microsoft.com/.default');

  const resp = await axios.post(
    `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
    params
  );
  return resp.data.access_token;
}

async function getExcelOptions(accessToken) {
  const site = process.env.SHAREPOINT_SITE;
  const drive = process.env.SHAREPOINT_DRIVE;
  const fileId = process.env.EXCEL_FILE_ID;
  const worksheet = process.env.WORKSHEET_NAME;
  const range = process.env.RANGE;

  const url = `https://graph.microsoft.com/v1.0/sites/${site}/drives/${drive}/items/${fileId}/workbook/worksheets/${worksheet}/range(address='${range}')`;

  const resp = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const values = resp.data.values || [];
  const items = values.flat();
  return items;
}

const PORT = process.env.PORT || 3000;
addon.register();
app.listen(PORT, () => {
  console.log(`App running on port ${PORT}`);
});

// Graph API client for downloading the call-list xlsx from SharePoint.
// App-only auth via MSAL client credentials. Hardcoded site + filename for
// defense-in-depth — even if creds leak, this code only ever asks for one file.

const { ConfidentialClientApplication } = require('@azure/msal-node');

// Locked-down targets — must match the Sites.Selected grant scope.
// Site: https://insitepg.sharepoint.com/sites/Operations  ("The Vault")
const ALLOWED_SITE_ID =
  'insitepg.sharepoint.com,0d8bebcd-d7ac-4913-96db-d16e02a57880,7287695f-7f19-447e-98af-189c324236da';
const ALLOWED_FILENAME = 'SecureSpace Store Call List.xlsx';
const ALLOWED_HOST = 'insitepg.sharepoint.com';
const ALLOWED_SITE_PATH_FRAGMENT = '/sites/Operations/';

// Stable SharePoint share URL — uses sourcedoc GUID, survives folder/library moves.
const CALLLIST_SHARE_URL =
  'https://insitepg.sharepoint.com/:x:/r/sites/Operations/_layouts/15/Doc.aspx?sourcedoc=%7B3F62577E-F46B-46D1-8FEF-6DA28B048FE0%7D&file=SecureSpace%20Store%20Call%20List.xlsx&action=default&mobileredirect=true';

function assertAllowedShareUrl(url) {
  const u = new URL(url);
  if (u.host.toLowerCase() !== ALLOWED_HOST) {
    throw new Error(`SECURITY: refused share URL host "${u.host}". Only ${ALLOWED_HOST} is permitted.`);
  }
  if (!u.pathname.includes(ALLOWED_SITE_PATH_FRAGMENT)) {
    throw new Error(`SECURITY: refused share URL path "${u.pathname}". Must include ${ALLOWED_SITE_PATH_FRAGMENT}.`);
  }
}

function encodeShareUrl(url) {
  // Graph /shares API encoding: "u!" + URL-safe base64 (no padding)
  const b64 = Buffer.from(url).toString('base64')
    .replace(/=+$/, '')
    .replace(/\//g, '_')
    .replace(/\+/g, '-');
  return 'u!' + b64;
}

let _msalClient = null;
function getMsalClient() {
  if (_msalClient) return _msalClient;
  const clientId = process.env.CALLDIR_AZURE_CLIENT_ID;
  const clientSecret = process.env.CALLDIR_AZURE_CLIENT_SECRET;
  const tenantId = process.env.AZURE_TENANT_ID;
  if (!clientId || !clientSecret || !tenantId) {
    throw new Error(
      'Missing env: CALLDIR_AZURE_CLIENT_ID, CALLDIR_AZURE_CLIENT_SECRET, AZURE_TENANT_ID'
    );
  }
  _msalClient = new ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
  });
  return _msalClient;
}

async function getAccessToken() {
  const result = await getMsalClient().acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });
  if (!result || !result.accessToken) {
    throw new Error('Graph token acquisition returned empty result');
  }
  return result.accessToken;
}

async function graphGet(token, urlPath, opts = {}) {
  const url = urlPath.startsWith('http')
    ? urlPath
    : `https://graph.microsoft.com/v1.0${urlPath}`;
  const res = await fetch(url, {
    method: 'GET',
    headers: { Authorization: `Bearer ${token}` },
    ...opts,
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph GET ${urlPath} → ${res.status}: ${body.slice(0, 500)}`);
  }
  return res;
}

// Resolve the call list share URL to a Graph drive item.
async function resolveCallListItem(token) {
  assertAllowedShareUrl(CALLLIST_SHARE_URL);
  const encoded = encodeShareUrl(CALLLIST_SHARE_URL);
  const res = await graphGet(token, `/shares/${encoded}/driveItem`);
  const item = await res.json();
  if (item.name !== ALLOWED_FILENAME) {
    throw new Error(`SECURITY: share URL resolved to unexpected file "${item.name}"`);
  }
  return item;
}

async function downloadXlsxByDriveAndItem(token, driveId, itemId) {
  const path = `/drives/${driveId}/items/${itemId}/content`;
  const res = await graphGet(token, path, { redirect: 'follow' });
  const arrayBuf = await res.arrayBuffer();
  return Buffer.from(arrayBuf);
}

// Public: download the call list xlsx as a Buffer + metadata.
async function fetchCallListXlsx() {
  const token = await getAccessToken();
  const item = await resolveCallListItem(token);
  const driveId = item.parentReference?.driveId;
  if (!driveId) throw new Error('Resolved item has no parentReference.driveId');
  const buffer = await downloadXlsxByDriveAndItem(token, driveId, item.id);
  return {
    buffer,
    itemId: item.id,
    driveId,
    eTag: item.eTag,
    webUrl: item.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
    sizeBytes: buffer.length,
  };
}

module.exports = {
  fetchCallListXlsx,
  getAccessToken,        // exported for diagnostics
  ALLOWED_SITE_ID,
  ALLOWED_FILENAME,
};

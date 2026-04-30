// Push parsed call directory data to all 3 dashboard repos via GitHub Contents API.
// Per-file SHA-256 diff: only writes when content actually changed.

const crypto = require('crypto');

// Hard-coded targets — only these 3 repos x 3 files. Anything else is refused.
const PUBLISH_TARGETS = [
  {
    repo: 'tvt0002/transitions-dashboard',
    branch: 'master',
    paths: {
      stores:    'public/_static/call-directory/data/stores.json',
      corporate: 'public/_static/call-directory/data/corporate.json',
      people:    'public/_static/call-directory/data/people.json',
    },
  },
  {
    repo: 'tvt0002/axxerion-workorder-flow',
    branch: 'master',
    paths: {
      stores:    'public/call-directory/data/stores.json',
      corporate: 'public/call-directory/data/corporate.json',
      people:    'public/call-directory/data/people.json',
    },
  },
  {
    repo: 'tvt0002/securespace-dashboards',
    branch: 'main',
    paths: {
      stores:    'public/org-chart/data/stores.json',
      corporate: 'public/org-chart/data/corporate.json',
      people:    'public/org-chart/data/people.json',
    },
  },
];

function sha256(buf) {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

// Match Python json.dumps(indent=2, ensure_ascii=True): escape any UTF-16 code
// unit > 0x7F as \uXXXX so written bytes equal the existing committed JSON.
function jsonString(obj) {
  const s = JSON.stringify(obj, null, 2);
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code <= 0x7F) {
      out += s[i];
    } else {
      out += '\\u' + code.toString(16).padStart(4, '0');
    }
  }
  return out;
}

async function ghRequest(token, method, urlPath, body) {
  const url = urlPath.startsWith('http') ? urlPath : `https://api.github.com${urlPath}`;
  const opts = {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'calldir-sync',
    },
  };
  if (body !== undefined) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  }
  return await fetch(url, opts);
}

async function getFile(token, repo, branch, filePath) {
  const res = await ghRequest(
    token,
    'GET',
    `/repos/${repo}/contents/${encodeURIComponent(filePath).replace(/%2F/g, '/')}?ref=${branch}`
  );
  if (res.status === 404) return null;
  if (!res.ok) {
    throw new Error(`GET ${repo}/${filePath}@${branch} -> ${res.status}: ${(await res.text()).slice(0, 300)}`);
  }
  const json = await res.json();
  return {
    sha: json.sha,
    contentBuf: Buffer.from(json.content, 'base64'),
  };
}

async function putFile(token, repo, branch, filePath, newContentBuf, prevSha, message) {
  const body = {
    message,
    content: newContentBuf.toString('base64'),
    branch,
  };
  if (prevSha) body.sha = prevSha;
  const res = await ghRequest(
    token,
    'PUT',
    `/repos/${repo}/contents/${encodeURIComponent(filePath).replace(/%2F/g, '/')}`,
    body
  );
  if (!res.ok) {
    throw new Error(`PUT ${repo}/${filePath}@${branch} -> ${res.status}: ${(await res.text()).slice(0, 300)}`);
  }
  return await res.json();
}

async function publishCallList(parsed, opts = {}) {
  const token = opts.token || process.env.GITHUB_PAT_CALLDIR;
  if (!token) throw new Error('Missing env: GITHUB_PAT_CALLDIR');

  const dryRun = !!opts.dryRun;
  const today = new Date().toISOString().slice(0, 10);
  const message = `chore: sync call directory from SharePoint (${today})`;

  const payloads = {
    stores:    Buffer.from(jsonString(parsed.stores), 'utf8'),
    corporate: Buffer.from(jsonString(parsed.corporate), 'utf8'),
    people:    Buffer.from(jsonString(parsed.people), 'utf8'),
  };

  const results = [];
  for (const target of PUBLISH_TARGETS) {
    for (const [key, filePath] of Object.entries(target.paths)) {
      const newBuf = payloads[key];
      const newHash = sha256(newBuf);
      const result = { repo: target.repo, branch: target.branch, key, filePath, newHash };
      try {
        const existing = await getFile(token, target.repo, target.branch, filePath);
        const oldHash = existing ? sha256(existing.contentBuf) : null;
        result.oldHash = oldHash;
        if (oldHash === newHash) {
          result.action = 'unchanged';
        } else if (dryRun) {
          result.action = 'would-update';
        } else {
          const commit = await putFile(
            token, target.repo, target.branch, filePath, newBuf,
            existing?.sha, message
          );
          result.action = 'updated';
          result.commitSha = commit.commit?.sha;
          result.commitUrl = commit.commit?.html_url;
        }
      } catch (err) {
        result.action = 'error';
        result.error = err.message;
      }
      results.push(result);
    }
  }
  return results;
}

module.exports = { publishCallList, PUBLISH_TARGETS, jsonString };

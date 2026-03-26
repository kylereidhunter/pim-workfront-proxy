const http = require('http');
const https = require('https');
const zlib = require('zlib');
const url = require('url');

const API_KEY = '5apg8xk2bywd7fvugvpzyz91x5zhmlba';
const WORKFRONT_HOST = 'athome.my.workfront.com';
const PORT = process.env.PORT || 3000;

// Helper to call Workfront API
function callWorkfront(endpoint, params) {
  return new Promise((resolve, reject) => {
    const queryString = Object.entries({ apiKey: API_KEY, ...params })
      .map(([k, v]) => `${k}=${encodeURIComponent(v)}`)
      .join('&');

    const reqUrl = `https://${WORKFRONT_HOST}/attask/api/v17.0/${endpoint}?${queryString}`;

    https.get(reqUrl, { headers: { 'Accept-Encoding': 'gzip, deflate' } }, (res) => {
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => {
        const buffer = Buffer.concat(chunks);
        const encoding = res.headers['content-encoding'];

        const parseResult = (str) => {
          try { resolve(JSON.parse(str)); }
          catch (e) { resolve({ error: 'Failed to parse', raw: str.substring(0, 500) }); }
        };

        if (encoding === 'gzip') {
          zlib.gunzip(buffer, (err, decoded) => {
            if (err) { resolve({ error: 'Decompression failed' }); return; }
            parseResult(decoded.toString());
          });
        } else if (encoding === 'deflate') {
          zlib.inflate(buffer, (err, decoded) => {
            if (err) { resolve({ error: 'Decompression failed' }); return; }
            parseResult(decoded.toString());
          });
        } else {
          parseResult(buffer.toString());
        }
      });
    }).on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);
  const path = parsed.pathname;
  const query = parsed.query;

  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Content-Type', 'application/json');

  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }

  try {
    // GET /search?name=FY27 - Search projects by name
    if (path === '/search' || path === '/search/') {
      const searchName = query.name || 'FY27';
      const status = query.status || 'CUR';

      const result = await callWorkfront('proj/search', {
        name: searchName,
        name_Mod: 'contains',
        status: status,
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Proof URL,DE:Project Type,owner:name,tasks:name,tasks:plannedCompletionDate',
        '$$LIMIT': '200'
      });

      // Extract review dates from tasks and add to project level
      if (result.data) {
        result.data.forEach(proj => {
          const tasks = proj.tasks || [];
          tasks.forEach(t => {
            const name = (t.name || '').toLowerCase();
            if (name.includes('r1 - creative review') || name.includes('r1 - proof due for creative')) {
              if (name.includes('proof due')) proj.proofDueCreativeReview = t.plannedCompletionDate;
              else proj.creativeReviewDate = t.plannedCompletionDate;
            }
            if (name.includes('r2 - marketing review') || name.includes('r2 - proof due for marketing')) {
              if (name.includes('proof due')) proj.proofDueMarketingReview = t.plannedCompletionDate;
              else proj.marketingReviewDate = t.plannedCompletionDate;
            }
            if (name.includes('r3 - exec') || name.includes('r3 - proof due for exec')) {
              if (name.includes('proof due')) proj.proofDueExecReview = t.plannedCompletionDate;
              else proj.execReviewDate = t.plannedCompletionDate;
            }
            if (name.includes('r5 - deliver final files') || name.includes('deliver final')) {
              proj.deliverDate = t.plannedCompletionDate;
            }
          });
          // Add PM name from owner
          if (proj.owner) {
            proj.pm = proj.owner.name;
            delete proj.owner;
          }
          // Clean up - remove full task list to keep response smaller
          delete proj.tasks;
        });
      }

      res.writeHead(200);
      res.end(JSON.stringify(result, null, 2));
    }

    // GET /project/:id - Get a specific project by ID
    else if (path.startsWith('/project/')) {
      const projectId = path.split('/project/')[1];
      const result = await callWorkfront(`proj/${projectId}`, {
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Proof URL,DE:Project Type,description,owner:name,tasks:name,tasks:status,tasks:assignedTo:name,tasks:plannedCompletionDate'
      });

      res.writeHead(200);
      res.end(JSON.stringify(result, null, 2));
    }

    // GET /tasks?projectId=xxx - Get tasks for a project
    else if (path === '/tasks' || path === '/tasks/') {
      const projectId = query.projectId;
      if (!projectId) {
        res.writeHead(400);
        res.end(JSON.stringify({ error: 'projectId is required' }));
        return;
      }

      const result = await callWorkfront('task/search', {
        projectID: projectId,
        fields: 'name,status,assignedTo:name,plannedCompletionDate,percentComplete',
        '$$LIMIT': '100'
      });

      res.writeHead(200);
      res.end(JSON.stringify(result, null, 2));
    }

    // GET /proofs?name=WK15 - Get proof status for projects matching a name
    else if (path === '/proofs' || path === '/proofs/') {
      const searchName = query.name || 'FY27';
      const result = await callWorkfront('docu/search', {
        'project:name': searchName,
        'project:name_Mod': 'contains',
        'currentVersion:proofID_Mod': 'notnull',
        fields: 'name,project:name,currentVersion:proofID,currentVersion:proofStatus,currentVersion:proofDecision,currentVersion:proofStatusDate,currentVersion:fileName',
        '$$LIMIT': '100'
      });

      // Clean up the response for readability
      if (result.data) {
        result.data = result.data.map(doc => ({
          documentName: doc.name,
          projectName: doc.project ? doc.project.name : 'N/A',
          fileName: doc.currentVersion ? doc.currentVersion.fileName : 'N/A',
          proofID: doc.currentVersion ? doc.currentVersion.proofID : null,
          proofStatus: doc.currentVersion ? doc.currentVersion.proofStatus : 'no proof',
          proofDecision: doc.currentVersion ? doc.currentVersion.proofDecision : 'N/A',
          proofStatusDate: doc.currentVersion ? doc.currentVersion.proofStatusDate : null,
          hasProof: !!(doc.currentVersion && doc.currentVersion.proofID)
        }));
      }

      res.writeHead(200);
      res.end(JSON.stringify(result, null, 2));
    }

    // GET /my-projects?assignedTo=username - Get projects assigned to a person
    else if (path === '/my-projects' || path === '/my-projects/') {
      const assignee = query.assignee || '';
      const result = await callWorkfront('proj/search', {
        status: 'CUR',
        'DE:Designer': assignee,
        'DE:Designer_Mod': 'contains',
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Proof URL',
        '$$LIMIT': '100'
      });

      res.writeHead(200);
      res.end(JSON.stringify(result, null, 2));
    }

    // GET /upcoming-reviews - Get projects with upcoming review dates
    else if (path === '/upcoming-reviews' || path === '/upcoming-reviews/') {
      const result = await callWorkfront('proj/search', {
        name: query.name || 'FY27',
        name_Mod: 'contains',
        status: 'CUR',
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Proof URL',
        '$$LIMIT': '200'
      });

      res.writeHead(200);
      res.end(JSON.stringify(result, null, 2));
    }

    // GET /health - Health check
    else if (path === '/health' || path === '/') {
      res.writeHead(200);
      res.end(JSON.stringify({ status: 'ok', message: 'Pim Workfront Proxy is running' }));
    }

    else {
      res.writeHead(404);
      res.end(JSON.stringify({ error: 'Not found', availableEndpoints: ['/search?name=FY27', '/project/:id', '/tasks?projectId=xxx', '/my-projects?assignee=name', '/upcoming-reviews', '/health'] }));
    }

  } catch (err) {
    res.writeHead(500);
    res.end(JSON.stringify({ error: err.message }));
  }
});

server.listen(PORT, () => {
  console.log(`Pim Workfront Proxy running on port ${PORT}`);
});

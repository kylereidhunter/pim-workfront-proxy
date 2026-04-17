const https = require('https');
const zlib = require('zlib');

const API_KEY = '5apg8xk2bywd7fvugvpzyz91x5zhmlba';
const WORKFRONT_HOST = 'athome.my.workfront.com';

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

function projectUrlFor(id) {
  return id ? `https://${WORKFRONT_HOST}/project/${id}/overview` : null;
}

// Workfront returns dates like "2026-04-21T15:00:00:000-0500" (colon before ms). Fix to a real ISO string.
function parseWFDate(s) {
  if (!s) return null;
  const fixed = s.replace(/(\d{2}):(\d{3})/, '$1.$2');
  const d = new Date(fixed);
  return isNaN(d.getTime()) ? null : d;
}

// Resolve a window keyword to a [start, end] date range (inclusive).
function resolveWindow(window, startDate, endDate) {
  if (startDate && endDate) {
    return [new Date(startDate + 'T00:00:00'), new Date(endDate + 'T23:59:59')];
  }
  const now = new Date();
  const sunday = new Date(now);
  sunday.setDate(now.getDate() - now.getDay());
  sunday.setHours(0, 0, 0, 0);
  const saturday = new Date(sunday);
  saturday.setDate(sunday.getDate() + 6);
  saturday.setHours(23, 59, 59, 999);
  if (window === 'thisweek') return [sunday, saturday];
  if (window === 'nextweek') {
    const ns = new Date(sunday); ns.setDate(sunday.getDate() + 7);
    const ne = new Date(saturday); ne.setDate(saturday.getDate() + 7);
    return [ns, ne];
  }
  if (window === 'last7' || window === 'next7') {
    const s = new Date(now); const e = new Date(now);
    if (window === 'last7') s.setDate(now.getDate() - 7); else e.setDate(now.getDate() + 7);
    s.setHours(0,0,0,0); e.setHours(23,59,59,999);
    return [s, e];
  }
  if (window === 'thismonth') {
    const s = new Date(now.getFullYear(), now.getMonth(), 1);
    const e = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
    return [s, e];
  }
  return [null, null];
}

function extractReviewDates(result) {
  if (!result.data) return result;
  result.data = result.data.map(proj => {
    const out = {
      ID: proj.ID,
      name: proj.name,
      status: proj.status,
      designer: proj['DE:Lead Designer'] || null,
      copywriter: proj['DE:Lead Copywriter'] || null,
      pm: proj.owner ? proj.owner.name : null,
      channel: proj['DE:Channel'] || null,
      projectType: proj['DE:Project Type'] || null,
      liveDate: proj['DE:Live Date'] || null,
      fiscalWeek: proj['DE:Fiscal Weeks'] || null,
      proofUrl: proj['DE:Proof URL'] || null,
      projectUrl: projectUrlFor(proj.ID),
    };
    (proj.tasks || []).forEach(t => {
      const n = (t.name || '').toLowerCase();
      const isProofDue = n.includes('proof due');
      if (n.includes('r1 - creative review') || n.includes('r1 - proof due for creative')) {
        if (isProofDue) out.proofDueCreativeReview = t.plannedCompletionDate;
        else out.creativeReviewDate = t.plannedCompletionDate;
      }
      if (n.includes('r2 - marketing review') || n.includes('r2 - proof due for marketing')) {
        if (isProofDue) out.proofDueMarketingReview = t.plannedCompletionDate;
        else out.marketingReviewDate = t.plannedCompletionDate;
      }
      if (n.includes('r3 - exec') || n.includes('r3 - proof due for exec')) {
        if (isProofDue) out.proofDueExecReview = t.plannedCompletionDate;
        else out.execReviewDate = t.plannedCompletionDate;
      }
      if (n.includes('r5 - deliver final files') || n.includes('deliver final')) {
        out.deliverDate = t.plannedCompletionDate;
      }
    });
    return out;
  });
  return result;
}

module.exports = async (req, res) => {
  const { pathname: path, searchParams } = new URL(req.url, `http://${req.headers.host}`);
  const query = Object.fromEntries(searchParams);

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Content-Type', 'application/json');

  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    // GET /search?name=FY27
    if (path === '/search' || path === '/search/') {
      const result = await callWorkfront('proj/search', {
        name: query.name || 'FY27',
        name_Mod: 'contains',
        status: query.status || 'CUR',
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Proof URL,DE:Project Type,owner:name,tasks:name,tasks:plannedCompletionDate',
        '$$LIMIT': '200'
      });
      return res.status(200).json(extractReviewDates(result));
    }

    // GET /project/:id
    else if (path.startsWith('/project/')) {
      const projectId = path.split('/project/')[1];
      const result = await callWorkfront(`proj/${projectId}`, {
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Proof URL,DE:Project Type,description,owner:name,tasks:name,tasks:status,tasks:assignedTo:name,tasks:plannedCompletionDate'
      });
      if (result.data) result.data.projectUrl = projectUrlFor(result.data.ID || projectId);
      return res.status(200).json(result);
    }

    // GET /tasks?projectId=xxx
    else if (path === '/tasks' || path === '/tasks/') {
      if (!query.projectId) return res.status(400).json({ error: 'projectId is required' });
      const result = await callWorkfront('task/search', {
        projectID: query.projectId,
        fields: 'name,status,assignedTo:name,plannedCompletionDate,percentComplete',
        '$$LIMIT': '100'
      });
      return res.status(200).json(result);
    }

    // GET /docs?name=FY27 — all documents on matching projects, with current
    // version + proof info. Used by the change detector to diff new uploads,
    // version bumps, and proof status changes.
    //
    // Matching is fuzzy on project name: spaces/underscores/hyphens are
    // treated as equivalent, and multi-word queries match if ALL words appear
    // anywhere in the normalized project name. This lets "coffee table"
    // match "FY27_WK15_5-14_Refresh_Your_Coffee_Table-Email".
    else if (path === '/docs' || path === '/docs/') {
      const userQuery = String(query.name || 'FY27').trim();
      const normalize = (s) => String(s || '').toLowerCase().replace(/[_\-\s]+/g, ' ').trim();
      const queryWords = normalize(userQuery).split(' ').filter(Boolean);
      // Use the longest raw (original-case) word for the Workfront server-side
      // filter — Workfront's name_Mod=contains is case-sensitive, so we must
      // Title-Case it. Client-side filter still uses the lowercase words.
      const rawWords = userQuery.replace(/[_\-]+/g, ' ').split(/\s+/).filter(Boolean);
      const longestRaw = rawWords.sort((a, b) => b.length - a.length)[0] || userQuery;
      const titleCase = (w) => w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : w;
      const wfSearchWord = titleCase(longestRaw);
      const result = await callWorkfront('docu/search', {
        'project:name': wfSearchWord,
        'project:name_Mod': 'contains',
        fields: 'ID,name,lastUpdateDate,project:ID,project:name,currentVersion:ID,currentVersion:version,currentVersion:entryDate,currentVersion:proofID,currentVersion:proofStatus,currentVersion:proofDecision,currentVersion:proofStatusDate,currentVersion:fileName',
        '$$LIMIT': '500',
      });
      if (result.data) {
        result.data = result.data
          .map(d => ({
            docID: d.ID,
            name: d.name,
            projectID: d.project ? d.project.ID : null,
            projectName: d.project ? d.project.name : null,
            lastUpdateDate: d.lastUpdateDate,
            version: d.currentVersion ? d.currentVersion.version : null,
            versionID: d.currentVersion ? d.currentVersion.ID : null,
            versionEntryDate: d.currentVersion ? d.currentVersion.entryDate : null,
            fileName: d.currentVersion ? d.currentVersion.fileName : null,
            proofID: d.currentVersion ? d.currentVersion.proofID : null,
            proofStatus: d.currentVersion ? d.currentVersion.proofStatus : null,
            proofDecision: d.currentVersion ? d.currentVersion.proofDecision : null,
            proofStatusDate: d.currentVersion ? d.currentVersion.proofStatusDate : null,
            hasProof: !!(d.currentVersion && d.currentVersion.proofID),
          }))
          // Client-side: all query words must appear in the normalized project name.
          .filter(d => {
            if (!queryWords.length) return true;
            const projNorm = normalize(d.projectName);
            return queryWords.every(w => projNorm.includes(w));
          });
      }
      return res.status(200).json(result);
    }

    // GET /updates?name=FY27&sinceHours=24 — journal notes on matching
    // projects' Updates tab. Default window: last 24 hours.
    else if (path === '/updates' || path === '/updates/') {
      const sinceHours = Math.max(1, Math.min(168, Number(query.sinceHours) || 24));
      const since = new Date(Date.now() - sinceHours * 3600 * 1000).toISOString().slice(0, 10);
      const result = await callWorkfront('note/search', {
        objCode: 'PROJ',
        entryDate: since,
        entryDate_Mod: 'gte',
        fields: 'ID,entryDate,objID,ownerID,owner:name,noteText',
        '$$LIMIT': '500',
      });
      // Filter client-side to only notes whose project matches the name query.
      // Workfront's note/search won't filter by project name directly.
      if (result.data && query.name) {
        const nameFilter = String(query.name).toLowerCase();
        // Need project names for each note's objID. Do a batched proj lookup.
        const objIds = [...new Set(result.data.map(n => n.objID).filter(Boolean))];
        const projectLookup = {};
        if (objIds.length) {
          const projResult = await callWorkfront('proj/search', {
            ID: objIds.join(','),
            ID_Mod: 'in',
            fields: 'ID,name',
            '$$LIMIT': String(objIds.length + 10),
          });
          (projResult.data || []).forEach(p => { projectLookup[p.ID] = p.name; });
        }
        result.data = result.data
          .map(n => ({
            noteID: n.ID,
            entryDate: n.entryDate,
            projectID: n.objID,
            projectName: projectLookup[n.objID] || null,
            ownerName: n.owner ? n.owner.name : null,
            text: n.noteText || '',
          }))
          .filter(n => n.projectName && n.projectName.toLowerCase().includes(nameFilter));
      }
      return res.status(200).json(result);
    }

    // GET /proofs?name=WK15
    else if (path === '/proofs' || path === '/proofs/') {
      const result = await callWorkfront('docu/search', {
        'project:name': query.name || 'FY27',
        'project:name_Mod': 'contains',
        'currentVersion:proofID_Mod': 'notnull',
        fields: 'name,project:name,currentVersion:proofID,currentVersion:proofStatus,currentVersion:proofDecision,currentVersion:proofStatusDate,currentVersion:fileName',
        '$$LIMIT': '100'
      });
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
      return res.status(200).json(result);
    }

    // GET /my-projects?assignee=name
    else if (path === '/my-projects' || path === '/my-projects/') {
      const result = await callWorkfront('proj/search', {
        status: 'CUR',
        'DE:Designer': query.assignee || '',
        'DE:Designer_Mod': 'contains',
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Proof URL',
        '$$LIMIT': '100'
      });
      if (result.data) {
        result.data.forEach(proj => { proj.projectUrl = projectUrlFor(proj.ID); });
      }
      return res.status(200).json(result);
    }

    // GET /upcoming-reviews
    else if (path === '/upcoming-reviews' || path === '/upcoming-reviews/') {
      const result = await callWorkfront('proj/search', {
        name: query.name || 'FY27',
        name_Mod: 'contains',
        status: 'CUR',
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Project Type,DE:Proof URL,owner:name,tasks:name,tasks:plannedCompletionDate',
        '$$LIMIT': '200'
      });
      return res.status(200).json(extractReviewDates(result));
    }

    // GET /reviews?reviewType=creative|marketing|exec|any&window=thisweek|nextweek|last7|next7|thismonth
    //        &startDate=YYYY-MM-DD&endDate=YYYY-MM-DD&name=FY27&channel=email|text-push|loyalty|all
    // Returns ONLY projects whose matching review date falls in the window — server-side filter.
    else if (path === '/reviews' || path === '/reviews/') {
      const reviewType = (query.reviewType || 'any').toLowerCase();
      const [start, end] = resolveWindow(query.window, query.startDate, query.endDate);
      if (!start || !end) {
        return res.status(400).json({
          error: 'Provide window=thisweek|nextweek|last7|next7|thismonth OR startDate=YYYY-MM-DD&endDate=YYYY-MM-DD',
        });
      }
      const channel = (query.channel || 'all').toLowerCase();
      const raw = await callWorkfront('proj/search', {
        name: query.name || 'FY27',
        name_Mod: 'contains',
        status: 'CUR',
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Project Type,DE:Proof URL,owner:name,tasks:name,tasks:plannedCompletionDate',
        '$$LIMIT': '200',
      });
      const extracted = extractReviewDates(raw);
      const typeToField = {
        creative: 'creativeReviewDate',
        marketing: 'marketingReviewDate',
        mkt: 'marketingReviewDate',
        exec: 'execReviewDate',
      };
      const matchesWindow = (proj) => {
        const fields = reviewType === 'any'
          ? ['creativeReviewDate', 'marketingReviewDate', 'execReviewDate']
          : [typeToField[reviewType]];
        return fields.some(f => {
          const d = parseWFDate(proj[f]);
          return d && d >= start && d <= end;
        });
      };
      const matchesChannel = (proj) => {
        if (channel === 'all') return true;
        const ch = (proj.channel || '').toLowerCase();
        const type = (proj.projectType || '').toLowerCase();
        const name = (proj.name || '').toLowerCase();
        if (channel === 'email') return ch.includes('email');
        if (channel === 'text-push' || channel === 'push' || channel === 'text' || channel === 'sms') {
          return ch.includes('text') || ch.includes('sms') || ch.includes('push');
        }
        if (channel === 'loyalty') return type.includes('loyalty') || name.includes('loyalty');
        return true;
      };
      const filtered = (extracted.data || []).filter(p => matchesWindow(p) && matchesChannel(p));
      return res.status(200).json({
        window: { start: start.toISOString().slice(0, 10), end: end.toISOString().slice(0, 10) },
        reviewType,
        channel,
        count: filtered.length,
        projects: filtered,
      });
    }

    // POST /upload-proof - Upload a proof to a Workfront project
    // Expects JSON body: { projectName: "WK15 Patriotic", fileBase64: "...", fileName: "proof.pdf" }
    else if (path === '/upload-proof' || path === '/upload-proof/') {
      if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Use POST method' });
      }

      // Parse request body
      const body = await new Promise((resolve) => {
        let data = '';
        req.on('data', chunk => data += chunk);
        req.on('end', () => {
          try { resolve(JSON.parse(data)); }
          catch (e) { resolve({}); }
        });
      });

      const { projectName, fileBase64, fileName, createProof, workflow } = body;

      // Workflow templates with reviewer IDs
      const WORKFLOWS = {
        'Creative Review': {
          stages: [{
            name: 'Stage 1',
            position: 1,
            activateOn: 1, // on proof creation
            lockOn: 0, // manually
            recipients: [
              { id: '6418bcb8003fe85fc9bf9eb78095ee14', role: 5, alerts: 4 },  // Kyle Hunter - Reviewer, Daily summary
              { id: '62a21ad6003022c08375916ea6756b8a', role: 6, alerts: 0 },  // Meghan Miller - Reviewer & Approver
              { id: '6080d3b5001143b2fd695433c963f0c4', role: 5, alerts: 8 },  // Sharon Wernert - Reviewer, Final decision
              { id: '60c7c956001d7533e9e1ebb9f5416ebd', role: 5, alerts: 0 },  // Ryan Creery - Reviewer
              { id: '6113f54c0012ff2e549c08be46d55263', role: 5, alerts: 8 },  // Meagan Goldberg - Reviewer, Final decision
              { id: '61fc2c490015176f176b8253f8b7ee80', role: 5, alerts: 4 },  // Alise Gray - Reviewer, Daily summary
              { id: '625d8bcc00ec74e8fc654cd882416420', role: 5, alerts: 0 },  // Danielle Maday - Reviewer
              { id: '676ecd4e00dd0481d77743410018fc37', role: 5, alerts: 0 },  // Peyton Mackay - Reviewer
              { id: '66954411001d7da52322ed33e756f04d', role: 5, alerts: 0 },  // Amber Mischo - Reviewer
              { id: '67e435340004215a520ea5cb7bdc4c15', role: 5, alerts: 0 },  // Gracie Chavez - Reviewer
              { id: '67e1cacb00055186317c9290247e16d0', role: 5, alerts: 0 },  // Traci Pruitt - Reviewer
              { id: '6080d4d00011a5bf15872c9ad5b7c835', role: 6, alerts: 0 },  // Hannah Bryant - Reviewer & Approver
              { id: '690d24330005054eaefccdcb745e72ac', role: 5, alerts: 0 },  // Meaghan Stevens - Reviewer
              { id: '64ee14b0002146bd55a0dbabd636b1a6', role: 10, alerts: 16 }, // Jilly Blythe - Author, Decisions
            ]
          }]
        }
      };

      const selectedWorkflow = workflow || 'Creative Review';

      if (!projectName || !fileBase64 || !fileName) {
        return res.status(400).json({ error: 'Required: projectName, fileBase64, fileName' });
      }

      // Step 1: Find the project
      const projResult = await callWorkfront('proj/search', {
        name: projectName,
        name_Mod: 'contains',
        status: 'CUR',
        fields: 'name'
      });

      if (!projResult.data || projResult.data.length === 0) {
        return res.status(404).json({ error: `No project found matching "${projectName}"` });
      }

      const project = projResult.data[0];

      // Step 2: Upload file to Workfront
      const fileBuffer = Buffer.from(fileBase64, 'base64');
      const boundary = '----PimUpload' + Date.now();

      const formParts = [
        `--${boundary}\r\n`,
        `Content-Disposition: form-data; name="uploadedFile"; filename="${fileName}"\r\n`,
        `Content-Type: application/octet-stream\r\n\r\n`
      ];
      const formEnd = `\r\n--${boundary}--\r\n`;

      const formHeader = Buffer.from(formParts.join(''));
      const formFooter = Buffer.from(formEnd);
      const formBody = Buffer.concat([formHeader, fileBuffer, formFooter]);

      const uploadResult = await new Promise((resolve, reject) => {
        const uploadReq = https.request({
          hostname: WORKFRONT_HOST,
          path: `/attask/api/v17.0/upload?apiKey=${API_KEY}`,
          method: 'POST',
          headers: {
            'Content-Type': `multipart/form-data; boundary=${boundary}`,
            'Content-Length': formBody.length,
            'Accept-Encoding': 'gzip, deflate'
          }
        }, (uploadRes) => {
          const chunks = [];
          uploadRes.on('data', c => chunks.push(c));
          uploadRes.on('end', () => {
            const buf = Buffer.concat(chunks);
            const enc = uploadRes.headers['content-encoding'];
            const parse = (s) => { try { resolve(JSON.parse(s)); } catch(e) { resolve({ error: s.substring(0,200) }); } };
            if (enc === 'gzip') { zlib.gunzip(buf, (e, d) => parse(d ? d.toString() : '')); }
            else { parse(buf.toString()); }
          });
        });
        uploadReq.on('error', reject);
        uploadReq.write(formBody);
        uploadReq.end();
      });

      if (!uploadResult.data || !uploadResult.data.handle) {
        return res.status(500).json({ error: 'Upload failed', details: uploadResult });
      }

      const handle = uploadResult.data.handle;

      // Step 3: Create document record linked to project
      const docParams = {
        name: fileName,
        handle: handle,
        docObjCode: 'PROJ',
        objID: project.ID
      };
      if (createProof !== false) {
        docParams.createProof = 'true';
      }

      const docResult = await callWorkfront('docu', {
        ...docParams,
        method: 'POST'
      });

      // Step 3b: Check if document already exists (for versioning)
      const existingDocs = await callWorkfront('docu/search', {
        'project:ID': project.ID,
        name: fileName,
        name_Mod: 'contains',
        fields: 'name,ID',
        '$$LIMIT': '1'
      });

      const proofWorkflow = body.workflow || 'Creative Review';

      // Create document (or new version if exists)
      const createDoc = await new Promise((resolve, reject) => {
        let postData;
        if (existingDocs.data && existingDocs.data.length > 0) {
          // Upload as new version of existing document
          const existingDocId = existingDocs.data[0].ID;
          postData = `docID=${existingDocId}&fileName=${encodeURIComponent(fileName)}&handle=${handle}&createProof=true`;
          // POST to docv (document version) for new version
          const docvReq = https.request({
            hostname: WORKFRONT_HOST,
            path: `/attask/api/v17.0/docv?apiKey=${API_KEY}`,
            method: 'POST',
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded',
              'Accept-Encoding': 'gzip, deflate'
            }
          }, (docRes) => {
            const chunks = [];
            docRes.on('data', c => chunks.push(c));
            docRes.on('end', () => {
              const buf = Buffer.concat(chunks);
              const enc = docRes.headers['content-encoding'];
              const parse = (s) => { try { resolve({ ...JSON.parse(s), isNewVersion: true }); } catch(e) { resolve({ error: s.substring(0,200) }); } };
              if (enc === 'gzip') { zlib.gunzip(buf, (e, d) => parse(d ? d.toString() : '')); }
              else { parse(buf.toString()); }
            });
          });
          docvReq.on('error', reject);
          docvReq.write(postData);
          docvReq.end();
          return;
        }

        // Create new document with workflow
        const wf = WORKFLOWS[selectedWorkflow];
        const advancedOptions = wf ? JSON.stringify(wf) : '';
        postData = `name=${encodeURIComponent(fileName)}&handle=${handle}&docObjCode=PROJ&objID=${project.ID}&createProof=true` + (advancedOptions ? `&advancedProofingOptions=${encodeURIComponent(advancedOptions)}` : '');
        const docReq = https.request({
          hostname: WORKFRONT_HOST,
          path: `/attask/api/v17.0/docu?apiKey=${API_KEY}`,
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept-Encoding': 'gzip, deflate'
          }
        }, (docRes) => {
          const chunks = [];
          docRes.on('data', c => chunks.push(c));
          docRes.on('end', () => {
            const buf = Buffer.concat(chunks);
            const enc = docRes.headers['content-encoding'];
            const parse = (s) => { try { resolve(JSON.parse(s)); } catch(e) { resolve({ error: s.substring(0,200) }); } };
            if (enc === 'gzip') { zlib.gunzip(buf, (e, d) => parse(d ? d.toString() : '')); }
            else { parse(buf.toString()); }
          });
        });
        docReq.on('error', reject);
        docReq.write(postData);
        docReq.end();
      });

      const isNewVersion = createDoc.isNewVersion || false;
      return res.status(200).json({
        success: true,
        message: isNewVersion
          ? `New version of "${fileName}" uploaded to project "${project.name}"`
          : `Proof "${fileName}" uploaded to project "${project.name}"`,
        projectId: project.ID,
        projectName: project.name,
        isNewVersion: isNewVersion,
        workflow: body.workflow || 'Creative Review',
        document: createDoc
      });
    }

    // GET /weekly-digest - Generate the formatted weekly digest message
    else if (path === '/weekly-digest' || path === '/weekly-digest/') {
      const result = await callWorkfront('proj/search', {
        name: 'FY27',
        name_Mod: 'contains',
        status: 'CUR',
        fields: 'name,status,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Live Date,owner:name,tasks:name,tasks:plannedCompletionDate',
        '$$LIMIT': '200'
      });

      // Extract review dates
      const projects = [];
      if (result.data) {
        result.data.forEach(proj => {
          const tasks = proj.tasks || [];
          const info = {
            name: proj.name,
            designer: proj['DE:Lead Designer'] || 'TBD',
            copywriter: proj['DE:Lead Copywriter'] || 'TBD',
            fiscalWeek: proj['DE:Fiscal Weeks'] || '',
            channel: proj['DE:Channel'] || '',
            liveDate: proj['DE:Live Date'] || '',
            pm: proj.owner ? proj.owner.name : 'TBD'
          };
          tasks.forEach(t => {
            const tname = (t.name || '').toLowerCase();
            if (tname.includes('r1 - creative review') && !tname.includes('proof')) info.creativeReview = t.plannedCompletionDate;
            if (tname.includes('r2 - marketing review') && !tname.includes('proof')) info.marketingReview = t.plannedCompletionDate;
            if (tname.includes('r3 - exec') && !tname.includes('proof')) info.execReview = t.plannedCompletionDate;
            if (tname.includes('r1 - proof due')) info.proofDueCR = t.plannedCompletionDate;
            if (tname.includes('r2 - proof due')) info.proofDueMKT = t.plannedCompletionDate;
            if (tname.includes('r3 - proof due')) info.proofDueExec = t.plannedCompletionDate;
          });
          // Only include email/SMS/push projects
          if (info.channel && info.channel.includes('Email')) projects.push(info);
        });
      }

      // Find next week's dates (Sunday through Saturday to catch all review days)
      const now = new Date();
      const dayOfWeek = now.getDay(); // 0=Sun, 1=Mon, ...
      const daysUntilNextSunday = 7 - dayOfWeek;
      const nextSunday = new Date(now);
      nextSunday.setDate(now.getDate() + daysUntilNextSunday);
      nextSunday.setHours(0, 0, 0, 0);
      const nextSaturday = new Date(nextSunday);
      nextSaturday.setDate(nextSunday.getDate() + 6);
      nextSaturday.setHours(23, 59, 59, 999);
      const nextMonday = new Date(nextSunday);
      nextMonday.setDate(nextSunday.getDate() + 1);
      const nextFriday = new Date(nextSunday);
      nextFriday.setDate(nextSunday.getDate() + 5);
      nextFriday.setHours(23, 59, 59, 999);

      // Group by review type for next week
      // Workfront dates use colon before ms (T15:00:00:000) - fix to dot (T15:00:00.000)
      const parseWFDate = (s) => s ? new Date(s.replace(/(\d{2}):(\d{3})/, '$1.$2')) : null;

      const crProjects = projects.filter(p => {
        if (!p.creativeReview) return false;
        const d = parseWFDate(p.creativeReview);
        return d && d >= nextSunday && d <= nextSaturday;
      });
      const mktProjects = projects.filter(p => {
        if (!p.marketingReview) return false;
        const d = parseWFDate(p.marketingReview);
        return d && d >= nextSunday && d <= nextSaturday;
      });
      const execProjects = projects.filter(p => {
        if (!p.execReview) return false;
        const d = parseWFDate(p.execReview);
        return d && d >= nextSunday && d <= nextSaturday;
      });

      // Format dates (fix Workfront colon-before-ms format)
      const fmt = (d) => {
        if (!d) return '';
        const dt = parseWFDate(d);
        if (!dt || isNaN(dt.getTime())) return '';
        return `${dt.getMonth()+1}/${dt.getDate()}`;
      };

      // Build the message with Pim's personality
      const openers = [
        "Happy Friday, team! ☀️ Your girl Pim has been crunching the numbers and here's what's on deck for next week!",
        "TGIF, creative crew! 🎨 Pim here with your weekly cheat sheet — let's make next week a masterpiece!",
        "Friday vibes activated! 🎉 Time for your favorite part of the week — Pim's email rundown!",
        "Rise and grind, beautiful people! ☕ Pim's got your next week all mapped out. Let's goooo!",
        "Another week, another batch of amazing emails! 💌 Your friendly neighborhood Pim is here with the scoop!",
        "Hey hey hey! 👋 It's Pim o'clock — which means it's time to peek at what's cooking for next week!",
        "Pop quiz: what's better than Friday? Friday with Pim's weekly digest! 📋 Here's the lineup!",
        "Gooood morning, rockstars! 🌟 Pim just pulled the latest from Workfront — here's what next week looks like!"
      ];
      const closers = [
        "That's the rundown! You've got this, team. Have an amazing weekend! 🙌 — Pim",
        "And that's a wrap! Go crush it. See you Monday! 💪 — Pim",
        "Questions? You know where to find me. Happy weekend, legends! ✨ — Pim",
        "That's all for now! Remember, proofs wait for no one. Have a great weekend! 🏖️ — Pim",
        "Boom. Done. Go enjoy your Friday — you've earned it! 🎊 — Pim"
      ];
      const opener = openers[Math.floor(Math.random() * openers.length)];
      const closer = closers[Math.floor(Math.random() * closers.length)];

      let message = `${opener}\n\n`;

      if (crProjects.length > 0) {
        const crDate = fmt(crProjects[0].creativeReview);
        const proofDue = crProjects[0].proofDueCR ? fmt(crProjects[0].proofDueCR) : '';
        message += `**Creative Review ${crDate}**\n`;
        message += `* Proof links due by 3 PM Monday ${proofDue}\n`;
        crProjects.forEach(p => {
          const shortName = p.name.replace('FY27_', '').replace(/_/g, ' ');
          message += `* ${shortName} - **${p.designer}** / ${p.copywriter}\n`;
        });
        message += '\n';
      }

      if (mktProjects.length > 0) {
        const mktDate = fmt(mktProjects[0].marketingReview);
        const proofDue = mktProjects[0].proofDueMKT ? fmt(mktProjects[0].proofDueMKT) : '';
        message += `**MKT Review ${mktDate}**\n`;
        message += `* Proof & JPEGs due by 1 PM Tuesday ${proofDue}\n`;
        const byWeek = {};
        mktProjects.forEach(p => {
          const wk = 'WK' + (p.fiscalWeek || '?');
          if (!byWeek[wk]) byWeek[wk] = [];
          byWeek[wk].push(p);
        });
        Object.entries(byWeek).forEach(([wk, ps]) => {
          message += `* ${wk}:\n`;
          ps.forEach(p => {
            const shortName = p.name.replace('FY27_', '').replace(/_/g, ' ');
            message += `  * ${shortName} - **${p.designer}** / ${p.copywriter}\n`;
          });
        });
        message += '\n';
      }

      if (execProjects.length > 0) {
        const execDate = fmt(execProjects[0].execReview);
        const proofDue = execProjects[0].proofDueExec ? fmt(execProjects[0].proofDueExec) : '';
        message += `**EXEC Review ${execDate}**\n`;
        message += `* Proof & JPEGs due by 1 PM Wednesday ${proofDue}\n`;
        execProjects.forEach(p => {
          const shortName = p.name.replace('FY27_', '').replace(/_/g, ' ');
          message += `* ${shortName} - **${p.designer}** / ${p.copywriter}\n`;
        });
        message += '\n';
      } else {
        message += `*No EXEC Review next week*\n\n`;
      }

      message += closer;

      return res.status(200).json({
        message: message,
        summary: {
          creativeReview: crProjects.length,
          mktReview: mktProjects.length,
          execReview: execProjects.length,
          total: crProjects.length + mktProjects.length + execProjects.length
        }
      });
    }

    // GET /cron — fires due scheduled messages AND polls Workfront for project
    // changes, DMing subscribed users about any updates. Called every 5 min by
    // cron-job.org. Merged into index.js because Vercel wasn't picking up
    // api/cron.js as its own function.
    else if (path === '/cron' || path === '/cron/') {
      const auth = req.headers.authorization || '';
      if (process.env.CRON_SECRET && !auth.includes(process.env.CRON_SECRET)) {
        return res.status(401).json({ error: 'unauthorized' });
      }
      const summary = { ok: true };
      try {
        // Part 1: fire due scheduled messages (weekly digests, reminders, etc.)
        const {
          getDueSchedules,
          getConversationRef,
          markFired,
        } = require('./lib/schedule-store');
        const { sendProactive } = require('./lib/proactive');
        const { buildMessage } = require('./lib/message-builder');
        const due = await getDueSchedules();
        const scheduleResults = [];
        for (const sched of due) {
          try {
            const ref = await getConversationRef(sched.conversationId);
            if (!ref) { scheduleResults.push({ id: sched.id, status: 'no-conv-ref' }); continue; }
            const text = await buildMessage(sched.messageKind, sched.messageArgs);
            await sendProactive(ref, text);
            await markFired(sched.id);
            scheduleResults.push({ id: sched.id, status: 'sent' });
          } catch (err) {
            scheduleResults.push({ id: sched.id, status: 'error', error: err.message });
          }
        }
        summary.fired = scheduleResults.length;
        summary.scheduleResults = scheduleResults;
      } catch (err) {
        summary.scheduleError = err.message;
      }
      try {
        // Part 2: poll Workfront, diff against last snapshot, DM subscribed
        // users about changes to their projects.
        const { detectAndNotify } = require('./lib/change-detector');
        const changeSummary = await detectAndNotify();
        summary.changes = changeSummary;
      } catch (err) {
        summary.changeError = err.message;
      }
      return res.status(200).json(summary);
    }

    // GET / or /health
    else if (path === '/health' || path === '/') {
      return res.status(200).json({ status: 'ok', message: 'Pim Workfront Proxy is running' });
    }

    else {
      return res.status(404).json({ error: 'Not found', availableEndpoints: ['/search?name=FY27', '/project/:id', '/tasks?projectId=xxx', '/proofs?name=WK15', '/my-projects?assignee=name', '/upcoming-reviews'] });
    }

  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
};

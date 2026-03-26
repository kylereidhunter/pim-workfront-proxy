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

function extractReviewDates(result) {
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
      if (proj.owner) {
        proj.pm = proj.owner.name;
        delete proj.owner;
      }
      delete proj.tasks;
    });
  }
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
      return res.status(200).json(result);
    }

    // GET /upcoming-reviews
    else if (path === '/upcoming-reviews' || path === '/upcoming-reviews/') {
      const result = await callWorkfront('proj/search', {
        name: query.name || 'FY27',
        name_Mod: 'contains',
        status: 'CUR',
        fields: 'name,status,plannedStartDate,plannedCompletionDate,DE:Creative Due Date,DE:Live Date,DE:Lead Designer,DE:Lead Copywriter,DE:Fiscal Weeks,DE:Channel,DE:Proof URL,owner:name,tasks:name,tasks:plannedCompletionDate',
        '$$LIMIT': '200'
      });
      return res.status(200).json(extractReviewDates(result));
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

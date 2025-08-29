# Bulk-Email-Sender-2
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Bulk Email Sender</title>
  <!-- SheetJS (Excel) & PapaParse (CSV) -->
  <script src="https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    :root { --bg:#0b1020; --card:#121934; --ink:#e9ecff; --muted:#b8c1ff; --accent:#8fa2ff; --good:#3ecf8e; --warn:#ffd166; --bad:#ff6b6b; }
    *{box-sizing:border-box}
    body{margin:0;font-family:Inter,system-ui,Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--ink)}
    .wrap{max-width:1100px;margin:24px auto;padding:16px}
    .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
    .card{background:var(--card);border:1px solid #1b2455;border-radius:16px;padding:16px;box-shadow:0 6px 24px rgba(0,0,0,.25)}
    h1{font-size:22px;margin:0 0 12px}
    h2{font-size:16px;margin:0 0 10px;color:var(--muted)}
    label{font-size:12px;color:var(--muted);display:block;margin:8px 0 4px}
    input,select,textarea{width:100%;background:#0d1430;color:var(--ink);border:1px solid #25306b;border-radius:10px;padding:10px}
    textarea{min-height:220px;font-family:ui-monospace,Menlo,monospace}
    .row{display:grid;grid-template-columns:1fr 1fr;gap:10px}
    .btn{cursor:pointer;padding:10px 14px;border-radius:12px;border:1px solid #25306b;background:#1a2557;color:#fff}
    .btn.primary{background:linear-gradient(180deg,#3b57ff,#2a45e8)}
    .btn:disabled{opacity:.6;cursor:not-allowed}
    .pill{display:inline-flex;align-items:center;gap:8px;padding:6px 10px;border-radius:999px;background:#0f173b;border:1px solid #23307c;color:#cdd5ff;font-size:12px}
    .table{overflow:auto;border:1px solid #23307c;border-radius:12px}
    table{border-collapse:collapse;width:100%}
    th,td{padding:8px 10px;border-bottom:1px solid #202a64;font-size:12px}
    th{position:sticky;top:0;background:#101949;text-align:left}
    .preview{background:#0d1430;border:1px dashed #2b3681;border-radius:12px;padding:12px;min-height:240px}
    .badge{font-size:11px;padding:3px 8px;border-radius:999px;background:#142058;border:1px solid #2a3a8f;color:#cdd5ff}
    .muted{color:#9aa6ff}
  </style>
</head>
<body>
<div class="wrap">
  <h1>Bulk Email Sender</h1>
  <div class="grid">
    <div class="card">
      <h2>1) Upload CSV or Excel</h2>
      <div class="row">
        <div>
          <label>Upload file (.csv, .xlsx, .xls)</label>
          <input id="fileInput" type="file" accept=".csv,.xlsx,.xls" />
        </div>
        <div>
          <label>Attachment (sent to everyone)</label>
          <input id="attachInput" type="file" />
        </div>
      </div>
      <div class="row">
        <div>
          <label>Column for recipient email</label>
          <select id="emailCol"></select>
        </div>
        <div>
          <label>Test cap (send first N emails only)</label>
          <input id="maxToSend" type="number" min="0" placeholder="e.g. 5 for test" />
        </div>
      </div>
      <div class="row">
        <div>
          <label>Delay between emails (ms)</label>
          <input id="delayMs" type="number" min="0" value="800" />
        </div>
        <div>
          <label>Reply-To (optional)</label>
          <input id="replyTo" type="email" placeholder="reply@yourdomain.com" />
        </div>
      </div>
      <div class="row">
        <div>
          <label>Sender name</label>
          <input id="senderName" placeholder="Rahul" />
        </div>
        <div>
          <label>Subject (supports {{vars}})</label>
          <input id="subject" placeholder="Hello {{Name}} – Your August update" />
        </div>
      </div>
    </div>

    <div class="card">
      <h2>2) Email Body (with {{Column}} variables)</h2>
      <label>Template</label>
      <textarea id="htmlTemplate"></textarea>

      <!-- CTA Inputs -->
      <div class="row" style="margin-top:10px">
        <div>
          <label for="cta-text">Button Text</label>
          <input id="cta-text" type="text" placeholder="e.g. Visit our website" />
        </div>
        <div>
          <label for="cta-url">Button URL</label>
          <input id="cta-url" type="url" placeholder="e.g. https://example.com" />
        </div>
      </div>

      <div style="display:flex;gap:8px;margin-top:10px">
        <button class="btn" id="loadSample">Load sample template</button>
        <button class="btn" id="previewBtn">Preview with first row</button>
      </div>
    </div>
  </div>

  <div class="grid" style="margin-top:16px">
    <div class="card">
      <h2>3) Data preview</h2>
      <div class="pill"><span id="count">0</span> rows loaded</div>
      <div class="table" style="margin-top:10px"><table id="dataTable"></table></div>
    </div>

    <div class="card">
      <h2>4) Live email preview</h2>
      <div class="badge">Subject:</div>
      <div id="previewSubject" class="muted" style="margin:6px 0 10px">(none)</div>
      <div class="preview" id="previewHtml">Your email preview will appear here.</div>
      <div style="display:flex;gap:8px;margin-top:12px">
        <button class="btn" id="sendBtn">Send Emails</button>
        <button class="btn" id="send5Btn">Send 5 (test)</button>
      </div>
      <div id="status" style="margin-top:10px" class="muted"></div>
    </div>
  </div>
</div>

<script>
  let ROWS = [];
  let HEADERS = [];
  let ATTACHMENT = null;

  const $ = (id)=>document.getElementById(id);
  function setStatus(msg){ $('status').textContent = msg; }

  function renderTable(rows){
    const tbl = $('dataTable');
    tbl.innerHTML = '';
    if (!rows.length) return;
    const headers = Object.keys(rows[0]);
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    headers.forEach(h=>{ const th=document.createElement('th'); th.textContent=h; trh.appendChild(th); });
    thead.appendChild(trh);
    tbl.appendChild(thead);
    const tbody = document.createElement('tbody');
    rows.slice(0,50).forEach(r=>{
      const tr=document.createElement('tr');
      headers.forEach(h=>{ const td=document.createElement('td'); td.textContent = r[h]; tr.appendChild(td); });
      tbody.appendChild(tr);
    });
    tbl.appendChild(tbody);
  }

  function parseCSVText(text){
    const out = Papa.parse(text, {header:true, skipEmptyLines:true});
    return out.data;
  }
  function parseExcel(data){
    const wb = XLSX.read(data, {type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, {defval:''});
  }

  function sampleTemplate(){
    return `
<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f4f6fb;padding:20px">
  <tr>
    <td align="center">
      <table role="presentation" width="640" cellspacing="0" cellpadding="0" style="max-width:640px;background:#ffffff;border-radius:8px;overflow:hidden;border:1px solid #e6e9f4">
        <tr>
          <td style="padding:20px 24px;background:#0b5cff;color:#ffffff;font:700 20px/1.2 Inter,Arial">THE PLUMBING HUB</td>
        </tr>
        <tr>
          <td style="padding:24px 24px 10px;font:600 18px/1.3 Inter,Arial;color:#111827">Hi {{Name}},</td>
        </tr>
        <tr>
          <td style="padding:0 24px 14px;font:400 14px/1.6 Inter,Arial;color:#374151">
            Thanks for your interest in our services. Your quote id is <b>{{QuoteId}}</b> and current status is <b>{{Status}}</b>.
          </td>
        </tr>
        <tr>
          <td style="padding:0 24px 18px;font:400 14px/1.6 Inter,Arial;color:#374151">
            If you have any questions, just reply to this email — we're happy to help.
          </td>
        </tr>
        <tr>
          <td style="padding:0 24px 24px">
            <a href="{{ButtonUrl}}" style="display:inline-block;background:#0b5cff;color:#fff;text-decoration:none;font:600 14px Inter,Arial;padding:10px 16px;border-radius:6px">{{ButtonText}}</a>
          </td>
        </tr>
        <tr>
          <td style="padding:16px 24px;background:#f8fafc;font:12px Inter,Arial;color:#6b7280">
            You are receiving this because you interacted with THE PLUMBING HUB.
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>`;
  }

  $('fileInput').addEventListener('change', async (e)=>{
    const file = e.target.files[0];
    if (!file) return;
    const name = file.name.toLowerCase();
    setStatus('Parsing '+file.name+' ...');
    if (name.endsWith('.csv')) {
      const text = await file.text();
      ROWS = parseCSVText(text);
    } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
      const buf = await file.arrayBuffer();
      ROWS = parseExcel(buf);
    } else {
      alert('Unsupported file type.');
      return;
    }
    HEADERS = ROWS.length ? Object.keys(ROWS[0]) : [];
    $('count').textContent = ROWS.length;
    renderTable(ROWS);
    renderEmailColOptions();
    setStatus('Loaded '+ROWS.length+' rows.');
  });

  function renderEmailColOptions(){
    const sel = $('emailCol');
    sel.innerHTML = '';
    HEADERS.forEach(h=>{
      const opt = document.createElement('option');
      opt.value = h; opt.textContent = h; sel.appendChild(opt);
    });
    const guess = HEADERS.find(h=>/mail/i.test(h)) || HEADERS[0];
    if (guess) sel.value = guess;
  }

  $('attachInput').addEventListener('change', async (e)=>{
    const file = e.target.files[0];
    if (!file) { ATTACHMENT = null; return; }
    const b64 = await toBase64(file);
    ATTACHMENT = { base64: b64.split(',')[1], mimeType: file.type || 'application/octet-stream', filename: file.name };
  });
  function toBase64(file){
    return new Promise((resolve, reject)=>{
      const r = new FileReader();
      r.onload = ()=>resolve(r.result);
      r.onerror = reject;
      r.readAsDataURL(file);
    });
  }

  $('loadSample').addEventListener('click', ()=>{ $('htmlTemplate').value = sampleTemplate(); });

  $('previewBtn').addEventListener('click', ()=>{
    if (!ROWS.length) return alert('Please upload CSV/Excel first.');
    const row = {...ROWS[0]};
    const html = $('htmlTemplate').value;
    const subj = $('subject').value;

    // Inject CTA values
    row.ButtonText = $('cta-text').value || 'Click Here';
    row.ButtonUrl = $('cta-url').value || 'https://example.com';

    google.script.run.withSuccessHandler(function(out){
      $('previewHtml').innerHTML = out.html || '(empty)';
      $('previewSubject').textContent = out.subject || '(no subject)';
    }).previewCompiledHtml(html, subj, row);
  });

  $('sendBtn').addEventListener('click', ()=>startSend());
  $('send5Btn').addEventListener('click', ()=>startSend(5));

  function startSend(testCap){
    if (!ROWS.length) return alert('Upload data first.');
    if (!$('emailCol').value) return alert('Choose email column.');
    if (!$('subject').value) return alert('Enter subject.');

    // Add CTA vars to every row
    const rowsWithCTA = ROWS.map(r => ({
      ...r,
      ButtonText: $('cta-text').value || 'Click Here',
      ButtonUrl: $('cta-url').value || 'https://example.com'
    }));

    const payload = {
      rows: rowsWithCTA,
      toField: $('emailCol').value,
      subject: $('subject').value,
      htmlTemplate: $('htmlTemplate').value,
      senderName: $('senderName').value,
      replyTo: $('replyTo').value,
      attachment: ATTACHMENT,
      delayMs: Number($('delayMs').value || 800),
      maxToSend: testCap || Number($('maxToSend').value || 0) || undefined
    };

    setStatus('Sending... This tab must stay open.');
    $('sendBtn').disabled = true; $('send5Btn').disabled = true;
    google.script.run.withSuccessHandler(function(res){
      $('sendBtn').disabled = false; $('send5Btn').disabled = false;
      if (!res || !res.results) { setStatus('Done, but no results returned.'); return; }
      const ok = res.results.filter(r=>r.status==='sent').length;
      const err = res.results.filter(r=>r.status==='error').length;
      setStatus(`Completed. Sent: ${ok}, Errors: ${err}, Total attempted: ${res.attempted}/${res.total}`);
      alert('Finished! Check your Sent mailbox.');
    }).sendBulkEmails(payload);
  }

  $('htmlTemplate').value = sampleTemplate();
</script>
</body>
</html>

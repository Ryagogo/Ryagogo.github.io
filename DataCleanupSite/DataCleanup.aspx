<%@ Page Title="Run Cleanup — DataCleanup Tool" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Linq" %>
<%@ Import Namespace="System.Web" %>

<script runat="server">
    // ============================================================
    // SERVER-SIDE C# — Replicates VBA FinalDataCleanup_Revised_V64545
    // ============================================================

    protected string ProcessedCsv   = "";
    protected string StatusMessage  = "";
    protected bool   ShowResults    = false;

    // Columns to delete by original 1-based index, right-to-left: I=9, B=2, A=1
    static readonly int[] DeleteColsRightToLeft = new[] { 9, 2, 1 };

    // Default delete values matching the VBA Case list
    static int[] DefaultDeleteValues = new[] { 640, 795, 797, 717, 646, 722, 930, 941 };

    protected void Page_Load(object sender, EventArgs e)
    {
        if (IsPostBack && Request.Files["xlFile"] != null && Request.Files["xlFile"].ContentLength > 0)
        {
            try   { RunCleanup(); }
            catch (Exception ex) { StatusMessage = "ERROR: " + HttpUtility.HtmlEncode(ex.Message); }
        }
    }

    private void RunCleanup()
    {
        var file             = Request.Files["xlFile"];
        string deleteRaw     = Request.Form["deleteValues"] ?? "";
        string hasHeaderStr  = Request.Form["hasHeader"]    ?? "1";
        bool   headerRow     = hasHeaderStr == "1";

        // Parse custom delete values
        int[] deleteVals = DefaultDeleteValues;
        if (!string.IsNullOrWhiteSpace(deleteRaw))
        {
            deleteVals = deleteRaw
                .Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(v => { int x; return int.TryParse(v.Trim(), out x) ? x : -1; })
                .Where(v => v >= 0).ToArray();
        }

        // ── Read CSV ──────────────────────────────────────────────
        var rows = new List<List<string>>();
        using (var reader = new StreamReader(file.InputStream, Encoding.UTF8))
        {
            string line;
            while ((line = reader.ReadLine()) != null)
                rows.Add(ParseCsvLine(line));
        }

        if (rows.Count == 0) { StatusMessage = "ERROR: Uploaded file is empty."; return; }

        int stats_initial = rows.Count;

        // ── Step 1: Delete cols I(9), B(2), A(1) right-to-left ───
        foreach (int colIdx in DeleteColsRightToLeft)
        {
            int idx = colIdx - 1;
            foreach (var row in rows)
                if (idx < row.Count) row.RemoveAt(idx);
        }

        // ── Step 2: Move Col D (index 3) → Col A (index 0) ───────
        foreach (var row in rows)
        {
            if (row.Count >= 4)
            {
                string val = row[3];
                row.RemoveAt(3);
                row.Insert(0, val);
            }
        }

        // ── Step 3: Text-to-Columns on Col A (tab-delimited) ─────
        var expanded = new List<List<string>>();
        foreach (var row in rows)
        {
            if (row.Count > 0)
            {
                var parts  = row[0].Split('\t').ToList();
                var newRow = new List<string>(parts);
                for (int c = 1; c < row.Count; c++) newRow.Add(row[c]);
                expanded.Add(newRow);
            }
            else expanded.Add(row);
        }
        rows = expanded;

        // ── Step 4: Deduplicate by Col A ─────────────────────────
        var seen   = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var deduped = new List<List<string>>();
        for (int i = 0; i < rows.Count; i++)
        {
            if (i == 0 && headerRow) { deduped.Add(rows[i]); continue; }
            string key = rows[i].Count > 0 ? rows[i][0] : "";
            if (seen.Add(key)) deduped.Add(rows[i]);
        }
        rows = deduped;
        int stats_after_dedup = rows.Count;

        // ── Step 5: Delete rows where Col D matches delete set ────
        var deleteSet = new HashSet<int>(deleteVals);
        var filtered  = new List<List<string>>();
        for (int i = 0; i < rows.Count; i++)
        {
            if (i == 0 && headerRow) { filtered.Add(rows[i]); continue; }
            bool del = false;
            if (rows[i].Count >= 4)
            {
                int val;
                if (int.TryParse(rows[i][3].Trim(), out val) && deleteSet.Contains(val))
                    del = true;
            }
            if (!del) filtered.Add(rows[i]);
        }
        rows = filtered;

        int stats_final = rows.Count;

        // ── Serialize back to CSV ─────────────────────────────────
        var sb = new StringBuilder();
        foreach (var row in rows)
            sb.AppendLine(string.Join(",", row.Select(c => CsvEscape(c))));

        ProcessedCsv    = Convert.ToBase64String(Encoding.UTF8.GetBytes(sb.ToString()));
        ShowResults     = true;
        int dupesRemoved = stats_initial - stats_after_dedup;
        int rowsDeleted  = stats_after_dedup - stats_final;
        StatusMessage   = $"OK|{stats_initial}|{stats_final}|{dupesRemoved}|{rowsDeleted}";
    }

    private List<string> ParseCsvLine(string line)
    {
        var result = new List<string>();
        var sb     = new StringBuilder();
        bool inQ   = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"')
            {
                if (inQ && i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                else inQ = !inQ;
            }
            else if (c == ',' && !inQ) { result.Add(sb.ToString()); sb.Clear(); }
            else sb.Append(c);
        }
        result.Add(sb.ToString());
        return result;
    }

    private string CsvEscape(string val)
    {
        if (val.Contains(",") || val.Contains("\"") || val.Contains("\n"))
            return "\"" + val.Replace("\"", "\"\"") + "\"";
        return val;
    }
</script>

<asp:Content ID="HeadContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style>
        /* ── Page-scoped styles ────────────────────────────────── */
        :root {
            --accent-dim: #00e5a018;
        }

        .page-header {
            margin-bottom: 40px;
        }

        .page-header h1 {
            font-family: var(--mono);
            font-size: 24px;
            font-weight: 700;
            letter-spacing: -0.4px;
            color: var(--text);
            margin-bottom: 6px;
        }

        .page-header p {
            font-size: 14px;
            color: var(--text-dim);
        }

        /* Pipeline indicator */
        .pipeline {
            display: flex;
            gap: 0;
            margin-bottom: 36px;
            overflow-x: auto;
            padding-bottom: 4px;
        }

        .step {
            display: flex;
            align-items: center;
            gap: 7px;
            background: var(--surface);
            border: 1px solid var(--border);
            padding: 9px 15px;
            font-family: var(--mono);
            font-size: 11px;
            color: var(--text-dim);
            white-space: nowrap;
        }

        .step:first-child { border-radius: 6px 0 0 6px; }
        .step:last-child  { border-radius: 0 6px 6px 0; }
        .step + .step     { border-left: none; }

        .step .num {
            width: 18px;
            height: 18px;
            background: var(--border);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 10px;
            font-weight: 700;
            color: var(--text-dim);
        }

        .step.active { background: var(--accent-dim); border-color: var(--accent); color: var(--accent); }
        .step.active .num { background: var(--accent); color: #000; }
        .step.done   { border-color: #2a4a3a; color: #4a9a70; }
        .step.done .num { background: #2a4a3a; color: #4a9a70; }

        /* Upload zone */
        .upload-zone {
            border: 2px dashed var(--border);
            border-radius: 8px;
            padding: 52px 24px;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s;
            position: relative;
            overflow: hidden;
        }

        .upload-zone:hover,
        .upload-zone.dragover { border-color: var(--accent); background: var(--accent-dim); }

        .upload-zone input[type="file"] {
            position: absolute;
            inset: 0;
            opacity: 0;
            cursor: pointer;
            width: 100%;
            height: 100%;
        }

        .upload-icon { font-size: 44px; margin-bottom: 12px; display: block; }

        .upload-zone h3 {
            font-family: var(--mono);
            font-size: 14px;
            color: var(--text);
            margin-bottom: 6px;
        }

        .upload-zone p { font-size: 13px; color: var(--text-dim); }

        .file-pill {
            display: none;
            align-items: center;
            gap: 10px;
            background: var(--accent-dim);
            border: 1px solid var(--accent);
            border-radius: 6px;
            padding: 11px 16px;
            margin-top: 14px;
            font-family: var(--mono);
            font-size: 13px;
            color: var(--accent);
        }

        /* Config */
        .config-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
        @media (max-width: 600px) { .config-grid { grid-template-columns: 1fr; } }

        .cfg-label {
            display: block;
            font-family: var(--mono);
            font-size: 11px;
            color: var(--text-dim);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 7px;
        }

        .cfg-input {
            width: 100%;
            background: var(--surface2);
            border: 1px solid var(--border);
            border-radius: 5px;
            padding: 10px 12px;
            font-family: var(--mono);
            font-size: 13px;
            color: var(--text);
            outline: none;
            transition: border-color 0.2s;
        }

        .cfg-input:focus { border-color: var(--accent); }

        /* Delete tags */
        .delete-tags { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; }

        .del-tag {
            background: var(--surface2);
            border: 1px solid var(--border);
            border-radius: 4px;
            padding: 5px 10px;
            font-family: var(--mono);
            font-size: 12px;
            color: #ff6b35;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .del-tag button {
            background: none;
            border: none;
            color: var(--text-dim);
            cursor: pointer;
            font-size: 14px;
            line-height: 1;
            padding: 0;
            transition: color 0.15s;
        }

        .del-tag button:hover { color: var(--danger); }

        .add-row { display: flex; gap: 8px; margin-top: 12px; }

        .add-row input {
            flex: 1;
            background: var(--surface2);
            border: 1px solid var(--border);
            border-radius: 5px;
            padding: 8px 12px;
            font-family: var(--mono);
            font-size: 13px;
            color: var(--text);
            outline: none;
            transition: border-color 0.2s;
        }

        .add-row input:focus { border-color: var(--accent); }

        /* Progress */
        .progress-wrap { display: none; margin-bottom: 24px; }

        .progress-meta {
            display: flex;
            justify-content: space-between;
            font-family: var(--mono);
            font-size: 11px;
            color: var(--text-dim);
            margin-bottom: 8px;
        }

        .progress-bar { height: 4px; background: var(--border); border-radius: 2px; overflow: hidden; }

        .progress-fill {
            height: 100%;
            background: var(--accent);
            border-radius: 2px;
            width: 0%;
            transition: width 0.3s ease;
        }

        .log-box {
            background: var(--surface2);
            border: 1px solid var(--border);
            border-radius: 6px;
            padding: 16px;
            font-family: var(--mono);
            font-size: 12px;
            color: var(--text-dim);
            max-height: 160px;
            overflow-y: auto;
            line-height: 1.9;
            display: none;
            margin-top: 16px;
        }

        .log-ok   { color: var(--success); }
        .log-warn { color: #ff6b35; }
        .log-err  { color: var(--danger); }

        /* Stats */
        .stats-row {
            display: none;
            grid-template-columns: repeat(4, 1fr);
            gap: 12px;
            margin-bottom: 24px;
        }

        @media (max-width: 640px) { .stats-row { grid-template-columns: 1fr 1fr; } }

        .stat-box {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 18px;
            text-align: center;
        }

        .stat-num {
            font-family: var(--mono);
            font-size: 26px;
            font-weight: 700;
            color: var(--accent);
        }

        .stat-label {
            font-size: 11px;
            color: var(--text-dim);
            margin-top: 5px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
        }

        /* Table preview */
        .preview-wrap { display: none; }

        .table-scroll {
            overflow-x: auto;
            border: 1px solid var(--border);
            border-radius: 8px;
        }

        table { width: 100%; border-collapse: collapse; font-family: var(--mono); font-size: 12px; }

        thead tr { background: var(--surface2); border-bottom: 2px solid var(--accent); }

        th {
            padding: 10px 14px;
            text-align: left;
            color: var(--accent);
            font-weight: 700;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            white-space: nowrap;
        }

        tbody tr { border-bottom: 1px solid var(--border); transition: background 0.1s; }
        tbody tr:hover { background: var(--surface2); }
        tbody tr:last-child { border-bottom: none; }

        td {
            padding: 9px 14px;
            color: var(--text);
            max-width: 200px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .row-count {
            font-family: var(--mono);
            font-size: 12px;
            color: var(--text-dim);
            margin-top: 12px;
        }

        .row-count span { color: var(--accent); }

        /* Action bar */
        .action-bar {
            display: flex;
            gap: 12px;
            align-items: center;
            flex-wrap: wrap;
            margin-top: 24px;
        }

        .ready-label {
            font-family: var(--mono);
            font-size: 11px;
            color: var(--text-dim);
        }

        .hidden { display: none; }
    </style>
</asp:Content>

<asp:Content ID="MainContent" ContentPlaceHolderID="MainContent" runat="server">
    <div class="wrapper">

        <!-- Page header -->
        <div class="page-header">
            <h1>Data Cleanup Pipeline</h1>
            <p>Upload a CSV export from Excel to run all five cleanup steps server-side.</p>
        </div>

        <!-- Pipeline indicator -->
        <div class="pipeline">
            <div class="step active" id="step1"><span class="num">1</span> Upload</div>
            <div class="step" id="step2"><span class="num">2</span> Configure</div>
            <div class="step" id="step3"><span class="num">3</span> Process</div>
            <div class="step" id="step4"><span class="num">4</span> Export</div>
        </div>

        <!-- Error banner -->
        <% if (!string.IsNullOrEmpty(StatusMessage) && StatusMessage.StartsWith("ERROR")) { %>
        <div class="card" style="border-color:var(--danger);margin-bottom:24px;">
            <div class="card-title" style="color:var(--danger);">Processing Error</div>
            <p style="font-family:var(--mono);font-size:13px;color:var(--danger);"><%: StatusMessage %></p>
        </div>
        <% } %>

        <!-- Stats (populated by JS after server response) -->
        <div class="stats-row" id="statsRow">
            <div class="stat-box"><div class="stat-num" id="statInitial">—</div><div class="stat-label">Initial Rows</div></div>
            <div class="stat-box"><div class="stat-num" id="statFinal">—</div><div class="stat-label">Final Rows</div></div>
            <div class="stat-box"><div class="stat-num" id="statDupes">—</div><div class="stat-label">Dupes Removed</div></div>
            <div class="stat-box"><div class="stat-num" id="statDeleted">—</div><div class="stat-label">Rows Deleted</div></div>
        </div>

        <!-- Form -->
        <form method="post" enctype="multipart/form-data" id="mainForm">

            <!-- Upload -->
            <div class="card">
                <div class="card-title">Upload Data File</div>
                <div class="upload-zone" id="uploadZone">
                    <input type="file" name="xlFile" id="xlFile" accept=".csv,.txt" />
                    <span class="upload-icon">📂</span>
                    <h3>Drop your CSV here or click to browse</h3>
                    <p>In Excel: <strong>File → Save As → CSV (Comma delimited)</strong> first</p>
                </div>
                <div class="file-pill" id="filePill">
                    <span>📄</span><span id="fileName">—</span>
                </div>
            </div>

            <!-- Config -->
            <div class="card">
                <div class="card-title">Pipeline Configuration</div>
                <div class="config-grid">
                    <div>
                        <label class="cfg-label">Header Row</label>
                        <select name="hasHeader" class="cfg-input">
                            <option value="1">Yes — first row is header</option>
                            <option value="0">No header row</option>
                        </select>
                    </div>
                    <div>
                        <label class="cfg-label">Columns Deleted (fixed)</label>
                        <input class="cfg-input" type="text" value="A, B, I  →  right-to-left" readonly />
                    </div>
                    <div>
                        <label class="cfg-label">Col D → Col A (fixed)</label>
                        <input class="cfg-input" type="text" value="Enabled" readonly />
                    </div>
                    <div>
                        <label class="cfg-label">Dedup Key (fixed)</label>
                        <input class="cfg-input" type="text" value="Column A after restructure" readonly />
                    </div>
                </div>

                <!-- Delete value editor -->
                <div style="margin-top:20px;">
                    <label class="cfg-label">Delete Rows Where Column D Equals</label>
                    <div class="delete-tags" id="deleteTagsWrap"></div>
                    <div class="add-row">
                        <input type="text" id="addValInput" placeholder="Add value, e.g. 640" />
                        <button type="button" class="btn btn-secondary btn-sm" onclick="addDeleteVal()">+ Add</button>
                        <button type="button" class="btn btn-secondary btn-sm" onclick="resetDeleteVals()">↺ Reset</button>
                    </div>
                    <input type="hidden" name="deleteValues" id="deleteValuesField" />
                </div>
            </div>

            <!-- Progress -->
            <div class="progress-wrap" id="progressWrap">
                <div class="progress-meta">
                    <span id="progressLabel">Processing…</span>
                    <span id="progressPct">0%</span>
                </div>
                <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
                <div class="log-box" id="logBox"></div>
            </div>

            <!-- Action bar -->
            <div class="action-bar">
                <button type="button" class="btn btn-primary" id="runBtn" onclick="runProcess()" disabled>
                    ▶ Run Cleanup
                </button>
                <span class="ready-label" id="readyLabel">Upload a file to begin</span>
            </div>

        </form>

        <!-- Results preview -->
        <div class="preview-wrap" id="previewWrap">
            <div class="card" style="margin-top:24px;">
                <div class="card-title">Output Preview</div>
                <div class="table-scroll">
                    <table>
                        <thead id="previewHead"></thead>
                        <tbody id="previewBody"></tbody>
                    </table>
                </div>
                <div class="row-count" id="rowCount"></div>
                <div class="action-bar">
                    <button type="button" class="btn btn-primary" onclick="downloadCsv()">⬇ Download CSV</button>
                    <button type="button" class="btn btn-secondary" onclick="showExcelTip()">📊 Re-import to Excel</button>
                    <button type="button" class="btn btn-secondary" onclick="location.reload()">↺ New File</button>
                </div>
            </div>
        </div>

    </div>

    <!-- Server output slots -->
    <div id="serverCsv"    class="hidden"><%: ProcessedCsv  %></div>
    <div id="serverStatus" class="hidden"><%: StatusMessage %></div>
</asp:Content>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="ScriptContent" runat="server">
<script>
    // ── Delete value state ──────────────────────────────
    let deleteVals = [640, 795, 797, 717, 646, 722, 930, 941];
    renderTags();

    function renderTags() {
        const wrap = document.getElementById('deleteTagsWrap');
        wrap.innerHTML = deleteVals.map((v, i) =>
            `<span class="del-tag">${v}<button onclick="removeVal(${i})" title="Remove">×</button></span>`
        ).join('');
        document.getElementById('deleteValuesField').value = deleteVals.join(',');
    }

    function addDeleteVal() {
        const inp = document.getElementById('addValInput');
        const v   = parseInt(inp.value.trim(), 10);
        if (!isNaN(v) && !deleteVals.includes(v)) { deleteVals.push(v); renderTags(); }
        inp.value = '';
    }

    function removeVal(i) { deleteVals.splice(i, 1); renderTags(); }

    function resetDeleteVals() { deleteVals = [640,795,797,717,646,722,930,941]; renderTags(); }

    document.getElementById('addValInput').addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); addDeleteVal(); }
    });

    // ── File selection ──────────────────────────────────
    document.getElementById('xlFile').addEventListener('change', function () {
        if (this.files[0]) {
            document.getElementById('fileName').textContent = this.files[0].name;
            document.getElementById('filePill').style.display = 'flex';
            document.getElementById('runBtn').disabled = false;
            document.getElementById('readyLabel').textContent = this.files[0].name + ' ready';
            setStep(2);
        }
    });

    const zone = document.getElementById('uploadZone');
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop',     e => { e.preventDefault(); zone.classList.remove('dragover'); });

    // ── Step indicator ──────────────────────────────────
    function setStep(n) {
        for (let i = 1; i <= 4; i++) {
            const el = document.getElementById('step' + i);
            el.className = 'step' + (i < n ? ' done' : i === n ? ' active' : '');
        }
    }

    // ── Progress simulation ─────────────────────────────
    let pAnim;

    function runProcess() {
        document.getElementById('runBtn').disabled = true;
        document.getElementById('progressWrap').style.display = 'block';
        document.getElementById('logBox').style.display = 'block';
        document.getElementById('logBox').innerHTML = '';
        setStep(3);

        let pct = 0;
        pAnim = setInterval(() => {
            pct = Math.min(pct + Math.random() * 9, 88);
            setProgress(pct, 'Processing…');
        }, 200);

        logLine('Uploading file…', 'ok');
        document.getElementById('mainForm').submit();
    }

    function setProgress(pct, label) {
        document.getElementById('progressFill').style.width = pct + '%';
        document.getElementById('progressLabel').textContent = label;
        document.getElementById('progressPct').textContent = Math.round(pct) + '%';
    }

    function logLine(msg, cls) {
        const b = document.getElementById('logBox');
        b.innerHTML += `<div class="log-${cls || ''}">${msg}</div>`;
        b.scrollTop = b.scrollHeight;
    }

    // ── On load: render server results if present ───────
    window.addEventListener('DOMContentLoaded', () => {
        const csvB64 = document.getElementById('serverCsv').textContent.trim();
        const status = document.getElementById('serverStatus').textContent.trim();

        if (csvB64 && status.startsWith('OK')) {
            clearInterval(pAnim);
            setProgress(100, 'Complete ✓');
            document.getElementById('progressWrap').style.display = 'block';
            document.getElementById('logBox').style.display = 'block';

            const parts = status.split('|');
            document.getElementById('statInitial').textContent = parts[1] || '—';
            document.getElementById('statFinal').textContent   = parts[2] || '—';
            document.getElementById('statDupes').textContent   = parts[3] || '—';
            document.getElementById('statDeleted').textContent = parts[4] || '—';
            document.getElementById('statsRow').style.display  = 'grid';

            logLine('✓ Columns A, B, I deleted (right-to-left)', 'ok');
            logLine('✓ Column D moved to Column A', 'ok');
            logLine('✓ Text-to-columns applied on Col A', 'ok');
            logLine('✓ Duplicates removed (key: Col A)', 'ok');
            logLine('✓ Rows deleted by Col D value filter', 'ok');
            logLine(`✓ Done — ${parts[2]} rows remaining`, 'ok');

            window._cleanedCsv = atob(csvB64);
            renderPreview(window._cleanedCsv);
            document.getElementById('previewWrap').style.display = 'block';
            setStep(4);
            DataCleanup.showToast('Cleanup complete — ' + parts[2] + ' rows remaining');
        }
    });

    // ── Preview table ───────────────────────────────────
    function renderPreview(csv) {
        const lines   = csv.trim().split('\n');
        const preview = lines.slice(0, 51);
        const head    = document.getElementById('previewHead');
        const body    = document.getElementById('previewBody');

        if (!preview.length) return;

        const hdrs = parseCsv(preview[0]);
        head.innerHTML = '<tr>' + hdrs.map(h => `<th>${esc(h)}</th>`).join('') + '</tr>';
        body.innerHTML = preview.slice(1).map(l =>
            '<tr>' + parseCsv(l).map(c => `<td title="${esc(c)}">${esc(c)}</td>`).join('') + '</tr>'
        ).join('');

        const total = lines.length - 1;
        document.getElementById('rowCount').innerHTML =
            `Showing <span>${Math.min(50, total)}</span> of <span>${total}</span> data rows` +
            (total > 50 ? ' — download for full dataset' : '');
    }

    function parseCsv(line) {
        const res = [], re = /("(?:[^"]|"")*"|[^,]*)/g;
        let m;
        while ((m = re.exec(line)) !== null) {
            if (m.index === line.length) break;
            let v = m[1];
            if (v.startsWith('"') && v.endsWith('"')) v = v.slice(1,-1).replace(/""/g,'"');
            res.push(v);
        }
        return res;
    }

    function esc(s) {
        return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    }

    // ── Download ────────────────────────────────────────
    function downloadCsv() {
        if (!window._cleanedCsv) return;
        const blob = new Blob([window._cleanedCsv], { type: 'text/csv;charset=utf-8;' });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement('a');
        a.href = url; a.download = 'cleaned_data.csv'; a.click();
        URL.revokeObjectURL(url);
        DataCleanup.showToast('Downloaded — use Data → From Text/CSV in Excel to re-import');
    }

    function showExcelTip() {
        DataCleanup.showToast('Excel: Data → Get Data → From Text/CSV → select cleaned_data.csv → Load');
    }
</script>
</asp:Content>

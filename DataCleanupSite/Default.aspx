<%@ Page Title="Home — DataCleanup Tool" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" %>

<asp:Content ID="TitleContent" ContentPlaceHolderID="TitleContent" runat="server">
    Home — DataCleanup Tool
</asp:Content>

<asp:Content ID="HeadContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style>
        .hero { padding: 72px 0 56px; }

        .hero-eyebrow {
            font-family: var(--mono);
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 2px;
            color: var(--accent);
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .hero-eyebrow::before {
            content: '';
            display: inline-block;
            width: 24px;
            height: 2px;
            background: var(--accent);
        }

        .hero h1 {
            font-family: var(--mono);
            font-size: clamp(28px, 4vw, 48px);
            font-weight: 700;
            line-height: 1.15;
            letter-spacing: -1px;
            color: var(--text);
            max-width: 680px;
            margin-bottom: 20px;
        }

        .hero h1 span { color: var(--accent); }

        .hero-sub {
            font-size: 16px;
            color: var(--text-dim);
            max-width: 560px;
            line-height: 1.7;
            margin-bottom: 36px;
        }

        .hero-actions { display: flex; gap: 12px; flex-wrap: wrap; }

        /* Pipeline diagram */
        .pipeline-section { margin-bottom: 48px; }

        .pipeline-title {
            font-family: var(--mono);
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            color: var(--text-dim);
            margin-bottom: 20px;
        }

        .pipeline-steps {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 0;
            position: relative;
        }

        @media (max-width: 700px) {
            .pipeline-steps { grid-template-columns: 1fr 1fr; gap: 12px; }
            .pipe-arrow { display: none; }
        }

        .pipe-step {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 20px 16px;
            text-align: center;
            position: relative;
        }

        .pipe-step:not(:last-child)::after {
            content: '→';
            position: absolute;
            right: -14px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--accent);
            font-family: var(--mono);
            font-size: 16px;
            z-index: 2;
        }

        @media (max-width: 700px) {
            .pipe-step::after { display: none; }
        }

        .pipe-num {
            width: 28px;
            height: 28px;
            background: var(--accent);
            color: #000;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: var(--mono);
            font-size: 12px;
            font-weight: 700;
            margin: 0 auto 10px;
        }

        .pipe-label {
            font-family: var(--mono);
            font-size: 11px;
            font-weight: 600;
            color: var(--text);
            margin-bottom: 4px;
        }

        .pipe-desc {
            font-size: 11px;
            color: var(--text-dim);
            line-height: 1.5;
        }

        /* Feature grid */
        .feature-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 16px;
            margin-bottom: 48px;
        }

        @media (max-width: 800px) { .feature-grid { grid-template-columns: 1fr 1fr; } }
        @media (max-width: 500px) { .feature-grid { grid-template-columns: 1fr; } }

        .feature-card {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 24px;
            transition: border-color 0.2s, transform 0.2s;
        }

        .feature-card:hover { border-color: var(--accent); transform: translateY(-2px); }

        .feat-icon {
            font-size: 28px;
            margin-bottom: 12px;
            display: block;
        }

        .feat-title {
            font-family: var(--mono);
            font-size: 13px;
            font-weight: 700;
            color: var(--text);
            margin-bottom: 6px;
        }

        .feat-desc {
            font-size: 13px;
            color: var(--text-dim);
            line-height: 1.6;
        }

        /* Quick start */
        .qs-list {
            list-style: none;
            counter-reset: qs;
        }

        .qs-list li {
            counter-increment: qs;
            display: flex;
            align-items: flex-start;
            gap: 14px;
            padding: 14px 0;
            border-bottom: 1px solid var(--border);
        }

        .qs-list li:last-child { border-bottom: none; }

        .qs-num {
            width: 24px;
            height: 24px;
            background: var(--accent-dim);
            border: 1px solid var(--accent);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: var(--mono);
            font-size: 11px;
            font-weight: 700;
            color: var(--accent);
            flex-shrink: 0;
            margin-top: 1px;
        }

        .qs-text strong {
            display: block;
            font-family: var(--mono);
            font-size: 13px;
            color: var(--text);
            margin-bottom: 3px;
        }

        .qs-text span {
            font-size: 13px;
            color: var(--text-dim);
        }

        code {
            background: var(--surface2);
            border: 1px solid var(--border);
            border-radius: 4px;
            padding: 2px 6px;
            font-family: var(--mono);
            font-size: 12px;
            color: var(--accent2);
        }
    </style>
</asp:Content>

<asp:Content ID="MainContent" ContentPlaceHolderID="MainContent" runat="server">
    <div class="wrapper">

        <!-- Hero -->
        <section class="hero">
            <div class="hero-eyebrow">Excel Data Processing</div>
            <h1>Clean your spreadsheet data<br /><span>server-side, instantly.</span></h1>
            <p class="hero-sub">
                Upload a CSV export from Excel and the pipeline automatically deletes columns,
                restructures data, removes duplicates, and filters rows — then hands you back
                a clean file ready to re-import.
            </p>
            <div class="hero-actions">
                <a href="~/DataCleanup.aspx" runat="server" class="btn btn-primary">▶ Start Cleanup</a>
                <a href="#quickstart" class="btn btn-secondary">How it works ↓</a>
            </div>
        </section>

        <!-- Pipeline -->
        <div class="pipeline-section">
            <div class="pipeline-title">Cleanup Pipeline</div>
            <div class="pipeline-steps">
                <div class="pipe-step">
                    <div class="pipe-num">1</div>
                    <div class="pipe-label">Delete Cols</div>
                    <div class="pipe-desc">Removes original columns A, B, I (right-to-left)</div>
                </div>
                <div class="pipe-step">
                    <div class="pipe-num">2</div>
                    <div class="pipe-label">Move Col D→A</div>
                    <div class="pipe-desc">Cuts column D and inserts it at position A</div>
                </div>
                <div class="pipe-step">
                    <div class="pipe-num">3</div>
                    <div class="pipe-label">Text→Columns</div>
                    <div class="pipe-desc">Splits tab-delimited values in column A</div>
                </div>
                <div class="pipe-step">
                    <div class="pipe-num">4</div>
                    <div class="pipe-label">Deduplicate</div>
                    <div class="pipe-desc">Removes duplicate rows keyed on column A</div>
                </div>
                <div class="pipe-step">
                    <div class="pipe-num">5</div>
                    <div class="pipe-label">Filter Rows</div>
                    <div class="pipe-desc">Deletes rows matching values in column D</div>
                </div>
            </div>
        </div>

        <!-- Features -->
        <div class="feature-grid">
            <div class="feature-card">
                <span class="feat-icon">⚡</span>
                <div class="feat-title">Faster than VBA</div>
                <div class="feat-desc">No Excel instance, no screen repaint, no COM overhead. Pure server-side C# finishes in milliseconds.</div>
            </div>
            <div class="feature-card">
                <span class="feat-icon">🔧</span>
                <div class="feat-title">Configurable Filters</div>
                <div class="feat-desc">Add or remove deletion values on the fly. Defaults match the original macro: 640, 795, 797, 717…</div>
            </div>
            <div class="feature-card">
                <span class="feat-icon">👁</span>
                <div class="feat-title">Live Preview</div>
                <div class="feat-desc">See the first 50 rows of cleaned data in-browser before downloading. Stats show exactly what changed.</div>
            </div>
            <div class="feature-card">
                <span class="feat-icon">📥</span>
                <div class="feat-title">CSV Round-trip</div>
                <div class="feat-desc">Export from Excel as CSV, clean here, re-import with Data → From Text/CSV. No macros needed.</div>
            </div>
            <div class="feature-card">
                <span class="feat-icon">🔒</span>
                <div class="feat-title">No Data Stored</div>
                <div class="feat-desc">Files are processed in-memory and never written to disk. Nothing persists after your session ends.</div>
            </div>
            <div class="feature-card">
                <span class="feat-icon">🖥</span>
                <div class="feat-title">Drop-in Deploy</div>
                <div class="feat-desc">Single ASPX file, no NuGet packages, no database. Works on any IIS host running .NET 4.x.</div>
            </div>
        </div>

        <!-- Quick start -->
        <div class="card" id="quickstart">
            <div class="card-title">Quick Start</div>
            <ol class="qs-list">
                <li>
                    <span class="qs-num">1</span>
                    <div class="qs-text">
                        <strong>Export from Excel</strong>
                        <span>Open your workbook → <code>File → Save As → CSV (Comma delimited)</code>. Save to desktop.</span>
                    </div>
                </li>
                <li>
                    <span class="qs-num">2</span>
                    <div class="qs-text">
                        <strong>Upload the CSV</strong>
                        <span>Click <strong>Start Cleanup</strong>, drag your CSV onto the upload zone, or click to browse.</span>
                    </div>
                </li>
                <li>
                    <span class="qs-num">3</span>
                    <div class="qs-text">
                        <strong>Confirm settings</strong>
                        <span>Check the delete-value list. Add or remove values to match your data. Verify the header-row toggle.</span>
                    </div>
                </li>
                <li>
                    <span class="qs-num">4</span>
                    <div class="qs-text">
                        <strong>Run &amp; Download</strong>
                        <span>Click <strong>Run Cleanup</strong>. Review the preview and stats, then click <strong>Download CSV</strong>.</span>
                    </div>
                </li>
                <li>
                    <span class="qs-num">5</span>
                    <div class="qs-text">
                        <strong>Re-import to Excel</strong>
                        <span>In Excel: <code>Data → Get Data → From Text/CSV</code> → select <code>cleaned_data.csv</code> → Load.</span>
                    </div>
                </li>
            </ol>
        </div>

    </div>
</asp:Content>

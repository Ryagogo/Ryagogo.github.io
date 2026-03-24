# DataCleanup Tool — ASP.NET Web Forms Site

Server-side replacement for the `FinalDataCleanup_Revised_V64545` VBA macro.
Upload a CSV, run all five cleanup steps in C#, download the result.

---

## Project Structure

```
DataCleanupSite/
├── DataCleanupSite.csproj   ← Visual Studio project file
├── Web.config               ← ASP.NET + IIS configuration
├── Global.asax              ← Application lifecycle hooks
├── Site.Master              ← Shared layout (nav, footer, CSS/JS refs)
├── Default.aspx             ← Home / landing page
├── DataCleanup.aspx         ← Main cleanup tool (drop your ASPX here)
├── Error.aspx               ← Friendly error page
├── Content/
│   └── site.css             ← Shared stylesheet
└── Scripts/
    └── site.js              ← Shared JS utilities (toast, nav highlight)
```

---

## Requirements

| Requirement         | Version          |
|---------------------|------------------|
| .NET Framework      | 4.8 (or 4.5+)    |
| IIS / IIS Express   | 8.0+             |
| Visual Studio       | 2019 / 2022 (optional) |

No NuGet packages. No database. No external dependencies.

---

## Deployment — IIS (Production)

1. **Copy the site folder** to your IIS web root, e.g. `C:\inetpub\wwwroot\DataCleanup`
2. **Open IIS Manager** → Sites → Add Website  
   - Physical path: `C:\inetpub\wwwroot\DataCleanup`  
   - Application pool: `.NET v4.8` (Classic or Integrated, both work)
3. **Verify Web.config** is present — IIS reads it automatically
4. Browse to `http://your-server/DataCleanup/`

> **Upload size**: `Web.config` allows 50 MB uploads by default.  
> Adjust `maxRequestLength` (KB) and `maxAllowedContentLength` (bytes) if needed.

---

## Deployment — Visual Studio / IIS Express (Development)

1. Open `DataCleanupSite.csproj` in Visual Studio 2019/2022
2. Press **F5** or click **IIS Express** — the site launches at `https://localhost:PORT/`

---

## Dropping In the ASPX File

The `DataCleanup.aspx` page is self-contained: all C# processing lives in its
`<script runat="server">` block. To update the cleanup logic:

1. Open `DataCleanup.aspx` in any text editor
2. Edit the `RunCleanup()` method inside `<script runat="server">`
3. Save and reload — no compile step needed for Web Forms inline code

---

## Cleanup Pipeline (mirrors the VBA macro)

| Step | Action | VBA Equivalent |
|------|--------|----------------|
| 1 | Delete columns A, B, I (right-to-left) | `ws.Columns("I:I").Delete` … |
| 2 | Move column D → column A | `ws.Columns("D:D").Cut` + Insert |
| 3 | Text-to-columns on new col A (tab-split) | `TextToColumns … Tab:=True` |
| 4 | Deduplicate by column A | `RemoveDuplicates Columns:=1` |
| 5 | Delete rows where col D ∈ filter set | `Select Case … Case 640, 795 …` |

---

## Excel Round-trip

**Export:**  
Excel → `File → Save As → CSV (Comma delimited)` → upload here

**Re-import after cleanup:**  
Excel → `Data → Get Data → From Text/CSV` → select `cleaned_data.csv` → Load

---

## Configuration

All filter values are editable in the browser UI before each run.  
Defaults match the original macro: `640, 795, 797, 717, 646, 722, 930, 941`

To change defaults permanently, edit `DefaultDeleteValues` in `DataCleanup.aspx`:

```csharp
static int[] DefaultDeleteValues = new[] { 640, 795, 797, 717, 646, 722, 930, 941 };
```

---

## Security Notes

- Files are processed **in memory only** — nothing is written to disk
- No data persists after the HTTP response is sent
- `customErrors mode="RemoteOnly"` hides stack traces from remote users
- Consider adding Windows Authentication if the tool should be internal-only

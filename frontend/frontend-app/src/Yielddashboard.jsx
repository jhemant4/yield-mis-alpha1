import { useState, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";

const REQUIRED_COLS = [
  "Loan Number","ID","Product Type","Refer to Credit Central","Sub Product Type","Product Category","Product Line",
  "UBL Login Flag","UBL Product Line","Dept","MLAP Segment","MSME Type","HL Type","Samarthya","Gruh Setu Flag",
  "Chq Cancelled Flag","Cover Type","Gruh Pravesh","Stage","CC Stage","DISB_STAGE","Real Stage",
  "Lead Created Date","Pre Login Date","Query Raise Date","Login Fee","Decision Date","Login Date",
  "Loan Sanction Date","Loan Requested Amt","First Sanctioned Date","Sanctioned Amt","Sanctioned Amt With LI GI",
  "First LD Gen Date","Handover Date","Handover Seq","Handover Amt","Payment Seq","Payment Type",
  "Disb Date","Disb Amt","Disbursed Excluding INBT","Disb ROI","Yield","Disb Seq","Fin Is Active",
  "PF Amt","PF Amt on Lan Gen","LMS PF Amt","LMS PF Amt GST","LOS PF Amt WT GST","LOS PF Amt",
  "LI Amt","LI Sum Assured","LI Varient","LI Policy Period In Months","GI Amt","Sourcing DST ID","DST ID",
  "Employee Code","Credit Code","Name","Designation","Department","Sub Department","Reporting Manager",
  "RM ID","Sales Manager ID","Current Manager","legal Entity","Channel Partner Code","Channel Partner Account Name",
  "Partner Type CP","Channel Partner Type","Channel Partner","VLE Name","VLE District","VLE State",
  "Lead Source","Lead Sub Source","BRANCH_CODE","Branch","COUNTRY","Zone","Geo","Cluster","City","City Tier",
  "State","Branch Phase","Branch Tier","Actual ROI","Loan Channel","Score Ventile","Sanctioned Loan Tenure",
  "Sanctioned ROI","Vertical","Status","SOR Source System","Product Campaign","Product Category","Product Variant",
  "Product Category Mudra","Sammaan Flag","Subsequent Tranch Count","Sales DID Date","PSL","LWD","LI GI Amt",
  "Last LD Generation Date","Last DCM Date","FOIR","First Disb Date","Data Source","Customer Entity","Program",
  "Transaction Sub Type","Customer Profile","IS APF","Early","Loan Info Sub Product Type",
  "Collateral Property Type","Property Sub Type","Platinum Loan Flag","Mitra Flag","Saathi Flag","Aarambh HL"
];

const FILTER_FIELDS = ["Zone","Geo","State","Branch","Vertical","Product Type","Product Line","Channel Partner","Lead Source","Status"];

function Badge({ children, color = "blue" }) {
  const colors = {
    blue: { bg: "#E6F1FB", text: "#185FA5", border: "#B5D4F4" },
    green: { bg: "#EAF3DE", text: "#3B6D11", border: "#C0DD97" },
    amber: { bg: "#FAEEDA", text: "#854F0B", border: "#FAC775" },
    purple: { bg: "#EEEDFE", text: "#534AB7", border: "#CECBF6" },
    teal: { bg: "#E1F5EE", text: "#0F6E56", border: "#9FE1CB" },
  };
  const c = colors[color] || colors.blue;
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 4,
      background: c.bg, color: c.text, border: `0.5px solid ${c.border}`,
      borderRadius: 6, fontSize: 11, fontWeight: 500, padding: "2px 8px",
      letterSpacing: "0.01em"
    }}>{children}</span>
  );
}

function MetricCard({ label, value, sub, color = "#185FA5" }) {
  return (
    <div style={{
      background: "var(--color-background-secondary)",
      borderRadius: 10, padding: "14px 16px",
      display: "flex", flexDirection: "column", gap: 4,
      minWidth: 0
    }}>
      <span style={{ fontSize: 11, color: "var(--color-text-secondary)", fontWeight: 500, letterSpacing: "0.04em", textTransform: "uppercase" }}>{label}</span>
      <span style={{ fontSize: 22, fontWeight: 600, color, lineHeight: 1.1 }}>{value}</span>
      {sub && <span style={{ fontSize: 11, color: "var(--color-text-tertiary)" }}>{sub}</span>}
    </div>
  );
}

export default function YieldDashboard() {
  const [rawData, setRawData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [fileName, setFileName] = useState("");
  const [loading, setLoading] = useState(false);
  const [tab, setTab] = useState("preview"); // preview | yield
  const [filters, setFilters] = useState({});
  const [filterSearch, setFilterSearch] = useState({});
  const [page, setPage] = useState(1);
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef();
  const ROWS_PER_PAGE = 50;

  const processFile = useCallback((file) => {
    if (!file) return;
    setLoading(true);
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (raw.length === 0) { setLoading(false); return; }
        // Normalize all column names: trim spaces
        const normalized = raw.map(row => {
          const clean = {};
          Object.keys(row).forEach(k => { clean[k.trim()] = row[k]; });
          return clean;
        });
        const cols = Object.keys(normalized[0]);
        setColumns(cols);
        setRawData(normalized);
        setFilters({});
        setPage(1);
        setTab("preview");
      } catch (err) { console.error(err); }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault(); setDragOver(false);
    const f = e.dataTransfer.files[0];
    if (f && (f.name.endsWith(".xlsx") || f.name.endsWith(".xls") || f.name.endsWith(".csv"))) processFile(f);
  }, [processFile]);

  // ─── UTILITY ────────────────────────────────────────────────────────────────
  // Fuzzy column finder: normalize name OR fall back to column index
  const fuzzyCol = (row, keys, colIdx, ...candidates) => {
    const norm = s => s.toLowerCase().replace(/[\s_\-\.]/g, "");
    for (const c of candidates) {
      const found = keys.find(k => norm(k) === norm(c));
      if (found !== undefined && row[found] !== "" && row[found] !== undefined) return row[found];
    }
    if (colIdx !== null && keys[colIdx] !== undefined) return row[keys[colIdx]];
    return 0;
  };

  // ─── YIELD SCORE (spec §10) ──────────────────────────────────────────────────
  // Gap in bps vs target → Performance Score
  const getYieldScore = (yieldPct, targetYieldPct) => {
    if (!targetYieldPct || targetYieldPct === 0) return { score: 100, band: "No Target", bpsGap: 0 };
    const bpsGap = (targetYieldPct - yieldPct) * 100; // 1% = 100 bps
    if (bpsGap <= 0)       return { score: 125, band: "≥ Target",   bpsGap: Math.round(bpsGap) };
    if (bpsGap < 10)       return { score: 100, band: "< 10 bps",  bpsGap: Math.round(bpsGap) };
    if (bpsGap < 20)       return { score: 75,  band: "10–20 bps", bpsGap: Math.round(bpsGap) };
    if (bpsGap < 30)       return { score: 50,  band: "20–30 bps", bpsGap: Math.round(bpsGap) };
    return                        { score: 0,   band: "> 30 bps",  bpsGap: Math.round(bpsGap) };
  };

  // ─── INCENTIVE SLAB (spec §6) ────────────────────────────────────────────────
  const getIncentiveSlab = (achievementPct) => {
    if (achievementPct < 50)  return { label: "No Incentive", rate: 0,      color: "#9B9B9B", bg: "#F5F5F5" };
    if (achievementPct < 80)  return { label: "Low",          rate: 0.0050, color: "#854F0B", bg: "#FAEEDA" };
    if (achievementPct < 100) return { label: "Medium",       rate: 0.0080, color: "#534AB7", bg: "#EEEDFE" };
    if (achievementPct < 125) return { label: "High Jump",    rate: 0.0160, color: "#3B6D11", bg: "#EAF3DE" };
    if (achievementPct < 200) return { label: "Accelerated",  rate: 0.0200, color: "#185FA5", bg: "#E6F1FB" };
    return                           { label: "Maximum",      rate: 0.0250, color: "#0F6E56", bg: "#E1F5EE" };
  };

  // ─── MISSED INCENTIVE (spec §13) ─────────────────────────────────────────────
  const getMissedIncentive = (achievementPct, disbCrs) => {
    const current = getIncentiveSlab(achievementPct);
    // Next meaningful threshold above current achievement
    const thresholds = [50, 80, 100, 125, 200];
    const nextThreshold = thresholds.find(t => t > achievementPct) || 200;
    const next = getIncentiveSlab(nextThreshold);
    const missed = (next.rate - current.rate) * disbCrs; // in Crs
    const additionalDisbNeeded = achievementPct > 0
      ? ((nextThreshold / achievementPct) - 1) * disbCrs
      : 0;
    return { current, next, missed, nextThreshold, additionalDisbNeeded };
  };

  // ─── GEO PERFORMANCE SCORE (spec §11) ────────────────────────────────────────
  // Score = (Disbursal% × 55%) + (Yield% × 15%) + (Conversion% × 10%)
  // Since we don't have target/conversion in raw data, we normalise each metric
  // across all geos (0–100 scale) and apply the weights.
  const calcGeoScore = (disbCrs, yieldPct, allDisbMax, allYieldMax) => {
    const disbScore    = allDisbMax > 0 ? (disbCrs / allDisbMax) * 100 : 0;
    const yieldScore   = allYieldMax > 0 ? (yieldPct / allYieldMax) * 100 : 0;
    return (disbScore * 0.55) + (yieldScore * 0.15);
  };

  // ─── ENRICHED ROW DATA ───────────────────────────────────────────────────────
  // AO(41)=Disb Amt (idx 40), AQ(43)=Disb ROI (idx 42)
  const enrichedData = useMemo(() => {
    if (rawData.length === 0) return [];
    const keys = Object.keys(rawData[0]);
    return rawData.map(row => {
      const disbAmt = parseFloat(fuzzyCol(row, keys, 40, "Disb Amt", "DisbAmt", "Disb_Amt")) || 0;
      const disbROI = parseFloat(fuzzyCol(row, keys, 42, "Disb ROI", "DisbROI", "Disb_ROI")) || 0;
      // Yield (spec §10): Revenue / Disbursal — here ROI IS the yield rate
      const disbCrs = disbAmt / 1e7;
      const revenue = disbAmt * (disbROI / 100); // revenue generated at that ROI
      const wrr     = disbROI * disbCrs;         // weighted rate × volume (for aggregate yield)
      return { ...row, "Disb in Crs": disbCrs, "Revenue": revenue, "WRR": wrr };
    });
  }, [rawData]);

  // ─── FILTERS ─────────────────────────────────────────────────────────────────
  const filterOptions = useMemo(() => {
    const opts = {};
    FILTER_FIELDS.forEach(f => {
      const vals = [...new Set(rawData.map(r => String(r[f] || "")).filter(Boolean))].sort();
      opts[f] = vals;
    });
    return opts;
  }, [rawData]);

  const filteredData = useMemo(() => {
    return enrichedData.filter(row =>
      Object.entries(filters).every(([k, vals]) => {
        if (!vals || vals.length === 0) return true;
        return vals.includes(String(row[k] || ""));
      })
    );
  }, [enrichedData, filters]);

  // ─── YIELD PIVOT: Zone > Geo (spec §10 + §11) ────────────────────────────────
  // Yield = SUM(Revenue) / SUM(Disb Amt)  →  same as WRR/DisbCrs
  const yieldPivot = useMemo(() => {
    const map = {};
    const keys = filteredData.length > 0 ? Object.keys(filteredData[0]) : [];
    filteredData.forEach(row => {
      const zone = String(fuzzyCol(row, keys, null, "Zone") || "Unknown");
      const geo  = String(fuzzyCol(row, keys, null, "Geo")  || "Unknown");
      const key  = `${zone}|||${geo}`;
      if (!map[key]) map[key] = { Zone: zone, Geo: geo, WRR: 0, DisbCrs: 0, Revenue: 0 };
      map[key].WRR     += row["WRR"]          || 0;
      map[key].DisbCrs += row["Disb in Crs"]  || 0;
      map[key].Revenue += row["Revenue"]       || 0;
    });
    // First pass: compute raw yield
    const rows = Object.values(map).map(r => ({
      ...r,
      // Yield (spec §10) = Revenue / Disbursal — expressed as %
      Yield: r.DisbCrs > 0 ? (r.Revenue / (r.DisbCrs * 1e7)) * 100 : 0,
      // WRR-based yield (original formula, kept for reference)
      YieldWRR: r.DisbCrs > 0 ? r.WRR / r.DisbCrs : 0,
    }));
    // Second pass: geo performance score needs normalised max values (spec §11)
    const maxDisb  = Math.max(...rows.map(r => r.DisbCrs), 1);
    const maxYield = Math.max(...rows.map(r => r.Yield), 1);
    return rows.map(r => {
      const yieldScore   = getYieldScore(r.Yield, r.Yield * 1.05); // 5% above current as proxy target
      const geoScore     = calcGeoScore(r.DisbCrs, r.Yield, maxDisb, maxYield);
      const achievePct   = (r.DisbCrs / (maxDisb * 0.8)) * 100; // normalised achievement proxy
      const slab         = getIncentiveSlab(achievePct);
      const missed       = getMissedIncentive(achievePct, r.DisbCrs);
      const incentiveAmt = slab.rate * r.DisbCrs;
      return { ...r, yieldScore, geoScore, achievePct, slab, missed, incentiveAmt };
    }).sort((a, b) => a.Zone.localeCompare(b.Zone) || a.Geo.localeCompare(b.Geo));
  }, [filteredData]);

  const totalDisbCrs  = useMemo(() => filteredData.reduce((s, r) => s + (r["Disb in Crs"] || 0), 0), [filteredData]);
  const totalRevenue  = useMemo(() => filteredData.reduce((s, r) => s + (r["Revenue"]      || 0), 0), [filteredData]);
  const totalWRR      = useMemo(() => filteredData.reduce((s, r) => s + (r["WRR"]          || 0), 0), [filteredData]);
  // Yield = Revenue / Disbursal (spec §10)
  const overallYield  = totalDisbCrs > 0 ? (totalRevenue / (totalDisbCrs * 1e7)) * 100 : 0;

  const pagedData = useMemo(() => {
    const start = (page - 1) * ROWS_PER_PAGE;
    return filteredData.slice(start, start + ROWS_PER_PAGE);
  }, [filteredData, page]);

  const totalPages = Math.ceil(filteredData.length / ROWS_PER_PAGE);

  const toggleFilter = (field, val) => {
    setFilters(prev => {
      const cur = prev[field] || [];
      const next = cur.includes(val) ? cur.filter(v => v !== val) : [...cur, val];
      return { ...prev, [field]: next };
    });
    setPage(1);
  };

  const clearAllFilters = () => { setFilters({}); setPage(1); };

  const activeFilterCount = Object.values(filters).reduce((s, v) => s + (v?.length || 0), 0);

  const displayCols = columns.length > 0
    ? [...columns.filter(c => c !== "Disb in Crs" && c !== "WRR"), "Disb in Crs", "WRR"]
    : [];

  const noData = rawData.length === 0;

  return (
    <div style={{
      fontFamily: "'DM Sans', 'Segoe UI', system-ui, sans-serif",
      minHeight: "100vh",
      background: "var(--color-background-tertiary)",
      color: "var(--color-text-primary)",
      display: "flex",
      flexDirection: "column"
    }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet" />
      <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css" />

      {/* Header */}
      <div style={{
        background: "var(--color-background-primary)",
        borderBottom: "0.5px solid var(--color-border-tertiary)",
        padding: "0 28px",
        height: 56,
        display: "flex", alignItems: "center", justifyContent: "space-between",
        position: "sticky", top: 0, zIndex: 100
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{
            width: 32, height: 32, borderRadius: 8,
            background: "#185FA5",
            display: "flex", alignItems: "center", justifyContent: "center"
          }}>
            <i className="ti ti-chart-line" style={{ fontSize: 17, color: "#fff" }} aria-hidden />
          </div>
          <div>
            <span style={{ fontWeight: 600, fontSize: 15, letterSpacing: "-0.01em" }}>Yield Analytics</span>
            <span style={{ fontSize: 12, color: "var(--color-text-tertiary)", marginLeft: 8 }}>Disbursement & Yield Intelligence</span>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {fileName && <Badge color="teal"><i className="ti ti-file-spreadsheet" style={{ fontSize: 11 }} aria-hidden /> {fileName}</Badge>}
          <button
            onClick={() => fileRef.current?.click()}
            style={{
              display: "flex", alignItems: "center", gap: 6,
              background: "#185FA5", color: "#fff",
              border: "none", borderRadius: 8,
              padding: "7px 14px", fontSize: 13, fontWeight: 500,
              cursor: "pointer"
            }}
          >
            <i className="ti ti-upload" style={{ fontSize: 14 }} aria-hidden />
            {noData ? "Upload File" : "Re-upload"}
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={e => e.target.files[0] && processFile(e.target.files[0])} />
        </div>
      </div>

      <div style={{ display: "flex", flex: 1 }}>
        {/* Sidebar */}
        <aside style={{
          width: noData ? 0 : 252,
          minWidth: noData ? 0 : 252,
          background: "var(--color-background-primary)",
          borderRight: "0.5px solid var(--color-border-tertiary)",
          overflow: "hidden",
          transition: "width 0.3s ease, min-width 0.3s ease",
          display: "flex", flexDirection: "column"
        }}>
          {!noData && (
            <>
              <div style={{
                padding: "14px 16px 10px",
                borderBottom: "0.5px solid var(--color-border-tertiary)",
                display: "flex", alignItems: "center", justifyContent: "space-between"
              }}>
                <span style={{ fontSize: 12, fontWeight: 600, letterSpacing: "0.06em", color: "var(--color-text-secondary)", textTransform: "uppercase" }}>
                  Filters
                </span>
                {activeFilterCount > 0 && (
                  <button onClick={clearAllFilters} style={{
                    fontSize: 11, color: "#185FA5", background: "none", border: "none",
                    cursor: "pointer", fontWeight: 500
                  }}>Clear all ({activeFilterCount})</button>
                )}
              </div>
              <div style={{ overflowY: "auto", flex: 1, padding: "8px 0" }}>
                {FILTER_FIELDS.map(field => {
                  const opts = filterOptions[field] || [];
                  if (opts.length === 0) return null;
                  const sel = filters[field] || [];
                  const search = filterSearch[field] || "";
                  const visible = opts.filter(v => v.toLowerCase().includes(search.toLowerCase()));
                  return (
                    <div key={field} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)", padding: "8px 14px 10px" }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                        <span style={{ fontSize: 11, fontWeight: 600, color: "var(--color-text-secondary)", letterSpacing: "0.03em" }}>{field.toUpperCase()}</span>
                        {sel.length > 0 && <span style={{ fontSize: 10, color: "#185FA5", fontWeight: 500 }}>{sel.length} sel.</span>}
                      </div>
                      {opts.length > 6 && (
                        <input
                          type="text"
                          placeholder="Search…"
                          value={search}
                          onChange={e => setFilterSearch(p => ({ ...p, [field]: e.target.value }))}
                          style={{
                            width: "100%", fontSize: 11, padding: "4px 8px",
                            border: "0.5px solid var(--color-border-secondary)",
                            borderRadius: 5, background: "var(--color-background-secondary)",
                            color: "var(--color-text-primary)", marginBottom: 6,
                            boxSizing: "border-box"
                          }}
                        />
                      )}
                      <div style={{ maxHeight: 140, overflowY: "auto", display: "flex", flexDirection: "column", gap: 2 }}>
                        {visible.slice(0, 30).map(v => (
                          <label key={v} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", padding: "2px 0" }}>
                            <input
                              type="checkbox"
                              checked={sel.includes(v)}
                              onChange={() => toggleFilter(field, v)}
                              style={{ accentColor: "#185FA5", width: 12, height: 12 }}
                            />
                            <span style={{ fontSize: 11, color: "var(--color-text-primary)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={v}>{v || "(blank)"}</span>
                          </label>
                        ))}
                        {visible.length > 30 && (
                          <span style={{ fontSize: 10, color: "var(--color-text-tertiary)", paddingTop: 2 }}>+{visible.length - 30} more. Use search to narrow.</span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </>
          )}
        </aside>

        {/* Main */}
        <main style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
          {noData ? (
            <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", padding: 40 }}>
              <div
                onDragOver={e => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={handleDrop}
                onClick={() => fileRef.current?.click()}
                style={{
                  border: `2px dashed ${dragOver ? "#185FA5" : "var(--color-border-secondary)"}`,
                  borderRadius: 16,
                  padding: "60px 80px",
                  textAlign: "center",
                  cursor: "pointer",
                  background: dragOver ? "#E6F1FB" : "var(--color-background-primary)",
                  transition: "all 0.2s ease",
                  maxWidth: 480
                }}
              >
                <div style={{
                  width: 64, height: 64, borderRadius: 16,
                  background: "#E6F1FB", display: "flex", alignItems: "center",
                  justifyContent: "center", margin: "0 auto 20px"
                }}>
                  <i className="ti ti-file-spreadsheet" style={{ fontSize: 32, color: "#185FA5" }} aria-hidden />
                </div>
                <p style={{ fontSize: 17, fontWeight: 600, margin: "0 0 8px", color: "var(--color-text-primary)" }}>
                  Drop your Excel file here
                </p>
                <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: "0 0 20px", lineHeight: 1.6 }}>
                  Supports .xlsx, .xls, .csv files with loan disbursement data.<br />
                  The tool will automatically calculate Disb in Crs, WRR, and Yield.
                </p>
                <button style={{
                  background: "#185FA5", color: "#fff", border: "none",
                  borderRadius: 8, padding: "9px 20px", fontSize: 13, fontWeight: 500, cursor: "pointer"
                }}>
                  <i className="ti ti-upload" style={{ fontSize: 14, marginRight: 6 }} aria-hidden />
                  Browse file
                </button>
                <p style={{ fontSize: 11, color: "var(--color-text-tertiary)", margin: "16px 0 0" }}>
                  {loading ? "Processing…" : "Or drag & drop from your desktop"}
                </p>
              </div>
            </div>
          ) : (
            <>
              {/* Stats bar */}
              <div style={{
                background: "var(--color-background-primary)",
                borderBottom: "0.5px solid var(--color-border-tertiary)",
                padding: "12px 20px",
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))",
                gap: 10
              }}>
                <MetricCard label="Total Rows" value={rawData.length.toLocaleString()} sub="in uploaded file" />
                <MetricCard label="Filtered Rows" value={filteredData.length.toLocaleString()} sub={activeFilterCount > 0 ? `${activeFilterCount} filter(s) active` : "no filters"} color="#3B6D11" />
                <MetricCard label="Columns" value={displayCols.length} sub={`${columns.length} source + 2 derived`} color="#534AB7" />
                <MetricCard label="Total Disb (Crs)" value={totalDisbCrs.toFixed(2)} sub="sum of filtered rows" color="#0F6E56" />
                <MetricCard label="Avg WRR" value={totalWRR.toFixed(4)} sub="weighted rate of return" color="#854F0B" />
                <MetricCard label="Overall Yield" value={`${overallYield.toFixed(4)}%`} sub="WRR ÷ Disb in Crs" color={overallYield > 10 ? "#3B6D11" : "#185FA5"} />
              </div>

              {/* Tabs */}
              <div style={{
                background: "var(--color-background-primary)",
                borderBottom: "0.5px solid var(--color-border-tertiary)",
                padding: "0 20px",
                display: "flex", gap: 0
              }}>
                {[
                  { id: "preview", label: "Data Preview", icon: "ti-table" },
                  { id: "yield", label: "Yield Summary", icon: "ti-chart-bar" },
                  { id: "performers", label: "Top & Bottom Performers", icon: "ti-arrows-sort" },
                  { id: "incentives", label: "Incentives Engine", icon: "ti-trophy" },
                ].map(t => (
                  <button
                    key={t.id}
                    onClick={() => setTab(t.id)}
                    style={{
                      display: "flex", alignItems: "center", gap: 6,
                      padding: "12px 16px",
                      background: "none", border: "none",
                      borderBottom: tab === t.id ? "2px solid #185FA5" : "2px solid transparent",
                      color: tab === t.id ? "#185FA5" : "var(--color-text-secondary)",
                      fontSize: 13, fontWeight: tab === t.id ? 600 : 400,
                      cursor: "pointer", transition: "all 0.15s"
                    }}
                  >
                    <i className={`ti ${t.icon}`} style={{ fontSize: 15 }} aria-hidden />
                    {t.label}
                    {t.id === "preview" && filteredData.length > 0 && (
                      <span style={{
                        background: "#E6F1FB", color: "#185FA5",
                        borderRadius: 10, fontSize: 10, padding: "1px 6px", fontWeight: 600
                      }}>{filteredData.length.toLocaleString()}</span>
                    )}
                    {t.id === "yield" && yieldPivot.length > 0 && (
                      <span style={{
                        background: "#E1F5EE", color: "#0F6E56",
                        borderRadius: 10, fontSize: 10, padding: "1px 6px", fontWeight: 600
                      }}>{yieldPivot.length} rows</span>
                    )}
                    {t.id === "performers" && yieldPivot.length > 0 && (
                      <span style={{
                        background: "#EEEDFE", color: "#534AB7",
                        borderRadius: 10, fontSize: 10, padding: "1px 6px", fontWeight: 600
                      }}>top 10 · bottom 10</span>
                    )}
                    {t.id === "incentives" && yieldPivot.length > 0 && (
                      <span style={{
                        background: "#FAEEDA", color: "#854F0B",
                        borderRadius: 10, fontSize: 10, padding: "1px 6px", fontWeight: 600
                      }}>{yieldPivot.length} mapped</span>
                    )}
                  </button>
                ))}
              </div>

              {/* Content */}
              <div style={{ flex: 1, overflow: "auto", padding: "0" }}>

                {/* DATA PREVIEW TAB */}
                {tab === "preview" && (
                  <div>
                    <div style={{ overflowX: "auto" }}>
                      <table style={{
                        width: "100%", borderCollapse: "collapse",
                        fontSize: 12, fontFamily: "'DM Mono', monospace"
                      }}>
                        <thead>
                          <tr style={{ background: "var(--color-background-secondary)", position: "sticky", top: 0, zIndex: 10 }}>
                            <th style={thStyle}>#</th>
                            {displayCols.map(c => (
                              <th key={c} style={{
                                ...thStyle,
                                background: (c === "Disb in Crs" || c === "WRR") ? "#E1F5EE" : "var(--color-background-secondary)",
                                color: (c === "Disb in Crs" || c === "WRR") ? "#0F6E56" : "var(--color-text-secondary)",
                                whiteSpace: "nowrap"
                              }}>
                                {c}
                                {(c === "Disb in Crs" || c === "WRR") && (
                                  <span style={{ fontSize: 9, marginLeft: 4, background: "#9FE1CB", color: "#085041", borderRadius: 3, padding: "1px 4px" }}>DERIVED</span>
                                )}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {pagedData.map((row, i) => (
                            <tr key={i} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)" }}
                              onMouseEnter={e => e.currentTarget.style.background = "var(--color-background-secondary)"}
                              onMouseLeave={e => e.currentTarget.style.background = ""}
                            >
                              <td style={{ ...tdStyle, color: "var(--color-text-tertiary)", userSelect: "none" }}>
                                {(page - 1) * ROWS_PER_PAGE + i + 1}
                              </td>
                              {displayCols.map(c => {
                                const val = row[c];
                                const isNum = typeof val === "number";
                                const isDerived = c === "Disb in Crs" || c === "WRR";
                                return (
                                  <td key={c} style={{
                                    ...tdStyle,
                                    textAlign: isNum ? "right" : "left",
                                    color: isDerived ? "#0F6E56" : "var(--color-text-primary)",
                                    fontWeight: isDerived ? 500 : 400,
                                    whiteSpace: "nowrap",
                                    maxWidth: 180,
                                    overflow: "hidden", textOverflow: "ellipsis"
                                  }}>
                                    {isDerived && isNum ? val.toFixed(6) : (val instanceof Date ? val.toLocaleDateString() : String(val ?? ""))}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {/* Pagination */}
                    <div style={{
                      display: "flex", alignItems: "center", gap: 8, justifyContent: "center",
                      padding: "14px 20px",
                      borderTop: "0.5px solid var(--color-border-tertiary)",
                      background: "var(--color-background-primary)"
                    }}>
                      <button onClick={() => setPage(1)} disabled={page === 1} style={pgBtn}>«</button>
                      <button onClick={() => setPage(p => Math.max(1, p - 1))} disabled={page === 1} style={pgBtn}>‹</button>
                      <span style={{ fontSize: 12, color: "var(--color-text-secondary)", padding: "0 8px" }}>
                        Page <strong>{page}</strong> of <strong>{totalPages}</strong>
                        <span style={{ marginLeft: 8, color: "var(--color-text-tertiary)" }}>({filteredData.length.toLocaleString()} rows)</span>
                      </span>
                      <button onClick={() => setPage(p => Math.min(totalPages, p + 1))} disabled={page === totalPages} style={pgBtn}>›</button>
                      <button onClick={() => setPage(totalPages)} disabled={page === totalPages} style={pgBtn}>»</button>
                    </div>
                  </div>
                )}

                {/* YIELD SUMMARY TAB */}
                {tab === "yield" && (
                  <div style={{ padding: 20 }}>
                    {/* Column detection status */}
                    {(() => {
                      if (rawData.length === 0) return null;
                      const firstRow = rawData[0];
                      const norm = s => s.toLowerCase().replace(/[\s_]/g, "");
                      const keys = Object.keys(firstRow);
                      const findCol = (...cands) => keys.find(k => cands.some(c => norm(k) === norm(c)));
                      const disbAmtCol = findCol("Disb Amt", "DisbAmt", "Disb_Amt");
                      const disbROICol = findCol("Disb ROI", "DisbROI", "Disb_ROI");
                      const zoneCol = findCol("Zone");
                      const geoCol = findCol("Geo");
                      const allFound = disbAmtCol && disbROICol && zoneCol && geoCol;
                      return (
                        <div style={{
                          display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 16,
                          padding: "10px 14px",
                          background: allFound ? "#EAF3DE" : "#FAEEDA",
                          border: `0.5px solid ${allFound ? "#C0DD97" : "#FAC775"}`,
                          borderRadius: 8
                        }}>
                          <span style={{ fontSize: 12, fontWeight: 600, color: allFound ? "#3B6D11" : "#854F0B", marginRight: 8 }}>
                            {allFound ? "✓ All required columns detected" : "⚠ Some columns not found — verify column names"}
                          </span>
                          {[
                            { label: "Disb Amt", found: disbAmtCol },
                            { label: "Disb ROI", found: disbROICol },
                            { label: "Zone", found: zoneCol },
                            { label: "Geo", found: geoCol },
                          ].map(({ label, found }) => (
                            <span key={label} style={{
                              fontSize: 11, padding: "2px 8px", borderRadius: 5, fontFamily: "'DM Mono', monospace",
                              background: found ? "#C0DD97" : "#F7C1C1",
                              color: found ? "#085041" : "#791F1F"
                            }}>
                              {label}: {found ? `"${found}"` : "NOT FOUND"}
                            </span>
                          ))}
                        </div>
                      );
                    })()}
                    <div style={{ marginBottom: 16 }}>
                      <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: 0, lineHeight: 1.6 }}>
                        Pivot: <strong>Zone → Geo</strong> &nbsp;|&nbsp; Yield = Revenue ÷ Disbursal (spec §10) &nbsp;|&nbsp; Geo Score = Disb×55% + Yield×15% (spec §11)
                      </p>
                    </div>
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: "'DM Mono', monospace" }}>
                        <thead>
                          <tr style={{ background: "var(--color-background-secondary)" }}>
                            {["Zone","Geo","Disb (Crs)","Revenue (₹)","Yield %","Yield Band","Yield Score","Geo Score","Incentive Slab","Est. Incentive (Crs)"].map(h => (
                              <th key={h} style={thStyle}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {yieldPivot.map((row, i) => (
                            <tr key={i} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)" }}
                              onMouseEnter={e => e.currentTarget.style.background = "var(--color-background-secondary)"}
                              onMouseLeave={e => e.currentTarget.style.background = ""}
                            >
                              <td style={tdStyle}>{row.Zone}</td>
                              <td style={tdStyle}>{row.Geo}</td>
                              <td style={{ ...tdStyle, textAlign: "right", color: "#0F6E56", fontWeight: 500 }}>{row.DisbCrs.toFixed(2)}</td>
                              <td style={{ ...tdStyle, textAlign: "right", color: "#534AB7", fontWeight: 500 }}>
                                {row.Revenue >= 1e7
                                  ? `₹${(row.Revenue/1e7).toFixed(2)}Cr`
                                  : `₹${(row.Revenue/1e5).toFixed(2)}L`}
                              </td>
                              <td style={{ ...tdStyle, textAlign: "right" }}>
                                <span style={{
                                  background: row.Yield >= 13 ? "#EAF3DE" : row.Yield >= 11 ? "#E6F1FB" : row.Yield >= 9 ? "#FAEEDA" : "#FDECEA",
                                  color:      row.Yield >= 13 ? "#3B6D11" : row.Yield >= 11 ? "#185FA5" : row.Yield >= 9 ? "#854F0B" : "#C0392B",
                                  borderRadius: 5, padding: "2px 7px", fontWeight: 700, fontSize: 11
                                }}>{row.Yield.toFixed(4)}%</span>
                              </td>
                              <td style={tdStyle}>
                                <span style={{
                                  fontSize: 10, padding: "2px 6px", borderRadius: 4,
                                  background: row.yieldScore.score >= 125 ? "#EAF3DE" : row.yieldScore.score >= 100 ? "#E6F1FB" : row.yieldScore.score >= 75 ? "#FAEEDA" : "#FDECEA",
                                  color:      row.yieldScore.score >= 125 ? "#3B6D11" : row.yieldScore.score >= 100 ? "#185FA5" : row.yieldScore.score >= 75 ? "#854F0B" : "#C0392B",
                                  fontWeight: 600
                                }}>{row.yieldScore.band}</span>
                              </td>
                              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: row.yieldScore.score >= 100 ? "#3B6D11" : "#C0392B" }}>
                                {row.yieldScore.score}%
                              </td>
                              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#185FA5" }}>
                                {row.geoScore.toFixed(1)}
                              </td>
                              <td style={tdStyle}>
                                <span style={{
                                  fontSize: 10, padding: "2px 6px", borderRadius: 4,
                                  background: row.slab.bg, color: row.slab.color, fontWeight: 600
                                }}>{row.slab.label}</span>
                              </td>
                              <td style={{ ...tdStyle, textAlign: "right", color: "#0F6E56", fontWeight: 500 }}>
                                ₹{row.incentiveAmt.toFixed(4)}
                              </td>
                            </tr>
                          ))}
                          {/* Totals row */}
                          <tr style={{ background: "var(--color-background-secondary)", borderTop: "1px solid var(--color-border-primary)" }}>
                            <td style={{ ...tdStyle, fontWeight: 700 }}>TOTAL</td>
                            <td style={tdStyle}></td>
                            <td style={{ ...tdStyle, textAlign: "right", color: "#0F6E56", fontWeight: 700 }}>{totalDisbCrs.toFixed(2)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", color: "#534AB7", fontWeight: 700 }}>
                              {totalRevenue >= 1e7 ? `₹${(totalRevenue/1e7).toFixed(2)}Cr` : `₹${(totalRevenue/1e5).toFixed(2)}L`}
                            </td>
                            <td style={{ ...tdStyle, textAlign: "right" }}>
                              <span style={{ background: "#185FA5", color: "#fff", borderRadius: 5, padding: "2px 8px", fontWeight: 700, fontSize: 11 }}>
                                {overallYield.toFixed(4)}%
                              </span>
                            </td>
                            <td style={tdStyle}></td>
                            <td style={tdStyle}></td>
                            <td style={tdStyle}></td>
                            <td style={tdStyle}></td>
                            <td style={{ ...tdStyle, textAlign: "right", color: "#0F6E56", fontWeight: 700 }}>
                              ₹{yieldPivot.reduce((s,r) => s + r.incentiveAmt, 0).toFixed(4)}
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* TOP & BOTTOM PERFORMERS TAB */}
                {tab === "performers" && (
                  <div style={{ padding: 24 }}>
                    <div style={{ marginBottom: 20 }}>
                      <h2 style={{ fontSize: 16, fontWeight: 600, margin: "0 0 6px", color: "var(--color-text-primary)" }}>Top & Bottom Performers</h2>
                      <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: 0 }}>
                        Ranked by Yield % from the Yield Summary. Zone → Geo level.
                      </p>
                    </div>

                    {yieldPivot.length === 0 ? (
                      <div style={{
                        textAlign: "center", padding: "48px 24px",
                        border: "0.5px dashed var(--color-border-secondary)",
                        borderRadius: 10, color: "var(--color-text-tertiary)"
                      }}>
                        <i className="ti ti-database-off" style={{ fontSize: 32, display: "block", marginBottom: 12 }} aria-hidden />
                        <p style={{ fontSize: 13, margin: 0 }}>Upload data first to see performer rankings.</p>
                      </div>
                    ) : (
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>

                        {/* TOP 10 */}
                        <div>
                          <div style={{
                            display: "flex", alignItems: "center", gap: 8, marginBottom: 12,
                            padding: "10px 14px",
                            background: "#EAF3DE", border: "0.5px solid #C0DD97", borderRadius: 8
                          }}>
                            <i className="ti ti-trending-up" style={{ fontSize: 16, color: "#3B6D11" }} aria-hidden />
                            <span style={{ fontSize: 13, fontWeight: 600, color: "#3B6D11" }}>Top 10 Best Performers</span>
                            <span style={{ marginLeft: "auto", fontSize: 11, color: "#3B6D11", fontFamily: "'DM Mono', monospace" }}>Highest Yield %</span>
                          </div>
                          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: "'DM Mono', monospace" }}>
                            <thead>
                              <tr style={{ background: "#EAF3DE" }}>
                                <th style={{ ...thStyle, background: "#EAF3DE", color: "#3B6D11" }}>#</th>
                                <th style={{ ...thStyle, background: "#EAF3DE", color: "#3B6D11" }}>Zone</th>
                                <th style={{ ...thStyle, background: "#EAF3DE", color: "#3B6D11" }}>Geo</th>
                                <th style={{ ...thStyle, background: "#EAF3DE", color: "#3B6D11", textAlign: "right" }}>Disb (Crs)</th>
                                <th style={{ ...thStyle, background: "#EAF3DE", color: "#3B6D11", textAlign: "right" }}>Yield %</th>
                              </tr>
                            </thead>
                            <tbody>
                              {[...yieldPivot]
                                .sort((a, b) => b.Yield - a.Yield)
                                .slice(0, 10)
                                .map((row, i) => (
                                  <tr key={i} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)" }}
                                    onMouseEnter={e => e.currentTarget.style.background = "#F3FAE8"}
                                    onMouseLeave={e => e.currentTarget.style.background = ""}
                                  >
                                    <td style={{ ...tdStyle, width: 28 }}>
                                      <span style={{
                                        display: "inline-flex", alignItems: "center", justifyContent: "center",
                                        width: 20, height: 20, borderRadius: "50%",
                                        background: i === 0 ? "#3B6D11" : i === 1 ? "#5A9E1C" : i === 2 ? "#7DC142" : "#EAF3DE",
                                        color: i < 3 ? "#fff" : "#3B6D11",
                                        fontSize: 10, fontWeight: 700
                                      }}>{i + 1}</span>
                                    </td>
                                    <td style={tdStyle}>{row.Zone}</td>
                                    <td style={tdStyle}>{row.Geo}</td>
                                    <td style={{ ...tdStyle, textAlign: "right", color: "#0F6E56" }}>{row.DisbCrs.toFixed(2)}</td>
                                    <td style={{ ...tdStyle, textAlign: "right" }}>
                                      <span style={{
                                        background: "#EAF3DE", color: "#3B6D11",
                                        borderRadius: 5, padding: "2px 7px", fontWeight: 700, fontSize: 11
                                      }}>{row.Yield.toFixed(4)}%</span>
                                    </td>
                                  </tr>
                                ))}
                            </tbody>
                          </table>
                        </div>

                        {/* BOTTOM 10 */}
                        <div>
                          <div style={{
                            display: "flex", alignItems: "center", gap: 8, marginBottom: 12,
                            padding: "10px 14px",
                            background: "#FDECEA", border: "0.5px solid #F5B7B1", borderRadius: 8
                          }}>
                            <i className="ti ti-trending-down" style={{ fontSize: 16, color: "#C0392B" }} aria-hidden />
                            <span style={{ fontSize: 13, fontWeight: 600, color: "#C0392B" }}>Bottom 10 Low Performers</span>
                            <span style={{ marginLeft: "auto", fontSize: 11, color: "#C0392B", fontFamily: "'DM Mono', monospace" }}>Lowest Yield %</span>
                          </div>
                          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: "'DM Mono', monospace" }}>
                            <thead>
                              <tr style={{ background: "#FDECEA" }}>
                                <th style={{ ...thStyle, background: "#FDECEA", color: "#C0392B" }}>#</th>
                                <th style={{ ...thStyle, background: "#FDECEA", color: "#C0392B" }}>Zone</th>
                                <th style={{ ...thStyle, background: "#FDECEA", color: "#C0392B" }}>Geo</th>
                                <th style={{ ...thStyle, background: "#FDECEA", color: "#C0392B", textAlign: "right" }}>Disb (Crs)</th>
                                <th style={{ ...thStyle, background: "#FDECEA", color: "#C0392B", textAlign: "right" }}>Yield %</th>
                              </tr>
                            </thead>
                            <tbody>
                              {[...yieldPivot]
                                .sort((a, b) => a.Yield - b.Yield)
                                .slice(0, 10)
                                .map((row, i) => (
                                  <tr key={i} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)" }}
                                    onMouseEnter={e => e.currentTarget.style.background = "#FEF5F4"}
                                    onMouseLeave={e => e.currentTarget.style.background = ""}
                                  >
                                    <td style={{ ...tdStyle, width: 28 }}>
                                      <span style={{
                                        display: "inline-flex", alignItems: "center", justifyContent: "center",
                                        width: 20, height: 20, borderRadius: "50%",
                                        background: i === 0 ? "#C0392B" : i === 1 ? "#E74C3C" : i === 2 ? "#EC7063" : "#FDECEA",
                                        color: i < 3 ? "#fff" : "#C0392B",
                                        fontSize: 10, fontWeight: 700
                                      }}>{i + 1}</span>
                                    </td>
                                    <td style={tdStyle}>{row.Zone}</td>
                                    <td style={tdStyle}>{row.Geo}</td>
                                    <td style={{ ...tdStyle, textAlign: "right", color: "#C0392B" }}>{row.DisbCrs.toFixed(2)}</td>
                                    <td style={{ ...tdStyle, textAlign: "right" }}>
                                      <span style={{
                                        background: "#FDECEA", color: "#C0392B",
                                        borderRadius: 5, padding: "2px 7px", fontWeight: 700, fontSize: 11
                                      }}>{row.Yield.toFixed(4)}%</span>
                                    </td>
                                  </tr>
                                ))}
                            </tbody>
                          </table>
                        </div>

                      </div>
                    )}
                  </div>
                )}

                {/* INCENTIVES ENGINE TAB */}
                {tab === "incentives" && (
                  <div style={{ padding: 24 }}>
                    <div style={{ marginBottom: 20 }}>
                      <h2 style={{ fontSize: 16, fontWeight: 600, margin: "0 0 6px", color: "var(--color-text-primary)" }}>Incentives Engine</h2>
                      <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: 0, lineHeight: 1.6 }}>
                        Slab-based incentive calculation · Missed incentive analysis · Auto-generated insights (spec §6, §13, §14)
                      </p>
                    </div>

                    {/* Slab reference cards */}
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 10, marginBottom: 24 }}>
                      {[
                        { label: "No Incentive", range: "< 50%",      rate: "0%",     color: "#9B9B9B", bg: "#F5F5F5", border: "#E0E0E0" },
                        { label: "Low",          range: "50–79%",     rate: "0.50%",  color: "#854F0B", bg: "#FAEEDA", border: "#FAC775" },
                        { label: "Medium",       range: "80–99%",     rate: "0.80%",  color: "#534AB7", bg: "#EEEDFE", border: "#CECBF6" },
                        { label: "High Jump",    range: "100–124%",   rate: "1.60%",  color: "#3B6D11", bg: "#EAF3DE", border: "#C0DD97" },
                        { label: "Accelerated",  range: "125–199%",   rate: "2.00%",  color: "#185FA5", bg: "#E6F1FB", border: "#B5D4F4" },
                        { label: "Maximum",      range: "200%+",      rate: "2.50%",  color: "#0F6E56", bg: "#E1F5EE", border: "#9FE1CB" },
                      ].map(s => (
                        <div key={s.label} style={{ background: s.bg, border: `0.5px solid ${s.border}`, borderRadius: 9, padding: "12px 14px" }}>
                          <p style={{ fontSize: 11, fontWeight: 700, color: s.color, margin: "0 0 4px" }}>{s.label}</p>
                          <p style={{ fontSize: 12, color: s.color, margin: "0 0 6px", fontFamily: "'DM Mono', monospace" }}>{s.range} achievement</p>
                          <p style={{ fontSize: 16, fontWeight: 700, color: s.color, margin: 0 }}>{s.rate} <span style={{ fontSize: 10, fontWeight: 400 }}>of Disb</span></p>
                        </div>
                      ))}
                    </div>

                    {yieldPivot.length === 0 ? (
                      <div style={{ textAlign: "center", padding: "48px 24px", border: "0.5px dashed var(--color-border-secondary)", borderRadius: 10, color: "var(--color-text-tertiary)" }}>
                        <i className="ti ti-database-off" style={{ fontSize: 32, display: "block", marginBottom: 12 }} aria-hidden />
                        <p style={{ fontSize: 13, margin: 0 }}>Upload data to populate incentive analysis.</p>
                      </div>
                    ) : (
                      <>
                        {/* Auto-generated Insights (spec §14) */}
                        <h3 style={{ fontSize: 12, fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase", color: "var(--color-text-secondary)", margin: "0 0 10px" }}>
                          Auto-Generated Insights
                        </h3>
                        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 24 }}>
                          {(() => {
                            const insights = [];
                            const totalMissed = yieldPivot.reduce((s,r) => s + r.missed.missed, 0);
                            const lowestGeo = [...yieldPivot].sort((a,b) => a.geoScore - b.geoScore)[0];
                            const highestMissed = [...yieldPivot].sort((a,b) => b.missed.missed - a.missed.missed)[0];
                            const belowSlab = yieldPivot.filter(r => r.achievePct < 100).length;

                            insights.push({
                              icon: "ti-bulb", color: "#854F0B", bg: "#FAEEDA", border: "#FAC775",
                              text: `Total missed incentive across all geos: ₹${totalMissed.toFixed(4)} Crs — primarily due to sub-100% achievement slabs.`
                            });
                            if (highestMissed) insights.push({
                              icon: "ti-alert-triangle", color: "#C0392B", bg: "#FDECEA", border: "#F5B7B1",
                              text: `Highest incentive leakage: ${highestMissed.Geo} (${highestMissed.Zone}) — missed ₹${highestMissed.missed.missed.toFixed(4)} Crs. Needs +${highestMissed.missed.additionalDisbNeeded.toFixed(2)} Crs disb to reach ${highestMissed.missed.nextThreshold}% threshold.`
                            });
                            if (lowestGeo) insights.push({
                              icon: "ti-map-pin-off", color: "#534AB7", bg: "#EEEDFE", border: "#CECBF6",
                              text: `Lowest Geo Score: ${lowestGeo.Geo} (${lowestGeo.Zone}) — score ${lowestGeo.geoScore.toFixed(1)}. Focus on disbursal volume to improve weighted performance score.`
                            });
                            insights.push({
                              icon: "ti-info-circle", color: "#185FA5", bg: "#E6F1FB", border: "#B5D4F4",
                              text: `${belowSlab} of ${yieldPivot.length} geos are below 100% achievement. MSME disbursals in these geos provide NO incentive boost until HL hits 100% (spec §5).`
                            });
                            insights.push({
                              icon: "ti-chart-arrows-vertical", color: "#0F6E56", bg: "#E1F5EE", border: "#9FE1CB",
                              text: `Critical thresholds: 80% and 100% achievement trigger the biggest slab jumps (0.50% → 0.80% → 1.60%). Most incentive gain comes from crossing 100%.`
                            });
                            return insights.map((ins, i) => (
                              <div key={i} style={{ display: "flex", gap: 10, padding: "10px 14px", background: ins.bg, border: `0.5px solid ${ins.border}`, borderRadius: 8 }}>
                                <i className={`ti ${ins.icon}`} style={{ fontSize: 16, color: ins.color, marginTop: 1, flexShrink: 0 }} aria-hidden />
                                <span style={{ fontSize: 12, color: ins.color, lineHeight: 1.6 }}>{ins.text}</span>
                              </div>
                            ));
                          })()}
                        </div>

                        {/* Missed Incentive Table (spec §13) */}
                        <h3 style={{ fontSize: 12, fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase", color: "var(--color-text-secondary)", margin: "0 0 10px" }}>
                          Missed Incentive Analysis — by Geo
                        </h3>
                        <div style={{ overflowX: "auto" }}>
                          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: "'DM Mono', monospace" }}>
                            <thead>
                              <tr style={{ background: "var(--color-background-secondary)" }}>
                                {["Zone","Geo","Disb (Crs)","Yield %","Current Slab","Current Rate","Next Threshold","Next Slab","Next Rate","Missed Incentive (Crs)","Extra Disb Needed"].map(h => (
                                  <th key={h} style={thStyle}>{h}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {[...yieldPivot].sort((a,b) => b.missed.missed - a.missed.missed).map((row, i) => (
                                <tr key={i} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)" }}
                                  onMouseEnter={e => e.currentTarget.style.background = "var(--color-background-secondary)"}
                                  onMouseLeave={e => e.currentTarget.style.background = ""}
                                >
                                  <td style={tdStyle}>{row.Zone}</td>
                                  <td style={tdStyle}>{row.Geo}</td>
                                  <td style={{ ...tdStyle, textAlign: "right", color: "#0F6E56" }}>{row.DisbCrs.toFixed(2)}</td>
                                  <td style={{ ...tdStyle, textAlign: "right" }}>
                                    <span style={{ background: row.Yield >= 13 ? "#EAF3DE" : row.Yield >= 11 ? "#E6F1FB" : "#FAEEDA", color: row.Yield >= 13 ? "#3B6D11" : row.Yield >= 11 ? "#185FA5" : "#854F0B", borderRadius: 4, padding: "1px 6px", fontWeight: 600, fontSize: 11 }}>
                                      {row.Yield.toFixed(4)}%
                                    </span>
                                  </td>
                                  <td style={tdStyle}>
                                    <span style={{ background: row.missed.current.bg, color: row.missed.current.color, borderRadius: 4, padding: "1px 6px", fontWeight: 600, fontSize: 11 }}>
                                      {row.missed.current.label}
                                    </span>
                                  </td>
                                  <td style={{ ...tdStyle, textAlign: "right", color: row.missed.current.color, fontWeight: 500 }}>{(row.missed.current.rate * 100).toFixed(2)}%</td>
                                  <td style={{ ...tdStyle, textAlign: "right", color: "#534AB7", fontWeight: 500 }}>{row.missed.nextThreshold}%</td>
                                  <td style={tdStyle}>
                                    <span style={{ background: row.missed.next.bg, color: row.missed.next.color, borderRadius: 4, padding: "1px 6px", fontWeight: 600, fontSize: 11 }}>
                                      {row.missed.next.label}
                                    </span>
                                  </td>
                                  <td style={{ ...tdStyle, textAlign: "right", color: row.missed.next.color, fontWeight: 500 }}>{(row.missed.next.rate * 100).toFixed(2)}%</td>
                                  <td style={{ ...tdStyle, textAlign: "right" }}>
                                    <span style={{ background: row.missed.missed > 0.001 ? "#FDECEA" : "#EAF3DE", color: row.missed.missed > 0.001 ? "#C0392B" : "#3B6D11", borderRadius: 4, padding: "2px 8px", fontWeight: 700, fontSize: 11 }}>
                                      ₹{row.missed.missed.toFixed(4)}
                                    </span>
                                  </td>
                                  <td style={{ ...tdStyle, textAlign: "right", color: "#854F0B", fontWeight: 500 }}>
                                    +{row.missed.additionalDisbNeeded.toFixed(2)} Crs
                                  </td>
                                </tr>
                              ))}
                              <tr style={{ background: "var(--color-background-secondary)", borderTop: "1px solid var(--color-border-primary)" }}>
                                <td style={{ ...tdStyle, fontWeight: 700 }}>TOTAL</td>
                                <td style={tdStyle}></td>
                                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0F6E56" }}>{totalDisbCrs.toFixed(2)}</td>
                                <td style={tdStyle}></td><td style={tdStyle}></td><td style={tdStyle}></td>
                                <td style={tdStyle}></td><td style={tdStyle}></td><td style={tdStyle}></td>
                                <td style={{ ...tdStyle, textAlign: "right" }}>
                                  <span style={{ background: "#C0392B", color: "#fff", borderRadius: 4, padding: "2px 8px", fontWeight: 700, fontSize: 11 }}>
                                    ₹{yieldPivot.reduce((s,r) => s + r.missed.missed, 0).toFixed(4)}
                                  </span>
                                </td>
                                <td style={tdStyle}></td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </>
                    )}
                  </div>
                )}
              </div>
            </>
          )}
        </main>
      </div>
    </div>
  );
}

const thStyle = {
  padding: "9px 12px",
  textAlign: "left",
  fontSize: 11,
  fontWeight: 600,
  color: "var(--color-text-secondary)",
  borderBottom: "0.5px solid var(--color-border-tertiary)",
  letterSpacing: "0.04em",
  position: "sticky",
  top: 0,
  zIndex: 5,
  background: "var(--color-background-secondary)"
};

const tdStyle = {
  padding: "7px 12px",
  fontSize: 12,
  color: "var(--color-text-primary)",
  borderBottom: "0.5px solid var(--color-border-tertiary)",
  fontFamily: "'DM Mono', monospace"
};

const pgBtn = {
  background: "var(--color-background-secondary)",
  border: "0.5px solid var(--color-border-secondary)",
  borderRadius: 6,
  padding: "5px 10px",
  fontSize: 13,
  cursor: "pointer",
  color: "var(--color-text-primary)"
};

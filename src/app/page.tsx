"use client";
import { useState, useEffect } from "react";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  Legend,
  CartesianGrid,
  AreaChart,
  Area,
  PieChart,
  Pie,
  Cell,
} from "recharts";
import * as XLSX from "xlsx";

/* ── SweetAlert2 loader ── */
const loadSwal = () =>
  new Promise((resolve) => {
    if (window.Swal) {
      resolve(window.Swal);
      return;
    }
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/sweetalert2@11";
    s.onload = () => resolve(window.Swal);
    document.head.appendChild(s);
  });

const swal = {
  fire: async (opts) => {
    const S = await loadSwal();
    return S.fire(opts);
  },
  confirm: async (opts) => {
    const S = await loadSwal();
    return S.fire({
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#b8924a",
      cancelButtonColor: "#8a8680",
      ...opts,
    });
  },
};

/* ─────────────── GLOBAL STYLES ─────────────── */
const css = `
  @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@300;400;500;600;700&family=Sarabun:wght@300;400;500;600&display=swap');
  @import url('https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --ink:       #0f0e0c;
    --ink-2:     #3a3832;
    --ink-3:     #8a8680;
    --stone:     #f5f3ef;
    --stone-2:   #edeae4;
    --stone-3:   #e0ddd6;
    --gold:      #b8924a;
    --gold-lt:   #e8d5b0;
    --gold-dim:  #f7f0e4;
    --white:     #faf9f7;
    --green:     #3d6b4f;
    --green-lt:  #e8f0eb;
    --red:       #8b3a3a;
    --red-lt:    #f5e8e8;
    --radius-sm: 4px;
    --radius-md: 8px;
    --radius-lg: 14px;
    --shadow-sm: 0 1px 4px rgba(15,14,12,0.06);
    --shadow-md: 0 4px 20px rgba(15,14,12,0.08);
    --shadow-lg: 0 8px 40px rgba(15,14,12,0.12);
    --font-display: 'Cormorant Garamond', 'Sarabun', serif;
    --font-body:    'Sarabun', sans-serif;
  }

  html, body { background: var(--stone); color: var(--ink); font-family: var(--font-body); font-size: 14px; line-height: 1.6; }
  button, select, input, label { font-family: var(--font-body); }

  .num { font-variant-numeric: tabular-nums; letter-spacing: -0.01em; font-family: var(--font-body); }

  /* ── Nav ── */
  .nav {
    background: var(--white);
    border-bottom: 1px solid var(--stone-3);
    height: 62px;
    padding: 0 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 100;
  }

  /* ── Gold divider ── */
  .gold-line { height: 1px; background: linear-gradient(90deg, transparent, var(--gold), transparent); margin: 0; }

  /* ── Cards ── */
  .card {
    background: var(--white);
    border: 1px solid var(--stone-3);
    border-radius: var(--radius-lg);
    transition: box-shadow .25s, transform .2s;
  }
  .card:hover { box-shadow: var(--shadow-md); transform: translateY(-1px); }

  .kpi-card {
    background: var(--white);
    border: 1px solid var(--stone-3);
    border-radius: var(--radius-lg);
    padding: 20px 22px 18px;
    position: relative;
    overflow: hidden;
    transition: box-shadow .25s, transform .2s;
  }
  .kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, var(--gold) 0%, var(--gold-lt) 100%);
  }
  .kpi-card:hover { box-shadow: var(--shadow-md); transform: translateY(-1px); }

  /* ── Table ── */
  .row-hover:hover td { background: var(--stone) !important; }

  /* ── Tabs ── */
  .tab-btn {
    background: none;
    border: none;
    cursor: pointer;
    padding: 16px 22px;
    font-size: 13px;
    font-family: var(--font-body);
    font-weight: 500;
    color: var(--ink-3);
    border-bottom: 2px solid transparent;
    margin-bottom: -1px;
    transition: color .15s, border-color .15s;
    white-space: nowrap;
    flex-shrink: 0;
    letter-spacing: .03em;
  }
  .tab-btn.active { color: var(--gold); border-bottom-color: var(--gold); font-weight: 600; }
  .tab-btn:hover:not(.active) { color: var(--ink-2); }

  /* ── Selects ── */
  .sel {
    background: var(--stone);
    border: 1px solid var(--stone-3);
    border-radius: var(--radius-sm);
    color: var(--ink);
    padding: 8px 12px;
    font-size: 12.5px;
    font-family: var(--font-body);
    cursor: pointer;
    outline: none;
    transition: border-color .15s;
    letter-spacing: .02em;
  }
  .sel:focus { border-color: var(--gold); }

  /* ── Buttons ── */
  .btn-primary {
    background: var(--ink);
    color: var(--white);
    border: none;
    border-radius: var(--radius-sm);
    padding: 9px 18px;
    font-size: 12.5px;
    font-family: var(--font-body);
    font-weight: 600;
    cursor: pointer;
    letter-spacing: .06em;
    text-transform: uppercase;
    transition: background .15s;
    white-space: nowrap;
    display: flex;
    align-items: center;
    gap: 7px;
  }
  .btn-primary:hover { background: var(--gold); }

  .btn-ghost {
    background: transparent;
    color: var(--red);
    border: 1px solid var(--red-lt);
    border-radius: var(--radius-sm);
    padding: 8px 14px;
    font-size: 12px;
    font-family: var(--font-body);
    font-weight: 500;
    cursor: pointer;
    letter-spacing: .04em;
    transition: all .15s;
    white-space: nowrap;
  }
  .btn-ghost:hover { background: var(--red-lt); border-color: var(--red); }

  /* ── Badge ── */
  .badge {
    display: inline-flex;
    align-items: center;
    gap: 3px;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .04em;
  }
  .badge.pos { background: var(--green-lt); color: var(--green); }
  .badge.neg { background: var(--red-lt); color: var(--red); }

  /* ── Mobile ── */
  .hamburger { display: none; background: none; border: 1px solid var(--stone-3); border-radius: var(--radius-sm); padding: 7px 11px; cursor: pointer; font-size: 16px; color: var(--ink-2); }
  .desktop-controls { display: flex; align-items: center; gap: 10px; }

  .mobile-panel { display: none; background: var(--white); border-bottom: 1px solid var(--stone-3); padding: 16px 20px; flex-direction: column; gap: 12px; }

  @media (max-width: 768px) {
    .hamburger { display: flex; }
    .desktop-controls { display: none !important; }
    .mobile-panel { display: flex; }
    .kpi-grid { grid-template-columns: repeat(2,1fr) !important; }
  }
  @media (max-width: 480px) {
    .kpi-grid { grid-template-columns: repeat(1,1fr) !important; }
  }

  .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  .data-table { min-width: 560px; }

  .tab-bar {
    display: flex;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
    scrollbar-width: none;
    border-bottom: 1px solid var(--stone-3);
  }
  .tab-bar::-webkit-scrollbar { display: none; }

  /* ── Section label ── */
  .section-label {
    font-size: 10px;
    font-weight: 600;
    letter-spacing: .12em;
    text-transform: uppercase;
    color: var(--ink-3);
    margin-bottom: 4px;
  }

  .chart-layout { display: flex; gap: 28px; flex-wrap: wrap; }
  .chart-main { flex: 2; min-width: 260px; }
  .chart-side { flex: 1; min-width: 200px; }

  /* ── Product quantity pills ── */
  .qty-pills { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 6px; }
  .qty-pill {
    display: flex;
    align-items: center;
    gap: 5px;
    background: var(--stone);
    border: 1px solid var(--stone-3);
    border-radius: 20px;
    padding: 3px 10px;
    font-size: 11px;
    color: var(--ink-2);
  }
  .qty-pill .dot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
  .qty-pill .qty-num { font-weight: 700; color: var(--ink); }

  /* ── Trend cards ── */
  .trend-cards { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 20px; }
  .trend-card { flex: 1; min-width: 110px; }

  .main-pad { padding: 28px 24px; }
  @media (max-width: 600px) { .main-pad { padding: 18px 14px; } }

  /* ── Ornament ── */
  .ornament { color: var(--gold); font-size: 11px; letter-spacing: .3em; opacity: .6; }

  /* ── Scrollbar ── */
  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--stone-3); border-radius: 10px; }
`;

const VALID_BRANCHES = [
  "Central Festival",
  "Central Airport",
  "Chotana",
  "Ruamchok Mall",
  "Thaweechok",
];
const BRANCH_COLORS = ["#b8924a", "#3d6b4f", "#8b3a3a", "#5b6b8b", "#7a5b8b"];
const PROD_COLORS = ["#b8924a", "#3d6b4f", "#8b5b3a"];
const PROD_NAMES = ["ถั่วงอก", "ข้าวหมาก", "ต้นอ่อน"];

/* ─── KPI Card ─────────────────────────────── */
function KpiCard({ label, value, sub, icon, isLarge = false }) {
  return (
    <div className="kpi-card">
      <div className="section-label">{label}</div>
      <div
        style={{
          display: "flex",
          alignItems: "baseline",
          gap: 8,
          marginTop: 4,
          marginBottom: 6,
        }}
      >
        <div
          className="num"
          style={{
            fontSize: isLarge ? 26 : 22,
            fontWeight: 700,
            color: "var(--ink)",
            letterSpacing: "-0.02em",
          }}
        >
          {value}
        </div>
      </div>
      {sub && (
        <div
          style={{
            fontSize: 11,
            color: "var(--gold)",
            fontWeight: 500,
            letterSpacing: ".04em",
          }}
        >
          {icon} {sub}
        </div>
      )}
    </div>
  );
}

/* ─── Tooltip ───────────────────────────────── */
function CustomTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div
      style={{
        background: "var(--white)",
        border: "1px solid var(--stone-3)",
        borderRadius: 8,
        padding: "10px 14px",
        fontSize: 12,
        boxShadow: "var(--shadow-lg)",
        minWidth: 160,
      }}
    >
      <div
        style={{
          fontWeight: 700,
          color: "var(--ink-2)",
          marginBottom: 8,
          letterSpacing: ".03em",
          fontSize: 11,
          textTransform: "uppercase",
        }}
      >
        {label}
      </div>
      {payload.map((p, i) => (
        <div
          key={i}
          style={{
            display: "flex",
            justifyContent: "space-between",
            gap: 16,
            marginBottom: 3,
          }}
        >
          <span style={{ color: p.color, fontWeight: 500, fontSize: 12 }}>
            ● {p.name}
          </span>
          <span
            className="num"
            style={{ fontWeight: 700, color: "var(--ink)" }}
          >
            ฿{Number(p.value).toLocaleString()}
          </span>
        </div>
      ))}
    </div>
  );
}

/* ─── Qty Tooltip (for product chart - shows ชิ้น not ฿) ── */
function QtyTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div
      style={{
        background: "var(--white)",
        border: "1px solid var(--stone-3)",
        borderRadius: 8,
        padding: "10px 14px",
        fontSize: 12,
        boxShadow: "var(--shadow-lg)",
        minWidth: 160,
      }}
    >
      <div
        style={{
          fontWeight: 700,
          color: "var(--ink-2)",
          marginBottom: 8,
          letterSpacing: ".03em",
          fontSize: 11,
          textTransform: "uppercase",
        }}
      >
        {label}
      </div>
      {payload.map((p, i) => (
        <div
          key={i}
          style={{
            display: "flex",
            justifyContent: "space-between",
            gap: 16,
            marginBottom: 3,
          }}
        >
          <span style={{ color: p.color, fontWeight: 500, fontSize: 12 }}>
            ● {p.name}
          </span>
          <span
            className="num"
            style={{ fontWeight: 700, color: "var(--ink)" }}
          >
            {Number(p.value).toLocaleString()} ชิ้น
          </span>
        </div>
      ))}
    </div>
  );
}

/* ─── Main ──────────────────────────────────── */
export default function Home() {
  const [allServerData, setAllServerData] = useState([]);
  const [availableDates, setAvailableDates] = useState([]);
  const [comparisonData, setComparisonData] = useState([]);
  const [dateA, setDateA] = useState("");
  const [dateB, setDateB] = useState("");
  const [selectedBranch, setSelectedBranch] = useState("ALL");
  const [isUploading, setIsUploading] = useState(false);
  const [activeTab, setActiveTab] = useState("compare");
  const [menuOpen, setMenuOpen] = useState(false);

  // Inject styles client-side only to avoid SSR hydration mismatch
  useEffect(() => {
    const el = document.createElement("style");
    el.setAttribute("data-dashboard", "1");
    el.textContent = css;
    document.head.appendChild(el);
    return () => {
      el.remove();
    };
  }, []);

  const fetchRecords = async () => {
    try {
      const res = await fetch("/api/insert");
      const data = await res.json();
      if (!Array.isArray(data)) return;
      setAllServerData(data);
      const dates = Array.from(
        new Set(data.filter((i) => i.sales_date).map((i) => i.sales_date)),
      );
      setAvailableDates(dates);
      if (!dates.includes(dateA)) setDateA(dates[0] || "");
      if (!dates.includes(dateB)) setDateB(dates[1] || "");
    } catch (e) {
      console.error(e);
    }
  };

  useEffect(() => {
    fetchRecords();
  }, []);

  useEffect(() => {
    if (!dateA) return;
    const merged = VALID_BRANCHES.map((branch) => {
      const a =
        allServerData.find(
          (d) => d.branch === branch && d.sales_date === dateA,
        ) || {};
      const b =
        allServerData.find(
          (d) => d.branch === branch && d.sales_date === dateB,
        ) || {};
      return {
        branch,
        short: branch.split(" ")[0],
        a_total: a.total_bath || 0,
        b_total: b.total_bath || 0,
        // tuangok/khaomak/tonon from Excel col3-5 = qty (units), not revenue
        a_tuangok: a.tuangok || 0,
        a_khaomak: a.khaomak || 0,
        a_tonon: a.tonon || 0,
        // qty: prefer explicit _qty field, fallback to tuangok/khaomak/tonon (same value from this Excel)
        a_tuangok_qty: a.tuangok_qty ?? a.tuangok ?? 0,
        a_khaomak_qty: a.khaomak_qty ?? a.khaomak ?? 0,
        a_tonon_qty: a.tonon_qty ?? a.tonon ?? 0,
      };
    });
    setComparisonData(merged);
  }, [dateA, dateB, allServerData]);

  const handleDelete = async (date) => {
    const result = await swal.confirm({
      title: `ลบข้อมูลวันที่ ${date}?`,
      text: "การดำเนินการนี้ไม่สามารถย้อนกลับได้",
      confirmButtonText: "ลบข้อมูล",
      cancelButtonText: "ยกเลิก",
    });
    if (!result.isConfirmed) return;
    try {
      const res = await fetch("/api/insert", {
        method: "DELETE",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sales_date: date }),
      });
      const r = await res.json();
      if (r.success) {
        await swal.fire({
          icon: "success",
          title: "ลบเรียบร้อย",
          text: `ข้อมูลวันที่ ${date} ถูกลบแล้ว`,
          confirmButtonColor: "#b8924a",
          timer: 2000,
          timerProgressBar: true,
        });
        await fetchRecords();
        window.location.reload();
      } else {
        await swal.fire({
          icon: "error",
          title: "ลบไม่สำเร็จ",
          text: r.error,
          confirmButtonColor: "#b8924a",
        });
      }
    } catch {
      await swal.fire({
        icon: "error",
        title: "เชื่อมต่อไม่ได้",
        text: "กรุณาตรวจสอบการเชื่อมต่อ",
        confirmButtonColor: "#b8924a",
      });
    }
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsUploading(true);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      const wb = XLSX.read(evt.target?.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const rows = [];
      raw.forEach((row) => {
        const b = row[1] ? String(row[1]).trim() : "";
        if (b && VALID_BRANCHES.includes(b) && row[2]) {
          // Excel layout: col1=ลำดับ, col1=สาขา, col2=วันที่,
          // col3=ถั่วงอก(qty), col4=ข้าวหมาก(qty), col5=ต้นอ่อน(qty), col6=Total Bath
          // col3-5 are QUANTITY (units), not revenue — save as both tuangok/khaomak/tonon AND _qty
          const tg = Number(row[3]) || 0;
          const km = Number(row[4]) || 0;
          const tn = Number(row[5]) || 0;
          rows.push({
            branch: b,
            sales_date: XLSX.SSF.format("yyyy-mm-dd", row[2]),
            tuangok: tg, // qty used for product chart
            khaomak: km,
            tonon: tn,
            total_bath: Number(row[6]) || 0,
            tuangok_qty: tg, // explicit qty fields for display
            khaomak_qty: km,
            tonon_qty: tn,
          });
        }
      });
      if (!rows.length) {
        await swal.fire({
          icon: "warning",
          title: "ไม่พบข้อมูลสาขา",
          text: "กรุณาตรวจสอบรูปแบบไฟล์ Excel ให้ถูกต้อง",
          confirmButtonColor: "#b8924a",
        });
        setIsUploading(false);
        return;
      }
      try {
        const res = await fetch("/api/insert", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(rows),
        });
        const r = await res.json();
        if (r.success) {
          await swal.fire({
            icon: "success",
            title: "บันทึกสำเร็จ",
            text: `นำเข้าข้อมูล ${rows.length} แถวเรียบร้อยแล้ว`,
            confirmButtonColor: "#b8924a",
            timer: 2000,
            timerProgressBar: true,
          });
          await fetchRecords();
          window.location.reload();
        } else {
          await swal.fire({
            icon: "error",
            title: "เกิดข้อผิดพลาด",
            text: r.error,
            confirmButtonColor: "#b8924a",
          });
        }
      } catch {
        await swal.fire({
          icon: "error",
          title: "ไม่สามารถเชื่อมต่อได้",
          text: "กรุณาตรวจสอบเซิร์ฟเวอร์",
          confirmButtonColor: "#b8924a",
        });
      } finally {
        setIsUploading(false);
      }
    };
    reader.readAsBinaryString(file);
  };

  /* ── Derived data ── */
  const filtered =
    selectedBranch === "ALL"
      ? comparisonData
      : comparisonData.filter((d) => d.branch === selectedBranch);
  const totalA = filtered.reduce((s, d) => s + d.a_total, 0);
  const totalB = filtered.reduce((s, d) => s + d.b_total, 0);
  const diff = totalA - totalB;
  const pct = totalB > 0 ? ((diff / totalB) * 100).toFixed(1) : null;
  const top = [...comparisonData].sort((a, b) => b.a_total - a.a_total)[0];

  // Product totals (revenue + qty)
  const prodData = [
    {
      name: "ถั่วงอก",
      qty: filtered.reduce((s, d) => s + d.a_tuangok_qty, 0),
      color: PROD_COLORS[0],
    },
    {
      name: "ข้าวหมาก",
      qty: filtered.reduce((s, d) => s + d.a_khaomak_qty, 0),
      color: PROD_COLORS[1],
    },
    {
      name: "ต้นอ่อน",
      qty: filtered.reduce((s, d) => s + d.a_tonon_qty, 0),
      color: PROD_COLORS[2],
    },
  ];
  const totalQty = prodData.reduce((s, d) => s + d.qty, 0);

  const trendData = [...availableDates].sort().map((date) => {
    const rows = allServerData.filter(
      (d) =>
        d.sales_date === date &&
        (selectedBranch === "ALL" || d.branch === selectedBranch),
    );
    return {
      date,
      total: rows.reduce((s, d) => s + (d.total_bath || 0), 0),
      qty: rows.reduce(
        (s, d) =>
          s +
          ((d.tuangok_qty ?? d.tuangok ?? 0) +
            (d.khaomak_qty ?? d.khaomak ?? 0) +
            (d.tonon_qty ?? d.tonon ?? 0)),
        0,
      ),
    };
  });

  /* ── Shared filter controls ── */
  const FilterControls = ({ layout = "row" }) => (
    <>
      <select
        className="sel"
        value={selectedBranch}
        onChange={(e) => setSelectedBranch(e.target.value)}
        style={layout === "col" ? { width: "100%" } : {}}
      >
        <option value="ALL">ทุกสาขา</option>
        {VALID_BRANCHES.map((b) => (
          <option key={b} value={b}>
            {b}
          </option>
        ))}
      </select>

      <select
        className="sel"
        value={dateA}
        onChange={(e) => setDateA(e.target.value)}
        style={
          layout === "col"
            ? {
                width: "100%",
                borderColor: "var(--gold-lt)",
                color: "var(--gold)",
                fontWeight: 700,
              }
            : {
                borderColor: "var(--gold-lt)",
                color: "var(--gold)",
                fontWeight: 700,
              }
        }
      >
        {availableDates.map((d) => (
          <option key={d} value={d}>
            {d}
          </option>
        ))}
      </select>

      {layout === "row" && (
        <span
          style={{
            fontSize: 11,
            color: "var(--ink-3)",
            fontWeight: 600,
            letterSpacing: ".08em",
          }}
        >
          VS
        </span>
      )}

      <select
        className="sel"
        value={dateB}
        onChange={(e) => setDateB(e.target.value)}
        style={
          layout === "col"
            ? {
                width: "100%",
                borderColor: "var(--green-lt)",
                color: "var(--green)",
                fontWeight: 700,
              }
            : {
                borderColor: "var(--green-lt)",
                color: "var(--green)",
                fontWeight: 700,
              }
        }
      >
        <option value="">— เปรียบเทียบ —</option>
        {availableDates
          .filter((d) => d !== dateA)
          .map((d) => (
            <option key={d} value={d}>
              {d}
            </option>
          ))}
      </select>

      {dateA && (
        <button className="btn-ghost" onClick={() => handleDelete(dateA)}>
          ลบ {dateA}
        </button>
      )}

      <label className="btn-primary" style={{ cursor: "pointer" }}>
        {isUploading ? (
          "⏳ กำลังโหลด..."
        ) : (
          <>
            <svg
              width="12"
              height="12"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" />
              <polyline points="17 8 12 3 7 8" />
              <line x1="12" y1="3" x2="12" y2="15" />
            </svg>
            นำเข้า Excel
          </>
        )}
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileUpload}
          style={{ display: "none" }}
        />
      </label>
    </>
  );

  const tabs = [
    { id: "compare", label: "เปรียบเทียบยอดรวม" },
    { id: "product", label: "แยกประเภทสินค้า" },
    { id: "trend", label: "แนวโน้มยอดขาย" },
  ];

  return (
    <>
      <div
        style={{
          minHeight: "100vh",
          background: "var(--stone)",
          color: "var(--ink)",
        }}
      >
        {/* ── NAVBAR ── */}
        <nav className="nav">
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div
              style={{
                borderRight: "1px solid var(--stone-3)",
                paddingRight: 14,
                lineHeight: 1,
              }}
            >
              <div className="ornament">✦✦✦</div>
            </div>
            <div>
              <div
                style={{
                  fontFamily: "var(--font-display)",
                  fontWeight: 600,
                  fontSize: 18,
                  color: "var(--ink)",
                  letterSpacing: ".04em",
                  lineHeight: 1.1,
                }}
              >
                Revenue Dashboard
              </div>
              <div
                style={{
                  fontSize: 10,
                  color: "var(--gold)",
                  letterSpacing: ".12em",
                  textTransform: "uppercase",
                  marginTop: 2,
                }}
              >
                ระบบรายงานยอดขาย
              </div>
            </div>
          </div>
          <div className="desktop-controls">
            <FilterControls layout="row" />
          </div>
          <button className="hamburger" onClick={() => setMenuOpen(!menuOpen)}>
            {menuOpen ? "✕" : "☰"}
          </button>
        </nav>
        <div className="gold-line" />

        {/* ── MOBILE PANEL ── */}
        {menuOpen && (
          <div className="mobile-panel">
            <FilterControls layout="col" />
          </div>
        )}

        <main className="main-pad" style={{ maxWidth: 1160, margin: "0 auto" }}>
          {/* Header */}
          <div
            style={{
              marginBottom: 24,
              display: "flex",
              alignItems: "flex-end",
              justifyContent: "space-between",
              flexWrap: "wrap",
              gap: 8,
            }}
          >
            <div>
              <div className="section-label" style={{ marginBottom: 6 }}>
                สรุปผลประกอบการ
              </div>
              <h1
                style={{
                  fontFamily: "var(--font-display)",
                  fontWeight: 600,
                  fontSize: 28,
                  color: "var(--ink)",
                  letterSpacing: ".02em",
                  lineHeight: 1.1,
                }}
              >
                Sales Overview
              </h1>
              <p
                style={{
                  fontSize: 12,
                  color: "var(--ink-3)",
                  marginTop: 6,
                  letterSpacing: ".03em",
                }}
              >
                {availableDates.length} รอบข้อมูล · {VALID_BRANCHES.length} สาขา
                {dateA && (
                  <>
                    {" "}
                    ·{" "}
                    <span style={{ color: "var(--gold)", fontWeight: 600 }}>
                      {dateA}
                    </span>
                  </>
                )}
                {dateB && (
                  <>
                    {" "}
                    เทียบ{" "}
                    <span style={{ color: "var(--green)", fontWeight: 600 }}>
                      {dateB}
                    </span>
                  </>
                )}
              </p>
            </div>
            <div className="ornament" style={{ fontSize: 18 }}>
              ◆
            </div>
          </div>

          {/* KPI Grid */}
          <div
            className="kpi-grid"
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(4,1fr)",
              gap: 14,
              marginBottom: 24,
            }}
          >
            <KpiCard
              label="ยอดขายรวม (รอบหลัก)"
              value={`฿${totalA.toLocaleString()}`}
              sub={dateA || "—"}
              icon="◈"
              isLarge
            />
            <KpiCard
              label="ยอดขายรวม (รอบเทียบ)"
              value={`฿${totalB.toLocaleString()}`}
              sub={dateB || "ยังไม่ได้เลือก"}
              icon="◈"
            />
            <KpiCard
              label="ส่วนต่างจากรอบเทียบ"
              value={pct && dateB ? `${diff >= 0 ? "+" : ""}${pct}%` : "—"}
              sub={
                dateB
                  ? `฿${Math.abs(diff).toLocaleString()} บาท`
                  : "เลือกรอบเทียบ"
              }
              icon={diff >= 0 ? "▲" : "▼"}
            />
            <KpiCard
              label="สาขาที่ขายดีสุด"
              value={top?.branch || "—"}
              sub={top ? `฿${top.a_total.toLocaleString()}` : "—"}
              icon="✦"
            />
          </div>

          {/* Chart Card */}
          <div className="card" style={{ overflow: "hidden" }}>
            <div className="tab-bar">
              {tabs.map((t) => (
                <button
                  key={t.id}
                  className={`tab-btn ${activeTab === t.id ? "active" : ""}`}
                  onClick={() => setActiveTab(t.id)}
                >
                  {t.label}
                </button>
              ))}
            </div>

            <div style={{ padding: "24px 24px 32px" }}>
              {/* ── TAB: เปรียบเทียบ ── */}
              {activeTab === "compare" && (
                <>
                  <div style={{ height: 280, marginBottom: 28 }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={filtered} barGap={4} barCategoryGap="32%">
                        <CartesianGrid
                          strokeDasharray="2 4"
                          stroke="var(--stone-2)"
                          vertical={false}
                        />
                        <XAxis
                          dataKey="short"
                          tick={{
                            fill: "var(--ink-3)",
                            fontSize: 11,
                            letterSpacing: 2,
                          }}
                          axisLine={false}
                          tickLine={false}
                        />
                        <YAxis
                          tick={{ fill: "var(--ink-3)", fontSize: 10 }}
                          axisLine={false}
                          tickLine={false}
                          tickFormatter={(v) =>
                            v >= 1000 ? `${(v / 1000).toFixed(0)}K` : String(v)
                          }
                          width={38}
                        />
                        <Tooltip
                          content={<CustomTooltip />}
                          cursor={{ fill: "var(--stone)", opacity: 0.8 }}
                        />
                        <Legend
                          wrapperStyle={{
                            fontSize: 11,
                            paddingTop: 14,
                            color: "var(--ink-3)",
                            letterSpacing: ".04em",
                          }}
                        />
                        {dateB && (
                          <Bar
                            dataKey="b_total"
                            name={dateB}
                            fill="var(--stone-2)"
                            stroke="var(--stone-3)"
                            strokeWidth={1}
                            radius={[3, 3, 0, 0]}
                          />
                        )}
                        <Bar
                          dataKey="a_total"
                          name={dateA}
                          fill="var(--gold)"
                          radius={[3, 3, 0, 0]}
                        />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>

                  {/* Table */}
                  <div className="table-wrap">
                    <table
                      className="data-table"
                      style={{
                        width: "100%",
                        borderCollapse: "separate",
                        borderSpacing: "0 2px",
                        fontSize: 12.5,
                      }}
                    >
                      <thead>
                        <tr
                          style={{
                            fontSize: 10,
                            color: "var(--ink-3)",
                            textTransform: "uppercase",
                            letterSpacing: ".1em",
                          }}
                        >
                          {[
                            "สาขา",
                            "รอบเทียบ",
                            "รอบหลัก",
                            "ส่วนต่าง",
                            "เติบโต",
                            "จำนวนสินค้า",
                          ].map((h, i) => (
                            <th
                              key={i}
                              style={{
                                padding: "6px 14px",
                                textAlign: i === 0 ? "left" : "right",
                                fontWeight: 600,
                              }}
                            >
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {filtered.map((item, i) => {
                          const d = item.a_total - item.b_total;
                          const p =
                            item.b_total > 0
                              ? ((d / item.b_total) * 100).toFixed(1)
                              : null;
                          const pos = d >= 0;
                          const totalItemQty =
                            item.a_tuangok_qty +
                            item.a_khaomak_qty +
                            item.a_tonon_qty;
                          return (
                            <tr key={i} className="row-hover">
                              <td
                                style={{
                                  padding: "12px 14px",
                                  borderRadius: "6px 0 0 6px",
                                }}
                              >
                                <div
                                  style={{
                                    display: "flex",
                                    alignItems: "center",
                                    gap: 8,
                                  }}
                                >
                                  <div
                                    style={{
                                      width: 8,
                                      height: 8,
                                      borderRadius: 2,
                                      background: BRANCH_COLORS[i % 5],
                                      flexShrink: 0,
                                    }}
                                  />
                                  <span
                                    style={{
                                      fontWeight: 600,
                                      color: "var(--ink)",
                                    }}
                                  >
                                    {item.branch}
                                  </span>
                                </div>
                                {/* per-branch product qty pills */}
                                <div
                                  className="qty-pills"
                                  style={{ marginTop: 5 }}
                                >
                                  {[
                                    {
                                      name: "ถั่วงอก",
                                      qty: item.a_tuangok_qty,
                                      color: PROD_COLORS[0],
                                    },
                                    {
                                      name: "ข้าวหมาก",
                                      qty: item.a_khaomak_qty,
                                      color: PROD_COLORS[1],
                                    },
                                    {
                                      name: "ต้นอ่อน",
                                      qty: item.a_tonon_qty,
                                      color: PROD_COLORS[2],
                                    },
                                  ]
                                    .filter((p) => p.qty > 0)
                                    .map((p, j) => (
                                      <span key={j} className="qty-pill">
                                        <span
                                          className="dot"
                                          style={{ background: p.color }}
                                        />
                                        {p.name}{" "}
                                        <span className="qty-num">
                                          {p.qty.toLocaleString()}
                                        </span>{" "}
                                        ชิ้น
                                      </span>
                                    ))}
                                </div>
                              </td>
                              <td
                                className="num"
                                style={{
                                  padding: "12px 14px",
                                  textAlign: "right",
                                  color: "var(--ink-3)",
                                }}
                              >
                                {item.b_total > 0
                                  ? `฿${item.b_total.toLocaleString()}`
                                  : "—"}
                              </td>
                              <td
                                className="num"
                                style={{
                                  padding: "12px 14px",
                                  textAlign: "right",
                                  color: "var(--gold)",
                                  fontWeight: 700,
                                }}
                              >
                                ฿{item.a_total.toLocaleString()}
                              </td>
                              <td
                                className="num"
                                style={{
                                  padding: "12px 14px",
                                  textAlign: "right",
                                  fontWeight: 600,
                                  color: pos ? "var(--green)" : "var(--red)",
                                }}
                              >
                                {dateB
                                  ? `${pos ? "+" : ""}฿${d.toLocaleString()}`
                                  : "—"}
                              </td>
                              <td
                                style={{
                                  padding: "12px 14px",
                                  textAlign: "right",
                                }}
                              >
                                {p && dateB ? (
                                  <span
                                    className={`badge ${pos ? "pos" : "neg"}`}
                                  >
                                    {pos ? "▲" : "▼"} {Math.abs(Number(p))}%
                                  </span>
                                ) : (
                                  "—"
                                )}
                              </td>
                              <td
                                className="num"
                                style={{
                                  padding: "12px 14px",
                                  textAlign: "right",
                                  borderRadius: "0 6px 6px 0",
                                  color: "var(--ink-2)",
                                  fontWeight: 600,
                                }}
                              >
                                {totalItemQty > 0
                                  ? `${totalItemQty.toLocaleString()} ชิ้น`
                                  : "—"}
                              </td>
                            </tr>
                          );
                        })}
                        {filtered.length > 1 && (
                          <tr style={{ borderTop: "1px solid var(--stone-3)" }}>
                            <td
                              style={{
                                padding: "12px 14px",
                                borderRadius: "6px 0 0 6px",
                                fontWeight: 700,
                                fontSize: 12,
                                color: "var(--ink-2)",
                                letterSpacing: ".04em",
                                textTransform: "uppercase",
                              }}
                            >
                              รวมทั้งหมด
                            </td>
                            <td
                              className="num"
                              style={{
                                padding: "12px 14px",
                                textAlign: "right",
                                color: "var(--ink-3)",
                                fontWeight: 600,
                              }}
                            >
                              {totalB > 0 ? `฿${totalB.toLocaleString()}` : "—"}
                            </td>
                            <td
                              className="num"
                              style={{
                                padding: "12px 14px",
                                textAlign: "right",
                                color: "var(--gold)",
                                fontWeight: 700,
                                fontSize: 14,
                              }}
                            >
                              ฿{totalA.toLocaleString()}
                            </td>
                            <td
                              className="num"
                              style={{
                                padding: "12px 14px",
                                textAlign: "right",
                                fontWeight: 700,
                                color:
                                  diff >= 0 ? "var(--green)" : "var(--red)",
                              }}
                            >
                              {dateB
                                ? `${diff >= 0 ? "+" : ""}฿${diff.toLocaleString()}`
                                : "—"}
                            </td>
                            <td
                              style={{
                                padding: "12px 14px",
                                textAlign: "right",
                              }}
                            >
                              {pct && dateB ? (
                                <span
                                  className={`badge ${diff >= 0 ? "pos" : "neg"}`}
                                >
                                  {diff >= 0 ? "▲" : "▼"}{" "}
                                  {Math.abs(Number(pct))}%
                                </span>
                              ) : (
                                "—"
                              )}
                            </td>
                            <td
                              className="num"
                              style={{
                                padding: "12px 14px",
                                textAlign: "right",
                                borderRadius: "0 6px 6px 0",
                                color: "var(--ink-2)",
                                fontWeight: 700,
                              }}
                            >
                              {totalQty > 0
                                ? `${totalQty.toLocaleString()} ชิ้น`
                                : "—"}
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </>
              )}

              {/* ── TAB: ประเภทสินค้า ── */}
              {activeTab === "product" && (
                <div className="chart-layout">
                  <div className="chart-main">
                    <div className="section-label" style={{ marginBottom: 14 }}>
                      จำนวนสินค้าแยกตามประเภท (ชิ้น) — {dateA}
                    </div>
                    <div style={{ height: 280 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={filtered}
                          barGap={3}
                          barCategoryGap="26%"
                        >
                          <CartesianGrid
                            strokeDasharray="2 4"
                            stroke="var(--stone-2)"
                            vertical={false}
                          />
                          <XAxis
                            dataKey="short"
                            tick={{ fill: "var(--ink-3)", fontSize: 11 }}
                            axisLine={false}
                            tickLine={false}
                          />
                          <YAxis
                            tick={{ fill: "var(--ink-3)", fontSize: 10 }}
                            axisLine={false}
                            tickLine={false}
                            tickFormatter={(v) =>
                              v >= 1000
                                ? `${(v / 1000).toFixed(0)}K`
                                : String(v)
                            }
                            width={38}
                          />
                          <Tooltip
                            content={<QtyTooltip />}
                            cursor={{ fill: "var(--stone)", opacity: 0.8 }}
                          />
                          <Legend
                            wrapperStyle={{
                              fontSize: 11,
                              paddingTop: 14,
                              color: "var(--ink-3)",
                            }}
                          />
                          <Bar
                            dataKey="a_tuangok_qty"
                            name="ถั่วงอก"
                            fill={PROD_COLORS[0]}
                            radius={[3, 3, 0, 0]}
                          />
                          <Bar
                            dataKey="a_khaomak_qty"
                            name="ข้าวหมาก"
                            fill={PROD_COLORS[1]}
                            radius={[3, 3, 0, 0]}
                          />
                          <Bar
                            dataKey="a_tonon_qty"
                            name="ต้นอ่อน"
                            fill={PROD_COLORS[2]}
                            radius={[3, 3, 0, 0]}
                          />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="chart-side">
                    <div className="section-label" style={{ marginBottom: 14 }}>
                      สัดส่วนและจำนวนสินค้า
                    </div>
                    <div style={{ height: 180 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={prodData}
                            cx="50%"
                            cy="50%"
                            innerRadius={48}
                            outerRadius={75}
                            paddingAngle={4}
                            dataKey="qty"
                            startAngle={90}
                            endAngle={-270}
                          >
                            {prodData.map((_, i) => (
                              <Cell key={i} fill={PROD_COLORS[i]} />
                            ))}
                          </Pie>
                          <Tooltip
                            formatter={(v) => [
                              `${Number(v).toLocaleString()} ชิ้น`,
                              "",
                            ]}
                            contentStyle={{
                              background: "var(--white)",
                              border: "1px solid var(--stone-3)",
                              borderRadius: 8,
                              fontSize: 12,
                            }}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>

                    {/* Product breakdown with qty */}
                    <div style={{ marginTop: 8 }}>
                      {prodData.map((item, i) => {
                        const share =
                          totalQty > 0
                            ? ((item.qty / totalQty) * 100).toFixed(0)
                            : 0;
                        return (
                          <div
                            key={i}
                            style={{
                              padding: "10px 0",
                              borderBottom: "1px solid var(--stone-2)",
                            }}
                          >
                            <div
                              style={{
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "space-between",
                                gap: 12,
                              }}
                            >
                              <div
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  gap: 8,
                                }}
                              >
                                <div
                                  style={{
                                    width: 8,
                                    height: 8,
                                    borderRadius: 2,
                                    background: item.color,
                                    flexShrink: 0,
                                  }}
                                />
                                <span
                                  style={{
                                    color: "var(--ink-2)",
                                    fontSize: 13,
                                  }}
                                >
                                  {item.name}
                                </span>
                              </div>
                              <div style={{ textAlign: "right" }}>
                                <div
                                  className="num"
                                  style={{
                                    fontWeight: 700,
                                    color:
                                      item.qty > 0
                                        ? item.color
                                        : "var(--ink-3)",
                                    fontSize: 14,
                                  }}
                                >
                                  {item.qty > 0
                                    ? `${item.qty.toLocaleString()} ชิ้น`
                                    : "—"}
                                </div>
                                <div
                                  className="num"
                                  style={{
                                    fontSize: 10,
                                    color: "var(--ink-3)",
                                    marginTop: 1,
                                  }}
                                >
                                  {share}%
                                </div>
                              </div>
                            </div>
                          </div>
                        );
                      })}

                      {/* Total qty summary */}
                      {totalQty > 0 && (
                        <div
                          style={{
                            marginTop: 12,
                            padding: "10px 14px",
                            background: "var(--gold-dim)",
                            border: "1px solid var(--gold-lt)",
                            borderRadius: 8,
                          }}
                        >
                          <div
                            style={{
                              fontSize: 10,
                              color: "var(--gold)",
                              fontWeight: 600,
                              letterSpacing: ".08em",
                              textTransform: "uppercase",
                              marginBottom: 4,
                            }}
                          >
                            จำนวนรวมทั้งหมด
                          </div>
                          <div
                            className="num"
                            style={{
                              fontWeight: 700,
                              fontSize: 18,
                              color: "var(--ink)",
                            }}
                          >
                            {totalQty.toLocaleString()}{" "}
                            <span
                              style={{ fontSize: 12, color: "var(--ink-3)" }}
                            >
                              ชิ้น
                            </span>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              )}

              {/* ── TAB: แนวโน้ม ── */}
              {activeTab === "trend" && (
                <>
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "space-between",
                      marginBottom: 16,
                      flexWrap: "wrap",
                      gap: 8,
                    }}
                  >
                    <div>
                      <div
                        className="section-label"
                        style={{ marginBottom: 4 }}
                      >
                        แนวโน้มยอดขายรวมทุกรอบ
                      </div>
                      {selectedBranch !== "ALL" && (
                        <div
                          style={{
                            color: "var(--gold)",
                            fontSize: 13,
                            fontWeight: 600,
                          }}
                        >
                          {selectedBranch}
                        </div>
                      )}
                    </div>
                    {totalQty > 0 && (
                      <div
                        style={{
                          padding: "8px 16px",
                          background: "var(--gold-dim)",
                          border: "1px solid var(--gold-lt)",
                          borderRadius: 8,
                          textAlign: "center",
                        }}
                      >
                        <div
                          style={{
                            fontSize: 9,
                            color: "var(--gold)",
                            fontWeight: 600,
                            letterSpacing: ".1em",
                            textTransform: "uppercase",
                          }}
                        >
                          จำนวนสินค้ารวม
                        </div>
                        <div
                          className="num"
                          style={{
                            fontWeight: 700,
                            fontSize: 16,
                            color: "var(--ink)",
                          }}
                        >
                          {totalQty.toLocaleString()} ชิ้น
                        </div>
                      </div>
                    )}
                  </div>

                  <div style={{ height: 260 }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={trendData}>
                        <defs>
                          <linearGradient
                            id="goldGrad"
                            x1="0"
                            y1="0"
                            x2="0"
                            y2="1"
                          >
                            <stop
                              offset="5%"
                              stopColor="var(--gold)"
                              stopOpacity={0.15}
                            />
                            <stop
                              offset="95%"
                              stopColor="var(--gold)"
                              stopOpacity={0}
                            />
                          </linearGradient>
                        </defs>
                        <CartesianGrid
                          strokeDasharray="2 4"
                          stroke="var(--stone-2)"
                          vertical={false}
                        />
                        <XAxis
                          dataKey="date"
                          tick={{ fill: "var(--ink-3)", fontSize: 11 }}
                          axisLine={false}
                          tickLine={false}
                        />
                        <YAxis
                          tick={{ fill: "var(--ink-3)", fontSize: 10 }}
                          axisLine={false}
                          tickLine={false}
                          tickFormatter={(v) =>
                            v >= 1000 ? `${(v / 1000).toFixed(0)}K` : String(v)
                          }
                          width={38}
                        />
                        <Tooltip
                          content={<CustomTooltip />}
                          cursor={{
                            stroke: "var(--gold)",
                            strokeWidth: 1,
                            strokeDasharray: "4 4",
                          }}
                        />
                        <Area
                          type="monotone"
                          dataKey="total"
                          name="ยอดรวม"
                          stroke="var(--gold)"
                          strokeWidth={2}
                          fill="url(#goldGrad)"
                          dot={{
                            fill: "var(--gold)",
                            r: 4,
                            strokeWidth: 2,
                            stroke: "var(--white)",
                          }}
                          activeDot={{
                            r: 6,
                            stroke: "var(--gold)",
                            strokeWidth: 2,
                            fill: "var(--white)",
                          }}
                        />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>

                  {/* Trend cards with qty */}
                  <div className="trend-cards">
                    {trendData.map((d, i) => {
                      const maxVal = Math.max(...trendData.map((t) => t.total));
                      const isTop = d.total === maxVal && maxVal > 0;
                      return (
                        <div
                          key={i}
                          className="trend-card"
                          style={{
                            background: isTop
                              ? "var(--gold-dim)"
                              : "var(--stone)",
                            border: `1px solid ${isTop ? "var(--gold-lt)" : "var(--stone-3)"}`,
                            borderTop: `2px solid ${isTop ? "var(--gold)" : "var(--stone-3)"}`,
                            borderRadius: 10,
                            padding: "12px 14px",
                          }}
                        >
                          <div
                            style={{
                              fontSize: 10,
                              color: "var(--ink-3)",
                              marginBottom: 4,
                              letterSpacing: ".04em",
                            }}
                          >
                            {d.date}
                          </div>
                          <div
                            className="num"
                            style={{
                              fontSize: 15,
                              fontWeight: 700,
                              color: isTop ? "var(--gold)" : "var(--ink)",
                            }}
                          >
                            ฿{d.total.toLocaleString()}
                          </div>
                          {d.qty > 0 && (
                            <div
                              style={{
                                fontSize: 10,
                                color: "var(--ink-3)",
                                marginTop: 3,
                              }}
                            >
                              <span
                                className="num"
                                style={{
                                  fontWeight: 600,
                                  color: "var(--ink-2)",
                                }}
                              >
                                {d.qty.toLocaleString()}
                              </span>{" "}
                              ชิ้น
                            </div>
                          )}
                          {isTop && (
                            <div
                              style={{
                                fontSize: 10,
                                color: "var(--gold)",
                                marginTop: 4,
                                fontWeight: 600,
                                letterSpacing: ".06em",
                              }}
                            >
                              ✦ สูงสุด
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </>
              )}
            </div>
          </div>

          {/* Footer */}
          <div style={{ textAlign: "center", marginTop: 32, paddingBottom: 8 }}>
            <div className="ornament" style={{ marginBottom: 6 }}>
              ✦ ✦ ✦
            </div>
            <div
              style={{
                fontSize: 10,
                color: "var(--ink-3)",
                letterSpacing: ".1em",
                textTransform: "uppercase",
              }}
            >
              Revenue Dashboard ·{" "}
              {new Date().toLocaleDateString("th-TH", {
                year: "numeric",
                month: "long",
                day: "numeric",
              })}
            </div>
          </div>
        </main>
      </div>
    </>
  );
}

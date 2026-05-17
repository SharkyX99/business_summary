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

const fontLink = `
  @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    background: #f5f6fa;
    font-family: 'Kurious Looped', 'Sarabun', 'Noto Sans Thai', sans-serif;
  }
  button, select, input, label {
    font-family: 'Kurious Looped', 'Sarabun', 'Noto Sans Thai', sans-serif;
  }
  .num {
    font-family: 'Kurious Looped', 'Sarabun', sans-serif;
    font-variant-numeric: tabular-nums;
    font-feature-settings: "tnum";
  }
  .card { transition: box-shadow 0.2s, transform 0.2s; }
  .card:hover { box-shadow: 0 8px 28px rgba(0,0,0,0.10) !important; transform: translateY(-1px); }
  .row-hover:hover { background: #f8fafc !important; }
  .tab-btn:hover { color: #1d4ed8 !important; }
  .upload-btn:hover { background: #1d4ed8 !important; }
  .del-btn:hover { background: #fef2f2 !important; color: #b91c1c !important; }

  /* ── Responsive ────────────────────────── */
  .kpi-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin-bottom: 24px;
  }
  @media (max-width: 1024px) {
    .kpi-grid { grid-template-columns: repeat(2, 1fr); }
  }
  @media (max-width: 600px) {
    .kpi-grid { grid-template-columns: repeat(2, 1fr); gap: 8px; }
  }

  .nav-controls {
    display: flex;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
  }
  @media (max-width: 768px) {
    .nav-controls { display: none; }
    .nav-controls.open { display: flex; flex-direction: column; align-items: stretch; }
  }

  .mobile-filter-bar {
    display: none;
    background: #fff;
    border-bottom: 1px solid #e5e7eb;
    padding: 12px 16px;
    flex-direction: column;
    gap: 10px;
  }
  @media (max-width: 768px) {
    .mobile-filter-bar { display: flex; }
  }

  .hamburger {
    display: none;
    background: none;
    border: 1px solid #e5e7eb;
    border-radius: 8px;
    padding: 6px 10px;
    cursor: pointer;
    font-size: 18px;
    color: #374151;
    align-items: center;
    justify-content: center;
  }
  @media (max-width: 768px) {
    .hamburger { display: flex; }
    .desktop-controls { display: none !important; }
  }

  .tab-bar {
    display: flex;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
    scrollbar-width: none;
    padding: 0 16px;
  }
  .tab-bar::-webkit-scrollbar { display: none; }

  .tab-btn {
    white-space: nowrap;
    flex-shrink: 0;
  }

  .chart-product-layout {
    display: flex;
    gap: 32px;
    flex-wrap: wrap;
  }
  .chart-product-main { flex: 2; min-width: 260px; }
  .chart-product-side { flex: 1; min-width: 200px; }
  @media (max-width: 600px) {
    .chart-product-main { min-width: 100%; }
    .chart-product-side { min-width: 100%; }
  }

  .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  .data-table { min-width: 520px; }

  .trend-cards {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-top: 20px;
  }
  .trend-card { min-width: 110px; flex: 1; }

  .main-pad { padding: 24px 20px; }
  @media (max-width: 600px) {
    .main-pad { padding: 16px 12px; }
  }

  .page-title { font-size: 20px; }
  @media (max-width: 600px) {
    .page-title { font-size: 17px; }
  }

  .kpi-value-large { font-size: 22px; }
  .kpi-value-medium { font-size: 16px; }
  @media (max-width: 600px) {
    .kpi-value-large { font-size: 18px; }
    .kpi-value-medium { font-size: 14px; }
  }

  select, .sel { max-width: 100%; }
`;

const VALID_BRANCHES = [
  "Central Festival",
  "Central Airport",
  "Chotana",
  "Ruamchok Mall",
  "Thaweechok",
];
const BRANCH_COLORS = ["#3b82f6", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6"];
const PROD_COLORS = ["#3b82f6", "#10b981", "#f59e0b"];

function KpiCard({
  label,
  value,
  sub,
  color,
  icon,
}: {
  label: string;
  value: string;
  sub?: string;
  color: string;
  icon: string;
}) {
  const isLong = value.length > 14;
  return (
    <div
      className="card"
      style={{
        background: "#fff",
        borderRadius: 12,
        padding: "16px 18px",
        border: "1px solid #e8eaf0",
        borderTop: `3px solid ${color}`,
        boxShadow: "0 2px 8px rgba(0,0,0,0.04)",
      }}
    >
      <div style={{ fontSize: 18, marginBottom: 8 }}>{icon}</div>
      <div
        style={{
          fontSize: 10,
          color: "#9ca3af",
          fontWeight: 600,
          letterSpacing: "0.07em",
          textTransform: "uppercase",
          marginBottom: 5,
        }}
      >
        {label}
      </div>
      <div
        className={`num ${isLong ? "kpi-value-medium" : "kpi-value-large"}`}
        style={{
          fontWeight: 700,
          color: "#111827",
          lineHeight: 1.25,
          wordBreak: "break-word",
        }}
      >
        {value}
      </div>
      {sub && (
        <div style={{ fontSize: 11, color, marginTop: 5, fontWeight: 500 }}>
          {sub}
        </div>
      )}
    </div>
  );
}

function CustomTooltip({ active, payload, label }: any) {
  if (!active || !payload?.length) return null;
  return (
    <div
      style={{
        background: "#fff",
        border: "1px solid #e5e7eb",
        borderRadius: 10,
        padding: "10px 14px",
        fontSize: 12,
        boxShadow: "0 4px 20px rgba(0,0,0,0.10)",
        maxWidth: 220,
      }}
    >
      <div style={{ fontWeight: 700, color: "#374151", marginBottom: 6 }}>
        {label}
      </div>
      {payload.map((p: any, i: number) => (
        <div
          key={i}
          style={{
            display: "flex",
            justifyContent: "space-between",
            gap: 12,
            marginBottom: 2,
          }}
        >
          <span style={{ color: p.color, fontWeight: 500 }}>● {p.name}</span>
          <span className="num" style={{ fontWeight: 700, color: "#111827" }}>
            ฿{Number(p.value).toLocaleString()}
          </span>
        </div>
      ))}
    </div>
  );
}

const selStyle: React.CSSProperties = {
  background: "#f9fafb",
  border: "1px solid #e5e7eb",
  borderRadius: 8,
  color: "#374151",
  padding: "8px 10px",
  fontSize: 13,
  cursor: "pointer",
  width: "100%",
};

export default function Home() {
  const [allServerData, setAllServerData] = useState<any[]>([]);
  const [availableDates, setAvailableDates] = useState<string[]>([]);
  const [comparisonData, setComparisonData] = useState<any[]>([]);
  const [dateA, setDateA] = useState("");
  const [dateB, setDateB] = useState("");
  const [selectedBranch, setSelectedBranch] = useState("ALL");
  const [isUploading, setIsUploading] = useState(false);
  const [activeTab, setActiveTab] = useState<"compare" | "product" | "trend">(
    "compare",
  );
  const [menuOpen, setMenuOpen] = useState(false);

  const fetchRecords = async () => {
    try {
      const res = await fetch("/api/insert");
      const data = await res.json();
      if (!Array.isArray(data)) return;
      setAllServerData(data);
      const dates: string[] = Array.from(
        new Set(
          data.filter((i: any) => i.sales_date).map((i: any) => i.sales_date),
        ),
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
        a_tuangok: a.tuangok || 0,
        a_khaomak: a.khaomak || 0,
        a_tonon: a.tonon || 0,
      };
    });
    setComparisonData(merged);
  }, [dateA, dateB, allServerData]);

  const handleDelete = async (date: string) => {
    if (!confirm(`ลบข้อมูลวันที่ ${date} ทั้งหมด?`)) return;
    try {
      const res = await fetch("/api/insert", {
        method: "DELETE",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sales_date: date }),
      });
      const r = await res.json();
      if (r.success) {
        alert("ลบเรียบร้อย");
        await fetchRecords();
        window.location.reload();
      } else alert("ลบไม่สำเร็จ: " + r.error);
    } catch {
      alert("เชื่อมต่อไม่ได้");
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsUploading(true);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      const wb = XLSX.read(evt.target?.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const rows: any[] = [];
      raw.forEach((row) => {
        const b = row[1] ? String(row[1]).trim() : "";
        if (b && VALID_BRANCHES.includes(b) && row[2]) {
          rows.push({
            branch: b,
            sales_date: XLSX.SSF.format("yyyy-mm-dd", row[2]),
            tuangok: Number(row[3]) || 0,
            khaomak: Number(row[4]) || 0,
            tonon: Number(row[5]) || 0,
            total_bath: Number(row[6]) || 0,
          });
        }
      });
      if (!rows.length) {
        alert("ไม่พบข้อมูลสาขาในไฟล์");
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
          alert("✅ บันทึกสำเร็จ");
          await fetchRecords();
          window.location.reload();
        } else alert("ผิดพลาด: " + r.error);
      } catch {
        alert("ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้");
      } finally {
        setIsUploading(false);
      }
    };
    reader.readAsBinaryString(file);
  };

  const filtered =
    selectedBranch === "ALL"
      ? comparisonData
      : comparisonData.filter((d) => d.branch === selectedBranch);
  const totalA = filtered.reduce((s, d) => s + d.a_total, 0);
  const totalB = filtered.reduce((s, d) => s + d.b_total, 0);
  const diff = totalA - totalB;
  const pct = totalB > 0 ? ((diff / totalB) * 100).toFixed(1) : null;
  const top = [...comparisonData].sort((a, b) => b.a_total - a.a_total)[0];

  const pieData = [
    { name: "ถั่วงอก", value: filtered.reduce((s, d) => s + d.a_tuangok, 0) },
    { name: "ข้าวหมาก", value: filtered.reduce((s, d) => s + d.a_khaomak, 0) },
    { name: "ต้นอ่อน", value: filtered.reduce((s, d) => s + d.a_tonon, 0) },
  ];

  const trendData = [...availableDates].sort().map((date) => {
    const rows = allServerData.filter(
      (d) =>
        d.sales_date === date &&
        (selectedBranch === "ALL" || d.branch === selectedBranch),
    );
    return {
      date,
      total: rows.reduce((s: number, d: any) => s + (d.total_bath || 0), 0),
    };
  });

  const tabs = [
    { id: "compare" as const, label: "เปรียบเทียบยอดรวม" },
    { id: "product" as const, label: "แยกประเภทสินค้า" },
    { id: "trend" as const, label: "แนวโน้มยอดขาย" },
  ];

  // Shared filter controls (used in both desktop navbar & mobile panel)
  const FilterControls = ({ layout }: { layout: "row" | "col" }) => (
    <>
      <div
        style={{
          display: "flex",
          flexDirection: layout === "col" ? "column" : "row",
          gap: 8,
          flex: layout === "col" ? 1 : undefined,
        }}
      >
        {layout === "col" && (
          <div
            style={{
              fontSize: 11,
              color: "#9ca3af",
              fontWeight: 600,
              textTransform: "uppercase",
              letterSpacing: "0.06em",
            }}
          >
            สาขา
          </div>
        )}
        <select
          value={selectedBranch}
          onChange={(e) => setSelectedBranch(e.target.value)}
          style={layout === "col" ? selStyle : { ...selStyle, width: "auto" }}
        >
          <option value="ALL">ทุกสาขา</option>
          {VALID_BRANCHES.map((b) => (
            <option key={b} value={b}>
              {b}
            </option>
          ))}
        </select>
      </div>

      <div
        style={{
          display: "flex",
          flexDirection: layout === "col" ? "column" : "row",
          gap: 8,
          flex: layout === "col" ? 1 : undefined,
        }}
      >
        {layout === "col" && (
          <div
            style={{
              fontSize: 11,
              color: "#9ca3af",
              fontWeight: 600,
              textTransform: "uppercase",
              letterSpacing: "0.06em",
            }}
          >
            รอบหลัก
          </div>
        )}
        <select
          value={dateA}
          onChange={(e) => setDateA(e.target.value)}
          style={
            layout === "col"
              ? {
                  ...selStyle,
                  borderColor: "#bfdbfe",
                  color: "#1d4ed8",
                  fontWeight: 700,
                }
              : {
                  ...selStyle,
                  width: "auto",
                  borderColor: "#bfdbfe",
                  color: "#1d4ed8",
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
      </div>

      {layout === "row" && (
        <span style={{ fontSize: 13, color: "#9ca3af", fontWeight: 600 }}>
          vs
        </span>
      )}

      <div
        style={{
          display: "flex",
          flexDirection: layout === "col" ? "column" : "row",
          gap: 8,
          flex: layout === "col" ? 1 : undefined,
        }}
      >
        {layout === "col" && (
          <div
            style={{
              fontSize: 11,
              color: "#9ca3af",
              fontWeight: 600,
              textTransform: "uppercase",
              letterSpacing: "0.06em",
            }}
          >
            รอบเทียบ
          </div>
        )}
        <select
          value={dateB}
          onChange={(e) => setDateB(e.target.value)}
          style={
            layout === "col"
              ? {
                  ...selStyle,
                  borderColor: "#a7f3d0",
                  color: "#059669",
                  fontWeight: 700,
                }
              : {
                  ...selStyle,
                  width: "auto",
                  borderColor: "#a7f3d0",
                  color: "#059669",
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
      </div>

      {dateA && (
        <button
          className="del-btn"
          onClick={() => handleDelete(dateA)}
          style={{
            background: "transparent",
            border: "1px solid #fca5a5",
            color: "#ef4444",
            borderRadius: 8,
            padding: "8px 12px",
            fontSize: 12,
            cursor: "pointer",
            fontWeight: 600,
            transition: "all 0.15s",
            whiteSpace: "nowrap",
            width: layout === "col" ? "100%" : "auto",
          }}
        >
          🗑 ลบ {dateA}
        </button>
      )}

      <label
        className="upload-btn"
        style={{
          background: "#2563eb",
          color: "#fff",
          borderRadius: 8,
          padding: "9px 16px",
          fontWeight: 700,
          fontSize: 13,
          cursor: "pointer",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          gap: 6,
          boxShadow: "0 2px 8px #2563eb33",
          transition: "background 0.15s",
          whiteSpace: "nowrap",
          width: layout === "col" ? "100%" : "auto",
        }}
      >
        {isUploading ? "⏳ กำลังโหลด..." : "📤 นำเข้า Excel"}
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileUpload}
          style={{ display: "none" }}
        />
      </label>
    </>
  );

  return (
    <>
      <style>{fontLink}</style>
      <div
        style={{ minHeight: "100vh", background: "#f5f6fa", color: "#111827" }}
      >
        {/* ── NAVBAR ── */}
        <nav
          style={{
            background: "#fff",
            borderBottom: "1px solid #e5e7eb",
            padding: "0 20px",
            height: 60,
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            position: "sticky",
            top: 0,
            zIndex: 50,
            boxShadow: "0 1px 4px rgba(0,0,0,0.06)",
          }}
        >
          {/* Logo */}
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div
              style={{
                width: 34,
                height: 34,
                borderRadius: 9,
                flexShrink: 0,
                background: "linear-gradient(135deg,#2563eb,#1d4ed8)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 17,
              }}
            >
              📊
            </div>
            <div>
              <div
                style={{
                  fontWeight: 700,
                  fontSize: 15,
                  color: "#111827",
                  lineHeight: 1.2,
                }}
              >
                Revenue Dashboard
              </div>
              <div style={{ fontSize: 10, color: "#9ca3af" }}>
                ระบบรายงานยอดขาย
              </div>
            </div>
          </div>

          {/* Desktop controls */}
          <div
            className="desktop-controls"
            style={{ display: "flex", alignItems: "center", gap: 8 }}
          >
            <FilterControls layout="row" />
          </div>

          {/* Mobile hamburger */}
          <button
            className="hamburger"
            onClick={() => setMenuOpen(!menuOpen)}
            aria-label="เปิดเมนู"
          >
            {menuOpen ? "✕" : "☰"}
          </button>
        </nav>

        {/* ── MOBILE FILTER PANEL ── */}
        {menuOpen && (
          <div className="mobile-filter-bar">
            <FilterControls layout="col" />
          </div>
        )}

        {/* ── BODY ── */}
        <main className="main-pad" style={{ maxWidth: 1180, margin: "0 auto" }}>
          {/* Page title */}
          <div style={{ marginBottom: 20 }}>
            <h1
              className="page-title"
              style={{ fontWeight: 700, color: "#111827", marginBottom: 4 }}
            >
              สรุปยอดขายสาขา
            </h1>
            <p style={{ fontSize: 13, color: "#6b7280", lineHeight: 1.6 }}>
              {availableDates.length} รอบข้อมูล · {VALID_BRANCHES.length} สาขา
              {dateA && (
                <>
                  {" "}
                  ·{" "}
                  <span style={{ color: "#2563eb", fontWeight: 600 }}>
                    แสดง {dateA}
                  </span>
                </>
              )}
              {dateB && (
                <>
                  {" "}
                  เทียบกับ{" "}
                  <span style={{ color: "#059669", fontWeight: 600 }}>
                    {dateB}
                  </span>
                </>
              )}
            </p>
          </div>

          {/* KPI Cards */}
          <div className="kpi-grid">
            <KpiCard
              label="ยอดขายรวม (หลัก)"
              icon="💰"
              value={`฿${totalA.toLocaleString()}`}
              sub={dateA || "—"}
              color="#2563eb"
            />
            <KpiCard
              label="ยอดขายรวม (เทียบ)"
              icon="📋"
              value={`฿${totalB.toLocaleString()}`}
              sub={dateB || "ยังไม่ได้เลือก"}
              color="#10b981"
            />
            <KpiCard
              label="ส่วนต่าง"
              icon={diff >= 0 ? "📈" : "📉"}
              value={pct && dateB ? `${diff >= 0 ? "+" : ""}${pct}%` : "—"}
              sub={
                dateB
                  ? `฿${Math.abs(diff).toLocaleString()} บาท`
                  : "เลือกรอบเทียบ"
              }
              color={diff >= 0 ? "#10b981" : "#ef4444"}
            />
            <KpiCard
              label="สาขาที่ขายดีสุด"
              icon="🏆"
              value={top?.branch || "—"}
              sub={top ? `฿${top.a_total.toLocaleString()}` : "—"}
              color="#f59e0b"
            />
          </div>

          {/* Main chart card */}
          <div
            style={{
              background: "#fff",
              borderRadius: 14,
              border: "1px solid #e5e7eb",
              boxShadow: "0 2px 8px rgba(0,0,0,0.05)",
              overflow: "hidden",
            }}
          >
            {/* Scrollable tab bar */}
            <div
              className="tab-bar"
              style={{ borderBottom: "1px solid #f3f4f6" }}
            >
              {tabs.map((t) => (
                <button
                  key={t.id}
                  className="tab-btn"
                  onClick={() => setActiveTab(t.id)}
                  style={{
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    padding: "14px 18px",
                    fontSize: 13,
                    fontWeight: activeTab === t.id ? 700 : 500,
                    color: activeTab === t.id ? "#2563eb" : "#6b7280",
                    borderBottom:
                      activeTab === t.id
                        ? "2px solid #2563eb"
                        : "2px solid transparent",
                    marginBottom: -1,
                    transition: "all 0.15s",
                  }}
                >
                  {t.label}
                </button>
              ))}
            </div>

            <div style={{ padding: "20px 20px 28px" }}>
              {/* ── TAB: เปรียบเทียบ ── */}
              {activeTab === "compare" && (
                <>
                  <div style={{ height: 280, marginBottom: 24 }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={filtered} barGap={4} barCategoryGap="30%">
                        <CartesianGrid
                          strokeDasharray="3 3"
                          stroke="#f3f4f6"
                          vertical={false}
                        />
                        <XAxis
                          dataKey="short"
                          tick={{ fill: "#6b7280", fontSize: 11 }}
                          axisLine={false}
                          tickLine={false}
                        />
                        <YAxis
                          tick={{ fill: "#9ca3af", fontSize: 11 }}
                          axisLine={false}
                          tickLine={false}
                          tickFormatter={(v) =>
                            v >= 1000 ? `${(v / 1000).toFixed(0)}K` : String(v)
                          }
                          width={40}
                        />
                        <Tooltip
                          content={<CustomTooltip />}
                          cursor={{ fill: "#f9fafb" }}
                        />
                        <Legend
                          wrapperStyle={{
                            fontSize: 12,
                            paddingTop: 12,
                            color: "#6b7280",
                          }}
                        />
                        {dateB && (
                          <Bar
                            dataKey="b_total"
                            name={`${dateB}`}
                            fill="#d1fae5"
                            stroke="#10b981"
                            strokeWidth={1.5}
                            radius={[5, 5, 0, 0]}
                          />
                        )}
                        <Bar
                          dataKey="a_total"
                          name={`${dateA}`}
                          fill="#2563eb"
                          radius={[5, 5, 0, 0]}
                        />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>

                  {/* Scrollable table on mobile */}
                  <div className="table-wrap">
                    <table
                      className="data-table"
                      style={{
                        width: "100%",
                        borderCollapse: "separate",
                        borderSpacing: "0 3px",
                        fontSize: 13,
                      }}
                    >
                      <thead>
                        <tr
                          style={{
                            fontSize: 10,
                            color: "#9ca3af",
                            textTransform: "uppercase",
                            letterSpacing: "0.07em",
                          }}
                        >
                          <th
                            style={{
                              padding: "8px 12px",
                              textAlign: "left",
                              fontWeight: 600,
                            }}
                          >
                            สาขา
                          </th>
                          <th
                            style={{
                              padding: "8px 12px",
                              textAlign: "right",
                              fontWeight: 600,
                            }}
                          >
                            รอบเทียบ
                          </th>
                          <th
                            style={{
                              padding: "8px 12px",
                              textAlign: "right",
                              fontWeight: 600,
                              color: "#2563eb",
                            }}
                          >
                            รอบหลัก
                          </th>
                          <th
                            style={{
                              padding: "8px 12px",
                              textAlign: "right",
                              fontWeight: 600,
                            }}
                          >
                            ส่วนต่าง
                          </th>
                          <th
                            style={{
                              padding: "8px 12px",
                              textAlign: "right",
                              fontWeight: 600,
                            }}
                          >
                            เติบโต
                          </th>
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
                          return (
                            <tr
                              key={i}
                              className="row-hover"
                              style={{
                                background: "#fff",
                                transition: "background 0.12s",
                              }}
                            >
                              <td
                                style={{
                                  padding: "11px 12px",
                                  borderRadius: "7px 0 0 7px",
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
                                      color: "#111827",
                                    }}
                                  >
                                    {item.branch}
                                  </span>
                                </div>
                              </td>
                              <td
                                className="num"
                                style={{
                                  padding: "11px 12px",
                                  textAlign: "right",
                                  color: "#9ca3af",
                                }}
                              >
                                {item.b_total > 0
                                  ? `฿${item.b_total.toLocaleString()}`
                                  : "—"}
                              </td>
                              <td
                                className="num"
                                style={{
                                  padding: "11px 12px",
                                  textAlign: "right",
                                  color: "#2563eb",
                                  fontWeight: 700,
                                }}
                              >
                                ฿{item.a_total.toLocaleString()}
                              </td>
                              <td
                                className="num"
                                style={{
                                  padding: "11px 12px",
                                  textAlign: "right",
                                  fontWeight: 600,
                                  color: pos ? "#059669" : "#dc2626",
                                }}
                              >
                                {dateB
                                  ? `${pos ? "+" : ""}฿${d.toLocaleString()}`
                                  : "—"}
                              </td>
                              <td
                                style={{
                                  padding: "11px 12px",
                                  textAlign: "right",
                                  borderRadius: "0 7px 7px 0",
                                }}
                              >
                                {p && dateB ? (
                                  <span
                                    className="num"
                                    style={{
                                      display: "inline-flex",
                                      alignItems: "center",
                                      gap: 2,
                                      background: pos ? "#f0fdf4" : "#fef2f2",
                                      color: pos ? "#15803d" : "#b91c1c",
                                      padding: "3px 9px",
                                      borderRadius: 20,
                                      fontSize: 11,
                                      fontWeight: 700,
                                    }}
                                  >
                                    {pos ? "▲" : "▼"} {Math.abs(Number(p))}%
                                  </span>
                                ) : (
                                  "—"
                                )}
                              </td>
                            </tr>
                          );
                        })}
                        {filtered.length > 1 && (
                          <tr style={{ background: "#f8fafc" }}>
                            <td
                              style={{
                                padding: "11px 12px",
                                borderRadius: "7px 0 0 7px",
                                fontWeight: 700,
                                fontSize: 12,
                                color: "#374151",
                              }}
                            >
                              รวมทั้งหมด
                            </td>
                            <td
                              className="num"
                              style={{
                                padding: "11px 12px",
                                textAlign: "right",
                                color: "#6b7280",
                                fontWeight: 600,
                              }}
                            >
                              {totalB > 0 ? `฿${totalB.toLocaleString()}` : "—"}
                            </td>
                            <td
                              className="num"
                              style={{
                                padding: "11px 12px",
                                textAlign: "right",
                                color: "#2563eb",
                                fontWeight: 700,
                              }}
                            >
                              ฿{totalA.toLocaleString()}
                            </td>
                            <td
                              className="num"
                              style={{
                                padding: "11px 12px",
                                textAlign: "right",
                                fontWeight: 700,
                                color: diff >= 0 ? "#059669" : "#dc2626",
                              }}
                            >
                              {dateB
                                ? `${diff >= 0 ? "+" : ""}฿${diff.toLocaleString()}`
                                : "—"}
                            </td>
                            <td
                              style={{
                                padding: "11px 12px",
                                textAlign: "right",
                                borderRadius: "0 7px 7px 0",
                              }}
                            >
                              {pct && dateB ? (
                                <span
                                  className="num"
                                  style={{
                                    display: "inline-flex",
                                    alignItems: "center",
                                    gap: 2,
                                    background:
                                      diff >= 0 ? "#f0fdf4" : "#fef2f2",
                                    color: diff >= 0 ? "#15803d" : "#b91c1c",
                                    padding: "3px 9px",
                                    borderRadius: 20,
                                    fontSize: 11,
                                    fontWeight: 700,
                                  }}
                                >
                                  {diff >= 0 ? "▲" : "▼"}{" "}
                                  {Math.abs(Number(pct))}%
                                </span>
                              ) : (
                                "—"
                              )}
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
                <div className="chart-product-layout">
                  <div className="chart-product-main">
                    <div
                      style={{
                        fontWeight: 600,
                        color: "#374151",
                        marginBottom: 16,
                        fontSize: 14,
                      }}
                    >
                      ยอดแยกตามประเภทสินค้า — {dateA}
                    </div>
                    <div style={{ height: 280 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={filtered}
                          barGap={3}
                          barCategoryGap="25%"
                        >
                          <CartesianGrid
                            strokeDasharray="3 3"
                            stroke="#f3f4f6"
                            vertical={false}
                          />
                          <XAxis
                            dataKey="short"
                            tick={{ fill: "#6b7280", fontSize: 11 }}
                            axisLine={false}
                            tickLine={false}
                          />
                          <YAxis
                            tick={{ fill: "#9ca3af", fontSize: 11 }}
                            axisLine={false}
                            tickLine={false}
                            tickFormatter={(v) =>
                              v >= 1000
                                ? `${(v / 1000).toFixed(0)}K`
                                : String(v)
                            }
                            width={40}
                          />
                          <Tooltip
                            content={<CustomTooltip />}
                            cursor={{ fill: "#f9fafb" }}
                          />
                          <Legend
                            wrapperStyle={{ fontSize: 12, paddingTop: 12 }}
                          />
                          <Bar
                            dataKey="a_tuangok"
                            name="ถั่วงอก"
                            fill={PROD_COLORS[0]}
                            radius={[4, 4, 0, 0]}
                          />
                          <Bar
                            dataKey="a_khaomak"
                            name="ข้าวหมาก"
                            fill={PROD_COLORS[1]}
                            radius={[4, 4, 0, 0]}
                          />
                          <Bar
                            dataKey="a_tonon"
                            name="ต้นอ่อน"
                            fill={PROD_COLORS[2]}
                            radius={[4, 4, 0, 0]}
                          />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                  <div className="chart-product-side">
                    <div
                      style={{
                        fontWeight: 600,
                        color: "#374151",
                        marginBottom: 16,
                        fontSize: 14,
                      }}
                    >
                      สัดส่วนสินค้า
                    </div>
                    <div style={{ height: 180 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={pieData}
                            cx="50%"
                            cy="50%"
                            innerRadius={45}
                            outerRadius={72}
                            paddingAngle={3}
                            dataKey="value"
                            startAngle={90}
                            endAngle={-270}
                          >
                            {pieData.map((_, i) => (
                              <Cell key={i} fill={PROD_COLORS[i]} />
                            ))}
                          </Pie>
                          <Tooltip
                            formatter={(v: any) => [
                              `฿${Number(v).toLocaleString()}`,
                              "",
                            ]}
                            contentStyle={{
                              background: "#fff",
                              border: "1px solid #e5e7eb",
                              borderRadius: 8,
                              fontSize: 12,
                            }}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                    <div style={{ marginTop: 8 }}>
                      {pieData.map((item, i) => {
                        const total = pieData.reduce((s, d) => s + d.value, 0);
                        const share =
                          total > 0
                            ? ((item.value / total) * 100).toFixed(0)
                            : 0;
                        return (
                          <div
                            key={i}
                            style={{
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "space-between",
                              padding: "9px 0",
                              borderBottom: "1px solid #f3f4f6",
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
                                  width: 9,
                                  height: 9,
                                  borderRadius: 2,
                                  background: PROD_COLORS[i],
                                }}
                              />
                              <span style={{ color: "#374151", fontSize: 13 }}>
                                {item.name}
                              </span>
                            </div>
                            <div style={{ textAlign: "right" }}>
                              <div
                                className="num"
                                style={{
                                  fontWeight: 700,
                                  color: "#111827",
                                  fontSize: 13,
                                }}
                              >
                                ฿{item.value.toLocaleString()}
                              </div>
                              <div
                                className="num"
                                style={{ fontSize: 10, color: "#9ca3af" }}
                              >
                                {share}%
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </div>
              )}

              {/* ── TAB: แนวโน้ม ── */}
              {activeTab === "trend" && (
                <>
                  <div
                    style={{
                      fontWeight: 600,
                      color: "#374151",
                      marginBottom: 16,
                      fontSize: 14,
                    }}
                  >
                    แนวโน้มยอดขายรวมทุกรอบ
                    {selectedBranch !== "ALL" && (
                      <span style={{ color: "#2563eb" }}>
                        {" "}
                        · {selectedBranch}
                      </span>
                    )}
                  </div>
                  <div style={{ height: 280 }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={trendData}>
                        <defs>
                          <linearGradient
                            id="blueGrad"
                            x1="0"
                            y1="0"
                            x2="0"
                            y2="1"
                          >
                            <stop
                              offset="5%"
                              stopColor="#2563eb"
                              stopOpacity={0.12}
                            />
                            <stop
                              offset="95%"
                              stopColor="#2563eb"
                              stopOpacity={0}
                            />
                          </linearGradient>
                        </defs>
                        <CartesianGrid
                          strokeDasharray="3 3"
                          stroke="#f3f4f6"
                          vertical={false}
                        />
                        <XAxis
                          dataKey="date"
                          tick={{ fill: "#6b7280", fontSize: 11 }}
                          axisLine={false}
                          tickLine={false}
                        />
                        <YAxis
                          tick={{ fill: "#9ca3af", fontSize: 11 }}
                          axisLine={false}
                          tickLine={false}
                          tickFormatter={(v) =>
                            v >= 1000 ? `${(v / 1000).toFixed(0)}K` : String(v)
                          }
                          width={40}
                        />
                        <Tooltip
                          content={<CustomTooltip />}
                          cursor={{
                            stroke: "#2563eb",
                            strokeWidth: 1,
                            strokeDasharray: "4 4",
                          }}
                        />
                        <Area
                          type="monotone"
                          dataKey="total"
                          name="ยอดรวม"
                          stroke="#2563eb"
                          strokeWidth={2.5}
                          fill="url(#blueGrad)"
                          dot={{
                            fill: "#2563eb",
                            r: 4,
                            strokeWidth: 2,
                            stroke: "#fff",
                          }}
                          activeDot={{
                            r: 6,
                            stroke: "#2563eb",
                            strokeWidth: 2,
                            fill: "#fff",
                          }}
                        />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                  <div className="trend-cards">
                    {trendData.map((d, i) => {
                      const maxVal = Math.max(...trendData.map((t) => t.total));
                      const isTop = d.total === maxVal && maxVal > 0;
                      return (
                        <div
                          key={i}
                          className="trend-card"
                          style={{
                            background: isTop ? "#eff6ff" : "#f9fafb",
                            border: `1px solid ${isTop ? "#bfdbfe" : "#e5e7eb"}`,
                            borderTop: `3px solid ${isTop ? "#2563eb" : "#e5e7eb"}`,
                            borderRadius: 10,
                            padding: "11px 14px",
                          }}
                        >
                          <div
                            style={{
                              fontSize: 10,
                              color: "#9ca3af",
                              marginBottom: 4,
                            }}
                          >
                            {d.date}
                          </div>
                          <div
                            className="num"
                            style={{
                              fontSize: 16,
                              fontWeight: 700,
                              color: isTop ? "#2563eb" : "#111827",
                            }}
                          >
                            ฿{d.total.toLocaleString()}
                          </div>
                          {isTop && (
                            <div
                              style={{
                                fontSize: 10,
                                color: "#2563eb",
                                marginTop: 3,
                                fontWeight: 600,
                              }}
                            >
                              🏆 สูงสุด
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

          <div
            style={{
              textAlign: "center",
              marginTop: 28,
              fontSize: 11,
              color: "#d1d5db",
            }}
          >
            Revenue Dashboard ·{" "}
            {new Date().toLocaleDateString("th-TH", {
              year: "numeric",
              month: "long",
              day: "numeric",
            })}
          </div>
        </main>
      </div>
    </>
  );
}

"use client";
import { useState, useMemo } from "react";
import jsPDF from "jspdf";
import html2canvas from "html2canvas-pro";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Legend,
} from "recharts";

const COLORS = ["#3B82F6", "#10B981", "#F59E0B", "#EF4444", "#8B5CF6", "#EC4899", "#14B8A6", "#F97316"];
const MAX_TABLE_ROWS = 10;
const MAX_CHART_GROUPS = 20; // beyond this, bucket the rest as "Other"

export default function CsvUploader() {
  const [rows, setRows] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [search, setSearch] = useState("");
  const [darkMode, setDarkMode] = useState(false);
  const [aiInsight, setAiInsight] = useState("");
  const [loadingAI, setLoadingAI] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [fileName, setFileName] = useState("");

  const [numericCols, setNumericCols] = useState<string[]>([]);
  const [categoricalCols, setCategoricalCols] = useState<string[]>([]);

  // User-overridable selections (default to auto-detected first columns)
  const [selectedMetric, setSelectedMetric] = useState<string>("");
  const [selectedGroup, setSelectedGroup] = useState<string>("");
  const [selectedLabel, setSelectedLabel] = useState<string>(""); // per-row label column, e.g. "Name"

  // ---------- Column type detection ----------
  const detectColumnTypes = (data: any[]) => {
    if (data.length === 0) return { numeric: [], categorical: [] };

    const cols = Object.keys(data[0]);
    const numeric: string[] = [];
    const categorical: string[] = [];

    cols.forEach((col) => {
      const sample = data.slice(0, 20).map((r) => r[col]);
      const numericCount = sample.filter(
        (v) => v !== "" && v !== null && v !== undefined && !isNaN(Number(v))
      ).length;

      if (numericCount / sample.length > 0.8) {
        numeric.push(col);
      } else {
        categorical.push(col);
      }
    });

    return { numeric, categorical };
  };

  // ---------- File upload (handles CSV, XLS, XLSX; big files won't freeze the tab) ----------
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setUploading(true);
    setFileName(file.name);

    const reader = new FileReader();

    reader.onload = (e) => {
      // yield to the browser before the heavy parse so the UI can repaint
      setTimeout(() => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: "array", dense: true });

          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData: any[] = XLSX.utils.sheet_to_json(sheet, { defval: "" });

          setRows(jsonData);

          if (jsonData.length > 0) {
            const cols = Object.keys(jsonData[0]);
            setColumns(cols);

            const { numeric, categorical } = detectColumnTypes(jsonData);
            setNumericCols(numeric);
            setCategoricalCols(categorical);

            // set sensible defaults, but let the user change them via dropdowns
            setSelectedMetric(numeric[0] || "");
            setSelectedGroup(categorical[0] || "");
            setSelectedLabel(categorical[0] || cols[0] || "");
          } else {
            setColumns([]);
            setNumericCols([]);
            setCategoricalCols([]);
          }
        } catch (err) {
          console.error("File parse error:", err);
          alert("Couldn't read that file. Make sure it's a valid CSV/XLS/XLSX.");
        } finally {
          setUploading(false);
        }
      }, 0);
    };

    reader.onerror = () => {
      setUploading(false);
      alert("Failed to read the file.");
    };

    reader.readAsArrayBuffer(file);
  };

  // ---------- Derived stats (all generic — driven by selectedMetric/selectedGroup) ----------
  const metricValues = useMemo(
    () =>
      selectedMetric
        ? rows.map((row) => Number(row[selectedMetric])).filter((v) => !isNaN(v))
        : [],
    [rows, selectedMetric]
  );

  const totalMetric = metricValues.reduce((a, b) => a + b, 0);
  const avgMetric = metricValues.length ? Number((totalMetric / metricValues.length).toFixed(2)) : 0;
  const maxMetric = metricValues.length ? Math.max(...metricValues) : 0;
  const minMetric = metricValues.length ? Math.min(...metricValues) : 0;

  const topRow = useMemo(() => {
    if (!selectedMetric || rows.length === 0) return null;
    return rows.reduce((prev, current) =>
      Number(prev[selectedMetric]) > Number(current[selectedMetric]) ? prev : current
    );
  }, [rows, selectedMetric]);

  const filteredRows = rows.filter((row) =>
    Object.values(row).join(" ").toLowerCase().includes(search.toLowerCase())
  );

  const chartData = useMemo(() => {
    if (!selectedLabel || !selectedMetric) return [];
    return rows.map((row) => ({
      name: String(row[selectedLabel] ?? ""),
      value: Number(row[selectedMetric]) || 0,
    }));
  }, [rows, selectedLabel, selectedMetric]);

  const groupedData: { name: string; value: number }[] = useMemo(() => {
    if (!selectedGroup || !selectedMetric) return [];

    const totals: Record<string, number> = {};
    rows.forEach((row) => {
      const key = String(row[selectedGroup] ?? "Unknown");
      totals[key] = (totals[key] || 0) + (Number(row[selectedMetric]) || 0);
    });

    const sorted = Object.entries(totals)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    if (sorted.length <= MAX_CHART_GROUPS) return sorted;

    const top = sorted.slice(0, MAX_CHART_GROUPS);
    const otherTotal = sorted.slice(MAX_CHART_GROUPS).reduce((sum, g) => sum + g.value, 0);
    return [...top, { name: "Other", value: otherTotal }];
  }, [rows, selectedGroup, selectedMetric]);

  const bestGroup = groupedData[0]?.name;

  // Cap employee-style bar chart to top N too, so huge files don't render thousands of bars
  const barChartData = useMemo(() => {
    const sorted = [...chartData].sort((a, b) => b.value - a.value);
    return sorted.slice(0, MAX_CHART_GROUPS);
  }, [chartData]);

  const downloadReport = () => {
    if (columns.length === 0) return;

    const header = columns.join(",");
    const body = rows
      .map((row) => columns.map((col) => `"${String(row[col] ?? "").replace(/"/g, '""')}"`).join(","))
      .join("\n");

    const blob = new Blob([`${header}\n${body}`], { type: "text/csv;charset=utf-8;" });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "data-report.csv";
    link.click();
    window.URL.revokeObjectURL(url);
  };

  const generateInsights = async () => {
    setLoadingAI(true);
    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rows: rows.slice(0, 500) }), // cap payload for huge files
      });

      const data = await response.json();
      setAiInsight(data.insight);
    } catch (err) {
      console.error("AI insight error:", err);
      setAiInsight("Couldn't generate insights right now.");
    } finally {
      setLoadingAI(false);
    }
  };

  const totalCells = rows.length * columns.length;
  const missingValues = rows.reduce(
    (count, row) =>
      count + Object.values(row).filter((value) => value === "" || value === null || value === undefined).length,
    0
  );
  const dataHealth = totalCells > 0 ? (((totalCells - missingValues) / totalCells) * 100).toFixed(1) : "100";

  const exportPDF = async () => {
    try {
      const dashboard = document.getElementById("dashboard-report");
      if (!dashboard) return;

      const canvas = await html2canvas(dashboard, {
        scale: 2,
        useCORS: true,
        backgroundColor: "#ffffff",
        logging: false,
      });

      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF("p", "mm", "a4");
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

      pdf.addImage(imgData, "PNG", 0, 0, pdfWidth, pdfHeight);
      pdf.save("AI_Data_Report.pdf");
    } catch (err) {
      console.error("PDF Error:", err);
      alert(String(err));
    }
  };

  return (
    <div
      id="dashboard-report"
      className={`rounded-2xl shadow-xl p-8 ${darkMode ? "bg-gray-900 text-white" : "bg-white"}`}
    >
      <div className="bg-white rounded-2xl shadow-xl p-8">
        <h2 className="text-3xl font-bold mb-6">Upload Dataset</h2>

        <div className="flex items-center gap-4 mb-6 flex-wrap">
          <input type="file" accept=".csv,.xlsx,.xls" onChange={handleFileUpload} />

          {uploading && <span className="text-sm text-gray-500">Processing {fileName}...</span>}

          <div className="flex gap-4">
            <button onClick={downloadReport} className="bg-green-600 text-white px-4 py-2 rounded-lg">
              Download Report
            </button>

            <button onClick={exportPDF} className="bg-red-600 text-white px-4 py-2 rounded-lg">
              Export PDF
            </button>

            <button onClick={() => setDarkMode(!darkMode)} className="bg-black text-white px-4 py-2 rounded-lg">
              Theme
            </button>
          </div>
        </div>

        {rows.length > 0 && (
          <>
            {/* Column selectors — lets the user correct auto-detection for any file shape */}
            <div className="flex gap-4 flex-wrap mb-8 text-sm">
              <label className="flex items-center gap-2">
                Value column:
                <select
                  value={selectedMetric}
                  onChange={(e) => setSelectedMetric(e.target.value)}
                  className="border rounded px-2 py-1"
                >
                  {numericCols.map((c) => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </label>

              <label className="flex items-center gap-2">
                Group by:
                <select
                  value={selectedGroup}
                  onChange={(e) => setSelectedGroup(e.target.value)}
                  className="border rounded px-2 py-1"
                >
                  {categoricalCols.map((c) => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </label>

              <label className="flex items-center gap-2">
                Row label:
                <select
                  value={selectedLabel}
                  onChange={(e) => setSelectedLabel(e.target.value)}
                  className="border rounded px-2 py-1"
                >
                  {columns.map((c) => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </label>

              <span className="text-gray-500">{rows.length.toLocaleString()} rows loaded</span>
            </div>

            <h3 className="text-2xl font-bold mb-4">Dashboard Overview</h3>

            {/* Executive KPIs */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 mb-8">
              <div className="bg-gradient-to-r from-green-500 to-green-700 text-white p-6 rounded-xl">
                <h4 className="text-lg font-semibold">Top {selectedLabel || "Row"}</h4>
                <p className="text-3xl font-bold mt-2">{topRow?.[selectedLabel] ?? "-"}</p>
                <p className="mt-2">
                  {selectedMetric}: {topRow?.[selectedMetric] ?? "-"}
                </p>
              </div>

              <div className="bg-gradient-to-r from-blue-500 to-blue-700 text-white p-6 rounded-xl">
                <h4 className="text-lg font-semibold">Top {selectedGroup || "Group"}</h4>
                <p className="text-3xl font-bold mt-2">{bestGroup ?? "-"}</p>
              </div>

              <div className="bg-gradient-to-r from-purple-500 to-purple-700 text-white p-6 rounded-xl">
                <h4 className="text-lg font-semibold">Data Health</h4>
                <p className="text-3xl font-bold mt-2">{dataHealth}%</p>
                <p>{missingValues === 0 ? "No Missing Values" : `${missingValues} Missing Values`}</p>
              </div>
            </div>

            {/* Search */}
            <input
              type="text"
              placeholder="Search dataset..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full p-3 border rounded-lg mb-6"
            />

            {/* Table — always capped, even on huge files */}
            <div className="bg-white rounded-xl shadow p-6 mb-8">
              <h3 className="text-xl font-bold mb-4">
                Data Preview (showing {Math.min(MAX_TABLE_ROWS, filteredRows.length)} of{" "}
                {filteredRows.length.toLocaleString()})
              </h3>

              <div className="overflow-x-auto">
                <table className="w-full border-collapse border border-gray-300">
                  <thead>
                    <tr>
                      {columns.map((col) => (
                        <th key={col} className="px-4 py-3 bg-gray-100 text-left font-semibold">
                          {col}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.slice(0, MAX_TABLE_ROWS).map((row, index) => (
                      <tr key={index}>
                        {columns.map((col) => (
                          <td key={col} className="px-4 py-3 border-b">
                            {String(row[col])}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* AI Insights */}
            <div className="bg-gradient-to-r from-blue-600 to-purple-600 text-white p-6 rounded-xl mt-8">
              <h3 className="text-2xl font-bold mb-4">AI Insights</h3>

              <ul className="space-y-2 mb-4">
                <li> Total {selectedMetric}: {totalMetric.toLocaleString()}</li>
                <li> Average {selectedMetric}: {avgMetric.toLocaleString()}</li>
                <li> Highest {selectedMetric}: {maxMetric.toLocaleString()}</li>
                <li> Lowest {selectedMetric}: {minMetric.toLocaleString()}</li>
                <li> Records Analyzed: {rows.length.toLocaleString()}</li>
              </ul>

              <button onClick={generateInsights} className="bg-white text-black px-4 py-2 rounded-lg mb-4">
                {loadingAI ? "Analyzing..." : "Generate AI Insights"}
              </button>

              {aiInsight && (
                <div className="bg-white text-black p-4 rounded-xl mb-4">
                  <h4 className="font-bold mb-2">AI Analysis</h4>
                  <p>{aiInsight}</p>
                </div>
              )}

              <div className="bg-black/20 p-4 rounded-xl">
                <h4 className="font-bold text-lg mb-2">Recommendations</h4>
                <ul className="space-y-2">
                  <li> Top value: {maxMetric.toLocaleString()} ({selectedMetric})</li>
                  <li> Average value: {avgMetric.toLocaleString()}</li>
                  <li> Focus on {selectedGroup} groups performing below average</li>
                </ul>
              </div>
            </div>

            {/* Charts — capped to top 20 groups/rows, rest bucketed as "Other" */}
            <div className="grid md:grid-cols-2 gap-6">
              <div className="bg-white rounded-xl shadow p-6">
                <h3 className="text-xl font-bold mb-4">
                  {selectedMetric} by {selectedLabel} {chartData.length > MAX_CHART_GROUPS ? "(top 20)" : ""}
                </h3>
                <ResponsiveContainer width="100%" height={350}>
                  <BarChart data={barChartData}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="name" />
                    <YAxis />
                    <Tooltip />
                    <Bar dataKey="value" fill="#3B82F6" radius={[8, 8, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="bg-white rounded-xl shadow p-6">
                <h3 className="text-xl font-bold mb-4">
                  {selectedMetric} by {selectedGroup} {groupedData.length > MAX_CHART_GROUPS ? "(top 20 + Other)" : ""}
                </h3>
                <ResponsiveContainer width="100%" height={350}>
                  <PieChart>
                    <Pie data={groupedData} dataKey="value" nameKey="name" outerRadius={120} label>
                      {groupedData.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                      ))}
                    </Pie>
                    <Tooltip />
                    <Legend />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

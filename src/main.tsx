import React from "react";
import ReactDOM from "react-dom/client";
import ReactECharts from "echarts-for-react";
import "./styles.css";

const sampleChartOption = {
  tooltip: {
    trigger: "axis"
  },
  grid: {
    left: 40,
    right: 24,
    top: 36,
    bottom: 36
  },
  xAxis: {
    type: "category",
    data: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
  },
  yAxis: {
    type: "value"
  },
  series: [
    {
      name: "Sales",
      type: "bar",
      data: [120, 180, 150, 240, 210, 280],
      itemStyle: {
        color: "#2563eb"
      }
    }
  ]
};

function App() {
  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div>
          <h1>Excel Visual Tool</h1>
          <p>Local workbook visualization workspace.</p>
        </div>
        <button type="button" className="primary-button">
          Open .xlsx
        </button>
        <section className="panel">
          <h2>Worksheets</h2>
          <div className="empty-state">No workbook selected</div>
        </section>
      </aside>

      <section className="workspace">
        <div className="preview-area">
          <div className="section-header">
            <h2>Data Preview</h2>
            <span>First 200 rows</span>
          </div>
          <div className="table-placeholder">
            <span>Select an Excel file to preview worksheet rows.</span>
          </div>
        </div>

        <div className="chart-area">
          <div className="section-header">
            <h2>Chart Preview</h2>
            <span>ECharts ready</span>
          </div>
          <ReactECharts option={sampleChartOption} className="chart" />
        </div>
      </section>

      <aside className="settings">
        <h2>Chart Settings</h2>
        <label>
          Chart type
          <select defaultValue="bar">
            <option value="bar">Bar</option>
            <option value="line">Line</option>
            <option value="pie">Pie</option>
            <option value="scatter">Scatter</option>
          </select>
        </label>
        <label>
          X axis
          <select disabled>
            <option>Select field</option>
          </select>
        </label>
        <label>
          Y axis
          <select disabled>
            <option>Select field</option>
          </select>
        </label>
      </aside>
    </main>
  );
}

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

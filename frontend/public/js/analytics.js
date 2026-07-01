let latestAnalytics = null;

const CHART_COLORS = [
  "#1f5fbf",
  "#0f766e",
  "#b45309",
  "#7c3aed",
  "#be123c",
  "#0369a1",
  "#15803d",
  "#a16207",
];

function drawPieChart(canvasId, values, labels, colors = []) {
  const canvas = document.getElementById(canvasId);
  const ctx = canvas.getContext("2d");
  const total = values.reduce((a, b) => a + b, 0);

  ctx.clearRect(0, 0, canvas.width, canvas.height);

  if (!total) {
    ctx.fillStyle = "#64748b";
    ctx.font = "14px Segoe UI";
    ctx.fillText("No contribution data available", 20, 40);
    return;
  }

  const cx = 150;
  const cy = 110;
  const radius = 80;

  let start = -Math.PI / 2;
  values.forEach((v, i) => {
    const slice = (v / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, radius, start, start + slice);
    ctx.closePath();
    ctx.fillStyle = colors[i] || CHART_COLORS[i % CHART_COLORS.length];
    ctx.fill();
    start += slice;
  });

  const legend = document.getElementById("pieLegend");
  legend.innerHTML = labels
    .map((label, i) => {
      const percent = ((values[i] / total) * 100).toFixed(1);
      const color = colors[i] || CHART_COLORS[i % CHART_COLORS.length];
      return `<span class="badge" style="background:${color}20; color:${color};">${label}: ${percent}%</span>`;
    })
    .join("");
}

function drawNetContributionPie(data) {
  const perUser = (data.currentDocument?.per_user || [])
    .map((x) => ({
      name: x.name,
      words: Number(x.words || 0),
      color: x.color || null,
    }))
    .filter((x) => x.words > 0);

  drawPieChart(
    "contributionPie",
    perUser.map((x) => x.words),
    perUser.map((x) => x.name),
    perUser.map((x, i) => x.color || CHART_COLORS[i % CHART_COLORS.length])
  );
}

async function downloadAnalyticsPdf() {
  const projectId = Number(document.getElementById("projectIdInput").value) || "project";
  const capture = document.getElementById("analyticsCapture");
  const canvas = await window.html2canvas(capture, {
    scale: 2,
    backgroundColor: "#ffffff",
    useCORS: true,
  });

  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF("p", "mm", "a4");
  const pageWidth = 210;
  const pageHeight = 297;
  const imageWidth = pageWidth;
  const imageHeight = (canvas.height * imageWidth) / canvas.width;
  const pageCanvasHeight = (canvas.width * pageHeight) / pageWidth;
  const totalPages = Math.max(1, Math.ceil(canvas.height / pageCanvasHeight));

  for (let pageIndex = 0; pageIndex < totalPages; pageIndex += 1) {
    const pageCanvas = document.createElement("canvas");
    pageCanvas.width = canvas.width;
    pageCanvas.height = Math.min(pageCanvasHeight, canvas.height - pageIndex * pageCanvasHeight);

    const ctx = pageCanvas.getContext("2d");
    ctx.drawImage(
      canvas,
      0,
      pageIndex * pageCanvasHeight,
      canvas.width,
      pageCanvas.height,
      0,
      0,
      canvas.width,
      pageCanvas.height
    );

    const imageData = pageCanvas.toDataURL("image/png");
    if (pageIndex > 0) {
      pdf.addPage();
    }
    const sliceHeightMm = (pageCanvas.height * imageWidth) / canvas.width;
    pdf.addImage(imageData, "PNG", 0, 0, imageWidth, sliceHeightMm);
  }

  pdf.save(`analytics_${projectId}.pdf`);
}

async function loadAnalytics(projectId) {
  const data = await API.get(`/api/analytics/${projectId}/summary`);
  latestAnalytics = data;

  document.getElementById("summaryCards").innerHTML = `
    <div class="card"><h4>Current Doc Words</h4><p>${data.currentDocument?.total_words || 0}</p></div>
    <div class="card"><h4>Active Contributors</h4><p>${(data.currentStateSummary || []).filter((x) => Number(x.net_words || 0) > 0).length}</p></div>
  `;

  const summaryBody = document.querySelector("#summaryTable tbody");
  summaryBody.innerHTML = "";
  (data.currentStateSummary || []).forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.name}</td>
      <td>${row.role || "-"}</td>
      <td>${row.net_words || 0}</td>
      <td>${row.net_percent || 0}%</td>
    `;
    summaryBody.appendChild(tr);
  });

  const historyBody = document.querySelector("#historyLogTable tbody");
  historyBody.innerHTML = "";
  (data.fullHistory?.logs || []).forEach((row) => {
    const time = row.timestamp ? new Date(row.timestamp).toLocaleString() : "-";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${time}</td>
      <td>${row.name}</td>
      <td>${row.action_type || "-"}</td>
      <td>${row.page_number || "-"}</td>
      <td>${row.word_count || 0}</td>
    `;
    historyBody.appendChild(tr);
  });

  drawNetContributionPie(data);
}

(async function init() {
  const user = await getCurrentUser();
  if (!user) {
    window.location.href = "/login";
    return;
  }

  document.getElementById("loadAnalyticsBtn").addEventListener("click", async () => {
    const projectId = Number(document.getElementById("projectIdInput").value);
    if (!projectId) return;
    await loadAnalytics(projectId);
  });

  document.getElementById("downloadPdfBtn").addEventListener("click", async () => {
    if (!latestAnalytics) return;
    await downloadAnalyticsPdf();
  });

  document.getElementById("toggleHistoryBtn").addEventListener("click", () => {
    const section = document.getElementById("historySection");
    const btn = document.getElementById("toggleHistoryBtn");
    if (!section) return;
    const open = section.style.display !== "none";
    section.style.display = open ? "none" : "block";
    btn.textContent = open ? "Show Full History" : "Hide Full History";
  });

  const historySection = document.getElementById("historySection");
  const toggleHistoryBtn = document.getElementById("toggleHistoryBtn");
  if (historySection && toggleHistoryBtn) {
    historySection.style.display = "none";
    toggleHistoryBtn.textContent = "Show Full History";
  }

  const initialProjectId = getQuery("projectId");
  if (initialProjectId) {
    document.getElementById("projectIdInput").value = initialProjectId;
    await loadAnalytics(Number(initialProjectId));
  }
})();

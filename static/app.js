const state = {
  scene: "eso",
  result: null,
  busy: false,
};

const sceneMeta = {
  eso: {
    title: "ESO未完成清单",
    primaryLabel: "零件俱乐部主清单",
    archiveLabel: "ESO归档清单（可选，用于回填）",
  },
  drawing: {
    title: "图纸/数模未完成清单",
    primaryLabel: "图纸/数模主清单",
    archiveLabel: "图纸发布/归档清单（可选，用于回填）",
  },
};

const $ = (selector) => document.querySelector(selector);

function setDefaultDate() {
  const day = new Date();
  day.setDate(day.getDate() - 1);
  $("#target-date").value = day.toISOString().slice(0, 10);
}

function setStatus(text, type = "info") {
  const el = $("#status");
  if (!text) {
    el.className = "status hidden";
    el.textContent = "";
    return;
  }
  el.className = `status ${type}`;
  el.textContent = text;
}

function switchScene(scene) {
  state.scene = scene;
  document.querySelectorAll(".scene-tab").forEach((button) => {
    button.classList.toggle("active", button.dataset.scene === scene);
  });
  $("#page-title").textContent = sceneMeta[scene].title;
  $("#primary-label").textContent = sceneMeta[scene].primaryLabel;
  $("#archive-label").textContent = sceneMeta[scene].archiveLabel;
  $("#primary-file").value = "";
  $("#archive-file").value = "";
  setStatus("");
}

function confidenceText(value) {
  return { high: "高", medium: "中", low: "低", missing: "缺失" }[value] || value;
}

function renderMappings(result) {
  const primary = result.mappings?.primary || [];
  const archive = result.mappings?.archive || [];
  const score = result.detected_tables?.primary?.mapping_score ?? 0;
  $("#mapping-score").textContent = `主清单 ${(score * 100).toFixed(0)}%`;

  const warnings = result.warnings || [];
  $("#warnings").innerHTML = warnings.map((item) => `<div class="warning-item">${escapeHtml(item)}</div>`).join("");

  const rows = [
    ...primary.map((item) => ({ ...item, role: "主清单" })),
    ...archive.map((item) => ({ ...item, role: "归档/发布" })),
  ];
  if (!rows.length) {
    $("#mapping-list").className = "mapping-list empty-state";
    $("#mapping-list").textContent = "暂无字段映射。";
    return;
  }
  $("#mapping-list").className = "mapping-list";
  $("#mapping-list").innerHTML = rows
    .map(
      (item) => `
      <div class="mapping-row" title="${escapeHtml(item.evidence || "")}">
        <strong>${escapeHtml(item.role)} · ${escapeHtml(item.label)}</strong>
        <span>${escapeHtml(item.column || "未识别")}</span>
        <em class="confidence ${escapeHtml(item.confidence)}">${confidenceText(item.confidence)}</em>
      </div>
    `,
    )
    .join("");
}

function renderSummary(summary) {
  const priority = ["未完成数量", "截至统计日期计划数量", "截至统计日期已完成数量", "本次按归档清单可回填数量", "本次按发布清单可回填数量", "D行排除数量", "D行排除未完成数量"];
  const entries = Object.entries(summary || {});
  const cards = priority
    .map((key) => entries.find(([name]) => name === key))
    .filter(Boolean)
    .slice(0, 4);
  if (!cards.length) {
    $("#summary-cards").className = "summary-grid empty-state";
    $("#summary-cards").textContent = "暂无结果。";
    return;
  }
  $("#summary-cards").className = "summary-grid";
  $("#summary-cards").innerHTML = cards
    .map(([key, value]) => `<div class="summary-card"><span>${escapeHtml(key)}</span><strong>${escapeHtml(String(value ?? ""))}</strong></div>`)
    .join("");
}

function renderStats(result) {
  const groups = result.group_stats || [];
  const types = result.type_stats || [];
  $("#stats-area").innerHTML = `
    ${renderMiniTable("功能组统计", groups.slice(0, 8))}
    ${renderMiniTable("类型统计", types.slice(0, 8))}
  `;
}

function renderMiniTable(title, rows) {
  if (!rows.length) {
    return `<div class="mini-table"><h4>${escapeHtml(title)}</h4><div class="status">暂无数据</div></div>`;
  }
  return `<div class="mini-table"><h4>${escapeHtml(title)}</h4>${renderTable(rows, 8)}</div>`;
}

function renderResultTable(rows) {
  $("#row-count").textContent = `${rows.length} 条`;
  if (!rows.length) {
    $("#result-table").className = "table-wrap empty-state";
    $("#result-table").textContent = "当前统计日期没有未完成明细。";
    return;
  }
  $("#result-table").className = "table-wrap";
  $("#result-table").innerHTML = renderTable(rows, 300);
}

function renderTable(rows, limit = 100) {
  const data = rows.slice(0, limit);
  if (!data.length) return "";
  const columns = Object.keys(data[0]);
  const head = columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("");
  const body = data
    .map((row) => `<tr>${columns.map((column) => `<td>${escapeHtml(String(row[column] ?? ""))}</td>`).join("")}</tr>`)
    .join("");
  return `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

function renderResult(result) {
  state.result = result;
  renderMappings(result);
  renderSummary(result.summary);
  renderStats(result);
  renderResultTable(result.rows || []);
  $("#export-link").classList.remove("disabled");
  $("#export-link").href = result.export_url;
}

function appendMessage(type, text, tableData = null) {
  const history = $("#chat-history");
  const message = document.createElement("div");
  message.className = `message ${type}`;
  const bubble = document.createElement("div");
  bubble.className = "bubble";
  bubble.textContent = text;
  if (tableData?.records?.length) {
    const wrap = document.createElement("div");
    wrap.className = "table-wrap";
    wrap.innerHTML = renderTable(tableData.records, 80);
    bubble.appendChild(wrap);
  }
  message.appendChild(bubble);
  history.appendChild(message);
  history.scrollTop = history.scrollHeight;
}

async function analyze(event) {
  event.preventDefault();
  if (state.busy) return;
  const primary = $("#primary-file").files[0];
  if (!primary) {
    setStatus("请先选择主清单 Excel 文件。", "error");
    return;
  }

  const data = new FormData();
  data.append("primary_file", primary);
  if ($("#archive-file").files[0]) data.append("archive_file", $("#archive-file").files[0]);
  if ($("#target-date").value) data.append("target_date", $("#target-date").value);

  state.busy = true;
  $("#analyze-button").disabled = true;
  setStatus("正在识别字段并计算未完成清单...");
  try {
    const response = await fetch(`/api/analyze/${state.scene}`, { method: "POST", body: data });
    const payload = await response.json();
    if (!response.ok) {
      const message = typeof payload.detail === "string" ? payload.detail : payload.detail?.message || "分析失败";
      throw new Error(message);
    }
    renderResult(payload);
    setStatus(`已生成 ${payload.rows.length} 条未完成明细。`);
    appendMessage("assistant", `已完成 ${payload.scene_label} 分析，当前未完成 ${payload.rows.length} 项。`);
  } catch (error) {
    setStatus(error.message || "分析失败", "error");
  } finally {
    state.busy = false;
    $("#analyze-button").disabled = false;
  }
}

async function sendChat(event) {
  event.preventDefault();
  const input = $("#chat-question");
  const question = input.value.trim();
  if (!question) return;
  appendMessage("user", question);
  input.value = "";
  try {
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.result?.session_id, scene: state.scene, question }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || "问答失败");
    appendMessage("assistant", payload.answer, payload.table_data);
  } catch (error) {
    appendMessage("assistant", error.message || "问答失败，请先生成清单。");
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

document.querySelectorAll(".scene-tab").forEach((button) => {
  button.addEventListener("click", () => switchScene(button.dataset.scene));
});
$("#analysis-form").addEventListener("submit", analyze);
$("#chat-form").addEventListener("submit", sendChat);
setDefaultDate();

/* Site de acompanhamento de obras - Triade 2025/2027
   Le a geometria real das ruas (ruas-geo.json, gerado a partir do KML)
   e cruza com os dados da planilha Ruas Triade.xlsx (nome, status, area,
   valores, % concluido). A planilha e relida periodicamente do disco,
   entao basta salvar o Excel para o site atualizar sozinho. */

const XLSX_FILE = "Ruas Triade.xlsx";
const GEO_FILE = "ruas-geo.json";
const SUBPREF_GEO_FILE = "subprefeituras-geo.json";
const REFRESH_MS = 15000;
const SUBPREF_BOUNDARY_COLOR = "#4a4f56";

const COLORS = {
  concluida: "#2ecc71",
  execucao: "#f5a623",
  naoiniciada: "#7d89a8",
  paralisada: "#e74c3c",
  estudo: "#4a90e2",
};

const LABELS = {
  concluida: "Concluída",
  execucao: "Em execução",
  naoiniciada: "Não iniciada",
  paralisada: "Paralisada",
  estudo: "Em estudo",
};

const SUBPREF = {
  SB: "Sapopemba",
  VP: "Vila Prudente",
  AF: "Aricanduva",
  VF: "Vila Formosa",
  CR: "Carrão",
};

// Ordem fixa usada tanto na legenda quanto no agrupamento da lista de ruas.
const STATUS_ORDER = ["concluida", "execucao", "naoiniciada", "paralisada", "estudo"];

let map, layerGroup, geoStreets = [];
let activeCard = null;
let lastMerged = [];
const activeFilters = new Set(STATUS_ORDER);

// Chaves das secoes/grupos que o usuario recolheu (minimizou). Persistem
// entre re-renders (refresh automatico, filtro, etc) ate a pagina recarregar.
const collapsedSections = new Set();

function toggleHtml(key) {
  const collapsed = collapsedSections.has(key);
  return `<button type="button" class="section-toggle${collapsed ? " collapsed" : ""}" data-key="${key}" aria-label="Minimizar/expandir">▾</button>`;
}

function normalizeStatus(raw) {
  const s = (raw || "").toString().normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase();
  if (!s) return "naoiniciada";
  if (s.includes("CONCLU")) return "concluida";
  if (s.includes("PARALIS")) return "paralisada";
  if (s.includes("ESTUD")) return "estudo";
  if (s.includes("EXEC") || s.includes("ANDAMENTO")) return "execucao";
  return "naoiniciada";
}

// SheetJS gives numeric cells back as real JS numbers already (currency/area
// cells are plain numbers, percentage cells come back as a 0..1 fraction).
// Only fall back to text-parsing (Brazilian format: "." thousands, "," decimal)
// when a cell was typed/formatted as text instead of a number.
function parseBRNumber(val) {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === "number") return val;
  const s = val.toString().replace(/[^0-9,.-]/g, "").trim();
  if (!s) return 0;
  const cleaned = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function parsePercent(val) {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === "number") return val <= 1 ? val * 100 : val;
  const n = parseFloat(val.toString().replace("%", "").replace(",", "."));
  return isNaN(n) ? 0 : n;
}

function formatCellDate(val) {
  if (!val) return "";
  if (val instanceof Date) return val.toLocaleDateString("pt-BR");
  return val.toString();
}

function fmtCurrency(n) {
  return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function fmtNumber(n, decimals = 0) {
  return n.toLocaleString("pt-BR", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function normName(s) {
  return (s || "").toString().trim().replace(/\s+/g, " ").toLowerCase();
}

function setStatusBar(state, text) {
  const dot = document.getElementById("status-dot");
  dot.className = "dot" + (state === "loading" ? " loading" : state === "error" ? " error" : "");
  document.getElementById("status-text").textContent = text;
}

async function loadGeo() {
  const res = await fetch(GEO_FILE + "?t=" + Date.now());
  if (!res.ok) throw new Error("Não foi possível ler " + GEO_FILE);
  geoStreets = await res.json();
}

async function loadSubprefGeo() {
  try {
    const res = await fetch(SUBPREF_GEO_FILE + "?t=" + Date.now());
    if (!res.ok) return;
    const data = await res.json();
    renderSubprefBoundaries(data);
  } catch (err) {
    console.error("Não foi possível carregar os perímetros das subprefeituras:", err);
  }
}

function renderSubprefBoundaries(subprefs) {
  const layer = L.featureGroup().addTo(map);
  subprefs.forEach((sp) => {
    L.polygon(sp.rings, {
      color: SUBPREF_BOUNDARY_COLOR,
      weight: 3,
      opacity: 0.95,
      dashArray: "10 7",
      lineCap: "butt",
      fill: true,
      fillColor: SUBPREF_BOUNDARY_COLOR,
      fillOpacity: 0.05,
      interactive: false,
    }).addTo(layer);

    if (sp.centroide) {
      L.marker(sp.centroide, {
        interactive: false,
        icon: L.divIcon({
          className: "subpref-label",
          html: sp.nome.toUpperCase(),
          iconSize: null,
        }),
      }).addTo(layer);
    }
  });
  layer.bringToBack();
}

async function loadPlanilha() {
  const res = await fetch(XLSX_FILE + "?t=" + Date.now());
  if (!res.ok) throw new Error("Não foi possível ler a planilha");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows
    .filter((r) => (r["Nome da Via"] || "").toString().trim() !== "")
    .map((r) => ({
      contrato: r["CONTRATO"] || "",
      subprefeitura: r["Subprefeitura"] || "",
      nome: r["Nome da Via"] || "",
      prazo: r["Prazo"] || "",
      status: r["Status"] || "",
      dataInicio: formatCellDate(r["Data Início"]),
      dataFim: formatCellDate(r["Data Fim"]),
      area: parseBRNumber(r["Área (m²)"]),
      valorEstimado: parseBRNumber(r["Valor Estimado (R$)"]),
      pctConcluido: parsePercent(r["% Concluído"]),
      valorMedido: parseBRNumber(r["Valor Medido (R$)"]),
      valorAMedir: parseBRNumber(r["Valor a Medir (R$)"]),
    }));
}

function mergeData(geo, planilha) {
  const byName = new Map();
  planilha.forEach((p) => byName.set(normName(p.nome), p));

  const merged = geo.map((g) => {
    const p = byName.get(normName(g.nome));
    byName.delete(normName(g.nome));
    if (p) {
      return { ...p, coords: g.coords, temGeometria: true, statusKey: normalizeStatus(p.status) };
    }
    return {
      nome: g.nome,
      contrato: "",
      subprefeitura: "",
      status: "Em estudo",
      statusKey: "estudo",
      prazo: "", dataInicio: "", dataFim: "",
      area: 0, valorEstimado: 0, pctConcluido: 0, valorMedido: 0, valorAMedir: 0,
      coords: g.coords, temGeometria: true,
    };
  });

  // linhas da planilha sem geometria correspondente no KML
  byName.forEach((p) => {
    merged.push({ ...p, coords: null, temGeometria: false, statusKey: normalizeStatus(p.status) });
  });

  return merged;
}

function initMap() {
  map = L.map("map", { zoomControl: true });

  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

  layerGroup = L.layerGroup().addTo(map);

  const legend = L.control({ position: "topright" });
  legend.onAdd = function () {
    const div = L.DomUtil.create("div", "legend-box");
    div.innerHTML =
      `<div class="legend-title">Filtrar por status</div>` +
      STATUS_ORDER.map(
        (k) =>
          `<button type="button" class="legend-row active" data-status="${k}">` +
          `<span class="legend-swatch" style="background:${COLORS[k]}"></span> ${LABELS[k]}</button>`
      ).join("");

    L.DomEvent.disableClickPropagation(div);
    div.querySelectorAll(".legend-row").forEach((btn) => {
      btn.addEventListener("click", () => {
        const status = btn.getAttribute("data-status");
        if (activeFilters.has(status)) {
          activeFilters.delete(status);
          btn.classList.remove("active");
        } else {
          activeFilters.add(status);
          btn.classList.add("active");
        }
        applyFilters();
      });
    });
    return div;
  };
  legend.addTo(map);
}

function applyFilters() {
  const visible = lastMerged.filter((s) => activeFilters.has(s.statusKey));
  const lineByName = renderMap(visible);
  renderList(visible, lineByName);
}

function popupHtml(s) {
  const rows = [];
  if (s.contrato) rows.push(["Contrato", s.contrato]);
  if (s.subprefeitura) rows.push(["Subprefeitura", SUBPREF[s.subprefeitura] || s.subprefeitura]);
  rows.push(["Status", LABELS[s.statusKey]]);
  if (s.dataInicio) rows.push(["Início", s.dataInicio]);
  if (s.dataFim) rows.push(["Fim previsto", s.dataFim]);
  if (s.area) rows.push(["Área", fmtNumber(s.area, 2) + " m²"]);
  if (s.valorEstimado) rows.push(["Valor estimado", fmtCurrency(s.valorEstimado)]);
  rows.push(["% concluído", fmtNumber(s.pctConcluido, 0) + "%"]);
  if (s.valorMedido) rows.push(["Valor medido", fmtCurrency(s.valorMedido)]);
  if (s.valorAMedir) rows.push(["Valor a medir", fmtCurrency(s.valorAMedir)]);
  return (
    `<div class="popup-title">${s.nome}</div>` +
    rows.map(([k, v]) => `<div class="popup-row"><span class="k">${k}</span><span>${v}</span></div>`).join("")
  );
}

function renderMap(streets) {
  layerGroup.clearLayers();
  const bounds = [];
  const lineByName = new Map();

  streets.forEach((s) => {
    if (!s.coords || !s.coords.length) return;
    const color = COLORS[s.statusKey] || COLORS.naoiniciada;
    const line = L.polyline(s.coords, {
      color,
      weight: 6,
      opacity: 0.9,
      dashArray: s.statusKey === "estudo" ? "6 6" : null,
    });
    line.bindTooltip(`${s.nome} — ${LABELS[s.statusKey]}`, { sticky: true });
    line.bindPopup(popupHtml(s));
    line.on("mouseover", () => line.setStyle({ weight: 10 }));
    line.on("mouseout", () => line.setStyle({ weight: 6 }));
    line.addTo(layerGroup);
    lineByName.set(normName(s.nome), line);
    s.coords.forEach((c) => bounds.push(c));
  });

  if (bounds.length && !map._fitted) {
    map.fitBounds(bounds, { padding: [30, 30] });
    map._fitted = true;
  }

  return lineByName;
}

function renderSummary(streets) {
  // Ruas em estudo ainda nao tem contrato - ficam fora do bloco principal e
  // fora do calculo de avanco geral, para nao diluir o % com vias que nem
  // comecaram a ser contratadas.
  const contratadas = streets.filter((s) => s.statusKey !== "estudo");
  const emEstudo = streets.filter((s) => s.statusKey === "estudo");

  const totalArea = contratadas.reduce((a, s) => a + s.area, 0);
  const totalEstimado = contratadas.reduce((a, s) => a + s.valorEstimado, 0);
  const totalMedido = contratadas.reduce((a, s) => a + s.valorMedido, 0);
  const totalAMedir = contratadas.reduce((a, s) => a + s.valorAMedir, 0);
  const pctMedio = totalArea > 0
    ? contratadas.reduce((a, s) => a + s.pctConcluido * s.area, 0) / totalArea
    : 0;
  const concluidas = contratadas.filter((s) => s.statusKey === "concluida").length;

  const contratadasCollapsed = collapsedSections.has("contratadas");
  document.getElementById("summary-contratadas").innerHTML = `
    <div class="summary-section-label">
      ${toggleHtml("contratadas")}
      Ruas contratadas
    </div>
    <div class="summary-grid${contratadasCollapsed ? " collapsed" : ""}">
      <div class="card">
        <div class="label">Ruas no contrato</div>
        <div class="value">${contratadas.length}</div>
      </div>
      <div class="card">
        <div class="label">Concluídas</div>
        <div class="value">${concluidas}</div>
      </div>
      <div class="card" style="grid-column: 1 / -1;">
        <div class="label">Avanço geral (ponderado por área)</div>
        <div class="value">${fmtNumber(pctMedio, 1)}%</div>
        <div class="progress-outer"><div class="progress-inner" style="width:${Math.min(pctMedio,100)}%"></div></div>
      </div>
      <div class="card">
        <div class="label">Área total</div>
        <div class="value">${fmtNumber(totalArea, 0)} m²</div>
      </div>
      <div class="card">
        <div class="label">Valor estimado</div>
        <div class="value">${fmtCurrency(totalEstimado)}</div>
      </div>
      <div class="card">
        <div class="label">Valor medido</div>
        <div class="value">${fmtCurrency(totalMedido)}</div>
      </div>
      <div class="card">
        <div class="label">Valor a medir</div>
        <div class="value">${fmtCurrency(totalAMedir)}</div>
      </div>
    </div>
  `;

  renderEstudoSummary(emEstudo);
}

function renderEstudoSummary(emEstudo) {
  const el = document.getElementById("summary-estudo");
  if (!emEstudo.length) {
    el.innerHTML = "";
    return;
  }
  const totalArea = emEstudo.reduce((a, s) => a + s.area, 0);
  const totalEstimado = emEstudo.reduce((a, s) => a + s.valorEstimado, 0);
  const estudoCollapsed = collapsedSections.has("estudo");

  el.innerHTML = `
    <div class="summary-section-label estudo">
      ${toggleHtml("estudo")}
      <span class="legend-swatch" style="background:${COLORS.estudo}"></span>
      Em estudo (fora do contrato)
    </div>
    <div class="summary-grid estudo-grid${estudoCollapsed ? " collapsed" : ""}">
      <div class="card estudo">
        <div class="label">Ruas em estudo</div>
        <div class="value">${emEstudo.length}</div>
      </div>
      <div class="card estudo">
        <div class="label">Área em estudo</div>
        <div class="value">${fmtNumber(totalArea, 0)} m²</div>
      </div>
      <div class="card estudo" style="grid-column: 1 / -1;">
        <div class="label">Valor estimado (em estudo)</div>
        <div class="value">${fmtCurrency(totalEstimado)}</div>
      </div>
    </div>
  `;
}

function renderList(streets, lineByName) {
  const withData = streets.filter((s) => s.contrato || s.status);
  const list = document.getElementById("street-list");
  list.innerHTML = "";
  activeCard = null;

  STATUS_ORDER.forEach((statusKey) => {
    const group = withData
      .filter((s) => s.statusKey === statusKey)
      .sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
    if (!group.length) return;

    const groupKey = "group-" + statusKey;
    const groupCollapsed = collapsedSections.has(groupKey);

    const header = document.createElement("div");
    header.className = "group-header";
    header.innerHTML =
      toggleHtml(groupKey) +
      `<span class="legend-swatch" style="background:${COLORS[statusKey]}"></span>` +
      `${LABELS[statusKey]} <span class="group-count">(${group.length})</span>`;
    list.appendChild(header);

    const body = document.createElement("div");
    body.className = "group-body" + (groupCollapsed ? " collapsed" : "");
    list.appendChild(body);

    group.forEach((s) => {
      const card = document.createElement("div");
      card.className = "street-card";
      card.style.borderLeftColor = COLORS[s.statusKey] || COLORS.naoiniciada;
      const badgeClass = s.statusKey === "naoiniciada" ? "naoiniciada" : s.statusKey;
      card.innerHTML = `
        <div class="name">${s.nome}</div>
        <div class="meta">
          <span class="badge ${badgeClass}">${LABELS[s.statusKey]}</span>
          <span>${fmtNumber(s.pctConcluido, 0)}%</span>
        </div>
        ${!s.temGeometria ? '<div class="meta" style="margin-top:4px;color:#e74c3c">Sem geometria no mapa (não encontrada no KML)</div>' : ""}
      `;
      card.addEventListener("click", () => {
        if (activeCard) activeCard.classList.remove("active");
        card.classList.add("active");
        activeCard = card;
        const line = lineByName.get(normName(s.nome));
        if (line) {
          map.fitBounds(line.getBounds(), { padding: [60, 60] });
          line.openPopup();
        }
      });
      body.appendChild(card);
    });
  });

  if (!withData.length) {
    list.innerHTML = '<div class="empty-list">Nenhuma rua com o filtro atual.</div>';
  }
}

async function refresh(isFirst) {
  try {
    setStatusBar("loading", isFirst ? "Carregando planilha..." : "Atualizando...");
    const planilha = await loadPlanilha();
    const merged = mergeData(geoStreets, planilha);
    lastMerged = merged;
    renderSummary(merged);
    applyFilters();
    const now = new Date().toLocaleTimeString("pt-BR");
    setStatusBar("ok", "Atualizado às " + now);
  } catch (err) {
    console.error(err);
    setStatusBar("error", "Erro: " + err.message);
  }
}

function setupCollapseToggles() {
  document.getElementById("sidebar").addEventListener("click", (ev) => {
    const btn = ev.target.closest(".section-toggle");
    if (!btn) return;
    const key = btn.getAttribute("data-key");
    if (collapsedSections.has(key)) {
      collapsedSections.delete(key);
    } else {
      collapsedSections.add(key);
    }
    renderSummary(lastMerged);
    applyFilters();
  });
}

async function main() {
  initMap();
  setupCollapseToggles();
  loadSubprefGeo();
  await loadGeo();
  await refresh(true);
  setInterval(() => refresh(false), REFRESH_MS);
  document.getElementById("refresh-btn").addEventListener("click", () => refresh(false));
}

main();

window.onerror = function (msg, src, line, col, err) {
  alert("JS ERROR:\n" + msg + "\nLine: " + line);
  console.error(err);
};
/* ======================
   GLOBAL STATE
====================== */
let markers = [];
let searchIndex = [];
let userLatLng = null;
let routingControl = null;
let petugasMarker = null;

let history = JSON.parse(
  localStorage.getItem("oyon_history") || "[]"
);

let originalRows = [];
let workingRows = [];

let filterHariBacaAktif = false;
let daftarHariBaca = [];
let indexHariBaca = 0;

let showDIJ = true;
let showDIL = true;

let isFullMap = false;
let hasSelectedCustomer = false;
let historyVisible = false;

let hiddenMarkers = [];

window.isCompareMode = false;
window.compareMarkers = [];

/* ======================
   GPS COMPATIBILITY ENGINE
====================== */

function getDIJLat(row) {
  return parseFloat(
    row["LAT DIJ"] ||
    row["LAT"] ||
    0
  );
}

function getDIJLon(row) {
  return parseFloat(
    row["LON DIJ"] ||
    row["LON"] ||
    0
  );
}

function getDILLat(row) {
  return parseFloat(
    row["LAT DIL"] ||
    row["LATTH"] ||
    row["LATDIL"] ||
    0
  );
}

function getDILLon(row) {
  return parseFloat(
    row["LON DIL"] ||
    row["LONTH"] ||
    row["LONDIL"] ||
    0
  );
}
/* ======================
   UTIL — HARI BACA DARI KDDK ACMT
====================== */
function getHariBaca(row) {

  const kddk = (
    row["KDDK"] ||
    row["KDDK ACMT"] ||
    row["KODEDK"] ||
    ""
  ).toString();

  if (!kddk) return "-";

  // huruf ke-6 dari belakang
  const hb = kddk.slice(-6, -5);

  return hb ? hb.toUpperCase() : "-";

}

function getNoMeter(row){

  return (
    row.NOMORKWH ||
    row.NOMET ||
    row.NO_METER ||
    row.NOMETER ||
    "-"
  );

}

function getMerkMeter(row){

  return (
    row.MEREKKWH ||
    row.MERK ||
    row.MERKMETER ||
    "-"
  );

}

function getKDDK(row){

  return (
    row.KDDK ||
    row.KODEDK ||
    row.KODUK ||
    "-"
  );

}

function renderMarker(m, map) {
  if (!m || !m.marker) return;

  const marker = m.marker;

  // kalau sudah ada di map, jangan di-add lagi
  if (map.hasLayer(marker)) return;

  clusterGroup.addLayer(marker);
}


const STORAGE_KEY = "RBM_WORKING_STATE";


// Reset history setiap hari
const today = new Date().toDateString();
const savedDate = localStorage.getItem("oyon_history_date");

function closeHistory() {
  const panel = document.getElementById("historyPanel");
  if (panel) {
    panel.style.display = "none";
  }
}

if (savedDate !== today) {
  history = [];
  localStorage.setItem("oyon_history", "[]");
  localStorage.setItem("oyon_history_date", today);
  renderHistory();

}

/* ======================
   DOM
====================== */
const upload = document.getElementById("upload");
const statusText = document.getElementById("statusText");
const searchBtn = document.getElementById("searchBtn");
const searchInput = document.getElementById("searchInput");
const zoomBtn = document.getElementById("zoomBtn");
const downloadBtn = document.getElementById("downloadBtn");

/* ======================
   INIT MAP
====================== */
const map = L.map("map", {
  preferCanvas: true,
  rotate: true,
  zoomControl: false,
  touchRotate: true,
  rotateControl: true,
  maxZoom: 19,
}).setView([-7.974155,112.6181065], 12);

/* ======================
   CLUSTER ENGINE
====================== */

const clusterGroup = L.markerClusterGroup({

  chunkedLoading: true,
  chunkInterval: 80,
  chunkDelay: 20,

  removeOutsideVisibleBounds: true,

  spiderfyOnMaxZoom: true,

  showCoverageOnHover: false,

  zoomToBoundsOnClick: true,

  disableClusteringAtZoom: 17,

  maxClusterRadius: 18

});

map.addLayer(clusterGroup);

function addMarkerToMap(marker) {

  if (!marker) return;

  if (!clusterGroup.hasLayer(marker)) {
    clusterGroup.addLayer(marker);
  }

}

L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom: 19
}).addTo(map);

/* ======================
   AUTO VIEWPORT RENDER
====================== */

map.on("moveend", renderVisibleMarkers);
map.on("zoomend", renderVisibleMarkers);


/* ======================
   IKUT KOMPAS
====================== */
let lastBearing = null;
let lastUpdate = 0;
let compassHandler = null;

function enableCompass() {
  if (compassHandler) return; // sudah aktif

  compassHandler = (event) => {
    const now = Date.now();
    if (now - lastUpdate < 120) return;
    if (event.alpha == null) return;

    const target = 360 - event.alpha;
    if (lastBearing == null) lastBearing = target;

    const smooth = lastBearing + (target - lastBearing) * 0.25;
    map.setBearing(smooth);

    lastBearing = smooth;
    lastUpdate = now;
  };

  window.addEventListener("deviceorientation", compassHandler);
  showToast("🧭 Kompas aktif");
}

function disableCompass() {
  if (compassHandler) {
    window.removeEventListener("deviceorientation", compassHandler);
    compassHandler = null;
    showToast("🛑 Kompas mati");
  }
}

/* ======================
   LOAD EXCEL
====================== */
upload.onchange = e => {
  const file = e.target.files[0];
  if (!file) return;

  // SIMPAN nama file
  localStorage.setItem("RBM_FILENAME", file.name);

  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(new Uint8Array(evt.target.result), { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    originalRows = XLSX.utils.sheet_to_json(sheet);
    
    workingRows = structuredClone(originalRows);
    saveState();

    buildMarkers();
    renderHistory();

    

    if (statusText) {
      statusText.innerText = "✅ Peta berhasil dibuat";
    }
    zoomBtn.style.display = "block";

    // auto zoom setelah upload
    zoomToAllMarkers();
  };
  reader.readAsArrayBuffer(file);

  document.addEventListener("DOMContentLoaded", () => {
  const statusText = document.getElementById("statusText");
});

};


/* ======================
   KUMPULKAN HARI BACA
====================== */
hariBacaList = [...new Set(
  workingRows
    .map(r => String(r["HARI BACA ACMT"]).trim())
    .filter(v => v !== "")
)].sort((a,b)=>a-b);

/* ======================
   SAAT LOAD
====================== */
window.addEventListener("load", () => {

  // Restore filename
  const savedFilename = localStorage.getItem("RBM_FILENAME");
  if (savedFilename) {
    document.getElementById("fileNameLabel").innerText = "📄 " + savedFilename;
  }

  const saved = localStorage.getItem(STORAGE_KEY);

  if (saved) {
    workingRows = JSON.parse(saved);

    map.whenReady(() => {
      buildMarkers();
      renderHistory();

      // Tampilkan tombol zoom
      if (workingRows.length > 0) {
        zoomBtn.style.display = "block";
      }

      // Auto zoom ke marker
      zoomToAllMarkers();
    });
  }
});

function zoomToAllMarkers() {
  if (!markers.length) return;
  const group = L.featureGroup(markers.map(m => m.marker));
  map.fitBounds(group.getBounds(), { padding: [40, 40] });
}
/* ======================
   FUNGSI FILTER MARKER HARI BACA
====================== */
function applyHariBacaFilter() {

  const showDIJ = document.getElementById("chkDIJ")?.checked ?? true;
  const showDIL  = document.getElementById("chkDIL")?.checked ?? true;

  markers.forEach(m => {

    let visible = true;

    // filter jenis marker
    if (m.type === "DIJ" && !showDIJ) visible = false;
    if (m.type === "DIL"  && !showDIL)  visible = false;

    // filter 1 hari baca
    if (filterHariBacaAktif) {
      const hari = getHariBaca(m.row);
      if (hari !== daftarHariBaca[indexHariBaca]) {
        visible = false;
      }
    }

    if (visible) {
      if (!clusterGroup.hasLayer(m.marker)) clusterGroup.addLayer(m.marker);
    } else {
      if (clusterGroup.hasLayer(m.marker)) clusterGroup.removeLayer(m.marker);
    }
  });

  document.getElementById("hariBacaLabel").innerText =
    filterHariBacaAktif
      ? `HARI BACA: ${daftarHariBaca[indexHariBaca]}`
      :"HARI BACA: ALL"

  renderVisibleMarkers();
}

function showNormalMarkers() {
  markers.forEach(m => {
    if (!clusterGroup.hasLayer(m.marker)) {
      addMarkerToMap(m.marker);
    }
  });
}
/* ======================
   VIEWPORT RENDER ENGINE
====================== */

function renderVisibleMarkers() {
  if (window.isCompareMode) return;

  const bounds = map.getBounds();

  markers.forEach(m => {

    let visible = true;

    // FILTER DIJ / DIL
    const showDIJ =
      document.getElementById("chkDIJ")?.checked ?? true;

    const showDIL =
      document.getElementById("chkDIL")?.checked ?? true;

    if (m.type === "DIJ" && !showDIJ) {
      visible = false;
    }

    if (m.type === "DIL" && !showDIL) {
      visible = false;
    }

    // FILTER HARI BACA
    if (filterHariBacaAktif) {

      const hari = getHariBaca(m.row);

      if (hari !== daftarHariBaca[indexHariBaca]) {
        visible = false;
      }
    }

    // VIEWPORT CHECK
    const latlng = m.marker.getLatLng();

    const insideViewport =
      bounds.contains(latlng);

    // ===== SHOW =====
    if (visible && insideViewport) {

      if (!clusterGroup.hasLayer(m.marker)) {
        clusterGroup.addLayer(m.marker);
      }

    }

    // ===== HIDE =====
    else {

      if (clusterGroup.hasLayer(m.marker)) {
        clusterGroup.removeLayer(m.marker);
      }

    }

  });

}

/* ======================
   SIMPAN STATE
====================== */
function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(workingRows));
}
/* ======================
   TAMPILKAN marker dij/Lokasi DIL saja
====================== */
function applyMarkerFilter() {}
/* ======================
function applyHariBacaFilter() {
  if (!filterHariBacaAktif) {
    applyMarkerFilter();
    return;
  }

  const hariAktif = daftarHariBaca[indexHariBaca];

  markers.forEach(m => {
    const hariRow = (m.row["HARI BACA ACMT"] || "").toUpperCase();

    if (hariRow === hariAktif) {
      addMarkerToMap(m.marker);
    } else {
      if (map.hasLayer(m.marker)) {
        clusterGroup.removeLayer(m.marker);
      }
    }
  });

  updateHariBacaLabel();
}
  ====================== */

function nextHari(){
  if (!filterHariBacaAktif) return;
  if (!daftarHariBaca.length) return;

  indexHariBaca++;

  if (indexHariBaca >= daftarHariBaca.length) {
    indexHariBaca = 0;
  }

  applyHariBacaFilter();
}

function prevHari(){
  if (!filterHariBacaAktif) return;
  if (!daftarHariBaca.length) return;

  indexHariBaca--;

  if (indexHariBaca < 0) {
    indexHariBaca = daftarHariBaca.length - 1;
  }

  applyHariBacaFilter();
}


function updateHariBacaLabel() {
  const el = document.getElementById("hariBacaLabel");
  if (!el) return;

  el.innerText = filterHariBacaAktif
    ? `HARI BACA: ${daftarHariBaca[indexHariBaca]}`
    : "HARI BACA: ALL";
}

window.addEventListener("DOMContentLoaded", () => {
  const cb = document.getElementById("oneDayCheck");
  if (!cb) return;

  cb.onchange = e => {
    filterHariBacaAktif = e.target.checked;

    if (filterHariBacaAktif && daftarHariBaca.length === 0) {
      showToast("⚠️ Data hari baca kosong");
      cb.checked = false;
      filterHariBacaAktif = false;
      return;
    }

    applyHariBacaFilter();
  };
});


/* ======================
   TAMPILKAN NAMA FILE SAAT REFRESH
====================== */

/* ======================
   BUILD MARKERS
====================== */
function buildMarkers(){
  // 🔥 HAPUS SEMUA MARKER LAMA DARI MAP
clusterGroup.clearLayers();

markers = [];
searchIndex = [];

  workingRows.forEach(row => {
    // ==========================
    // MARKER 1 — LOKASI DIJ
    // ==========================
    const lat = getDIJLat(row);
    const lng = getDIJLon(row);

    if(lat && lng){
      const hariBaca = getHariBaca(row);
      const color = getColor(hariBaca);


      const iconDIJ = L.divIcon({
        html: `<div class="custom-marker" style="background:${color}"></div>`,
        iconSize: [16,16],
        iconAnchor: [8,16],
        className: ""
      });

      const markerDIJ = L.marker([lat, lng], { icon: iconDIJ });

      markerDIJ.bindTooltip(row.NAMA, {
        permanent: false,
        direction: "top",
        offset: [0, -10],
        opacity: 0.9,
        className: "marker-label"
      });

      markerDIJ.on("click", () => openDetail(row));

      markers.push({
        marker: markerDIJ,
        row: row,
        type: "DIJ"
      });

            searchIndex.push({
        idpel: String(row.IDPEL||"").replace(/\D/g,""),
        meter: String(getNoMeter(row)||"").replace(/\D/g,""),
        marker: markerDIJ,
        row
      });
      
    }
    

    // ==========================
// MARKER 2 — TAGGING Lokasi DIL (HITAM)
// ==========================
const latDIL = getDILLat(row);
const lonDIL = getDILLon(row);

const hasDILLocation =
  latDIL && lonDIL &&
  !isNaN(latDIL) &&
  !isNaN(lonDIL);

if (hasDILLocation) {

  // ICON Lokasi DIL (WAJIB DULU)
  const iconDIL = L.divIcon({
    html: `<div class="custom-marker" style="background:#000"></div>`,
    iconSize: [12, 12],
    iconAnchor: [6, 12],
    className: ""
  });

  // MARKER Lokasi DIL
  const markerDIL = L.marker([latDIL, lonDIL], { icon: iconDIL });

  markerDIL.bindTooltip(
    `<b>${row.NAMA}</b><br>IDPEL: ${row.IDPEL}`,
    {
      permanent: false,
      direction: "top",
      offset: [0, -10],
      opacity: 0.8,
      className: "marker-label"
    }
  );

  markerDIL.on("click", () => openDetail(row));

  markers.push({
    marker: markerDIL,
    row,
    type: "DIL"
  });
}

  }); // ← TUTUP workingRows.forEach

  // ===== KUMPULKAN HARI BACA UNIK (SETELAH BUILD MARKER) =====
daftarHariBaca = [...new Set(
  workingRows
    .map(r => getHariBaca(r))
    .filter(v => v)
)].sort();

// ⛔ JANGAN reset index kalau filter sudah aktif
if (!filterHariBacaAktif) {
  indexHariBaca = 0;
}

updateHariBacaLabel();
applyHariBacaFilter();
renderVisibleMarkers();

setTimeout(() => {
  renderVisibleMarkers();
}, 200);
}     // ← TUTUP function buildMarkers

/* ======================
markers.forEach(m => {
  if (m.marker && m.marker.addTo) {
    clusterGroup.addLayer(m.marker);
  }
});====================== */

/* ======================
   HIGHLIGHT MARKER
====================== */
function highlightMarker(marker){
  const el = marker.getElement();
  if(!el) return;

  el.classList.add("highlight");

  setTimeout(() => {
    el.classList.remove("highlight");
  }, 1500);
}

function showLabel(marker) {
  marker.openTooltip();

  setTimeout(() => {
    marker.closeTooltip();
  }, 4000); // nama tampil 4 detik lalu hilang otomatis
}


function toggleRoute(cb) {
  if (cb.checked) {
    enableCompass();
    showToast("🧭 Ikuti kompas aktif");
  } else {
    disableCompass();
    showToast("🛑 Ikuti kompas mati");
  }
}

/* ======================
   SEARCH
====================== */
searchBtn.onclick = () => {

  const key =
    searchInput.value.trim();

  if(key.length < 4){

    showToast("Masukkan minimal 4 digit");
    return;

  }

  if(searchIndex.length === 0){

    showToast("Data belum siap, upload dulu");
    return;

  }

  // MULTI RESULT
  const results = searchIndex.filter(o =>

    o.idpel.endsWith(key) ||
    o.meter.endsWith(key)

  );

  // TIDAK ADA
  if(!results.length){

    showToast("Tidak ditemukan");
    return;

  }

  // SATU HASIL
  if(results.length === 1){

    focusSearchResult(results[0]);
    return;

  }

  // MULTI RESULT UI
  showSearchResults(results);

};

function focusSearchResult(found){

  addMarkerToMap(found.marker);

  clusterGroup.addLayer(found.marker);

  map.setView(
    found.marker.getLatLng(),
    17,
    { animate:true }
  );

  highlightMarker(found.marker);

  showLabel(found.marker);

  openDetail(found.row);

  // AUTO MINIMIZE
  const uiBar =
    document.getElementById("uiBar");

  const toggleBtn =
    document.getElementById("toggleBtn");

  if(
    uiBar &&
    !uiBar.classList.contains("minimized")
  ){

    uiBar.classList.add("minimized");

    if(toggleBtn)
      toggleBtn.innerText = "➕";

  }

}

function showSearchResults(results){

  let html = `

    <div id="searchPopup">

      <div class="searchPopupTitle">
        Pilih Data
      </div>

  `;

  results.forEach((r, i) => {

    html += `

      <div
        class="searchResultItem"
        onclick="selectSearchResult(${i})"
      >

        <strong>
          ${r.row.NAMA || "-"}
        </strong><br>

        IDPEL:
        ${r.row.IDPEL || "-"}<br>

        NOMET:
        ${getNoMeter(r.row)}

      </div>

    `;

  });

  html += `</div>`;

  window.searchResultsTemp = results;

  document.body.insertAdjacentHTML(
    "beforeend",
    html
  );

}

function selectSearchResult(index){

  const found =
    window.searchResultsTemp[index];

  focusSearchResult(found);

  document
    .getElementById("searchPopup")
    ?.remove();

}

/* ======================
   ZOOM ALL
====================== */
zoomBtn.onclick = () => {
  const active = markers
    .map(m => m.marker)
    .filter(m => clusterGroup.hasLayer(m));

  if (!active.length) {
    showToast("Tidak ada marker aktif");
    return;
  }

  const group = L.featureGroup(active);
  map.fitBounds(group.getBounds(), { padding: [40,40] });
};

/* ======================
   DETAIL PANEL
====================== */
function openDetail(r){
  hasSelectedCustomer = true;
  const panel = document.getElementById("detailPanel");
  const left = document.getElementById("detailLeft");
  const right = document.getElementById("detailRight");

  /* FORCE SHOW MARKERS FOR COMPARE */

    // simpan IDPEL yang sedang dibuka (global state kecil)
  window.currentCompareIDPEL = r.IDPEL;
  
  const cardHTML = (label, lat, lon, isDIL = false) => `
  <div class="${isDIL ? "card-DIL" : ""}">
    <div style="font-size:13px;color:${isDIL ? "#9ca3af" : "#666"};margin-bottom:4px">
      ${label}
    </div>

    <h3 style="margin:4px 0">${r.NAMA}</h3>
    <div style="display:flex;align-items:center;justify-content:space-between;">
      <span>🔢 NOMET: ${getNoMeter(r)}</span>
      <button 
        style="background:#2563eb;color:white;border:none;padding:4px 8px;border-radius:6px;font-size:11px;cursor:pointer;"
        onclick="copyNomet('${getNoMeter(r)}')">
        Copy Nomet
      </button>
      
    </div>
    
    <div>⚡ DAYA: ${r.DAYA}</div>
    <div>🔧 MERK: ${getMerkMeter(r)}</div>
    <div style="display:flex;align-items:center;justify-content:space-between;margin-top:4px;">
        <span>📟 IDPEL: ${r.IDPEL}</span>

        <button
            style="
            background:#059669;
            color:white;
            border:none;
            padding:4px 8px;
            border-radius:6px;
            font-size:11px;
            cursor:pointer;
            "
            onclick="copyIDPEL('${r.IDPEL}')">

            Copy IDPEL
        </button>
        </div>
    <div>📅 HARI BACA: ${getHariBaca(r)}</div>
    <div>🧭 ALAMAT: ${r["ALAMAT"]}</div>

    <div style="margin-top:10px;display:flex;gap:6px;">
      ${
        (!isDIL || (lat && lon))
          ? `<button class="btn-go" onclick="goTo(${lat},${lon})">🧭 Tampil Rute</button>`
          : ``
      }
      <button class="btn-done" onclick="markDone('${r.IDPEL}')">✅ Selesai</button>
    </div>
  </div>
`;

  // tampilkan compare hanya sekali di atas
    left.innerHTML = cardHTML(
    "📍 Lokasi DIJ",
    getDIJLat(r),
    getDIJLon(r),
    false
    );

  const dilLat = getDILLat(r);
        const dilLon = getDILLon(r);

        if (dilLat && dilLon) {

        right.innerHTML = cardHTML(
            "📍 Lokasi DIL",
            dilLat,
            dilLon,
            true
        );

        } else {

        right.innerHTML = `
            <div style="
            padding:20px;
            text-align:center;
            color:#666;
            ">
            📭 Lokasi DIL tidak tersedia
            </div>
        `;
        }

    panel.classList.add("open");
  }


function closeDetail(){
  document.getElementById("detailPanel").classList.remove("open");
  hasSelectedCustomer = false;
}

function toggleFullMap() {
  const uiBar = document.getElementById("uiBar");
  const detailPanel = document.getElementById("detailPanel");
  const btn = document.getElementById("fullMapBtn");

  isFullMap = !isFullMap;

  if (isFullMap) {
    // 🔥 MASUK MODE FULL MAP
    if (uiBar) uiBar.style.display = "none";
    if (detailPanel) detailPanel.classList.remove("open");

    btn.innerText = "📋 SHOW MENU";
    showToast("🗺️ Mode peta penuh");

  } else {
    // 🔥 KELUAR DARI FULL MAP
    if (uiBar) uiBar.style.display = "block";

    // hanya tampilkan bottom bar jika ada pelanggan aktif
    if (hasSelectedCustomer && detailPanel) {
      detailPanel.classList.add("open");
    }

    btn.innerText = "🗺️ FULL MAP";
    showToast("📍 Mode normal");
  }
}


/* ======================



function toggleCompare(checkbox) {
  const idpel = window.currentCompareIDPEL;
  if (!idpel) return;

  if (checkbox.checked) {
    const active = [];

    // 🔥 hide semua marker
    markers.forEach(m => {
      if (map.hasLayer(m.marker)) {
        clusterGroup.removeLayer(m.marker);
      }
    });

    // 🔥 tampilkan hanya marker pelanggan ini
    markers.forEach(m => {
      if (String(m.row.IDPEL) === String(idpel)) {
        addMarkerToMap(m.marker);
        active.push(m.marker);
      }
    });

    // zoom ke 2 marker
    if (active.length > 0) {
      const group = L.featureGroup(active);
      map.fitBounds(group.getBounds(), { padding: [60, 60] });
    }

    showToast("🔍 Mode bandingkan aktif");
  } else {
    // 🔥 KEMBALIKAN SESUAI CHECKBOX
    applyHariBacaFilter();
    showToast("📍 Semua marker ditampilkan kembali");
  }
}
====================== */

/* ======================
   minimize menu saat bandingkan
====================== */

function toggleCompare(checkbox) {
  const idpel = String(window.currentCompareIDPEL || "");
  if (!idpel) return;

  const uiBar = document.getElementById("uiBar");
  const toggleBtn = document.getElementById("toggleBtn");

  if (checkbox.checked) {
    // =========================
    // MODE BANDINGKAN ON
    // =========================
    window.isCompareMode = true;

    // auto minimize menu
    if (uiBar && !uiBar.classList.contains("minimized")) {
      uiBar.classList.add("minimized");
      if (toggleBtn) toggleBtn.innerText = "➕";
    }

    // sembunyikan semua marker
    markers.forEach(m => {
      if (clusterGroup.hasLayer(m.marker)) {
        clusterGroup.removeLayer(m.marker);
      }
    });

    // ambil 2 marker milik IDPEL aktif
    const activeMarkers = markers.filter(
      m => String(m.row.IDPEL) === idpel
    );

    // tampilkan hanya marker tsb
    activeMarkers.forEach(m => {
      clusterGroup.addLayer(m.marker);
      setTimeout(() => {
      renderVisibleMarkers();
    }, 100);
    });

    // zoom ke 2 marker
    if (activeMarkers.length) {
      const group = L.featureGroup(activeMarkers.map(m => m.marker));
      map.flyToBounds(group.getBounds(), {
        padding: [80, 80],
        duration: 0.5
      });
    }

    showToast("🔍 Mode bandingkan aktif");

  } else {
    exitCompareMode();
  }
}
function exitCompareMode() {
  window.isCompareMode = false;

  // bersihkan map
  markers.forEach(m => {
    if (clusterGroup.hasLayer(m.marker)) {
      clusterGroup.removeLayer(m.marker);
    }
  });

  // tampilkan kembali marker SESUAI FILTER
  applyHariBacaFilter();

  // restore menu
  const uiBar = document.getElementById("uiBar");
  const toggleBtn = document.getElementById("toggleBtn");

  if (uiBar && uiBar.classList.contains("minimized")) {
    uiBar.classList.remove("minimized");
    if (toggleBtn) toggleBtn.innerText = "➖";
  }

  showToast("📍 Mode bandingkan selesai");
}


function onFinishCompare() {
  const checkbox = document.getElementById("compareCheckbox");
  if (checkbox) checkbox.checked = false;

  exitCompareMode();
}


function minimizeMenu() {
  const panel = document.getElementById("uiContent");
  if (!panel) return;

  panel.classList.add("minimized");
}
    
/* ======================
   ROUTING
====================== */
function goTo(lat,lng){
  if(!userLatLng) return alert("GPS belum aktif");

  if(routingControl) map.removeControl(routingControl);

      routingControl = L.Routing.control({
      waypoints: [L.latLng(userLatLng), L.latLng(lat,lng)],
      addWaypoints: false,
      draggableWaypoints: false,
      createMarker: () => null
    }).addTo(map);


  document.getElementById("cancelRouteBtn").style.display="block";
}
  

function cancelRoute(){
  if(routingControl) map.removeControl(routingControl);
  routingControl = null;
  document.getElementById("cancelRouteBtn").style.display="none";
}

/* ======================
   DONE + HISTORY
====================== */
function markDone(idpel){

  if (window.isCompareMode) {
  const cb = document.querySelector('#detailPanel input[type="checkbox"]');
  if (cb) cb.checked = false;
  exitCompareMode();
}


  idpel = String(idpel);

  // 1️⃣ ambil data pelanggan
  const found = workingRows.find(r => String(r.IDPEL) === idpel);
  if (!found) return;

  // 2️⃣ simpan ke history
  history.push({
    row: found,
    time: new Date().toLocaleTimeString("id-ID")
  });
  localStorage.setItem("oyon_history", JSON.stringify(history));

  // 3️⃣ hapus dari workingRows
  workingRows = workingRows.filter(r => String(r.IDPEL) !== idpel);
  saveState();

  // 4️⃣ HAPUS MARKER DIJ + DIL DENGAN IDPEL YANG SAMA
  markers = markers.filter(obj => {
    if (String(obj.row.IDPEL) === idpel) {
      if (map.hasLayer(obj.marker)) {
        clusterGroup.removeLayer(obj.marker);
      }
      return false; // buang dari array
    }
    return true;
  });

  // 5️⃣ update UI
  renderHistory();
  closeDetail();

  showToast("✅ Pelanggan ditandai selesai");
}

/* ======================
   DONE + HISTORY
====================== */
function renderHistory() {
  const el = document.getElementById("historyList");
  if (!el) return;

  if (!history || history.length === 0) {
    el.innerHTML = `
      <div style="padding:10px;color:#666;text-align:center">
        📭 Belum ada pelanggan selesai hari ini
      </div>`;
    return;
  }

  el.innerHTML = history.map((item, index) => `

  <div style="
    border-bottom:1px solid #eee;
    padding:8px;
    display:flex;
    justify-content:space-between;
    align-items:center;
  ">

    <div>

      <strong>${item.row.NAMA || "-"}</strong><br>

      NOMET: ${getNoMeter(item.row)}<br>

      MERK: ${getMerkMeter(item.row)}<br>

      HARI BACA: ${getHariBaca(item.row)}<br>

      <small>${item.time}</small>

    </div>

    <button
      onclick="undoHistory(${index})"
      style="
        background:#2563eb;
        color:white;
        border:none;
        padding:6px 10px;
        border-radius:8px;
        font-size:12px;
      "
    >
      Undo
    </button>

  </div>

`).join("");
}


function undoHistory(index) {
  const item = history[index];
  if (!item) return;

  // keluar dari compare mode
  if (window.isCompareMode) {
    const cb = document.getElementById("compareCheckbox");
    if (cb) cb.checked = false;
    exitCompareMode();
  }

  // kembalikan data
  workingRows.push(item.row);
  saveState();

  // hapus dari history
  history.splice(index, 1);
  localStorage.setItem("oyon_history", JSON.stringify(history));

  // 🔥 REBUILD DULU (INI KUNCI)
  buildMarkers();

  // 🔁 BARU SET FILTER HARI BACA
  const hariUndo = getHariBaca(item.row);
  const idx = daftarHariBaca.indexOf(hariUndo);

  if (idx !== -1) {
    filterHariBacaAktif = true;
    indexHariBaca = idx;
  }

  // 🔥 APPLY SETELAH STATE SIAP
  applyHariBacaFilter();

  renderHistory();
  showToast(`↩️ Dikembalikan ke HARI BACA ${hariUndo}`);
}





/* ======================
   TAMPILAN KEMBALI KE AREA TAGGING
====================== */

function saveMapView() {
  const center = map.getCenter();
  const zoom = map.getZoom();

  localStorage.setItem("RBM_MAP_VIEW", JSON.stringify({
    lat: center.lat,
    lng: center.lng,
    zoom: zoom
  }));
}

function restoreMapView() {
  const savedView = localStorage.getItem("RBM_MAP_VIEW");

  if (savedView) {
    const view = JSON.parse(savedView);
    map.setView([view.lat, view.lng], view.zoom);
  } else {
    zoomToAllMarkers(); // fallback
  }
}


/* ======================
   reset history 23.00
====================== */



function copyNomet(nomet){
  if (!nomet) return;

  navigator.clipboard.writeText(nomet)
    .then(() => showToast("📋 Nomor meter disalin"))
    .catch(() => showToast("❌ Gagal menyalin"));
}

function copyIDPEL(idpel){

  if (!idpel) return;

  navigator.clipboard.writeText(idpel)
    .then(() => showToast("📋 IDPEL disalin"))
    .catch(() => showToast("❌ Gagal menyalin IDPEL"));
}

/* ======================
   TOAST
====================== */
function showToast(msg){
  const t = document.getElementById("toast");
  t.innerHTML = msg;
  t.classList.add("show");
  setTimeout(()=>t.classList.remove("show"),5000);
}
/* ======================
   FINISH DAN SIMPAN RBM UPDATE
====================== */
downloadBtn.onclick = () => {
  if (workingRows.length === 0) {
    showToast("⚠️ Tidak ada data untuk diunduh");
    return;
  }

  // Ambil KODE PETUGAS dari baris pertama
  const kodePetugas = (workingRows[0]["KODE PETUGAS"] || "UNKNOWN")
    .toString()
    .replace(/\s+/g, "_");

  // Buat timestamp
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, "0");
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const min = String(now.getMinutes()).padStart(2, "0");

  const filename = `RBM_${kodePetugas}_${dd}-${mm}_${hh}-${min}.xlsx`;

  // ===== CLEAN EXPORT ENGINE =====

const cleanRows = workingRows.map(r => ({

  NO: r.NO || "",

  IDPEL: r.IDPEL || "",

  NAMA: r.NAMA || "",

  NIK: r.NIK || "",

  KDDK:
    r.KDDK ||
    r["KDDK ACMT"] ||
    "",

  MEREKKWH:
    r.MEREKKWH ||
    r.MERK ||
    "",

  NOMORKWH:
    r.NOMORKWH ||
    r["NOMOR METER"] ||
    "",

  DAYA: r.DAYA || "",

  ALAMAT: r.ALAMAT || "",

  "LAT DIJ": getDIJLat(r),

  "LON DIJ": getDIJLon(r),

  "LAT DIL": getDILLat(r),

  "LON DIL": getDILLon(r)

}));

// Generate Excel
const ws =
  XLSX.utils.json_to_sheet(cleanRows);

const wb =
  XLSX.utils.book_new();

wb.Props = {
  Title: "OYONID MAP EXPORT"
};

XLSX.utils.book_append_sheet(
  wb,
  ws,
  "RBM Update"
);

XLSX.writeFile(wb, filename, {
  compression: true
});
};


/* ======================
   GPS
====================== */
navigator.geolocation?.watchPosition(pos => {
  userLatLng = [pos.coords.latitude, pos.coords.longitude];

  const icon = L.divIcon({
    html: `
      <div class="petugas-marker">
        <img src="https://cdn-icons-png.flaticon.com/512/4140/4140048.png" />
      </div>
    `,
    iconSize: [38,38],
    iconAnchor: [19,19],
    className: ""
  });

  if (!petugasMarker) {
    petugasMarker = L.marker(userLatLng, { icon }).addTo(map);
    petugasMarker.bindTooltip("📍 Lokasi Anda", {
      permanent: false,
      direction: "top"
    });
  } else {
    petugasMarker.setLatLng(userLatLng);
  }

}, err => {
  console.warn("GPS error:", err);
}, {
  enableHighAccuracy: true,
  maximumAge: 1000,
  timeout: 10000
});

/* ======================
   COLOR MAP
====================== */
/* ======================
   COLOR MAP — A–Z UNIK
====================== */
function getColor(v){
  const colors = [
    "#22c55e", // A
    "#f97316", // B
    "#3b82f6", // C
    "#ef4444", // D
    "#a855f7", // E
    "#14b8a6", // F
    "#eab308", // G
    "#ec4899", // H
    "#6366f1", // I
    "#84cc16", // J
    "#06b6d4", // K
    "#f43f5e", // L
    "#7c3aed", // M
    "#f59e0b", // N
    "#0ea5e9", // O
    "#10b981", // P
    "#db2777", // Q
    "#65a30d", // R
    "#0284c7", // S
    "#dc2626", // T
    "#9333ea", // U
    "#059669", // V
    "#ca8a04", // W
    "#be123c", // X
    "#2563eb", // Y
    "#0891b2"  // Z
  ];

  if (!v) return "#64748b";

  const code = v.toUpperCase().charCodeAt(0) - 65;
  return colors[code] || "#64748b";
}


renderHistory();
window.addEventListener("DOMContentLoaded", () => {

  function openHistory(){
  const panel = document.getElementById("historyPanel");
  panel.style.display = "block";
  renderHistory();
}

function closeHistory(){
  const panel = document.getElementById("historyPanel");
  panel.style.display = "none";
}

document.getElementById("historyToggleBtn").onclick = openHistory;


});
/* ======================
   minimize btn
====================== */
window.addEventListener("DOMContentLoaded", () => {
  const uiBar = document.getElementById("uiBar");
  const toggleBtn = document.getElementById("toggleBtn");

  toggleBtn.onclick = () => {
    uiBar.classList.toggle("minimized");
    toggleBtn.innerText = uiBar.classList.contains("minimized") ? "➕" : "➖";
  };

  document.getElementById("fullMapBtn").onclick = toggleFullMap;
});

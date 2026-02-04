const tooltip = document.getElementById("tooltip");

// Detect mobile device
const isMobile = () => window.innerWidth <= 768 || 'ontouchstart' in window;

// Set initial map position based on device
const getInitialMapSettings = () => {
  if (isMobile()) {
    return { zoom: 0.32, pan: { x: 30, y: 10 } };
  }
  return { zoom: 0.70, pan: { x: 260, y: 25 } };
};

const initialMapSettings = getInitialMapSettings();

const state = {
  rawRows: [],
  partyList: [],
  districtWinners: [],
  parties: [],
  selectedParties: new Set(),
  pinnedPartyId: null,
  includePartyList: true,
  activeDistrict: null,
  provinceLayout: null,
  editedProvincePositions: new Map(),
  mapZoom: initialMapSettings.zoom,
  mapPan: { ...initialMapSettings.pan },
  isDragging: false,
  dragStart: { x: 0, y: 0 },
  draggingProvince: null,
  provinceDragStart: { x: 0, y: 0 },
};

const elements = {
  countedTotal: document.getElementById("countedTotal"),
  partyListSeatsTotal: document.getElementById("partyListSeatsTotal"),
  partyOverview: document.getElementById("partyOverview"),
  topPartyList: document.getElementById("topPartyList"),
  provinceFilter: document.getElementById("provinceFilter"),
  partyFilter: document.getElementById("partyFilter"),
  districtSearch: document.getElementById("districtSearch"),
  mapMarkers: document.getElementById("mapMarkers"),
  top3List: document.getElementById("top3List"),
  mapViewport: document.getElementById("mapViewport"),
  mapContent: document.getElementById("mapContent"),
  zoomIn: document.getElementById("zoomIn"),
  zoomOut: document.getElementById("zoomOut"),
  partyListTable: document.getElementById("partyListTable"),
  partyListInputTable: document.getElementById("partyListInputTable"),
  seatSplit: document.getElementById("seatSplit"),
  coalitionList: document.getElementById("coalitionList"),
  coalitionTotal: document.getElementById("coalitionTotal"),
  coalitionStatus: document.getElementById("coalitionStatus"),
  coalitionTags: document.getElementById("coalitionTags"),
  includePartyList: document.getElementById("includePartyList"),
  inputTable: document.getElementById("inputTable"),
  downloadTemplate: document.getElementById("downloadTemplate"),
  downloadPartyListTemplate: document.getElementById("downloadPartyListTemplate"),
  csvUpload: document.getElementById("csvUpload"),
  partyListUpload: document.getElementById("partyListUpload"),
  applyCsv: document.getElementById("applyCsv"),
  addPartyRow: document.getElementById("addPartyRow"),
  exportCoordinates: document.getElementById("exportCoordinates"),
  partyVotesList: document.getElementById("partyVotesList"),
  top3Title: document.getElementById("top3Title"),
  importFromSheet: document.getElementById("importFromSheet"),
  importPartyListFromSheet: document.getElementById("importPartyListFromSheet"),
};

const formatNumber = (value) =>
  new Intl.NumberFormat("th-TH").format(value || 0);

let tooltipTimeout = null;

const updateTooltip = (event, html) => {
  // Clear any pending auto-hide
  if (tooltipTimeout) {
    clearTimeout(tooltipTimeout);
    tooltipTimeout = null;
  }
  
  // Handle both mouse and touch events
  const x = event.touches ? event.touches[0].clientX : event.clientX;
  const y = event.touches ? event.touches[0].clientY : event.clientY;
  
  tooltip.innerHTML = html;
  tooltip.style.left = `${x + 12}px`;
  tooltip.style.top = `${y + 12}px`;
  tooltip.classList.add("show");
  
  // Auto-hide after 3 seconds on touch devices
  if ("ontouchstart" in window) {
    tooltipTimeout = setTimeout(hideTooltip, 3000);
  }
};

const hideTooltip = () => {
  tooltip.classList.remove("show");
  if (tooltipTimeout) {
    clearTimeout(tooltipTimeout);
    tooltipTimeout = null;
  }
};

// Hide tooltip when tapping anywhere else (for mobile)
document.addEventListener("touchstart", (event) => {
  const target = event.target;
  const isDistrictBox = target.closest(".map-province__district");
  const isPartyRow = target.closest(".party-seat-row, .party-list tbody tr, .overview-table tbody tr");
  
  if (!isDistrictBox && !isPartyRow) {
    hideTooltip();
  }
});

// Party Winners Modal functions
const showPartyWinnersModal = (partyName, partyColor) => {
  const modal = document.getElementById("partyWinnersModal");
  const titleEl = modal.querySelector(".modal-title");
  const countEl = modal.querySelector(".modal-count");
  const colorEl = modal.querySelector(".modal-party-color");
  const listEl = document.getElementById("modalWinnersList");
  
  // Filter winners by party and sort by province, then district
  const winners = state.districtWinners
    .filter((w) => w.party === partyName)
    .sort((a, b) => {
      const provinceCompare = a.province.localeCompare(b.province, "th");
      if (provinceCompare !== 0) return provinceCompare;
      return a.district - b.district;
    });
  
  // Set party color as CSS variable
  modal.style.setProperty("--modal-accent", partyColor);
  colorEl.style.background = partyColor;
  
  // Update header
  titleEl.textContent = `ผู้ชนะจากพรรค${partyName}`;
  countEl.textContent = `${winners.length} เขต`;
  
  // Populate table
  listEl.innerHTML = winners
    .map((w) => `
      <tr>
        <td>${w.name}</td>
        <td>${w.province}</td>
        <td>${w.district}</td>
        <td>${formatNumber(w.votes)}</td>
      </tr>
    `)
    .join("");
  
  // Show modal
  modal.classList.remove("hidden");
  
  // Prevent body scroll
  document.body.style.overflow = "hidden";
};

const hidePartyWinnersModal = () => {
  const modal = document.getElementById("partyWinnersModal");
  modal.classList.add("hidden");
  document.body.style.overflow = "";
};

// Modal event listeners
document.addEventListener("DOMContentLoaded", () => {
  const modal = document.getElementById("partyWinnersModal");
  const closeBtn = document.getElementById("modalClose");
  
  if (closeBtn) {
    closeBtn.addEventListener("click", hidePartyWinnersModal);
  }
  
  if (modal) {
    // Close on overlay click
    modal.addEventListener("click", (event) => {
      if (event.target === modal) {
        hidePartyWinnersModal();
      }
    });
  }
  
  // Close on Escape key
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      hidePartyWinnersModal();
    }
  });
});

const palette = [
  "#f38b00",
  "#d92d27",
  "#1aa260",
  "#2f6fed",
  "#2e7be5",
  "#8b5cf6",
  "#0ea5e9",
  "#f97316",
  "#22c55e",
  "#64748b",
];

// Party colors from official party logos
const partyColors = {
  "ไทยทรัพย์ทวี": "#D91C24",
  "เพื่อชาติไทย": "#C9A227",
  "ใหม่": "#0B2A5B",
  "มิติใหม่": "#2E7D7A",
  "รวมใจไทย": "#D32F2F",
  "รวมไทยสร้างชาติ": "#1B2A8A",
  "พลวัต": "#2DBE60",
  "ประชาธิปไตยใหม่": "#F26C1A",
  "เพื่อไทย": "#E30613",
  "ทางเลือกใหม่": "#1E66D0",
  "เศรษฐกิจ": "#D1A000",
  "เสรีรวมไทย": "#D6B11E",
  "รวมพลังประชาชน": "#F28C28",
  "ท้องที่ไทย": "#2E8B57",
  "อนาคตไทย": "#D9282A",
  "พลังเพื่อไทย": "#E21D24",
  "ไทยชนะ": "#1D3C8F",
  "พลังสังคมใหม่": "#8B1E2D",
  "สังคมประชาธิปไตยไทย": "#1E57A5",
  "ฟิวชัน": "#1B7F3A",
  "ไทรวมพลัง": "#A000B5",
  "ก้าวอิสระ": "#6B4EFF",
  "ปวงชนไทย": "#E53935",
  "วิชชั่นใหม่": "#1D4ED8",
  "เพื่อชีวิตใหม่": "#D4AF37",
  "คลองไทย": "#005BBB",
  "ประชาธิปัตย์": "#1E88E5",
  "ไทยก้าวหน้า": "#1E3A8A",
  "ไทยภักดี": "#2E7D32",
  "แรงงานสร้างชาติ": "#6A1B9A",
  "ประชากรไทย": "#1E4FA1",
  "ครูไทยเพื่อประชาชน": "#F57C00",
  "ประชาชาติ": "#C6A200",
  "สร้างอนาคตไทย": "#1E4FA1",
  "รักชาติ": "#006B3C",
  "ไทยพร้อม": "#1E4FA1",
  "ภูมิใจไทย": "#312682",
  "พลังธรรมใหม่": "#1E3A8A",
  "กรีน": "#2E8B57",
  "ไทยธรรม": "#5B2B90",
  "แผ่นดินธรรม": "#6D4C41",
  "กล้าธรรม": "#90EE90",
  "พลังประชารัฐ": "#0B6B3A",
  "โอกาสใหม่": "#FF69B4",
  "เป็นธรรม": "#1E57A5",
  "ประชาชน": "#F57A36",
  "ประชาไทย": "#1E57A5",
  "ไทยสร้างไทย": "#6F42C1",
  "ไทยก้าวใหม่": "#D4E000",
  "ประชาอาสาชาติ": "#E53935",
  "พร้อม": "#E11D2E",
  "เครือข่ายชาวนาแห่งประเทศไทย": "#0B4FA3",
  "ไทยพิทักษ์ธรรม": "#1E3A8A",
  "ความหวังใหม่": "#F2C300",
  "ไทยรวมไทย": "#1E4FA1",
  "เพื่อบ้านเมือง": "#F5D000",
  "พลังไทยรักชาติ": "#00A0B0",
  "ก้าวไกล": "#F97316",
  "ชาติไทยพัฒนา": "#4CAF50",
  "ชาติพัฒนากล้า": "#8BC34A",
};

// Pixel-based layout matching reference image
const defaultProvinceLayout = {
  provinces: [
    { name: "แม่ฮ่องสอน", x: -13, y: -2, columns: 2 },
    { name: "เชียงราย", x: 66, y: -22, columns: 4 },
    { name: "เชียงใหม่", x: -50, y: 47, columns: 5 },
    { name: "พะเยา", x: 101, y: 35, columns: 3 },
    { name: "น่าน", x: 131, y: 77, columns: 3 },
    { name: "ลำปาง", x: 48, y: 64, columns: 4 },
    { name: "ลำพูน", x: -2, y: 110, columns: 2 },
    { name: "แพร่", x: 124, y: 115, columns: 3 },
    { name: "อุตรดิตถ์", x: 147, y: 156, columns: 3 },
    { name: "ตาก", x: -23, y: 147, columns: 3 },
    { name: "สุโขทัย", x: 69, y: 126, columns: 3 },
    { name: "พิษณุโลก", x: 98, y: 191, columns: 4 },
    { name: "เพชรบูรณ์", x: 198, y: 196, columns: 6 },
    { name: "กำแพงเพชร", x: 3, y: 199, columns: 4 },
    { name: "พิจิตร", x: 172, y: 232, columns: 3 },
    { name: "นครสวรรค์", x: 94, y: 250, columns: 5 },
    { name: "อุทัยธานี", x: 40, y: 238, columns: 2 },
    { name: "ชัยนาท", x: 54, y: 296, columns: 2 },
    { name: "สิงห์บุรี", x: 117, y: 304, columns: 1 },
    { name: "ลพบุรี", x: 190, y: 271, columns: 3 },
    { name: "สระบุรี", x: 232, y: 307, columns: 3 },
    { name: "สุพรรณบุรี", x: 7, y: 337, columns: 5 },
    { name: "อ่างทอง", x: 156, y: 329, columns: 2 },
    { name: "พระนครศรีอยุธยา", x: 75, y: 378, columns: 4 },
    { name: "กาญจนบุรี", x: -24, y: 383, columns: 5 },
    { name: "นครปฐม", x: 93, y: 418, columns: 3 },
    { name: "ปทุมธานี", x: 199, y: 361, columns: 4 },
    { name: "นนทบุรี", x: 152, y: 413, columns: 4 },
    { name: "ราชบุรี", x: 10, y: 465, columns: 5 },
    { name: "สมุทรสาคร", x: 100, y: 465, columns: 3 },
    { name: "สมุทรปราการ", x: 178, y: 464, columns: 5 },
    { name: "สมุทรสงคราม", x: 100, y: 515, columns: 1 },
    { name: "กรุงเทพมหานคร", x: 200, y: 573, columns: 7 },
    { name: "เพชรบุรี", x: 42, y: 540, columns: 3 },
    { name: "ประจวบคีรีขันธ์", x: 62, y: 593, columns: 3 },
    { name: "นครนายก", x: 291, y: 361, columns: 2 },
    { name: "ปราจีนบุรี", x: 303, y: 403, columns: 3 },
    { name: "สระแก้ว", x: 368, y: 400, columns: 3 },
    { name: "ฉะเชิงเทรา", x: 233, y: 422, columns: 4 },
    { name: "ชลบุรี", x: 274, y: 461, columns: 5 },
    { name: "ระยอง", x: 358, y: 506, columns: 4 },
    { name: "จันทบุรี", x: 359, y: 460, columns: 3 },
    { name: "ตราด", x: 430, y: 515, columns: 1 },
    { name: "หนองคาย", x: 362, y: 11, columns: 3 },
    { name: "บึงกาฬ", x: 425, y: 10, columns: 3 },
    { name: "อุดรธานี", x: 325, y: 58, columns: 4 },
    { name: "หนองบัวลำภู", x: 275, y: 127, columns: 3 },
    { name: "เลย", x: 221, y: 143, columns: 3 },
    { name: "สกลนคร", x: 407, y: 62, columns: 4 },
    { name: "นครพนม", x: 479, y: 64, columns: 4 },
    { name: "มุกดาหาร", x: 474, y: 111, columns: 2 },
    { name: "ขอนแก่น", x: 287, y: 169, columns: 6 },
    { name: "กาฬสินธุ์", x: 362, y: 129, columns: 4 },
    { name: "ร้อยเอ็ด", x: 404, y: 191, columns: 5 },
    { name: "มหาสารคาม", x: 336, y: 231, columns: 4 },
    { name: "ยโสธร", x: 440, y: 144, columns: 3 },
    { name: "อำนาจเจริญ", x: 486, y: 192, columns: 2 },
    { name: "ชัยภูมิ", x: 250, y: 232, columns: 5 },
    { name: "นครราชสีมา", x: 298, y: 286, columns: 8 },
    { name: "อุบลราชธานี", x: 483, y: 238, columns: 6 },
    { name: "บุรีรัมย์", x: 369, y: 345, columns: 5 },
    { name: "สุรินทร์", x: 420, y: 260, columns: 4 },
    { name: "ศรีสะเกษ", x: 445, y: 313, columns: 5 },
    { name: "ชุมพร", x: 44, y: 656, columns: 3 },
    { name: "ระนอง", x: 10, y: 695, columns: 2 },
    { name: "สุราษฎร์ธานี", x: 53, y: 730, columns: 4 },
    { name: "พังงา", x: 7, y: 776, columns: 2 },
    { name: "นครศรีธรรมราช", x: 49, y: 805, columns: 5 },
    { name: "ภูเก็ต", x: -36, y: 815, columns: 3 },
    { name: "กระบี่", x: -5, y: 886, columns: 3 },
    { name: "ตรัง", x: 26, y: 856, columns: 4 },
    { name: "พัทลุง", x: 85, y: 885, columns: 3 },
    { name: "สงขลา", x: 85, y: 940, columns: 5 },
    { name: "ปัตตานี", x: 156, y: 996, columns: 3 },
    { name: "สตูล", x: 39, y: 932, columns: 2 },
    { name: "ยะลา", x: 85, y: 1005, columns: 3 },
    { name: "นราธิวาส", x: 189, y: 1066, columns: 5 },
  ],
};

const setMapZoom = (value) => {
  const minZoom = isMobile() ? 0.3 : 0.4;
  const zoom = Math.min(2.5, Math.max(minZoom, value));
  state.mapZoom = zoom;
  updateMapTransform();
};

const updateMapTransform = () => {
  if (elements.mapContent) {
    elements.mapContent.style.transform = `translate(${state.mapPan.x}px, ${state.mapPan.y}px) scale(${state.mapZoom})`;
  }
};

const initMapDrag = () => {
  if (!elements.mapViewport) return;
  
  elements.mapViewport.addEventListener("mousedown", (event) => {
    if (event.button !== 0) return;
    // Don't start map drag if we're dragging a province
    if (state.draggingProvince) return;
    state.isDragging = true;
    state.dragStart = { x: event.clientX - state.mapPan.x, y: event.clientY - state.mapPan.y };
    elements.mapViewport.style.cursor = "grabbing";
    event.preventDefault();
  });
  
  document.addEventListener("mousemove", (event) => {
    // Handle province dragging
    if (state.draggingProvince) {
      const newX = (event.clientX - state.provinceDragStart.x) / state.mapZoom;
      const newY = (event.clientY - state.provinceDragStart.y) / state.mapZoom;
      state.draggingProvince.style.left = `${Math.round(newX)}px`;
      state.draggingProvince.style.top = `${Math.round(newY)}px`;
      return;
    }
    // Handle map panning
    if (!state.isDragging) return;
    state.mapPan.x = event.clientX - state.dragStart.x;
    state.mapPan.y = event.clientY - state.dragStart.y;
    updateMapTransform();
  });
  
  document.addEventListener("mouseup", (event) => {
    // Handle province drag end
    if (state.draggingProvince) {
      const provinceName = state.draggingProvince.dataset.province;
      const newX = Math.round((event.clientX - state.provinceDragStart.x) / state.mapZoom);
      const newY = Math.round((event.clientY - state.provinceDragStart.y) / state.mapZoom);
      const columns = parseInt(state.draggingProvince.dataset.columns, 10) || 3;
      state.editedProvincePositions.set(provinceName, { x: newX, y: newY, columns });
      state.draggingProvince.classList.remove("dragging");
      state.draggingProvince = null;
      return;
    }
    // Handle map drag end
    if (state.isDragging) {
      state.isDragging = false;
      if (elements.mapViewport) {
        elements.mapViewport.style.cursor = "grab";
      }
    }
  });

  // Touch support for mobile devices
  const touchInstruction = document.getElementById("touchInstruction");
  let touchState = {
    lastTouchDistance: 0,
    lastTouchCenter: { x: 0, y: 0 },
    isTwoFingerGesture: false,
    holdTimer: null,
  };

  const getTouchDistance = (touches) => {
    const dx = touches[0].clientX - touches[1].clientX;
    const dy = touches[0].clientY - touches[1].clientY;
    return Math.sqrt(dx * dx + dy * dy);
  };

  const getTouchCenter = (touches) => {
    return {
      x: (touches[0].clientX + touches[1].clientX) / 2,
      y: (touches[0].clientY + touches[1].clientY) / 2,
    };
  };

  const showTouchInstruction = () => {
    if (touchInstruction) {
      touchInstruction.classList.remove("hidden");
    }
  };

  const hideTouchInstruction = () => {
    if (touchInstruction) {
      touchInstruction.classList.add("hidden");
    }
  };

  elements.mapViewport.addEventListener("touchstart", (event) => {
    // Clear any existing hold timer
    if (touchState.holdTimer) {
      clearTimeout(touchState.holdTimer);
      touchState.holdTimer = null;
    }

    if (event.touches.length === 1) {
      // Single finger - start hold timer for instruction popup
      touchState.isTwoFingerGesture = false;
      touchState.holdTimer = setTimeout(() => {
        showTouchInstruction();
      }, 300);
    } else if (event.touches.length === 2) {
      // Two fingers - start pan/zoom gesture
      event.preventDefault();
      hideTouchInstruction();
      touchState.isTwoFingerGesture = true;
      touchState.lastTouchDistance = getTouchDistance(event.touches);
      touchState.lastTouchCenter = getTouchCenter(event.touches);
    }
  }, { passive: false });

  elements.mapViewport.addEventListener("touchmove", (event) => {
    // Clear hold timer on any movement
    if (touchState.holdTimer) {
      clearTimeout(touchState.holdTimer);
      touchState.holdTimer = null;
    }

    if (event.touches.length === 2) {
      event.preventDefault();
      hideTouchInstruction();
      touchState.isTwoFingerGesture = true;

      // Calculate pinch zoom
      const currentDistance = getTouchDistance(event.touches);
      const currentCenter = getTouchCenter(event.touches);

      if (touchState.lastTouchDistance > 0) {
        // Zoom based on pinch
        const scale = currentDistance / touchState.lastTouchDistance;
        const minZoom = isMobile() ? 0.3 : 0.4;
        const newZoom = Math.min(2.5, Math.max(minZoom, state.mapZoom * scale));
        
        // Get viewport-relative center for zoom
        const rect = elements.mapViewport.getBoundingClientRect();
        const centerX = currentCenter.x - rect.left;
        const centerY = currentCenter.y - rect.top;
        
        // Calculate map point under pinch center
        const mapX = (centerX - state.mapPan.x) / state.mapZoom;
        const mapY = (centerY - state.mapPan.y) / state.mapZoom;
        
        // Update zoom
        state.mapZoom = newZoom;
        
        // Adjust pan to keep pinch center fixed
        state.mapPan.x = centerX - mapX * newZoom;
        state.mapPan.y = centerY - mapY * newZoom;
      }

      // Pan based on center movement
      if (touchState.lastTouchCenter.x !== 0) {
        const dx = currentCenter.x - touchState.lastTouchCenter.x;
        const dy = currentCenter.y - touchState.lastTouchCenter.y;
        state.mapPan.x += dx;
        state.mapPan.y += dy;
      }

      touchState.lastTouchDistance = currentDistance;
      touchState.lastTouchCenter = currentCenter;
      updateMapTransform();
    }
  }, { passive: false });

  elements.mapViewport.addEventListener("touchend", (event) => {
    // Clear hold timer
    if (touchState.holdTimer) {
      clearTimeout(touchState.holdTimer);
      touchState.holdTimer = null;
    }

    // Hide instruction popup
    hideTouchInstruction();

    // Reset touch state when all fingers lifted
    if (event.touches.length === 0) {
      touchState.lastTouchDistance = 0;
      touchState.lastTouchCenter = { x: 0, y: 0 };
      touchState.isTwoFingerGesture = false;
    } else if (event.touches.length === 1) {
      // Transitioning from 2 fingers to 1
      touchState.lastTouchDistance = 0;
      touchState.lastTouchCenter = { x: 0, y: 0 };
    }
  });

  elements.mapViewport.addEventListener("touchcancel", () => {
    // Clear hold timer
    if (touchState.holdTimer) {
      clearTimeout(touchState.holdTimer);
      touchState.holdTimer = null;
    }
    hideTouchInstruction();
    touchState.lastTouchDistance = 0;
    touchState.lastTouchCenter = { x: 0, y: 0 };
    touchState.isTwoFingerGesture = false;
  });
};

const getProvinceList = (rows) => {
  const set = new Set();
  rows.forEach((row) => {
    const normalized = normalizeRow(row);
    if (normalized) {
      set.add(normalized.province);
    }
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, "th"));
};

const normalizeLayout = (layout, rows) => {
  if (layout && Array.isArray(layout.provinces) && layout.provinces.length) {
    return layout;
  }

  if (layout && Array.isArray(layout.regions) && layout.regions.length) {
    const provinces = [];
    layout.regions.forEach((region) => {
      (region.provinces || []).forEach((province, index) => {
        const colOffset = index % (region.columns || 3);
        const rowOffset = Math.floor(index / (region.columns || 3));
        const x = (region.startX ?? 10) + colOffset * (region.xGap || 6);
        const y = (region.startY ?? 10) + rowOffset * (region.yGap || 7);
        provinces.push({
          name: province,
          x,
          y,
          columns: region.columns || 3,
        });
      });
    });
    if (provinces.length) {
      return { provinces };
    }
  }

  const provinces = getProvinceList(rows);
  const columns = 6;
  const xGap = 14;
  const yGap = 10;
  const generated = provinces.map((province, index) => ({
    name: province,
    x: 6 + (index % columns) * xGap,
    y: 6 + Math.floor(index / columns) * yGap,
    columns: 3,
  }));
  return { provinces: generated };
};

const normalizeRow = (row) => {
  const region = `${row.region ?? ""}`.trim();
  const province = `${row.province ?? ""}`.trim();
  const name = `${row.name ?? ""}`.trim();
  const party = `${row.party ?? ""}`.trim();
  const district = Number(`${row.district ?? ""}`.trim());
  const votes = Number(`${row.votes ?? ""}`.trim());
  const counted = `${row.counted ?? ""}`.trim();
  if (!province || Number.isNaN(district)) {
    return null;
  }
  return {
    region,
    province,
    district,
    name,
    party,
    votes: Number.isNaN(votes) ? 0 : votes,
    counted,
  };
};

// Loading overlay helpers
const updateLoadingScreen = (status, progress, substatus = "") => {
  const overlay = document.getElementById("loadingOverlay");
  const statusEl = document.getElementById("loadingStatus");
  const progressBar = document.getElementById("loadingProgressBar");
  const substatusEl = document.getElementById("loadingSubstatus");
  
  if (statusEl) statusEl.textContent = status;
  if (progressBar) progressBar.style.width = `${progress}%`;
  if (substatusEl) substatusEl.textContent = substatus;
};

const hideLoadingScreen = (delay = 500) => {
  const overlay = document.getElementById("loadingOverlay");
  if (overlay) {
    setTimeout(() => {
      overlay.classList.add("hidden");
    }, delay);
  }
};

const showLoadingError = (message) => {
  const overlay = document.getElementById("loadingOverlay");
  const titleEl = overlay?.querySelector(".loading-title");
  
  if (overlay) overlay.classList.add("error");
  if (titleEl) titleEl.textContent = "เกิดข้อผิดพลาด";
  updateLoadingScreen(message, 100, "กำลังใช้ข้อมูลสำรอง...");
};

const loadData = async () => {
  let data = null;
  let useGoogleSheets = true;
  
  // Load province layout first (needed for map)
  updateLoadingScreen("กำลังโหลดข้อมูลแผนที่...", 5);
  try {
    const layoutResponse = await fetch("data/province_layout.json");
    if (!layoutResponse.ok) {
      throw new Error("layout not found");
    }
    state.provinceLayout = await layoutResponse.json();
  } catch (error) {
    state.provinceLayout = defaultProvinceLayout;
  }
  
  // Try to auto-fetch from Google Sheets
  updateLoadingScreen("กำลังเชื่อมต่อ Google Sheets...", 15);
  
  try {
    // Fetch district data from Google Sheets
    updateLoadingScreen("กำลังดาวน์โหลดข้อมูลรายเขต...", 25);
    const cacheBuster = `&_t=${Date.now()}`;
    const districtUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${DISTRICT_GID}${cacheBuster}`;
    
    const districtText = await fetchWithProxy(districtUrl);
    const districtData = parseCsv(districtText);
    
    if (districtData.length > 0) {
      state.rawRows = districtData;
      updateLoadingScreen("โหลดข้อมูลรายเขตสำเร็จ", 50, `${districtData.length} แถว`);
    } else {
      throw new Error("ไม่พบข้อมูลรายเขต");
    }
    
    // Fetch party list data from Google Sheets
    updateLoadingScreen("กำลังดาวน์โหลดข้อมูลบัญชีรายชื่อ...", 60);
    const partyListUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${PARTY_LIST_GID}${cacheBuster}`;
    
    const partyListText = await fetchWithProxy(partyListUrl);
    const partyListData = parsePartyListCsv(partyListText);
    
    if (partyListData.length > 0) {
      state.partyList = partyListData;
      const totalSeats = partyListData.reduce((sum, p) => sum + p.seats, 0);
      updateLoadingScreen("โหลดข้อมูลบัญชีรายชื่อสำเร็จ", 80, `${partyListData.length} พรรค, ${totalSeats} ที่นั่ง`);
    }
    
    updateLoadingScreen("กำลังประมวลผลข้อมูล...", 90);
    
  } catch (error) {
    console.warn("Google Sheets fetch failed, using fallback:", error);
    showLoadingError(error.message || "ไม่สามารถเชื่อมต่อ Google Sheets ได้");
    
    // Fall back to local data
    try {
      const response = await fetch("data/results.json");
      if (response.ok) {
        data = await response.json();
      }
    } catch (e) {
      data = window.__RESULTS__ || { candidates: [], partyList: [] };
    }
    
    if (!data) {
      data = window.__RESULTS__ || { candidates: [], partyList: [] };
    }
    
    state.rawRows = (data.candidates || [])
      .map((row) => normalizeRow(row))
      .filter(Boolean);
    state.partyList = data.partyList || [];
  }
  
  // Initialize map and render
  updateLoadingScreen("กำลังแสดงผลหน้าจอ...", 95);
  updateMapTransform();
  initMapDrag();
  recalculate();
  
  // Hide loading screen
  updateLoadingScreen("โหลดเสร็จสมบูรณ์!", 100);
  hideLoadingScreen(800);
};

const recalculate = () => {
  state.districtWinners = calculateDistrictWinners(state.rawRows);
  state.parties = buildPartySummary(state.districtWinners, state.partyList);
  renderAll();
};

const calculateDistrictWinners = (rows) => {
  const grouped = new Map();
  rows
    .map((row) => normalizeRow(row))
    .filter(Boolean)
    .forEach((row) => {
      const key = `${row.province}|||${row.district}`;
      const current = grouped.get(key);
      if (!current || row.votes > current.votes) {
        grouped.set(key, row);
      }
    });
  // Only include districts where winner has votes > 0 (i.e., data has arrived)
  return Array.from(grouped.values())
    .filter((row) => row.votes > 0)
    .map((row) => ({
      ...row,
      status: "นับแล้ว",
    }));
};

const buildPartySummary = (winners, partyList) => {
  const map = new Map();
  winners.forEach((row) => {
    const entry = map.get(row.party) || {
      id: row.party.toLowerCase().replace(/\s+/g, "-"),
      name: row.party,
      color: partyColors[row.party] || null,
      districtSeats: 0,
      listSeats: 0,
      partyListVotes: 0,
    };
    entry.districtSeats += 1;
    map.set(row.party, entry);
  });

  partyList.forEach((row, index) => {
    const entry = map.get(row.party) || {
      id: row.party.toLowerCase().replace(/\s+/g, "-"),
      name: row.party,
      color: partyColors[row.party] || null,
      districtSeats: 0,
      listSeats: 0,
      partyListVotes: 0,
    };
    entry.partyListVotes = row.votes || 0;
    entry.listSeats = row.seats || 0;
    entry.color = partyColors[row.party] || row.color || palette[index % palette.length];
    map.set(row.party, entry);
  });

  return Array.from(map.values()).map((entry, index) => ({
    ...entry,
    color: partyColors[entry.name] || entry.color || palette[index % palette.length],
    totalSeats: entry.districtSeats + entry.listSeats,
  }));
};

const renderAll = () => {
  renderOverview();
  renderDistrictFilters();
  renderDistrictMap();
  renderPartyVotes();
  renderTop3();
  renderPartyList();
  renderCoalition();
  renderInputTable();
  renderPartyListInput();
};

const renderOverview = () => {
  const totalCounted = state.districtWinners.length;
  const totalDistricts = getTotalDistricts(state.rawRows);
  elements.countedTotal.textContent = formatNumber(totalCounted);
  
  // Update party list seats total
  const totalPartyListSeats = state.partyList.reduce((sum, p) => sum + (p.seats || 0), 0);
  if (elements.partyListSeatsTotal) {
    elements.partyListSeatsTotal.textContent = formatNumber(totalPartyListSeats);
  }

  const list = [...state.parties]
    .map((party) => ({
      ...party,
      displayTotal: state.includePartyList
        ? party.totalSeats
        : party.districtSeats,
    }))
    .sort((a, b) => b.displayTotal - a.displayTotal)
    .map((party) => {
      const row = document.createElement("div");
      row.className = "party-row";
      if (state.pinnedPartyId === party.id) {
        row.classList.add("pinned");
      }
      row.innerHTML = `
        <div class="party-tag">
          <span class="party-color" style="background:${party.color}"></span>
          <strong>${party.name}</strong>
        </div>
        <span>${party.districtSeats} เขต / ${party.listSeats} บัญชี</span>
        <span>${party.displayTotal} ที่นั่ง</span>
      `;
      row.addEventListener("mouseenter", (event) => {
        updateTooltip(
          event,
          `<strong>${party.name}</strong><br/>ส.ส.เขต: ${party.districtSeats}<br/>ส.ส.บัญชีรายชื่อ: ${party.listSeats}<br/>รวมที่นั่ง: ${party.displayTotal}`
        );
      });
      row.addEventListener("mousemove", (event) => {
        if (tooltip.classList.contains("show")) {
          updateTooltip(
            event,
            `<strong>${party.name}</strong><br/>ส.ส.เขต: ${party.districtSeats}<br/>ส.ส.บัญชีรายชื่อ: ${party.listSeats}<br/>รวมที่นั่ง: ${party.displayTotal}`
          );
        }
      });
      row.addEventListener("mouseleave", hideTooltip);
      row.addEventListener("click", () => {
        hideTooltip();
        showPartyWinnersModal(party.name, party.color);
      });
      return row;
    });

  elements.partyOverview.innerHTML = "";
  list.forEach((row) => elements.partyOverview.appendChild(row));

  elements.topPartyList.innerHTML = "";
  list.slice(0, 5).forEach((row, index) => {
    const li = document.createElement("li");
    li.innerHTML = `
      <span>อันดับ ${index + 1} ${row.querySelector("strong").textContent}</span>
      <strong>${row.querySelector("span:last-child").textContent}</strong>
    `;
    elements.topPartyList.appendChild(li);
  });
};

const renderDistrictFilters = () => {
  const provinces = [
    "ทุกจังหวัด",
    ...new Set(state.districtWinners.map((item) => item.province)),
  ];
  const parties = [
    "ทุกพรรค",
    ...new Set(state.districtWinners.map((item) => item.party)),
  ];

  elements.provinceFilter.innerHTML = provinces
    .map((value) => `<option value="${value}">${value}</option>`)
    .join("");
  elements.partyFilter.innerHTML = parties
    .map((value) => `<option value="${value}">${value}</option>`)
    .join("");
};

const renderPartyVotes = () => {
  // Calculate seat count per party from all district winners
  const partySeatsMap = new Map();
  state.districtWinners.forEach((winner) => {
    const current = partySeatsMap.get(winner.party) || { 
      party: winner.party, 
      seatCount: 0 
    };
    current.seatCount += 1;
    partySeatsMap.set(winner.party, current);
  });
  
  // Convert to array and sort from highest to lowest by seat count
  const partySeatsArray = Array.from(partySeatsMap.values())
    .sort((a, b) => b.seatCount - a.seatCount);
  
  // Get party colors
  const partyColorMap = getPartyColorMap(state.parties);
  
  // Find max seats for bar width calculation
  const maxSeats = partySeatsArray.length > 0 ? partySeatsArray[0].seatCount : 1;
  
  elements.partyVotesList.innerHTML = "";
  
  if (partySeatsArray.length === 0) {
    elements.partyVotesList.innerHTML = '<div class="muted">ไม่พบข้อมูล</div>';
    return;
  }
  
  partySeatsArray.forEach((item, index) => {
    const row = document.createElement("div");
    row.className = "party-seat-row";
    const color = partyColorMap.get(item.party) || partyColors[item.party] || "#64748b";
    const barWidth = Math.max(30, (item.seatCount / maxSeats) * 100);
    
    row.innerHTML = `
      <span class="rank">${index + 1}</span>
      <div class="bar-container">
        <div class="bar" style="background:${color}; width:${barWidth}%">
          <span class="party-name">${item.party}</span>
        </div>
        <span class="seat-count">${item.seatCount}</span>
      </div>
    `;
    row.addEventListener("mouseenter", (event) => {
      updateTooltip(
        event,
        `<strong>${item.party}</strong><br/>ส.ส.เขต: ${item.seatCount} ที่นั่ง`
      );
    });
    row.addEventListener("mousemove", (event) => {
      if (tooltip.classList.contains("show")) {
        updateTooltip(
          event,
          `<strong>${item.party}</strong><br/>ส.ส.เขต: ${item.seatCount} ที่นั่ง`
        );
      }
    });
    row.addEventListener("mouseleave", hideTooltip);
    row.addEventListener("click", () => {
      hideTooltip();
      showPartyWinnersModal(item.party, color);
    });
    row.style.cursor = "pointer";
    elements.partyVotesList.appendChild(row);
  });
  
  renderTop3();
};

const renderDistrictMap = () => {
  const normalizedLayout = normalizeLayout(state.provinceLayout, state.rawRows);
  if (!normalizedLayout.provinces.length) {
    return;
  }
  const activeProvince = elements.provinceFilter.value;
  const winnersMap = new Map(
    state.districtWinners.map((row) => [
      `${row.province}|||${row.district}`,
      row,
    ])
  );
  const totalsByProvince = getProvinceDistrictTotals(state.rawRows);
  const partyColorMap = getPartyColorMap(state.parties);

  elements.mapMarkers.innerHTML = "";

  (normalizedLayout.provinces || []).forEach((provinceDef) => {
    const totalDistricts =
      provinceDef.districts || totalsByProvince.get(provinceDef.name) || 0;

    // Check for edited position
    const editedPos = state.editedProvincePositions.get(provinceDef.name);
    const posX = editedPos ? editedPos.x : provinceDef.x;
    const posY = editedPos ? editedPos.y : provinceDef.y;

    const provinceEl = document.createElement("div");
    provinceEl.className = "map-province";
    provinceEl.dataset.province = provinceDef.name;
    provinceEl.dataset.columns = provinceDef.columns || 3;
    // Use pixel-based positioning
    provinceEl.style.left = `${posX}px`;
    provinceEl.style.top = `${posY}px`;
    provinceEl.innerHTML = `<div class="map-province__name">${provinceDef.name}</div>`;

    // Province dragging disabled
    // provinceEl.addEventListener("mousedown", (event) => {
    //   if (event.button !== 0) return;
    //   if (event.target.classList.contains("district-box")) return;
    //   event.stopPropagation();
    //   state.draggingProvince = provinceEl;
    //   const currentX = parseInt(provinceEl.style.left, 10) || 0;
    //   const currentY = parseInt(provinceEl.style.top, 10) || 0;
    //   state.provinceDragStart = {
    //     x: event.clientX - currentX * state.mapZoom,
    //     y: event.clientY - currentY * state.mapZoom,
    //   };
    //   provinceEl.classList.add("dragging");
    //   event.preventDefault();
    // });

    const grid = document.createElement("div");
    grid.className = "district-grid";
    grid.style.gridTemplateColumns = `repeat(${provinceDef.columns || 3}, 14px)`;

    for (let district = 1; district <= totalDistricts; district += 1) {
      const key = `${provinceDef.name}|||${district}`;
      const winner = winnersMap.get(key);
      const box = document.createElement("div");
      box.className = "district-box";
      box.textContent = district;
      if (!winner) {
        box.classList.add("no-data");
      } else {
        box.style.background = partyColorMap.get(winner.party) || "#64748b";
      }
      if (
        activeProvince === provinceDef.name &&
        state.activeDistrict === district
      ) {
        box.classList.add("active");
      }
      box.addEventListener("mouseenter", (event) => {
        if (state.isDragging) return;
        if (winner) {
          updateTooltip(
            event,
            `<strong>${provinceDef.name} เขต ${district}</strong><br/>${winner.name}<br/>${winner.party}<br/>คะแนน: ${formatNumber(
              winner.votes
            )}`
          );
        } else {
          updateTooltip(
            event,
            `<strong>${provinceDef.name} เขต ${district}</strong><br/>ยังไม่มีข้อมูล`
          );
        }
      });
      box.addEventListener("mousemove", (event) => {
        if (state.isDragging) return;
        if (!tooltip.classList.contains("show")) {
          return;
        }
        if (winner) {
          updateTooltip(
            event,
            `<strong>${provinceDef.name} เขต ${district}</strong><br/>${winner.name}<br/>${winner.party}<br/>คะแนน: ${formatNumber(
              winner.votes
            )}`
          );
        } else {
          updateTooltip(
            event,
            `<strong>${provinceDef.name} เขต ${district}</strong><br/>ยังไม่มีข้อมูล`
          );
        }
      });
      box.addEventListener("mouseleave", hideTooltip);
      box.addEventListener("click", (event) => {
        if (state.isDragging) return;
        event.stopPropagation();
        if (
          state.activeDistrict === district &&
          elements.provinceFilter.value === provinceDef.name
        ) {
          state.activeDistrict = null;
        } else {
          state.activeDistrict = district;
        }
        elements.provinceFilter.value = provinceDef.name;
        renderPartyVotes();
        renderDistrictMap();
        renderTop3();
      });
      grid.appendChild(box);
    }

    provinceEl.appendChild(grid);
    elements.mapMarkers.appendChild(provinceEl);
  });
};

const getProvinceDistrictTotals = (rows) => {
  const map = new Map();
  rows.forEach((row) => {
    const normalized = normalizeRow(row);
    if (!normalized) {
      return;
    }
    const current = map.get(normalized.province) || 0;
    if (normalized.district > current) {
      map.set(normalized.province, normalized.district);
    }
  });
  return map;
};

const getPartyColorMap = (parties) =>
  new Map(parties.map((party) => [party.name, party.color]));

const getProvincesByRegion = (rows) => {
  const map = new Map();
  rows.forEach((row) => {
    const normalized = normalizeRow(row);
    if (!normalized) {
      return;
    }
    const region = normalized.region || "กลาง";
    if (!map.has(region)) {
      map.set(region, new Set());
    }
    map.get(region).add(normalized.province);
  });
  return new Map(
    Array.from(map.entries()).map(([region, set]) => [region, Array.from(set)])
  );
};


const renderTop3 = () => {
  const province = elements.provinceFilter.value;
  const district = state.activeDistrict;
  
  if (!province || province === "ทุกจังหวัด" || !district) {
    if (elements.top3Title) elements.top3Title.textContent = "Top 3 คะแนน";
    elements.top3List.innerHTML = `<div class="muted">ยังไม่ได้เลือกเขต</div>`;
    return;
  }

  if (elements.top3Title) elements.top3Title.textContent = `Top 3 คะแนน (${province} เขต ${district})`;

  const candidates = state.rawRows
    .filter(
      (row) => row.province === province && Number(row.district) === district
    )
    .sort((a, b) => b.votes - a.votes);
  
  // Get the highest "counted" value for this district
  let countedValue = "";
  candidates.forEach((row) => {
    if (row.counted && row.counted.trim()) {
      // If we already have a value, compare and keep the higher one
      if (countedValue) {
        // Try to extract numbers for comparison (e.g., "4/20" -> compare first number)
        const currentMatch = countedValue.match(/(\d+)/);
        const newMatch = row.counted.match(/(\d+)/);
        if (currentMatch && newMatch) {
          if (parseInt(newMatch[1]) > parseInt(currentMatch[1])) {
            countedValue = row.counted.trim();
          }
        }
      } else {
        countedValue = row.counted.trim();
      }
    }
  });

  const top3 = candidates.slice(0, 3);

  if (!top3.length) {
    elements.top3List.innerHTML = `<div class="muted">ยังไม่มีข้อมูลสำหรับเขตนี้</div>`;
    return;
  }

  // Build the counted display HTML
  const countedHtml = countedValue 
    ? `<div class="counted-status">
        <span class="counted-label">หน่วยที่นับได้</span>
        <span class="counted-value">${countedValue}</span>
      </div>`
    : "";

  elements.top3List.innerHTML = countedHtml + top3
    .map(
      (candidate, index) => `
      <div class="top3-item">
        <span>#${index + 1} ${candidate.name} (${candidate.party})</span>
        <strong>${formatNumber(candidate.votes)}</strong>
      </div>
    `
    )
    .join("");
};

const renderPartyList = () => {
  elements.partyListTable.innerHTML = "";
  const sorted = [...state.parties].sort(
    (a, b) => b.listSeats - a.listSeats
  );
  sorted.forEach((party) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>
        <span class="party-tag">
          <span class="party-color" style="background:${party.color}"></span>
          ${party.name}
        </span>
      </td>
      <td>${party.listSeats}</td>
      <td>${party.totalSeats}</td>
    `;
    if (state.pinnedPartyId === party.id) {
      row.style.background = "#e9f2ff";
    }
    elements.partyListTable.appendChild(row);
  });

  elements.seatSplit.innerHTML = sorted
    .map(
      (party) => `
      <div class="party-row">
        <div class="party-tag">
          <span class="party-color" style="background:${party.color}"></span>
          <strong>${party.name}</strong>
        </div>
        <span>${party.districtSeats} / ${party.listSeats}</span>
        <span>${party.totalSeats}</span>
      </div>
    `
    )
    .join("");
};

const renderCoalition = () => {
  elements.coalitionList.innerHTML = "";
  const sorted = [...state.parties].sort((a, b) => b.totalSeats - a.totalSeats);
  sorted.forEach((party) => {
    const card = document.createElement("div");
    card.className = "party-card";
    if (state.selectedParties.has(party.id)) {
      card.classList.add("active");
    }
    card.innerHTML = `
      <span class="party-color" style="background:${party.color}"></span>
      <strong>${party.name}</strong>
      <span>${party.totalSeats} ที่นั่ง</span>
    `;
    card.addEventListener("click", () => {
      if (state.selectedParties.has(party.id)) {
        state.selectedParties.delete(party.id);
      } else {
        state.selectedParties.add(party.id);
      }
      updateCoalitionScore();
      renderCoalition();
    });
    elements.coalitionList.appendChild(card);
  });

  updateCoalitionScore();
};

const updateCoalitionScore = () => {
  const selected = state.parties
    .filter((party) => state.selectedParties.has(party.id))
    .sort((a, b) => b.totalSeats - a.totalSeats);
  const baseTotal = selected.reduce((sum, party) => sum + party.totalSeats, 0);
  const total = baseTotal;
  elements.coalitionTotal.textContent = formatNumber(total);
  elements.coalitionStatus.textContent =
    total >= 250 ? "เสียงเพียงพอจัดตั้งรัฐบาล" : `ขาดอีก ${250 - total} เสียง`;
  elements.coalitionTags.innerHTML = selected
    .map((party) => `<span>${party.name}</span>`)
    .join("");
};

const renderInputTable = () => {
  elements.inputTable.innerHTML = "";
  state.rawRows.forEach((item, index) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td contenteditable="true">${item.region}</td>
      <td contenteditable="true">${item.province}</td>
      <td contenteditable="true">${item.district}</td>
      <td contenteditable="true">${item.name}</td>
      <td contenteditable="true">${item.party}</td>
      <td contenteditable="true">${item.votes}</td>
    `;
    row.addEventListener("input", () => {
      const cells = row.querySelectorAll("td");
      const normalized = normalizeRow({
        region: cells[0].textContent.trim(),
        province: cells[1].textContent.trim(),
        district: cells[2].textContent.trim(),
        name: cells[3].textContent.trim(),
        party: cells[4].textContent.trim(),
        votes: cells[5].textContent.trim(),
      });
      if (normalized) {
        state.rawRows[index] = normalized;
      }
      recalculate();
    });
    elements.inputTable.appendChild(row);
  });
};

const renderPartyListInput = () => {
  elements.partyListInputTable.innerHTML = "";
  state.partyList.forEach((item, index) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td contenteditable="true">${item.party}</td>
      <td contenteditable="true">${item.seats}</td>
    `;
    row.addEventListener("input", () => {
      const cells = row.querySelectorAll("td");
      state.partyList[index] = {
        party: cells[0].textContent.trim(),
        votes: 0,
        seats: Number(cells[1].textContent.trim()),
        color: item.color,
      };
      recalculate();
    });
    elements.partyListInputTable.appendChild(row);
  });
};

const addPartyRow = () => {
  state.partyList.push({
    party: "พรรคใหม่",
    votes: 0,
    seats: 0,
    color: palette[state.partyList.length % palette.length],
  });
  renderPartyListInput();
  recalculate();
};

const buildProvinceSummary = (rows) => {
  const districtMap = new Map();
  rows
    .map((row) => normalizeRow(row))
    .filter(Boolean)
    .forEach((row) => {
      const key = `${row.province}|||${row.district}`;
      if (!districtMap.has(key)) {
        districtMap.set(key, row);
      }
    });

  const provinceSummary = new Map();
  districtMap.forEach((row) => {
    const existing = provinceSummary.get(row.province) || {
      region: row.region || "กลาง",
      province: row.province,
      total: 0,
      counted: 0,
    };
    existing.total += 1;
    existing.counted += 1;
    provinceSummary.set(row.province, existing);
  });

  return Array.from(provinceSummary.values()).sort((a, b) =>
    a.province.localeCompare(b.province, "th")
  );
};

const getTotalDistricts = (rows) => {
  const set = new Set();
  rows
    .map((row) => normalizeRow(row))
    .filter(Boolean)
    .forEach((row) => set.add(`${row.province}|||${row.district}`));
  return set.size;
};

const getRegionColumn = (region) => {
  const mapping = {
    เหนือ: 2,
    อีสาน: 4,
    กลาง: 3,
    ตะวันออก: 5,
    ตะวันตก: 2,
    ใต้: 3,
  };
  return mapping[region] || 3;
};

const getRegionRowOffset = (region) => {
  const mapping = {
    เหนือ: 1,
    อีสาน: 4,
    กลาง: 7,
    ตะวันออก: 10,
    ตะวันตก: 10,
    ใต้: 13,
  };
  return mapping[region] || 7;
};

const parseCsv = (text) => {
  const lines = text.trim().split(/\r?\n/);
  const headers = lines.shift().split(",");
  const headerMap = headers.map((header) => normalizeHeader(header));
  return lines
    .map((line) => {
      const cells = line.split(",");
      const record = {};
      headerMap.forEach((key, index) => {
        record[key] = (cells[index] || "").trim();
      });
      return normalizeRow({
        region: record.region,
        province: record.province,
        district: record.district,
        name: record.name,
        party: record.party,
        votes: record.votes,
        counted: record.counted,
      });
    })
    .filter(Boolean);
};

const parsePartyListCsv = (text) => {
  const lines = text.trim().split(/\r?\n/);
  const headers = lines.shift().split(",");
  console.log("Party list headers found:", headers);
  const headerMap = headers.map((header) => normalizePartyListHeader(header));
  console.log("Mapped headers:", headerMap);
  return lines
    .map((line) => {
      const cells = line.split(",");
      const record = {};
      headerMap.forEach((key, index) => {
        record[key] = (cells[index] || "").trim();
      });
      const party = `${record.party ?? ""}`.trim();
      if (!party) {
        return null;
      }
      return {
        party,
        votes: Number(`${record.votes ?? ""}`.trim()) || 0,
        seats: Number(`${record.seats ?? ""}`.trim()) || 0,
      };
    })
    .filter(Boolean);
};

const normalizeHeader = (header) => {
  const clean = header.trim().toLowerCase();
  if (clean === "ภาค") return "region";
  if (clean === "จังหวัด") return "province";
  if (clean === "เขต") return "district";
  if (clean === "ชื่อ") return "name";
  if (clean === "พรรค 2569") return "party";
  if (clean === "คะแนน") return "votes";
  if (clean === "นับคะแนนแล้ว") return "counted";
  if (clean === "region") return "region";
  if (clean === "province") return "province";
  if (clean === "district") return "district";
  if (clean === "candidate") return "name";
  if (clean === "party") return "party";
  if (clean === "votes") return "votes";
  if (clean === "counted") return "counted";
  return clean;
};

const normalizePartyListHeader = (header) => {
  const clean = header.trim().toLowerCase();
  // Party name variations
  if (clean === "พรรค" || clean === "ชื่อพรรค" || clean === "party") return "party";
  // Votes variations
  if (clean === "คะแนน" || clean === "votes") return "votes";
  // Seats variations - many possible names
  if (clean === "ที่นั่ง" || 
      clean === "จำนวนที่นั่ง" || 
      clean === "ส.ส." ||
      clean === "สส" ||
      clean === "ส.ส. บัญชีรายชื่อ" ||
      clean === "สส บัญชีรายชื่อ" ||
      clean === "ส.ส.บัญชีรายชื่อ" ||
      clean === "สสบัญชีรายชื่อ" ||
      clean === "บัญชีรายชื่อ" ||
      clean === "seats" ||
      clean === "จำนวน" ||
      clean === "จำนวน สส บัญชีรายชื่อ" ||
      clean === "จำนวน ส.ส. บัญชีรายชื่อ" ||
      clean === "จำนวน ส.ส.บัญชีรายชื่อ" ||
      clean === "จำนวนสสบัญชีรายชื่อ") return "seats";
  return clean;
};

const handleCsvUpload = (file) => {
  const reader = new FileReader();
  reader.onload = () => {
    const mapped = parseCsv(reader.result);
    state.rawRows = mapped;
    recalculate();
  };
  reader.readAsText(file);
};

const handlePartyListUpload = (file) => {
  const reader = new FileReader();
  reader.onload = () => {
    state.partyList = parsePartyListCsv(reader.result);
    recalculate();
  };
  reader.readAsText(file);
};

const downloadTemplate = async () => {
  const response = await fetch("data/input_template.csv");
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "เทมเพลต_ผลเขต.csv";
  link.click();
  URL.revokeObjectURL(url);
};

const downloadPartyListTemplate = async () => {
  const response = await fetch("data/party_list_template.csv");
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "เทมเพลต_บัญชีรายชื่อ.csv";
  link.click();
  URL.revokeObjectURL(url);
};

const SHEET_ID = "19cLkQfXtcwbnVFR6ilNd7ZeqRTUYU2jrmCCRsyFJ61w";
const DISTRICT_GID = "0";
const PARTY_LIST_GID = "170759107";

// Multiple CORS proxies as fallback
const CORS_PROXIES = [
  "https://corsproxy.io/?",
  "https://api.allorigins.win/raw?url=",
  "https://api.codetabs.com/v1/proxy?quest="
];

const fetchWithProxy = async (targetUrl) => {
  for (const proxy of CORS_PROXIES) {
    try {
      const url = proxy + encodeURIComponent(targetUrl);
      const response = await fetch(url, { cache: "no-store" });
      if (response.ok) {
        return await response.text();
      }
    } catch (e) {
      console.log(`Proxy ${proxy} failed, trying next...`);
    }
  }
  throw new Error("ไม่สามารถดึงข้อมูลจาก Google Sheet ได้");
};

// Progress indicator helpers
const updateProgress = (progressEl, step, label, percent) => {
  if (!progressEl) return;
  progressEl.classList.remove("hidden", "error", "success");
  const stepEl = progressEl.querySelector(".step");
  const labelEl = progressEl.querySelector(".step-label");
  const fillEl = progressEl.querySelector(".import-progress__fill");
  if (stepEl) stepEl.textContent = `ขั้นตอน ${step}/3`;
  if (labelEl) labelEl.textContent = label;
  if (fillEl) fillEl.style.width = `${percent}%`;
};

const showProgressSuccess = (progressEl, message) => {
  if (!progressEl) return;
  progressEl.classList.add("success");
  const labelEl = progressEl.querySelector(".step-label");
  const stepEl = progressEl.querySelector(".step");
  if (stepEl) stepEl.textContent = "สำเร็จ";
  if (labelEl) labelEl.textContent = message;
  setTimeout(() => {
    progressEl.classList.add("hidden");
  }, 3000);
};

const showProgressError = (progressEl, message) => {
  if (!progressEl) return;
  progressEl.classList.add("error");
  const labelEl = progressEl.querySelector(".step-label");
  const stepEl = progressEl.querySelector(".step");
  if (stepEl) stepEl.textContent = "ผิดพลาด";
  if (labelEl) labelEl.textContent = message;
  setTimeout(() => {
    progressEl.classList.add("hidden");
  }, 5000);
};

const importFromGoogleSheet = async () => {
  const progressEl = document.getElementById("districtImportProgress");
  const button = document.getElementById("importFromSheet");
  
  // Disable button during import
  if (button) button.disabled = true;
  
  try {
    // Step 1: Connecting
    updateProgress(progressEl, 1, "กำลังเชื่อมต่อ Google Sheet...", 10);
    
    const cacheBuster = `&_t=${Date.now()}`;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${DISTRICT_GID}${cacheBuster}`;
    
    // Step 2: Downloading
    updateProgress(progressEl, 2, "กำลังดาวน์โหลดข้อมูล...", 40);
    const text = await fetchWithProxy(sheetUrl);
    
    // Step 3: Processing
    updateProgress(progressEl, 3, "กำลังประมวลผลข้อมูล...", 70);
    const mapped = parseCsv(text);
    state.rawRows = mapped;
    
    updateProgress(progressEl, 3, "กำลังอัปเดตหน้าจอ...", 90);
    recalculate();
    
    showProgressSuccess(progressEl, `นำเข้าสำเร็จ! (${mapped.length} แถว)`);
  } catch (error) {
    showProgressError(progressEl, error.message);
  } finally {
    if (button) button.disabled = false;
  }
};

const importPartyListFromGoogleSheet = async () => {
  const progressEl = document.getElementById("partyListImportProgress");
  const button = document.getElementById("importPartyListFromSheet");
  
  // Disable button during import
  if (button) button.disabled = true;
  
  try {
    // Step 1: Connecting
    updateProgress(progressEl, 1, "กำลังเชื่อมต่อ Google Sheet...", 10);
    
    const cacheBuster = `&_t=${Date.now()}`;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${PARTY_LIST_GID}${cacheBuster}`;
    
    // Step 2: Downloading
    updateProgress(progressEl, 2, "กำลังดาวน์โหลดข้อมูล...", 40);
    const text = await fetchWithProxy(sheetUrl);
    
    // Step 3: Processing
    updateProgress(progressEl, 3, "กำลังประมวลผลข้อมูล...", 70);
    console.log("Raw CSV text (first 500 chars):", text.substring(0, 500));
    state.partyList = parsePartyListCsv(text);
    console.log("Parsed party list:", state.partyList);
    
    updateProgress(progressEl, 3, "กำลังอัปเดตหน้าจอ...", 90);
    recalculate();
    
    const totalSeats = state.partyList.reduce((sum, p) => sum + p.seats, 0);
    showProgressSuccess(progressEl, `นำเข้าสำเร็จ! (${state.partyList.length} พรรค, ${totalSeats} ที่นั่ง)`);
  } catch (error) {
    showProgressError(progressEl, error.message);
  } finally {
    if (button) button.disabled = false;
  }
};

const exportProvinceCoordinates = () => {
  const normalizedLayout = normalizeLayout(state.provinceLayout, state.rawRows);
  const provinces = (normalizedLayout.provinces || []).map((provinceDef) => {
    const editedPos = state.editedProvincePositions.get(provinceDef.name);
    return {
      name: provinceDef.name,
      x: editedPos ? editedPos.x : provinceDef.x,
      y: editedPos ? editedPos.y : provinceDef.y,
      columns: editedPos ? editedPos.columns : (provinceDef.columns || 3),
    };
  });

  const json = JSON.stringify({ provinces }, null, 2);
  const blob = new Blob([json], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "province_layout_export.json";
  link.click();
  URL.revokeObjectURL(url);
};

const applyCsv = () => {
  recalculate();
};

document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    document
      .querySelectorAll(".tab")
      .forEach((item) => item.classList.remove("active"));
    tab.classList.add("active");
    document
      .querySelectorAll(".panel")
      .forEach((panel) => panel.classList.remove("active"));
    document.getElementById(tab.dataset.tab).classList.add("active");
  });
});

elements.districtSearch.addEventListener("input", renderPartyVotes);
elements.provinceFilter.addEventListener("change", () => {
  state.activeDistrict = null;
  renderPartyVotes();
  renderDistrictMap();
});
elements.partyFilter.addEventListener("change", renderPartyVotes);
elements.includePartyList.addEventListener("change", (event) => {
  state.includePartyList = event.target.checked;
  renderOverview();
});
elements.downloadTemplate.addEventListener("click", downloadTemplate);
elements.downloadPartyListTemplate.addEventListener(
  "click",
  downloadPartyListTemplate
);
elements.importFromSheet.addEventListener("click", importFromGoogleSheet);
elements.importPartyListFromSheet.addEventListener("click", importPartyListFromGoogleSheet);
elements.csvUpload.addEventListener("change", (event) => {
  if (event.target.files.length) {
    handleCsvUpload(event.target.files[0]);
  }
});
elements.partyListUpload.addEventListener("change", (event) => {
  if (event.target.files.length) {
    handlePartyListUpload(event.target.files[0]);
  }
});
elements.applyCsv.addEventListener("click", applyCsv);
elements.addPartyRow.addEventListener("click", addPartyRow);
if (elements.exportCoordinates) {
  elements.exportCoordinates.addEventListener("click", exportProvinceCoordinates);
}

if (elements.zoomIn && elements.zoomOut && elements.mapViewport) {
  elements.zoomIn.addEventListener("click", () => setMapZoom(state.mapZoom + 0.1));
  elements.zoomOut.addEventListener("click", () => setMapZoom(state.mapZoom - 0.1));
  elements.mapViewport.addEventListener(
    "wheel",
    (event) => {
      event.preventDefault();
      const delta = event.deltaY > 0 ? -0.02 : 0.02;
      const minZoom = isMobile() ? 0.3 : 0.4;
      const newZoom = Math.min(2.5, Math.max(minZoom, state.mapZoom + delta));
      
      // Get mouse position relative to viewport
      const rect = elements.mapViewport.getBoundingClientRect();
      const mouseX = event.clientX - rect.left;
      const mouseY = event.clientY - rect.top;
      
      // Calculate the point in map space under the cursor
      const mapX = (mouseX - state.mapPan.x) / state.mapZoom;
      const mapY = (mouseY - state.mapPan.y) / state.mapZoom;
      
      // Update zoom
      state.mapZoom = newZoom;
      
      // Adjust pan so the same map point stays under the cursor
      state.mapPan.x = mouseX - mapX * newZoom;
      state.mapPan.y = mouseY - mapY * newZoom;
      
      updateMapTransform();
    },
    { passive: false }
  );
}

window.addEventListener("click", (event) => {
  if (!event.target.closest(".party-row")) {
    state.pinnedPartyId = null;
    renderOverview();
    renderPartyList();
  }
});

loadData();

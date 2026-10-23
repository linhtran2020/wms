/**
 * Honeywell WMS - GHN Fulfillment HCM - Q12
 * Application Core Logic: InboundQC, Multi-Location Putaway (PL.Tam.01 - 100), Real-time Inventory
 */

// Global State
let currentTab = 'home';
let pendingParsedOrder = null;
let currentReceiveOrderId = null;
let currentPutawayTask = null;
let putawayAllocations = [];
let allLocations = [];

// Initialize on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
  setupNavigation();
  setupDatalistLocations();
  setupEventListeners();
  loadLocations();
  loadInboundOrders();
  loadPutawayData();
  loadInventoryData();
  loadDashboardStats();
});

// Toast Notification
function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  const icon = type === 'success' ? '✅' : (type === 'error' ? '❌' : 'ℹ️');
  toast.innerHTML = `<span>${icon}</span> <div>${message}</div>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 250);
  }, 3500);
}

// Modal Helpers
function openModal(id) {
  const modal = document.getElementById(id);
  if (modal) modal.classList.add('active');
}

function closeModal(id) {
  const modal = document.getElementById(id);
  if (modal) modal.classList.remove('active');
}

// 1. Navigation & Sidebar Setup
function setupNavigation() {
  // Sidebar items
  document.querySelectorAll('.hw-sidebar .menu-item').forEach(item => {
    item.addEventListener('click', () => {
      const tab = item.dataset.tab;
      if (tab) switchTab(tab);
    });
  });

  // Top workspace tabs
  document.querySelectorAll('.workspace-tabs-bar .wms-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;
      if (tab) switchTab(tab);
    });
  });

  // Toggle Sidebar Button
  const btnToggle = document.getElementById('btn-toggle-sidebar');
  const sidebar = document.getElementById('sidebar');
  if (btnToggle && sidebar) {
    btnToggle.addEventListener('click', () => {
      if (sidebar.style.display === 'none') {
        sidebar.style.display = 'flex';
      } else {
        sidebar.style.display = 'none';
      }
    });
  }
}

function switchTab(tabId) {
  currentTab = tabId;

  // Update Top Workspace Tabs
  document.querySelectorAll('.workspace-tabs-bar .wms-tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tabId);
  });

  // Update Sidebar Active state
  document.querySelectorAll('.hw-sidebar .menu-item').forEach(item => {
    item.classList.toggle('active', item.dataset.tab === tabId);
  });

  // Show Active Tab Page
  document.querySelectorAll('.tab-page').forEach(page => {
    page.classList.toggle('active', page.id === `page-${tabId}`);
  });

  // Refresh tab data
  if (tabId === 'inbound') {
    loadInboundOrders();
  } else if (tabId === 'putaway') {
    loadPutawayData();
  } else if (tabId === 'inventory') {
    loadInventoryData();
  } else if (tabId === 'outbound') {
    loadOutboundData();
  }

  loadDashboardStats();
}

// 2. Setup Valid Locations Datalist (PL.Tam.01 to PL.Tam.100)
function setupDatalistLocations() {
  const datalist = document.getElementById('valid-locations-datalist');
  if (!datalist) return;
  datalist.innerHTML = '';
  for (let i = 1; i <= 100; i++) {
    const num = i < 10 ? `0${i}` : `${i}`;
    const opt = document.createElement('option');
    opt.value = `PL.Tam.${num}`;
    opt.textContent = `Vị trí Pallet Tạm ${num}`;
    datalist.appendChild(opt);
  }
}

// 3. Event Listeners Setup
function setupEventListeners() {
  // Quick Inbound Sample Load
  document.getElementById('btn-quick-inbound')?.addEventListener('click', loadSampleInboundFile);

  // File Input Change
  const fileInput = document.getElementById('inbound-file-input');
  if (fileInput) {
    fileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) handleFileUpload(e.target.files[0]);
    });
  }

  // Cancel Preview
  document.getElementById('btn-cancel-preview')?.addEventListener('click', () => {
    pendingParsedOrder = null;
    document.getElementById('inbound-preview-card').style.display = 'none';
  });

  // Confirm Create Inbound Order
  document.getElementById('btn-confirm-create-order')?.addEventListener('click', confirmCreateInboundOrder);

  // Confirm Goods Receipt into InboundQC
  document.getElementById('btn-confirm-receive')?.addEventListener('click', submitReceiveGoods);

  // Putaway Multi-Location Actions
  document.getElementById('btn-add-location-row')?.addEventListener('click', () => addPutawayAllocationRow());
  document.getElementById('btn-auto-split-locations')?.addEventListener('click', autoSplitPutawayLocations);
  document.getElementById('btn-submit-putaway-sku')?.addEventListener('click', submitPutawaySku);

  // Inventory View Switches
  document.getElementById('btn-inv-summary')?.addEventListener('click', () => switchInventoryView('summary'));
  document.getElementById('btn-inv-details')?.addEventListener('click', () => switchInventoryView('details'));
  document.getElementById('btn-inv-movements')?.addEventListener('click', () => switchInventoryView('movements'));

  // Inventory Search & Filter
  document.getElementById('inv-search-input')?.addEventListener('input', () => loadInventoryData());
  document.getElementById('inv-filter-location')?.addEventListener('change', () => loadInventoryData());

  // Outbound Load Sheet OB
  document.getElementById('btn-load-outbound')?.addEventListener('click', loadSheetOBOutbound);
}

// 4. File Upload & Sheet IB Parsing
function handleFileUpload(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    showToast('Vui lòng chọn file Excel (.xlsx hoặc .xls).', 'error');
    return;
  }

  const formData = new FormData();
  formData.append('excelFile', file);

  showToast(`Đang đọc file ${file.name} (Sheet IB)...`, 'info');

  fetch('/api/inbound/upload', {
    method: 'POST',
    body: formData
  })
    .then(res => res.json())
    .then(result => {
      if (result.success) {
        showToast(result.message, 'success');
        renderParsedInboundPreview(result.data);
      } else {
        showToast(result.message, 'error');
      }
    })
    .catch(err => showToast('Lỗi: ' + err.message, 'error'));
}

function loadSampleInboundFile() {
  showToast('Đang nạp file "Booking GIao hàng Nhanh ngày 16.07.2026.xlsx"...', 'info');

  fetch('/api/inbound/load-sample', { method: 'POST' })
    .then(res => res.json())
    .then(result => {
      if (result.success) {
        showToast(result.message, 'success');
        renderParsedInboundPreview(result.data);
      } else {
        showToast(result.message, 'error');
      }
    })
    .catch(err => showToast('Lỗi: ' + err.message, 'error'));
}

function renderParsedInboundPreview(data) {
  pendingParsedOrder = data;
  const card = document.getElementById('inbound-preview-card');
  const tbody = document.getElementById('inbound-preview-body');
  const metaBadge = document.getElementById('preview-meta-badge');

  metaBadge.textContent = `${data.itemCount} dòng bánh • Tổng ${data.totalExpectedQty} cái`;
  tbody.innerHTML = '';

  data.items.forEach(it => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="color:#64748b;">${it.stt}</td>
      <td><span class="sku-code">${it.sku}</span></td>
      <td><strong>${it.itemName}</strong></td>
      <td><span class="loc-badge" style="background:#f1f5f9; color:#475569; border-color:#cbd5e1;">${it.uom}</span></td>
      <td style="text-align:right; font-weight:700; color:#0369a1;">${it.expectedQty.toLocaleString('vi-VN')}</td>
    `;
    tbody.appendChild(tr);
  });

  card.style.display = 'block';
  card.scrollIntoView({ behavior: 'smooth' });
}

// 5. Confirm Create Inbound Order (Status: "Đã tạo đơn hàng nhập")
function confirmCreateInboundOrder() {
  if (!pendingParsedOrder || !pendingParsedOrder.items) return;

  const payload = {
    supplier: 'Giao Hàng Nhanh (GHN) - Lấy Date Mới trên 70%',
    orderDate: '2026-07-16',
    sourceFileName: pendingParsedOrder.sourceFileName,
    note: 'Đơn nhập hàng tươi theo kế hoạch phân bổ GHN',
    items: pendingParsedOrder.items
  };

  fetch('/api/inbound/create', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  })
    .then(res => res.json())
    .then(result => {
      if (result.success) {
        showToast(result.message, 'success');
        pendingParsedOrder = null;
        document.getElementById('inbound-preview-card').style.display = 'none';
        loadInboundOrders();
        loadDashboardStats();
      } else {
        showToast(result.message, 'error');
      }
    })
    .catch(err => showToast('Lỗi: ' + err.message, 'error'));
}

// 6. Load Inbound Orders
function loadInboundOrders() {
  fetch('/api/inbound')
    .then(res => res.json())
    .then(res => {
      if (!res.success) return;
      const orders = res.data || [];
      const tbody = document.getElementById('inbound-orders-body');
      tbody.innerHTML = '';

      if (orders.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:30px; color:#64748b;">Chưa có đơn nhập nào. Bấm nút <strong>"⚡ Nạp Sheet IB từ file có sẵn"</strong> ở trên để tạo đơn.</td></tr>`;
        return;
      }

      orders.forEach(ord => {
        const tr = document.createElement('tr');

        let statusHtml = '';
        let putawayAllowedHtml = '';
        let actionsHtml = '';

        if (ord.status === 'CREATED') {
          statusHtml = `<span class="status-pill created">Đã tạo đơn hàng nhập</span>`;
          putawayAllowedHtml = `<span style="color:#64748b; font-size:12px;">🔒 Chưa nhận hàng</span>`;
          actionsHtml = `
            <button class="btn btn-red btn-sm" onclick="openReceiveModal('${ord.id}')">
              <span>📥</span> Nhận Hàng (Vào InboundQC)
            </button>
          `;
        } else if (ord.status === 'RECEIVED' || ord.status === 'PUTAWAY_IN_PROGRESS') {
          statusHtml = `<span class="status-pill received">${ord.status === 'RECEIVED' ? 'Đã nhận hàng' : 'Đang cất kho'}</span>`;
          putawayAllowedHtml = `<span style="color:#16a34a; font-weight:700;">✅ ĐƯỢC PHÉP PUTAWAY</span>`;
          actionsHtml = `
            <button class="btn btn-primary btn-sm" onclick="switchTab('putaway')">
              <span>🚀</span> Cất Kho (Putaway PL.Tam)
            </button>
          `;
        } else if (ord.status === 'PUTAWAY_COMPLETED') {
          statusHtml = `<span class="status-pill putaway-complete">Đã cất kho hoàn tất</span>`;
          putawayAllowedHtml = `<span style="color:#16a34a; font-weight:600;">🎉 Hoàn tất</span>`;
          actionsHtml = `<span style="color:#16a34a; font-size:12px;">Đã lưu vào PL.Tam</span>`;
        }

        tr.innerHTML = `
          <td><strong class="sku-code" style="color:#1e293b;">${ord.orderCode}</strong></td>
          <td>${ord.orderDate}</td>
          <td>${ord.supplier}</td>
          <td><span class="loc-badge qc">📍 ${ord.receivingLocation || 'InboundQC'}</span></td>
          <td style="text-align:center;"><strong>${ord.items.length}</strong></td>
          <td style="text-align:right; font-weight:700; color:#0369a1;">${ord.totalExpectedQty.toLocaleString('vi-VN')} cái</td>
          <td style="text-align:center;">${statusHtml}</td>
          <td style="text-align:center;">${putawayAllowedHtml}</td>
          <td style="text-align:center;">${actionsHtml}</td>
        `;

        tbody.appendChild(tr);
      });
    });
}

// 7. Receive Goods Modal (Receives into InboundQC)
function openReceiveModal(orderId) {
  fetch(`/api/inbound/${orderId}`)
    .then(res => res.json())
    .then(result => {
      if (!result.success) return;
      const order = result.data;
      currentReceiveOrderId = orderId;

      const tbody = document.getElementById('receive-modal-tbody');
      tbody.innerHTML = '';

      order.items.forEach(it => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td style="color:#64748b;">${it.stt}</td>
          <td><span class="sku-code">${it.sku}</span></td>
          <td><strong>${it.itemName}</strong></td>
          <td><span class="loc-badge" style="background:#f1f5f9; color:#475569; border-color:#cbd5e1;">${it.uom}</span></td>
          <td style="text-align:right; font-weight:600;">${it.expectedQty}</td>
          <td style="text-align:right;">
            <input type="number" class="receive-qty-input" data-sku="${it.sku}" value="${it.expectedQty}" min="0" style="width:85px; padding:4px 8px; text-align:right; font-weight:700;">
          </td>
          <td>
            <input type="date" class="receive-exp-input" data-sku="${it.sku}" value="${it.expiryDate}" style="padding:4px 8px; font-size:12px; border:1px solid #cbd5e1; border-radius:4px;">
          </td>
        `;
        tbody.appendChild(tr);
      });

      openModal('modal-receive-goods');
    });
}

function submitReceiveGoods() {
  if (!currentReceiveOrderId) return;

  const note = document.getElementById('receive-note-input').value;
  const qtyInputs = document.querySelectorAll('.receive-qty-input');
  const receivedItems = [];

  qtyInputs.forEach(inp => {
    const sku = inp.dataset.sku;
    const qty = parseInt(inp.value, 10) || 0;
    const expInput = document.querySelector(`.receive-exp-input[data-sku="${sku}"]`);
    const expDate = expInput ? expInput.value : '';

    receivedItems.push({
      sku,
      receivedQty: qty,
      expiryDate: expDate
    });
  });

  fetch(`/api/inbound/${currentReceiveOrderId}/receive`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      receiverNote: note,
      receivedItems
    })
  })
    .then(res => res.json())
    .then(result => {
      if (result.success) {
        showToast(result.message, 'success');
        closeModal('modal-receive-goods');
        loadInboundOrders();
        loadPutawayData();
        loadInventoryData();
        loadDashboardStats();

        // Prompt to go to Putaway
        setTimeout(() => {
          if (confirm('Hàng đã được tiếp nhận vào vị trí [InboundQC] và ĐƯỢC PHÉP PUTAWAY. Chuyển sang module Putaway để cất vào các kệ PL.Tam.01 - PL.Tam.100?')) {
            switchTab('putaway');
          }
        }, 500);
      } else {
        showToast(result.message, 'error');
      }
    })
    .catch(err => showToast('Lỗi: ' + err.message, 'error'));
}

// 8. Putaway Module: List items in InboundQC ready for Putaway
function loadPutawayData() {
  fetch('/api/putaway/ready')
    .then(res => res.json())
    .then(result => {
      if (!result.success) return;
      const { tasks } = result.data;
      const tbody = document.getElementById('putaway-table-body');
      const badge = document.getElementById('putaway-tasks-badge');

      badge.textContent = `${tasks.length} mặt hàng chờ cất`;
      tbody.innerHTML = '';

      if (tasks.length === 0) {
        tbody.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:30px; color:#64748b;">Không có mặt hàng nào chờ cất kho. Các đơn hàng ở trạng thái <strong>"Đã nhận hàng"</strong> tại InboundQC sẽ hiển thị ở đây.</td></tr>`;
        return;
      }

      tasks.forEach(t => {
        const tr = document.createElement('tr');

        // Formatted putaway allocations summary
        let allocSummaryHtml = '<span style="color:#94a3b8; font-size:12px;">Chưa cất</span>';
        if (t.putawayAllocations && t.putawayAllocations.length > 0) {
          allocSummaryHtml = t.putawayAllocations.map(a => `<span class="loc-badge" style="margin:2px;">${a.locationCode}: ${a.qty}</span>`).join(' ');
        }

        tr.innerHTML = `
          <td><strong class="sku-code" style="color:#1e293b;">${t.orderCode}</strong></td>
          <td><span class="sku-code">${t.sku}</span></td>
          <td><strong>${t.itemName}</strong></td>
          <td><span class="loc-badge" style="background:#f1f5f9; color:#475569; border-color:#cbd5e1;">${t.uom}</span></td>
          <td><span class="loc-badge qc">📍 InboundQC</span></td>
          <td style="text-align:right; font-weight:600;">${t.receivedQty}</td>
          <td style="text-align:right; font-weight:600; color:#16a34a;">${t.putawayQty}</td>
          <td style="text-align:right; font-weight:700; color:var(--hw-red); font-size:14px;">${t.remainingQty}</td>
          <td>${allocSummaryHtml}</td>
          <td style="text-align:center;">
            <button class="btn btn-red btn-sm" onclick='openPutawaySkuModal(${JSON.stringify(t)})'>
              <span>🚀</span> Cất Kho (Đa Vị Trí)
            </button>
          </td>
        `;

        tbody.appendChild(tr);
      });
    });
}

// 9. Open Multi-Location Putaway Modal for a SKU
function openPutawaySkuModal(task) {
  currentPutawayTask = task;
  putawayAllocations = [];

  document.getElementById('modal-putaway-sku-code').textContent = task.sku;
  document.getElementById('modal-putaway-sku-name').textContent = task.itemName;
  document.getElementById('modal-putaway-uom').textContent = task.uom;
  document.getElementById('modal-putaway-total-qty').textContent = task.remainingQty;

  // Initial allocation: if suggested locations available, add first one
  const firstLoc = (task.suggestedLocations && task.suggestedLocations[0]) || 'PL.Tam.01';
  putawayAllocations.push({
    locationCode: firstLoc,
    qty: task.remainingQty
  });

  renderPutawayAllocations();
  openModal('modal-putaway-sku');
}

function renderPutawayAllocations() {
  const container = document.getElementById('allocations-container');
  container.innerHTML = '';

  const totalRequired = currentPutawayTask.remainingQty;
  let totalAllocated = 0;

  putawayAllocations.forEach((alloc, idx) => {
    totalAllocated += alloc.qty || 0;

    const row = document.createElement('div');
    row.className = 'alloc-row';
    row.innerHTML = `
      <span style="font-size:12px; font-weight:700; color:#64748b; width:65px;">Vị trí ${idx + 1}:</span>
      <input type="text" class="alloc-loc-input" list="valid-locations-datalist" value="${alloc.locationCode}" placeholder="PL.Tam.01 - 100" data-idx="${idx}">
      <span style="font-size:12px; color:#64748b;">Số lượng:</span>
      <input type="number" class="alloc-qty-input" value="${alloc.qty}" min="1" max="${totalRequired}" data-idx="${idx}">
      <span style="font-size:12px; color:#475569;">${currentPutawayTask.uom}</span>
      <button class="btn btn-secondary btn-sm" style="color:#dc2626; margin-left:auto;" onclick="removePutawayAllocationRow(${idx})" title="Xóa dòng này">
        🗑️
      </button>
    `;
    container.appendChild(row);
  });

  // Attach change events
  container.querySelectorAll('.alloc-loc-input').forEach(inp => {
    inp.addEventListener('input', (e) => {
      const idx = parseInt(e.target.dataset.idx, 10);
      putawayAllocations[idx].locationCode = e.target.value.trim();
      updatePutawayAllocCounters();
    });
  });

  container.querySelectorAll('.alloc-qty-input').forEach(inp => {
    inp.addEventListener('input', (e) => {
      const idx = parseInt(e.target.dataset.idx, 10);
      putawayAllocations[idx].qty = parseInt(e.target.value, 10) || 0;
      updatePutawayAllocCounters();
    });
  });

  updatePutawayAllocCounters();
}

function updatePutawayAllocCounters() {
  const totalRequired = currentPutawayTask ? currentPutawayTask.remainingQty : 0;
  const totalAllocated = putawayAllocations.reduce((sum, a) => sum + (parseInt(a.qty, 10) || 0), 0);
  const remaining = totalRequired - totalAllocated;

  const allocatedEl = document.getElementById('modal-putaway-allocated-qty');
  const remainingEl = document.getElementById('modal-putaway-remaining-qty');

  allocatedEl.textContent = totalAllocated;
  remainingEl.textContent = remaining;

  if (remaining === 0) {
    allocatedEl.style.color = '#16a34a';
    remainingEl.style.color = '#16a34a';
  } else if (remaining < 0) {
    allocatedEl.style.color = '#dc2626';
    remainingEl.style.color = '#dc2626';
  } else {
    allocatedEl.style.color = '#b45309';
    remainingEl.style.color = '#dc2626';
  }
}

function addPutawayAllocationRow() {
  const nextIdx = putawayAllocations.length + 1;
  const nextNum = nextIdx < 10 ? `0${nextIdx}` : `${nextIdx}`;
  const nextLoc = `PL.Tam.${nextNum}`;

  putawayAllocations.push({
    locationCode: nextLoc,
    qty: 0
  });
  renderPutawayAllocations();
}

function removePutawayAllocationRow(idx) {
  if (putawayAllocations.length <= 1) {
    showToast('Phải có ít nhất 1 vị trí tiếp nhận hàng!', 'error');
    return;
  }
  putawayAllocations.splice(idx, 1);
  renderPutawayAllocations();
}

// Auto split high-quantity SKU across multiple empty PL.Tam locations (e.g. 85 -> 45 in PL.Tam.01, 40 in PL.Tam.02)
function autoSplitPutawayLocations() {
  if (!currentPutawayTask) return;
  const total = currentPutawayTask.remainingQty;
  const capacityPerPallet = 45; // Reasonable pallet size
  const neededLocations = Math.max(1, Math.ceil(total / capacityPerPallet));

  putawayAllocations = [];
  let remainingToDistribute = total;

  for (let i = 1; i <= neededLocations; i++) {
    const locNum = i < 10 ? `0${i}` : `${i}`;
    const allocQty = Math.min(capacityPerPallet, remainingToDistribute);
    putawayAllocations.push({
      locationCode: `PL.Tam.${locNum}`,
      qty: allocQty
    });
    remainingToDistribute -= allocQty;
  }

  renderPutawayAllocations();
  showToast(`Đã tự động gợi ý phân bổ ${total} cái vào ${neededLocations} vị trí kệ PL.Tam.`, 'info');
}

// 10. Submit Multi-location Putaway for the SKU
function submitPutawaySku() {
  if (!currentPutawayTask || putawayAllocations.length === 0) return;

  const totalRequired = currentPutawayTask.remainingQty;
  const totalAllocated = putawayAllocations.reduce((sum, a) => sum + (parseInt(a.qty, 10) || 0), 0);

  if (totalAllocated === 0) {
    showToast('Vui lòng nhập số lượng cất kho lớn hơn 0!', 'error');
    return;
  }

  if (totalAllocated > totalRequired) {
    showToast(`Tổng số lượng phân bổ (${totalAllocated}) vượt quá số lượng cần cất (${totalRequired})!`, 'error');
    return;
  }

  // Validate all location codes: must be PL.Tam.01 to PL.Tam.100
  for (const alloc of putawayAllocations) {
    const code = (alloc.locationCode || '').trim();
    const match = code.match(/^PL\.Tam\.(\d{1,3})$/i);
    if (!match) {
      showToast(`Vị trí "${code}" không hợp lệ! Vị trí được phép chỉ từ PL.Tam.01 đến PL.Tam.100.`, 'error');
      return;
    }
    const num = parseInt(match[1], 10);
    if (num < 1 || num > 100) {
      showToast(`Vị trí "${code}" nằm ngoài phạm vi được phép (PL.Tam.01 đến PL.Tam.100)!`, 'error');
      return;
    }
    // Normalize format
    const formatted = num < 10 ? `0${num}` : `${num}`;
    alloc.locationCode = `PL.Tam.${formatted}`;
  }

  const payload = {
    orderId: currentPutawayTask.orderId,
    items: [
      {
        sku: currentPutawayTask.sku,
        allocations: putawayAllocations,
        batchNo: currentPutawayTask.batchNo,
        expiryDate: currentPutawayTask.expiryDate
      }
    ]
  };

  fetch('/api/putaway/execute', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  })
    .then(res => res.json())
    .then(result => {
      if (result.success) {
        showToast(result.message, 'success');
        closeModal('modal-putaway-sku');
        loadPutawayData();
        loadInboundOrders();
        loadInventoryData();
        loadDashboardStats();
      } else {
        showToast(result.message, 'error');
      }
    })
    .catch(err => showToast('Lỗi: ' + err.message, 'error'));
}

// 11. Inventory Management Views
function switchInventoryView(mode) {
  const sumView = document.getElementById('inv-view-summary');
  const detView = document.getElementById('inv-view-details');
  const movView = document.getElementById('inv-view-movements');

  const btnSum = document.getElementById('btn-inv-summary');
  const btnDet = document.getElementById('btn-inv-details');
  const btnMov = document.getElementById('btn-inv-movements');

  [sumView, detView, movView].forEach(v => v.style.display = 'none');
  [btnSum, btnDet, btnMov].forEach(b => {
    b.classList.remove('btn-primary');
    b.classList.add('btn-secondary');
  });

  if (mode === 'summary') {
    sumView.style.display = 'block';
    btnSum.classList.remove('btn-secondary');
    btnSum.classList.add('btn-primary');
  } else if (mode === 'details') {
    detView.style.display = 'block';
    btnDet.classList.remove('btn-secondary');
    btnDet.classList.add('btn-primary');
  } else if (mode === 'movements') {
    movView.style.display = 'block';
    btnMov.classList.remove('btn-secondary');
    btnMov.classList.add('btn-primary');
    loadStockMovements();
  }
}

function loadInventoryData() {
  const search = document.getElementById('inv-search-input')?.value || '';
  const locFilter = document.getElementById('inv-filter-location')?.value || '';

  const url = `/api/inventory?search=${encodeURIComponent(search)}&location=${encodeURIComponent(locFilter)}`;

  fetch(url)
    .then(res => res.json())
    .then(result => {
      if (!result.success) return;
      const { skuSummary, details } = result.data;

      // 1. Render SKU Summary
      const sumTbody = document.getElementById('inv-summary-tbody');
      sumTbody.innerHTML = '';

      if (skuSummary.length === 0) {
        sumTbody.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:30px; color:#64748b;">Kho hiện tại đang trống. Hàng nhận vào InboundQC hoặc Putaway vào PL.Tam sẽ hiển thị tại đây.</td></tr>`;
      } else {
        skuSummary.forEach(item => {
          const tr = document.createElement('tr');

          // Build multi-location breakdown pills
          const locPills = item.locationBreakdown.map(lb => {
            const isQc = lb.location === 'InboundQC';
            const badgeClass = isQc ? 'loc-badge qc' : 'loc-badge';
            return `<span class="${badgeClass}" style="margin:2px 4px 2px 0;">${lb.location}: <strong>${lb.qty}</strong></span>`;
          }).join(' ');

          tr.innerHTML = `
            <td><span class="sku-code">${item.sku}</span></td>
            <td><strong>${item.itemName}</strong></td>
            <td><span class="loc-badge" style="background:#f1f5f9; color:#475569; border-color:#cbd5e1;">${item.uom}</span></td>
            <td style="text-align:right; font-weight:800; color:#0369a1; font-size:15px;">${item.totalQuantity.toLocaleString('vi-VN')}</td>
            <td style="text-align:right; font-weight:700; color:#b45309;">${item.inboundQcQty > 0 ? item.inboundQcQty : '-'}</td>
            <td>${locPills}</td>
          `;
          sumTbody.appendChild(tr);
        });
      }

      // 2. Render Details View (Per Location and Batch)
      const detTbody = document.getElementById('inv-details-tbody');
      detTbody.innerHTML = '';
      details.forEach(item => {
        const tr = document.createElement('tr');
        const isQc = item.location === 'InboundQC';
        const locBadge = isQc ? '<span class="loc-badge qc">📍 InboundQC</span>' : `<span class="loc-badge">${item.location}</span>`;
        const updatedDate = item.lastUpdated ? new Date(item.lastUpdated).toLocaleDateString('vi-VN') : '-';

        tr.innerHTML = `
          <td>${locBadge}</td>
          <td><span class="sku-code">${item.sku}</span></td>
          <td><strong>${item.itemName}</strong></td>
          <td><span class="loc-badge" style="background:#f1f5f9; color:#475569; border-color:#cbd5e1;">${item.uom}</span></td>
          <td><span class="sku-code" style="color:#64748b; background:#f8fafc; border-color:transparent;">${item.batchNo}</span></td>
          <td style="font-size:12px; color:#b45309;">${item.expiryDate || '-'}</td>
          <td style="text-align:right; font-weight:700; color:#0369a1;">${item.quantity.toLocaleString('vi-VN')}</td>
          <td style="font-size:12px; color:#64748b;">${updatedDate}</td>
        `;
        detTbody.appendChild(tr);
      });
    });
}

function loadStockMovements() {
  fetch('/api/inventory/movements')
    .then(res => res.json())
    .then(result => {
      if (!result.success) return;
      const movements = result.data || [];
      const tbody = document.getElementById('inv-movements-tbody');
      tbody.innerHTML = '';

      movements.forEach(m => {
        const tr = document.createElement('tr');
        const timeStr = new Date(m.timestamp).toLocaleString('vi-VN');
        const isReceive = m.type === 'INBOUND_RECEIVE';
        const typePill = isReceive
          ? `<span class="status-pill created">Nhận Hàng (InboundQC)</span>`
          : `<span class="status-pill putaway-complete">Putaway (PL.Tam)</span>`;

        tr.innerHTML = `
          <td style="font-size:12px; color:#64748b;">${timeStr}</td>
          <td>${typePill}</td>
          <td><strong class="sku-code" style="color:#1e293b;">${m.orderId}</strong></td>
          <td><span class="sku-code">${m.sku}</span></td>
          <td><strong>${m.itemName}</strong></td>
          <td><span class="loc-badge ${m.fromLocation === 'InboundQC' ? 'qc' : ''}">${m.fromLocation}</span></td>
          <td><span class="loc-badge ${m.toLocation === 'InboundQC' ? 'qc' : ''}">${m.toLocation}</span></td>
          <td style="text-align:right; font-weight:700; color:#0369a1;">+${m.quantity}</td>
          <td style="font-size:12px; color:#64748b;">${m.note || '-'}</td>
        `;
        tbody.appendChild(tr);
      });
    });
}

// 12. Load Locations for Filter Dropdown
function loadLocations() {
  fetch('/api/locations')
    .then(res => res.json())
    .then(result => {
      if (!result.success) return;
      allLocations = result.data || [];

      const select = document.getElementById('inv-filter-location');
      if (select) {
        select.innerHTML = '<option value="">Tất cả vị trí (InboundQC & PL.Tam)</option>';
        allLocations.forEach(loc => {
          select.innerHTML += `<option value="${loc.code}">${loc.code} (${loc.name})</option>`;
        });
      }
    });
}

// 13. Outbound Module (Sheet OB)
function loadSheetOBOutbound() {
  showToast('Đang trích xuất 183 dòng đơn hàng Vinmart+ từ sheet "OB"...', 'info');
  fetch('/api/outbound/load-default', { method: 'POST' })
    .then(res => res.json())
    .then(result => {
      if (result.success) {
        showToast(result.message, 'success');
        loadOutboundData();
      } else {
        showToast(result.message, 'error');
      }
    });
}

function loadOutboundData() {
  fetch('/api/outbound')
    .then(res => res.json())
    .then(result => {
      if (!result.success) return;
      const orders = result.data || [];
      const tbody = document.getElementById('outbound-tbody');
      const badge = document.getElementById('outbound-badge');

      badge.textContent = `${orders.length} đơn hàng`;
      tbody.innerHTML = '';

      if (orders.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:30px; color:#64748b;">Chưa có đơn xuất nào. Bấm nút <strong>"⚡ Nạp Đơn Xuất Từ Sheet OB"</strong> ở trên để tải dữ liệu.</td></tr>`;
        return;
      }

      orders.forEach(o => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><strong class="sku-code" style="color:#1e293b;">${o.id}</strong></td>
          <td><strong>${o.customerName}</strong></td>
          <td style="text-align:center;">${o.items.length} SKU</td>
          <td style="text-align:right; font-weight:700; color:#0369a1;">${o.totalQty} cái</td>
          <td style="text-align:center;"><span class="status-pill received">Chờ xuất kho</span></td>
        `;
        tbody.appendChild(tr);
      });
    });
}

// 14. Dashboard Stats
function loadDashboardStats() {
  fetch('/api/dashboard/stats')
    .then(res => res.json())
    .then(result => {
      // stats can be used in topbar or tooltips
    });
}

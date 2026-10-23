const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Configure Multer for file uploads (memory storage: works on Vercel's read-only filesystem)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 4 * 1024 * 1024 }
});

// ===== Storage backend =====
// On Vercel: Upstash Redis via REST (env vars from Vercel Marketplace Upstash integration or Upstash console)
const REDIS_URL = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
const REDIS_TOKEN = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
const USE_REDIS = Boolean(REDIS_URL && REDIS_TOKEN);
const REDIS_KEY = process.env.WMS_DB_KEY || 'wms:db';

// Local mode only: ensure data directory exists
if (!USE_REDIS) {
  try {
    const p = path.join(__dirname, 'data');
    if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true });
  } catch (e) {}
}

const DB_PATH = path.join(__dirname, 'data', 'db.json');
let dbCache = null; // In-memory copy, hydrated from Redis at the start of each API request
let pendingSave = null; // Promise of the latest Redis write

async function redisCommand(args) {
  const resp = await fetch(REDIS_URL, {
    method: 'POST',
    headers: { Authorization: `Bearer ${REDIS_TOKEN}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(args)
  });
  const json = await resp.json();
  if (!resp.ok || json.error) throw new Error(json.error || `Redis HTTP ${resp.status}`);
  return json.result;
}

function emptyDB() {
  return {
    inboundOrders: [],
    putawayTasks: [],
    inventory: [],
    stockMovements: [],
    outboundOrders: [],
    locations: generateDefaultLocations()
  };
}

// Middleware: hydrate DB from Redis before API handlers, and wait for Redis writes before responding
app.use('/api', async (req, res, next) => {
  if (!USE_REDIS) return next();
  try {
    const raw = await redisCommand(['GET', REDIS_KEY]);
    dbCache = raw ? JSON.parse(raw) : null;
  } catch (err) {
    console.error('Redis read error:', err);
    return res.status(500).json({ success: false, message: 'Không đọc được database (Redis): ' + err.message });
  }
  pendingSave = null;
  const originalJson = res.json.bind(res);
  res.json = (body) => {
    Promise.resolve(pendingSave)
      .then(() => originalJson(body))
      .catch(err => {
        res.status(500);
        originalJson({ success: false, message: 'Không lưu được database (Redis): ' + err.message });
      });
    return res;
  };
  next();
});
const DEFAULT_EXCEL_FILE = path.join(__dirname, 'Booking GIao hàng Nhanh ngày 16.07.2026.xlsx');

// Generate 100 PL.Tam locations: PL.Tam.01 to PL.Tam.100 + InboundQC
function generateDefaultLocations() {
  const locs = [
    {
      code: 'InboundQC',
      zone: 'Khu Nhận Hàng & Kiểm Soát Chất Lượng (Inbound QC Area)',
      name: 'Vị trí Nhận hàng Mặc định InboundQC',
      capacity: 999999,
      current: 0,
      isDefaultInbound: true
    }
  ];

  for (let i = 1; i <= 100; i++) {
    const numStr = i < 10 ? `0${i}` : `${i}`;
    const code = `PL.Tam.${numStr}`;
    locs.push({
      code,
      zone: 'Khu Vực Lưu Trữ Tạm (PL.Tam Zone)',
      name: `Vị trí Pallet Tạm ${numStr}`,
      capacity: 100,
      current: 0,
      isPutawayAllowed: true
    });
  }

  return locs;
}

// Validate Putaway Location: must be PL.Tam.01 to PL.Tam.100
function isValidPutawayLocation(code) {
  if (!code || typeof code !== 'string') return false;
  const match = code.trim().match(/^PL\.Tam\.(\d{2,3})$/i);
  if (!match) return false;
  const num = parseInt(match[1], 10);
  return num >= 1 && num <= 100;
}

// Helper to normalize location format: e.g. PL.Tam.1 -> PL.Tam.01
function normalizePutawayLocation(code) {
  if (!code) return '';
  const match = code.trim().match(/^PL\.Tam\.(\d{1,3})$/i);
  if (!match) return code.trim();
  const num = parseInt(match[1], 10);
  const formatted = num < 10 ? `0${num}` : `${num}`;
  return `PL.Tam.${formatted}`;
}

// Database Read/Write Helper
function loadDB() {
  if (USE_REDIS) {
    if (!dbCache) {
      dbCache = emptyDB();
      saveDB(dbCache);
    } else if (!dbCache.locations || dbCache.locations.length < 50 || !dbCache.locations.some(l => l.code === 'InboundQC')) {
      dbCache.locations = generateDefaultLocations();
      saveDB(dbCache);
    }
    return dbCache;
  }
  try {
    if (!fs.existsSync(DB_PATH)) {
      const initData = {
        inboundOrders: [],
        putawayTasks: [],
        inventory: [],
        stockMovements: [],
        outboundOrders: [],
        locations: generateDefaultLocations()
      };
      saveDB(initData);
      return initData;
    }
    const raw = fs.readFileSync(DB_PATH, 'utf-8');
    const parsed = JSON.parse(raw);

    // Ensure InboundQC and PL.Tam.01-100 exist
    if (!parsed.locations || parsed.locations.length < 50 || !parsed.locations.some(l => l.code === 'InboundQC')) {
      parsed.locations = generateDefaultLocations();
      saveDB(parsed);
    }
    return parsed;
  } catch (err) {
    console.error('Error reading db.json:', err);
    return {
      inboundOrders: [],
      putawayTasks: [],
      inventory: [],
      stockMovements: [],
      outboundOrders: [],
      locations: generateDefaultLocations()
    };
  }
}

function saveDB(data) {
  if (USE_REDIS) {
    dbCache = data;
    pendingSave = redisCommand(['SET', REDIS_KEY, JSON.stringify(data)]);
    pendingSave.catch(err => console.error('Redis write error:', err));
    return;
  }
  try {
    fs.writeFileSync(DB_PATH, JSON.stringify(data, null, 2), 'utf-8');
  } catch (err) {
    console.error('Error saving db.json:', err);
  }
}

// Smart Location Suggestion among PL.Tam.01 to PL.Tam.100
function suggestPutawayLocations(neededLocationsCount, locations, inventory) {
  const suggestions = [];
  for (let i = 1; i <= 100; i++) {
    const numStr = i < 10 ? `0${i}` : `${i}`;
    const code = `PL.Tam.${numStr}`;
    const loc = locations.find(l => l.code === code);
    const storedQty = (loc && loc.current) || 0;
    if (storedQty < (loc?.capacity || 100)) {
      suggestions.push(code);
      if (suggestions.length >= neededLocationsCount) break;
    }
  }
  return suggestions.length > 0 ? suggestions : ['PL.Tam.01'];
}

// Excel Parsing Logic for Inbound (Sheet 'IB')
function parseInboundWorkbook(workbook, fileName = 'Uploaded.xlsx') {
  const sheetNames = workbook.SheetNames;
  let targetSheetName = sheetNames.find(n => n.toUpperCase().trim() === 'IB');
  if (!targetSheetName) {
    targetSheetName = sheetNames.find(n => n.toLowerCase().includes('inbound') || n.toLowerCase().includes('nhập')) || sheetNames[0];
  }

  const sheet = workbook.Sheets[targetSheetName];
  if (!sheet) {
    throw new Error(`Không tìm thấy sheet IB trong file Excel.`);
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  if (rows.length < 2) {
    throw new Error('Sheet IB không có dữ liệu hàng hóa hợp lệ.');
  }

  let headerIndex = -1;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const r = rows[i] || [];
    const str = r.map(c => String(c || '').toLowerCase()).join(' ');
    if (str.includes('sku') || str.includes('tên bánh') || str.includes('hóa đơn')) {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) headerIndex = 0;

  const header = (rows[headerIndex] || []).map(c => String(c || '').trim());
  let skuCol = header.findIndex(h => /sku/i.test(h));
  let nameCol = header.findIndex(h => /tên bánh|tên hàng|item name|description/i.test(h));
  let uomCol = header.findIndex(h => /uom|đơn vị|dvt/i.test(h));
  let qtyCol = header.findIndex(h => /hóa đơn|sl|số lượng|quantity|qty/i.test(h));

  if (skuCol === -1) skuCol = 1;
  if (nameCol === -1) nameCol = 2;
  if (uomCol === -1) uomCol = 3;
  if (qtyCol === -1) qtyCol = 4;

  const items = [];
  let sttCounter = 1;

  for (let i = headerIndex + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length === 0) continue;

    const rawSku = r[skuCol];
    const rawName = r[nameCol];
    const rawUom = r[uomCol];
    const rawQty = r[qtyCol];

    if (!rawSku && !rawName) continue;

    const skuStr = String(rawSku || '').trim();
    const nameStr = String(rawName || '').trim();

    if (/grand\s*total|tổng\s*cộng|tổng/i.test(skuStr) || /grand\s*total|tổng\s*cộng|tổng/i.test(nameStr)) {
      continue;
    }

    const qty = parseInt(rawQty, 10);
    if (isNaN(qty) || qty <= 0) continue;

    items.push({
      stt: sttCounter++,
      sku: skuStr,
      itemName: nameStr || `Mặt hàng SKU-${skuStr}`,
      uom: String(rawUom || 'CAI').trim().toUpperCase(),
      expectedQty: qty,
      receivedQty: qty,
      putawayQty: 0,
      batchNo: `LOT-${new Date().toISOString().slice(2, 10).replace(/-/g, '')}`,
      expiryDate: new Date(Date.now() + 15 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10)
    });
  }

  const totalExpectedQty = items.reduce((sum, item) => sum + item.expectedQty, 0);

  return {
    sourceFileName: fileName,
    sheetName: targetSheetName,
    items,
    itemCount: items.length,
    totalExpectedQty
  };
}

// Excel Parsing Logic for Outbound (Sheet 'OB')
function parseOutboundWorkbook(workbook, fileName = 'Uploaded.xlsx') {
  const sheetNames = workbook.SheetNames;
  let targetSheetName = sheetNames.find(n => n.toUpperCase().trim() === 'OB') ||
                        sheetNames.find(n => n.toLowerCase().includes('outbound') || n.toLowerCase().includes('xuất'));

  if (!targetSheetName) return { orders: [], totalOrders: 0 };

  const sheet = workbook.Sheets[targetSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  if (rows.length < 2) return { orders: [], totalOrders: 0 };

  const ordersMap = new Map();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length < 10) continue;

    const orderNum = String(r[1] || '').trim();
    const customer = String(r[4] || 'Vinmart+ HCM').trim();
    const sku = String(r[8] || '').trim();
    const itemName = String(r[9] || '').trim();
    const uom = String(r[10] || 'CAI').trim();
    const qty = parseInt(r[11], 10) || 0;

    if (!orderNum || !sku || qty <= 0) continue;

    if (!ordersMap.has(orderNum)) {
      ordersMap.set(orderNum, {
        id: `OB-${orderNum}`,
        orderNumber: orderNum,
        customerName: customer,
        status: 'PENDING',
        items: [],
        totalQty: 0,
        createdAt: new Date().toISOString()
      });
    }

    const order = ordersMap.get(orderNum);
    order.items.push({
      sku,
      itemName,
      uom,
      orderQty: qty,
      pickedQty: 0
    });
    order.totalQty += qty;
  }

  return {
    orders: Array.from(ordersMap.values()),
    totalOrders: ordersMap.size
  };
}

// ======================== API ROUTES ======================== //

// 1. Dashboard Overview Stats
app.get('/api/dashboard/stats', (req, res) => {
  const db = loadDB();
  const totalInbound = db.inboundOrders.length;
  const createdInbound = db.inboundOrders.filter(o => o.status === 'CREATED').length;
  const receivedInbound = db.inboundOrders.filter(o => o.status === 'RECEIVED' || o.status === 'PUTAWAY_IN_PROGRESS').length;
  const completedInbound = db.inboundOrders.filter(o => o.status === 'PUTAWAY_COMPLETED').length;

  const totalInventoryUnits = db.inventory.reduce((sum, i) => sum + (i.quantity || 0), 0);
  const totalSKUs = new Set(db.inventory.filter(i => i.quantity > 0).map(i => i.sku)).size;

  // Quantity currently waiting in InboundQC
  const inboundQCUnits = db.inventory
    .filter(i => i.location === 'InboundQC')
    .reduce((sum, i) => sum + (i.quantity || 0), 0);

  // Capacity in PL.Tam.01 to PL.Tam.100
  const plTamLocations = db.locations.filter(l => l.code.startsWith('PL.Tam.'));
  const totalCapacity = plTamLocations.reduce((sum, l) => sum + (l.capacity || 100), 0);
  const usedCapacity = plTamLocations.reduce((sum, l) => sum + (l.current || 0), 0);
  const occupancyRate = totalCapacity > 0 ? Math.round((usedCapacity / totalCapacity) * 100) : 0;

  res.json({
    success: true,
    data: {
      inbound: {
        total: totalInbound,
        created: createdInbound,
        received: receivedInbound,
        completed: completedInbound,
        inboundQCUnits
      },
      inventory: {
        totalUnits: totalInventoryUnits,
        inboundQCUnits,
        skuCount: totalSKUs,
        totalLocations: db.locations.length,
        occupancyRate,
        usedCapacity,
        totalCapacity
      },
      recentMovements: db.stockMovements.slice(-10).reverse()
    }
  });
});

// 2. Parse uploaded Excel file (Sheet 'IB')
app.post('/api/inbound/upload', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: 'Vui lòng đính kèm file Excel (.xlsx, .xls).' });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const parsed = parseInboundWorkbook(workbook, req.file.originalname);

    res.json({
      success: true,
      message: `Đã đọc thành công sheet "${parsed.sheetName}" với ${parsed.itemCount} mặt hàng (Tổng: ${parsed.totalExpectedQty} cái).`,
      data: parsed
    });
  } catch (err) {
    console.error('Error parsing uploaded file:', err);
    res.status(500).json({ success: false, message: err.message || 'Lỗi khi đọc file Excel.' });
  }
});

// 3. Load default Excel file directly from workspace
app.post('/api/inbound/load-sample', (req, res) => {
  try {
    if (!fs.existsSync(DEFAULT_EXCEL_FILE)) {
      return res.status(404).json({ success: false, message: 'Không tìm thấy file mẫu Booking Giao hàng Nhanh trong thư mục.' });
    }

    const workbook = XLSX.readFile(DEFAULT_EXCEL_FILE);
    const parsed = parseInboundWorkbook(workbook, path.basename(DEFAULT_EXCEL_FILE));

    res.json({
      success: true,
      message: `Đã nạp file mẫu thành công: Sheet "${parsed.sheetName}" có ${parsed.itemCount} dòng hàng, tổng ${parsed.totalExpectedQty} cái.`,
      data: parsed
    });
  } catch (err) {
    console.error('Error loading default excel:', err);
    res.status(500).json({ success: false, message: err.message || 'Lỗi khi mở file mẫu.' });
  }
});

// 4. Create Inbound Order (Trạng thái: "Đã tạo đơn hàng nhập")
app.post('/api/inbound/create', (req, res) => {
  try {
    const { orderCode, supplier, orderDate, note, items, sourceFileName } = req.body;

    if (!items || !Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ success: false, message: 'Danh sách mặt hàng nhập không được rỗng.' });
    }

    const db = loadDB();

    const dateStr = (orderDate || new Date().toISOString().slice(0, 10)).replace(/-/g, '');
    const countToday = db.inboundOrders.filter(o => o.orderCode && o.orderCode.includes(dateStr)).length + 1;
    const finalOrderCode = (orderCode && orderCode.trim()) || `IB-${dateStr}-${String(countToday).padStart(4, '0')}`;

    const totalExpectedQty = items.reduce((sum, it) => sum + (parseInt(it.expectedQty, 10) || 0), 0);

    const newOrder = {
      id: finalOrderCode,
      orderCode: finalOrderCode,
      supplier: supplier || 'Giao Hàng Nhanh (GHN) - Lấy Date Mới trên 70%',
      orderDate: orderDate || new Date().toISOString().slice(0, 10),
      sourceFileName: sourceFileName || 'Booking GIao hàng Nhanh ngày 16.07.2026.xlsx',
      note: note || 'Đơn nhập hàng tươi theo kế hoạch phân bổ GHN',
      status: 'CREATED', // YÊU CẦU: Trạng thái ban đầu là "Đã tạo đơn hàng nhập"
      statusLabel: 'Đã tạo đơn hàng nhập',
      putawayAllowed: false, // Chưa được phép putaway cho đến khi nhận hàng
      receivingLocation: 'InboundQC', // YÊU CẦU: Vị trí nhận hàng mặc định là InboundQC
      items: items.map((it, idx) => ({
        stt: idx + 1,
        sku: String(it.sku),
        itemName: it.itemName,
        uom: it.uom || 'CAI',
        expectedQty: parseInt(it.expectedQty, 10) || 0,
        receivedQty: parseInt(it.receivedQty || it.expectedQty, 10) || 0,
        putawayQty: 0,
        putawayAllocations: [], // Sẽ lưu chi tiết nhiều vị trí PL.Tam và số lượng tương ứng
        batchNo: it.batchNo || `LOT-${dateStr.slice(2)}`,
        expiryDate: it.expiryDate || new Date(Date.now() + 15 * 86400000).toISOString().slice(0, 10)
      })),
      totalExpectedQty,
      totalReceivedQty: 0,
      createdAt: new Date().toISOString(),
      receivedAt: null,
      putawayAt: null
    };

    db.inboundOrders.unshift(newOrder);
    saveDB(db);

    res.json({
      success: true,
      message: `Tạo thành công đơn nhập ${newOrder.orderCode}. Trạng thái: "Đã tạo đơn hàng nhập".`,
      data: newOrder
    });
  } catch (err) {
    console.error('Error creating inbound order:', err);
    res.status(500).json({ success: false, message: err.message || 'Lỗi tạo đơn hàng nhập.' });
  }
});

// 5. Get all Inbound Orders
app.get('/api/inbound', (req, res) => {
  const db = loadDB();
  res.json({ success: true, data: db.inboundOrders });
});

// 6. Get Single Inbound Order by ID
app.get('/api/inbound/:id', (req, res) => {
  const db = loadDB();
  const order = db.inboundOrders.find(o => o.id === req.params.id);
  if (!order) {
    return res.status(404).json({ success: false, message: 'Không tìm thấy đơn hàng nhập.' });
  }
  res.json({ success: true, data: order });
});

// 7. Receive Goods (Xác nhận nhận hàng vào vị trí mặc định InboundQC)
// YÊU CẦU: "Sau khi nhận hàng vào kho sẽ nhận vào vị trí mặc định là InboundQC"
// "sau khi nhận hàng xong thì trạng thái chuyển thành đã nhận hàng và được phép putaway"
app.post('/api/inbound/:id/receive', (req, res) => {
  try {
    const { receivedItems, receiverNote } = req.body;
    const db = loadDB();
    const orderIndex = db.inboundOrders.findIndex(o => o.id === req.params.id);

    if (orderIndex === -1) {
      return res.status(404).json({ success: false, message: 'Không tìm thấy đơn hàng nhập.' });
    }

    const order = db.inboundOrders[orderIndex];

    if (order.status !== 'CREATED' && order.status !== 'RECEIVED') {
      return res.status(400).json({ success: false, message: `Đơn hàng đang ở trạng thái "${order.statusLabel}", không thể nhận lại.` });
    }

    let totalReceived = 0;
    if (receivedItems && Array.isArray(receivedItems)) {
      order.items = order.items.map(orig => {
        const updated = receivedItems.find(r => String(r.sku) === String(orig.sku));
        const recQty = updated ? (parseInt(updated.receivedQty, 10) || 0) : orig.expectedQty;
        totalReceived += recQty;
        return {
          ...orig,
          receivedQty: recQty,
          batchNo: updated?.batchNo || orig.batchNo,
          expiryDate: updated?.expiryDate || orig.expiryDate
        };
      });
    } else {
      order.items = order.items.map(orig => {
        const qty = orig.expectedQty;
        totalReceived += qty;
        return { ...orig, receivedQty: qty };
      });
    }

    order.totalReceivedQty = totalReceived;
    order.receiverNote = receiverNote || 'Nhận hàng tại cửa nhận, chuyển về InboundQC chờ kiểm tra';
    order.receivedAt = new Date().toISOString();

    // Trạng thái chuyển thành "Đã nhận hàng" & ĐƯỢC PHÉP PUTAWAY
    order.status = 'RECEIVED';
    order.statusLabel = 'Đã nhận hàng';
    order.putawayAllowed = true;
    order.receivingLocation = 'InboundQC';

    // Cập nhật tồn kho vào vị trí mặc định InboundQC
    order.items.forEach(it => {
      const existingInboundQC = db.inventory.find(
        inv => inv.sku === it.sku && inv.location === 'InboundQC' && inv.batchNo === it.batchNo
      );

      if (existingInboundQC) {
        existingInboundQC.quantity += it.receivedQty;
        existingInboundQC.lastUpdated = new Date().toISOString();
      } else {
        db.inventory.push({
          id: `INV-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
          sku: it.sku,
          itemName: it.itemName,
          uom: it.uom,
          location: 'InboundQC', // YÊU CẦU: Nhận vào vị trí mặc định InboundQC
          batchNo: it.batchNo,
          expiryDate: it.expiryDate,
          quantity: it.receivedQty,
          lastUpdated: new Date().toISOString()
        });
      }

      // Record Stock Movement to InboundQC
      db.stockMovements.push({
        id: `MOV-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        timestamp: new Date().toISOString(),
        type: 'INBOUND_RECEIVE',
        typeLabel: 'Nhận hàng vào kho (InboundQC)',
        orderId: order.id,
        sku: it.sku,
        itemName: it.itemName,
        fromLocation: 'CỬA NHẬN HÀNG (RECEIVING DOCK)',
        toLocation: 'InboundQC',
        quantity: it.receivedQty,
        batchNo: it.batchNo,
        expiryDate: it.expiryDate,
        note: `Nhận hàng đơn ${order.orderCode} vào vị trí InboundQC`
      });
    });

    // Update InboundQC location occupancy
    const inboundQcLoc = db.locations.find(l => l.code === 'InboundQC');
    if (inboundQcLoc) {
      inboundQcLoc.current = (inboundQcLoc.current || 0) + totalReceived;
    }

    db.inboundOrders[orderIndex] = order;
    saveDB(db);

    res.json({
      success: true,
      message: `Đơn ${order.orderCode} đã nhận ${totalReceived} cái vào vị trí mặc định [InboundQC]. Trạng thái: "Đã nhận hàng" & ĐƯỢC PHÉP PUTAWAY!`,
      data: order
    });
  } catch (err) {
    console.error('Error receiving goods:', err);
    res.status(500).json({ success: false, message: err.message || 'Lỗi khi nhận hàng.' });
  }
});

// 8. Putaway: Get items ready for putaway from InboundQC
app.get('/api/putaway/ready', (req, res) => {
  const db = loadDB();
  const readyOrders = db.inboundOrders.filter(o => o.putawayAllowed && o.status !== 'PUTAWAY_COMPLETED');

  const tasks = [];
  readyOrders.forEach(order => {
    order.items.forEach(it => {
      const remainingQty = (it.receivedQty || it.expectedQty) - (it.putawayQty || 0);
      if (remainingQty > 0) {
        // Find suggested PL.Tam locations (e.g. 1 location per 50 units)
        const locCountNeeded = Math.max(1, Math.ceil(remainingQty / 45));
        const suggestedLocs = suggestPutawayLocations(locCountNeeded, db.locations, db.inventory);

        tasks.push({
          orderId: order.id,
          orderCode: order.orderCode,
          sku: it.sku,
          itemName: it.itemName,
          uom: it.uom,
          batchNo: it.batchNo,
          expiryDate: it.expiryDate,
          receivedQty: it.receivedQty,
          putawayQty: it.putawayQty || 0,
          remainingQty,
          currentLocation: 'InboundQC',
          suggestedLocations: suggestedLocs,
          putawayAllocations: it.putawayAllocations || []
        });
      }
    });
  });

  res.json({
    success: true,
    data: {
      orders: readyOrders,
      tasks,
      allowedLocationRange: {
        from: 'PL.Tam.01',
        to: 'PL.Tam.100'
      }
    }
  });
});

// 9. Putaway: Execute Putaway with Multi-location Allocation per SKU
// YÊU CẦU:
// "Khi putaway theo SKU sẽ có trường hợp 1 SKU với số lượng nhiều sẽ chiếm nhiều vị trí
// nên khi cập nhật vào tồn kho sẽ cần phải cập nhật đầy đủ vị trí và số lượng tương ứng tại mỗi vị trí"
// "Vị trí được phép putaway được phép nhập là PL.Tam.01 đến PL.Tam.100"
app.post('/api/putaway/execute', (req, res) => {
  try {
    const { orderId, items } = req.body;
    // items: [ { sku, allocations: [ { locationCode: 'PL.Tam.01', qty: 40 }, { locationCode: 'PL.Tam.02', qty: 45 } ] } ]
    if (!orderId || !items || !Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ success: false, message: 'Dữ liệu phân bổ cất kho không hợp lệ.' });
    }

    const db = loadDB();
    const orderIndex = db.inboundOrders.findIndex(o => o.id === orderId);
    if (orderIndex === -1) {
      return res.status(404).json({ success: false, message: 'Không tìm thấy đơn hàng nhập.' });
    }

    const order = db.inboundOrders[orderIndex];
    if (!order.putawayAllowed) {
      return res.status(400).json({ success: false, message: 'Đơn hàng này chưa nhận hàng, không được phép Putaway!' });
    }

    const errors = [];
    const executionPlan = [];

    // 1. Validate all requested allocations
    items.forEach(itPlan => {
      const sku = String(itPlan.sku);
      const orderItem = order.items.find(i => String(i.sku) === sku);
      if (!orderItem) {
        errors.push(`SKU ${sku} không thuộc đơn hàng này.`);
        return;
      }

      const remainingNeed = (orderItem.receivedQty || orderItem.expectedQty) - (orderItem.putawayQty || 0);
      let totalAllocatedThisSku = 0;

      if (!itPlan.allocations || !Array.isArray(itPlan.allocations) || itPlan.allocations.length === 0) {
        return;
      }

      itPlan.allocations.forEach((alloc, aIdx) => {
        const rawLoc = String(alloc.locationCode || '').trim();
        const normLoc = normalizePutawayLocation(rawLoc);
        const qty = parseInt(alloc.qty, 10);

        if (!qty || qty <= 0) return;

        // YÊU CẦU: Vị trí được phép putaway là PL.Tam.01 đến PL.Tam.100
        if (!isValidPutawayLocation(normLoc)) {
          errors.push(`Vị trí "${rawLoc}" không hợp lệ! Vị trí cất kho chỉ được phép nhập từ PL.Tam.01 đến PL.Tam.100.`);
          return;
        }

        totalAllocatedThisSku += qty;
        executionPlan.push({
          orderItem,
          sku,
          locationCode: normLoc,
          qty,
          batchNo: itPlan.batchNo || orderItem.batchNo,
          expiryDate: itPlan.expiryDate || orderItem.expiryDate
        });
      });

      if (totalAllocatedThisSku > remainingNeed) {
        errors.push(`Mã SKU ${sku} (${orderItem.itemName}): Tổng số lượng cất (${totalAllocatedThisSku}) vượt quá số lượng cần cất (${remainingNeed}).`);
      }
    });

    if (errors.length > 0) {
      return res.status(400).json({ success: false, message: errors.join(' | ') });
    }

    if (executionPlan.length === 0) {
      return res.status(400).json({ success: false, message: 'Chưa có dòng phân bổ vị trí nào có số lượng > 0.' });
    }

    // 2. Execute Stock Transfers: from InboundQC to respective PL.Tam locations
    executionPlan.forEach(plan => {
      const { orderItem, sku, locationCode, qty, batchNo, expiryDate } = plan;

      // Deduct from InboundQC in inventory
      const inboundQcInv = db.inventory.find(
        inv => inv.sku === sku && inv.location === 'InboundQC' && inv.batchNo === batchNo
      );
      if (inboundQcInv) {
        inboundQcInv.quantity = Math.max(0, inboundQcInv.quantity - qty);
        inboundQcInv.lastUpdated = new Date().toISOString();
      }

      // Add to destination location in inventory
      const destInv = db.inventory.find(
        inv => inv.sku === sku && inv.location === locationCode && inv.batchNo === batchNo
      );
      if (destInv) {
        destInv.quantity += qty;
        destInv.lastUpdated = new Date().toISOString();
      } else {
        db.inventory.push({
          id: `INV-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
          sku,
          itemName: orderItem.itemName,
          uom: orderItem.uom,
          location: locationCode, // PL.Tam.xx
          batchNo,
          expiryDate,
          quantity: qty,
          lastUpdated: new Date().toISOString()
        });
      }

      // Update location current occupancy
      const locObj = db.locations.find(l => l.code === locationCode);
      if (locObj) {
        locObj.current = (locObj.current || 0) + qty;
      }
      const inboundQcLoc = db.locations.find(l => l.code === 'InboundQC');
      if (inboundQcLoc) {
        inboundQcLoc.current = Math.max(0, (inboundQcLoc.current || 0) - qty);
      }

      // Update item in order
      orderItem.putawayQty = (orderItem.putawayQty || 0) + qty;
      if (!orderItem.putawayAllocations) orderItem.putawayAllocations = [];
      orderItem.putawayAllocations.push({
        locationCode,
        qty,
        timestamp: new Date().toISOString()
      });

      // Record Stock Movement: InboundQC -> PL.Tam.xx
      db.stockMovements.push({
        id: `MOV-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        timestamp: new Date().toISOString(),
        type: 'PUTAWAY',
        typeLabel: 'Putaway Cất kho (SKU đa vị trí)',
        orderId: order.id,
        sku,
        itemName: orderItem.itemName,
        fromLocation: 'InboundQC',
        toLocation: locationCode,
        quantity: qty,
        batchNo,
        expiryDate,
        note: `Cất kho đơn ${order.orderCode}: ${qty} cái vào kệ ${locationCode}`
      });
    });

    // Clean up empty InboundQC rows
    db.inventory = db.inventory.filter(inv => inv.quantity > 0);

    // Check if order is completely put away
    const allCompleted = order.items.every(it => (it.putawayQty || 0) >= (it.receivedQty || it.expectedQty));
    if (allCompleted) {
      order.status = 'PUTAWAY_COMPLETED';
      order.statusLabel = 'Đã cất kho hoàn tất';
      order.putawayAt = new Date().toISOString();
    } else {
      order.status = 'PUTAWAY_IN_PROGRESS';
      order.statusLabel = 'Đang cất kho';
    }

    db.inboundOrders[orderIndex] = order;
    saveDB(db);

    res.json({
      success: true,
      message: `Putaway thành công! Đã chuyển hàng từ InboundQC vào các vị trí PL.Tam và cập nhật tồn kho chính xác.`,
      data: {
        order,
        status: order.status
      }
    });
  } catch (err) {
    console.error('Error executing putaway:', err);
    res.status(500).json({ success: false, message: err.message || 'Lỗi khi thực hiện Putaway.' });
  }
});

// 10. Inventory API: Real-time stock levels
app.get('/api/inventory', (req, res) => {
  const db = loadDB();
  const { search, location, sku } = req.query;

  let items = db.inventory.filter(i => i.quantity > 0);

  if (search) {
    const s = search.toLowerCase();
    items = items.filter(i => i.sku.toLowerCase().includes(s) || i.itemName.toLowerCase().includes(s) || i.location.toLowerCase().includes(s));
  }
  if (location) {
    items = items.filter(i => i.location === location);
  }
  if (sku) {
    items = items.filter(i => i.sku === sku);
  }

  // Aggregate by SKU
  const skuSummaryMap = new Map();
  items.forEach(it => {
    if (!skuSummaryMap.has(it.sku)) {
      skuSummaryMap.set(it.sku, {
        sku: it.sku,
        itemName: it.itemName,
        uom: it.uom,
        totalQuantity: 0,
        inboundQcQty: 0,
        locations: new Map() // location -> qty
      });
    }
    const entry = skuSummaryMap.get(it.sku);
    entry.totalQuantity += it.quantity;
    if (it.location === 'InboundQC') {
      entry.inboundQcQty += it.quantity;
    }
    const currentLocQty = entry.locations.get(it.location) || 0;
    entry.locations.set(it.location, currentLocQty + it.quantity);
  });

  const skuSummary = Array.from(skuSummaryMap.values()).map(v => ({
    sku: v.sku,
    itemName: v.itemName,
    uom: v.uom,
    totalQuantity: v.totalQuantity,
    inboundQcQty: v.inboundQcQty,
    locationBreakdown: Array.from(v.locations.entries()).map(([loc, q]) => ({ location: loc, qty: q }))
  }));

  res.json({
    success: true,
    data: {
      details: items,
      skuSummary,
      totalUnits: items.reduce((sum, i) => sum + i.quantity, 0),
      totalSkus: skuSummary.length
    }
  });
});

// 11. Inventory: Stock Movements / Thẻ kho
app.get('/api/inventory/movements', (req, res) => {
  const db = loadDB();
  res.json({
    success: true,
    data: db.stockMovements.slice().reverse()
  });
});

// 12. Locations Master API
app.get('/api/locations', (req, res) => {
  const db = loadDB();
  const enrichedLocations = db.locations.map(loc => {
    const storedItems = db.inventory.filter(i => i.location === loc.code && i.quantity > 0);
    const totalStored = storedItems.reduce((sum, i) => sum + i.quantity, 0);
    return {
      ...loc,
      current: totalStored,
      occupancyPercent: loc.capacity > 0 ? Math.min(100, Math.round((totalStored / loc.capacity) * 100)) : 0,
      storedItems
    };
  });

  res.json({
    success: true,
    data: enrichedLocations
  });
});

// 13. Outbound Module (Sheet 'OB')
app.post('/api/outbound/load-default', (req, res) => {
  try {
    if (!fs.existsSync(DEFAULT_EXCEL_FILE)) {
      return res.status(404).json({ success: false, message: 'Không tìm thấy file mẫu.' });
    }
    const workbook = XLSX.readFile(DEFAULT_EXCEL_FILE);
    const parsed = parseOutboundWorkbook(workbook);

    const db = loadDB();
    let addedCount = 0;
    parsed.orders.forEach(ord => {
      if (!db.outboundOrders.some(o => o.orderNumber === ord.orderNumber)) {
        db.outboundOrders.push(ord);
        addedCount++;
      }
    });
    saveDB(db);

    res.json({
      success: true,
      message: `Đã nạp ${addedCount} đơn xuất hàng từ sheet OB vào hệ thống!`,
      data: db.outboundOrders
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/api/outbound', (req, res) => {
  const db = loadDB();
  res.json({ success: true, data: db.outboundOrders });
});

// 14. Reset System Data
app.post('/api/system/reset-demo', (req, res) => {
  const cleanDB = {
    inboundOrders: [],
    putawayTasks: [],
    inventory: [],
    stockMovements: [],
    outboundOrders: [],
    locations: generateDefaultLocations()
  };
  saveDB(cleanDB);
  res.json({ success: true, message: 'Đã thiết lập lại dữ liệu hệ thống Honeywell WMS (InboundQC & PL.Tam.01 đến PL.Tam.100).' });
});

// Export for Vercel serverless (api/index.js)
module.exports = app;

// Start Server only when run directly (node server.js)
if (require.main === module) {
  const server = app.listen(PORT, () => {
    console.log(`====================================================`);
    console.log(`🚀 Honeywell WMS Server đang chạy tại: http://localhost:${PORT}`);
    console.log(`💾 Storage: ${USE_REDIS ? 'Upstash Redis' : 'data/db.json (local)'}`);
    console.log(`====================================================`);
  });

  server.on('error', (e) => {
    if (e.code === 'EADDRINUSE') {
      const nextPort = Number(PORT) + 1;
      app.listen(nextPort, () => {
        console.log(`🚀 Honeywell WMS Server đang chạy tại: http://localhost:${nextPort}`);
      });
    }
  });
}

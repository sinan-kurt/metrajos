import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';
import html2pdf from 'html2pdf.js';

const state = {
  workbooks: [null, null],
  selectedSheets: [null, null],
  files: [null, null],
  parsedItems: [null, null],
  columnMappings: [null, null],
  diffs: [],
  allDiffs: [],
  summary: null,
  expandedRow: null,
  importLog: [],
  feedback: [],
  isDemo: false,
};

const metrics = {
  sessionId: crypto.randomUUID(),
  startTime: Date.now(),
  events: [],

  track(event, data = {}) {
    this.events.push({
      ts: Date.now(),
      elapsed: Date.now() - this.startTime,
      event,
      ...data
    });
  }
};

const CATEGORIES = {
  added: { label: 'Eklenen', icon: '+' },
  removed: { label: 'Kaldırılan', icon: '−' },
  increased: { label: 'Artan', icon: '↑' },
  decreased: { label: 'Azalan', icon: '↓' },
  unchanged: { label: 'Değişmeyen', icon: '=' },
  name_changed: { label: 'İsim Değişti', icon: '~' },
  unit_changed: { label: 'Birim Değişti', icon: '⚠' },
  duplicate_suspect: { label: 'Şüpheli Kopya', icon: '⚡' },
};

const COL_PATTERNS = {
  positionCode: ['POZ_NO', 'POZNO', 'POZ NO', 'POZ', 'POSITION CODE', 'ITEM CODE', 'POZ NUMARASI', 'SIRA NO'],
  itemName: ['IMALAT_ADI', 'İMALAT_ADI', 'IMALATADI', 'İMALATADI', 'IMALAT', 'İMALAT', 'AÇIKLAMA', 'DESCRIPTION', 'MALZEME', 'TANIM', 'İŞ KALEMİ', 'IS KALEMI'],
  unit: ['BIRIM', 'BİRİM', 'UNIT', 'ÜNITE', 'ÜNİTE', 'OLCU BIRIMI', 'ÖLÇÜ BİRİMİ'],
  quantity: ['MIKTAR', 'MİKTAR', 'QUANTITY', 'QTY', 'ADET', 'TOPLAM MIKTAR'],
  unitPrice: ['BIRIM_FIYAT', 'BİRİMFİYAT', 'BIRIMFIYAT', 'BIRIM FIYAT', 'BİRİM FİYAT', 'UNIT PRICE', 'FIYAT', 'FİYAT'],
  discipline: ['DISIPLIN', 'DİSİPLİN', 'DISCIPLINE', 'BRANŞ', 'BRANS'],
  zone: ['MAHAL', 'ZONE', 'LOCATION', 'YER', 'MEKAN'],
};

const UNIT_MAP = {
  'm2': 'm²', 'm²': 'm²', 'metrekare': 'm²', 'metre kare': 'm²', 'sqm': 'm²',
  'm3': 'm³', 'm³': 'm³', 'metreküp': 'm³', 'metre küp': 'm³', 'cbm': 'm³',
  'kg': 'kg', 'kilogram': 'kg', 'kilo': 'kg',
  't': 'ton', 'ton': 'ton', 'tn': 'ton',
  'adet': 'adet', 'ad': 'adet', 'piece': 'adet', 'pcs': 'adet',
  'm': 'm', 'metre': 'm', 'mt': 'm',
  'cm': 'cm', 'mm': 'mm',
  'lt': 'lt', 'litre': 'lt', 'liter': 'lt', 'l': 'lt',
  'tk': 'tk', 'takım': 'tk', 'set': 'set',
  'saat': 'saat', 'h': 'saat',
  'gün': 'gün', 'gun': 'gün',
};

const DEMO_SCENARIOS = {
  revision: {
    name: 'Revizyon Farkı',
    description: 'Klasik revizyon karşılaştırması — miktar değişiklikleri, isim güncellemeleri, birim dönüşümleri',
    icon: '📋',
    v1: [
      { positionCode: '15.150.001', itemName: 'C30 Beton', unit: 'm³', quantity: 120, unitPrice: 2900, discipline: 'Betonarme', zone: 'Temel', rowIndex: 2 },
      { positionCode: '15.210.005', itemName: 'Nervürlü Donatı', unit: 'kg', quantity: 18500, unitPrice: 14, discipline: 'Betonarme', zone: 'Temel', rowIndex: 3 },
      { positionCode: '15.310.010', itemName: 'Gazbeton Duvar 19 cm', unit: 'm²', quantity: 640, unitPrice: 780, discipline: 'Duvar', zone: 'Dış Duvar', rowIndex: 4 },
      { positionCode: '15.410.003', itemName: 'İç Cephe Sıva', unit: 'm²', quantity: 1180, unitPrice: 45, discipline: 'Sıva', zone: 'İç Duvar', rowIndex: 5 },
      { positionCode: '15.510.002', itemName: 'Seramik Kaplama', unit: 'm²', quantity: 420, unitPrice: 380, discipline: 'Döşeme', zone: 'Döşeme', rowIndex: 6 },
      { positionCode: '15.610.001', itemName: 'Alçıpan Tavan', unit: 'm²', quantity: 360, unitPrice: 520, discipline: 'Tavan', zone: 'Tavan', rowIndex: 7 },
      { positionCode: '15.710.004', itemName: 'Boyası', unit: 'm²', quantity: 1250, unitPrice: 28, discipline: 'Boyama', zone: 'İç', rowIndex: 8 },
      { positionCode: '15.810.006', itemName: 'PVC Doğrama', unit: 'm²', quantity: 96, unitPrice: 1450, discipline: 'Doğrama', zone: 'Cephe', rowIndex: 9 },
      { positionCode: '15.910.002', itemName: 'Ahşap Kapı', unit: 'adet', quantity: 42, unitPrice: 850, discipline: 'Doğrama', zone: 'İç', rowIndex: 10 },
      { positionCode: '16.010.003', itemName: 'Kablo Tavası', unit: 'm', quantity: 180, unitPrice: 95, discipline: 'Elektrik', zone: 'Elektrik', rowIndex: 11 },
      { positionCode: '16.110.001', itemName: 'Klima Seti', unit: 'set', quantity: 18, unitPrice: 12000, discipline: 'HVAC', zone: 'Kuru Alaş', rowIndex: 12 },
      { positionCode: '16.210.005', itemName: 'Beton Bordür', unit: 'm', quantity: 110, unitPrice: 180, discipline: 'Dış', zone: 'Bahçe', rowIndex: 13 },
    ],
    v2: [
      { positionCode: '15.150.001', itemName: 'C30 Beton', unit: 'm³', quantity: 128, unitPrice: 2900, discipline: 'Betonarme', zone: 'Temel', rowIndex: 2 },
      { positionCode: '15.210.005', itemName: 'Nervürlü Donatı', unit: 'kg', quantity: 19150, unitPrice: 14, discipline: 'Betonarme', zone: 'Temel', rowIndex: 3 },
      { positionCode: '15.310.010', itemName: '19 cm Gazbeton Duvar', unit: 'm²', quantity: 640, unitPrice: 780, discipline: 'Duvar', zone: 'Dış Duvar', rowIndex: 4 },
      { positionCode: '15.410.003', itemName: 'İç Cephe Sıva', unit: 'm²', quantity: 1120, unitPrice: 45, discipline: 'Sıva', zone: 'İç Duvar', rowIndex: 5 },
      { positionCode: '15.510.002', itemName: 'Seramik Kaplama', unit: 'm²', quantity: 455, unitPrice: 380, discipline: 'Döşeme', zone: 'Döşeme', rowIndex: 6 },
      { positionCode: '15.610.001', itemName: 'Alçıpan Tavan', unit: 'm²', quantity: 360, unitPrice: 520, discipline: 'Tavan', zone: 'Tavan', rowIndex: 7 },
      { positionCode: '15.710.004', itemName: 'Boyası', unit: 'm²', quantity: 1325, unitPrice: 28, discipline: 'Boyama', zone: 'İç', rowIndex: 8 },
      { positionCode: '15.810.006', itemName: 'PVC Doğrama', unit: 'adet', quantity: 28, unitPrice: 1450, discipline: 'Doğrama', zone: 'Cephe', rowIndex: 9 },
      { positionCode: '15.910.002', itemName: 'Ahşap Kapı', unit: 'adet', quantity: 40, unitPrice: 850, discipline: 'Doğrama', zone: 'İç', rowIndex: 10 },
      { positionCode: '16.010.003', itemName: 'Kablo Tavası', unit: 'm', quantity: 205, unitPrice: 95, discipline: 'Elektrik', zone: 'Elektrik', rowIndex: 11 },
      { positionCode: '16.110.001', itemName: 'Klima Seti', unit: 'set', quantity: 20, unitPrice: 12000, discipline: 'HVAC', zone: 'Kuru Alaş', rowIndex: 12 },
      { positionCode: '16.210.005', itemName: 'Beton Bordür', unit: 'm', quantity: 110, unitPrice: 180, discipline: 'Dış', zone: 'Bahçe', rowIndex: 13 },
    ],
  },
  addremove: {
    name: 'Ekleme / Çıkarma',
    description: 'Yoğun ekleme/silme senaryosu',
    icon: '➕➖',
    v1: [
      { positionCode: '20.010.001', itemName: 'Temel Kazı', unit: 'm³', quantity: 1200, unitPrice: 85, discipline: 'Temel', zone: 'Temel', rowIndex: 2 },
      { positionCode: '20.020.001', itemName: 'Grobeton', unit: 'm³', quantity: 450, unitPrice: 650, discipline: 'Temel', zone: 'Temel', rowIndex: 3 },
      { positionCode: '20.030.001', itemName: 'Betonarme', unit: 'm³', quantity: 320, unitPrice: 2200, discipline: 'Betonarme', zone: 'Temel', rowIndex: 4 },
      { positionCode: '20.040.001', itemName: 'Kalıp İşleri', unit: 'm²', quantity: 2800, unitPrice: 95, discipline: 'Kalıp', zone: 'Temel', rowIndex: 5 },
      { positionCode: '20.050.001', itemName: 'Su Yalıtımı', unit: 'm²', quantity: 1800, unitPrice: 75, discipline: 'Yalıtım', zone: 'Temel', rowIndex: 6 },
      { positionCode: '20.060.001', itemName: 'Isı Yalıtımı', unit: 'm³', quantity: 380, unitPrice: 320, discipline: 'Yalıtım', zone: 'Temel', rowIndex: 7 },
      { positionCode: '20.070.001', itemName: 'Drenaj Borusu', unit: 'm', quantity: 850, unitPrice: 120, discipline: 'Drenaj', zone: 'Temel', rowIndex: 8 },
      { positionCode: '20.080.001', itemName: 'Topraklama', unit: 'adet', quantity: 16, unitPrice: 450, discipline: 'Elektrik', zone: 'Temel', rowIndex: 9 },
      { positionCode: '20.090.001', itemName: 'Yangın Dolabı', unit: 'adet', quantity: 8, unitPrice: 3200, discipline: 'Yangın', zone: 'Ortak', rowIndex: 10 },
      { positionCode: '20.100.001', itemName: 'Asansör Kuyusu', unit: 'adet', quantity: 2, unitPrice: 18000, discipline: 'Mekanik', zone: 'Ortak', rowIndex: 11 },
    ],
    v2: [
      { positionCode: '20.010.001', itemName: 'Temel Kazı', unit: 'm³', quantity: 1400, unitPrice: 85, discipline: 'Temel', zone: 'Temel', rowIndex: 2 },
      { positionCode: '20.020.001', itemName: 'Grobeton', unit: 'm³', quantity: 550, unitPrice: 650, discipline: 'Temel', zone: 'Temel', rowIndex: 3 },
      { positionCode: '20.030.001', itemName: 'Betonarme', unit: 'm³', quantity: 350, unitPrice: 2200, discipline: 'Betonarme', zone: 'Temel', rowIndex: 4 },
      { positionCode: '20.040.001', itemName: 'Kalıp İşleri', unit: 'm²', quantity: 3000, unitPrice: 95, discipline: 'Kalıp', zone: 'Temel', rowIndex: 5 },
      { positionCode: '20.050.001', itemName: 'Su Yalıtımı', unit: 'm²', quantity: 1950, unitPrice: 75, discipline: 'Yalıtım', zone: 'Temel', rowIndex: 6 },
      { positionCode: '20.110.001', itemName: 'Sondaj Kazığı', unit: 'adet', quantity: 48, unitPrice: 8500, discipline: 'Temel', zone: 'Temel', rowIndex: 7 },
      { positionCode: '20.120.001', itemName: 'Jet Grout', unit: 'm³', quantity: 220, unitPrice: 1800, discipline: 'Temel', zone: 'Temel', rowIndex: 8 },
      { positionCode: '20.130.001', itemName: 'Palplanş', unit: 'ton', quantity: 140, unitPrice: 5200, discipline: 'Temel', zone: 'Temel', rowIndex: 9 },
      { positionCode: '20.090.001', itemName: 'Yangın Dolabı', unit: 'adet', quantity: 9, unitPrice: 3200, discipline: 'Yangın', zone: 'Ortak', rowIndex: 10 },
      { positionCode: '20.140.001', itemName: 'Zemin İyileştirme', unit: 'm³', quantity: 600, unitPrice: 420, discipline: 'Temel', zone: 'Temel', rowIndex: 11 },
      { positionCode: '20.150.001', itemName: 'Ankraj Sistemi', unit: 'adet', quantity: 32, unitPrice: 2800, discipline: 'Temel', zone: 'Temel', rowIndex: 12 },
    ],
  },
  unitname: {
    name: 'Birim & İsim Karmaşası',
    description: 'Birim dönüşümleri ve isim farklılıkları',
    icon: '⚠️',
    v1: [
      { positionCode: '30.010.001', itemName: 'Çelik Profil HEA200', unit: 'm', quantity: 850, unitPrice: 950, discipline: 'Çelik', zone: 'Yapı', rowIndex: 2 },
      { positionCode: '30.020.001', itemName: 'Alüminyum Doğrama', unit: 'm²', quantity: 380, unitPrice: 2100, discipline: 'Doğrama', zone: 'Cephe', rowIndex: 3 },
      { positionCode: '30.030.001', itemName: 'Epoksi Zemin Kaplama', unit: 'm²', quantity: 2400, unitPrice: 180, discipline: 'Döşeme', zone: 'Endüstri', rowIndex: 4 },
      { positionCode: '30.040.001', itemName: 'Mermer Merdiven Basamak', unit: 'adet', quantity: 28, unitPrice: 4200, discipline: 'Dekorasyon', zone: 'Merdiven', rowIndex: 5 },
      { positionCode: '30.050.001', itemName: 'Yangın Merdiveni Korkuluk', unit: 'm', quantity: 145, unitPrice: 850, discipline: 'Yangın', zone: 'Yangın Merdiveni', rowIndex: 6 },
      { positionCode: '30.060.001', itemName: 'Gümrük Demiri', unit: 'kg', quantity: 12500, unitPrice: 18, discipline: 'Demirçilik', zone: 'Pencere', rowIndex: 7 },
      { positionCode: '30.070.001', itemName: 'Cam Fiber Levha', unit: 'm²', quantity: 350, unitPrice: 280, discipline: 'Yalıtım', zone: 'Duvar', rowIndex: 8 },
      { positionCode: '30.080.001', itemName: 'Çatı Kaplama Malzeme', unit: 't', quantity: 45, unitPrice: 3100, discipline: 'Çatı', zone: 'Çatı', rowIndex: 9 },
    ],
    v2: [
      { positionCode: '30.010.001', itemName: 'Çelik Profil HEA200', unit: 'ton', quantity: 6.8, unitPrice: 950000, discipline: 'Çelik', zone: 'Yapı', rowIndex: 2 },
      { positionCode: '30.020.001', itemName: 'Alüminyum Doğrama', unit: 'adet', quantity: 115, unitPrice: 7000, discipline: 'Doğrama', zone: 'Cephe', rowIndex: 3 },
      { positionCode: '30.030.001', itemName: 'Epoksi Kaplama (Zemin)', unit: 'm²', quantity: 2400, unitPrice: 180, discipline: 'Döşeme', zone: 'Endüstri', rowIndex: 4 },
      { positionCode: '30.040.001', itemName: 'Doğal Taş Merdiven Kaplaması', unit: 'adet', quantity: 28, unitPrice: 4200, discipline: 'Dekorasyon', zone: 'Merdiven', rowIndex: 5 },
      { positionCode: '30.050.001', itemName: 'Yangın Merdiveni Korkuluğu', unit: 'm', quantity: 145, unitPrice: 850, discipline: 'Yangın', zone: 'Yangın Merdiveni', rowIndex: 6 },
      { positionCode: '30.060.001', itemName: 'Gümrük Demiri', unit: 'kg', quantity: 13100, unitPrice: 18, discipline: 'Demirçilik', zone: 'Pencere', rowIndex: 7 },
      { positionCode: '30.070.001', itemName: 'Cam Fiber Levha', unit: 'm²', quantity: 350, unitPrice: 280, discipline: 'Yalıtım', zone: 'Duvar', rowIndex: 8 },
      { positionCode: '30.080.001', itemName: 'Çatı Kaplama Malzeme', unit: 't', quantity: 50, unitPrice: 3100, discipline: 'Çatı', zone: 'Çatı', rowIndex: 9 },
    ],
  },
};

function normalizeUnit(val) {
  if (!val) return '';
  const s = String(val).toLowerCase().trim();
  return UNIT_MAP[s] || s;
}

function normalizeItemName(val) {
  if (!val) return '';
  return String(val).toLowerCase().trim();
}

function parseNumber(val) {
  if (typeof val === 'number') return val;
  let s = String(val || '').trim();
  if (!s) return NaN;
  s = s.replace(/[\s₺TL$€]/gi, '');
  if (/^\d{1,3}(\.\d{3})+,\d+$/.test(s)) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (/^\d+,\d+$/.test(s)) {
    s = s.replace(',', '.');
  }
  return parseFloat(s);
}

function findColumn(headers, patterns) {
  if (!headers || !patterns) return -1;
  const normalized = headers.map(h => String(h).toUpperCase().trim());

  for (const pattern of patterns) {
    const idx = normalized.indexOf(pattern.toUpperCase());
    if (idx >= 0) return idx;
  }

  for (const pattern of patterns) {
    if (pattern.length < 4) continue;
    for (let i = 0; i < normalized.length; i++) {
      if (normalized[i].includes(pattern.toUpperCase())) return i;
    }
  }

  for (const pattern of patterns) {
    for (let i = 0; i < normalized.length; i++) {
      if (normalized[i].length < 4) continue;
      if (pattern.toUpperCase().includes(normalized[i])) return i;
    }
  }

  return -1;
}

function calculateSimilarity(a, b) {
  if (!a || !b) return 0;
  const aLower = normalizeItemName(a);
  const bLower = normalizeItemName(b);
  if (aLower === bLower) return 1;

  const aWords = aLower.split(/\s+/);
  const bWords = bLower.split(/\s+/);
  const common = aWords.filter(w => bWords.includes(w)).length;
  const total = Math.max(aWords.length, bWords.length);

  return total > 0 ? common / total : 0;
}

function detectDuplicateRows(items) {
  const seen = new Map();
  const duplicates = [];
  for (const item of items) {
    const key = [item.positionCode || '', normalizeItemName(item.itemName), item.quantity, normalizeUnit(item.unit)].join('||');
    if (seen.has(key)) {
      duplicates.push({ item, firstOccurrence: seen.get(key) });
    } else {
      seen.set(key, item.rowIndex);
    }
  }
  return duplicates;
}

['dropV1', 'dropV2'].forEach((id, idx) => {
  const el = document.getElementById(id);
  el.addEventListener('dragover', e => {
    e.preventDefault();
    el.classList.add('dragover');
  });
  el.addEventListener('dragleave', () => el.classList.remove('dragover'));
  el.addEventListener('drop', e => {
    e.preventDefault();
    el.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      const input = document.getElementById('fileV' + (idx + 1));
      input.files = e.dataTransfer.files;
      handleFile(idx + 1, input);
    }
  });
});

function loadDemo(scenarioKey) {
  metrics.track('demo_load', { scenario: scenarioKey });
  const scenario = DEMO_SCENARIOS[scenarioKey];
  if (!scenario) return;

  state.isDemo = true;
  state.parsedItems[0] = scenario.v1;
  state.parsedItems[1] = scenario.v2;

  state.columnMappings[0] = {
    headers: ['positionCode', 'itemName', 'unit', 'quantity', 'unitPrice', 'discipline', 'zone'],
    mapping: { positionCode: 0, itemName: 1, unit: 2, quantity: 3, unitPrice: 4, discipline: 5, zone: 6 }
  };
  state.columnMappings[1] = {
    headers: ['positionCode', 'itemName', 'unit', 'quantity', 'unitPrice', 'discipline', 'zone'],
    mapping: { positionCode: 0, itemName: 1, unit: 2, quantity: 3, unitPrice: 4, discipline: 5, zone: 6 }
  };

  goToStep(3);
}

function handleFile(version, input) {
  const file = input.files[0];
  if (!file) return;

  metrics.track('file_upload', { version, fileName: file.name, fileSize: file.size });

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const workbook = XLSX.read(e.target.result, { type: 'array' });
      state.workbooks[version - 1] = workbook;
      state.files[version - 1] = file;

      const sheetSelect = document.getElementById('sheetsV' + version);
      sheetSelect.innerHTML = '';
      workbook.SheetNames.forEach(name => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.text = name;
        sheetSelect.appendChild(opt);
      });

      state.selectedSheets[version - 1] = workbook.SheetNames[0];

      const infoDiv = document.getElementById('infoV' + version);
      infoDiv.innerHTML = `<strong>${file.name}</strong><br>${file.size} bytes<br>${workbook.SheetNames.length} sheet${workbook.SheetNames.length > 1 ? 's' : ''}`;
      infoDiv.style.display = 'block';

      document.getElementById('sheetSelV' + version).style.display = 'block';
      document.getElementById('dropV' + version).classList.add('loaded');

      checkUploadComplete();
    } catch (err) {
      alert('Hata: Dosya yüklenemedi. Excel dosyası olduğundan emin olun.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function selectSheet(version) {
  const sel = document.getElementById('sheetsV' + version);
  state.selectedSheets[version - 1] = sel.value;
}

function checkUploadComplete() {
  const btn = document.getElementById('btnContinueStep1');
  btn.disabled = !state.workbooks[0] || !state.workbooks[1];
}

function parseSheet(version) {
  const workbook = state.workbooks[version - 1];
  if (!workbook) return null;
  const sheetName = state.selectedSheets[version - 1];
  if (!sheetName) return null;
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;

  const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  if (data.length === 0) return null;

  const headers = Object.keys(data[0]);

  const mapping = {};
  mapping.positionCode = findColumn(headers, COL_PATTERNS.positionCode);
  mapping.itemName = findColumn(headers, COL_PATTERNS.itemName);
  mapping.unit = findColumn(headers, COL_PATTERNS.unit);
  mapping.quantity = findColumn(headers, COL_PATTERNS.quantity);
  mapping.unitPrice = findColumn(headers, COL_PATTERNS.unitPrice);
  mapping.discipline = findColumn(headers, COL_PATTERNS.discipline);
  mapping.zone = findColumn(headers, COL_PATTERNS.zone);

  state.columnMappings[version - 1] = { headers, mapping };
  return { data, headers, mapping };
}

function showMappingUI() {
  state.importLog = [];
  const p1 = parseSheet(1);
  const p2 = parseSheet(2);

  if (!p1 || !p2) {
    alert('Hata: Excel sayfaları boş');
    return;
  }

  const fields = [
    { key: 'itemName', label: 'İmalat Adı', required: true },
    { key: 'unit', label: 'Birim', required: true },
    { key: 'quantity', label: 'Miktar', required: true },
    { key: 'positionCode', label: 'Poz No', required: false },
    { key: 'unitPrice', label: 'Birim Fiyat', required: false },
    { key: 'discipline', label: 'Disiplin', required: false },
    { key: 'zone', label: 'Mahal/Yer', required: false },
  ];

  let html = '<tr><th>Alan</th><th>V1 Sütun</th><th>V2 Sütun</th><th>Durum</th></tr>';

  fields.forEach(field => {
    const v1Idx = p1.mapping[field.key];
    const v2Idx = p2.mapping[field.key];
    const v1Col = v1Idx >= 0 ? p1.headers[v1Idx] : '-';
    const v2Col = v2Idx >= 0 ? p2.headers[v2Idx] : '-';
    const v1Ok = v1Idx >= 0;
    const v2Ok = v2Idx >= 0;
    const status = v1Ok && v2Ok ? 'ok' : (field.required ? 'error' : 'warning');
    const statusText = v1Ok && v2Ok ? '✓ Eşleşti' : (field.required ? '✗ Gerekli' : '⚠ Opsiyonel');

    html += `<tr>
      <td>${field.label}</td>
      <td><select data-version="1" data-field="${field.key}"><option>-</option>${p1.headers.map((h, i) => `<option value="${i}" ${v1Idx === i ? 'selected' : ''}>${h}</option>`).join('')}</select></td>
      <td><select data-version="2" data-field="${field.key}"><option>-</option>${p2.headers.map((h, i) => `<option value="${i}" ${v2Idx === i ? 'selected' : ''}>${h}</option>`).join('')}</select></td>
      <td><span class="status-badge status-${status}">${statusText}</span></td>
    </tr>`;
  });

  document.getElementById('mappingTable').innerHTML = html;

  document.querySelectorAll('#mappingTable select').forEach(sel => {
    sel.addEventListener('change', () => {
      const version = parseInt(sel.dataset.version);
      const field = sel.dataset.field;
      const idx = parseInt(sel.value);
      if (idx >= 0) {
        state.columnMappings[version - 1].mapping[field] = idx;
      }
    });
  });

  const warnings = [];
  const d1 = detectDuplicateRows(itemsFromParsed(p1, state.columnMappings[0].mapping));
  const d2 = detectDuplicateRows(itemsFromParsed(p2, state.columnMappings[1].mapping));
  if (d1.length > 0) warnings.push(`V1'de ${d1.length} şüpheli kopya satır bulundu`);
  if (d2.length > 0) warnings.push(`V2'de ${d2.length} şüpheli kopya satır bulundu`);

  let whtml = '';
  if (warnings.length > 0) {
    whtml = '<div class="warnings-box"><strong>Uyarılar:</strong>';
    warnings.forEach(w => whtml += `<div class="warning-item">${w}</div>`);
    whtml += '</div>';
  }
  document.getElementById('warningsBox').innerHTML = whtml;

  let logHtml = '';
  if (state.importLog.length > 0) {
    logHtml = `<div class="import-log">
      <div class="log-header" onclick="document.querySelector('.log-body').classList.toggle('visible')">⚠️ İçe Aktarma Notları (${state.importLog.length}) ▾</div>
      <div class="log-body">
        <table>`;
    state.importLog.forEach(log => {
      const rowClass = `log-${log.type}`;
      logHtml += `<tr class="${rowClass}"><td>Satır ${log.row}</td><td>${log.reason}</td></tr>`;
    });
    logHtml += `</table></div></div>`;
  }
  document.getElementById('importLogSection').innerHTML = logHtml;

  showPreview(p1, 1);
  showPreview(p2, 2);

  document.getElementById('btnContinueStep2').disabled = false;
}

function showPreview(parsed, version) {
  const { data, headers, mapping } = parsed;
  const items = itemsFromParsed(parsed, mapping);
  const previewData = data.slice(0, 5);

  let html = '<tr>';
  if (mapping.positionCode >= 0) html += `<th>${headers[mapping.positionCode]}</th>`;
  if (mapping.itemName >= 0) html += `<th>${headers[mapping.itemName]}</th>`;
  if (mapping.quantity >= 0) html += `<th>${headers[mapping.quantity]}</th>`;
  if (mapping.unit >= 0) html += `<th>${headers[mapping.unit]}</th>`;
  html += '</tr>';

  previewData.forEach(row => {
    html += '<tr>';
    if (mapping.positionCode >= 0) html += `<td>${row[headers[mapping.positionCode]] || ''}</td>`;
    if (mapping.itemName >= 0) html += `<td>${row[headers[mapping.itemName]] || ''}</td>`;
    if (mapping.quantity >= 0) html += `<td>${row[headers[mapping.quantity]] || ''}</td>`;
    if (mapping.unit >= 0) html += `<td>${row[headers[mapping.unit]] || ''}</td>`;
    html += '</tr>';
  });

  document.getElementById('previewV' + version).innerHTML = html;
}

function itemsFromParsed(parsed, mapping) {
  const { data, headers } = parsed;
  return data.map((row, idx) => ({
    positionCode: mapping.positionCode >= 0 ? String(row[headers[mapping.positionCode]] || '').trim() : '',
    itemName: mapping.itemName >= 0 ? String(row[headers[mapping.itemName]] || '').trim() : '',
    unit: mapping.unit >= 0 ? normalizeUnit(row[headers[mapping.unit]] || '') : '',
    quantity: mapping.quantity >= 0 ? parseNumber(row[headers[mapping.quantity]] || 0) : 0,
    unitPrice: mapping.unitPrice >= 0 ? parseNumber(row[headers[mapping.unitPrice]] || 0) : 0,
    discipline: mapping.discipline >= 0 ? String(row[headers[mapping.discipline]] || '').trim() : '',
    zone: mapping.zone >= 0 ? String(row[headers[mapping.zone]] || '').trim() : '',
    rowIndex: idx + 2,
  }));
}

function matchItems(baseItems, compItems, threshold = 0.72) {
  const matches = [];
  const usedComp = new Set();

  for (const base of baseItems) {
    for (let i = 0; i < compItems.length; i++) {
      if (usedComp.has(i)) continue;
      const comp = compItems[i];
      if (base.positionCode && base.positionCode === comp.positionCode && normalizeItemName(base.itemName) === normalizeItemName(comp.itemName)) {
        matches.push({ base, comp, confidence: 1 });
        usedComp.add(i);
        break;
      }
    }
  }

  for (const base of baseItems) {
    if (matches.some(m => m.base === base)) continue;
    if (!base.positionCode) continue;
    for (let i = 0; i < compItems.length; i++) {
      if (usedComp.has(i)) continue;
      const comp = compItems[i];
      if (base.positionCode === comp.positionCode) {
        const sim = calculateSimilarity(base.itemName, comp.itemName);
        if (sim >= threshold) {
          matches.push({ base, comp, confidence: sim });
          usedComp.add(i);
          break;
        }
      }
    }
  }

  for (const base of baseItems) {
    if (matches.some(m => m.base === base)) continue;
    if (base.positionCode) continue;
    for (let i = 0; i < compItems.length; i++) {
      if (usedComp.has(i)) continue;
      const comp = compItems[i];
      if (!comp.positionCode) {
        const sim = calculateSimilarity(base.itemName, comp.itemName);
        if (sim >= threshold) {
          matches.push({ base, comp, confidence: sim });
          usedComp.add(i);
          break;
        }
      }
    }
  }

  for (const base of baseItems) {
    if (matches.some(m => m.base === base)) continue;
    let bestMatch = null;
    let bestSim = 0;
    for (let i = 0; i < compItems.length; i++) {
      if (usedComp.has(i)) continue;
      const comp = compItems[i];
      const sim = calculateSimilarity(base.itemName, comp.itemName);
      if (sim > bestSim) {
        bestSim = sim;
        bestMatch = i;
      }
    }
    if (bestSim >= threshold && bestMatch !== null) {
      matches.push({ base, comp: compItems[bestMatch], confidence: bestSim });
      usedComp.add(bestMatch);
    }
  }

  for (const base of baseItems) {
    if (!matches.some(m => m.base === base)) {
      matches.push({ base, comp: null, confidence: 0 });
    }
  }

  for (let i = 0; i < compItems.length; i++) {
    if (!usedComp.has(i)) {
      matches.push({ base: null, comp: compItems[i], confidence: 0 });
    }
  }

  return matches;
}

function getExplanation(diff) {
  switch (diff.cat) {
    case 'added':
      return `Bu kalem V1'de bulunmuyor. V2'de yeni eklenmiş. Birim: ${diff.v2Unit}, Miktar: ${diff.v2Qty}`;
    case 'removed':
      return `Bu kalem V2'de bulunmuyor. Revizyon sırasında çıkarılmış. Eski Birim: ${diff.v1Unit}, Eski Miktar: ${diff.v1Qty}`;
    case 'increased':
      const upPercent = ((diff.v2Qty - diff.v1Qty) / diff.v1Qty * 100).toFixed(1);
      return `Miktar ${diff.v1Qty} → ${diff.v2Qty} (↑${upPercent}%). Birim Fiyat: ${diff.v1Item?.unitPrice || 0}₺`;
    case 'decreased':
      const downPercent = ((diff.v2Qty - diff.v1Qty) / diff.v1Qty * 100).toFixed(1);
      return `Miktar ${diff.v1Qty} → ${diff.v2Qty} (↓${Math.abs(downPercent)}%). Birim Fiyat: ${diff.v1Item?.unitPrice || 0}₺`;
    case 'name_changed':
      return `Kalem adı değişti: '${diff.v1Item?.itemName}' → '${diff.v2Item?.itemName}'. Miktar: ${diff.v1Qty} ${diff.v1Unit}`;
    case 'unit_changed':
      return `Birim değişti: ${diff.v1Unit} → ${diff.v2Unit}. Miktarlar karşılaştırılamaz. Manuel inceleme gerekli.`;
    case 'unchanged':
      return `Kalem her iki revizyonda da aynı: ${diff.v1Qty} ${diff.v1Unit}, ${diff.v1Item?.unitPrice || 0}₺/birim`;
    default:
      return '';
  }
}

function generateExecutiveSummary(summary, diffs) {
  const criticalCount = summary.severityCounts.critical || 0;
  const highCount = summary.severityCounts.high || 0;
  const totalMoney = summary.totalAmount;
  
  let moneyText = totalMoney > 0 ? `+${totalMoney.toLocaleString('tr-TR')} ₺ maliyet artışı` 
                : (totalMoney < 0 ? `${totalMoney.toLocaleString('tr-TR')} ₺ maliyet azalışı` : `ciddi bir maliyet etkisi`);

  // Bulunan en büyük bütçe etkisine sahip disiplini tespit edelim (varsa)
  const discImpacts = {};
  diffs.forEach(d => {
    const disc = d.v1Item?.discipline || d.v2Item?.discipline;
    if (disc) {
      discImpacts[disc] = (discImpacts[disc] || 0) + Math.abs(d.amountDelta || 0);
    }
  });
  const topDisc = Object.keys(discImpacts).sort((a,b) => discImpacts[b] - discImpacts[a])[0];
  let discText = topDisc ? `En büyük değişim hareketliliği **${topDisc}** disiplininde gözlemlendi.` : '';

  return `Kıyaslanan revizyonlar sonucunda toplam **${summary.totalDiffs} kalemde** değişiklik (silinme, eklenme, miktar veya isim değişimi) tespit edildi. Bunların **${criticalCount} tanesi kritik**, **${highCount} tanesi yüksek** etkili olarak işaretlendi. Bu revizyonun projenin toplam bütçesine etkisi **${moneyText}** olarak hesaplanmıştır. ${discText}`;
}

function generateDiffs(matches) {
  const diffs = [];

  for (const match of matches) {
    const base = match.base;
    const comp = match.comp;
    const conf = match.confidence || 0;

    if (base && comp) {
      if (normalizeUnit(base.unit) !== normalizeUnit(comp.unit)) {
        diffs.push({
          positionCode: base.positionCode || comp.positionCode || '',
          itemName: base.itemName || comp.itemName || '',
          cat: 'unit_changed',
          severity: 'medium',
          v1Qty: base.quantity,
          v2Qty: comp.quantity,
          v1Unit: base.unit,
          v2Unit: comp.unit,
          deltaQty: 0,
          deltaPercent: 0,
          amountDelta: 0,
          activeUnitPrice: base?.unitPrice || comp?.unitPrice || 0,
          confidence: conf,
          v1Item: base,
          v2Item: comp,
        });
      } else if (base.quantity === comp.quantity) {
        const nameChanged = normalizeItemName(base.itemName) !== normalizeItemName(comp.itemName);
        diffs.push({
          positionCode: base.positionCode || '',
          itemName: nameChanged ? base.itemName : (base.itemName || ''),
          cat: nameChanged ? 'name_changed' : 'unchanged',
          severity: 'low',
          v1Qty: base.quantity,
          v2Qty: comp.quantity,
          v1Unit: base.unit,
          v2Unit: comp.unit,
          deltaQty: 0,
          deltaPercent: 0,
          amountDelta: 0,
          activeUnitPrice: base?.unitPrice || comp?.unitPrice || 0,
          confidence: conf,
          v1Item: base,
          v2Item: comp,
        });
      } else if (comp.quantity > base.quantity) {
        const delta = comp.quantity - base.quantity;
        const percent = base.quantity > 0 ? (delta / base.quantity) * 100 : 0;
        diffs.push({
          positionCode: base.positionCode || '',
          itemName: base.itemName || '',
          cat: 'increased',
          severity: calculateSeverity(delta, percent, base.quantity),
          v1Qty: base.quantity,
          v2Qty: comp.quantity,
          v1Unit: base.unit,
          v2Unit: comp.unit,
          deltaQty: delta,
          deltaPercent: percent,
          amountDelta: delta * (base.unitPrice || 0),
          activeUnitPrice: base.unitPrice || 0,
          confidence: conf,
          v1Item: base,
          v2Item: comp,
        });
      } else {
        const delta = comp.quantity - base.quantity;
        const percent = base.quantity > 0 ? (delta / base.quantity) * 100 : 0;
        diffs.push({
          positionCode: base.positionCode || '',
          itemName: base.itemName || '',
          cat: 'decreased',
          severity: calculateSeverity(delta, percent, base.quantity),
          v1Qty: base.quantity,
          v2Qty: comp.quantity,
          v1Unit: base.unit,
          v2Unit: comp.unit,
          deltaQty: delta,
          deltaPercent: percent,
          amountDelta: delta * (base.unitPrice || 0),
          activeUnitPrice: base.unitPrice || 0,
          confidence: conf,
          v1Item: base,
          v2Item: comp,
        });
      }
    } else if (base) {
      diffs.push({
        positionCode: base.positionCode || '',
        itemName: base.itemName || '',
        cat: 'removed',
        severity: 'high',
        v1Qty: base.quantity,
        v2Qty: 0,
        v1Unit: base.unit,
        v2Unit: '',
        deltaQty: -base.quantity,
        deltaPercent: -100,
        amountDelta: -(base.quantity * (base.unitPrice || 0)),
        activeUnitPrice: base.unitPrice || 0,
        confidence: 0,
        v1Item: base,
        v2Item: null,
      });
    } else {
      diffs.push({
        positionCode: comp.positionCode || '',
        itemName: comp.itemName || '',
        cat: 'added',
        severity: 'high',
        v1Qty: 0,
        v2Qty: comp.quantity,
        v1Unit: '',
        v2Unit: comp.unit,
        deltaQty: comp.quantity,
        deltaPercent: 100,
        amountDelta: comp.quantity * (comp.unitPrice || 0),
        activeUnitPrice: comp.unitPrice || 0,
        confidence: 0,
        v1Item: null,
        v2Item: comp,
      });
    }
  }

  return diffs;
}

function calculateSeverity(deltaQty, deltaPercent, baseQty) {
  if (Math.abs(deltaPercent) > 50) return 'critical';
  if (Math.abs(deltaPercent) > 20) return 'high';
  if (Math.abs(deltaQty) > 10) return 'medium';
  return 'low';
}

function showResults() {
  const t0 = Date.now();

  let items1, items2;
  if (state.isDemo) {
    items1 = state.parsedItems[0];
    items2 = state.parsedItems[1];
  } else {
    const p1 = parseSheet(1);
    const p2 = parseSheet(2);
    if (!p1 || !p2) {
      alert('Dosya verileri okunamadı. Lütfen dosyaları tekrar yükleyin.');
      return;
    }
    items1 = itemsFromParsed(p1, state.columnMappings[0]?.mapping);
    items2 = itemsFromParsed(p2, state.columnMappings[1]?.mapping);
  }

  state.parsedItems[0] = items1;
  state.parsedItems[1] = items2;

  const tMatch0 = Date.now();
  const threshold = 0.72;
  const matches = matchItems(items1, items2, threshold);
  const tMatch1 = Date.now();

  const tDiff0 = Date.now();
  state.allDiffs = generateDiffs(matches);
  const tDiff1 = Date.now();

  state.diffs = [...state.allDiffs];

  const totalAmount = state.diffs.reduce((s, d) => s + (d.amountDelta || 0), 0);
  const matchedCount = state.diffs.filter(d => d.v1Item && d.v2Item).length;
  const matchRate = items1.length > 0 ? ((matchedCount / items1.length) * 100).toFixed(1) : 0;
  const manualReviewCount = state.diffs.filter(d => ['added', 'removed', 'unit_changed', 'duplicate_suspect'].includes(d.cat)).length;

  const categoryCounts = {};
  const severityCounts = { critical: 0, high: 0, medium: 0, low: 0 };
  const disciplines = new Set();

  state.diffs.forEach(d => {
    categoryCounts[d.cat] = (categoryCounts[d.cat] || 0) + 1;
    severityCounts[d.severity] = (severityCounts[d.severity] || 0) + 1;
    if (d.v1Item?.discipline) disciplines.add(d.v1Item.discipline);
    if (d.v2Item?.discipline) disciplines.add(d.v2Item.discipline);
  });

  state.summary = {
    totalDiffs: state.diffs.length,
    totalAmount,
    matchRate,
    matchedCount,
    manualReviewCount,
    categoryCounts,
    severityCounts,
    disciplines: Array.from(disciplines),
  };

  metrics.track('comparison_run', {
    v1Items: items1.length,
    v2Items: items2.length,
    matchDuration: tMatch1 - tMatch0,
    diffDuration: tDiff1 - tDiff0,
    totalTime: Date.now() - t0
  });

  metrics.perfTiming = {
    parse1: tMatch0 - t0,
    match: tMatch1 - tMatch0,
    diff: tDiff1 - tDiff0,
    total: Date.now() - t0
  };

  renderResults();
  
  saveSession(); 
}

function renderResults() {
  const s = state.summary;
  const execSummary = generateExecutiveSummary(s, state.diffs);

  let kpiHtml = `
    <div class="executive-summary-panel" style="background: rgba(59, 130, 246, 0.05); padding: 16px; border-radius: 8px; border-left: 4px solid var(--primary); margin-bottom: 20px;">
      <h3 style="margin-bottom: 8px; font-weight: 700; color: var(--primary); font-size: 15px;">📋 Yönetici Özeti & Etki Analizi</h3>
      <p style="font-size: 13px; color: var(--text); line-height: 1.6;">${execSummary}</p>
    </div>
    <div class="kpi-row kpi-main">
      <div class="kpi-card kpi-large">
        <div class="kpi-value">${s.totalDiffs}</div>
        <div class="kpi-label">Toplam Kalem</div>
        <div class="kpi-sub">V1: ${state.parsedItems[0]?.length || 0} | V2: ${state.parsedItems[1]?.length || 0}</div>
      </div>
      <div class="kpi-card kpi-large">
        <div class="kpi-value ${s.totalAmount > 0 ? 'kpi-positive' : 'kpi-negative'}">${s.totalAmount > 0 ? '+' : ''}${s.totalAmount.toFixed(0)} ₺</div>
        <div class="kpi-label">Mali Etki</div>
        <div class="kpi-sub">Artış ve azalış</div>
      </div>
      <div class="kpi-card kpi-large">
        <div class="kpi-value">${s.matchRate}%</div>
        <div class="kpi-label">Eşleşme Oranı</div>
        <div class="kpi-sub">${s.matchedCount} / ${state.parsedItems[0]?.length || 0} kalem</div>
      </div>
      <div class="kpi-card kpi-large kpi-attention">
        <div class="kpi-value">${s.manualReviewCount}</div>
        <div class="kpi-label">Manuel İnceleme</div>
        <div class="kpi-sub">Dikkat gerekli</div>
      </div>
    </div>
    <div class="kpi-row kpi-categories">
  `;

  Object.keys(CATEGORIES).forEach(cat => {
    const count = s.categoryCounts[cat] || 0;
    kpiHtml += `<div class="kpi-card"><div class="kpi-value">${count}</div><div class="kpi-label">${CATEGORIES[cat].label}</div></div>`;
  });

  kpiHtml += '</div>';
  document.getElementById('kpiPanel').innerHTML = kpiHtml;

  let perfHtml = `<div class="perf-panel">
    <div class="perf-header" onclick="document.querySelector('.perf-body').classList.toggle('visible')">⚙️ Performans ▾</div>
    <div class="perf-body">
Dosya okuma:    ${(metrics.perfTiming?.parse1 || 0).toFixed(0)}ms<br>
Eşleştirme:     ${(metrics.perfTiming?.match || 0).toFixed(0)}ms<br>
Fark üretimi:   ${(metrics.perfTiming?.diff || 0).toFixed(0)}ms<br>
Toplam:         ${(metrics.perfTiming?.total || 0).toFixed(0)}ms
    </div>
  </div>`;
  document.getElementById('perfPanelSection').innerHTML = perfHtml;

  let discHtml = '<option value="">Tüm Disiplinler</option>';
  s.disciplines.forEach(d => {
    discHtml += `<option value="${d}">${d}</option>`;
  });
  document.getElementById('filterDiscipline').innerHTML = discHtml;

  applyFilters();

  document.getElementById('chartsRow').style.display = 'flex';
  
  if (window.costChartInstance) window.costChartInstance.destroy();
  if (window.typeChartInstance) window.typeChartInstance.destroy();

  const ctxCost = document.getElementById('costChart').getContext('2d');
  const ctxType = document.getElementById('typeChart').getContext('2d');

  const topCostDiffs = [...state.diffs]
    .sort((a,b) => Math.abs(b.amountDelta||0) - Math.abs(a.amountDelta||0))
    .slice(0, 10);

  window.costChartInstance = new Chart(ctxCost, {
    type: 'bar',
    data: {
      labels: topCostDiffs.map(d => d.positionCode),
      datasets: [{
        label: 'Mali Etki (₺)',
        data: topCostDiffs.map(d => d.amountDelta || 0),
        backgroundColor: topCostDiffs.map(d => (d.amountDelta||0) > 0 ? '#ef4444' : '#10b981')
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: { display: true, text: 'En Yüksek Etkiye Sahip 10 Kalem', color: '#94a3b8' }
      },
      scales: {
        y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(148, 163, 184, 0.1)' } },
        x: { ticks: { color: '#94a3b8' }, grid: { display: false } }
      }
    }
  });

  const categoryLabels = Object.keys(CATEGORIES).map(k => CATEGORIES[k].label);
  const categoryData = Object.keys(CATEGORIES).map(k => s.categoryCounts[k] || 0);

  window.typeChartInstance = new Chart(ctxType, {
    type: 'doughnut',
    data: {
      labels: categoryLabels,
      datasets: [{
        data: categoryData,
        backgroundColor: ['#ef4444','#10b981','#f59e0b','#3b82f6','#8b5cf6','#64748b','#ec4899','#14b8a6'],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { color: '#94a3b8' } },
        title: { display: true, text: 'Değişim Kategori Dağılımı', color: '#94a3b8' }
      }
    }
  });
}

function applyFilters() {
  const catFilter = document.getElementById('filterCategory').value;
  const sevFilter = document.getElementById('filterSeverity').value;
  const discFilter = document.getElementById('filterDiscipline').value;
  const searchVal = document.getElementById('filterSearch').value.toLowerCase();
  const sortBy = document.getElementById('filterSort').value;

  let filtered = state.allDiffs.filter(d => {
    if (catFilter && d.cat !== catFilter) return false;
    if (sevFilter && d.severity !== sevFilter) return false;
    if (discFilter && d.v1Item?.discipline !== discFilter && d.v2Item?.discipline !== discFilter) return false;
    if (searchVal && !d.itemName.toLowerCase().includes(searchVal) && !d.positionCode.toLowerCase().includes(searchVal)) return false;
    return true;
  });

  filtered.sort((a, b) => {
    if (sortBy === 'severity') {
      const sevOrder = { critical: 0, high: 1, medium: 2, low: 3 };
      return sevOrder[a.severity] - sevOrder[b.severity];
    } else if (sortBy === 'amount') {
      return Math.abs(b.amountDelta || 0) - Math.abs(a.amountDelta || 0);
    } else if (sortBy === 'delta') {
      return Math.abs(b.deltaPercent) - Math.abs(a.deltaPercent);
    } else {
      return a.itemName.localeCompare(b.itemName);
    }
  });

  state.diffs = filtered;
  renderTable();
  document.getElementById('filterCount').textContent = `${filtered.length} / ${state.allDiffs.length} fark`;
}

function renderTable() {
  let html = '';
  state.diffs.forEach((d, i) => {
    const deltaCls = d.deltaQty > 0 ? 'delta-pos' : (d.deltaQty < 0 ? 'delta-neg' : 'delta-neutral');
    const amountCls = (d.amountDelta || 0) > 0 ? 'delta-pos' : ((d.amountDelta || 0) < 0 ? 'delta-neg' : 'delta-neutral');


    html += `<tr onclick="toggleDetail(${i})">
      <td class="poz">${d.positionCode}</td>
      <td class="item-name">
        ${d.itemName}
        ${d.confidence > 0 && d.confidence < 0.95 ? `<br><span style="font-size: 11px; font-weight: 600; color: #D97706; background: rgba(217, 119, 6, 0.1); padding: 2px 6px; border-radius: 4px; margin-top: 4px; display: inline-block;">⚠️ Şüpheli Eşleşme (Güven: %${Math.round(d.confidence*100)})</span>` : ''}
      </td>
      <td><span class="badge badge-${d.cat}">${CATEGORIES[d.cat].label}</span></td>
      <td><span class="badge badge-${d.severity}">${d.severity.charAt(0).toUpperCase() + d.severity.slice(1)}</span></td>
      <td>${d.v1Qty.toFixed(2)} ${d.v1Unit}</td>
      <td>${d.v2Qty.toFixed(2)} ${d.v2Unit}</td>
      <td class="${deltaCls}">${d.deltaQty > 0 ? '+' : ''}${d.deltaQty.toFixed(2)} (${d.deltaPercent.toFixed(0)}%)</td>
      <td>
        <input type="number" 
               class="inline-editor" 
               value="${d.activeUnitPrice || 0}" 
               step="0.01"
               onclick="event.stopPropagation()"
               onchange="updateUnitPrice(${i}, this.value)" />
      </td>
      <td class="${amountCls}">${(d.amountDelta || 0) > 0 ? '+' : ''}${(d.amountDelta || 0).toFixed(2)} ₺</td>
    </tr>`;

    if (state.expandedRow === i) {
      const v1Disc = d.v1Item?.discipline || '-';
      const v2Disc = d.v2Item?.discipline || '-';
      const v1Zone = d.v1Item?.zone || '-';
      const v2Zone = d.v2Item?.zone || '-';
      html += `<tr class="detail-row">
        <td colspan="8">
          <div class="detail-grid">
            <div>
              <div class="dlabel">V1 Disiplin</div>
              <div class="dval">${v1Disc}</div>
            </div>
            <div>
              <div class="dlabel">V2 Disiplin</div>
              <div class="dval">${v2Disc}</div>
            </div>
            <div>
              <div class="dlabel">Eşleştirme Güvenliği</div>
              <div class="dval">${(d.v1Item && d.v2Item ? (d.confidence ? '%'+Math.round(d.confidence*100) : '%100') : (d.v1Item || d.v2Item ? 'Eşleşmemiş' : '-'))}</div>
            </div>
            <div>
              <div class="dlabel">V1 Mahal</div>
              <div class="dval">${v1Zone}</div>
            </div>
            <div>
              <div class="dlabel">V2 Mahal</div>
              <div class="dval">${v2Zone}</div>
            </div>
            <div></div>
            <div class="detail-explanation">
              ${getExplanation(d)}
            </div>
            <div class="feedback-actions">
              <button class="fb-btn fb-correct" onclick="submitItemFeedback(${i}, 'correct')">✅ Doğru</button>
              <button class="fb-btn fb-wrong" onclick="submitItemFeedback(${i}, 'wrong')">❌ Yanlış</button>
              <button class="fb-btn fb-note" onclick="submitItemFeedback(${i}, 'note')">💬 Not</button>
            </div>
          </div>
        </td>
      </tr>`;
    }
  });
  document.getElementById('diffBody').innerHTML = html;
}

function toggleDetail(idx) {
  state.expandedRow = state.expandedRow === idx ? null : idx;
  renderTable();
}

function submitItemFeedback(idx, type) {
  const d = state.diffs[idx];
  const note = type === 'note' ? prompt('Not ekleyin:') : '';
  state.feedback.push({ item: d.itemName, type, note, timestamp: new Date() });
  showToast(type === 'correct' ? '✅ Geri bildirim kaydedildi' : (type === 'wrong' ? '❌ Hata kaydedildi' : '💬 Not kaydedildi'));
  metrics.track('feedback_submit', { type, item: d.itemName });
}

function clearFilters() {
  document.getElementById('filterCategory').value = '';
  document.getElementById('filterSeverity').value = '';
  document.getElementById('filterDiscipline').value = '';
  document.getElementById('filterSort').value = 'severity';
  document.getElementById('filterSearch').value = '';
  applyFilters();
}

function exportExcel() {
  metrics.track('export_excel');
  const wb = XLSX.utils.book_new();

  const summaryData = [
    ['MetrajOS - Karşılaştırma Özeti'],
    [],
    ['Metrik', 'Değer'],
    ['Toplam Fark Sayısı', state.summary.totalDiffs],
    ['Mali Etki (₺)', state.summary.totalAmount],
    ['Eşleşme Oranı (%)', state.summary.matchRate],
    ['V1 Toplam Kalem', state.parsedItems[0].length],
    ['V2 Toplam Kalem', state.parsedItems[1].length],
    ['Manuel İnceleme Gerekli', state.summary.manualReviewCount],
    [],
    ['Kategori Dağılımı'],
    ['Kategori', 'Sayı'],
  ];

  Object.keys(CATEGORIES).forEach(cat => {
    summaryData.push([CATEGORIES[cat].label, state.summary.categoryCounts[cat] || 0]);
  });

  const ws1 = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, ws1, 'Özet');

  const detailData = [
    ['Poz No', 'İmalat Adı', 'Kategori', 'Şiddet', 'V1 Miktar', 'V1 Birim', 'V2 Miktar', 'V2 Birim', 'Fark Miktarı', 'Fark %', 'Mali Fark (₺)', 'Açıklama'],
  ];

  state.allDiffs.forEach(d => {
    detailData.push([
      d.positionCode,
      d.itemName,
      CATEGORIES[d.cat].label,
      d.severity,
      d.v1Qty,
      d.v1Unit,
      d.v2Qty,
      d.v2Unit,
      d.deltaQty,
      d.deltaPercent.toFixed(2),
      (d.amountDelta || 0).toFixed(2),
      getExplanation(d),
    ]);
  });

  const ws2 = XLSX.utils.aoa_to_sheet(detailData);
  XLSX.utils.book_append_sheet(wb, ws2, 'Fark Detayları');

  XLSX.writeFile(wb, 'MetrajOS_Karşılaştırma.xlsx');
}

function exportFeedback() {
  metrics.track('export_feedback');
  const data = {
    sessionId: metrics.sessionId,
    timestamp: new Date(),
    feedback: state.feedback,
    summary: state.summary,
    totalEvents: metrics.events.length
  };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'MetrajOS_Geri_Bildirim.json';
  a.click();
}

function openFeedbackModal() {
  document.getElementById('feedbackModal').classList.add('visible');
}

function closeFeedbackModal() {
  document.getElementById('feedbackModal').classList.remove('visible');
}

function submitFeedback() {
  const type = document.getElementById('feedbackType').value;
  const text = document.getElementById('feedbackText').value;

  if (!type || !text) {
    alert('Lütfen tür ve metin girin');
    return;
  }

  state.feedback.push({ type, text, timestamp: new Date() });
  metrics.track('feedback_submit', { type });

  document.getElementById('feedbackType').value = '';
  document.getElementById('feedbackText').value = '';
  closeFeedbackModal();
  showToast('Geri bildiriminiz kaydedildi');
}

function showToast(msg) {
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 3000);
}

function goToStep(n) {
  for (let i = 1; i <= 3; i++) {
    const sec = document.getElementById('section' + i);
    const step = document.getElementById('step' + i + 'Indicator');
    if (sec) sec.classList.remove('active');
    if (step) step.classList.remove('active', 'completed');
  }

  const activeSection = document.getElementById('section' + n);
  const activeStep = document.getElementById('step' + n + 'Indicator');
  if (activeSection) activeSection.classList.add('active');
  if (activeStep) activeStep.classList.add('active');

  for (let i = 1; i < n; i++) {
    const prevStep = document.getElementById('step' + i + 'Indicator');
    if (prevStep) prevStep.classList.add('completed');
  }

  if (n === 2) {
    showMappingUI();
  } else if (n === 3) {
    showResults();
  }
}

function runTests() {
  const results = [];

  function assert(condition, message) {
    if (condition) {
      results.push(`<span class="test-pass">✓ ${message}</span>`);
    } else {
      results.push(`<span class="test-fail">✗ ${message}</span>`);
    }
  }

  assert(normalizeUnit('m2') === 'm²', 'Unit: m2 → m²');
  assert(normalizeUnit('M²') === 'm²', 'Unit: M² → m²');
  assert(normalizeUnit('metrekare') === 'm²', 'Unit: metrekare → m²');
  assert(parseNumber('1.250,50') === 1250.50, 'Number: Turkish 1.250,50 → 1250.50');
  assert(parseNumber('1250.50') === 1250.50, 'Number: 1250.50 → 1250.50');
  assert(isNaN(parseNumber('')), 'Number: empty → NaN');

  const testHeaders = ['proje_kodu', 'poz_no', 'imalat_adi', 'birim', 'miktar', 'birim_fiyat_tl'];
  assert(findColumn(testHeaders, COL_PATTERNS.positionCode) === 1, 'Column: poz_no at index 1');
  assert(findColumn(testHeaders, COL_PATTERNS.itemName) === 2, 'Column: imalat_adi at index 2');

  assert(calculateSimilarity('Gazbeton Duvar 19 cm', '19 cm Gazbeton Duvar') > 0.7, 'Similarity: word reorder > 0.7');
  assert(calculateSimilarity('C30 Beton', 'Lineer LED') < 0.3, 'Similarity: different < 0.3');

  const baseItems = [
    { positionCode: '15.150.100', itemName: 'C30 Beton', unit: 'm3', quantity: 100, zone: 'Temel', discipline: '', unitPrice: 2900, rowIndex: 2 },
  ];
  const compItems = [
    { positionCode: '15.150.100', itemName: 'C30 Beton', unit: 'm3', quantity: 120, zone: 'Temel', discipline: '', unitPrice: 2900, rowIndex: 2 },
  ];
  const matches = matchItems(baseItems, compItems, 0.72);
  const diffs = generateDiffs(matches);
  assert(diffs.some(d => d.cat === 'increased'), 'Diff: has increased');

  assert(DEMO_SCENARIOS.revision.v1.length === 12, 'Demo: Revision scenario has 12 V1 items');
  assert(DEMO_SCENARIOS.addremove.v2.length === 12, 'Demo: Addremove scenario has 12 V2 items');
  assert(DEMO_SCENARIOS.unitname.v1.length === 8, 'Demo: Unitname scenario has 8 V1 items');

  return results;
}

function openTestModal() {
  const results = runTests();
  const html = results.join('<br>');
  document.getElementById('testResults').innerHTML = html;
  document.getElementById('testModal').classList.add('visible');
}

function closeTestModal() {
  document.getElementById('testModal').classList.remove('visible');
}

document.addEventListener('keydown', e => {
  if (e.ctrlKey && e.key === 't') {
    e.preventDefault();
    openTestModal();
  }
  if (e.ctrlKey && e.key === 'e') {
    e.preventDefault();
    exportExcel();
  }
  if (e.ctrlKey && e.key === 'd') {
    e.preventDefault();
    openFeedbackModal();
  }
});

document.addEventListener('click', e => {
  if (e.target.id === 'feedbackModal') {
    closeFeedbackModal();
  }
  if (e.target.id === 'testModal') {
    closeTestModal();
  }
});

// Expose functions to global scope for HTML inline handlers
window.normalizeUnit=normalizeUnit;
window.normalizeItemName=normalizeItemName;
window.parseNumber=parseNumber;
window.findColumn=findColumn;
window.calculateSimilarity=calculateSimilarity;
window.detectDuplicateRows=detectDuplicateRows;
window.loadDemo=loadDemo;
window.handleFile=handleFile;
window.selectSheet=selectSheet;
window.checkUploadComplete=checkUploadComplete;
window.parseSheet=parseSheet;
window.showMappingUI=showMappingUI;
window.showPreview=showPreview;
window.itemsFromParsed=itemsFromParsed;
window.matchItems=matchItems;
window.getExplanation=getExplanation;
window.generateDiffs=generateDiffs;
window.calculateSeverity=calculateSeverity;
window.showResults=showResults;
window.renderResults=renderResults;
window.applyFilters=applyFilters;
window.renderTable=renderTable;
window.toggleDetail=toggleDetail;
window.submitItemFeedback=submitItemFeedback;
window.clearFilters=clearFilters;
window.exportExcel=exportExcel;
window.exportFeedback=exportFeedback;
window.openFeedbackModal=openFeedbackModal;
window.closeFeedbackModal=closeFeedbackModal;
window.submitFeedback=submitFeedback;
window.showToast=showToast;
window.goToStep=goToStep;
window.runTests=runTests;
window.openTestModal=openTestModal;
window.closeTestModal=closeTestModal;
window.toggleDarkMode=toggleDarkMode;
window.exportPDF=exportPDF;
window.updateUnitPrice=updateUnitPrice;
window.loadPastSession=loadPastSession;
window.clearPastSession=clearPastSession;

function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
  const isDark = document.body.classList.contains('dark-mode');
  localStorage.setItem('metraj_dark_mode', isDark);
}

function exportPDF() {
  const element = document.getElementById('section3');
  const opt = {
    margin:       0.5,
    filename:     'Metraj_Karsilastirma_Raporu.pdf',
    image:        { type: 'jpeg', quality: 0.98 },
    html2canvas:  { scale: 2, useCORS: true },
    jsPDF:        { unit: 'in', format: 'a4', orientation: 'landscape' }
  };
  
  const originalActionBars = element.querySelectorAll('.action-bar');
  originalActionBars.forEach(bar => bar.style.display = 'none');
  
  html2pdf().set(opt).from(element).save().then(() => {
    originalActionBars.forEach(bar => bar.style.display = '');
  });
}

function updateUnitPrice(idx, newVal) {
  const val = parseFloat(newVal) || 0;
  const d = state.diffs[idx];
  d.activeUnitPrice = val;
  d.amountDelta = d.deltaQty * val;
  
  // Update state.allDiffs item
  const allIdx = state.allDiffs.indexOf(state.diffs[idx]);
  if (allIdx >= 0) state.allDiffs[allIdx] = d;

  // re-calculate summary
  const totalAmount = state.diffs.reduce((s, _d) => s + (_d.amountDelta || 0), 0);
  state.summary.totalAmount = totalAmount;
  
  // Re-render
  renderResults();
}

function saveSession() {
  const sessionData = {
    v1Length: state.parsedItems[0]?.length || 0,
    v2Length: state.parsedItems[1]?.length || 0,
    summary: state.summary,
    allDiffs: state.allDiffs.map(d => ({
      positionCode: d.positionCode,
      itemName: d.itemName,
      cat: d.cat,
      severity: d.severity,
      v1Qty: d.v1Qty,
      v2Qty: d.v2Qty,
      v1Unit: d.v1Unit,
      v2Unit: d.v2Unit,
      deltaQty: d.deltaQty,
      deltaPercent: d.deltaPercent,
      amountDelta: d.amountDelta,
      activeUnitPrice: d.activeUnitPrice,
      v1Item: d.v1Item ? { rowIndex: d.v1Item.rowIndex, discipline: d.v1Item.discipline } : null,
      v2Item: d.v2Item ? { rowIndex: d.v2Item.rowIndex, discipline: d.v2Item.discipline } : null
    })),
    timestamp: Date.now()
  };
  localStorage.setItem('metraj_session', JSON.stringify(sessionData));
}

function loadPastSession() {
  const data = localStorage.getItem('metraj_session');
  if(!data) return;
  try {
    const sessionData = JSON.parse(data);
    state.allDiffs = sessionData.allDiffs;
    state.diffs = [...state.allDiffs];
    state.summary = sessionData.summary;
    state.parsedItems = [ new Array(sessionData.v1Length), new Array(sessionData.v2Length) ];
    goToStep(3);
    document.getElementById('sessionLoaderContainer').innerHTML = '';
  } catch(err) {
    console.error("Session load error", err);
  }
}

function clearPastSession() {
  localStorage.removeItem('metraj_session');
  document.getElementById('sessionLoaderContainer').innerHTML = '';
}

function initUI() {
  if (localStorage.getItem('metraj_dark_mode') === 'true') {
    document.body.classList.add('dark-mode');
  }
  const data = localStorage.getItem('metraj_session');
  if(data) {
    try {
      const sessionData = JSON.parse(data);
      const d = new Date(sessionData.timestamp).toLocaleString('tr-TR');
      document.getElementById('sessionLoaderContainer').innerHTML = `
        <div class="session-loader">
          <div class="session-info">
            <strong>Kayıtlı oturumunuz bulundu</strong>
            <span>${d} tarihindeki karşılaştırma</span>
          </div>
          <div style="display:flex; gap:8px;">
            <button class="btn btn-outline btn-small" onclick="clearPastSession()">İptal Et</button>
            <button class="btn btn-primary btn-small" onclick="loadPastSession()">Geri Yükle</button>
          </div>
        </div>
      `;
    } catch(err) {}
  }
}

document.addEventListener('DOMContentLoaded', initUI);
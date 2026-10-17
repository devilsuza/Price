import os
import uuid
import sqlite3
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string
import pandas as pd

app = Flask(__name__)
UPLOAD_FOLDER = "temp_uploads"
DB_NAME = "store_prices.db"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# --- DATABASE SETUP ---
def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS products (
                barcode TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                price TEXT NOT NULL,
                category TEXT,
                source_tab TEXT,
                updated_at TEXT
            )
        """)
        conn.commit()


init_db()


def clean_barcode(val):
    """Cleans Excel float formatting (e.g., converts '89012345.0' to '89012345')."""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


# --- HTML TEMPLATE (LOGIN + SCANNER + MANAGER PORTAL) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Store Price Checker & Manager Portal</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://unpkg.com/html5-qrcode"></script>
</head>
<body class="bg-slate-100 min-h-screen font-sans text-slate-800">

    <!-- ================= LOGIN MODAL OVERLAY ================= -->
    <div id="login-overlay" class="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white w-full max-w-sm rounded-2xl shadow-2xl p-6 border border-slate-200 text-center">
            <div class="w-14 h-14 bg-indigo-100 text-indigo-700 rounded-2xl flex items-center justify-center mx-auto mb-3 text-2xl font-bold shadow-sm">
                🔐
            </div>
            <h2 class="text-xl font-extrabold text-slate-800">Store Portal Login</h2>
            <p class="text-xs text-slate-500 mt-1 mb-5">Enter PIN for Employee or Manager access</p>

            <input type="password" id="login-pin" autofocus
                placeholder="Enter Password (123 or 321)"
                class="w-full px-4 py-3 text-center text-xl tracking-widest font-bold border-2 border-indigo-300 rounded-xl focus:outline-none focus:border-indigo-600 mb-3">
            
            <p id="login-error" class="text-xs font-bold text-red-600 mb-3 hidden">❌ Wrong Password! Try again.</p>

            <button onclick="handleLogin()" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 rounded-xl shadow-md transition">
                Login Now
            </button>
        </div>
    </div>

    <!-- ================= TOP NAVIGATION ================= -->
    <nav class="bg-indigo-700 text-white shadow-lg">
        <div class="max-w-5xl mx-auto px-4 py-3 flex flex-wrap justify-between items-center gap-2">
            <div class="flex items-center space-x-2">
                <span class="text-xl font-bold tracking-wide">🏪 Store Price Portal</span>
                <span id="role-badge" class="text-xs bg-indigo-900 text-indigo-200 px-2.5 py-0.5 rounded-full font-semibold hidden"></span>
            </div>
            <div class="flex items-center space-x-2">
                <button onclick="switchMode('scanner')" id="btn-scanner" class="px-3.5 py-1.5 rounded-lg text-sm font-medium bg-white text-indigo-700 shadow">
                    🔍 Scanner
                </button>
                <button onclick="switchMode('manager')" id="btn-manager" class="px-3.5 py-1.5 rounded-lg text-sm font-medium text-indigo-100 hover:bg-indigo-600">
                    ⚙️ Manager Upload
                </button>
                <button onclick="logoutPortal()" class="px-3 py-1.5 rounded-lg text-sm font-medium bg-red-500/90 hover:bg-red-600 text-white shadow">
                    🔒 Logout
                </button>
            </div>
        </div>
    </nav>

    <div class="max-w-4xl mx-auto p-4 mt-4">

        <!-- ================= EMPLOYEE SCANNER VIEW ================= -->
        <div id="view-scanner" class="space-y-6">
            <div class="bg-white p-6 rounded-2xl shadow-md border border-slate-200">
                <h2 class="text-lg font-bold text-slate-700 mb-3">Scan Article Barcode or Search Name</h2>
                
                <!-- Search Input with Inline Small Scanner Icon inside the Box -->
                <div class="flex flex-col sm:flex-row gap-3">
                    <div class="relative flex-1">
                        <input type="text" id="scan-input"
                            placeholder="Scan barcode here or type article code / name..."
                            class="w-full pl-4 pr-14 py-3 text-lg border-2 border-indigo-300 rounded-xl focus:outline-none focus:border-indigo-600">
                        
                        <!-- Small Inline Barcode/QR Scanner Button Inside Input -->
                        <button type="button" onclick="toggleCamera()" id="inline-scan-btn"
                            title="Open QR / Barcode Camera Scanner"
                            class="absolute right-2 top-1/2 -translate-y-1/2 p-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-lg border border-indigo-200 transition flex items-center justify-center shadow-sm">
                            <!-- Barcode/QR SVG Icon -->
                            <svg xmlns="http://www.w3.org/2000/svg" class="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M3 7V5a2 2 0 012-2h2m10 0h2a2 2 0 012 2v2m0 10v2a2 2 0 01-2 2h-2M7 21H5a2 2 0 01-2-2v-2M7 8h1v8H7V8zm3 0h2v8h-2V8zm4 0h1v8h-1V8zm3 0h1v8h-1V8z" />
                            </svg>
                        </button>
                    </div>

                    <button onclick="searchArticle()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold px-6 py-3 rounded-xl shadow">
                        Check Price
                    </button>
                </div>

                <!-- Mobile/Web Camera Container (Opens when small icon is clicked) -->
                <div id="reader-container" class="hidden mt-4 bg-slate-900 p-3 rounded-2xl max-w-sm mx-auto text-center">
                    <div class="flex justify-between items-center text-white text-xs font-semibold mb-2 px-1">
                        <span>📷 Point camera at Barcode or QR Code</span>
                        <button onclick="stopCamera()" class="bg-red-500 hover:bg-red-600 text-white px-2.5 py-1 rounded-md">✕ Close</button>
                    </div>
                    <div id="reader" class="w-full overflow-hidden rounded-xl border-2 border-emerald-400 bg-black"></div>
                </div>
            </div>

            <!-- Price Display Result Card -->
            <div id="price-result" class="hidden"></div>
        </div>

        <!-- ================= MANAGER UPLOAD VIEW ================= -->
        <div id="view-manager" class="hidden space-y-6">
            
            <!-- Step 1: Upload Excel -->
            <div class="bg-white p-6 rounded-2xl shadow-md border border-slate-200">
                <div class="flex justify-between items-center mb-4">
                    <div>
                        <h2 class="text-xl font-bold text-slate-800">1. Upload Daily Price Excel</h2>
                        <p class="text-sm text-slate-500">Upload .xlsx, .xls, or .csv file to map tabs and columns</p>
                    </div>
                    <span class="bg-indigo-50 text-indigo-700 text-xs font-bold px-3 py-1.5 rounded-full border border-indigo-200" id="total-count-badge">
                        Total Articles: {{ total_products }}
                    </span>
                </div>

                <div class="flex flex-col sm:flex-row items-center gap-4">
                    <input type="file" id="excel-file" accept=".xlsx,.xls,.csv"
                        class="block w-full text-sm text-slate-600 file:mr-4 file:py-2.5 file:px-4 file:rounded-xl file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100 border border-slate-200 rounded-xl p-1.5">
                    <button onclick="uploadForPreview()" id="btn-upload-preview"
                        class="w-full sm:w-auto whitespace-nowrap bg-indigo-600 hover:bg-indigo-700 text-white font-semibold px-6 py-2.5 rounded-xl shadow">
                        Read Tabs & Columns
                    </button>
                </div>
                <div id="upload-status" class="mt-3 text-sm font-medium hidden"></div>
            </div>

            <!-- Step 2: Dynamic Tab & Column Mapping -->
            <div id="mapping-section" class="hidden bg-white p-6 rounded-2xl shadow-md border border-slate-200 space-y-5">
                <div>
                    <h2 class="text-xl font-bold text-slate-800">2. Select Which Column for Which Tab</h2>
                    <p class="text-sm text-slate-500">Check the tabs you want to import and choose the matching column for each field.</p>
                </div>

                <div id="tabs-container" class="space-y-6"></div>

                <div class="pt-4 border-t border-slate-200 flex justify-end">
                    <button onclick="confirmAndImport()" id="btn-confirm-import"
                        class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-8 py-3 rounded-xl shadow-lg text-base">
                        ✅ Save & Update Daily Prices
                    </button>
                </div>
            </div>

        </div>

    </div>

    <script>
        // --- PASSWORDS ---
        const EMPLOYEE_PASS = "123";
        const MANAGER_PASS = "321";
        let currentRole = null; // 'employee' or 'manager'

        let currentTempFile = null;
        let html5QrCode = null;
        let cameraActive = false;

        // --- SUPERMARKET BEEP SOUND (Web Audio API - No external MP3 needed) ---
        function playBeep() {
            try {
                const AudioCtx = window.AudioContext || window.webkitAudioContext;
                const ctx = new AudioCtx();
                const osc = ctx.createOscillator();
                const gain = ctx.createGain();

                osc.type = "sine";
                osc.frequency.setValueAtTime(1350, ctx.currentTime); // High-pitched store scanner beep
                gain.gain.setValueAtTime(0.25, ctx.currentTime);

                osc.connect(gain);
                gain.connect(ctx.destination);

                osc.start();
                osc.stop(ctx.currentTime + 0.14); // 140ms crisp beep
            } catch (e) {
                console.log("Audio beep error:", e);
            }
        }

        // --- LOGIN & ROLE MANAGEMENT ---
        document.getElementById('login-pin').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') handleLogin();
        });

        function handleLogin() {
            const pinInput = document.getElementById('login-pin');
            const errEl = document.getElementById('login-error');
            const pin = pinInput.value.trim();

            if (pin === EMPLOYEE_PASS) {
                currentRole = "employee";
                errEl.classList.add('hidden');
                document.getElementById('login-overlay').classList.add('hidden');
                updateRoleBadge("👤 Employee");
                switchMode('scanner');
            } else if (pin === MANAGER_PASS) {
                currentRole = "manager";
                errEl.classList.add('hidden');
                document.getElementById('login-overlay').classList.add('hidden');
                updateRoleBadge("👔 Manager");
                switchMode('manager');
            } else {
                errEl.classList.remove('hidden');
                pinInput.value = "";
                pinInput.focus();
            }
        }

        function updateRoleBadge(text) {
            const badge = document.getElementById('role-badge');
            badge.innerText = text;
            badge.classList.remove('hidden');
        }

        function logoutPortal() {
            stopCamera();
            currentRole = null;
            document.getElementById('login-pin').value = "";
            document.getElementById('login-error').classList.add('hidden');
            document.getElementById('login-overlay').classList.remove('hidden');
            document.getElementById('login-pin').focus();
        }

        // Switch between Employee Scanner and Manager View
        function switchMode(mode) {
            if (mode === 'manager' && currentRole !== 'manager') {
                const askPin = prompt("🔒 Enter Manager Password to access Upload section:");
                if (askPin === MANAGER_PASS) {
                    currentRole = 'manager';
                    updateRoleBadge("👔 Manager");
                } else {
                    if (askPin !== null) alert("❌ Invalid Manager Password!");
                    return;
                }
            }

            const scannerView = document.getElementById('view-scanner');
            const managerView = document.getElementById('view-manager');
            const btnScanner = document.getElementById('btn-scanner');
            const btnManager = document.getElementById('btn-manager');

            if (mode === 'scanner') {
                scannerView.classList.remove('hidden');
                managerView.classList.add('hidden');
                btnScanner.className = "px-3.5 py-1.5 rounded-lg text-sm font-medium bg-white text-indigo-700 shadow";
                btnManager.className = "px-3.5 py-1.5 rounded-lg text-sm font-medium text-indigo-100 hover:bg-indigo-600";
                document.getElementById('scan-input').focus();
            } else {
                scannerView.classList.add('hidden');
                managerView.classList.remove('hidden');
                btnManager.className = "px-3.5 py-1.5 rounded-lg text-sm font-medium bg-white text-indigo-700 shadow";
                btnScanner.className = "px-3.5 py-1.5 rounded-lg text-sm font-medium text-indigo-100 hover:bg-indigo-600";
                stopCamera();
            }
        }

        // Listen for Enter key from USB Barcode Scanner or keyboard
        document.getElementById('scan-input').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                playBeep();
                searchArticle();
            }
        });

        // Search Article by Barcode or Name in Database
        async function searchArticle(overrideQuery = null) {
            const inputEl = document.getElementById('scan-input');
            const query = (overrideQuery !== null ? overrideQuery : inputEl.value).trim();
            if (!query) return;

            const resultBox = document.getElementById('price-result');
            resultBox.classList.remove('hidden');
            resultBox.innerHTML = `<div class="bg-white p-6 rounded-2xl shadow text-center text-slate-500">Searching article in database...</div>`;

            try {
                const res = await fetch(`/api/search?q=${encodeURIComponent(query)}`);
                const data = await res.json();

                if (!data.found || data.items.length === 0) {
                    resultBox.innerHTML = `
                        <div class="bg-red-50 border-2 border-red-300 p-6 rounded-2xl text-center">
                            <p class="text-2xl font-bold text-red-600">❌ Article Not Found</p>
                            <p class="text-slate-600 mt-1">No price record found for: <strong>${query}</strong></p>
                        </div>`;
                } else {
                    resultBox.innerHTML = data.items.map(item => `
                        <div class="bg-white border-2 border-emerald-500 p-6 rounded-2xl shadow-lg mb-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                            <div>
                                <span class="text-xs font-bold uppercase tracking-wider bg-slate-100 text-slate-600 px-2.5 py-1 rounded-md">
                                    Barcode: ${item.barcode}
                                </span>
                                ${item.category ? `<span class="ml-2 text-xs font-bold uppercase bg-indigo-50 text-indigo-700 px-2.5 py-1 rounded-md">${item.category}</span>` : ''}
                                <h3 class="text-2xl font-extrabold text-slate-800 mt-2">${item.name}</h3>
                                <p class="text-xs text-slate-400 mt-1">Tab: ${item.source_tab} • Updated: ${item.updated_at}</p>
                            </div>
                            <div class="bg-emerald-50 border border-emerald-200 px-6 py-4 rounded-2xl text-right w-full sm:w-auto">
                                <span class="text-xs uppercase font-bold text-emerald-700 block">Selling Price</span>
                                <span class="text-4xl font-black text-emerald-600">${item.price}</span>
                            </div>
                        </div>
                    `).join('');
                }
                inputEl.select();
            } catch (err) {
                resultBox.innerHTML = `<div class="bg-red-100 p-4 rounded-xl text-red-700">Error fetching price from database.</div>`;
            }
        }

        // --- INLINE CAMERA BARCODE & QR SCANNER ---
        function toggleCamera() {
            if (cameraActive) {
                stopCamera();
            } else {
                const container = document.getElementById('reader-container');
                const inlineBtn = document.getElementById('inline-scan-btn');
                container.classList.remove('hidden');
                inlineBtn.classList.replace('bg-indigo-50', 'bg-emerald-100');
                inlineBtn.classList.replace('text-indigo-700', 'text-emerald-700');

                html5QrCode = new Html5Qrcode("reader");
                cameraActive = true;

                html5QrCode.start(
                    { facingMode: "environment" },
                    { fps: 12, qrbox: { width: 250, height: 150 } },
                    (decodedText) => {
                        // 1. Play Beep Sound immediately
                        playBeep();
                        // 2. Put barcode into search input
                        document.getElementById('scan-input').value = decodedText;
                        // 3. Stop camera automatically
                        stopCamera();
                        // 4. Fetch and display database details below
                        searchArticle(decodedText);
                    },
                    (errorMessage) => {}
                ).catch(err => {
                    alert("Camera open nahi ho paya. Please allow camera permission in browser.");
                    stopCamera();
                });
            }
        }

        function stopCamera() {
            const inlineBtn = document.getElementById('inline-scan-btn');
            inlineBtn.classList.replace('bg-emerald-100', 'bg-indigo-50');
            inlineBtn.classList.replace('text-emerald-700', 'text-indigo-700');

            if (html5QrCode && cameraActive) {
                html5QrCode.stop().then(() => {
                    html5QrCode.clear();
                }).catch(() => {});
            }
            cameraActive = false;
            document.getElementById('reader-container').classList.add('hidden');
        }

        // Step 1: Upload Excel and Get Tabs + Columns
        async function uploadForPreview() {
            const fileInput = document.getElementById('excel-file');
            const statusEl = document.getElementById('upload-status');
            if (!fileInput.files.length) {
                alert("Please select an Excel file first.");
                return;
            }

            const formData = new FormData();
            formData.append("file", fileInput.files[0]);

            statusEl.classList.remove('hidden');
            statusEl.className = "mt-3 text-sm font-medium text-indigo-600";
            statusEl.innerText = "⏳ Reading Excel tabs and columns...";

            try {
                const res = await fetch("/api/preview-excel", {
                    method: "POST",
                    body: formData
                });
                const data = await res.json();

                if (!data.success) {
                    statusEl.className = "mt-3 text-sm font-medium text-red-600";
                    statusEl.innerText = "❌ " + data.error;
                    return;
                }

                currentTempFile = data.temp_filename;
                statusEl.className = "mt-3 text-sm font-medium text-emerald-600";
                statusEl.innerText = `✅ Found ${data.tabs.length} tab(s) in workbook. Map the columns below:`;

                renderTabMappings(data.tabs);
            } catch (err) {
                statusEl.className = "mt-3 text-sm font-medium text-red-600";
                statusEl.innerText = "❌ Failed to read Excel file.";
            }
        }

        // Helper to auto-guess column names
        function guessColumn(columns, keywords) {
            for (const col of columns) {
                const lower = col.toLowerCase();
                if (keywords.some(k => lower.includes(k))) return col;
            }
            return "";
        }

        // Render Column Selectors for Each Tab
        function renderTabMappings(tabs) {
            const section = document.getElementById('mapping-section');
            const container = document.getElementById('tabs-container');
            section.classList.remove('hidden');

            container.innerHTML = tabs.map((tab) => {
                const barcodeGuess = guessColumn(tab.columns, ['barcode', 'code', 'sku', 'art', 'ean', 'id']);
                const nameGuess = guessColumn(tab.columns, ['name', 'desc', 'product', 'article', 'item', 'nome']);
                const priceGuess = guessColumn(tab.columns, ['price', 'rate', 'mrp', 'preco', 'preço', 'valor', 'amount']);
                const catGuess = guessColumn(tab.columns, ['cat', 'dept', 'group', 'unit', 'brand']);

                const makeOptions = (selectedVal, allowEmpty = false) => {
                    let opts = allowEmpty ? `<option value="">-- Optional / Skip --</option>` : `<option value="">-- Select Column --</option>`;
                    for (const col of tab.columns) {
                        const isSel = (col === selectedVal) ? "selected" : "";
                        opts += `<option value="${col}" ${isSel}>${col}</option>`;
                    }
                    return opts;
                };

                return `
                <div class="border-2 border-slate-200 rounded-xl p-5 bg-slate-50 tab-config-card" data-tab-name="${tab.sheet_name}">
                    <div class="flex items-center justify-between pb-3 mb-4 border-b border-slate-200">
                        <label class="flex items-center space-x-3 cursor-pointer">
                            <input type="checkbox" class="tab-enable-checkbox w-5 h-5 text-indigo-600 rounded" checked>
                            <span class="text-lg font-bold text-indigo-900">📑 Tab: ${tab.sheet_name}</span>
                            <span class="text-xs bg-slate-200 text-slate-700 px-2.5 py-0.5 rounded-full font-semibold">${tab.row_count} rows</span>
                        </label>
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                        <div>
                            <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Barcode / Article Code *</label>
                            <select class="col-barcode w-full p-2.5 bg-white border border-slate-300 rounded-lg font-medium text-sm">
                                ${makeOptions(barcodeGuess, false)}
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Article Name *</label>
                            <select class="col-name w-full p-2.5 bg-white border border-slate-300 rounded-lg font-medium text-sm">
                                ${makeOptions(nameGuess, false)}
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Selling Price *</label>
                            <select class="col-price w-full p-2.5 bg-white border border-slate-300 rounded-lg font-medium text-sm">
                                ${makeOptions(priceGuess, false)}
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Category / Unit (Optional)</label>
                            <select class="col-category w-full p-2.5 bg-white border border-slate-300 rounded-lg font-medium text-sm">
                                ${makeOptions(catGuess, true)}
                            </select>
                        </div>
                    </div>

                    <!-- Sample Preview Table -->
                    <div class="mt-4 overflow-x-auto">
                        <p class="text-xs font-semibold text-slate-500 mb-1">Sample Preview (First 2 Rows):</p>
                        <table class="w-full text-xs text-left bg-white border border-slate-200 rounded-lg">
                            <thead class="bg-slate-100 text-slate-700">
                                <tr>${tab.columns.map(c => `<th class="px-2 py-1 border-b">${c}</th>`).join('')}</tr>
                            </thead>
                            <tbody>
                                ${tab.sample.map(row => `
                                    <tr class="border-b">${tab.columns.map(c => `<td class="px-2 py-1 text-slate-600">${row[c] ?? ''}</td>`).join('')}</tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>`;
            }).join('');
        }

        // Step 2: Confirm Column Mappings and Update Database
        async function confirmAndImport() {
            const cards = document.querySelectorAll('.tab-config-card');
            const mappings = [];

            for (const card of cards) {
                const enabled = card.querySelector('.tab-enable-checkbox').checked;
                if (!enabled) continue;

                const sheetName = card.getAttribute('data-tab-name');
                const barcodeCol = card.querySelector('.col-barcode').value;
                const nameCol = card.querySelector('.col-name').value;
                const priceCol = card.querySelector('.col-price').value;
                const categoryCol = card.querySelector('.col-category').value;

                if (!barcodeCol || !nameCol || !priceCol) {
                    alert(`Please select Barcode, Article Name, and Price columns for tab: "${sheetName}"`);
                    return;
                }

                mappings.push({
                    sheet_name: sheetName,
                    barcode_col: barcodeCol,
                    name_col: nameCol,
                    price_col: priceCol,
                    category_col: categoryCol
                });
            }

            if (mappings.length === 0) {
                alert("Please select at least one tab to import.");
                return;
            }

            const btn = document.getElementById('btn-confirm-import');
            btn.disabled = true;
            btn.innerText = "⏳ Updating Prices...";

            try {
                const res = await fetch("/api/import-mapped", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        temp_filename: currentTempFile,
                        mappings: mappings
                    })
                });
                const data = await res.json();

                btn.disabled = false;
                btn.innerText = "✅ Save & Update Daily Prices";

                if (data.success) {
                    alert(`✅ Success! Updated ${data.imported_count} articles across ${mappings.length} tab(s).`);
                    document.getElementById('total-count-badge').innerText = `Total Articles: ${data.total_products}`;
                    document.getElementById('mapping-section').classList.add('hidden');
                } else {
                    alert("❌ Error: " + data.error);
                }
            } catch (err) {
                btn.disabled = false;
                btn.innerText = "✅ Save & Update Daily Prices";
                alert("❌ Failed to import data.");
            }
        }
    </script>
</body>
</html>
"""


# --- ROUTES ---
@app.route("/")
def index():
    with sqlite3.connect(DB_NAME) as conn:
        total = conn.execute("SELECT COUNT(*) FROM products").fetchone()[0]
    return render_template_string(HTML_TEMPLATE, total_products=total)


@app.route("/api/search")
def api_search():
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"found": False, "items": []})

    with sqlite3.connect(DB_NAME) as conn:
        conn.row_factory = sqlite3.Row
        # 1. Exact Barcode Match first
        rows = conn.execute(
            "SELECT * FROM products WHERE barcode = ?", (q,)
        ).fetchall()

        # 2. If not found by exact barcode, search by Article Name or partial barcode
        if not rows:
            like_q = f"%{q}%"
            rows = conn.execute(
                "SELECT * FROM products WHERE name LIKE ? OR barcode LIKE ? LIMIT 15",
                (like_q, like_q),
            ).fetchall()

    items = [dict(r) for r in rows]
    return jsonify({"found": len(items) > 0, "items": items})


@app.route("/api/preview-excel", methods=["POST"])
def preview_excel():
    if "file" not in request.files:
        return jsonify({"success": False, "error": "No file uploaded"})

    f = request.files["file"]
    if not f.filename:
        return jsonify({"success": False, "error": "Empty filename"})

    ext = os.path.splitext(f.filename)[1].lower()
    temp_filename = f"{uuid.uuid4().hex}{ext}"
    filepath = os.path.join(UPLOAD_FOLDER, temp_filename)
    f.save(filepath)

    try:
        tabs_info = []
        if ext == ".csv":
            df = pd.read_csv(filepath)
            df = df.dropna(how="all")
            cols = [str(c) for c in df.columns]
            sample = df.head(2).fillna("").astype(str).to_dict(orient="records")
            tabs_info.append({
                "sheet_name": "CSV_Data",
                "columns": cols,
                "row_count": len(df),
                "sample": sample,
            })
        else:
            excel_file = pd.ExcelFile(filepath)
            for sheet in excel_file.sheet_names:
                df = excel_file.parse(sheet)
                df = df.dropna(how="all")
                if len(df.columns) == 0:
                    continue
                cols = [str(c) for c in df.columns]
                sample = (
                    df.head(2).fillna("").astype(str).to_dict(orient="records")
                )
                tabs_info.append({
                    "sheet_name": sheet,
                    "columns": cols,
                    "row_count": len(df),
                    "sample": sample,
                })

        if not tabs_info:
            return jsonify({
                "success": False,
                "error": "No valid data tabs found in the file.",
            })

        return jsonify({
            "success": True,
            "temp_filename": temp_filename,
            "tabs": tabs_info,
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/import-mapped", methods=["POST"])
def import_mapped():
    data = request.get_json()
    temp_filename = data.get("temp_filename")
    mappings = data.get("mappings", [])

    if not temp_filename or not mappings:
        return jsonify(
            {"success": False, "error": "Missing file or column mappings"}
        )

    filepath = os.path.join(UPLOAD_FOLDER, temp_filename)
    if not os.path.exists(filepath):
        return jsonify({
            "success": False,
            "error": "Temporary file expired. Please upload again.",
        })

    ext = os.path.splitext(filepath)[1].lower()
    now_str = datetime.now().strftime("%d-%b-%Y %H:%M")
    imported_count = 0

    try:
        with sqlite3.connect(DB_NAME) as conn:
            for m in mappings:
                sheet = m["sheet_name"]
                b_col = m["barcode_col"]
                n_col = m["name_col"]
                p_col = m["price_col"]
                c_col = m.get("category_col", "")

                if ext == ".csv":
                    df = pd.read_csv(filepath)
                else:
                    df = pd.read_excel(filepath, sheet_name=sheet)

                df = df.dropna(how="all")

                records_to_upsert = []
                for _, row in df.iterrows():
                    barcode = clean_barcode(row.get(b_col))
                    name = (
                        str(row.get(n_col, "")).strip()
                        if not pd.isna(row.get(n_col))
                        else ""
                    )
                    price_val = row.get(p_col)

                    if not barcode or not name or pd.isna(price_val):
                        continue

                    if isinstance(price_val, (int, float)):
                        price_str = f"{price_val:,.2f}"
                    else:
                        price_str = str(price_val).strip()

                    category = ""
                    if c_col and c_col in df.columns and not pd.isna(row.get(c_col)):
                        category = str(row.get(c_col)).strip()

                    records_to_upsert.append(
                        (barcode, name, price_str, category, sheet, now_str)
                    )

                conn.executemany(
                    """
                    INSERT INTO products (barcode, name, price, category, source_tab, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(barcode) DO UPDATE SET
                        name=excluded.name,
                        price=excluded.price,
                        category=excluded.category,
                        source_tab=excluded.source_tab,
                        updated_at=excluded.updated_at
                """,
                    records_to_upsert,
                )

                imported_count += len(records_to_upsert)
            conn.commit()

            total_products = conn.execute(
                "SELECT COUNT(*) FROM products"
            ).fetchone()[0]

        try:
            os.remove(filepath)
        except OSError:
            pass

        return jsonify({
            "success": True,
            "imported_count": imported_count,
            "total_products": total_products,
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

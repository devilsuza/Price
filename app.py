import os
import uuid
import sqlite3
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string, Response
import pandas as pd

app = Flask(__name__)
UPLOAD_FOLDER = "temp_uploads"
DB_NAME = "store_prices.db"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# --- DATABASE SETUP & AUTO-MIGRATION ---
def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS products (
                barcode TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                price_cx TEXT DEFAULT '',
                price_emb TEXT DEFAULT '',
                price_pc_kgm TEXT DEFAULT '',
                price_bag TEXT DEFAULT '',
                updated_at TEXT
            )
        """)
        existing_cols = [
            row[1] for row in conn.execute("PRAGMA table_info(products)").fetchall()
        ]
        for col in ["price_cx", "price_emb", "price_pc_kgm", "price_bag", "updated_at"]:
            if col not in existing_cols:
                conn.execute(f"ALTER TABLE products ADD COLUMN {col} TEXT DEFAULT ''")
        conn.commit()


init_db()


def clean_barcode(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def format_price(val):
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (int, float)):
        if float(val) == 0:
            return ""
        return f"{val:,.2f}"
    s = str(val).strip()
    if s in ["", "-", "0", "0.0", "0.00", "nan", "None"]:
        return ""
    return s


# --- PWA MANIFEST JSON ---
MANIFEST_JSON = """
{
  "name": "Store Price Portal",
  "short_name": "PricePortal",
  "start_url": "/",
  "display": "standalone",
  "background_color": "#f1f5f9",
  "theme_color": "#4338ca",
  "icons": [
    {
      "src": "https://cdn-icons-png.flaticon.com/512/3081/3081559.png",
      "sizes": "512x512",
      "type": "image/png"
    }
  ]
}
"""


# --- SERVICE WORKER FOR OFFLINE CACHING & PWA ---
SERVICE_WORKER_JS = """
const CACHE_NAME = 'store-price-portal-v4';
const ASSETS_TO_CACHE = [
    '/',
    'https://cdn.tailwindcss.com',
    'https://unpkg.com/html5-qrcode',
    'https://cdn-icons-png.flaticon.com/512/3081/3081559.png'
];

self.addEventListener('install', (event) => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return cache.addAll(ASSETS_TO_CACHE).catch(() => {});
        })
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((keys) =>
            Promise.all(keys.map((k) => k !== CACHE_NAME && caches.delete(k)))
        )
    );
    self.clients.claim();
});

self.addEventListener('fetch', (event) => {
    if (event.request.method !== 'GET') return;
    const url = new URL(event.request.url);

    if (url.pathname.startsWith('/api/')) return;

    event.respondWith(
        fetch(event.request)
            .then((response) => {
                const cloned = response.clone();
                caches.open(CACHE_NAME).then((cache) => cache.put(event.request, cloned));
                return response;
            })
            .catch(() => caches.match(event.request))
    );
});
"""


# --- HTML TEMPLATE (WITH MANAGER EDITABLE PRODUCT TABLE) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Store Price Checker & Manager Portal</title>
    <link rel="manifest" href="/manifest.json">
    <meta name="theme-color" content="#4338ca">
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
            <p class="text-xs text-slate-500 mt-1 mb-5">Enter PIN for Employee or Manager</p>

            <input type="password" id="login-pin" autofocus
                placeholder="Enter Password"
                class="w-full px-4 py-3 text-center text-xl tracking-widest font-bold border-2 border-indigo-300 rounded-xl focus:outline-none focus:border-indigo-600 mb-3">
            
            <p id="login-error" class="text-xs font-bold text-red-600 mb-3 hidden">❌ Wrong Password! Try again.</p>

            <button onclick="handleLogin()" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 rounded-xl shadow-md transition">
                Login Now
            </button>
        </div>
    </div>

    <!-- ================= TOP NAVIGATION ================= -->
    <nav class="bg-indigo-700 text-white shadow-lg">
        <div class="max-w-7xl mx-auto px-4 py-3 flex flex-wrap justify-between items-center gap-2">
            <div class="flex items-center flex-wrap gap-2">
                <span class="text-xl font-bold tracking-wide">🏪 Store Price Portal</span>
                <span id="role-badge" class="text-xs bg-indigo-900 text-indigo-200 px-2.5 py-0.5 rounded-full font-semibold hidden"></span>
                
                <div class="flex items-center space-x-1 bg-indigo-800/90 px-3 py-1 rounded-full border border-indigo-500 text-xs font-semibold">
                    <span id="net-status-dot" class="w-2.5 h-2.5 rounded-full bg-emerald-400 inline-block"></span>
                    <span id="net-status-text">Checking...</span>
                    <button onclick="syncDatabaseInBackground(true)" id="btn-manual-sync" title="Sync Database Now" class="ml-1.5 text-indigo-200 hover:text-white">
                        🔄
                    </button>
                </div>
            </div>

            <div class="flex items-center space-x-2">
                <button id="btn-install-app" onclick="installPWA()" class="hidden px-3 py-1.5 rounded-lg text-sm font-bold bg-emerald-500 hover:bg-emerald-600 text-white shadow-md flex items-center gap-1">
                    📱 Install App
                </button>
                <button onclick="switchMode('scanner')" id="btn-scanner" class="px-3.5 py-1.5 rounded-lg text-sm font-medium bg-white text-indigo-700 shadow">
                    🔍 Scanner
                </button>
                <button onclick="switchMode('manager')" id="btn-manager" class="px-3.5 py-1.5 rounded-lg text-sm font-medium text-indigo-100 hover:bg-indigo-600">
                    ⚙️ Manager Panel
                </button>
                <button onclick="logoutPortal()" class="px-3 py-1.5 rounded-lg text-sm font-medium bg-red-500/90 hover:bg-red-600 text-white shadow">
                    🔒 Logout
                </button>
            </div>
        </div>
    </nav>

    <div class="max-w-5xl mx-auto p-4 mt-2">

        <!-- Offline Mode Banner -->
        <div id="offline-banner" class="hidden mb-4 bg-amber-50 border-2 border-amber-400 text-amber-900 px-4 py-3 rounded-xl flex items-center justify-between text-sm font-semibold shadow-sm">
            <div>
                ⚡ <strong>Offline Mode Active:</strong> Internet band hai, par scanning chalu hai! Data local memory se aa raha hai.
            </div>
            <span id="offline-count-label" class="text-xs bg-amber-200 text-amber-900 px-2.5 py-1 rounded-lg font-bold whitespace-nowrap ml-2"></span>
        </div>

        <!-- ================= EMPLOYEE SCANNER VIEW ================= -->
        <div id="view-scanner" class="space-y-6">
            <div class="bg-white p-6 rounded-2xl shadow-md border border-slate-200">
                <div class="flex justify-between items-center mb-3">
                    <h2 class="text-lg font-bold text-slate-700">Scan Article Barcode or Search Name</h2>
                    <span id="last-sync-time" class="text-[11px] text-slate-400 font-medium"></span>
                </div>
                
                <div class="flex flex-col sm:flex-row gap-3">
                    <div class="relative flex-1">
                        <input type="text" id="scan-input" inputmode="none" autocomplete="off" autocapitalize="off"
                            placeholder="Scan barcode here or type article code / name..."
                            class="w-full pl-4 pr-14 py-3 text-lg border-2 border-indigo-300 rounded-xl focus:outline-none focus:border-indigo-600">
                        
                        <button type="button" onclick="toggleCamera()" id="inline-scan-btn"
                            title="Open QR / Barcode Camera Scanner"
                            class="absolute right-2 top-1/2 -translate-y-1/2 p-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-lg border border-indigo-200 transition flex items-center justify-center shadow-sm">
                            <svg xmlns="http://www.w3.org/2000/svg" class="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M3 7V5a2 2 0 012-2h2m10 0h2a2 2 0 012 2v2m0 10v2a2 2 0 01-2 2h-2M7 21H5a2 2 0 01-2-2v-2M7 8h1v8H7V8zm3 0h2v8h-2V8zm4 0h1v8h-1V8zm3 0h1v8h-1V8z" />
                            </svg>
                        </button>
                    </div>

                    <button onclick="searchArticle()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold px-6 py-3 rounded-xl shadow">
                        Check Price
                    </button>
                </div>

                <div id="reader-container" class="hidden mt-4 bg-slate-900 p-3 rounded-2xl max-w-sm mx-auto text-center">
                    <div class="flex justify-between items-center text-white text-xs font-semibold mb-2 px-1">
                        <span>📷 Point camera at Barcode or QR Code</span>
                        <button onclick="stopCamera()" class="bg-red-500 hover:bg-red-600 text-white px-2.5 py-1 rounded-md">✕ Close</button>
                    </div>
                    <div id="reader" class="w-full overflow-hidden rounded-xl border-2 border-emerald-400 bg-black"></div>
                </div>
            </div>

            <div id="price-result" class="hidden space-y-4"></div>
        </div>

        <!-- ================= MANAGER UPLOAD & EDIT VIEW ================= -->
        <div id="view-manager" class="hidden space-y-6">
            
            <!-- Step 1: Upload Excel -->
            <div class="bg-white p-6 rounded-2xl shadow-md border border-slate-200">
                <div class="flex flex-wrap justify-between items-center gap-2 mb-4">
                    <div>
                        <h2 class="text-xl font-bold text-slate-800">1. Upload Daily Price Excel (Auto-Merge)</h2>
                        <p class="text-sm text-slate-500">New articles added & existing prices updated automatically</p>
                    </div>
                    <span class="bg-indigo-50 text-indigo-700 text-xs font-bold px-3 py-1.5 rounded-full border border-indigo-200" id="total-count-badge">
                        Total Articles in DB: {{ total_products }}
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
                    <h2 class="text-xl font-bold text-slate-800">2. Select Columns for Each Tab</h2>
                    <p class="text-sm text-slate-500">Map Barcode, Article Name, and the 4 Price columns (CX, EMB, PC/KGM, BAG).</p>
                </div>

                <div id="tabs-container" class="space-y-6"></div>

                <div class="pt-4 border-t border-slate-200 flex justify-end">
                    <button onclick="confirmAndImport()" id="btn-confirm-import"
                        class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-8 py-3 rounded-xl shadow-lg text-base">
                        ✅ Merge & Update Prices
                    </button>
                </div>
            </div>

            <!-- Step 3: Searchable Product List & Inline Editing Table -->
            <div class="bg-white p-6 rounded-2xl shadow-md border border-slate-200 space-y-4">
                <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                    <div>
                        <h2 class="text-xl font-bold text-slate-800">3. Live Product Catalog & Manual Price Editor</h2>
                        <p class="text-sm text-slate-500">Search any article and click inside any price box to edit and save instantly.</p>
                    </div>
                    <button onclick="loadManagerProductTable()" class="bg-slate-800 hover:bg-slate-900 text-white font-semibold px-4 py-2 rounded-xl text-sm shadow">
                        🔄 Refresh Table
                    </button>
                </div>

                <!-- Search Filter Bar -->
                <div class="relative">
                    <input type="text" id="manager-search-input" oninput="filterManagerTable()"
                        placeholder="Type article code or name to filter table..."
                        class="w-full pl-4 pr-10 py-2.5 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-indigo-600 text-sm">
                </div>

                <!-- Product Table Container -->
                <div class="overflow-x-auto border border-slate-200 rounded-xl max-h-[500px]">
                    <table class="w-full text-xs text-left bg-white">
                        <thead class="bg-slate-100 text-slate-700 sticky top-0 uppercase font-bold tracking-wider z-10">
                            <tr>
                                <th class="px-3 py-3 border-b">Article Code</th>
                                <th class="px-3 py-3 border-b">Article Name</th>
                                <th class="px-3 py-3 border-b text-center">CX</th>
                                <th class="px-3 py-3 border-b text-center">EMB</th>
                                <th class="px-3 py-3 border-b text-center">PC / KGM</th>
                                <th class="px-3 py-3 border-b text-center">BAG</th>
                                <th class="px-3 py-3 border-b text-center">Action</th>
                            </tr>
                        </thead>
                        <tbody id="manager-table-body" class="divide-y divide-slate-200">
                            <tr><td colspan="7" class="text-center py-6 text-slate-400">Loading products... Click Refresh Table if empty.</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>

        </div>

    </div>

    <script>
        const EMPLOYEE_PASS = "123";
        const MANAGER_PASS = "2026";
        let currentRole = null;

        let currentTempFile = null;
        let html5QrCode = null;
        let cameraActive = false;
        let allManagerProducts = [];

        // PWA Install Prompt
        let deferredPrompt = null;
        window.addEventListener('beforeinstallprompt', (e) => {
            e.preventDefault();
            deferredPrompt = e;
            document.getElementById('btn-install-app').classList.remove('hidden');
        });

        async function installPWA() {
            if (!deferredPrompt) {
                alert("App already installed or browser doesn't support direct installation.");
                return;
            }
            deferredPrompt.prompt();
            const { outcome } = await deferredPrompt.userChoice;
            if (outcome === 'accepted') {
                document.getElementById('btn-install-app').classList.add('hidden');
            }
            deferredPrompt = null;
        }

        // Offline IndexedDB Engine
        const IDB_NAME = "StorePriceOfflineDB";
        const IDB_VERSION = 1;
        const STORE_NAME = "products";
        let dbPromise = null;
        let localItemCount = 0;
        let isSyncing = false;

        if ('serviceWorker' in navigator) {
            window.addEventListener('load', () => {
                navigator.serviceWorker.register('/sw.js').catch(err => console.log('SW Error:', err));
            });
        }

        function openLocalDB() {
            if (dbPromise) return dbPromise;
            dbPromise = new Promise((resolve, reject) => {
                const request = indexedDB.open(IDB_NAME, IDB_VERSION);
                request.onupgradeneeded = (e) => {
                    const db = e.target.result;
                    if (!db.objectStoreNames.contains(STORE_NAME)) {
                        db.createObjectStore(STORE_NAME, { keyPath: "barcode" });
                    }
                };
                request.onsuccess = (e) => resolve(e.target.result);
                request.onerror = (e) => reject(e.target.error);
            });
            return dbPromise;
        }

        async function saveAllToLocalDB(items) {
            const db = await openLocalDB();
            return new Promise((resolve, reject) => {
                const tx = db.transaction(STORE_NAME, "readwrite");
                const store = tx.objectStore(STORE_NAME);
                store.clear();
                for (const item of items) {
                    store.put(item);
                }
                tx.oncomplete = () => {
                    localItemCount = items.length;
                    const nowStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                    localStorage.setItem("last_sync_time", nowStr);
                    localStorage.setItem("local_item_count", String(items.length));
                    resolve(items.length);
                };
                tx.onerror = (e) => reject(e.target.error);
            });
        }

        async function searchInLocalDB(query) {
            const db = await openLocalDB();
            const cleanQ = query.trim();
            const lowerQ = cleanQ.toLowerCase();

            return new Promise((resolve) => {
                const tx = db.transaction(STORE_NAME, "readonly");
                const store = tx.objectStore(STORE_NAME);

                const exactReq = store.get(cleanQ);
                exactReq.onsuccess = () => {
                    if (exactReq.result) {
                        resolve([exactReq.result]);
                    } else {
                        const matches = [];
                        const cursorReq = store.openCursor();
                        cursorReq.onsuccess = (e) => {
                            const cursor = e.target.result;
                            if (cursor && matches.length < 15) {
                                const item = cursor.value;
                                const bMatch = item.barcode && item.barcode.toLowerCase().includes(lowerQ);
                                const nMatch = item.name && item.name.toLowerCase().includes(lowerQ);
                                if (bMatch || nMatch) {
                                    matches.push(item);
                                }
                                cursor.continue();
                            } else {
                                resolve(matches);
                            }
                        };
                        cursorReq.onerror = () => resolve([]);
                    }
                };
                exactReq.onerror = () => resolve([]);
            });
        }

        function updateNetworkUI(syncingState = false) {
            const dot = document.getElementById('net-status-dot');
            const text = document.getElementById('net-status-text');
            const banner = document.getElementById('offline-banner');
            const offlineCountLbl = document.getElementById('offline-count-label');
            const lastSyncEl = document.getElementById('last-sync-time');

            const savedCount = localStorage.getItem("local_item_count") || localItemCount || 0;
            const savedTime = localStorage.getItem("last_sync_time") || "";

            if (savedTime && lastSyncEl) {
                lastSyncEl.innerText = `Local DB Synced at ${savedTime} (${savedCount} articles)`;
            }

            if (syncingState) {
                dot.className = "w-2.5 h-2.5 rounded-full bg-blue-400 animate-ping inline-block";
                text.innerText = "Syncing DB...";
                return;
            }

            if (navigator.onLine) {
                dot.className = "w-2.5 h-2.5 rounded-full bg-emerald-400 inline-block";
                text.innerText = `Online (${savedCount} Offline Ready)`;
                banner.classList.add('hidden');
            } else {
                dot.className = "w-2.5 h-2.5 rounded-full bg-amber-400 inline-block";
                text.innerText = `Offline Mode (${savedCount} items)`;
                banner.classList.remove('hidden');
                if (offlineCountLbl) offlineCountLbl.innerText = `${savedCount} Articles Ready`;
            }
        }

        async function syncDatabaseInBackground(showAlert = false) {
            if (!navigator.onLine || isSyncing) {
                updateNetworkUI();
                if (showAlert && !navigator.onLine) {
                    alert("⚡ Aap abhi Offline hain! Internet connect hone par automatic sync ho jayega.");
                }
                return;
            }

            isSyncing = true;
            updateNetworkUI(true);

            try {
                const res = await fetch("/api/sync-all", { cache: "no-store" });
                const data = await res.json();
                if (data.success && Array.isArray(data.items)) {
                    await saveAllToLocalDB(data.items);
                    document.getElementById('total-count-badge').innerText = `Total Articles in DB: ${data.items.length}`;
                    if (showAlert) {
                        alert(`✅ Sync Complete! ${data.items.length} articles saved for Offline Mode.`);
                    }
                }
            } catch (err) {
                console.log("Background sync skipped:", err);
            } finally {
                isSyncing = false;
                updateNetworkUI(false);
            }
        }

        window.addEventListener('online', () => {
            updateNetworkUI();
            syncDatabaseInBackground(false);
        });

        window.addEventListener('offline', () => {
            updateNetworkUI();
        });

        window.addEventListener('DOMContentLoaded', () => {
            updateNetworkUI();
            syncDatabaseInBackground(false);
            setInterval(() => {
                if (navigator.onLine) syncDatabaseInBackground(false);
            }, 10 * 60 * 1000);
        });


        // Login & Scanner Logic
        function playBeep() {
            try {
                const AudioCtx = window.AudioContext || window.webkitAudioContext;
                const ctx = new AudioCtx();
                const osc = ctx.createOscillator();
                const gain = ctx.createGain();

                osc.type = "sine";
                osc.frequency.setValueAtTime(1350, ctx.currentTime);
                gain.gain.setValueAtTime(0.25, ctx.currentTime);

                osc.connect(gain);
                gain.connect(ctx.destination);

                osc.start();
                osc.stop(ctx.currentTime + 0.14);
            } catch (e) {}
        }

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

        function switchMode(mode) {
            if (mode === 'manager' && currentRole !== 'manager') {
                const askPin = prompt("🔒 Enter Manager Password to access Manager Panel:");
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
                prepareScannerInput();
            } else {
                scannerView.classList.add('hidden');
                managerView.classList.remove('hidden');
                btnManager.className = "px-3.5 py-1.5 rounded-lg text-sm font-medium bg-white text-indigo-700 shadow";
                btnScanner.className = "px-3.5 py-1.5 rounded-lg text-sm font-medium text-indigo-100 hover:bg-indigo-600";
                stopCamera();
                loadManagerProductTable(); // Load table when entering manager mode
            }
        }

        function prepareScannerInput() {
            const el = document.getElementById('scan-input');
            el.blur();
            el.setAttribute('inputmode', 'none');
        }
        document.getElementById('scan-input').addEventListener('pointerdown', function() {
            this.setAttribute('inputmode', 'text');
        });
        document.getElementById('scan-input').addEventListener('blur', function() {
            this.setAttribute('inputmode', 'none');
        });

        document.getElementById('scan-input').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                playBeep();
                searchArticle();
            }
        });

        function renderPriceBox(label, subLabel, value, colorClass) {
            const hasPrice = value && value.trim() !== "";
            return `
                <div class="${hasPrice ? colorClass : 'bg-slate-50 border-slate-200 opacity-60'} border-2 rounded-2xl p-4 text-center shadow-sm transition">
                    <span class="text-xs font-extrabold uppercase tracking-wider block text-slate-600">${label}</span>
                    <span class="text-[11px] text-slate-400 block mb-1">${subLabel}</span>
                    <span class="${hasPrice ? 'text-2xl sm:text-3xl font-black text-slate-900' : 'text-xl font-bold text-slate-400'} block mt-1">
                        ${hasPrice ? value : '--'}
                    </span>
                </div>
            `;
        }

        function displaySearchResults(items, query, sourceLabel) {
            const resultBox = document.getElementById('price-result');
            resultBox.classList.remove('hidden');

            if (!items || items.length === 0) {
                resultBox.innerHTML = `
                    <div class="bg-red-50 border-2 border-red-300 p-6 rounded-2xl text-center">
                        <p class="text-2xl font-bold text-red-600">❌ Article Not Found</p>
                        <p class="text-slate-600 mt-1">No record found for: <strong>${query}</strong> (${sourceLabel})</p>
                    </div>`;
                return;
            }

            resultBox.innerHTML = items.map(item => `
                <div class="bg-white border-2 border-indigo-500 p-6 rounded-2xl shadow-lg space-y-5">
                    <div class="border-b border-slate-200 pb-4 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <div class="flex items-center gap-2">
                                <span class="text-xs font-extrabold uppercase tracking-wider bg-indigo-100 text-indigo-800 px-3 py-1 rounded-lg">
                                    Article Code: ${item.barcode}
                                </span>
                                <span class="text-[11px] font-semibold px-2.5 py-0.5 rounded-md ${sourceLabel.includes('Offline') ? 'bg-amber-100 text-amber-800' : 'bg-emerald-100 text-emerald-800'}">
                                    ${sourceLabel}
                                </span>
                            </div>
                            <h3 class="text-2xl sm:text-3xl font-black text-slate-900 mt-2">${item.name}</h3>
                        </div>
                    </div>

                    <div class="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
                        ${renderPriceBox('Price', '1', item.price_cx, 'bg-emerald-50 border-emerald-400')}
                        ${renderPriceBox('Price', '2', item.price_emb, 'bg-blue-50 border-blue-400')}
                        ${renderPriceBox('Price', '3', item.price_pc_kgm, 'bg-amber-50 border-amber-400')}
                        ${renderPriceBox('Price', '4', item.price_bag, 'bg-purple-50 border-purple-400')}
                    </div>
                </div>
            `).join('');
        }

        async function searchArticle(overrideQuery = null) {
            const inputEl = document.getElementById('scan-input');
            const query = (overrideQuery !== null ? overrideQuery : inputEl.value).trim();
            if (!query) return;

            const resultBox = document.getElementById('price-result');
            resultBox.classList.remove('hidden');

            if (!navigator.onLine) {
                const localItems = await searchInLocalDB(query);
                displaySearchResults(localItems, query, "⚡ Offline Local DB");
                inputEl.select();
                return;
            }

            resultBox.innerHTML = `<div class="bg-white p-6 rounded-2xl shadow text-center text-slate-500 font-medium">⏳ Fetching price details...</div>`;

            try {
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), 2500);

                const res = await fetch(`/api/search?q=${encodeURIComponent(query)}`, {
                    signal: controller.signal
                });
                clearTimeout(timeoutId);

                const data = await res.json();
                displaySearchResults(data.items || [], query, "🟢 Online Live DB");
                inputEl.select();
            } catch (err) {
                const localItems = await searchInLocalDB(query);
                displaySearchResults(localItems, query, "⚡ Offline Local DB");
                inputEl.select();
            }
        }

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
                        playBeep();
                        const scanBox = document.getElementById('scan-input');
                        scanBox.value = decodedText;
                        scanBox.blur();
                        scanBox.setAttribute('inputmode', 'none');
                        stopCamera();
                        searchArticle(decodedText);
                    },
                    (errorMessage) => {}
                ).catch(err => {
                    alert("Camera open nahi ho paya. Please allow camera permission.");
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

        // =========================================================================
        // MANAGER EDITABLE PRODUCT TABLE LOGIC
        // =========================================================================
        async function loadManagerProductTable() {
            const tbody = document.getElementById('manager-table-body');
            tbody.innerHTML = `<tr><td colspan="7" class="text-center py-6 text-slate-400">Loading products from database...</td></tr>`;

            try {
                const res = await fetch("/api/sync-all");
                const data = await res.json();
                if (data.success && Array.isArray(data.items)) {
                    allManagerProducts = data.items;
                    renderManagerTable(allManagerProducts);
                } else {
                    tbody.innerHTML = `<tr><td colspan="7" class="text-center py-6 text-red-500">Failed to load products.</td></tr>`;
                }
            } catch (err) {
                tbody.innerHTML = `<tr><td colspan="7" class="text-center py-6 text-red-500">Network error loading products.</td></tr>`;
            }
        }

        function renderManagerTable(items) {
            const tbody = document.getElementById('manager-table-body');
            if (items.length === 0) {
                tbody.innerHTML = `<tr><td colspan="7" class="text-center py-6 text-slate-400">No products found in database.</td></tr>`;
                return;
            }

            tbody.innerHTML = items.map(item => `
                <tr class="hover:bg-slate-50 transition" id="row-${item.barcode}">
                    <td class="px-3 py-2.5 font-bold text-slate-700">${item.barcode}</td>
                    <td class="px-3 py-2.5 font-medium text-slate-900 max-w-xs truncate" title="${item.name}">${item.name}</td>
                    <td class="px-2 py-2 text-center"><input type="text" id="cx-${item.barcode}" value="${item.price_cx || ''}" class="w-20 p-1.5 text-center border border-slate-300 rounded-lg text-xs font-semibold focus:border-indigo-600 bg-slate-50"></td>
                    <td class="px-2 py-2 text-center"><input type="text" id="emb-${item.barcode}" value="${item.price_emb || ''}" class="w-20 p-1.5 text-center border border-slate-300 rounded-lg text-xs font-semibold focus:border-indigo-600 bg-slate-50"></td>
                    <td class="px-2 py-2 text-center"><input type="text" id="pc-${item.barcode}" value="${item.price_pc_kgm || ''}" class="w-20 p-1.5 text-center border border-slate-300 rounded-lg text-xs font-semibold focus:border-indigo-600 bg-slate-50"></td>
                    <td class="px-2 py-2 text-center"><input type="text" id="bag-${item.barcode}" value="${item.price_bag || ''}" class="w-20 p-1.5 text-center border border-slate-300 rounded-lg text-xs font-semibold focus:border-indigo-600 bg-slate-50"></td>
                    <td class="px-2 py-2 text-center">
                        <button onclick="saveInlineProduct('${item.barcode}')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow">
                            💾 Save
                        </button>
                    </td>
                </tr>
            `).join('');
        }

        function filterManagerTable() {
            const query = document.getElementById('manager-search-input').value.toLowerCase().trim();
            if (!query) {
                renderManagerTable(allManagerProducts);
                return;
            }
            const filtered = allManagerProducts.filter(item => 
                (item.barcode && item.barcode.toLowerCase().includes(query)) || 
                (item.name && item.name.toLowerCase().includes(query))
            );
            renderManagerTable(filtered);
        }

        async function saveInlineProduct(barcode) {
            const price_cx = document.getElementById(`cx-${barcode}`).value.trim();
            const price_emb = document.getElementById(`emb-${barcode}`).value.trim();
            const price_pc_kgm = document.getElementById(`pc-${barcode}`).value.trim();
            const price_bag = document.getElementById(`bag-${barcode}`).value.trim();

            try {
                const res = await fetch("/api/update-single", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ barcode, price_cx, price_emb, price_pc_kgm, price_bag })
                });
                const data = await res.json();
                if (data.success) {
                    // Update local cache array
                    const idx = allManagerProducts.findIndex(p => p.barcode === barcode);
                    if (idx !== -1) {
                        allManagerProducts[idx].price_cx = price_cx;
                        allManagerProducts[idx].price_emb = price_emb;
                        allManagerProducts[idx].price_pc_kgm = price_pc_kgm;
                        allManagerProducts[idx].price_bag = price_bag;
                    }
                    await syncDatabaseInBackground(false); // Resync offline DB
                    alert(`✅ Prices updated successfully for article: ${barcode}`);
                } else {
                    alert("❌ Error updating: " + data.error);
                }
            } catch (err) {
                alert("❌ Failed to update price.");
            }
        }


        // =========================================================================
        // EXCEL UPLOAD MAPPING LOGIC
        // =========================================================================
        async function uploadForPreview() {
            if (!navigator.onLine) {
                alert("❌ Excel upload karne ke liye internet connection chahiye.");
                return;
            }
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
                statusEl.innerText = `✅ Found ${data.tabs.length} tab(s). Select columns below:`;

                renderTabMappings(data.tabs);
            } catch (err) {
                statusEl.className = "mt-3 text-sm font-medium text-red-600";
                statusEl.innerText = "❌ Failed to read Excel file.";
            }
        }

        function guessColumn(columns, keywords) {
            for (const col of columns) {
                const lower = col.toLowerCase();
                if (keywords.some(k => lower.includes(k))) return col;
            }
            return "";
        }

        function renderTabMappings(tabs) {
            const section = document.getElementById('mapping-section');
            const container = document.getElementById('tabs-container');
            section.classList.remove('hidden');

            container.innerHTML = tabs.map((tab) => {
                const barcodeGuess = guessColumn(tab.columns, ['barcode', 'code', 'cod', 'sku', 'art', 'ean', 'id']);
                const nameGuess = guessColumn(tab.columns, ['name', 'desc', 'product', 'article', 'item', 'nome']);
                const cxGuess = guessColumn(tab.columns, ['cx', 'caixa', 'box', 'carton']);
                const embGuess = guessColumn(tab.columns, ['emb', 'pack', 'pct', 'fardo']);
                const pcGuess = guessColumn(tab.columns, ['pc', 'kgm', 'kg', 'un', 'unit', 'price', 'preco', 'preço']);
                const bagGuess = guessColumn(tab.columns, ['bag', 'saco', 'sc']);

                const makeOptions = (selectedVal, placeholderText) => {
                    let opts = `<option value="">${placeholderText}</option>`;
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

                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                        <div>
                            <label class="block text-xs font-bold uppercase text-slate-700 mb-1">Article Code / Barcode *</label>
                            <select class="col-barcode w-full p-2.5 bg-white border-2 border-indigo-200 rounded-lg font-medium text-sm">
                                ${makeOptions(barcodeGuess, '-- Select Barcode Column --')}
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase text-slate-700 mb-1">Article Name *</label>
                            <select class="col-name w-full p-2.5 bg-white border-2 border-indigo-200 rounded-lg font-medium text-sm">
                                ${makeOptions(nameGuess, '-- Select Name Column --')}
                            </select>
                        </div>
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4 bg-white p-4 rounded-xl border border-slate-200">
                        <div>
                            <label class="block text-xs font-bold uppercase text-emerald-700 mb-1">Price</label>
                            <select class="col-price-cx w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-sm">
                                ${makeOptions(cxGuess, '-- None / Skip --')}
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase text-blue-700 mb-1">Price</label>
                            <select class="col-price-emb w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-sm">
                                ${makeOptions(embGuess, '-- None / Skip --')}
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase text-amber-700 mb-1">Price</label>
                            <select class="col-price-pc w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-sm">
                                ${makeOptions(pcGuess, '-- None / Skip --')}
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase text-purple-700 mb-1">Price</label>
                            <select class="col-price-bag w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-sm">
                                ${makeOptions(bagGuess, '-- None / Skip --')}
                            </select>
                        </div>
                    </div>

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

        async function confirmAndImport() {
            const cards = document.querySelectorAll('.tab-config-card');
            const mappings = [];

            for (const card of cards) {
                const enabled = card.querySelector('.tab-enable-checkbox').checked;
                if (!enabled) continue;

                const sheetName = card.getAttribute('data-tab-name');
                const barcodeCol = card.querySelector('.col-barcode').value;
                const nameCol = card.querySelector('.col-name').value;
                const priceCxCol = card.querySelector('.col-price-cx').value;
                const priceEmbCol = card.querySelector('.col-price-emb').value;
                const pricePcCol = card.querySelector('.col-price-pc').value;
                const priceBagCol = card.querySelector('.col-price-bag').value;

                if (!barcodeCol || !nameCol) {
                    alert(`Please select Article Code and Article Name columns for tab: "${sheetName}"`);
                    return;
                }

                mappings.push({
                    sheet_name: sheetName,
                    barcode_col: barcodeCol,
                    name_col: nameCol,
                    price_cx_col: priceCxCol,
                    price_emb_col: priceEmbCol,
                    price_pc_col: pricePcCol,
                    price_bag_col: priceBagCol
                });
            }

            if (mappings.length === 0) {
                alert("Please select at least one tab to import.");
                return;
            }

            const btn = document.getElementById('btn-confirm-import');
            btn.disabled = true;
            btn.innerText = "⏳ Merging New & Existing Prices...";

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
                btn.innerText = "✅ Merge & Update Prices";

                if (data.success) {
                    await syncDatabaseInBackground(false);
                    loadManagerProductTable(); // Refresh table
                    alert(
                        `✅ Merge & Offline Sync Complete!\n\n` +
                        `🆕 New Added: ${data.new_added}\n` +
                        `🔄 Updated: ${data.existing_updated}\n` +
                        `📦 Total in DB: ${data.total_products}`
                    );
                    document.getElementById('total-count-badge').innerText = `Total Articles in DB: ${data.total_products}`;
                    document.getElementById('mapping-section').classList.add('hidden');
                } else {
                    alert("❌ Error: " + data.error);
                }
            } catch (err) {
                btn.disabled = false;
                btn.innerText = "✅ Merge & Update Prices";
                alert("❌ Failed to merge data.");
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


@app.route("/manifest.json")
def manifest():
    return Response(MANIFEST_JSON, mimetype="application/json")


@app.route("/sw.js")
def service_worker():
    return Response(SERVICE_WORKER_JS, mimetype="application/javascript")


@app.route("/api/sync-all")
def api_sync_all():
    with sqlite3.connect(DB_NAME) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT barcode, name, price_cx, price_emb, price_pc_kgm, price_bag, updated_at FROM products"
        ).fetchall()
    items = [dict(r) for r in rows]
    return jsonify({"success": True, "count": len(items), "items": items})


@app.route("/api/update-single", methods=["POST"])
def update_single():
    data = request.get_json()
    barcode = data.get("barcode")
    if not barcode:
        return jsonify({"success": False, "error": "Missing barcode"})

    p_cx = data.get("price_cx", "").strip()
    p_emb = data.get("price_emb", "").strip()
    p_pc = data.get("price_pc_kgm", "").strip()
    p_bag = data.get("price_bag", "").strip()
    now_str = datetime.now().strftime("%d-%b-%Y %H:%M")

    try:
        with sqlite3.connect(DB_NAME) as conn:
            conn.execute(
                """
                UPDATE products SET
                    price_cx = ?,
                    price_emb = ?,
                    price_pc_kgm = ?,
                    price_bag = ?,
                    updated_at = ?
                WHERE barcode = ?
            """,
                (p_cx, p_emb, p_pc, p_bag, now_str, barcode),
            )
            conn.commit()
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/search")
def api_search():
    q = clean_barcode(request.args.get("q", ""))
    if not q:
        return jsonify({"found": False, "items": []})
    with sqlite3.connect(DB_NAME) as conn:
        conn.row_factory = sqlite3.Row
        resolved = conn.execute("SELECT article_code FROM ean_aliases WHERE ean_code = ?", (q,)).fetchone()
        article_code = resolved[0] if resolved else q
        rows = conn.execute(
            "SELECT barcode, name, price_cx, price_emb, price_pc_kgm, price_bag FROM products WHERE barcode = ?",
            (article_code,),
        ).fetchall()
        if not rows:
            like_q = f"%{q}%"
            rows = conn.execute(
                """SELECT DISTINCT p.barcode, p.name, p.price_cx, p.price_emb, p.price_pc_kgm, p.price_bag
                   FROM products p LEFT JOIN ean_aliases e ON e.article_code = p.barcode
                   WHERE p.name LIKE ? OR p.barcode LIKE ? OR e.ean_code LIKE ? LIMIT 15""",
                (like_q, like_q, like_q),
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
    new_added = 0
    existing_updated = 0

    try:
        with sqlite3.connect(DB_NAME) as conn:
            existing_barcodes = {
                row[0]
                for row in conn.execute("SELECT barcode FROM products").fetchall()
            }

            for m in mappings:
                sheet = m["sheet_name"]
                b_col = m["barcode_col"]
                n_col = m["name_col"]
                cx_col = m.get("price_cx_col", "")
                emb_col = m.get("price_emb_col", "")
                pc_col = m.get("price_pc_col", "")
                bag_col = m.get("price_bag_col", "")

                if ext == ".csv":
                    df = pd.read_csv(filepath)
                else:
                    df = pd.read_excel(filepath, sheet_name=sheet)

                df = df.dropna(how="all")

                records_to_upsert = []
                for _, row in df.iterrows():
                    uploaded_id = clean_barcode(row.get(b_col))
                    if not uploaded_id:
                        continue
                    alias_row = conn.execute(
                        "SELECT article_code FROM ean_aliases WHERE ean_code = ?", (uploaded_id,)
                    ).fetchone()
                    barcode = alias_row[0] if alias_row else uploaded_id
                    name = (
                        str(row.get(n_col, "")).strip()
                        if not pd.isna(row.get(n_col))
                        else ""
                    )

                    p_cx = format_price(row.get(cx_col)) if cx_col else ""
                    p_emb = format_price(row.get(emb_col)) if emb_col else ""
                    p_pc = format_price(row.get(pc_col)) if pc_col else ""
                    p_bag = format_price(row.get(bag_col)) if bag_col else ""

                    if not any([p_cx, p_emb, p_pc, p_bag]) and not name:
                        continue

                    if barcode in existing_barcodes:
                        existing_updated += 1
                    else:
                        new_added += 1
                        existing_barcodes.add(barcode)

                    records_to_upsert.append(
                        (barcode, name, p_cx, p_emb, p_pc, p_bag, now_str)
                    )

                conn.executemany(
                    """
                    INSERT INTO products (barcode, name, price_cx, price_emb, price_pc_kgm, price_bag, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(barcode) DO UPDATE SET
                        name = CASE WHEN excluded.name != '' THEN excluded.name ELSE products.name END,
                        price_cx = CASE WHEN excluded.price_cx != '' THEN excluded.price_cx ELSE products.price_cx END,
                        price_emb = CASE WHEN excluded.price_emb != '' THEN excluded.price_emb ELSE products.price_emb END,
                        price_pc_kgm = CASE WHEN excluded.price_pc_kgm != '' THEN excluded.price_pc_kgm ELSE products.price_pc_kgm END,
                        price_bag = CASE WHEN excluded.price_bag != '' THEN excluded.price_bag ELSE products.price_bag END,
                        updated_at = excluded.updated_at
                    """,
                    records_to_upsert,
                )

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
            "new_added": new_added,
            "existing_updated": existing_updated,
            "total_products": total_products,
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

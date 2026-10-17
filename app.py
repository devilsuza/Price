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
    """Cleans Excel float formatting (e.g., converts '89012345.0' to '89012345')."""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def format_price(val):
    """Formats price values cleanly from Excel cells."""
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


# --- SERVICE WORKER FOR OFFLINE PAGE CACHING ---
SERVICE_WORKER_JS = """
const CACHE_NAME = 'store-price-portal-v2';
const ASSETS_TO_CACHE = [
    '/',
    'https://cdn.tailwindcss.com',
    'https://unpkg.com/html5-qrcode'
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

    // Do not intercept API calls in Service Worker (handled by IndexedDB logic)
    if (url.pathname.startsWith('/api/')) return;

    // Network-first for HTML/Assets, fallback to Offline Cache if internet is off
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


# --- HTML TEMPLATE (OFFLINE-FIRST + LOGIN + SCANNER + 4-PRICE MANAGER PORTAL) ---
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
        <div class="max-w-5xl mx-auto px-4 py-3 flex flex-wrap justify-between items-center gap-2">
            <div class="flex items-center flex-wrap gap-2">
                <span class="text-xl font-bold tracking-wide">🏪 Store Price Portal</span>
                <span id="role-badge" class="text-xs bg-indigo-900 text-indigo-200 px-2.5 py-0.5 rounded-full font-semibold hidden"></span>
                
                <!-- Online / Offline Sync Status Badge -->
                <div class="flex items-center space-x-1 bg-indigo-800/90 px-3 py-1 rounded-full border border-indigo-500 text-xs font-semibold">
                    <span id="net-status-dot" class="w-2.5 h-2.5 rounded-full bg-emerald-400 inline-block"></span>
                    <span id="net-status-text">Checking...</span>
                    <button onclick="syncDatabaseInBackground(true)" id="btn-manual-sync" title="Sync Database Now" class="ml-1.5 text-indigo-200 hover:text-white">
                        🔄
                    </button>
                </div>
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

    <div class="max-w-4xl mx-auto p-4 mt-2">

        <!-- Offline Mode Banner (Shows automatically when Internet is Off) -->
        <div id="offline-banner" class="hidden mb-4 bg-amber-50 border-2 border-amber-400 text-amber-900 px-4 py-3 rounded-xl flex items-center justify-between text-sm font-semibold shadow-sm">
            <div>
                ⚡ <strong>Offline Mode Active:</strong> Internet band hai, par aap barcode scan kar sakte hain! Data local memory se aa raha hai.
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
                            <svg xmlns="http://www.w3.org/2000/svg" class="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M3 7V5a2 2 0 012-2h2m10 0h2a2 2 0 012 2v2m0 10v2a2 2 0 01-2 2h-2M7 21H5a2 2 0 01-2-2v-2M7 8h1v8H7V8zm3 0h2v8h-2V8zm4 0h1v8h-1V8zm3 0h1v8h-1V8z" />
                            </svg>
                        </button>
                    </div>

                    <button onclick="searchArticle()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold px-6 py-3 rounded-xl shadow">
                        Check Price
                    </button>
                </div>

                <!-- Mobile/Web Camera Container -->
                <div id="reader-container" class="hidden mt-4 bg-slate-900 p-3 rounded-2xl max-w-sm mx-auto text-center">
                    <div class="flex justify-between items-center text-white text-xs font-semibold mb-2 px-1">
                        <span>📷 Point camera at Barcode or QR Code</span>
                        <button onclick="stopCamera()" class="bg-red-500 hover:bg-red-600 text-white px-2.5 py-1 rounded-md">✕ Close</button>
                    </div>
                    <div id="reader" class="w-full overflow-hidden rounded-xl border-2 border-emerald-400 bg-black"></div>
                </div>
            </div>

            <!-- Price Display Result Card -->
            <div id="price-result" class="hidden space-y-4"></div>
        </div>

        <!-- ================= MANAGER UPLOAD VIEW ================= -->
        <div id="view-manager" class="hidden space-y-6">
            
            <!-- Step 1: Upload Excel -->
            <div class="bg-white p-6 rounded-2xl shadow-md border border-slate-200">
                <div class="flex flex-wrap justify-between items-center gap-2 mb-4">
                    <div>
                        <h2 class="text-xl font-bold text-slate-800">1. Upload Daily Price Excel (Auto-Merge)</h2>
                        <p class="text-sm text-slate-500">New articles will be added & existing article prices will be updated automatically</p>
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
                    <p class="text-sm text-slate-500">Map Barcode, Article Name, and the 4 Price columns (CX, EMB, PC/KGM, BAG) for each tab.</p>
                </div>

                <div id="tabs-container" class="space-y-6"></div>

                <div class="pt-4 border-t border-slate-200 flex justify-end">
                    <button onclick="confirmAndImport()" id="btn-confirm-import"
                        class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-8 py-3 rounded-xl shadow-lg text-base">
                        ✅ Merge & Update Prices
                    </button>
                </div>
            </div>

        </div>

    </div>

    <script>
        // --- PASSWORDS ---
        const EMPLOYEE_PASS = "123";
        const MANAGER_PASS = "2026";
        let currentRole = null;

        let currentTempFile = null;
        let html5QrCode = null;
        let cameraActive = false;

        // =========================================================================
        // 1. OFFLINE INDEXEDDB ENGINE & AUTOMATIC BACKGROUND SYNC
        // =========================================================================
        const IDB_NAME = "StorePriceOfflineDB";
        const IDB_VERSION = 1;
        const STORE_NAME = "products";
        let dbPromise = null;
        let localItemCount = 0;
        let isSyncing = false;

        // Register Service Worker so page works even if refreshed offline
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

        // Save full product list into IndexedDB
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

        // Search inside Local IndexedDB (Works 100% Offline)
        async function searchInLocalDB(query) {
            const db = await openLocalDB();
            const cleanQ = query.trim();
            const lowerQ = cleanQ.toLowerCase();

            return new Promise((resolve, reject) => {
                const tx = db.transaction(STORE_NAME, "readonly");
                const store = tx.objectStore(STORE_NAME);

                // 1. Check Exact Barcode Match First (Instant O(1) lookup)
                const exactReq = store.get(cleanQ);
                exactReq.onsuccess = () => {
                    if (exactReq.result) {
                        resolve([exactReq.result]);
                    } else {
                        // 2. Fallback: Search partial barcode or Article Name
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

        // Update Online / Offline UI Status Indicator
        function updateNetworkUI(syncingState = false) {
            const dot = document.getElementById('net-status-dot');
            const text = document.getElementById('net-status-text');
            const banner = document.getElementById('offline-banner');
            const offlineCountLbl = document.getElementById('offline-count-label');
            const lastSyncEl = document.getElementById('last-sync-time');

            const savedCount = localStorage.getItem("local_item_count") || localItemCount || 0;
            const savedTime = localStorage.getItem("last_sync_time") || "";

            if (savedTime) {
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
                offlineCountLbl.innerText = `${savedCount} Articles Ready`;
            }
        }

        // Automatic Background Sync Function
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
                console.log("Background sync skipped (server unreachable):", err);
            } finally {
                isSyncing = false;
                updateNetworkUI(false);
            }
        }

        // Listen for Internet Disconnect / Reconnect Events Automatically
        window.addEventListener('online', () => {
            updateNetworkUI();
            syncDatabaseInBackground(false); // Auto-sync immediately when internet returns!
        });

        window.addEventListener('offline', () => {
            updateNetworkUI();
        });

        // Initial load: open local DB & sync in background
        window.addEventListener('DOMContentLoaded', () => {
            updateNetworkUI();
            syncDatabaseInBackground(false);
            // Also auto-sync every 10 minutes in background when online
            setInterval(() => {
                if (navigator.onLine) syncDatabaseInBackground(false);
            }, 10 * 60 * 1000);
        });


        // =========================================================================
        // 2. SUPERMARKET BEEP SOUND & LOGIN SYSTEM
        // =========================================================================
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
            } catch (e) {
                console.log("Audio beep error:", e);
            }
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

        // =========================================================================
        // 3. HYBRID ONLINE + OFFLINE ARTICLE SEARCH & 4-PRICE DISPLAY
        // =========================================================================
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
                    <!-- Article Code & Article Name Header -->
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

                    <!-- 4 Price Types Grid -->
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

            // 1. IF OFFLINE -> Fetch Immediately from Local IndexedDB
            if (!navigator.onLine) {
                const localItems = await searchInLocalDB(query);
                displaySearchResults(localItems, query, "⚡ Offline Local DB");
                inputEl.select();
                return;
            }

            // 2. IF ONLINE -> Try Server first (with 2.5s timeout), fallback to Local DB if connection drops
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
                // Network failed or timed out -> Automatic fallback to Local IndexedDB!
                console.log("Server unreachable, using Offline Local DB fallback...");
                const localItems = await searchInLocalDB(query);
                displaySearchResults(localItems, query, "⚡ Offline Local DB");
                inputEl.select();
            }
        }

        // =========================================================================
        // 4. INLINE CAMERA BARCODE & QR SCANNER
        // =========================================================================
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
                        document.getElementById('scan-input').value = decodedText;
                        stopCamera();
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

        // =========================================================================
        // 5. MANAGER EXCEL UPLOAD & COLUMN MAPPING
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
                statusEl.innerText = `✅ Found ${data.tabs.length} tab(s). Select the columns below for each tab:`;

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

                    <!-- Article Code & Name Mapping -->
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

                    <!-- 4 Price Columns Mapping -->
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

                if (!priceCxCol && !priceEmbCol && !pricePcCol && !priceBagCol) {
                    alert(`Please select at least one Price column (CX, EMB, PC/KGM, or BAG) for tab: "${sheetName}"`);
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
                    // Immediately sync the newly updated database to local Offline IndexedDB!
                    await syncDatabaseInBackground(false);

                    alert(
                        `✅ Merge & Offline Sync Complete!\n\n` +
                        `🆕 New Articles Added: ${data.new_added}\n` +
                        `🔄 Existing Articles Updated: ${data.existing_updated}\n` +
                        `📦 Total Articles in Database: ${data.total_products}`
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


@app.route("/sw.js")
def service_worker():
    return Response(SERVICE_WORKER_JS, mimetype="application/javascript")


@app.route("/api/sync-all")
def api_sync_all():
    """Returns all products for browser Offline IndexedDB caching."""
    with sqlite3.connect(DB_NAME) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT barcode, name, price_cx, price_emb, price_pc_kgm, price_bag, updated_at FROM products"
        ).fetchall()
    items = [dict(r) for r in rows]
    return jsonify({"success": True, "count": len(items), "items": items})


@app.route("/api/search")
def api_search():
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"found": False, "items": []})

    with sqlite3.connect(DB_NAME) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT barcode, name, price_cx, price_emb, price_pc_kgm, price_bag FROM products WHERE barcode = ?",
            (q,),
        ).fetchall()

        if not rows:
            like_q = f"%{q}%"
            rows = conn.execute(
                """SELECT barcode, name, price_cx, price_emb, price_pc_kgm, price_bag
                   FROM products WHERE name LIKE ? OR barcode LIKE ? LIMIT 15""",
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
                    barcode = clean_barcode(row.get(b_col))
                    name = (
                        str(row.get(n_col, "")).strip()
                        if not pd.isna(row.get(n_col))
                        else ""
                    )

                    if not barcode:
                        continue

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

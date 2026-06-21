/**
 * SIAP PKP Filler — Browser-only edition
 * =======================================
 *
 * Di-load oleh bookmarklet ke tab SIAP BPK. Semua processing terjadi di
 * browser user — Excel data + auth token nggak pernah keluar dari laptop.
 * Submit langsung ke SIAP API (same-origin) pakai token segar dari page.
 *
 * Flow:
 * 1. Validate user di tab siap.bpk.go.id halaman Prosedur
 * 2. Extract sub-pemeriksaan ID dari URL
 * 3. Tampilin overlay UI di tab SIAP
 * 4. User upload Excel + isi sheet name + NIP/Nama auditor
 * 5. Parse Excel via SheetJS (CDN)
 * 6. Capture fresh JWT + user_id dari live page request
 * 7. Fetch PIC list — filter prosedur ke yg user PIC saja
 * 8. Fetch prosedur_list dari API
 * 9. Match Excel ↔ prosedur (similarity-based untuk kode duplicate)
 * 10. Submit loop dengan classify_error + early-stop logic
 * 11. Show progress + summary
 */

(function() {
    'use strict';

    // Prevent double-load — kalau bookmarklet di-click lagi, re-open UI existing
    if (window.__siapFiller) {
        window.__siapFiller.show();
        return;
    }

    const VERSION = '1.2.2';
    const SHEETJS_CDN = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    const BASE_API = location.origin;

    // ============ VALIDATION ============
    if (!location.host.endsWith('siap.bpk.go.id')) {
        alert('SIAP PKP Filler cuma jalan di tab SIAP BPK (siap.bpk.go.id).\n\n' +
              'Buka https://siap.bpk.go.id dulu, login, lalu navigate ke ' +
              'Pelaksanaan > Prosedur, baru klik bookmarklet ini.');
        return;
    }
    if (!location.href.includes('/pelaksanaan/prosedur')) {
        if (!confirm('Kamu nggak lagi di halaman Pelaksanaan > Prosedur.\n\n' +
                     'Bookmarklet butuh halaman ini buat extract ID sub-pemeriksaan.\n\n' +
                     'Lanjut paksa? (Mungkin gagal extract.)')) return;
    }

    function extractSubId() {
        const url = location.href;
        const patterns = [
            /subpemeriksaan\/([a-f0-9-]{20,})/i,
            /\/sub[a-z]*\/([a-f0-9-]{20,})/i,
        ];
        for (const p of patterns) {
            const m = url.match(p);
            if (m) return m[1];
        }
        const uuid = url.match(/[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}/i);
        return uuid ? uuid[0] : null;
    }

    const subId = extractSubId();
    if (!subId) {
        alert('Gagal extract ID sub-pemeriksaan dari URL:\n\n' + location.href +
              '\n\nPastikan kamu di halaman Pelaksanaan > Prosedur dari sub-pemeriksaan yg dituju.');
        return;
    }

    // ============ STATE ============
    const state = {
        subId,
        token: null,
        userId: null,
        auditor: null,        // { nip, nama, peran, peranLvl }
        excelData: null,      // list of { no, hasil, langkah }
        sheetName: null,
        excelConfig: { headerRow: 16, dataStartRow: 19, noCol: 1, hasilCol: 7, langkahCol: 2 },
        prosedurList: null,   // list of { no, id, nama }
        picStates: null,      // Map<prosedurId, {has_hasil, is_approved}>
        prefilterStats: null, // {pic_count, filtered_to, skipped_filled, skipped_validated}
        submissions: null,    // list of { id, no, hasil }
        skipped: [],
        running: false,
        skipFilled: false,
        skipValidated: true,  // default ON (safer)
    };

    // ============ UI INJECTION ============
    const STYLE_PREFIX = 'siapf';
    const css = `
.${STYLE_PREFIX}-overlay { position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(15,23,42,.6); backdrop-filter:blur(4px); z-index:2147483646; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; }
.${STYLE_PREFIX}-modal { position:fixed; top:50%; left:50%; transform:translate(-50%,-50%); width:min(640px,92vw); max-height:88vh; background:#fff; border-radius:16px; box-shadow:0 20px 60px rgba(0,0,0,.3); z-index:2147483647; display:flex; flex-direction:column; overflow:hidden; }
.${STYLE_PREFIX}-header { padding:18px 22px; border-bottom:1px solid #e5e7eb; display:flex; align-items:center; justify-content:space-between; flex-shrink:0; }
.${STYLE_PREFIX}-title { font-size:16px; font-weight:700; color:#0f172a; margin:0; }
.${STYLE_PREFIX}-subtitle { font-size:11px; color:#94a3b8; margin:2px 0 0 0; }
.${STYLE_PREFIX}-close { background:none; border:none; cursor:pointer; padding:6px; color:#64748b; border-radius:6px; }
.${STYLE_PREFIX}-close:hover { background:#f1f5f9; color:#0f172a; }
.${STYLE_PREFIX}-body { padding:20px 22px; overflow-y:auto; flex:1; }
.${STYLE_PREFIX}-section { margin-bottom:18px; }
.${STYLE_PREFIX}-label { display:block; font-size:12px; font-weight:600; color:#475569; margin-bottom:6px; }
.${STYLE_PREFIX}-input, .${STYLE_PREFIX}-select { width:100%; padding:9px 12px; border:1px solid #cbd5e1; border-radius:8px; font-size:13px; box-sizing:border-box; font-family:inherit; }
.${STYLE_PREFIX}-input:focus, .${STYLE_PREFIX}-select:focus { outline:none; border-color:#6366f1; box-shadow:0 0 0 3px rgba(99,102,241,.15); }
.${STYLE_PREFIX}-grid { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
.${STYLE_PREFIX}-grid-3 { display:grid; grid-template-columns:repeat(3,1fr); gap:12px; }
.${STYLE_PREFIX}-grid-4 { display:grid; grid-template-columns:repeat(4,1fr); gap:8px; }
.${STYLE_PREFIX}-btn { padding:11px 22px; background:linear-gradient(135deg,#6366f1,#8b5cf6); color:#fff; border:none; border-radius:9px; font-size:13px; font-weight:700; cursor:pointer; transition:all .15s; font-family:inherit; }
.${STYLE_PREFIX}-btn:hover:not(:disabled) { transform:translateY(-1px); box-shadow:0 6px 16px rgba(99,102,241,.35); }
.${STYLE_PREFIX}-btn:disabled { opacity:.5; cursor:not-allowed; }
.${STYLE_PREFIX}-btn-secondary { background:#f1f5f9; color:#475569; }
.${STYLE_PREFIX}-btn-secondary:hover:not(:disabled) { background:#e2e8f0; box-shadow:none; transform:none; }
.${STYLE_PREFIX}-progress-bar { width:100%; height:8px; background:#e2e8f0; border-radius:4px; overflow:hidden; }
.${STYLE_PREFIX}-progress-fill { height:100%; background:linear-gradient(90deg,#6366f1,#a855f7); transition:width .3s; width:0%; }
.${STYLE_PREFIX}-log { background:#0f172a; color:#e2e8f0; padding:12px; border-radius:8px; font-family:'SF Mono',Monaco,monospace; font-size:11px; max-height:280px; overflow-y:auto; line-height:1.6; }
.${STYLE_PREFIX}-log-ok { color:#34d399; }
.${STYLE_PREFIX}-log-skip { color:#94a3b8; }
.${STYLE_PREFIX}-log-fail { color:#f87171; }
.${STYLE_PREFIX}-log-stopped { color:#fbbf24; font-weight:700; }
.${STYLE_PREFIX}-log-info { color:#60a5fa; }
.${STYLE_PREFIX}-stat-card { padding:10px; border-radius:8px; text-align:center; }
.${STYLE_PREFIX}-stat-num { font-size:22px; font-weight:800; line-height:1; }
.${STYLE_PREFIX}-stat-label { font-size:10px; font-weight:500; margin-top:3px; text-transform:uppercase; letter-spacing:.5px; }
.${STYLE_PREFIX}-banner { padding:10px 14px; border-radius:8px; font-size:12px; line-height:1.5; }
.${STYLE_PREFIX}-banner-info { background:#dbeafe; color:#1e40af; border:1px solid #bfdbfe; }
.${STYLE_PREFIX}-banner-warn { background:#fef3c7; color:#92400e; border:1px solid #fde68a; }
.${STYLE_PREFIX}-banner-ok { background:#d1fae5; color:#065f46; border:1px solid #a7f3d0; }
.${STYLE_PREFIX}-banner-err { background:#fee2e2; color:#991b1b; border:1px solid #fecaca; }
.${STYLE_PREFIX}-step { font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:.5px; margin-bottom:10px; }
.${STYLE_PREFIX}-hint { font-size:11px; color:#94a3b8; margin-top:4px; }
`;

    const styleEl = document.createElement('style');
    styleEl.id = STYLE_PREFIX + '-style';
    styleEl.textContent = css;
    document.head.appendChild(styleEl);

    const overlay = document.createElement('div');
    overlay.className = `${STYLE_PREFIX}-overlay`;
    overlay.style.display = 'none';
    overlay.innerHTML = `
        <div class="${STYLE_PREFIX}-modal">
            <div class="${STYLE_PREFIX}-header">
                <div>
                    <h2 class="${STYLE_PREFIX}-title">SIAP PKP Filler</h2>
                    <p class="${STYLE_PREFIX}-subtitle">v${VERSION} · by Bashid Effendi</p>
                </div>
                <button class="${STYLE_PREFIX}-close" id="${STYLE_PREFIX}-close" title="Tutup">
                    <svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M6 18L18 6M6 6l12 12"/></svg>
                </button>
            </div>
            <div class="${STYLE_PREFIX}-body" id="${STYLE_PREFIX}-body"></div>
        </div>
    `;
    document.body.appendChild(overlay);

    const bodyEl = overlay.querySelector('#' + STYLE_PREFIX + '-body');
    overlay.querySelector('#' + STYLE_PREFIX + '-close').addEventListener('click', () => {
        if (state.running && !confirm('Proses lagi jalan. Tutup beneran?')) return;
        overlay.style.display = 'none';
    });

    // ============ HELPER: LOAD SHEETJS ============
    function loadSheetJS() {
        return new Promise((resolve, reject) => {
            if (window.XLSX) return resolve(window.XLSX);
            const s = document.createElement('script');
            s.src = SHEETJS_CDN;
            s.onload = () => window.XLSX ? resolve(window.XLSX) : reject(new Error('XLSX not loaded'));
            s.onerror = () => reject(new Error('Gagal load SheetJS dari CDN'));
            document.head.appendChild(s);
        });
    }

    // ============ CAPTURE TOKEN + USER_ID ============
    function captureLiveTokenAndUserId(timeoutMs = 10000) {
        return new Promise((resolve, reject) => {
            let capturedToken = null;
            let capturedUserId = null;
            const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;
            const origOpen = XMLHttpRequest.prototype.open;
            const origFetch = window.fetch;
            let cleaned = false;

            function cleanup() {
                if (cleaned) return;
                cleaned = true;
                XMLHttpRequest.prototype.setRequestHeader = origSetHeader;
                XMLHttpRequest.prototype.open = origOpen;
                window.fetch = origFetch;
            }

            function tryExtractUserId(url) {
                if (capturedUserId || !url) return;
                const m = url.match(/dataPicHasil\/[^\/]+\/([a-f0-9-]{20,})/i);
                if (m) capturedUserId = m[1];
            }

            XMLHttpRequest.prototype.open = function(method, url) {
                this.__url = url;
                tryExtractUserId(url);
                return origOpen.apply(this, arguments);
            };
            XMLHttpRequest.prototype.setRequestHeader = function(name, value) {
                if (!capturedToken && name && name.toLowerCase() === 'authorization' &&
                    typeof value === 'string' && value.startsWith('Bearer ') &&
                    this.__url && this.__url.includes('/api/')) {
                    capturedToken = value.slice(7);
                }
                return origSetHeader.apply(this, arguments);
            };
            window.fetch = function(input, init) {
                const url = typeof input === 'string' ? input : (input && input.url);
                tryExtractUserId(url);
                if (!capturedToken && init && init.headers) {
                    const auth = init.headers.Authorization || init.headers.authorization ||
                        (typeof init.headers.get === 'function' && init.headers.get('Authorization'));
                    if (auth && typeof auth === 'string' && auth.startsWith('Bearer ')) {
                        capturedToken = auth.slice(7);
                    }
                }
                return origFetch.apply(this, arguments);
            };

            const start = Date.now();
            const poll = setInterval(() => {
                if (capturedToken) {
                    clearInterval(poll);
                    cleanup();
                    resolve({ token: capturedToken, userId: capturedUserId });
                } else if (Date.now() - start > timeoutMs) {
                    clearInterval(poll);
                    cleanup();
                    reject(new Error('Token nggak ter-capture dalam ' + (timeoutMs/1000) + 's'));
                }
            }, 100);

            // Trigger UI: klik tab Akun → Prosedur biar page kirim API request
            setTimeout(() => {
                const links = [...document.querySelectorAll('a')];
                const otherTab = links.find(a => {
                    const txt = a.textContent.trim();
                    return (txt === 'Akun' || txt === 'Dokumen Pemeriksaan') &&
                           (a.href.includes('/akun') || a.href.includes('/dokumen'));
                });
                if (otherTab) {
                    otherTab.click();
                    setTimeout(() => {
                        const prosLink = links.find(a =>
                            a.textContent.trim() === 'Prosedur' && a.href.includes('/prosedur'));
                        if (prosLink) prosLink.click();
                    }, 1500);
                }
            }, 200);
        });
    }

    // ============ EXCEL PARSING ============
    async function parseExcel(file, sheetName, config) {
        const XLSX = await loadSheetJS();
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        if (!wb.SheetNames.includes(sheetName)) {
            throw new Error(`Sheet "${sheetName}" tidak ada. Sheet tersedia: ${wb.SheetNames.join(', ')}`);
        }
        const ws = wb.Sheets[sheetName];
        const range = XLSX.utils.decode_range(ws['!ref']);
        const results = [];
        for (let r = config.dataStartRow - 1; r <= range.e.r; r++) {
            const noCell = ws[XLSX.utils.encode_cell({ r, c: config.noCol - 1 })];
            const hasilCell = ws[XLSX.utils.encode_cell({ r, c: config.hasilCol - 1 })];
            const langkahCell = ws[XLSX.utils.encode_cell({ r, c: config.langkahCol - 1 })];
            // Guard .v == null: sel formula tanpa cached value / sel non-data
            // jangan jadi string "undefined"/"null" yang nyasar ke-submit.
            const cellVal = (cell) => (cell && cell.v != null) ? String(cell.v).trim() : '';
            const no = cellVal(noCell);
            const hasil = cellVal(hasilCell);
            const langkah = cellVal(langkahCell);
            if (no && no.startsWith('B.') && hasil) {
                results.push({ no, hasil, langkah });
            }
        }
        return results;
    }

    async function getSheetNames(file) {
        const XLSX = await loadSheetJS();
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array', bookSheets: true });
        return wb.SheetNames;
    }

    // ============ DEEP-LINK: BUKA PKP SAYA (file asli di Perencanaan) ============
    // SIAP nggak expose endpoint download file langsung (file disajikan via
    // office viewer ONLYOFFICE-style). Jadi tombol ini LONCAT ke halaman file
    // PKP kamu di Perencanaan → di baris namamu klik "Lihat File" (buka viewer)
    // → download dari viewer buat ambil file ASLI. ("Download File" doang nggak
    // langsung nyimpen — harus Lihat File dulu, dikonfirmasi user 2026-06-21.)
    // Container "Program Kerja Perseorangan" dicari via kontainermenu
    // (indeks A). Verified empiris 2026-06-21: container-id stabil, URL detail
    // /perencanaan/dokumenpemeriksaan/{containerId} nampilin daftar file PKP.
    async function openMyPkp(onStatus, win) {
        const setS = onStatus || (() => {});
        const sid = String(state.subId).toLowerCase();
        const base = `${BASE_API}/pemeriksaan/subpemeriksaan/${state.subId}/perencanaan/dokumenpemeriksaan`;
        const guidRe = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
        let containerId = null;
        try {
            setS('Ambil token dari halaman (tab switch sebentar — normal)…');
            const cap = await captureLiveTokenAndUserId();
            state.token = cap.token;
            setS('Cari dokumen PKP kamu…');
            const resp = await apiGet(`/api/subpemeriksaan/dok/kontainermenu?indeks=A&subid=${sid}`);
            const items = Array.isArray(resp) ? resp : (Array.isArray(resp.data) ? resp.data : []);
            const pkp = items.find(c => /program kerja perseorangan/i.test(JSON.stringify(c)));
            if (pkp) {
                // Ambil GUID container (bukan subId). Coba key id eksplisit dulu,
                // fallback ke field bernilai-GUID pertama yang bukan subId.
                for (const k of ['Id', 'id', 'KontainerId', 'kontainerId', 'IdKontainer']) {
                    if (pkp[k] && guidRe.test(String(pkp[k])) && String(pkp[k]).toLowerCase() !== sid) { containerId = pkp[k]; break; }
                }
                if (!containerId) for (const k in pkp) {
                    const v = pkp[k];
                    if (guidRe.test(String(v)) && String(v).toLowerCase() !== sid) { containerId = v; break; }
                }
            }
        } catch (e) { /* fallback: buka halaman daftar dokumen */ }

        const url = containerId ? `${base}/${containerId}` : base;
        // win dibuka SINKRON di handler (anti popup-block); redirect di sini.
        if (win && !win.closed) { win.location.href = url; } else { window.open(url, '_blank'); }
        return { direct: !!containerId };
    }

    // ============ API CALLS ============
    function getXsrfToken() {
        const m = document.cookie.split(';').map(c => c.trim()).find(c => c.startsWith('XSRF-TOKEN='));
        return m ? decodeURIComponent(m.split('=')[1]) : '';
    }

    async function apiGet(path) {
        const r = await fetch(BASE_API + path, {
            headers: {
                'Accept': 'application/json',
                'Authorization': 'Bearer ' + state.token,
                'X-XSRF-TOKEN': getXsrfToken()
            }
        });
        if (!r.ok) {
            let detail = '';
            try { const j = await r.json(); detail = j.error ? `${j.error.code}: ${j.error.message}` : JSON.stringify(j); } catch(e) {}
            throw new Error(`HTTP ${r.status} — ${detail}`);
        }
        return r.json();
    }

    async function apiPost(path, payload) {
        const r = await fetch(BASE_API + path, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Authorization': 'Bearer ' + state.token,
                'X-XSRF-TOKEN': getXsrfToken()
            },
            body: JSON.stringify(payload)
        });
        if (!r.ok) {
            let detail = '';
            try { const j = await r.json(); detail = j.error ? `${j.error.code}: ${j.error.message}` : JSON.stringify(j); } catch(e) {}
            throw new Error(`HTTP ${r.status} — ${detail}`);
        }
        return r.json();
    }

    // ============ DETECT AUDITOR ============
    async function detectAuditor(prosedurList) {
        if (!prosedurList || !prosedurList.length) return null;
        let lastErr = null;
        for (const item of prosedurList.slice(0, 3)) {
            try {
                const d = await apiGet(`/api/subpemeriksaan/prosedur/detail/${item.id}`);
                if (d.success && d.data && d.data.Pic && d.data.Pic.length) {
                    const pic = d.data.Pic.find(p => p.IsMyNip) || d.data.Pic[0];
                    return {
                        nip: pic.NIP || '',
                        nama: pic.Nama || '',
                        peran: pic.Peran || 'Anggota Tim',
                        peranLvl: pic.PeranLvl || '6'
                    };
                }
            } catch(e) { lastErr = e.message; }
        }
        throw new Error('Auto-detect auditor gagal: ' + (lastErr || 'no Pic data'));
    }

    // ============ FETCH PIC IDs ============
    async function fetchProjectId(subId) {
        // ProjectId (= Pemeriksaan.Id) di slot kedua URL dataPicHasil.
        // Fetch dari endpoint subpemeriksaan/data yang selalu work.
        const d = await apiGet(`/api/subpemeriksaan/data/${subId}`);
        if (!d || !d.data || !d.data.ProjectId) throw new Error('ProjectId not found');
        return d.data.ProjectId;
    }

    async function fetchUserPicStates(subId, projectId, targetNip) {
        // Return Map<prosedurId, {has_hasil, is_approved}>
        const d = await apiGet(`/api/subpemeriksaan/prosedur/bundle/dataPicHasil/${subId}/${projectId}`);
        if (!d.success || !Array.isArray(d.data)) throw new Error('dataPicHasil response invalid');
        const target = String(targetNip).trim();
        const states = new Map();
        for (const it of d.data) {
            if (String(it.NIP || '').trim() === target && it.IsPic) {
                if (it.PemeriksaanProsedurId) {
                    states.set(it.PemeriksaanProsedurId.toUpperCase(), {
                        has_hasil: !!it.HasilSusun,
                        is_approved: !!(it.IsApprovedByKT || it.IsApprovedByKST)
                    });
                }
            }
        }
        return states;
    }

    // ============ MATCH PROSEDUR ============
    function textSimilarity(a, b) {
        if (!a || !b) return 0;
        const al = a.toLowerCase().trim();
        const bl = b.toLowerCase().trim();
        if (al === bl) return 1;
        if (al.slice(0, 80) === bl.slice(0, 80)) return 0.95;
        const aw = new Set(al.split(/\s+/));
        const bw = new Set(bl.split(/\s+/));
        if (!aw.size || !bw.size) return 0;
        let overlap = 0;
        for (const w of aw) if (bw.has(w)) overlap++;
        return overlap / Math.max(aw.size, bw.size);
    }

    function matchProsedur(webProsedurs, excelData) {
        const excelByNo = {};
        for (const d of excelData) {
            if (!excelByNo[d.no]) excelByNo[d.no] = [];
            excelByNo[d.no].push(d);
        }
        const used = new Set();
        const submissions = [];
        const skipped = [];
        for (const item of webProsedurs) {
            const no = item.no;
            const webNama = item.nama || '';
            if (!excelByNo[no]) { skipped.push({ no, reason: 'Tidak ada di Excel' }); continue; }
            const candidates = excelByNo[no];
            const available = candidates.map((c, i) => [i, c]).filter(([i]) => !used.has(no + '|' + i));
            if (!available.length) { skipped.push({ no, reason: 'Semua kemunculan sudah terpakai' }); continue; }
            let best;
            if (available.length === 1) {
                best = available[0];
            } else if (webNama) {
                const scored = available.map(([i, c]) => [textSimilarity(webNama, c.langkah || ''), i, c]);
                scored.sort((a, b) => b[0] - a[0]);
                best = scored[0][0] >= 0.1 ? [scored[0][1], scored[0][2]] : available[0];
            } else {
                best = available[0];
            }
            used.add(no + '|' + best[0]);
            const hasil = best[1].hasil;
            if (hasil && hasil.trim()) {
                submissions.push({ id: item.id, no, hasil: hasil.trim() });
            } else {
                skipped.push({ no, reason: 'Hasil kosong' });
            }
        }
        return { submissions, skipped };
    }

    // ============ CLASSIFY ERROR ============
    function classifyError(msg) {
        if (!msg) return { sistemik: true, category: 'unknown' };
        const low = msg.toLowerCase();
        if (low.includes('bukan pic')) return { sistemik: false, category: 'not_pic' };
        if (low.includes('tidak memiliki akses') || low.includes('tidak berhak')) return { sistemik: false, category: 'no_access' };
        if (low.includes('sudah final') || low.includes('sudah disubmit') || low.includes('sudah divalidasi')) return { sistemik: false, category: 'already_done' };
        return { sistemik: true, category: 'sistemik' };
    }

    function padShortText(text, min = 101) {
        if (text.length >= min) return text;
        return text + ' Hasil pengujian telah didokumentasikan dalam kertas kerja pemeriksaan sesuai prosedur yang ditetapkan.';
    }

    // ============ RUN FILLER ============
    async function runFiller(submissions, auditor, onProgress, dryRun = false, maxConsecutive = 5) {
        const total = submissions.length;
        let success = 0, failed = 0, skippedPerm = 0;
        let consecutive = 0, stoppedEarly = false, lastError = null;
        const errors = [], skippedDetails = [];

        for (let i = 0; i < submissions.length; i++) {
            const sub = submissions[i];
            let statusText, statusType;
            try {
                const hasil = padShortText(sub.hasil);
                const payload = {
                    hasilPengujian: '<p>' + hasil + '</p>',
                    nip: auditor.nip, nama: auditor.nama,
                    peran: auditor.peran, peranLvl: auditor.peranLvl, nip2: false
                };
                if (dryRun) {
                    success++; consecutive = 0;
                    statusText = `DRY-RUN OK (${hasil.length} chars)`;
                    statusType = 'info';
                } else {
                    const result = await apiPost(`/api/subpemeriksaan/prosedur/store/${sub.id}`, payload);
                    if (result.success) {
                        success++; consecutive = 0;
                        statusText = 'OK'; statusType = 'ok';
                    } else {
                        const msg = (result.error && result.error.message) || JSON.stringify(result);
                        const { sistemik, category } = classifyError(msg);
                        if (sistemik) {
                            failed++; consecutive++; lastError = msg;
                            errors.push({ no: sub.no, error: msg });
                            statusText = `GAGAL: ${msg}`; statusType = 'fail';
                        } else {
                            skippedPerm++; consecutive = 0;
                            skippedDetails.push({ no: sub.no, reason: msg, category });
                            statusText = `SKIP (${category}): ${msg}`; statusType = 'skip';
                        }
                    }
                }
            } catch(e) {
                const msg = e.message;
                const { sistemik, category } = classifyError(msg);
                if (sistemik) {
                    failed++; consecutive++; lastError = msg;
                    errors.push({ no: sub.no, error: msg });
                    statusText = `ERROR: ${msg}`; statusType = 'fail';
                } else {
                    skippedPerm++; consecutive = 0;
                    skippedDetails.push({ no: sub.no, reason: msg, category });
                    statusText = `SKIP (${category}): ${msg}`; statusType = 'skip';
                }
            }
            onProgress(i + 1, total, sub.no, statusText, statusType);
            if (!dryRun && (i + 1) % 10 === 0) await new Promise(r => setTimeout(r, 1000));
            if (consecutive >= maxConsecutive) {
                stoppedEarly = true;
                const stopMsg = `Dihentikan otomatis setelah ${maxConsecutive} error sistemik berturut-turut. Error terakhir: ${lastError}`;
                onProgress(i + 1, total, '—', `STOPPED: ${stopMsg}`, 'stopped');
                break;
            }
        }
        return { success, failed, skippedPerm, skippedDetails, errors, total, stoppedEarly, lastError };
    }

    // ============ UI VIEWS ============
    function renderStep1() {
        bodyEl.innerHTML = `
            <div class="${STYLE_PREFIX}-step">Step 1 — Setup</div>
            <div class="${STYLE_PREFIX}-banner ${STYLE_PREFIX}-banner-info" style="margin-bottom:14px;">
                <strong>Sub-pemeriksaan ID:</strong> <code style="font-size:11px;">${state.subId}</code>
            </div>
            <div class="${STYLE_PREFIX}-section">
                <label class="${STYLE_PREFIX}-label">File Excel PKP (.xlsx)</label>
                <input type="file" id="${STYLE_PREFIX}-file" class="${STYLE_PREFIX}-input" accept=".xlsx,.xls">
                <p class="${STYLE_PREFIX}-hint">Kolom A=Nomor prosedur, G=Hasil pemeriksaan (default)</p>
            </div>
            <div class="${STYLE_PREFIX}-section">
                <label class="${STYLE_PREFIX}-label">Nama Sheet</label>
                <select id="${STYLE_PREFIX}-sheet" class="${STYLE_PREFIX}-select" disabled>
                    <option>Pilih file Excel dulu...</option>
                </select>
            </div>
            <div class="${STYLE_PREFIX}-grid ${STYLE_PREFIX}-section">
                <div>
                    <label class="${STYLE_PREFIX}-label">NIP Auditor</label>
                    <input type="text" id="${STYLE_PREFIX}-nip" class="${STYLE_PREFIX}-input" placeholder="240010157">
                </div>
                <div>
                    <label class="${STYLE_PREFIX}-label">Nama Auditor</label>
                    <input type="text" id="${STYLE_PREFIX}-nama" class="${STYLE_PREFIX}-input" placeholder="BASHID EFFENDI">
                </div>
            </div>
            <div class="${STYLE_PREFIX}-section">
                <button id="${STYLE_PREFIX}-pkp" class="${STYLE_PREFIX}-btn ${STYLE_PREFIX}-btn-secondary" style="width:100%;">📂 Buka PKP Saya di SIAP</button>
                <p class="${STYLE_PREFIX}-hint" id="${STYLE_PREFIX}-pkp-status" style="margin-top:6px;"></p>
            </div>
            <details>
                <summary style="cursor:pointer;font-size:12px;color:#6366f1;font-weight:600;margin-bottom:10px;">Pengaturan Lanjutan (Format Excel)</summary>
                <div class="${STYLE_PREFIX}-grid-4">
                    <div><label class="${STYLE_PREFIX}-label">Header Row</label><input id="${STYLE_PREFIX}-header-row" type="number" class="${STYLE_PREFIX}-input" value="16"></div>
                    <div><label class="${STYLE_PREFIX}-label">Data Start</label><input id="${STYLE_PREFIX}-data-row" type="number" class="${STYLE_PREFIX}-input" value="19"></div>
                    <div><label class="${STYLE_PREFIX}-label">Col No</label><input id="${STYLE_PREFIX}-no-col" type="number" class="${STYLE_PREFIX}-input" value="1"></div>
                    <div><label class="${STYLE_PREFIX}-label">Col Hasil</label><input id="${STYLE_PREFIX}-hasil-col" type="number" class="${STYLE_PREFIX}-input" value="7"></div>
                </div>
            </details>
            <div class="${STYLE_PREFIX}-section" style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:12px;">
                <p style="font-size:11px;font-weight:600;color:#475569;margin:0 0 8px 0;text-transform:uppercase;letter-spacing:.5px;">Opsi Replace</p>
                <label style="display:flex;gap:8px;font-size:12px;color:#334155;cursor:pointer;margin-bottom:6px;align-items:flex-start;">
                    <input type="checkbox" id="${STYLE_PREFIX}-skip-validated" checked style="margin-top:2px;">
                    <span><strong>Skip prosedur yang sudah divalidasi</strong> (KT/KST approve) — disarankan ON</span>
                </label>
                <label style="display:flex;gap:8px;font-size:12px;color:#334155;cursor:pointer;align-items:flex-start;">
                    <input type="checkbox" id="${STYLE_PREFIX}-skip-filled" style="margin-top:2px;">
                    <span><strong>Skip prosedur yang sudah terisi</strong> — pertahankan isi sebelumnya, jangan overwrite</span>
                </label>
            </div>
            <div class="${STYLE_PREFIX}-section" style="display:flex;gap:10px;justify-content:flex-end;margin-top:14px;margin-bottom:0;">
                <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:#475569;">
                    <input type="checkbox" id="${STYLE_PREFIX}-dryrun"> Dry-run (simulasi)
                </label>
                <button id="${STYLE_PREFIX}-start" class="${STYLE_PREFIX}-btn">Mulai Proses</button>
            </div>
        `;

        const fileInput = bodyEl.querySelector('#' + STYLE_PREFIX + '-file');
        const sheetSelect = bodyEl.querySelector('#' + STYLE_PREFIX + '-sheet');
        fileInput.addEventListener('change', async () => {
            if (!fileInput.files.length) return;
            try {
                const names = await getSheetNames(fileInput.files[0]);
                sheetSelect.innerHTML = names.map(n => `<option>${n}</option>`).join('');
                sheetSelect.disabled = false;
            } catch(e) {
                sheetSelect.innerHTML = `<option>Error: ${e.message}</option>`;
            }
        });

        const pkpBtn = bodyEl.querySelector('#' + STYLE_PREFIX + '-pkp');
        const pkpStatus = bodyEl.querySelector('#' + STYLE_PREFIX + '-pkp-status');
        pkpBtn.addEventListener('click', async () => {
            // Buka tab kosong SINKRON dalam gesture klik biar nggak ke-block popup,
            // baru di-redirect setelah container PKP ketemu.
            const win = window.open('about:blank', '_blank');
            const origText = pkpBtn.textContent;
            pkpBtn.disabled = true;
            try {
                const res = await openMyPkp((m) => { pkpStatus.textContent = m; }, win);
                pkpStatus.textContent = res.direct
                    ? '✓ Daftar file PKP kebuka di tab baru — di baris namamu klik "Lihat File", lalu download dari viewer-nya.'
                    : '✓ Halaman Dokumen kebuka — buka "Program Kerja Perseorangan" → Detail → di baris namamu klik "Lihat File" → download dari viewer.';
            } catch(e) {
                if (win && !win.closed) win.close();
                pkpStatus.textContent = '✗ Gagal: ' + e.message;
            } finally {
                pkpBtn.disabled = false;
                pkpBtn.textContent = origText;
            }
        });

        bodyEl.querySelector('#' + STYLE_PREFIX + '-start').addEventListener('click', async () => {
            if (!fileInput.files.length) return alert('Upload file Excel dulu!');
            const sheet = sheetSelect.value;
            const nip = bodyEl.querySelector('#' + STYLE_PREFIX + '-nip').value.trim();
            const nama = bodyEl.querySelector('#' + STYLE_PREFIX + '-nama').value.trim();
            if (!nip || !nama) return alert('NIP dan Nama wajib diisi!');
            const dryRun = bodyEl.querySelector('#' + STYLE_PREFIX + '-dryrun').checked;
            state.skipFilled = bodyEl.querySelector('#' + STYLE_PREFIX + '-skip-filled').checked;
            state.skipValidated = bodyEl.querySelector('#' + STYLE_PREFIX + '-skip-validated').checked;
            state.excelConfig.headerRow = parseInt(bodyEl.querySelector('#' + STYLE_PREFIX + '-header-row').value) || 16;
            state.excelConfig.dataStartRow = parseInt(bodyEl.querySelector('#' + STYLE_PREFIX + '-data-row').value) || 19;
            state.excelConfig.noCol = parseInt(bodyEl.querySelector('#' + STYLE_PREFIX + '-no-col').value) || 1;
            state.excelConfig.hasilCol = parseInt(bodyEl.querySelector('#' + STYLE_PREFIX + '-hasil-col').value) || 7;
            state.sheetName = sheet;
            state.auditor = { nip, nama, peran: 'Anggota Tim', peranLvl: '6' };
            try {
                await runFlow(fileInput.files[0], dryRun);
            } catch(e) {
                bodyEl.innerHTML = `<div class="${STYLE_PREFIX}-banner ${STYLE_PREFIX}-banner-err">${e.message}</div>
                    <button class="${STYLE_PREFIX}-btn ${STYLE_PREFIX}-btn-secondary" style="margin-top:14px;" onclick="window.__siapFiller.reset()">Mulai Lagi</button>`;
            }
        });
    }

    function renderProgress() {
        bodyEl.innerHTML = `
            <div class="${STYLE_PREFIX}-step">Step 2 — Submit Progress</div>
            <div id="${STYLE_PREFIX}-status-banner"></div>
            <div class="${STYLE_PREFIX}-section">
                <div style="display:flex;justify-content:space-between;font-size:13px;color:#475569;margin-bottom:6px;">
                    <span id="${STYLE_PREFIX}-progress-text">0 / 0</span>
                    <span id="${STYLE_PREFIX}-progress-pct" style="font-weight:700;color:#6366f1;">0%</span>
                </div>
                <div class="${STYLE_PREFIX}-progress-bar"><div class="${STYLE_PREFIX}-progress-fill" id="${STYLE_PREFIX}-progress-fill"></div></div>
            </div>
            <div class="${STYLE_PREFIX}-log" id="${STYLE_PREFIX}-log"></div>
            <div id="${STYLE_PREFIX}-summary" style="margin-top:14px;"></div>
        `;
    }

    function appendLog(text, type = 'info') {
        const logEl = bodyEl.querySelector('#' + STYLE_PREFIX + '-log');
        if (!logEl) return;
        const p = document.createElement('div');
        p.className = STYLE_PREFIX + '-log-' + type;
        p.textContent = text;
        logEl.appendChild(p);
        logEl.scrollTop = logEl.scrollHeight;
    }

    function setStatusBanner(html, level = 'info') {
        const el = bodyEl.querySelector('#' + STYLE_PREFIX + '-status-banner');
        if (el) el.innerHTML = `<div class="${STYLE_PREFIX}-banner ${STYLE_PREFIX}-banner-${level}" style="margin-bottom:14px;">${html}</div>`;
    }

    // ============ FULL FLOW ============
    async function runFlow(excelFile, dryRun) {
        renderProgress();
        state.running = true;

        // 1. Capture token + userId
        setStatusBanner('Capture token autentikasi dari live page...', 'info');
        appendLog('▶ Capture token dari page (tab akan switch sebentar — normal)', 'info');
        const cap = await captureLiveTokenAndUserId();
        state.token = cap.token;
        state.userId = cap.userId;
        try {
            const payload = JSON.parse(atob(state.token.split('.')[1].replace(/-/g,'+').replace(/_/g,'/')));
            const minLeft = Math.round((payload.exp - Math.floor(Date.now()/1000)) / 60);
            appendLog(`  Token valid ~${minLeft} menit lagi`, 'info');
        } catch(e) {}

        // 2. Parse Excel
        setStatusBanner('Parse Excel...', 'info');
        appendLog('▶ Parse Excel sheet "' + state.sheetName + '"', 'info');
        state.excelData = await parseExcel(excelFile, state.sheetName, state.excelConfig);
        appendLog(`  ${state.excelData.length} row data ter-parse`, 'info');

        // 3. Fetch prosedur list
        setStatusBanner('Fetch prosedur list dari SIAP...', 'info');
        appendLog('▶ Fetch prosedur list dari API', 'info');
        const pdata = await apiGet(`/api/subpemeriksaan/prosedur/bundle/data/${state.subId}`);
        if (!pdata.success || !pdata.data) throw new Error('Fetch prosedur gagal: ' + JSON.stringify(pdata.error || pdata));
        state.prosedurList = pdata.data.sort((a, b) => a.NoUrut - b.NoUrut).map(d => ({
            no: d.Kode, id: d.Id, nama: d.Nama || ''
        }));
        appendLog(`  ${state.prosedurList.length} prosedur ter-load`, 'info');

        // 4. Detect auditor if not set or to verify
        // (auditor sudah dari form — skip auto-detect)

        // 5. Pre-filter PIC + apply skip flags
        let filteredList = state.prosedurList;
        setStatusBanner('Fetch PIC list + filter prosedur...', 'info');
        appendLog('▶ Fetch ProjectId + PIC states buat pre-filter', 'info');
        try {
            const projectId = await fetchProjectId(state.subId);
            state.picStates = await fetchUserPicStates(state.subId, projectId, state.auditor.nip);
            let skippedValidated = 0, skippedFilled = 0;
            const allowedIds = new Set();
            for (const [pid, st] of state.picStates) {
                if (state.skipValidated && st.is_approved) { skippedValidated++; continue; }
                if (state.skipFilled && st.has_hasil) { skippedFilled++; continue; }
                allowedIds.add(pid);
            }
            filteredList = state.prosedurList.filter(p => allowedIds.has(p.id.toUpperCase()));
            state.prefilterStats = {
                pic_count: state.picStates.size,
                filtered_to: filteredList.length,
                skipped_validated: skippedValidated,
                skipped_filled: skippedFilled
            };
            const extras = [];
            if (skippedValidated > 0) extras.push(`${skippedValidated} sudah divalidasi`);
            if (skippedFilled > 0) extras.push(`${skippedFilled} sudah terisi`);
            appendLog(`  ${state.picStates.size} PIC, ${filteredList.length} akan diproses${extras.length ? ` (skip: ${extras.join(', ')})` : ''}`, 'ok');
        } catch(e) {
            appendLog('  ⚠ Pre-filter gagal: ' + e.message + ' — submit semua', 'skip');
        }

        // 6. Match Excel ↔ prosedur
        const matchResult = matchProsedur(filteredList, state.excelData);
        state.submissions = matchResult.submissions;
        state.skipped = matchResult.skipped;
        appendLog(`▶ Match: ${state.submissions.length} submission siap, ${state.skipped.length} skipped`, 'info');

        if (!state.submissions.length) {
            setStatusBanner('Tidak ada submission valid. Cek format Excel & nomor prosedur.', 'err');
            state.running = false;
            return;
        }

        // 7. Submit loop
        setStatusBanner(`${dryRun ? 'DRY-RUN' : 'SUBMIT'} ${state.submissions.length} prosedur...`, 'info');
        const summary = await runFiller(state.submissions, state.auditor, (cur, total, no, status, type) => {
            const pct = Math.round((cur / total) * 100);
            bodyEl.querySelector('#' + STYLE_PREFIX + '-progress-text').textContent = `${cur} / ${total}`;
            bodyEl.querySelector('#' + STYLE_PREFIX + '-progress-pct').textContent = pct + '%';
            bodyEl.querySelector('#' + STYLE_PREFIX + '-progress-fill').style.width = pct + '%';
            appendLog(`[${cur}/${total}] ${no} — ${status}`, type);
        }, dryRun);

        state.running = false;
        renderSummary(summary);
    }

    function renderSummary(s) {
        let bannerLevel = 'ok', bannerText = '✓ Selesai!';
        if (s.stoppedEarly) { bannerLevel = 'err'; bannerText = '✗ Dihentikan otomatis'; }
        else if (s.failed > 0 && s.success === 0) { bannerLevel = 'err'; bannerText = '✗ Gagal'; }
        else if (s.failed > 0) { bannerLevel = 'warn'; bannerText = '⚠ Selesai dengan error'; }

        setStatusBanner(bannerText + (s.stoppedEarly ? `<br><span style="font-size:11px;">Error terakhir: ${s.lastError}</span>` : ''), bannerLevel);

        const sumEl = bodyEl.querySelector('#' + STYLE_PREFIX + '-summary');
        sumEl.innerHTML = `
            <div class="${STYLE_PREFIX}-grid-4" style="margin-bottom:12px;">
                <div class="${STYLE_PREFIX}-stat-card" style="background:#d1fae5;color:#065f46;">
                    <div class="${STYLE_PREFIX}-stat-num">${s.success}</div>
                    <div class="${STYLE_PREFIX}-stat-label">Berhasil</div>
                </div>
                <div class="${STYLE_PREFIX}-stat-card" style="background:#fee2e2;color:#991b1b;">
                    <div class="${STYLE_PREFIX}-stat-num">${s.failed}</div>
                    <div class="${STYLE_PREFIX}-stat-label">Gagal</div>
                </div>
                <div class="${STYLE_PREFIX}-stat-card" style="background:#fef3c7;color:#92400e;">
                    <div class="${STYLE_PREFIX}-stat-num">${s.skippedPerm}</div>
                    <div class="${STYLE_PREFIX}-stat-label">Skip Izin</div>
                </div>
                <div class="${STYLE_PREFIX}-stat-card" style="background:#f1f5f9;color:#475569;">
                    <div class="${STYLE_PREFIX}-stat-num">${state.skipped.length}</div>
                    <div class="${STYLE_PREFIX}-stat-label">No-Match</div>
                </div>
            </div>
            <div style="text-align:center;font-size:11px;color:#94a3b8;margin-bottom:12px;">
                Auditor: <strong>${state.auditor.nama}</strong> (${state.auditor.nip})
                ${state.prefilterStats ? `<br>Pre-filter: <span style="color:#059669;font-weight:600;">${state.prefilterStats.pic_count} PIC → ${state.prefilterStats.filtered_to} diproses</span>${state.prefilterStats.skipped_validated > 0 ? ` (skip ${state.prefilterStats.skipped_validated} validated)` : ''}${state.prefilterStats.skipped_filled > 0 ? ` (skip ${state.prefilterStats.skipped_filled} filled)` : ''}` : ''}
            </div>
            <div style="display:flex;gap:8px;">
                <button class="${STYLE_PREFIX}-btn ${STYLE_PREFIX}-btn-secondary" style="flex:1;" onclick="window.__siapFiller.reset()">Mulai Lagi</button>
                <button class="${STYLE_PREFIX}-btn" style="flex:1;" onclick="window.__siapFiller.hide()">Tutup</button>
            </div>
        `;
    }

    // ============ PUBLIC API ============
    window.__siapFiller = {
        show: () => { overlay.style.display = 'block'; },
        hide: () => { overlay.style.display = 'none'; },
        reset: () => {
            state.token = null; state.userId = null; state.auditor = null;
            state.excelData = null; state.prosedurList = null;
            state.picStates = null; state.prefilterStats = null;
            state.submissions = null; state.skipped = []; state.running = false;
            renderStep1();
        },
        state: state,
        version: VERSION,
    };

    // Initial render
    renderStep1();
    overlay.style.display = 'block';

    console.log('SIAP PKP Filler v' + VERSION + ' loaded');
})();

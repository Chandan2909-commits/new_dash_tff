// ═══════════════════════════════════════════════════════════════════
//  LOGIN GATE
// ═══════════════════════════════════════════════════════════════════
const ADMIN_USER = 'admin';
const ADMIN_PASS = 'admin@123';

(function initLoginGate() {
    const gate = document.getElementById('login-gate');

    // Already logged in this session → hide gate immediately
    if (sessionStorage.getItem('pd_auth') === '1') {
        gate.style.display = 'none';
        return;
    }

    // Block the dashboard from being visible until auth passes
    gate.style.display = 'flex';

    // Password eye toggle
    document.getElementById('toggle-password').addEventListener('click', () => {
        const pwInput = document.getElementById('login-password');
        const icon = document.getElementById('pw-eye-icon');
        const isHidden = pwInput.type === 'password';
        pwInput.type = isHidden ? 'text' : 'password';
        icon.className = isHidden ? 'fa-solid fa-eye-slash' : 'fa-solid fa-eye';
    });

    // Form submit
    document.getElementById('login-form').addEventListener('submit', (e) => {
        e.preventDefault();
        const user = document.getElementById('login-username').value.trim();
        const pass = document.getElementById('login-password').value;
        const errEl = document.getElementById('login-error');

        if (user === ADMIN_USER && pass === ADMIN_PASS) {
            // ✅ Correct — animate out and unlock dashboard
            sessionStorage.setItem('pd_auth', '1');
            gate.classList.add('login-gate--fade-out');
            gate.addEventListener('animationend', () => {
                gate.style.display = 'none';
                gate.classList.remove('login-gate--fade-out');
            }, { once: true });
        } else {
            // ❌ Wrong — shake the card and show error
            errEl.style.display = 'flex';
            const card = gate.querySelector('.login-card');
            card.classList.remove('shake');
            void card.offsetWidth; // reflow to restart animation
            card.classList.add('shake');
        }
    });
})();

// ═══════════════════════════════════════════════════════════════════
//  CONFIGURATION
// ═══════════════════════════════════════════════════════════════════

// Main Google Sheet API (reads all leads)
const API_URL = 'https://script.google.com/macros/s/AKfycbwRoeDbrZYDc8sY_LTiKGFj3o0JD_oOjpCRKJmuO8bVRDtYo5pRIZVPCG5sdfKEvYXM/exec';

// Closed Queries Google Sheet API
// After deploying closed_queries_appscript.js as a Web App, paste the URL below:
const CLOSED_QUERIES_API_URL = 'https://script.google.com/macros/s/AKfycbyUN55bQ4tKBiTObgZJM16G0_0BJlR8yG64pJuBciavbqPqLmVKR78odvlR2x2m90XpVw/exec';


// ═══════════════════════════════════════════════════════════════════
//  STATE
// ═══════════════════════════════════════════════════════════════════
let sheetData = [];
let sheetHeaders = [];
let filteredData = [];
let closedQueries = new Set(); // "email_phone_date" exact keys
let closedEmailPhoneSet = new Set(); // "email_phone" fallback
let closedEmailSet = new Set(); // email-only fallback — most reliable
let closedQueriesDetails = {};       // exactKey → { email, phone, date, messages, closedOn }
let queriesResolved = 0;

// ── Pagination ────────────────────────────────────────────────────
const ROWS_PER_PAGE = 10;
let currentPage = 1;

// ── Status Filter (all | open | closed) ───────────────────────────
let statusFilter = 'all';

// ═══════════════════════════════════════════════════════════════════
//  DATE NORMALIZER — single source of truth for date formatting
//  Used in both syncClosedQueriesFromSheet() and renderTable()/getGroupedData()
//  so keys always match regardless of raw date string format.
// ═══════════════════════════════════════════════════════════════════
function normalizeDate(val) {
    if (!val) return '';
    const d = new Date(val);
    return isNaN(d) ? String(val).trim() : d.toLocaleDateString();
}

// ═══════════════════════════════════════════════════════════════════
//  CLOSED-ENTRY LOOKUP HELPER
//
//  TWO-LEVEL match (most specific → least specific):
//  Level 1: email + phone + date  (exact)
//  Level 2: email + phone         (immune to date-format / timezone drift)
// ═══════════════════════════════════════════════════════════════════
function isEntryClosed(email, phone, date) {
    if (!email) return false;
    const e = email.trim().toLowerCase();
    const exactKey = `${email}_${phone}_${date}`;
    const looseKey = `${email}_${phone}`;
    return (
        closedQueries.has(exactKey) ||        // L1: exact key
        closedEmailPhoneSet.has(looseKey) ||  // L2: email+phone
        closedEmailSet.has(e)                 // L3: email only (most reliable)
    );
}

// ═══════════════════════════════════════════════════════════════════
//  DOM ELEMENTS
// ═══════════════════════════════════════════════════════════════════
const tableBody = document.querySelector('#data-table tbody');
const refreshBtn = document.getElementById('refresh-btn');
const modal = document.getElementById('edit-modal');
const closeModalSpan = document.querySelector('.close-modal');
const searchInput = document.getElementById('search-input');
const queryModal = document.getElementById('query-modal');
const closeQueryModalSpan = document.querySelector('.close-query-modal');

// Chart instances
let pieChart, barChart, histogramChart;

// ═══════════════════════════════════════════════════════════════════
//  INITIAL LOAD
// ═══════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', async () => {
    // Load localStorage instantly so badge & stats are populated before the
    // network call finishes; fetchData() will re-sync from the Sheet and
    // re-render the table with closed entries correctly filtered.
    loadClosedQueriesFromStorage();
    updateClosedBadge();
    fetchData();   // ← syncs from closed-queries Sheet internally before rendering
    initNavigation();
    initCharts();

    document.getElementById('back-btn').addEventListener('click', () => {
        document.getElementById('page-loader').style.display = 'flex';
        setTimeout(() => {
            document.getElementById('query-section').style.display = 'none';
            document.getElementById('dashboard-section').style.display = 'block';
            document.getElementById('page-loader').style.display = 'none';
        }, 2000);
    });

    document.getElementById('close-all-btn').addEventListener('click', closeAllQueries);

    // ── Status filter pills ──────────────────────────────────────
    document.getElementById('status-filter-group').addEventListener('click', (e) => {
        const btn = e.target.closest('.filter-pill');
        if (!btn) return;
        // Update active pill
        document.querySelectorAll('.filter-pill').forEach(p => p.classList.remove('filter-pill--active'));
        btn.classList.add('filter-pill--active');
        // Apply filter and reset to page 1
        statusFilter = btn.dataset.filter;
        currentPage = 1;
        renderTable();
    });

    // ── User avatar dropdown ──────────────────────────────────────
    const userAvatar = document.getElementById('user-avatar');
    const userDropdown = document.getElementById('user-dropdown');
    const logoutBtn = document.getElementById('logout-btn');

    userAvatar.addEventListener('click', (e) => {
        e.stopPropagation();
        userDropdown.classList.toggle('show');
    });

    document.addEventListener('click', () => {
        userDropdown.classList.remove('show');
    });

    logoutBtn.addEventListener('click', () => {
        sessionStorage.removeItem('pd_auth');
        location.reload();
    });
});

// ═══════════════════════════════════════════════════════════════════
//  LOCAL STORAGE — persistence across page reloads
// ═══════════════════════════════════════════════════════════════════
function loadClosedQueriesFromStorage() {
    const savedKeys = localStorage.getItem('closedQueries');
    if (savedKeys) {
        closedQueries = new Set(JSON.parse(savedKeys));
        queriesResolved = closedQueries.size;
        // Rebuild fallback sets immediately so isEntryClosed() works
        // before the Google Sheet network call finishes.
        closedEmailPhoneSet = new Set();
        closedEmailSet = new Set();
        closedQueries.forEach(key => {
            const parts = key.split('_');
            // key format: "email_phone_date"
            if (parts.length >= 2) closedEmailPhoneSet.add(`${parts[0]}_${parts[1]}`);
        });
    }
    const savedDetails = localStorage.getItem('closedQueriesDetails');
    if (savedDetails) {
        closedQueriesDetails = JSON.parse(savedDetails);
    }
}

function saveClosedQueriesToStorage() {
    localStorage.setItem('closedQueries', JSON.stringify([...closedQueries]));
    localStorage.setItem('closedQueriesDetails', JSON.stringify(closedQueriesDetails));
}

// ═══════════════════════════════════════════════════════════════════
//  GOOGLE SHEET — push a closed query row
// ═══════════════════════════════════════════════════════════════════
async function postClosedQueryToSheet(detail) {
    if (CLOSED_QUERIES_API_URL === 'YOUR_CLOSED_QUERIES_WEB_APP_URL_HERE') {
        console.warn('CLOSED_QUERIES_API_URL not set — skipping Google Sheet push. Using localStorage only.');
        return;
    }
    try {
        const params = new URLSearchParams({
            action: 'add',
            date: detail.date || '',
            email: detail.email || '',
            phone: detail.phone || '',
            messages: Array.isArray(detail.messages) ? detail.messages.join(' | ') : (detail.messages || ''),
            closedOn: detail.closedOn || '',
            status: 'Closed'
        });
        await fetch(`${CLOSED_QUERIES_API_URL}?${params.toString()}`, {
            method: 'GET',
            mode: 'no-cors'
        });
        console.log('Closed query pushed to Google Sheet.');
    } catch (err) {
        console.error('Failed to push closed query to Google Sheet:', err);
    }
}

// ═══════════════════════════════════════════════════════════════════
//  GOOGLE SHEET — sync closed queries (SOURCE OF TRUTH)
//
//  This function is called on EVERY page load, refresh, and manual
//  Refresh button click.  It rebuilds closedQueries + closedEmailPhoneSet
//  from scratch using the closed-queries Google Sheet as the single source
//  of truth.  localStorage is used only as an instant pre-fill while the
//  network request is in-flight.
// ═══════════════════════════════════════════════════════════════════
async function syncClosedQueriesFromSheet() {
    const sheetRows = await fetchClosedQueriesFromSheet();

    if (sheetRows && sheetRows.length > 0) {
        // ── Full rebuild — wipe ALL stale state first ──
        closedQueries = new Set();
        closedEmailPhoneSet = new Set();
        closedEmailSet = new Set();
        closedQueriesDetails = {};

        sheetRows.forEach(row => {
            const email = (row['Email'] || '').trim();
            const rawPhone = (row['Phone'] || '').trim();
            const rawDate = (row['Date'] || '').trim();

            if (!email) return;

            const phoneIsValid = rawPhone && !rawPhone.startsWith('#') && rawPhone !== 'undefined';
            const phone = phoneIsValid ? rawPhone : '';
            const date = normalizeDate(rawDate);

            const exactKey = `${email}_${phone}_${date}`;
            const looseKey = `${email}_${phone}`;

            closedQueries.add(exactKey);
            closedEmailPhoneSet.add(looseKey);
            closedEmailSet.add(email.toLowerCase());   // L3: email-only match
            closedQueriesDetails[exactKey] = {
                date, email,
                phone: phoneIsValid ? rawPhone : '',
                messages: (row['Messages'] || '').split(' | ').filter(Boolean),
                closedOn: row['Closed On'] || '',
                status: 'Closed'
            };
        });

        queriesResolved = closedQueries.size;
        saveClosedQueriesToStorage();
        console.log(`[sync] ✔ ${closedQueries.size} closed quer(ies) loaded from Google Sheet.`);
    } else {
        // Sheet returned nothing — rebuild fallback sets from localStorage
        closedEmailPhoneSet = new Set();
        closedEmailSet = new Set();
        closedQueries.forEach(key => {
            const parts = key.split('_');
            if (parts.length >= 2) closedEmailPhoneSet.add(`${parts[0]}_${parts[1]}`);
        });
        console.warn('[sync] ⚠️ No data from sheet — using localStorage fallback.');
    }
}

async function fetchClosedQueriesFromSheet() {
    if (CLOSED_QUERIES_API_URL === 'YOUR_CLOSED_QUERIES_WEB_APP_URL_HERE') {
        return null;
    }
    try {
        // Add a cache-busting timestamp so browsers/CDNs never serve a stale response
        const res = await fetch(`${CLOSED_QUERIES_API_URL}?t=${Date.now()}`);
        const json = await res.json();
        if (json.success && Array.isArray(json.data)) {
            return json.data; // [{ Date, Email, Phone, Messages, 'Closed On', Status }]
        }
    } catch (err) {
        console.error('[sync] Failed to fetch closed queries from Google Sheet:', err);
    }
    return null;
}

// ═══════════════════════════════════════════════════════════════════
//  NAVIGATION
// ═══════════════════════════════════════════════════════════════════
function initNavigation() {
    document.querySelectorAll('.nav-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
            link.classList.add('active');

            const section = link.dataset.section;
            document.querySelectorAll('.content-section').forEach(s => s.style.display = 'none');

            if (section === 'analytics') {
                document.getElementById('analytics-section').style.display = 'block';
                updateCharts();
            } else if (section === 'dashboard') {
                document.getElementById('dashboard-section').style.display = 'block';
            } else if (section === 'closed-queries') {
                document.getElementById('closed-queries-section').style.display = 'block';
                renderClosedQueriesSection();
            }
        });
    });
}

// ═══════════════════════════════════════════════════════════════════
//  SEARCH
// ═══════════════════════════════════════════════════════════════════
searchInput.addEventListener('input', (e) => {
    const term = e.target.value.toLowerCase();
    filteredData = term
        ? sheetData.filter(row => {
            const email = (row.Email || '').toLowerCase();
            const phone = (row.Phone || '').toLowerCase();
            return email.includes(term) || phone.includes(term);
        })
        : sheetData;
    currentPage = 1;   // reset to first page on every new search
    renderTable();
});

// ═══════════════════════════════════════════════════════════════════
//  MODAL CLOSE HANDLERS
// ═══════════════════════════════════════════════════════════════════
refreshBtn.addEventListener('click', fetchData);

closeModalSpan.onclick = () => { modal.style.display = 'none'; };
closeQueryModalSpan.onclick = () => { queryModal.style.display = 'none'; };
window.onclick = (event) => {
    if (event.target === modal) modal.style.display = 'none';
    if (event.target === queryModal) queryModal.style.display = 'none';
};

// ═══════════════════════════════════════════════════════════════════
//  FETCH LEADS DATA
// ═══════════════════════════════════════════════════════════════════
async function fetchData() {
    if (API_URL === 'YOUR_WEB_APP_URL_HERE') {
        mockData();
        return;
    }
    try {
        refreshBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Loading...';

        // ── Fetch BOTH sheets in parallel for speed ─────────────────
        // syncClosedQueriesFromSheet() rebuilds the closed-set from the
        // Google Sheet (source of truth) BEFORE we render anything, so
        // every page open / refresh always shows only open entries.
        const [data] = await Promise.all([
            fetch(`${API_URL}?t=${Date.now()}`).then(r => r.json()),
            syncClosedQueriesFromSheet()
        ]);

        updateClosedBadge();

        sheetData = data;
        filteredData = data;

        if (data.length > 0) {
            sheetHeaders = ['Date', 'Email', 'Phone', 'Message'];
            renderTable();
            updateStats();
            updateCharts();
        }
    } catch (error) {
        console.error('Error fetching data:', error);
        alert('Failed to load data. Ensure "doGet" is implemented in your Apps Script.');
    } finally {
        refreshBtn.innerHTML = '<i class="fa-solid fa-rotate-right"></i> Refresh';
    }
}

// Mock Data (fallback)
function mockData() {
    sheetHeaders = ['Date', 'Email', 'Phone', 'Message'];
    sheetData = [
        { Date: '2023-10-27T10:00:00.000Z', Email: 'john@example.com', Phone: '1234567890', Messages: ['Interested in services', 'Need more info'] },
        { Date: '2023-10-26T14:30:00.000Z', Email: 'jane@test.com', Phone: '9876543210', Messages: ['Callback requested'] },
        { Date: '2023-10-27T09:15:00.000Z', Email: 'mike@demo.com', Phone: '5555555555', Messages: ['Pricing inquiry', 'Demo needed'] }
    ];
    filteredData = sheetData;
    renderTable();
    updateStats();
}

// ═══════════════════════════════════════════════════════════════════
//  RENDER MAIN TABLE
// ═══════════════════════════════════════════════════════════════════
function renderTable() {
    tableBody.innerHTML = '';
    const grouped = {};
    const dataToRender = (filteredData.length > 0 || searchInput.value) ? filteredData : sheetData;

    dataToRender.forEach(row => {
        const dateValue = row.Date || row.date || row.timestamp || row.Timestamp;
        const date = normalizeDate(dateValue);
        const email = row.Email || row.email || '';
        const phone = row.Phone || row.phone || '';
        const messages = row.Messages || row.messages || [];
        const key = `${email}_${phone}_${date}`;

        if (!grouped[key]) {
            grouped[key] = { date, email, phone, messages: [] };
        }
        if (Array.isArray(messages)) {
            grouped[key].messages.push(...messages);
        } else if (messages) {
            grouped[key].messages.push(messages);
        }
    });

    let sortedGroups = Object.values(grouped).sort((a, b) => new Date(b.date) - new Date(a.date));

    // ── Status filter — uses dual-key lookup (exact + loose email+phone) ──
    if (statusFilter === 'closed') {
        sortedGroups = sortedGroups.filter(g => isEntryClosed(g.email, g.phone, g.date));
    } else {
        // 'all' and 'open' — ALWAYS hide closed entries from main view
        sortedGroups = sortedGroups.filter(g => !isEntryClosed(g.email, g.phone, g.date));
    }

    // ── Pagination slice ─────────────────────────────────────────
    const totalPages = Math.max(1, Math.ceil(sortedGroups.length / ROWS_PER_PAGE));
    if (currentPage > totalPages) currentPage = totalPages;
    const startIdx = (currentPage - 1) * ROWS_PER_PAGE;
    const pageGroups = sortedGroups.slice(startIdx, startIdx + ROWS_PER_PAGE);

    pageGroups.forEach(group => {
        const rowKey = `${group.email}_${group.phone}_${group.date}`;
        // Use dual-key lookup: exact email+phone+date AND loose email+phone fallback
        const isClosed = isEntryClosed(group.email, group.phone, group.date);

        const tr = document.createElement('tr');

        // Date
        const tdDate = document.createElement('td');
        tdDate.textContent = group.date;
        tr.appendChild(tdDate);

        // Email
        const tdEmail = document.createElement('td');
        tdEmail.textContent = group.email;
        tr.appendChild(tdEmail);

        // Phone
        const tdPhone = document.createElement('td');
        tdPhone.textContent = group.phone;
        tr.appendChild(tdPhone);

        // Message — View Queries button
        const tdMsg = document.createElement('td');
        const queriesDropdown = document.createElement('div');
        queriesDropdown.className = 'queries-dropdown';

        const dropdownBtn = document.createElement('button');
        dropdownBtn.className = 'dropdown-btn';
        dropdownBtn.innerHTML = '<i class="fa-solid fa-messages"></i> View Queries';
        dropdownBtn.onclick = (e) => {
            e.stopPropagation();
            showAllQueries(group.messages, rowKey);
        };

        const dropdownMenu = document.createElement('div');
        dropdownMenu.className = 'dropdown-menu';
        queriesDropdown.appendChild(dropdownBtn);
        queriesDropdown.appendChild(dropdownMenu);
        tdMsg.appendChild(queriesDropdown);
        tr.appendChild(tdMsg);

        // Status
        const tdStatus = document.createElement('td');
        tdStatus.textContent = isClosed ? 'Chat Closed' : 'Open';
        tdStatus.style.color = isClosed ? '#10b981' : '#f59e0b';
        tdStatus.style.fontWeight = '600';
        tr.appendChild(tdStatus);

        // Action — always show button; grey when closed
        const tdAction = document.createElement('td');
        const closeBtn = document.createElement('button');
        if (isClosed) {
            closeBtn.className = 'btn-close-query btn-close-query--closed';
            closeBtn.innerHTML = '<i class="fa-solid fa-check-circle"></i> Query Closed';
            closeBtn.disabled = true;
        } else {
            closeBtn.className = 'btn-close-query';
            closeBtn.innerHTML = '<i class="fa-solid fa-times-circle"></i> Close Query';
            closeBtn.onclick = () => closeQuery(rowKey, group);
        }
        tdAction.appendChild(closeBtn);
        tr.appendChild(tdAction);

        tableBody.appendChild(tr);
    });

    // ── Render pagination bar ────────────────────────────────────
    renderPagination(totalPages, sortedGroups.length);

    // Collapse dropdowns on outside click
    document.addEventListener('click', () => {
        document.querySelectorAll('.dropdown-menu').forEach(m => m.style.display = 'none');
    });
}

// ═══════════════════════════════════════════════════════════════════
//  PAGINATION BAR
// ═══════════════════════════════════════════════════════════════════
function renderPagination(totalPages, totalEntries) {
    // Remove any existing pagination bar
    const existing = document.getElementById('pagination-bar');
    if (existing) existing.remove();

    if (totalPages <= 1) return;   // No need for pagination

    const bar = document.createElement('div');
    bar.id = 'pagination-bar';
    bar.className = 'pagination-bar';

    // Info label  e.g. "Showing 1–10 of 15 entries"
    const startEntry = (currentPage - 1) * ROWS_PER_PAGE + 1;
    const endEntry = Math.min(currentPage * ROWS_PER_PAGE, totalEntries);
    const info = document.createElement('span');
    info.className = 'pagination-info';
    info.textContent = `Showing ${startEntry}–${endEntry} of ${totalEntries} entries`;
    bar.appendChild(info);

    const controls = document.createElement('div');
    controls.className = 'pagination-controls';

    // Prev button
    const prevBtn = document.createElement('button');
    prevBtn.className = 'page-btn' + (currentPage === 1 ? ' page-btn--disabled' : '');
    prevBtn.innerHTML = '<i class="fa-solid fa-chevron-left"></i>';
    prevBtn.disabled = currentPage === 1;
    prevBtn.onclick = () => { currentPage--; renderTable(); };
    controls.appendChild(prevBtn);

    // Page number pills (show max 5 around current page)
    const range = buildPageRange(currentPage, totalPages);
    range.forEach(p => {
        if (p === '...') {
            const dots = document.createElement('span');
            dots.className = 'page-dots';
            dots.textContent = '…';
            controls.appendChild(dots);
        } else {
            const btn = document.createElement('button');
            btn.className = 'page-btn' + (p === currentPage ? ' page-btn--active' : '');
            btn.textContent = p;
            btn.onclick = () => { currentPage = p; renderTable(); };
            controls.appendChild(btn);
        }
    });

    // Next button
    const nextBtn = document.createElement('button');
    nextBtn.className = 'page-btn' + (currentPage === totalPages ? ' page-btn--disabled' : '');
    nextBtn.innerHTML = '<i class="fa-solid fa-chevron-right"></i>';
    nextBtn.disabled = currentPage === totalPages;
    nextBtn.onclick = () => { currentPage++; renderTable(); };
    controls.appendChild(nextBtn);

    bar.appendChild(controls);

    // Insert bar below the table-responsive div
    const tableSection = document.querySelector('#data-table').closest('.data-table-section');
    tableSection.appendChild(bar);
}

/** Returns an array like [1, '...', 4, 5, 6, '...', 12] */
function buildPageRange(current, total) {
    if (total <= 7) return Array.from({ length: total }, (_, i) => i + 1);
    const pages = [];
    pages.push(1);
    if (current > 3) pages.push('...');
    for (let p = Math.max(2, current - 1); p <= Math.min(total - 1, current + 1); p++) {
        pages.push(p);
    }
    if (current < total - 2) pages.push('...');
    pages.push(total);
    return pages;
}

// ═══════════════════════════════════════════════════════════════════
//  QUERY VIEW (full-page query list)
// ═══════════════════════════════════════════════════════════════════
let currentQueryRowKey = '';

function showAllQueries(messages, rowKey) {
    currentQueryRowKey = rowKey;
    document.getElementById('page-loader').style.display = 'flex';
    setTimeout(() => {
        const queryListContent = document.getElementById('query-list-content');
        queryListContent.innerHTML = '';

        for (let i = 0; i < Math.min(5, messages.length); i++) {
            const queryItem = document.createElement('div');
            queryItem.className = 'query-item';

            const queryText = document.createElement('div');
            queryText.className = 'query-text';
            queryText.innerHTML = `<strong>Query ${i + 1}:</strong> ${messages[i]}`;

            queryItem.appendChild(queryText);
            queryListContent.appendChild(queryItem);
        }

        document.getElementById('dashboard-section').style.display = 'none';
        document.getElementById('analytics-section').style.display = 'none';
        document.getElementById('query-section').style.display = 'block';
        document.getElementById('page-loader').style.display = 'none';
    }, 2000);
}

// ═══════════════════════════════════════════════════════════════════
//  CLOSE QUERY ACTIONS
// ═══════════════════════════════════════════════════════════════════

/** Called from "Close All Queries" button inside the query view page */
function closeAllQueries() {
    const grouped = getGroupedData();
    const entry = grouped[currentQueryRowKey];

    const detail = {
        date: entry ? entry.date : currentQueryRowKey.split('_')[2] || '',
        email: entry ? entry.email : currentQueryRowKey.split('_')[0] || '',
        phone: entry ? entry.phone : currentQueryRowKey.split('_')[1] || '',
        messages: entry ? entry.messages : [],
        closedOn: new Date().toLocaleString(),
        status: 'Closed'
    };

    closedQueriesDetails[currentQueryRowKey] = detail;
    closedQueries.add(currentQueryRowKey);
    queriesResolved++;

    saveClosedQueriesToStorage();
    postClosedQueryToSheet(detail);   // Push to Google Sheet
    updateStats();
    updateClosedBadge();

    document.getElementById('page-loader').style.display = 'flex';
    setTimeout(() => {
        document.getElementById('query-section').style.display = 'none';
        document.getElementById('dashboard-section').style.display = 'block';
        document.getElementById('page-loader').style.display = 'none';
        renderTable();
        updateCharts();
    }, 2000);
}

/** Called from "Close Query" button in the main table */
function closeQuery(rowKey, group) {
    const detail = {
        date: group.date,
        email: group.email,
        phone: group.phone,
        messages: group.messages || [],
        closedOn: new Date().toLocaleString(),
        status: 'Closed'
    };

    closedQueriesDetails[rowKey] = detail;
    closedQueries.add(rowKey);
    queriesResolved++;

    saveClosedQueriesToStorage();
    postClosedQueryToSheet(detail);   // Push to Google Sheet
    updateStats();
    updateClosedBadge();
    renderTable();
    updateCharts();
}

// ═══════════════════════════════════════════════════════════════════
//  CLOSED QUERIES SECTION — render in sidebar panel
// ═══════════════════════════════════════════════════════════════════
async function renderClosedQueriesSection() {
    const tbody = document.getElementById('closed-queries-tbody');
    tbody.innerHTML = `
        <tr>
            <td colspan="6" style="text-align:center;padding:30px;color:var(--text-secondary);">
                <i class="fa-solid fa-spinner fa-spin" style="font-size:1.5rem;"></i><br>Loading closed queries...
            </td>
        </tr>`;

    // Try fetching from Google Sheet first; fall back to localStorage
    const sheetRows = await fetchClosedQueriesFromSheet();

    tbody.innerHTML = '';

    if (sheetRows && sheetRows.length > 0) {
        // Render rows fetched from Google Sheet
        sheetRows.forEach(row => {
            renderClosedRow(tbody, {
                date: row['Date'] || '',
                email: row['Email'] || '',
                phone: row['Phone'] || '',
                messages: (row['Messages'] || '').split(' | ').filter(Boolean),
                closedOn: row['Closed On'] || '',
            });
        });
    } else if (closedQueries.size > 0) {
        // Fall back to localStorage data
        [...closedQueries].forEach(key => {
            const d = closedQueriesDetails[key] || {};
            renderClosedRow(tbody, {
                date: d.date || key.split('_')[2] || '',
                email: d.email || key.split('_')[0] || '',
                phone: d.phone || key.split('_')[1] || '',
                messages: d.messages || [],
                closedOn: d.closedOn || '',
            });
        });
    } else {
        // Empty state
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = 6;
        td.style.cssText = 'text-align:center;color:var(--text-secondary);padding:50px;';
        td.innerHTML = '<i class="fa-solid fa-inbox" style="font-size:2.5rem;display:block;margin-bottom:12px;opacity:0.5;"></i>No closed queries yet.';
        tr.appendChild(td);
        tbody.appendChild(tr);
    }
}

function renderClosedRow(tbody, { date, email, phone, messages, closedOn }) {
    const tr = document.createElement('tr');

    const cells = [
        date,
        email,
        phone,
        messages.length > 0
            ? `<span class="closed-msg-count">${messages.length} message${messages.length > 1 ? 's' : ''}</span>`
            : '-',
        closedOn,
        '<span class="status-closed-badge"><i class="fa-solid fa-circle-check"></i> Closed</span>'
    ];

    cells.forEach((content, i) => {
        const td = document.createElement('td');
        if (i === 3 || i === 5) {
            td.innerHTML = content;
        } else {
            td.textContent = content;
        }
        tr.appendChild(td);
    });

    tbody.appendChild(tr);
}

// ═══════════════════════════════════════════════════════════════════
//  BADGE & STATS HELPERS
// ═══════════════════════════════════════════════════════════════════
function updateClosedBadge() {
    const badge = document.getElementById('closed-count-badge');
    if (badge) badge.textContent = closedQueries.size;
}

function updateStats() {
    // Use grouped data so each unique lead (email+phone+date) counts as 1,
    // not once per raw sheet row / message.
    const grouped = getGroupedData();
    const totalLeads = Object.keys(grouped).length;
    const closedCount = closedQueries.size;

    // Count open queries using isEntryClosed() so loose-key fallback is included
    const openCount = Object.values(grouped).filter(g => !isEntryClosed(g.email, g.phone, g.date)).length;

    // Total unique leads
    document.getElementById('total-users').innerText = totalLeads;

    // Open queries = unique leads that are NOT closed
    document.getElementById('active-chats').innerText = openCount;

    // New Today = unique grouped leads whose date matches today
    const today = new Date().toLocaleDateString();
    const newTodayCount = Object.values(grouped).filter(g => {
        try { return new Date(g.date).toLocaleDateString() === today; }
        catch (e) { return false; }
    }).length;
    document.getElementById('new-today').innerText = newTodayCount;

    // Closed Queries
    document.getElementById('queries-resolved').innerText = closedCount;
}

function getGroupedData() {
    const grouped = {};
    sheetData.forEach(row => {
        const dateValue = row.Date || row.date || row.timestamp || row.Timestamp;
        const date = normalizeDate(dateValue);
        const email = row.Email || row.email || '';
        const phone = row.Phone || row.phone || '';
        const messages = row.Messages || row.messages || [];
        const key = `${email}_${phone}_${date}`;
        if (!grouped[key]) {
            grouped[key] = { email, phone, date, messages: [] };
        }
        if (Array.isArray(messages)) {
            grouped[key].messages.push(...messages);
        } else if (messages) {
            grouped[key].messages.push(messages);
        }
    });
    return grouped;
}

// ═══════════════════════════════════════════════════════════════════
//  EXCEL EXPORT (fallback / offline)
// ═══════════════════════════════════════════════════════════════════
async function exportClosedQueriesToExcel() {
    // Try to get fresh data from Google Sheet
    const sheetRows = await fetchClosedQueriesFromSheet();

    let rows = [['Date', 'Email', 'Phone', 'Messages', 'Closed On', 'Status']];

    if (sheetRows && sheetRows.length > 0) {
        sheetRows.forEach(r => {
            rows.push([r['Date'] || '', r['Email'] || '', r['Phone'] || '',
            r['Messages'] || '', r['Closed On'] || '', 'Closed']);
        });
    } else if (closedQueries.size > 0) {
        [...closedQueries].forEach(key => {
            const d = closedQueriesDetails[key] || {};
            rows.push([
                d.date || key.split('_')[2] || '',
                d.email || key.split('_')[0] || '',
                d.phone || key.split('_')[1] || '',
                (d.messages || []).join(' | '),
                d.closedOn || '',
                'Closed'
            ]);
        });
    } else {
        alert('No closed queries to export.');
        return;
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Closed Queries');
    XLSX.writeFile(wb, 'closed_queries.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
//  UPDATE DATA (DISABLED)
// ═══════════════════════════════════════════════════════════════════
async function updateData(updates) {
    alert('Editing is disabled because the current Google Apps Script only supports adding new rows.');
}

// ═══════════════════════════════════════════════════════════════════
//  CHARTS
// ═══════════════════════════════════════════════════════════════════
function initCharts() {
    const chartOptions = {
        responsive: true,
        maintainAspectRatio: true,
        plugins: { legend: { labels: { color: '#f8fafc' } } }
    };

    pieChart = new Chart(document.getElementById('pieChart'), {
        type: 'pie',
        data: {
            labels: ['Open Queries', 'Closed Queries'],
            datasets: [{
                data: [0, 0],
                backgroundColor: ['#f59e0b', '#10b981'],
                borderColor: '#1e293b',
                borderWidth: 2
            }]
        },
        options: chartOptions
    });

    barChart = new Chart(document.getElementById('barChart'), {
        type: 'bar',
        data: {
            labels: ['Open Queries', 'Closed Queries'],
            datasets: [{
                label: 'Count',
                data: [0, 0],
                backgroundColor: ['#f59e0b', '#10b981'],
                borderColor: ['#f59e0b', '#10b981'],
                borderWidth: 1
            }]
        },
        options: {
            ...chartOptions,
            scales: {
                y: { beginAtZero: true, ticks: { color: '#94a3b8' }, grid: { color: '#334155' } },
                x: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } }
            }
        }
    });

    const last7Days = [];
    for (let i = 6; i >= 0; i--) {
        const d = new Date();
        d.setDate(d.getDate() - i);
        last7Days.push(d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }));
    }

    histogramChart = new Chart(document.getElementById('histogramChart'), {
        type: 'bar',
        data: {
            labels: last7Days,
            datasets: [
                { label: 'Total Queries', data: [0, 0, 0, 0, 0, 0, 0], backgroundColor: '#3b82f6', borderColor: '#3b82f6', borderWidth: 1 },
                { label: 'Open Queries', data: [0, 0, 0, 0, 0, 0, 0], backgroundColor: '#f59e0b', borderColor: '#f59e0b', borderWidth: 1 },
                { label: 'Closed Queries', data: [0, 0, 0, 0, 0, 0, 0], backgroundColor: '#10b981', borderColor: '#10b981', borderWidth: 1 }
            ]
        },
        options: {
            ...chartOptions,
            scales: {
                y: { beginAtZero: true, ticks: { color: '#94a3b8', stepSize: 1 }, grid: { color: '#334155' } },
                x: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } }
            }
        }
    });
}

function updateCharts() {
    const totalQueries = Object.keys(getGroupedData()).length;
    const closedCount = closedQueries.size;
    const openCount = totalQueries - closedCount;

    pieChart.data.datasets[0].data = [openCount, closedCount];
    pieChart.update();

    barChart.data.datasets[0].data = [openCount, closedCount];
    barChart.update();

    const totalData = [0, 0, 0, 0, 0, 0, 0];
    const openData = [0, 0, 0, 0, 0, 0, 0];
    const closedData = [0, 0, 0, 0, 0, 0, 0];

    // Build a lookup of the last-7-days locale strings → array index
    const today = new Date();
    const dayLocale = [];
    for (let i = 0; i < 7; i++) {
        const d = new Date();
        d.setDate(today.getDate() - (6 - i));
        dayLocale.push(d.toLocaleDateString());
    }

    // ── Total & Open: bucket by the query's CREATION date ──────────
    const grouped = getGroupedData();
    Object.keys(grouped).forEach(key => {
        const dateStr = key.split('_')[2];
        try {
            const idx = dayLocale.indexOf(new Date(dateStr).toLocaleDateString());
            if (idx === -1) return;
            totalData[idx]++;
            if (!closedQueries.has(key)) {
                openData[idx]++;
            }
        } catch (e) { }
    });

    // ── Closed: bucket by the date the query was ACTUALLY CLOSED ───
    // This way a query created Feb 17 but closed Feb 20 shows on Feb 20
    Object.entries(closedQueriesDetails).forEach(([key, detail]) => {
        if (!detail || !detail.closedOn) return;
        try {
            // closedOn is stored as new Date().toLocaleString()
            const closedDate = new Date(detail.closedOn).toLocaleDateString();
            const idx = dayLocale.indexOf(closedDate);
            if (idx !== -1) {
                closedData[idx]++;
            }
        } catch (e) { }
    });

    histogramChart.data.datasets[0].data = totalData;
    histogramChart.data.datasets[1].data = openData;
    histogramChart.data.datasets[2].data = closedData;
    histogramChart.update();
}


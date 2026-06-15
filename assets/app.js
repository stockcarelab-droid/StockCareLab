const StockCare = (function () {
    const DB_KEY = "stockcarelab_database";
    const SESSION_KEY = "stockcarelab_session";
    const categoryLabels = {
        reagen: "Reagen",
        bmhp: "BMHP"
    };
    const fallbackDatabase = {
        users: [
            {
                id_user: 1,
                nama_lengkap: "Administrator Lab",
                username: "admin",
                password: "admin123",
                role: "admin",
                status: "aktif"
            },
            {
                id_user: 2,
                nama_lengkap: "CPNS Labkes 2026",
                username: "CPNSLabkes2026",
                password: "CPNSLabkes2026",
                role: "user",
                status: "aktif"
            }
        ],
        inventory_items: [],
        inventory_item_details: []
    };

    function qs(selector) {
        return document.querySelector(selector);
    }

    function qsa(selector, root = document) {
        return Array.from(root.querySelectorAll(selector));
    }

    function params() {
        return new URLSearchParams(window.location.search);
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
    }

    function getStoredDatabase() {
        const raw = localStorage.getItem(DB_KEY);
        if (!raw) {
            return null;
        }

        try {
            return JSON.parse(raw);
        } catch {
            return null;
        }
    }

    function saveDatabase(database) {
        localStorage.setItem(DB_KEY, JSON.stringify(database));
    }

    function syncDefaultUsers(database) {
        const existingUsers = Array.isArray(database.users) ? database.users : [];
        let changed = !Array.isArray(database.users);

        fallbackDatabase.users.forEach((defaultUser) => {
            const exists = existingUsers.some((user) => user.username === defaultUser.username);
            if (!exists) {
                existingUsers.push(structuredClone(defaultUser));
                changed = true;
            }
        });

        database.users = existingUsers;

        if (changed) {
            saveDatabase(database);
        }

        return database;
    }

    function normalizeSearch(value) {
        return String(value ?? "").trim().toLowerCase();
    }

    function includesSearch(row, query, fields) {
        if (!query) {
            return true;
        }

        return fields.some((field) => normalizeSearch(row[field]).includes(query));
    }

    function formatDuration(totalSeconds) {
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        return [hours, minutes, seconds].map((value) => String(value).padStart(2, "0")).join(":");
    }

    function initSystemUptime() {
        const uptime = qs("#systemUptime");
        if (!uptime) {
            return;
        }

        const activeSession = session();
        if (!activeSession) {
            return;
        }

        if (!Number(activeSession.login_started_at)) {
            activeSession.login_started_at = Date.now();
            sessionStorage.setItem(SESSION_KEY, JSON.stringify(activeSession));
        }

        const startedAt = Number(activeSession.login_started_at);
        function update() {
            uptime.textContent = formatDuration(Math.floor((Date.now() - startedAt) / 1000));
        }

        update();
        setInterval(update, 1000);
    }

    function downloadExcel(filename, sheetTitle, headers, rows) {
        const tableRows = [
            headers,
            ...rows
        ].map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("");
        const html = `
            <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
                <head><meta charset="UTF-8"></head>
                <body>
                    <table>
                        <tr><th colspan="${headers.length}">${escapeHtml(sheetTitle)}</th></tr>
                        ${tableRows}
                    </table>
                </body>
            </html>
        `;
        const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8;" });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        URL.revokeObjectURL(link.href);
        link.remove();
    }

    async function loadInitialDatabase() {
        const stored = getStoredDatabase();
        if (stored) {
            return syncDefaultUsers(stored);
        }

        try {
            const response = await fetch("assets/database.json", { cache: "no-store" });
            if (!response.ok) {
                throw new Error("Database JSON tidak tersedia.");
            }
            const database = await response.json();
            saveDatabase(database);
            return database;
        } catch {
            saveDatabase(fallbackDatabase);
            return structuredClone(fallbackDatabase);
        }
    }

    function database() {
        const stored = getStoredDatabase();
        if (stored) {
            return syncDefaultUsers(stored);
        }

        saveDatabase(fallbackDatabase);
        return structuredClone(fallbackDatabase);
    }

    function nextId(rows, field) {
        return rows.reduce((highest, row) => Math.max(highest, Number(row[field]) || 0), 0) + 1;
    }

    function session() {
        const raw = sessionStorage.getItem(SESSION_KEY) || localStorage.getItem(SESSION_KEY);
        if (!raw) {
            return null;
        }

        try {
            return JSON.parse(raw);
        } catch {
            return null;
        }
    }

    function requireLogin() {
        const activeSession = session();
        if (!activeSession) {
            window.location.href = "index.html";
            return null;
        }
        return activeSession;
    }

    function logout() {
        sessionStorage.removeItem(SESSION_KEY);
        localStorage.removeItem(SESSION_KEY);
        window.location.href = "index.html";
    }

    function bindLogout() {
        qsa("[data-logout]").forEach((button) => button.addEventListener("click", logout));
    }

    function openOverlay(overlay) {
        overlay.classList.add("show");
        overlay.setAttribute("aria-hidden", "false");
    }

    function closeOverlay(overlay) {
        overlay.classList.remove("show");
        overlay.setAttribute("aria-hidden", "true");
    }

    function updateStock(database, idItem) {
        const stock = database.inventory_item_details
            .filter((detail) => Number(detail.id_item) === Number(idItem))
            .reduce((total, detail) => total + Number(detail.jumlah_masuk) - Number(detail.jumlah_keluar), 0);
        const item = database.inventory_items.find((row) => Number(row.id_item) === Number(idItem));
        if (item) {
            item.stock_in_hand = stock;
        }
    }

    async function initLoginPage() {
        await loadInitialDatabase();
        if (session()) {
            window.location.href = "dashboard.html";
            return;
        }

        const form = qs("#loginForm");
        const alert = qs("#loginAlert");
        form.addEventListener("submit", function (event) {
            event.preventDefault();
            const db = database();
            const username = qs("#username").value.trim();
            const password = qs("#password").value;
            const user = db.users.find((row) => row.username === username && row.password === password && row.status === "aktif");

            if (!user) {
                alert.hidden = false;
                return;
            }

            const activeSession = {
                id_user: user.id_user,
                nama_lengkap: user.nama_lengkap,
                username: user.username,
                role: user.role,
                login_started_at: Date.now()
            };
            sessionStorage.setItem(SESSION_KEY, JSON.stringify(activeSession));
            window.location.href = "dashboard.html";
        });
    }

    function initDashboardPage() {
        const activeSession = requireLogin();
        if (!activeSession) {
            return;
        }

        bindLogout();
        initSystemUptime();
        qs("#userName").textContent = activeSession.nama_lengkap;

        let deleteId = null;
        const category = categoryLabels[params().get("kategori")] ? params().get("kategori") : "";
        const panel = qs("#stockPanel");
        const sheet = qs("#itemSheet");
        const deleteModal = qs("#deleteModal");

        if (!category) {
            return;
        }

        panel.hidden = false;
        qs("#categoryLabel").textContent = categoryLabels[category];
        qs("#sheetCategoryLabel").textContent = "Data " + categoryLabels[category];
        qs("#panelTitle").textContent = "Data " + categoryLabels[category];
        qs("#" + category + "Tab").classList.add("active");

        function dashboardRows() {
            const query = normalizeSearch(qs("#dashboardSearch").value);
            const db = database();
            return db.inventory_items
                .filter((item) => item.kategori === category)
                .filter((item) => includesSearch(item, query, ["item", "stock_in_hand", "expire_date", "distributor"]))
                .sort((a, b) => Number(b.id_item) - Number(a.id_item));
        }

        function renderItems() {
            const rows = dashboardRows();

            qs("#itemRows").innerHTML = rows.length ? rows.map((item, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td><a class="item-link" href="item_detail.html?id_item=${item.id_item}">${escapeHtml(item.item)}</a></td>
                    <td>${escapeHtml(item.stock_in_hand)}</td>
                    <td>${escapeHtml(item.expire_date)}</td>
                    <td>${escapeHtml(item.distributor)}</td>
                    <td>
                        <div class="table-actions">
                            <button type="button" class="cancel-link" data-edit-id="${item.id_item}">Edit</button>
                            <button type="button" class="delete-button" data-delete-id="${item.id_item}" data-delete-name="${escapeHtml(item.item)}">Hapus</button>
                        </div>
                    </td>
                </tr>
            `).join("") : `<tr><td colspan="6" class="empty-state">Data tidak ditemukan.</td></tr>`;

            qsa("[data-edit-id]").forEach((button) => button.addEventListener("click", () => openItemForm(Number(button.dataset.editId))));
            qsa("[data-delete-id]").forEach((button) => button.addEventListener("click", () => {
                deleteId = Number(button.dataset.deleteId);
                qs("#deleteMessage").textContent = "Yakin ingin menghapus " + button.dataset.deleteName + "?";
                openOverlay(deleteModal);
            }));
        }

        function exportDashboard() {
            const label = categoryLabels[category];
            const rows = dashboardRows().map((item, index) => [
                index + 1,
                item.item,
                item.stock_in_hand,
                item.expire_date,
                item.distributor
            ]);
            downloadExcel(
                `stockcare-${category}-dashboard.xls`,
                `Data ${label}`,
                ["Nomor", "Item", "Stock In Hand", "Expire Date", "Distributor"],
                rows
            );
        }

        function setAddOnlyVisible(visible) {
            qsa("[data-add-only]").forEach((element) => {
                element.style.display = visible ? "" : "none";
                qsa("input", element).forEach((input) => input.required = visible);
            });
        }

        function openItemForm(idItem = null) {
            const form = qs("#itemForm");
            form.reset();
            qs("#formAlert").hidden = true;
            qs("#editItemId").value = idItem || "";
            const isEdit = Boolean(idItem);
            qs("#sheetTitle").textContent = isEdit ? "Edit item" : "Tambah item";
            qs("#submitItemButton").textContent = isEdit ? "Simpan" : "Tambah";
            setAddOnlyVisible(!isEdit);

            if (isEdit) {
                const item = database().inventory_items.find((row) => Number(row.id_item) === Number(idItem));
                qs("#item").value = item.item;
                qs("#expire_date").value = item.expire_date;
                qs("#distributor").value = item.distributor;
            }

            openOverlay(sheet);
        }

        qs("#openFormButton").addEventListener("click", () => openItemForm());
        qs("#dashboardSearch").addEventListener("input", renderItems);
        qs("#exportDashboardButton").addEventListener("click", exportDashboard);
        qsa("[data-close-sheet]").forEach((button) => button.addEventListener("click", () => closeOverlay(sheet)));
        qs("#cancelDelete").addEventListener("click", () => closeOverlay(deleteModal));
        qs("#confirmDelete").addEventListener("click", () => {
            const db = database();
            db.inventory_items = db.inventory_items.filter((item) => Number(item.id_item) !== deleteId);
            db.inventory_item_details = db.inventory_item_details.filter((detail) => Number(detail.id_item) !== deleteId);
            saveDatabase(db);
            closeOverlay(deleteModal);
            renderItems();
        });

        qs("#itemForm").addEventListener("submit", function (event) {
            event.preventDefault();
            const db = database();
            const idItem = Number(qs("#editItemId").value);
            const payload = {
                item: qs("#item").value.trim(),
                expire_date: qs("#expire_date").value,
                distributor: qs("#distributor").value.trim()
            };

            if (idItem) {
                const item = db.inventory_items.find((row) => Number(row.id_item) === idItem);
                Object.assign(item, payload);
                db.inventory_item_details
                    .filter((detail) => Number(detail.id_item) === idItem)
                    .forEach((detail) => {
                        detail.expire_date = payload.expire_date;
                        detail.distributor = payload.distributor;
                    });
            } else {
                const newItem = {
                    id_item: nextId(db.inventory_items, "id_item"),
                    kategori: category,
                    stock_in_hand: 0,
                    ...payload
                };
                db.inventory_items.push(newItem);
                db.inventory_item_details.push({
                    id_detail: nextId(db.inventory_item_details, "id_detail"),
                    id_item: newItem.id_item,
                    tanggal: qs("#tanggal").value,
                    jumlah_masuk: Number(qs("#jumlah_masuk").value) || 0,
                    nomor_batch: qs("#nomor_batch").value.trim(),
                    expire_date: payload.expire_date,
                    satuan: qs("#satuan").value.trim(),
                    jumlah_keluar: Number(qs("#jumlah_keluar").value) || 0,
                    distributor: payload.distributor
                });
                updateStock(db, newItem.id_item);
            }

            saveDatabase(db);
            closeOverlay(sheet);
            renderItems();
        });

        renderItems();
    }

    function initDetailPage() {
        const activeSession = requireLogin();
        if (!activeSession) {
            return;
        }

        bindLogout();
        const idItem = Number(params().get("id_item"));
        let db = database();
        let item = db.inventory_items.find((row) => Number(row.id_item) === idItem);

        if (!item) {
            window.location.href = "dashboard.html";
            return;
        }

        const sheet = qs("#detailSheet");
        const deleteModal = qs("#deleteDetailModal");
        let deleteId = null;

        function refreshItem() {
            db = database();
            item = db.inventory_items.find((row) => Number(row.id_item) === idItem);
            qs("#categoryLabel").textContent = categoryLabels[item.kategori] || item.kategori;
            qs("#itemTitle").textContent = item.item;
            qs("#sheetItemName").textContent = item.item;
            qs("#stockInHand").textContent = item.stock_in_hand;
            qs("#backToCategory").href = "dashboard.html?kategori=" + encodeURIComponent(item.kategori);
            qs("#backToCategory").textContent = "Kembali ke " + (categoryLabels[item.kategori] || item.kategori);
            document.title = item.item + " | StockCare Lab";
        }

        function detailRowsWithStock() {
            let runningStock = 0;
            const query = normalizeSearch(qs("#detailSearch").value);
            const rows = database().inventory_item_details
                .filter((detail) => Number(detail.id_item) === idItem)
                .sort((a, b) => String(a.tanggal).localeCompare(String(b.tanggal)) || Number(a.id_detail) - Number(b.id_detail))
                .map((detail) => {
                    runningStock += Number(detail.jumlah_masuk) - Number(detail.jumlah_keluar);
                    return { ...detail, jumlah_stok: runningStock };
                })
                .filter((detail) => includesSearch(detail, query, [
                    "tanggal",
                    "jumlah_masuk",
                    "nomor_batch",
                    "expire_date",
                    "satuan",
                    "jumlah_keluar",
                    "jumlah_stok",
                    "distributor"
                ]));
            return qs("#urutan").value === "terbaru" ? rows.reverse() : rows;
        }

        function renderDetails() {
            refreshItem();
            const rows = detailRowsWithStock();
            qs("#detailRows").innerHTML = rows.length ? rows.map((detail) => `
                <tr>
                    <td>${escapeHtml(detail.tanggal)}</td>
                    <td>${escapeHtml(detail.jumlah_masuk)}</td>
                    <td>${escapeHtml(detail.nomor_batch)}</td>
                    <td>${escapeHtml(detail.expire_date)}</td>
                    <td>${escapeHtml(detail.satuan)}</td>
                    <td>${escapeHtml(detail.jumlah_keluar)}</td>
                    <td>${escapeHtml(detail.jumlah_stok)}</td>
                    <td>${escapeHtml(detail.distributor)}</td>
                    <td>
                        <div class="table-actions">
                            <button type="button" class="cancel-link" data-edit-detail="${detail.id_detail}">Edit</button>
                            <button type="button" class="delete-button" data-delete-detail="${detail.id_detail}" data-delete-name="${escapeHtml(detail.tanggal + " - " + detail.nomor_batch)}">Hapus</button>
                        </div>
                    </td>
                </tr>
            `).join("") : `<tr><td colspan="9" class="empty-state">Belum ada detail untuk item ini.</td></tr>`;

            qsa("[data-edit-detail]").forEach((button) => button.addEventListener("click", () => openDetailForm(Number(button.dataset.editDetail))));
            qsa("[data-delete-detail]").forEach((button) => button.addEventListener("click", () => {
                deleteId = Number(button.dataset.deleteDetail);
                qs("#deleteDetailMessage").textContent = "Yakin ingin menghapus detail " + button.dataset.deleteName + "? Stock In Hand akan dihitung ulang.";
                openOverlay(deleteModal);
            }));
        }

        function openDetailForm(idDetail = null) {
            qs("#detailForm").reset();
            qs("#detailAlert").hidden = true;
            qs("#editDetailId").value = idDetail || "";
            qs("#detailSheetTitle").textContent = idDetail ? "Edit detail item" : "Tambah update item";
            qs("#submitDetailButton").textContent = idDetail ? "Simpan" : "Tambah";

            if (idDetail) {
                const detail = database().inventory_item_details.find((row) => Number(row.id_detail) === Number(idDetail));
                qs("#tanggal").value = detail.tanggal;
                qs("#jumlah_masuk").value = detail.jumlah_masuk;
                qs("#nomor_batch").value = detail.nomor_batch;
                qs("#expire_date").value = detail.expire_date;
                qs("#satuan").value = detail.satuan;
                qs("#jumlah_keluar").value = detail.jumlah_keluar;
                qs("#distributor").value = detail.distributor;
            } else {
                qs("#expire_date").value = item.expire_date;
                qs("#distributor").value = item.distributor;
                qs("#jumlah_masuk").value = 0;
                qs("#jumlah_keluar").value = 0;
            }

            openOverlay(sheet);
        }

        function exportDetails() {
            const rows = detailRowsWithStock().map((detail) => [
                detail.tanggal,
                detail.jumlah_masuk,
                detail.nomor_batch,
                detail.expire_date,
                detail.satuan,
                detail.jumlah_keluar,
                detail.jumlah_stok,
                detail.distributor
            ]);
            downloadExcel(
                `stockcare-${item.kategori}-${item.item}.xls`.replace(/[\\/:*?"<>|]+/g, "-"),
                `Detail ${item.item}`,
                ["Tanggal", "Jumlah Masuk", "Nomor Batch", "Expire Date", "Satuan", "Jumlah Keluar", "Jumlah Stok", "Distributor"],
                rows
            );
        }

        qs("#urutan").addEventListener("change", renderDetails);
        qs("#detailSearch").addEventListener("input", renderDetails);
        qs("#exportDetailButton").addEventListener("click", exportDetails);
        qs("#openDetailFormButton").addEventListener("click", () => openDetailForm());
        qsa("[data-close-sheet]").forEach((button) => button.addEventListener("click", () => closeOverlay(sheet)));
        qs("#cancelDeleteDetail").addEventListener("click", () => closeOverlay(deleteModal));
        qs("#confirmDeleteDetail").addEventListener("click", () => {
            const currentDb = database();
            currentDb.inventory_item_details = currentDb.inventory_item_details.filter((detail) => Number(detail.id_detail) !== deleteId);
            updateStock(currentDb, idItem);
            saveDatabase(currentDb);
            closeOverlay(deleteModal);
            renderDetails();
        });

        qs("#detailForm").addEventListener("submit", function (event) {
            event.preventDefault();
            const currentDb = database();
            const idDetail = Number(qs("#editDetailId").value);
            const payload = {
                id_item: idItem,
                tanggal: qs("#tanggal").value,
                jumlah_masuk: Number(qs("#jumlah_masuk").value) || 0,
                nomor_batch: qs("#nomor_batch").value.trim(),
                expire_date: qs("#expire_date").value,
                satuan: qs("#satuan").value.trim(),
                jumlah_keluar: Number(qs("#jumlah_keluar").value) || 0,
                distributor: qs("#distributor").value.trim()
            };

            if (idDetail) {
                const detail = currentDb.inventory_item_details.find((row) => Number(row.id_detail) === idDetail);
                Object.assign(detail, payload);
            } else {
                currentDb.inventory_item_details.push({
                    id_detail: nextId(currentDb.inventory_item_details, "id_detail"),
                    ...payload
                });
            }

            updateStock(currentDb, idItem);
            saveDatabase(currentDb);
            closeOverlay(sheet);
            renderDetails();
        });

        renderDetails();
    }

    return {
        initLoginPage,
        initDashboardPage,
        initDetailPage
    };
})();

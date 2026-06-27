const StockCare = (function () {
    const SESSION_KEY = "stockcarelab_session";
    const SUPABASE_URL = "https://fsdxdcvcvvxlnhaahxep.supabase.co";
    const SUPABASE_KEY = "sb_publishable_g7wfUvGnf1A19Z0LVIXcdw_dFfIFCi8";
    const categoryLabels = {
        reagen: "Reagen",
        bmhp: "BMHP"
    };

    let supabaseClient = null;

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

    function normalizeSearch(value) {
        return String(value ?? "").trim().toLowerCase();
    }

    function includesSearch(row, query, fields) {
        if (!query) {
            return true;
        }

        return fields.some((field) => normalizeSearch(row[field]).includes(query));
    }

    function client() {
        if (!window.supabase) {
            throw new Error("Supabase SDK belum dimuat.");
        }

        if (!supabaseClient) {
            supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
        }

        return supabaseClient;
    }

    function showAlert(selector, message) {
        const alert = qs(selector);
        if (!alert) {
            return;
        }
        alert.textContent = message;
        alert.hidden = false;
    }

    function handleError(error, fallbackMessage = "Data gagal diproses. Coba lagi.") {
        console.error(error);
        return error?.message || fallbackMessage;
    }

    async function fetchTable(table, orderBy = null) {
        let query = client().from(table).select("*");
        if (orderBy) {
            query = query.order(orderBy.column, { ascending: orderBy.ascending });
        }
        const { data, error } = await query;
        if (error) {
            throw error;
        }
        return data || [];
    }

    async function database() {
        const [users, inventoryItems, inventoryItemDetails] = await Promise.all([
            fetchTable("users", { column: "id_user", ascending: true }),
            fetchTable("inventory_items", { column: "id_item", ascending: true }),
            fetchTable("inventory_item_details", { column: "id_detail", ascending: true })
        ]);

        return {
            users,
            inventory_items: inventoryItems,
            inventory_item_details: inventoryItemDetails
        };
    }

    async function getItem(idItem) {
        const { data, error } = await client()
            .from("inventory_items")
            .select("*")
            .eq("id_item", idItem)
            .single();

        if (error) {
            throw error;
        }

        return data;
    }

    async function updateItemStock(idItem) {
        const { error } = await client().rpc("refresh_item_stock", { item_id: idItem });
        if (error) {
            throw error;
        }
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

    async function initLoginPage() {
        if (session()) {
            window.location.href = "dashboard.html";
            return;
        }

        const form = qs("#loginForm");
        const alert = qs("#loginAlert");
        form.addEventListener("submit", async function (event) {
            event.preventDefault();
            alert.hidden = true;

            try {
                const username = qs("#username").value.trim();
                const password = qs("#password").value;
                const { data: user, error } = await client()
                    .from("users")
                    .select("*")
                    .eq("username", username)
                    .eq("password", password)
                    .eq("status", "aktif")
                    .maybeSingle();

                if (error) {
                    throw error;
                }

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
            } catch (error) {
                alert.textContent = handleError(error, "Tidak bisa terhubung ke database.");
                alert.hidden = false;
            }
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
        let cachedItems = [];
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

        async function loadItems() {
            const { data, error } = await client()
                .from("inventory_items")
                .select("*")
                .eq("kategori", category)
                .order("id_item", { ascending: false });

            if (error) {
                throw error;
            }

            cachedItems = data || [];
        }

        function dashboardRows() {
            const query = normalizeSearch(qs("#dashboardSearch").value);
            return cachedItems
                .filter((item) => includesSearch(item, query, ["item", "stock_in_hand", "expire_date", "distributor"]));
        }

        async function renderItems() {
            try {
                await loadItems();
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
            } catch (error) {
                qs("#itemRows").innerHTML = `<tr><td colspan="6" class="empty-state">${escapeHtml(handleError(error, "Data gagal dimuat."))}</td></tr>`;
            }
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
                const item = cachedItems.find((row) => Number(row.id_item) === Number(idItem));
                qs("#item").value = item.item;
                qs("#expire_date").value = item.expire_date;
                qs("#distributor").value = item.distributor;
            }

            openOverlay(sheet);
        }

        qs("#openFormButton").addEventListener("click", () => openItemForm());
        qs("#dashboardSearch").addEventListener("input", () => {
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
        });
        qs("#exportDashboardButton").addEventListener("click", exportDashboard);
        qsa("[data-close-sheet]").forEach((button) => button.addEventListener("click", () => closeOverlay(sheet)));
        qs("#cancelDelete").addEventListener("click", () => closeOverlay(deleteModal));
        qs("#confirmDelete").addEventListener("click", async () => {
            try {
                const { error } = await client().from("inventory_items").delete().eq("id_item", deleteId);
                if (error) {
                    throw error;
                }
                closeOverlay(deleteModal);
                await renderItems();
            } catch (error) {
                showAlert("#formAlert", handleError(error, "Item gagal dihapus."));
            }
        });

        qs("#itemForm").addEventListener("submit", async function (event) {
            event.preventDefault();
            qs("#formAlert").hidden = true;

            try {
                const idItem = Number(qs("#editItemId").value);
                const payload = {
                    item: qs("#item").value.trim(),
                    expire_date: qs("#expire_date").value,
                    distributor: qs("#distributor").value.trim()
                };

                if (idItem) {
                    const { error: itemError } = await client()
                        .from("inventory_items")
                        .update(payload)
                        .eq("id_item", idItem);
                    if (itemError) {
                        throw itemError;
                    }

                    const { error: detailError } = await client()
                        .from("inventory_item_details")
                        .update({
                            expire_date: payload.expire_date,
                            distributor: payload.distributor
                        })
                        .eq("id_item", idItem);
                    if (detailError) {
                        throw detailError;
                    }
                } else {
                    const { data: newItem, error: itemError } = await client()
                        .from("inventory_items")
                        .insert({
                            kategori: category,
                            stock_in_hand: 0,
                            ...payload
                        })
                        .select()
                        .single();
                    if (itemError) {
                        throw itemError;
                    }

                    const { error: detailError } = await client()
                        .from("inventory_item_details")
                        .insert({
                            id_item: newItem.id_item,
                            tanggal: qs("#tanggal").value,
                            jumlah_masuk: Number(qs("#jumlah_masuk").value) || 0,
                            nomor_batch: qs("#nomor_batch").value.trim(),
                            expire_date: payload.expire_date,
                            satuan: qs("#satuan").value.trim(),
                            jumlah_keluar: Number(qs("#jumlah_keluar").value) || 0,
                            distributor: payload.distributor
                        });
                    if (detailError) {
                        throw detailError;
                    }
                    await updateItemStock(newItem.id_item);
                }

                closeOverlay(sheet);
                await renderItems();
            } catch (error) {
                showAlert("#formAlert", handleError(error, "Item gagal disimpan."));
            }
        });

        renderItems();
    }

    async function initDetailPage() {
        const activeSession = requireLogin();
        if (!activeSession) {
            return;
        }

        bindLogout();
        const idItem = Number(params().get("id_item"));
        let item = null;
        let cachedDetails = [];

        try {
            item = await getItem(idItem);
        } catch {
            window.location.href = "dashboard.html";
            return;
        }

        const sheet = qs("#detailSheet");
        const deleteModal = qs("#deleteDetailModal");
        let deleteId = null;

        async function refreshItem() {
            item = await getItem(idItem);
            qs("#categoryLabel").textContent = categoryLabels[item.kategori] || item.kategori;
            qs("#itemTitle").textContent = item.item;
            qs("#sheetItemName").textContent = item.item;
            qs("#stockInHand").textContent = item.stock_in_hand;
            qs("#backToCategory").href = "dashboard.html?kategori=" + encodeURIComponent(item.kategori);
            qs("#backToCategory").textContent = "Kembali ke " + (categoryLabels[item.kategori] || item.kategori);
            document.title = item.item + " | StockCare Lab";
        }

        async function loadDetails() {
            const { data, error } = await client()
                .from("inventory_item_details")
                .select("*")
                .eq("id_item", idItem)
                .order("tanggal", { ascending: true })
                .order("id_detail", { ascending: true });

            if (error) {
                throw error;
            }

            let runningStock = 0;
            cachedDetails = (data || []).map((detail) => {
                runningStock += Number(detail.jumlah_masuk) - Number(detail.jumlah_keluar);
                return { ...detail, jumlah_stok: runningStock };
            });
        }

        function detailRowsWithStock() {
            const query = normalizeSearch(qs("#detailSearch").value);
            const rows = cachedDetails.filter((detail) => includesSearch(detail, query, [
                "tanggal",
                "jumlah_masuk",
                "nomor_batch",
                "expire_date",
                "satuan",
                "jumlah_keluar",
                "jumlah_stok",
                "distributor"
            ]));
            return qs("#urutan").value === "terbaru" ? [...rows].reverse() : rows;
        }

        async function renderDetails() {
            try {
                await refreshItem();
                await loadDetails();
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
            } catch (error) {
                qs("#detailRows").innerHTML = `<tr><td colspan="9" class="empty-state">${escapeHtml(handleError(error, "Data detail gagal dimuat."))}</td></tr>`;
            }
        }

        function renderFilteredDetails() {
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
                const detail = cachedDetails.find((row) => Number(row.id_detail) === Number(idDetail));
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

        qs("#urutan").addEventListener("change", renderFilteredDetails);
        qs("#detailSearch").addEventListener("input", renderFilteredDetails);
        qs("#exportDetailButton").addEventListener("click", exportDetails);
        qs("#openDetailFormButton").addEventListener("click", () => openDetailForm());
        qsa("[data-close-sheet]").forEach((button) => button.addEventListener("click", () => closeOverlay(sheet)));
        qs("#cancelDeleteDetail").addEventListener("click", () => closeOverlay(deleteModal));
        qs("#confirmDeleteDetail").addEventListener("click", async () => {
            try {
                const { error } = await client().from("inventory_item_details").delete().eq("id_detail", deleteId);
                if (error) {
                    throw error;
                }
                await updateItemStock(idItem);
                closeOverlay(deleteModal);
                await renderDetails();
            } catch (error) {
                showAlert("#detailAlert", handleError(error, "Detail gagal dihapus."));
            }
        });

        qs("#detailForm").addEventListener("submit", async function (event) {
            event.preventDefault();
            qs("#detailAlert").hidden = true;

            try {
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

                const request = idDetail
                    ? client().from("inventory_item_details").update(payload).eq("id_detail", idDetail)
                    : client().from("inventory_item_details").insert(payload);
                const { error } = await request;
                if (error) {
                    throw error;
                }

                await updateItemStock(idItem);
                closeOverlay(sheet);
                await renderDetails();
            } catch (error) {
                showAlert("#detailAlert", handleError(error, "Detail gagal disimpan."));
            }
        });

        renderDetails();
    }

    return {
        initLoginPage,
        initDashboardPage,
        initDetailPage,
        database
    };
})();

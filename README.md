<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Librería y Variedades Amaya - Sistema POS & Catálogo</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS para Excel -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: rgba(0,0,0,0.05); }
        ::-webkit-scrollbar-thumb { background: rgba(0,0,0,0.2); border-radius: 4px; }
    </style>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen flex flex-col">

    <!-- BARRA SUPERIOR DE NAVEGACIÓN -->
    <header class="bg-slate-900 text-white shadow-md sticky top-0 z-30">
        <div class="max-w-7xl mx-auto px-4 flex flex-wrap justify-between items-center h-16">
            <div class="flex items-center space-x-3">
                <div class="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center text-white font-bold text-xl shadow-lg">
                    <i class="fa-solid fa-store"></i>
                </div>
                <div>
                    <h1 id="headerStoreName" class="font-bold text-base leading-tight">Librería y Variedades Amaya</h1>
                    <span class="text-xs text-slate-400">Sensuntepeque, El Salvador</span>
                </div>
            </div>

            <!-- BOTONES DEL MENÚ DE NAVEGACIÓN -->
            <nav class="flex items-center space-x-1 overflow-x-auto py-1">
                <button onclick="openMenu('pos')" id="btnMenuPos" class="menu-btn px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition bg-indigo-600 text-white shadow whitespace-nowrap">
                    <i class="fa-solid fa-cart-shopping"></i> POS / Ventas
                </button>
                <button onclick="openMenu('catalog')" id="btnMenuCatalog" class="menu-btn px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition text-slate-300 hover:text-white hover:bg-slate-800 whitespace-nowrap">
                    <i class="fa-solid fa-globe text-emerald-400"></i> Catálogo & WhatsApp
                </button>
                <button onclick="openMenu('webOrders')" id="btnMenuWebOrders" class="menu-btn px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition text-slate-300 hover:text-white hover:bg-slate-800 whitespace-nowrap">
                    <i class="fa-solid fa-bag-shopping"></i> Pedidos Web
                </button>
                <button onclick="openMenu('inventory')" id="btnMenuInventory" class="menu-btn px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition text-slate-300 hover:text-white hover:bg-slate-800 whitespace-nowrap">
                    <i class="fa-solid fa-boxes-stacked"></i> Inventario
                </button>
                <button onclick="openMenu('history')" id="btnMenuHistory" class="menu-btn px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition text-slate-300 hover:text-white hover:bg-slate-800 whitespace-nowrap">
                    <i class="fa-solid fa-clock-rotate-left"></i> Historial
                </button>
                <button onclick="openMenu('settings')" id="btnMenuSettings" class="menu-btn px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition text-slate-300 hover:text-white hover:bg-slate-800 whitespace-nowrap">
                    <i class="fa-solid fa-sliders"></i> Ajustes
                </button>
            </nav>
        </div>
    </header>

    <div id="toastContainer" class="fixed bottom-4 right-4 z-50 flex flex-col gap-2"></div>

    <main class="flex-grow max-w-7xl w-full mx-auto p-4">

        <!-- 1. VISTA POS -->
        <section id="sectionPos" class="app-section grid grid-cols-1 lg:grid-cols-12 gap-6">
            <div class="lg:col-span-8 flex flex-col space-y-4">
                <div class="bg-white p-4 rounded-xl shadow-sm border flex gap-3 items-center">
                    <i class="fa-solid fa-barcode text-slate-400"></i>
                    <input type="text" id="posSearch" onkeydown="handleBarcode(event)" oninput="renderPosList()" placeholder="Escanea o busca un producto y presiona ENTER..." class="w-full text-xs font-bold outline-none">
                </div>
                <div class="bg-white rounded-xl shadow-sm border overflow-hidden max-h-[500px] overflow-y-auto">
                    <table class="w-full text-left text-xs">
                        <thead class="bg-slate-100 uppercase border-b">
                            <tr><th class="p-3">Código</th><th class="p-3">Producto</th><th class="p-3">Categoría</th><th class="p-3 text-right">Precio</th><th class="p-3 text-center">Stock</th><th class="p-3 text-center">Acción</th></tr>
                        </thead>
                        <tbody id="posTableBody" class="divide-y"></tbody>
                    </table>
                </div>
            </div>

            <div class="lg:col-span-4">
                <div class="bg-white rounded-xl shadow border p-4 space-y-3 sticky top-20 text-xs">
                    <h3 class="font-bold border-b pb-2"><i class="fa-solid fa-receipt text-indigo-600 mr-2"></i>Facturación</h3>
                    <div id="cartList" class="max-h-60 overflow-y-auto space-y-2 divide-y"></div>
                    <div class="border-t pt-2 space-y-2">
                        <div class="flex justify-between font-bold text-sm"><span>TOTAL:</span><span id="cartTotal" class="text-indigo-600">$0.00</span></div>
                        <input type="number" id="paidAmount" oninput="calcChange()" placeholder="Efectivo $" class="w-full border p-2 rounded">
                        <div class="flex justify-between font-bold"><span>Cambio:</span><span id="changeVal">$0.00</span></div>
                        <button onclick="checkout()" class="w-full bg-emerald-600 text-white font-bold py-2.5 rounded shadow">Cobrar e Imprimir</button>
                    </div>
                </div>
            </div>
        </section>

        <!-- 2. VISTA CATÁLOGOY WHATSAPP -->
        <section id="sectionCatalog" class="app-section hidden space-y-6">
            <div class="bg-gradient-to-r from-emerald-800 to-slate-900 text-white p-6 rounded-2xl shadow flex justify-between items-center">
                <div>
                    <h2 class="text-2xl font-black">Catálogo Digital & Pedidos</h2>
                    <p class="text-xs text-slate-300">Tus clientes pueden armar su pedido y enviarlo directamente a WhatsApp.</p>
                </div>
                <button onclick="copyPromoText()" class="bg-purple-600 hover:bg-purple-500 text-white px-4 py-2 rounded-lg font-bold text-xs">
                    <i class="fa-brands fa-instagram mr-1"></i> Texto para Instagram
                </button>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
                <div class="lg:col-span-8">
                    <div id="catalogGrid" class="grid grid-cols-2 sm:grid-cols-3 gap-4"></div>
                </div>
                <div class="lg:col-span-4 bg-white p-4 rounded-xl border shadow text-xs space-y-3">
                    <h3 class="font-bold border-b pb-2">Tu Pedido en Línea</h3>
                    <div id="onlineCartList" class="space-y-2 max-h-50 overflow-y-auto"></div>
                    <div class="border-t pt-2 space-y-2">
                        <div class="flex justify-between font-bold text-sm"><span>Total:</span><span id="onlineTotal" class="text-emerald-600">$0.00</span></div>
                        <input type="text" id="custName" placeholder="Nombre completo" class="w-full border p-2 rounded">
                        <input type="text" id="custAddr" placeholder="Dirección de entrega" class="w-full border p-2 rounded">
                        <button onclick="sendWhatsApp()" class="w-full bg-emerald-600 text-white font-bold py-2.5 rounded shadow flex justify-center items-center gap-2">
                            <i class="fa-brands fa-whatsapp text-base"></i> Enviar Pedido por WhatsApp
                        </button>
                    </div>
                </div>
            </div>
        </section>

        <!-- 3. VISTA PEDIDOS WEB INTERNOS -->
        <section id="sectionWebOrders" class="app-section hidden space-y-4">
            <h2 class="font-bold text-base"><i class="fa-solid fa-globe text-indigo-600 mr-2"></i>Pedidos Web Recibidos</h2>
            <div id="webOrdersGrid" class="grid grid-cols-1 md:grid-cols-3 gap-4"></div>
        </section>

        <!-- 4. VISTA INVENTARIO Y EXCEL -->
        <section id="sectionInventory" class="app-section hidden space-y-4">
            <div class="bg-white p-4 rounded-xl border shadow flex justify-between items-center">
                <input type="text" id="invSearch" oninput="renderInventory()" placeholder="Buscar en inventario..." class="border p-2 rounded text-xs w-64">
                <div class="flex gap-2">
                    <button onclick="exportExcel()" class="bg-slate-700 text-white font-bold px-3 py-1.5 rounded text-xs">Exportar Excel</button>
                    <button onclick="openModal()" class="bg-indigo-600 text-white font-bold px-3 py-1.5 rounded text-xs">+ Producto</button>
                </div>
            </div>
            <div class="bg-white rounded-xl border overflow-x-auto">
                <table class="w-full text-left text-xs">
                    <thead class="bg-slate-100 uppercase border-b">
                        <tr><th class="p-3">Código</th><th class="p-3">Nombre</th><th class="p-3">Categoría</th><th class="p-3 text-right">Precio</th><th class="p-3 text-center">Stock</th><th class="p-3 text-right">Acción</th></tr>
                    </thead>
                    <tbody id="inventoryBody" class="divide-y"></tbody>
                </table>
            </div>
        </section>

        <!-- 5. VISTA HISTORIAL -->
        <section id="sectionHistory" class="app-section hidden space-y-4">
            <h2 class="font-bold text-base"><i class="fa-solid fa-clock-rotate-left text-indigo-600 mr-2"></i>Historial de Ventas</h2>
            <div class="bg-white rounded-xl border overflow-x-auto">
                <table class="w-full text-left text-xs">
                    <thead class="bg-slate-100 border-b">
                        <tr><th class="p-3"># Ticket</th><th class="p-3">Fecha</th><th class="p-3">Ítems</th><th class="p-3">Total</th></tr>
                    </thead>
                    <tbody id="historyBody" class="divide-y"></tbody>
                </table>
            </div>
        </section>

        <!-- 6. VISTA AJUSTES -->
        <section id="sectionSettings" class="app-section hidden space-y-4 max-w-md mx-auto bg-white p-6 rounded-xl border shadow text-xs">
            <h2 class="font-bold text-base border-b pb-2">Ajustes del Sistema</h2>
            <div><label class="block font-semibold mb-1">Nombre</label><input type="text" id="cfgName" class="w-full border p-2 rounded"></div>
            <div><label class="block font-semibold mb-1">WhatsApp</label><input type="text" id="cfgPhone" class="w-full border p-2 rounded" placeholder="+50323456789"></div>
            <div><label class="block font-semibold mb-1">Dirección</label><input type="text" id="cfgAddr" class="w-full border p-2 rounded"></div>
            <button onclick="saveSettings()" class="w-full bg-indigo-600 text-white font-bold py-2 rounded shadow">Guardar Cambios</button>
        </section>
    </main>

    <!-- MODAL NUEVO PRODUCTO -->
    <div id="productModal" class="fixed inset-0 bg-slate-900/50 hidden flex items-center justify-center p-4 z-50">
        <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-5 space-y-3 text-xs">
            <h3 class="font-bold border-b pb-2">Registrar Producto</h3>
            <form onsubmit="saveProd(event)" class="space-y-2">
                <div><label class="font-semibold">Código</label><input type="text" id="pCode" required class="w-full border p-1.5 rounded"></div>
                <div><label class="font-semibold">Nombre</label><input type="text" id="pName" required class="w-full border p-1.5 rounded"></div>
                <div><label class="font-semibold">Categoría</label><input type="text" id="pCat" placeholder="Papelería" class="w-full border p-1.5 rounded"></div>
                <div class="grid grid-cols-2 gap-2">
                    <div><label class="font-semibold">Precio ($)</label><input type="number" id="pPrice" step="0.01" required class="w-full border p-1.5 rounded"></div>
                    <div><label class="font-semibold">Stock</label><input type="number" id="pStock" required class="w-full border p-1.5 rounded"></div>
                </div>
                <div class="flex justify-end gap-2 pt-2">
                    <button type="button" onclick="closeModal()" class="px-3 py-1 border rounded">Cancelar</button>
                    <button type="submit" class="px-3 py-1 bg-indigo-600 text-white font-bold rounded">Guardar</button>
                </div>
            </form>
        </div>
    </div>

    <!-- SCRIPT DE NAVEGACIÓN Y CONTROL -->
    <script>
        let products = JSON.parse(localStorage.getItem('pos_products')) || [
            { id: '1', code: 'C0001', name: 'Cuaderno Espiral 200h', category: 'Papelería', price: 2.25, stock: 45 },
            { id: '2', code: 'B0001', name: 'Bolígrafo Gel Azul', category: 'Papelería', price: 0.60, stock: 100 }
        ];
        let cart = [];
        let onlineCart = [];
        let salesHistory = JSON.parse(localStorage.getItem('pos_sales')) || [];
        let storeSettings = JSON.parse(localStorage.getItem('pos_settings')) || { name: 'Librería y Variedades Amaya', phone: '+50323456789', address: 'Sensuntepeque' };

        // NAVEGACIÓN DIRECTA ENTRE MENÚS
        function openMenu(menuKey) {
            // 1. Ocultar todas las secciones
            document.querySelectorAll('.app-section').forEach(sec => sec.classList.add('hidden'));

            // 2. Desactivar estilos en todos los botones
            document.querySelectorAll('.menu-btn').forEach(btn => {
                btn.className = "menu-btn px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition text-slate-300 hover:text-white hover:bg-slate-800 whitespace-nowrap";
            });

            // 3. Mostrar la sección seleccionada
            const targetSection = document.getElementById('section' + menuKey.charAt(0).toUpperCase() + menuKey.slice(1));
            if (targetSection) targetSection.classList.remove('hidden');

            // 4. Activar el botón seleccionado
            const activeBtn = document.getElementById('btnMenu' + menuKey.charAt(0).toUpperCase() + menuKey.slice(1));
            if (activeBtn) {
                activeBtn.className = "menu-btn px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition bg-indigo-600 text-white shadow whitespace-nowrap";
            }

            // 5. Cargar datos específicos de la vista
            if (menuKey === 'pos') renderPosList();
            if (menuKey === 'catalog') renderCatalog();
            if (menuKey === 'inventory') renderInventory();
            if (menuKey === 'history') renderHistory();
            if (menuKey === 'settings') loadSettings();
        }

        document.addEventListener('DOMContentLoaded', () => { openMenu('pos'); });

        function showToast(msg) {
            const container = document.getElementById('toastContainer');
            const toast = document.createElement('div');
            toast.className = "bg-emerald-600 text-white text-xs px-3 py-2 rounded shadow font-semibold";
            toast.textContent = msg;
            container.appendChild(toast);
            setTimeout(() => toast.remove(), 2500);
        }

        // --- POS Y LECTURA DE BARRAS ---
        function renderPosList() {
            const tbody = document.getElementById('posTableBody');
            const search = document.getElementById('posSearch').value.toLowerCase().trim();
            const filtered = products.filter(p => p.name.toLowerCase().includes(search) || p.code.toLowerCase().includes(search));

            tbody.innerHTML = filtered.map(p => `
                <tr class="hover:bg-slate-50 border-b">
                    <td class="p-3 font-mono font-bold">${p.code}</td>
                    <td class="p-3 font-semibold">${p.name}</td>
                    <td class="p-3 text-slate-500">${p.category}</td>
                    <td class="p-3 text-right font-bold text-indigo-600">$${p.price.toFixed(2)}</td>
                    <td class="p-3 text-center font-bold">${p.stock}</td>
                    <td class="p-3 text-center">
                        <button onclick="addToCart('${p.id}')" class="bg-indigo-600 text-white font-bold px-2 py-1 rounded text-xs"><i class="fa-solid fa-plus"></i></button>
                    </td>
                </tr>
            `).join('');
        }

        function handleBarcode(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                const match = products.find(p => p.code.toLowerCase() === e.target.value.trim().toLowerCase());
                if (match) { addToCart(match.id); e.target.value = ''; }
                else showToast("Código no encontrado");
            }
        }

        function addToCart(id) {
            const prod = products.find(p => p.id === id);
            if (!prod || prod.stock <= 0) return showToast("Sin existencias");
            const item = cart.find(i => i.id === id);
            if (item) item.qty++;
            else cart.push({ id: prod.id, name: prod.name, price: prod.price, qty: 1 });
            renderCart();
        }

        function renderCart() {
            const list = document.getElementById('cartList');
            let total = 0;
            list.innerHTML = cart.map(i => {
                total += i.price * i.qty;
                return `<div class="flex justify-between py-1"><span>${i.name} (x${i.qty})</span><span class="font-bold">$${(i.price*i.qty).toFixed(2)}</span></div>`;
            }).join('');
            document.getElementById('cartTotal').textContent = `$${total.toFixed(2)}`;
            calcChange();
        }

        function calcChange() {
            const total = parseFloat(document.getElementById('cartTotal').textContent.replace('$', '')) || 0;
            const paid = parseFloat(document.getElementById('paidAmount').value) || 0;
            document.getElementById('changeVal').textContent = `$${Math.max(0, paid - total).toFixed(2)}`;
        }

        function checkout() {
            if (cart.length === 0) return;
            cart.forEach(item => {
                const p = products.find(prod => prod.id === item.id);
                if (p) p.stock -= item.qty;
            });
            salesHistory.push({ ticketNum: salesHistory.length + 1, date: new Date().toISOString(), items: [...cart], total: parseFloat(document.getElementById('cartTotal').textContent.replace('$', '')) });
            localStorage.setItem('pos_sales', JSON.stringify(salesHistory));
            localStorage.setItem('pos_products', JSON.stringify(products));
            cart = [];
            renderCart();
            renderPosList();
            showToast("Venta realizada");
        }

        // --- CATÁLOGOWHATSAPP ---
        function renderCatalog() {
            const grid = document.getElementById('catalogGrid');
            grid.innerHTML = products.filter(p => p.stock > 0).map(p => `
                <div class="bg-white border rounded-xl p-3 shadow-sm flex flex-col justify-between space-y-2">
                    <div><span class="text-[9px] font-bold text-slate-400">${p.category}</span><h4 class="font-bold text-xs mt-1">${p.name}</h4></div>
                    <div class="flex justify-between items-center border-t pt-2">
                        <span class="font-bold text-emerald-600">$${p.price.toFixed(2)}</span>
                        <button onclick="addOnline('${p.id}')" class="bg-emerald-600 text-white font-bold px-2 py-1 rounded text-xs"><i class="fa-solid fa-plus"></i></button>
                    </div>
                </div>
            `).join('');
        }

        function addOnline(id) {
            const p = products.find(prod => prod.id === id);
            const item = onlineCart.find(i => i.id === id);
            if (item) item.qty++;
            else onlineCart.push({ id: p.id, name: p.name, price: p.price, qty: 1 });
            renderOnlineCart();
        }

        function renderOnlineCart() {
            const list = document.getElementById('onlineCartList');
            let total = 0;
            list.innerHTML = onlineCart.map(i => {
                total += i.price * i.qty;
                return `<div class="flex justify-between text-xs py-1"><span>${i.qty}x ${i.name}</span><span class="font-bold">$${(i.price*i.qty).toFixed(2)}</span></div>`;
            }).join('');
            document.getElementById('onlineTotal').textContent = `$${total.toFixed(2)}`;
        }

        function sendWhatsApp() {
            if (onlineCart.length === 0) return alert("Agrega artículos al pedido");
            const name = document.getElementById('custName').value.trim();
            const addr = document.getElementById('custAddr').value.trim();
            if (!name || !addr) return alert("Ingresa Nombre y Dirección");

            let text = `*NUEVO PEDIDO DE CLIENTE*\n\n*Cliente:* ${name}\n*Dirección:* ${addr}\n\n*Productos:*\n`;
            let total = 0;
            onlineCart.forEach(i => { total += i.price * i.qty; text += `• ${i.qty}x ${i.name} ($${(i.price*i.qty).toFixed(2)})\n`; });
            text += `\n*TOTAL: $${total.toFixed(2)}*`;

            const phone = storeSettings.phone.replace(/[^0-9]/g, '');
            window.open(`https://wa.me/${phone}?text=${encodeURIComponent(text)}`, '_blank');
        }

        function copyPromoText() {
            navigator.clipboard.writeText(`📚 *${storeSettings.name}*\n¡Escríbenos para hacer tu pedido!\n📱 WhatsApp: ${storeSettings.phone}`);
            showToast("Texto para Instagram copiado");
        }

        // --- INVENTARIO Y CONFIGURACIÓN ---
        function renderInventory() {
            const tbody = document.getElementById('inventoryBody');
            tbody.innerHTML = products.map(p => `
                <tr class="hover:bg-slate-50 border-b">
                    <td class="p-3 font-mono font-bold">${p.code}</td>
                    <td class="p-3 font-semibold">${p.name}</td>
                    <td class="p-3 text-slate-500">${p.category}</td>
                    <td class="p-3 text-right font-bold">$${p.price.toFixed(2)}</td>
                    <td class="p-3 text-center font-bold">${p.stock}</td>
                    <td class="p-3 text-right"><button onclick="deleteProd('${p.id}')" class="text-rose-600"><i class="fa-solid fa-trash"></i></button></td>
                </tr>
            `).join('');
        }

        function openModal() { document.getElementById('productModal').classList.remove('hidden'); }
        function closeModal() { document.getElementById('productModal').classList.add('hidden'); }

        function saveProd(e) {
            e.preventDefault();
            products.push({
                id: Date.now().toString(),
                code: document.getElementById('pCode').value,
                name: document.getElementById('pName').value,
                category: document.getElementById('pCat').value || 'Papelería',
                price: parseFloat(document.getElementById('pPrice').value),
                stock: parseInt(document.getElementById('pStock').value)
            });
            localStorage.setItem('pos_products', JSON.stringify(products));
            closeModal();
            renderInventory();
            showToast("Producto agregado");
        }

        function deleteProd(id) {
            products = products.filter(p => p.id !== id);
            localStorage.setItem('pos_products', JSON.stringify(products));
            renderInventory();
        }

        function renderHistory() {
            const tbody = document.getElementById('historyBody');
            tbody.innerHTML = salesHistory.map(s => `
                <tr class="border-b"><td class="p-3 font-bold">#${s.ticketNum}</td><td class="p-3">${new Date(s.date).toLocaleDateString()}</td><td class="p-3">${s.items.map(i=>i.name).join(', ')}</td><td class="p-3 font-bold text-emerald-600">$${s.total.toFixed(2)}</td></tr>
            `).join('');
        }

        function exportExcel() {
            const ws = XLSX.utils.json_to_sheet(products);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Inventario");
            XLSX.writeFile(wb, "Inventario_Amaya.xlsx");
        }

        function loadSettings() {
            document.getElementById('cfgName').value = storeSettings.name;
            document.getElementById('cfgPhone').value = storeSettings.phone;
            document.getElementById('cfgAddr').value = storeSettings.address;
        }

        function saveSettings() {
            storeSettings.name = document.getElementById('cfgName').value;
            storeSettings.phone = document.getElementById('cfgPhone').value;
            storeSettings.address = document.getElementById('cfgAddr').value;
            localStorage.setItem('pos_settings', JSON.stringify(storeSettings));
            showToast("Ajustes guardados");
        }
    </script>
</body>
</html>

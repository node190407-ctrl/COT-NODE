// DOM Elements
const workspace = document.getElementById('workspace');
const recordCount = document.getElementById('recordCount');
const catalogList = document.getElementById('catalogList');
const aiSearch = document.getElementById('aiSearch');
const clearSearch = document.getElementById('clearSearch');
const aiIndicator = document.getElementById('aiIndicator');
const quoteItemsContainer = document.getElementById('quoteItems');
const quoteSummary = document.getElementById('quoteSummary');
const subtotalVal = document.getElementById('subtotalVal');
const taxVal = document.getElementById('taxVal');
const totalVal = document.getElementById('totalVal');
const aiModal = document.getElementById('aiModal');
const exportQuoteBtn = document.getElementById('exportQuote');
const exportWppBtn = document.getElementById('exportWpp');

// Nuevos Elementos
const currencyToggle = document.getElementById('currencyToggle');
const clientNameInput = document.getElementById('clientName');
const clientCompanyInput = document.getElementById('clientCompany');
const authorNameInput = document.getElementById('authorName');

// IA Mode Elements
const btnModeManual = document.getElementById('btnModeManual');
const btnModeAI = document.getElementById('btnModeAI');
const modeSelectionOverlay = document.getElementById('modeSelectionOverlay');
const bgVideo = document.getElementById('bgVideo');
const videoSource = document.getElementById('videoSource');
const catalogTitle = document.getElementById('catalogTitle');
const aiPromptArea = document.getElementById('aiPromptArea');
const aiPromptInput = document.getElementById('aiPromptInput');
const btnGenerateAI = document.getElementById('btnGenerateAI');
const aiThinkingStatus = document.getElementById('aiThinkingStatus');

// State
let catalogData = []; 
let shoppingCart = []; 
let currentCurrency = 'MXN';
const EXCHANGE_RATE = 15.0; // Factor de conversión a USD

// Precargar Logo PDF
const logoImage = new Image();
logoImage.src = 'logo.png';

// INIT
let appMode = 'MANUAL';
document.addEventListener('DOMContentLoaded', () => {
    aiModal.classList.remove('hidden');
    loadLocalExcel();
    setupEventListeners();
    setupModeSelection();
});

// Auto-Load Excel File
async function loadLocalExcel() {
    try {
        const response = await fetch('1.Productos NODE.xlsx');
        if (!response.ok) throw new Error("No se pudo cargar el archivo");
        
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, {type: 'array'});
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Usar header:1 nos asegura leer las columnas estrictamente como un arreglo (Array) por cada fila
        const jsonArr = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false});
        
        if (jsonArr.length < 2) {
            alert("El archivo base está vacío o no tiene suficientes filas.");
            aiModal.classList.add('hidden');
            return;
        }
        
        // Saltar la cabecera
        const dataRows = jsonArr.slice(1);
        
        // Give UI time to breathe
        setTimeout(() => processDataWithAI(dataRows), 1200);
        
    } catch(error) {
        console.error(error);
        alert("Error leyendo el archivo excel predefinido. Probablemente tengas que abrirlo a través de un servidor (Live Server) por las políticas CORS de los navegadores locales.");
        aiModal.classList.add('hidden');
    }
}

// "AI" Data Structure
function processDataWithAI(rows) {
    catalogData = rows
    .filter(row => row && row.length >= 3 && row[1]) // Asegurar que exista al menos el nombre (Columna B)
    .map((row, index) => {
        // Columna C: Precio (índice 2)
        // Quitamos símbolo de dólar, comas (que en MXN son separadores de miles) y espacios
        let rawPrice = String(row[2] || '0').replace(/[$,\s]/g, '');
        let priceNum = parseFloat(rawPrice);

        return {
            id: String(row[0] || '').trim() || generateCyberId(index), // Col A
            name: String(row[1] || '').trim(),                           // Col B
            price: isNaN(priceNum) ? 0 : priceNum,                       // Col C estructurada
            searchString: row.map(v => String(v)).join(' ').toLowerCase()
        };
    });

    finishLoading();
}

function generateCyberId(index) {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
    let res = 'SKU-';
    for(let i=0; i<4; i++) res += chars.charAt(Math.floor(Math.random() * chars.length));
    return res + '-' + index;
}

function finishLoading() {
    aiModal.classList.add('hidden');
    
    recordCount.innerText = catalogData.length;
    renderCatalog(catalogData);
}

// Render Engine
function renderCatalog(items) {
    catalogList.innerHTML = '';
    
    if (items.length === 0) {
        catalogList.innerHTML = '<div style="color:var(--text-muted); padding:20px;">No se encontraron resultados del Engine.</div>';
        return;
    }

    const fragment = document.createDocumentFragment();
    
    const toRender = items.slice(0, 100); 

    toRender.forEach(item => {
        const card = document.createElement('div');
        card.className = 'product-card';
        card.innerHTML = `
            
            <div class="prod-name">${item.name}</div>
            <div class="prod-price">${formatCurrency(item.price)}</div>
            <button class="add-btn" onclick="addToQuote('${item.id}')">
                <i class="ph ph-plus-circle"></i> Añadir
            </button>
        `;
        fragment.appendChild(card);
    });

    catalogList.appendChild(fragment);
}

// Search
let searchTimeout;
aiSearch.addEventListener('input', (e) => {
    const term = e.target.value.toLowerCase();
    
    aiIndicator.classList.remove('hidden');
    
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => {
        aiIndicator.classList.add('hidden');
        if(term.trim() === '') {
            renderCatalog(catalogData);
            return;
        }

        const words = term.split(' ').filter(w => w.length > 0);
        const filtered = catalogData.filter(item => {
            return words.every(w => item.searchString.includes(w));
        });
        
        renderCatalog(filtered);
    }, 300);
});

clearSearch.addEventListener('click', () => {
    aiSearch.value = '';
    renderCatalog(catalogData);
    aiSearch.focus();
});

// Cart Logistics
window.addToQuote = function(id) {
    const item = catalogData.find(p => p.id === id);
    if (!item) return;

    const existing = shoppingCart.find(p => p.id === id);
    if (existing) {
        existing.qty += 1;
    } else {
        shoppingCart.push({ ...item, qty: 1 });
    }
    
    renderCart();
};

window.updateQty = function(id, delta) {
    const item = shoppingCart.find(p => p.id === id);
    if(!item) return;
    
    item.qty += delta;
    if(item.qty <= 0) {
        removeFromQuote(id);
    } else {
        renderCart();
    }
}

window.removeFromQuote = function(id) {
    shoppingCart = shoppingCart.filter(p => p.id !== id);
    renderCart();
}

function renderCart() {
    quoteItemsContainer.innerHTML = '';
    
    if (shoppingCart.length === 0) {
        quoteItemsContainer.innerHTML = '<div class="empty-quote">No hay productos en la cotización</div>';
        quoteSummary.classList.add('hidden');
        exportQuoteBtn.classList.add('hidden');
        exportWppBtn.classList.add('hidden');
        return;
    }
    
    exportQuoteBtn.classList.remove('hidden');
    exportWppBtn.classList.remove('hidden');
    quoteSummary.classList.remove('hidden');

    const fragment = document.createDocumentFragment();
    let subtotal = 0;

    shoppingCart.forEach(item => {
        const itemTotal = item.price * item.qty;
        subtotal += itemTotal;

        const el = document.createElement('div');
        el.className = 'cart-item';
        el.innerHTML = `
            <div class="cart-item-header">
                <div class="cart-item-name">${item.name}</div>
                <button class="remove-btn" onclick="removeFromQuote('${item.id}')"><i class="ph ph-trash"></i></button>
            </div>
            <div class="cart-item-controls">
                <div class="qty-controls">
                    <button class="qty-btn" onclick="updateQty('${item.id}', -1)"><i class="ph ph-minus"></i></button>
                    <span>${item.qty}</span>
                    <button class="qty-btn" onclick="updateQty('${item.id}', 1)"><i class="ph ph-plus"></i></button>
                </div>
                <div class="cart-item-price">${formatCurrency(itemTotal)}</div>
            </div>
        `;
        fragment.appendChild(el);
    });

    quoteItemsContainer.appendChild(fragment);
    
    const tax = subtotal * 0.16;
    const total = subtotal + tax;
    
    subtotalVal.innerText = formatCurrency(subtotal);
    taxVal.innerText = formatCurrency(tax);
    totalVal.innerText = formatCurrency(total);
}

function formatCurrency(val) {
    let finalVal = currentCurrency === 'USD' ? (val / EXCHANGE_RATE) : val;
    finalVal = Math.ceil(finalVal);
    return new Intl.NumberFormat(currentCurrency === 'MXN' ? 'es-MX' : 'en-US', {
        style: 'currency',
        currency: currentCurrency,
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(finalVal);
}

function setupEventListeners() {
    exportQuoteBtn.addEventListener('click', generatePDF);
    document.getElementById('finalizeQuote').addEventListener('click', () => {
        alert("Cotización procesada exitosamente en el sistema de pruebas.");
    });

    // Toggle Moneda
    currencyToggle.addEventListener('click', (e) => {
        // Permitir click en switch o en labels
        let newCurr = null;
        if(e.target.classList.contains('curr-label')){
            newCurr = e.target.getAttribute('data-curr');
        } else {
            // switch o slider pulsado: alternar
            newCurr = currentCurrency === 'MXN' ? 'USD' : 'MXN';
        }

        if(newCurr && newCurr !== currentCurrency) {
            currentCurrency = newCurr;
            
            // Actualizar UI Toggle
            if(currentCurrency === 'USD') {
                currencyToggle.classList.add('is-usd');
            } else {
                currencyToggle.classList.remove('is-usd');
            }
            
            Array.from(currencyToggle.querySelectorAll('.curr-label')).forEach(el => {
                if(el.getAttribute('data-curr') === currentCurrency) el.classList.add('active');
                else el.classList.remove('active');
            });

            // Re-renderizar catálogo y carrito con la nueva moneda
            renderCatalog(catalogData);
            renderCart();
        }
    });
}

// ═══════════════════════════════════════════════════════════════════════
// generatePDF — NODE Soluciones Tecnológicas · Identity v2.0
// ═══════════════════════════════════════════════════════════════════════
function generatePDF() {
    if (!window.jspdf) {
        alert("El motor PDF está cargando. Verifica tu conexión e intenta de nuevo.");
        return;
    }

    // ── Datos del formulario ─────────────────────────────────────────────
    const cName    = clientNameInput.value.trim()    || "Cliente No Registrado";
    const cCompany = clientCompanyInput.value.trim() || "";
    const aName    = authorNameInput.value.trim()    || "NODE Soluciones Tecnológicas";

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "pt", format: "a4" });

    // ── Paleta NODE Identity v2.0 ────────────────────────────────────────
    const COLOR = {
        ink:        [30, 45, 64],
        indigo:     [67, 56, 202],
        teal:       [13, 148, 136],
        midnight:   [13, 18, 37],
        white:      [255, 255, 255],
        neutral50:  [248, 249, 251],
        neutral100: [241, 243, 247],
        neutral200: [228, 232, 240],
        neutral400: [148, 163, 184],
        neutral500: [100, 116, 139],
        neutral600: [71, 85, 105],
        indigoLight:[238, 242, 255],
        indigoMid:  [199, 210, 254],
    };

    const PAGE_W  = doc.internal.pageSize.getWidth();
    const PAGE_H  = doc.internal.pageSize.getHeight();
    const MARGIN  = 40;
    const CONTENT = PAGE_W - MARGIN * 2;

    // ═══════════════════════════════════════════════════════════════════
    // 1.  HEADER BAND — NODE MEDIANOCHE
    // ═══════════════════════════════════════════════════════════════════
    const HEADER_H = 82;
    doc.setFillColor(...COLOR.midnight);
    doc.rect(0, 0, PAGE_W, HEADER_H, "F");

    if (logoImage.complete && logoImage.naturalHeight > 0) {
        const logoW = 52;
        const logoH = (logoImage.naturalHeight / logoImage.naturalWidth) * logoW;
        doc.addImage(logoImage, "PNG", MARGIN, (HEADER_H - logoH) / 2, logoW, logoH);
    } else {
        doc.setFont("helvetica", "bold");
        doc.setFontSize(22);
        doc.setTextColor(...COLOR.white);
        doc.text("NODE", MARGIN, HEADER_H / 2 + 8);
    }

    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(...COLOR.neutral400);
    doc.text("SOLUCIONES TECNOLÓGICAS", MARGIN, HEADER_H - 14);

    const folio = `CTZ-${new Date().getFullYear()}-${String(Math.floor(Math.random() * 9000) + 1000)}`;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.setTextColor(...COLOR.neutral400);
    doc.text("COTIZACIÓN", PAGE_W - MARGIN, 26, { align: "right" });

    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.setTextColor(...COLOR.white);
    doc.text(folio, PAGE_W - MARGIN, 46, { align: "right" });

    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(...COLOR.neutral500);
    const dateStr = new Date().toLocaleDateString("es-MX", { year: "numeric", month: "long", day: "numeric" });
    doc.text(dateStr, PAGE_W - MARGIN, 62, { align: "right" });

    // ═══════════════════════════════════════════════════════════════════
    // 2.  ACCENT BAR — índigo → teal
    // ═══════════════════════════════════════════════════════════════════
    const BAR_H = 4;
    const BAR_Y = HEADER_H;
    doc.setFillColor(...COLOR.indigo);
    doc.rect(0, BAR_Y, PAGE_W / 2, BAR_H, "F");
    doc.setFillColor(...COLOR.teal);
    doc.rect(PAGE_W / 2, BAR_Y, PAGE_W / 2, BAR_H, "F");

    // ═══════════════════════════════════════════════════════════════════
    // 3.  BLOQUE DE INFORMACIÓN DOBLE COLUMNA
    // ═══════════════════════════════════════════════════════════════════
    let Y = HEADER_H + BAR_H + 24;
    const COL_W = CONTENT / 2 - 12;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(7);
    doc.setTextColor(...COLOR.neutral400);
    doc.text("PARA", MARGIN, Y);

    Y += 12;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(...COLOR.ink);
    doc.text(cName, MARGIN, Y);

    if (cCompany) {
        Y += 14;
        doc.setFont("helvetica", "normal");
        doc.setFontSize(9);
        doc.setTextColor(...COLOR.neutral600);
        doc.text(cCompany, MARGIN, Y);
    }

    const RX = MARGIN + COL_W + 24;
    let RY = HEADER_H + BAR_H + 24;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(7);
    doc.setTextColor(...COLOR.neutral400);
    doc.text("ELABORADO POR", RX, RY);

    RY += 12;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(...COLOR.ink);
    doc.text(aName, RX, RY);

    RY += 14;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(...COLOR.neutral600);
    doc.text("contacto@node.mx", RX, RY);

    RY += 18;
    const BADGE_W = 140;
    const BADGE_H = 18;
    doc.setFillColor(...COLOR.indigoLight);
    doc.roundedRect(RX, RY, BADGE_W, BADGE_H, 4, 4, "F");
    doc.setDrawColor(...COLOR.indigoMid);
    doc.setLineWidth(0.5);
    doc.roundedRect(RX, RY, BADGE_W, BADGE_H, 4, 4, "S");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8);
    doc.setTextColor(...COLOR.indigo);
    doc.text("Válida 15 días naturales", RX + BADGE_W / 2, RY + 11.5, { align: "center" });

    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(...COLOR.neutral400);
    doc.text(`Moneda: ${currentCurrency}`, PAGE_W - MARGIN, RY + 11.5, { align: "right" });

    Y = Math.max(Y + (cCompany ? 14 : 0), RY + BADGE_H) + 20;
    doc.setDrawColor(...COLOR.neutral200);
    doc.setLineWidth(0.5);
    doc.line(MARGIN, Y, PAGE_W - MARGIN, Y);
    Y += 16;

    // ═══════════════════════════════════════════════════════════════════
    // 4.  TABLA DE PRODUCTOS
    // ═══════════════════════════════════════════════════════════════════
    const tableRows = shoppingCart.map(item => {
        const itemTotal = item.price * item.qty;
        return [
            item.id,
            item.name.length > 48 ? item.name.substring(0, 47) + "…" : item.name,
            formatCurrency(item.price),
            String(item.qty),
            formatCurrency(itemTotal),
        ];
    });

    doc.autoTable({
        head: [["ID", "Producto", "P. Unitario", "Cant.", "Total"]],
        body: tableRows,
        startY: Y,
        margin: { left: MARGIN, right: MARGIN },
        theme: "plain",
        styles: {
            font:       "helvetica",
            fontSize:   9,
            cellPadding: { top: 9, bottom: 9, left: 10, right: 10 },
            textColor:  COLOR.neutral600,
            lineColor:  COLOR.neutral100,
            lineWidth:  0.4,
        },
        headStyles: {
            fillColor:    COLOR.neutral50,
            textColor:    COLOR.neutral500,
            fontSize:     7.5,
            fontStyle:    "bold",
            halign:       "left",
            cellPadding:  { top: 8, bottom: 8, left: 10, right: 10 },
            lineColor:    COLOR.neutral200,
            lineWidth:    { bottom: 0.8 },
        },
        alternateRowStyles: {
            fillColor: [250, 250, 252],
        },
        columnStyles: {
            0: { cellWidth: 52,  textColor: COLOR.indigo,  fontStyle: "bold", fontSize: 8 },
            1: { cellWidth: "auto", textColor: COLOR.ink,  fontStyle: "normal" },
            2: { cellWidth: 80,  halign: "right" },
            3: { cellWidth: 44,  halign: "center", fontStyle: "bold", textColor: COLOR.ink },
            4: { cellWidth: 80,  halign: "right",  fontStyle: "bold", textColor: COLOR.ink },
        },
        didDrawCell(data) {
            if (data.section === "body" && data.column.index === 0) {
                const { x, y, width, height } = data.cell;
                doc.setFillColor(...COLOR.indigoLight);
                doc.roundedRect(x + 4, y + (height - 12) / 2, width - 8, 12, 2, 2, "F");
                doc.setFont("helvetica", "bold");
                doc.setFontSize(7.5);
                doc.setTextColor(...COLOR.indigo);
                doc.text(
                    String(data.cell.raw),
                    x + width / 2,
                    y + height / 2 + 1,
                    { align: "center", baseline: "middle" }
                );
            }
        },
    });

    // ═══════════════════════════════════════════════════════════════════
    // 5.  BLOQUE DE TOTALES
    // ═══════════════════════════════════════════════════════════════════
    const TABLE_BOTTOM = doc.lastAutoTable.finalY || Y + 40;
    const TOTALS_X     = PAGE_W - MARGIN - 200;
    let TY             = TABLE_BOTTOM + 16;

    const drawTotalRow = (label, value, opts = {}) => {
        const { bold = false, big = false, accent = false, topLine = false } = opts;
        if (topLine) {
            doc.setDrawColor(...COLOR.neutral200);
            doc.setLineWidth(1);
            doc.line(TOTALS_X, TY - 4, PAGE_W - MARGIN, TY - 4);
        }
        doc.setFont("helvetica", bold ? "bold" : "normal");
        doc.setFontSize(big ? 12 : 9);
        doc.setTextColor(...(accent ? COLOR.indigo : (bold ? COLOR.ink : COLOR.neutral600)));
        doc.text(label, TOTALS_X, TY);
        doc.text(value, PAGE_W - MARGIN, TY, { align: "right" });
        TY += big ? 20 : 16;
    };

    const subtotal = shoppingCart.reduce((s, i) => s + i.price * i.qty, 0);
    const tax      = subtotal * 0.16;
    const total    = subtotal + tax;

    drawTotalRow("Subtotal", formatCurrency(subtotal));
    drawTotalRow("IVA (16%)", formatCurrency(tax));
    drawTotalRow(
        `TOTAL ${currentCurrency}`,
        formatCurrency(total),
        { bold: true, big: true, accent: true, topLine: true }
    );

    // ═══════════════════════════════════════════════════════════════════
    // 6.  NOTAS DE CONDICIONES
    // ═══════════════════════════════════════════════════════════════════
    TY += 10;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(...COLOR.neutral400);
    doc.text(
        "Anticipo del 50% al firmar contrato · Saldo contra entrega · Precios expresados sin IVA",
        MARGIN, TY
    );
    TY += 11;
    doc.text(
        "Esta cotización no constituye un compromiso de venta hasta recibir anticipo y contrato firmado.",
        MARGIN, TY
    );

    // ═══════════════════════════════════════════════════════════════════
    // 7.  FOOTER BAND — NODE MEDIANOCHE
    // ═══════════════════════════════════════════════════════════════════
    const FOOTER_H = 38;
    const FOOTER_Y = PAGE_H - FOOTER_H;

    doc.setFillColor(...COLOR.midnight);
    doc.rect(0, FOOTER_Y, PAGE_W, FOOTER_H, "F");

    doc.setFillColor(...COLOR.indigo);
    doc.circle(MARGIN + 5, FOOTER_Y + FOOTER_H / 2, 3, "F");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.5);
    doc.setTextColor(...COLOR.white);
    doc.text("NODE.MX", MARGIN + 14, FOOTER_Y + FOOTER_H / 2 + 3);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(...COLOR.neutral500);
    doc.text("contacto@node.mx  ·  node.mx", MARGIN + 14, FOOTER_Y + FOOTER_H / 2 + 14);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(...COLOR.neutral500);
    doc.text(`Folio: ${folio}`, PAGE_W - MARGIN, FOOTER_Y + FOOTER_H / 2 + 3, { align: "right" });

    // ═══════════════════════════════════════════════════════════════════
    // 8.  GUARDAR
    // ═══════════════════════════════════════════════════════════════════
    const filename = `NODE_Cotizacion_${cName.replace(/\s+/g, "_")}_${folio}.pdf`;
    doc.save(filename);
}

function setupModeSelection() {
    btnModeManual.addEventListener('click', () => {
        modeSelectionOverlay.classList.add('hidden');
        appMode = 'MANUAL';
    });

    btnModeAI.addEventListener('click', () => {
        modeSelectionOverlay.classList.add('hidden');
        appMode = 'AI';
        
        // Cambiar a Video de AI
        videoSource.src = 'NODE IA.mp4';
        bgVideo.load();
        
        // Esconder barra manual
        document.querySelector('.ai-search-bar').classList.add('hidden');
        catalogTitle.innerText = "Asistente de Cotización con IA";
        aiPromptArea.classList.remove('hidden');
        catalogList.classList.add('hidden'); // Ocultar lista genérica
    });
}

// Lógica de WhatsApp
exportWppBtn.addEventListener('click', () => {
    if(shoppingCart.length === 0) return;
    
    const cName = clientNameInput.value.trim() || 'No Definido';
    const aName = authorNameInput.value.trim() || 'Sistema Automático NODE';
    
    let text = `*Cotización NODE*\nCliente: ${cName}\nElaborado por: ${aName}\nMoneda: ${currentCurrency}\n\n`;
    shoppingCart.forEach(i => {
        const itemTotal = i.price * i.qty;
        let finalVal = currentCurrency === 'USD' ? (itemTotal / EXCHANGE_RATE) : itemTotal;
        finalVal = Math.ceil(finalVal);
        const priceStr = new Intl.NumberFormat(currentCurrency === 'MXN' ? 'es-MX' : 'en-US', {
            style: 'currency', currency: currentCurrency, minimumFractionDigits: 0, maximumFractionDigits: 0
        }).format(finalVal);
        
        text += `- ${i.qty}x ${i.name.substring(0, 30)}... (${priceStr})\n`;
    });
    
    text += `\n*Subtotal:* ${subtotalVal.innerText}`;
    text += `\n*IVA (16%):* ${taxVal.innerText}`;
    text += `\n*Total:* ${totalVal.innerText}`;
    
    const url = `https://wa.me/?text=${encodeURIComponent(text)}`;
    window.open(url, '_blank');
});

// Lógica de API GROQ
const GROQ_API_KEY = "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";

btnGenerateAI.addEventListener('click', async () => {
    const prompt = aiPromptInput.value.trim();
    if(!prompt) return alert("Por favor describe lo que necesitas.");
    
    btnGenerateAI.disabled = true;
    aiThinkingStatus.classList.remove('hidden');
    catalogList.classList.add('hidden');
    
    // Samplear catálogo
    const maxItems = 150; 
    const sample = catalogData.slice(0, maxItems).map(i => `ID:${i.id} | Name:${i.name} | Price:${i.price}`).join('\n');
    
    const sysPrompt = `Eres un cotizador experto. A partir del catálogo de productos que recibes, crea una lista de compra recomendada según lo pedido por el usuario.
Catálogo:
${sample}

Regla vital: Responde ESTRICTAMENTE en formato JSON plano con la siguiente estructura:
[
  {"id": "AQUÍ_ID_DEL_PRODUCTO", "qty": AQUÍ_CANTIDAD_NUMERICA}
]
Nada de explicaciones. Si no encuentras, devuelve [].`;

    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${GROQ_API_KEY}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                model: "llama-3.1-8b-instant",
                messages: [
                    { role: "system", content: sysPrompt },
                    { role: "user", content: prompt }
                ],
                temperature: 0.1
            })
        });
        
        const data = await response.json();
        const text = data.choices[0].message.content;
        
        const jsonMatch = text.match(/\[.*\]/s);
        let products = [];
        if (jsonMatch) {
            products = JSON.parse(jsonMatch[0]);
        } else {
            products = JSON.parse(text);
        }
        
        shoppingCart = [];
        products.forEach(p => {
             const item = catalogData.find(c => c.id === p.id);
             if (item) {
                 shoppingCart.push({ ...item, qty: p.qty || 1 });
             }
        });
        renderCart();
        
        if (shoppingCart.length > 0) {
            catalogList.innerHTML = '';
            const fragment = document.createDocumentFragment();
            shoppingCart.forEach(item => {
                const card = document.createElement('div');
                card.className = 'product-card';
                card.innerHTML = `
                    
                    <div class="prod-name" style="margin-top:5px;">${item.name}</div>
                    <div class="prod-price">${formatCurrency(item.price)}</div>
                    <button class="add-btn" onclick="addToQuote('${item.id}')">
                        <i class="ph ph-plus-circle"></i> Añadir extra manual
                    </button>
                `;
                fragment.appendChild(card);
            });
            catalogList.appendChild(fragment);
            catalogList.classList.remove('hidden');
        } else {
            alert("La IA procesó la solicitud pero no pudo hallar productos que coincidan completamente.");
        }
        
    } catch (e) {
        console.error(e);
        alert("Ocurrió un error leyendo la IA. Revisa consola o intenta ser más directo indicando qué objetos en específico requieres.");
    } finally {
        btnGenerateAI.disabled = false;
        aiThinkingStatus.classList.add('hidden');
    }
});

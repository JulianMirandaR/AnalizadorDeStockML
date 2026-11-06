// Elements
const fileMl = document.getElementById('file-ml');
const fileSys = document.getElementById('file-sys');
const statusMl = document.getElementById('status-ml');
const statusSys = document.getElementById('status-sys');
const btnProcess = document.getElementById('btn-process');
const resultsPanel = document.getElementById('results-panel');
const resultsBody = document.getElementById('results-body');
const btnDownload = document.getElementById('btn-download');

// Drop zones
const dropZoneMl = document.getElementById('drop-zone-ml');
const dropZoneSys = document.getElementById('drop-zone-sys');

// Stats data
const stats = {
    mlMayor: document.getElementById('stat-ml-mayor'),
    missingMl: document.getElementById('stat-missing-ml'),
    warning: document.getElementById('stat-warning'),
    other: document.getElementById('stat-other')
};

let dataMl = null;
let dataSys = null;
let finalResults = [];

// NEW TABS LOGIC
const tabBtns = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');

tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        tabBtns.forEach(b => b.classList.remove('active'));
        tabContents.forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');
    });
});

// --- COPIAR SKU AL PORTAPAPELES ---
// Copia el SKU y muestra un check breve en el botón como confirmación.
function copySku(sku, btn) {
    const text = String(sku == null ? '' : sku).trim();
    if (!text) return;
    const done = () => {
        if (!btn) return;
        const prev = btn.innerHTML;
        btn.classList.add('copied');
        btn.innerHTML = '<i class="fa-solid fa-check"></i>';
        setTimeout(() => { btn.classList.remove('copied'); btn.innerHTML = prev; }, 1200);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(done).catch(() => fallbackCopySku(text, done));
    } else {
        fallbackCopySku(text, done);
    }
}
function fallbackCopySku(text, cb) {
    try {
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.focus(); ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        if (cb) cb();
    } catch (e) { console.error('No se pudo copiar el SKU', e); }
}

// Drag and drop handlers
function setupDropZone(dropZone, fileInput, statusElement, type) {
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            handleFileSelect(fileInput, statusElement, type);
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length === 0) return;
        handleFileSelect(fileInput, statusElement, type);
    });
}

// Auditoría Margaría retirada: sus elementos ya no existen en el HTML, se protege el acceso.
if (dropZoneMl && dropZoneSys) {
    setupDropZone(dropZoneMl, fileMl, statusMl, 'ml');
    setupDropZone(dropZoneSys, fileSys, statusSys, 'sys');
}

function handleFileSelect(input, statusElement, type) {
    if (input.files.length === 0) return;
    
    const file = input.files[0];
    statusElement.textContent = file.name;
    statusElement.classList.add('uploaded');
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Select sheet based on file type
            let targetSheetName = workbook.SheetNames[0];
            
            // ML exports use the 'Publicaciones' sheet
            if (type === 'ml') {
                if (workbook.SheetNames.includes('Publicaciones')) {
                    targetSheetName = 'Publicaciones';
                } else if (workbook.SheetNames.length > 2) {
                    targetSheetName = workbook.SheetNames[2]; // Usually the 3rd sheet if name was changed
                } else if (workbook.SheetNames.length > 1) {
                    targetSheetName = workbook.SheetNames[1];
                }
            }
            
            const worksheet = workbook.Sheets[targetSheetName];
            
            // Convert to array of arrays
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            if (type === 'ml') {
                dataMl = parseMl(json);
            } else {
                dataSys = parseSys(json);
            }
            
            checkReady();
        } catch (error) {
            Swal.fire('Error', 'No se pudo leer el archivo Excel.', 'error');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function checkReady() {
    if (dataMl && dataSys) {
        btnProcess.disabled = false;
    }
}

// parsing logic
// Columna del depósito "Coronel Gil" (Caja Gil 681) en la planilla de ML. El export trae una
// columna de stock por depósito (Asicha, Caja Gil 681, Distri BS AS...), así que se busca por
// el nombre en los encabezados (primeras 5 filas) en vez de por posición.
// Ese depósito de ML tiene publicado el stock de TODAS las sucursales (Coronel Gil + Santiago
// Marzo + 1435), por eso se compara contra el total del sistema y no solo contra Coronel Gil.
// Devuelve el índice de la columna, -1 si la planilla no tiene stock por depósito (export viejo
// de un solo depósito), y lanza error si tiene depósitos pero ninguno es el de Coronel Gil.
function colStockGilMl(rows) {
    let porDeposito = false;
    for (let r = 0; r < Math.min(rows.length, 5); r++) {
        const row = rows[r] || [];
        for (let c = 0; c < row.length; c++) {
            const val = row[c] == null ? '' : String(row[c]);
            if (/^STORE_STOCK_QUANTITY/i.test(val)) porDeposito = true;
            else if (val.length < 40 && /\bgil\b/i.test(val)) return c;
        }
    }
    if (porDeposito) {
        const err = new Error('Sin columna de depósito Coronel Gil en la planilla de ML');
        err.userMessage = 'La planilla de ML tiene stock por depósito pero no encontré el de <b>Coronel Gil</b> (Caja Gil 681). Descargala de nuevo incluyendo ese depósito.';
        throw err;
    }
    return -1;
}

function parseMl(rows) {
    const map = {};
    const colGil = colStockGilMl(rows);
    const stockCol = colGil >= 0 ? colGil : 7; // export viejo: columna H
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        let sku = row[4];
        let stock = row[stockCol];
        let title = row[5];
        
        if (sku !== undefined && sku !== null && sku !== '') {
            sku = sku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            if (sku === "SKU") continue;
            
            const stockVal = parseInt(stock, 10);
            if (!isNaN(stockVal)) {
                map[sku] = {
                    stock: stockVal,
                    title: title ? title.toString().trim() : 'Sin título'
                };
            }
        }
    }
    return map;
}

function parseSys(rows) {
    const map = {};
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        
        let sku = row[3];
        let stock = row[9];
        
        if (sku !== undefined && sku !== null && sku !== '') {
            sku = sku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            
            stock = parseInt(stock, 10);
            if (!isNaN(stock)) {
                map[sku] = stock;
            }
        }
    }
    return map;
}

function parseSysNuevos(rows) {
    const map = {};
    for (let i = 1; i < rows.length; i++) {
        let row = rows[i];
        if (!row || row.length === 0) continue;
        
        if (row.length === 1 && typeof row[0] === 'string' && row[0].includes(';')) {
            row = row[0].split(';');
        } else if (typeof row[0] === 'string' && row[0].includes(';')) {
            let joinedRow = row.join(';');
            row = joinedRow.split(';');
        }
        
        if (row.length < 9) continue;
        
        let sku = row[1];
        let desc = row[2];
        let stockCoronel = parseInt(row[5], 10) || 0;      // Col F = Coronel Gil
        let stockSantiago = parseInt(row[6], 10) || 0;     // Col G = Santiago Marzo
        let stockSantiago1435 = parseInt(row[7], 10) || 0; // Col H = Santiago Marzo 1435
        let stockFull = parseInt(row[8], 10) || 0;         // Col I = ML Full
        // El stock de ML Full NO se cuenta en el total de sucursales: es el depósito
        // de Mercado Libre, no stock propio. Solo suman Coronel Gil, Santiago Marzo y 1435.
        let totalStock = stockCoronel + stockSantiago + stockSantiago1435;
        
        if (sku !== undefined && sku !== null && sku !== '') {
            sku = sku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            if (sku === "SKU") continue;
            
            const descVal = desc ? desc.toString().trim() : 'Neumático Sistema';
            
            if (map.hasOwnProperty(sku)) {
                map[sku].stock += totalStock;
                map[sku].gilStock += stockCoronel;
                map[sku].marzoStock += stockSantiago;
                map[sku].marzo1435Stock += stockSantiago1435;
                map[sku].fullStock += stockFull;
            } else {
                map[sku] = {
                    stock: totalStock,
                    descripcion: descVal,
                    gilStock: stockCoronel,
                    marzoStock: stockSantiago,
                    marzo1435Stock: stockSantiago1435,
                    fullStock: stockFull
                };
            }
        }
    }
    return map;
}

if (btnProcess) btnProcess.addEventListener('click', () => {
    btnProcess.disabled = true;
    btnProcess.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Procesando...';
    
    setTimeout(() => {
        analizarDatos();
        renderTable();
        
        resultsPanel.classList.remove('hidden');
        
        btnProcess.innerHTML = '<i class="fa-solid fa-bolt"></i> Analizar y Comparar Stock';
        btnProcess.disabled = false;
        
        Swal.fire({
            title: '¡Análisis Completado!',
            text: `Se encontraron ${finalResults.length} inconsistencias.`,
            icon: 'success',
            confirmButtonColor: '#3b82f6'
        });
    }, 500); // Small delay for UX
});

function analizarDatos() {
    finalResults = [];
    
    let cMlMayor = 0;
    let cMissingMl = 0;
    let cWarning = 0;
    let cOther = 0;
    
    // Compare ML against System
    for (const rawSku in dataMl) {
        let originalMlStock = dataMl[rawSku].stock;
        let title = dataMl[rawSku].title;
        let skuForSystem = rawSku;
        let multiplier = 1;
        
        // Detectar si es un kit (ej: KITX2-32277 o 32277X2)
        const kitMatchPrefix = rawSku.match(/^KITX(\d+)-(.*)/);
        const kitMatchSuffix = rawSku.match(/^(.*)X(\d+)$/);
        
        if (kitMatchPrefix) {
            multiplier = parseInt(kitMatchPrefix[1], 10);
            skuForSystem = kitMatchPrefix[2]; // el SKU real en el sistema
        } else if (kitMatchSuffix) {
            multiplier = parseInt(kitMatchSuffix[2], 10);
            skuForSystem = kitMatchSuffix[1]; // el SKU real base
        }
        
        const mlStockEq = originalMlStock * multiplier;
        
        if (dataSys.hasOwnProperty(skuForSystem)) {
            const sysStock = dataSys[skuForSystem];
            
            if (mlStockEq !== sysStock) {
                const diff = mlStockEq - sysStock;
                let motivo = '';
                let badgeClass = '';
                
                if (mlStockEq > sysStock && mlStockEq < 10 && sysStock < 10) {
                    motivo = 'Stock ML mayor que Sistema (Riesgo)';
                    badgeClass = 'bg-danger';
                    cMlMayor++;
                } else if (Math.abs(diff) <= 2 && mlStockEq < 8 && sysStock < 8) {
                    motivo = 'Diferencia de 2 unidades o menos (Advertencia)';
                    badgeClass = 'bg-warning';
                    cWarning++;
                } else if (mlStockEq < 4 && sysStock > 6) {
                    motivo = 'Actualizar ML (Sistema tiene más stock)';
                    badgeClass = 'bg-info';
                    cOther++;
                } else {
                    // No cumple con ninguna de las reglas prioritarias de corrección
                    continue;
                }
                
                finalResults.push({
                    SKU: rawSku,
                    Title: title,
                    'Stock ML': originalMlStock,
                    'Multiplier': multiplier, // para mostrarlo en la tabla
                    'Stock Sistema': sysStock,
                    Diferencia: diff,
                    Motivo: motivo,
                    _badgeClass: badgeClass
                });
            }
        } else {
            // Note: User didn't request this case, but good to have if needed. Omitting for exact compliance.
        }
    }
    
    // Compare System against ML (for missing items)
    for (const sku in dataSys) {
        const sysStock = dataSys[sku];
        
        // Rule 3: System has SKU with stock, but it's not in ML
        if (!dataMl.hasOwnProperty(sku) && sysStock > 0) {
            cMissingMl++;
            finalResults.push({
                SKU: sku,
                Title: 'No en Mercado Libre',
                'Stock ML': 'No existe',
                'Stock Sistema': sysStock,
                Diferencia: -sysStock,
                Motivo: 'SKU no existe en ML (Publicar/Activar)',
                _badgeClass: 'bg-success'
            });
        }
    }
    
    // Update stats
    stats.mlMayor.textContent = cMlMayor;
    stats.missingMl.textContent = cMissingMl;
    stats.warning.textContent = cWarning;
    stats.other.textContent = cOther;
}

function renderTable() {
    resultsBody.innerHTML = '';
    
    // Sort by absolute difference from highest to lowest
    finalResults.sort((a, b) => {
        return Math.abs(b.Diferencia) - Math.abs(a.Diferencia);
    });
    
    finalResults.forEach(item => {
        const tr = document.createElement('tr');
        
        const mlDisplay = item.Multiplier > 1 
            ? `${item['Stock ML']} <span style="font-size: 0.8rem; color: var(--warning);">(x${item.Multiplier})</span>`
            : item['Stock ML'];
            
        tr.innerHTML = `
            <td><strong>${item.SKU}</strong></td>
            <td><span style="font-size: 0.85rem; color: var(--text-secondary);">${item.Title || 'Sin detalle'}</span></td>
            <td>${mlDisplay}</td>
            <td>${item['Stock Sistema']}</td>
            <td><span class="status-badge ${item._badgeClass}">${item.Motivo}</span></td>
        `;
        
        resultsBody.appendChild(tr);
    });
}

if (btnDownload) btnDownload.addEventListener('click', () => {
    // Preparar data para XLSX sin la propiedad _badgeClass
    const excelData = finalResults.map(item => {
        return {
            'SKU': item.SKU,
            'Título de Publicación': item.Title || 'Sin detalle',
            'Stock Mercado Libre': item['Stock ML'],
            'Stock Sistema': item['Stock Sistema'],
            'Motivo de Corrección': item.Motivo
        };
    });
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    
    // Ajustar anchos
    const wscols = [
        {wch: 20}, // SKU
        {wch: 45}, // Título de Publicación
        {wch: 20}, // ML
        {wch: 20}, // Sys
        {wch: 50}  // Motivo
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Correcciones Stock");
    
    // Save file
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Conciliacion_Stock_ML_${dateStr}.xlsx`);
});

// --- NUEVOS NEUMATICOS LOGIC ---
const fileMlNuevos = document.getElementById('file-ml-nuevos');
const fileSysNuevos = document.getElementById('file-sys-nuevos');
const statusMlNuevos = document.getElementById('status-ml-nuevos');
const statusSysNuevos = document.getElementById('status-sys-nuevos');
const btnProcessNuevos = document.getElementById('btn-process-nuevos');
const resultsPanelNuevos = document.getElementById('results-panel-nuevos');
const resultsBodyNuevos = document.getElementById('results-body-nuevos');
const btnDownloadNuevos = document.getElementById('btn-download-nuevos');

const dropZoneMlNuevos = document.getElementById('drop-zone-ml-nuevos');
const dropZoneSysNuevos = document.getElementById('drop-zone-sys-nuevos');

const statsNuevos = {
    mlMayor: document.getElementById('stat-ml-mayor-nuevos'),
    missingMl: document.getElementById('stat-missing-ml-nuevos'),
    warning: document.getElementById('stat-warning-nuevos'),
    other: document.getElementById('stat-other-nuevos')
};

let dataMlNuevos = null;
let dataSysNuevos = null;
let finalResultsNuevos = [];
let mlArrayBufferNuevos = null; // bytes originales de la planilla de ML (para re-subir)
let mlSheetNameNuevos = null;   // hoja de datos de esa planilla ("Publicaciones")

function setupDropZoneNuevos(dropZone, fileInput, statusElement, type) {
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            handleFileSelectNuevos(fileInput, statusElement, type);
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length === 0) return;
        handleFileSelectNuevos(fileInput, statusElement, type);
    });
}

setupDropZoneNuevos(dropZoneMlNuevos, fileMlNuevos, statusMlNuevos, 'ml');
setupDropZoneNuevos(dropZoneSysNuevos, fileSysNuevos, statusSysNuevos, 'sys');

function handleFileSelectNuevos(input, statusElement, type) {
    if (input.files.length === 0) return;
    const file = input.files[0];
    statusElement.textContent = file.name;
    statusElement.classList.add('uploaded');
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            let targetSheetName = workbook.SheetNames[0];
            if (type === 'ml') {
                if (workbook.SheetNames.includes('Publicaciones')) {
                    targetSheetName = 'Publicaciones';
                } else if (workbook.SheetNames.length > 2) {
                    targetSheetName = workbook.SheetNames[2];
                } else if (workbook.SheetNames.length > 1) {
                    targetSheetName = workbook.SheetNames[1];
                }
            }
            
            const worksheet = workbook.Sheets[targetSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            if (type === 'ml') {
                dataMlNuevos = parseMl(json);
                // Se guarda el archivo original en crudo para poder devolver la planilla
                // de ML re-subible. Se re-lee limpio en cada descarga (sin arrastrar ediciones).
                mlArrayBufferNuevos = e.target.result;
                mlSheetNameNuevos = targetSheetName;
            } else if (type === 'sys') {
                dataSysNuevos = parseSysNuevos(json);
            }

            checkReadyNuevos();
        } catch (error) {
            Swal.fire({ title: 'Error', html: error.userMessage || 'No se pudo leer el archivo Excel/CSV.', icon: 'error' });
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function checkReadyNuevos() {
    if (dataMlNuevos && dataSysNuevos) {
        btnProcessNuevos.disabled = false;
    }
}

btnProcessNuevos.addEventListener('click', () => {
    btnProcessNuevos.disabled = true;
    btnProcessNuevos.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Procesando...';
    
    setTimeout(() => {
        analizarDatosNuevos();
        renderTableNuevos();
        
        resultsPanelNuevos.classList.remove('hidden');
        
        btnProcessNuevos.innerHTML = '<i class="fa-solid fa-bolt"></i> Analizar y Comparar Stock';
        btnProcessNuevos.disabled = false;
        
        Swal.fire({
            title: '¡Análisis Completado!',
            text: `Se encontraron ${finalResultsNuevos.length} inconsistencias.`,
            icon: 'success',
            confirmButtonColor: '#3b82f6'
        });
    }, 500);
});

function analizarDatosNuevos() {
    finalResultsNuevos = [];
    
    let cMlMayor = 0;
    let cMissingMl = 0;
    let cWarning = 0;
    let cOther = 0;
    
    const combinedSys = dataSysNuevos || {};
    
    // Compare ML against System
    for (const rawSku in dataMlNuevos) {
        let originalMlStock = dataMlNuevos[rawSku].stock;
        let title = dataMlNuevos[rawSku].title;
        let skuForSystem = rawSku;
        let multiplier = 1;
        
        const kitMatchPrefix = rawSku.match(/^KITX(\d+)-(.*)/);
        const kitMatchSuffix = rawSku.match(/^(.*)X(\d+)$/);
        
        if (kitMatchPrefix) {
            multiplier = parseInt(kitMatchPrefix[1], 10);
            skuForSystem = kitMatchPrefix[2];
        } else if (kitMatchSuffix) {
            multiplier = parseInt(kitMatchSuffix[2], 10);
            skuForSystem = kitMatchSuffix[1];
        }
        
        const mlStockEq = originalMlStock * multiplier;
        
        if (combinedSys.hasOwnProperty(skuForSystem)) {
            const sysItem = combinedSys[skuForSystem];
            const sysStock = sysItem.stock;
            const sysDesc = sysItem.descripcion;

            if (mlStockEq !== sysStock) {
                const diff = mlStockEq - sysStock;
                let motivo = '';
                let badgeClass = '';
                
                if (mlStockEq > sysStock && mlStockEq < 10 && sysStock < 10) {
                    motivo = 'Stock ML mayor que Sistema (Riesgo)';
                    badgeClass = 'bg-danger';
                    cMlMayor++;
                } else if (Math.abs(diff) <= 2 && mlStockEq < 8 && sysStock < 8) {
                    motivo = 'Diferencia de 2 unidades o menos (Advertencia)';
                    badgeClass = 'bg-warning';
                    cWarning++;
                } else if (mlStockEq < 4 && sysStock > 6) {
                    motivo = 'Actualizar ML (Sistema tiene más stock)';
                    badgeClass = 'bg-info';
                    cOther++;
                } else {
                    continue;
                }
                
                finalResultsNuevos.push({
                    SKU: rawSku,
                    Title: title || sysDesc || 'Sin detalle',
                    'Stock ML': originalMlStock,
                    'Multiplier': multiplier,
                    'Stock Sistema': sysStock,
                    Diferencia: diff,
                    Motivo: motivo,
                    _badgeClass: badgeClass
                });
            }
        }
    }
    
    // Compare System against ML (for missing items)
    for (const sku in combinedSys) {
        const sysItem = combinedSys[sku];
        const sysStock = sysItem.stock;
        const sysDesc = sysItem.descripcion;

        if (!dataMlNuevos.hasOwnProperty(sku) && sysStock > 0) {
            cMissingMl++;
            finalResultsNuevos.push({
                SKU: sku,
                Title: sysDesc || 'Sin detalle',
                'Stock ML': 'No existe',
                'Stock Sistema': sysStock,
                Diferencia: -sysStock,
                Motivo: 'SKU no existe en ML (Publicar/Activar)',
                _badgeClass: 'bg-success'
            });
        }
    }
    
    statsNuevos.mlMayor.textContent = cMlMayor;
    statsNuevos.missingMl.textContent = cMissingMl;
    statsNuevos.warning.textContent = cWarning;
    statsNuevos.other.textContent = cOther;
}

function renderTableNuevos() {
    resultsBodyNuevos.innerHTML = '';
    
    finalResultsNuevos.sort((a, b) => {
        return Math.abs(b.Diferencia) - Math.abs(a.Diferencia);
    });
    
    finalResultsNuevos.forEach(item => {
        const tr = document.createElement('tr');
        
        const mlDisplay = item.Multiplier > 1 
            ? `${item['Stock ML']} <span style="font-size: 0.8rem; color: var(--warning);">(x${item.Multiplier})</span>`
            : item['Stock ML'];
            
        const enML = item['Stock ML'] !== 'No existe';
        const stockInput = enML
            ? `<div class="ml-stock-field" title="Stock que se subirá a Mercado Libre para este SKU. Dejalo vacío si no querés cambiarlo.">
                    <i class="fa-solid fa-arrow-right-to-bracket"></i>
                    <input type="number" class="ml-stock-input" data-sku="${item.SKU}" placeholder="${item['Stock Sistema']}" min="0" step="1" inputmode="numeric">
               </div>`
            : `<div class="ml-stock-field ml-stock-na" title="No existe como publicación en ML: no se puede actualizar por planilla.">—</div>`;

        tr.innerHTML = `
            <td>
                <span class="sku-cell">
                    <strong>${item.SKU}</strong>
                    <button type="button" class="copy-sku-btn" data-sku="${item.SKU}" title="Copiar SKU" aria-label="Copiar SKU"><i class="fa-regular fa-copy"></i></button>
                </span>
                ${stockInput}
            </td>
            <td><span style="font-size: 0.85rem; color: var(--text-secondary);">${item.Title || 'Sin detalle'}</span></td>
            <td>${mlDisplay}</td>
            <td>${item['Stock Sistema']}</td>
            <td><span class="status-badge ${item._badgeClass}">${item.Motivo}</span></td>
        `;

        resultsBodyNuevos.appendChild(tr);
    });
}

// Botón "copiar SKU" en la tabla de auditoría (delegación de eventos).
resultsBodyNuevos.addEventListener('click', (e) => {
    const btn = e.target.closest('.copy-sku-btn');
    if (btn) copySku(btn.dataset.sku, btn);
});

btnDownloadNuevos.addEventListener('click', () => {
    const excelData = finalResultsNuevos.map(item => {
        return {
            'SKU': item.SKU,
            'Título de Publicación': item.Title || 'Sin detalle',
            'Stock Mercado Libre': item['Stock ML'],
            'Stock Sistema': item['Stock Sistema'],
            'Motivo de Corrección': item.Motivo
        };
    });
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    
    const wscols = [
        {wch: 20}, // SKU
        {wch: 45}, // Título
        {wch: 20}, // ML
        {wch: 20}, // Sys
        {wch: 50}  // Motivo
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Correcciones Nuevos");
    
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Conciliacion_Nuevos_${dateStr}.xlsx`);
});

// --- DESCARGAR PLANILLA DE ML CON EL STOCK ACTUALIZADO ---
// Edita los BYTES internos del .xlsx original (es un zip) y cambia únicamente las celdas de stock
// (columna STOCK_FLEX / "En mi depósito") de los SKU que el usuario completó. NO se regenera el
// archivo con la librería: así ML recibe su plantilla idéntica (hojas ocultas, imágenes, estilos,
// validaciones) con solo los números tocados, y la acepta al re-subir. Las publicaciones que no
// tocaste quedan con su stock original, así ML no las modifica.
function normSkuMl(s) {
    return String(s == null ? '' : s).replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
}

// Ubica el XML de una hoja por su nombre, dentro del zip ya descomprimido (map name -> path).
function rutaHojaXml(zipObj, sheetName) {
    const dec = new TextDecoder();
    const wbXml = zipObj['xl/workbook.xml'] ? dec.decode(zipObj['xl/workbook.xml']) : '';
    const relsXml = zipObj['xl/_rels/workbook.xml.rels'] ? dec.decode(zipObj['xl/_rels/workbook.xml.rels']) : '';
    if (!wbXml || !relsXml) return null;
    const unesc = s => String(s).replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
    let rid = null;
    const tags = wbXml.match(/<sheet\b[^>]*>/g) || [];
    for (const t of tags) {
        const nm = (t.match(/name="([^"]*)"/) || [])[1];
        const id = (t.match(/r:id="([^"]*)"/) || [])[1];
        if (nm != null && unesc(nm) === sheetName) { rid = id; break; }
    }
    if (!rid) return null;
    const rel = relsXml.match(new RegExp('<Relationship\\b[^>]*Id="' + rid + '"[^>]*>'));
    if (!rel) return null;
    let target = (rel[0].match(/Target="([^"]*)"/) || [])[1];
    if (!target) return null;
    target = target.replace(/^\//, '');
    if (!target.startsWith('xl/')) target = 'xl/' + target;
    return target;
}

const btnDownloadMlPlanilla = document.getElementById('btn-download-ml-planilla');
if (btnDownloadMlPlanilla) btnDownloadMlPlanilla.addEventListener('click', () => {
    if (!mlArrayBufferNuevos) {
        Swal.fire('Falta la planilla de ML', 'Primero subí y analizá el Excel de Mercado Libre.', 'info');
        return;
    }
    if (typeof fflate === 'undefined') {
        Swal.fire('Error', 'No se pudo cargar la librería para editar el Excel (fflate). Revisá tu conexión y recargá.', 'error');
        return;
    }

    // Junta los valores que el usuario cargó en los casilleros.
    const updates = new Map(); // SKU normalizado -> stock (número)
    document.querySelectorAll('#results-body-nuevos .ml-stock-input').forEach(inp => {
        const v = String(inp.value).trim();
        if (v === '') return;
        const n = parseInt(v, 10);
        if (isNaN(n) || n < 0) return;
        updates.set(normSkuMl(inp.dataset.sku), n);
    });

    if (updates.size === 0) {
        Swal.fire('Sin cambios', 'Completá el stock de al menos un SKU (los que están en rojo o azul) antes de descargar.', 'info');
        return;
    }

    // 1) Con la librería solo LOCALIZAMOS qué celdas de stock hay que cambiar (no reescribimos nada).
    const wb = XLSX.read(mlArrayBufferNuevos, { type: 'array' });
    const sheetName = (mlSheetNameNuevos && wb.Sheets[mlSheetNameNuevos])
        ? mlSheetNameNuevos
        : (wb.SheetNames.includes('Publicaciones') ? 'Publicaciones' : wb.SheetNames[wb.SheetNames.length - 1]);
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) {
        Swal.fire('Error', 'No se pudo leer la hoja de publicaciones de la planilla de ML.', 'error');
        return;
    }

    const range = XLSX.utils.decode_range(ws['!ref']);
    let skuCol = 4, stockCol = 7; // defaults E y H
    for (let r = range.s.r; r <= Math.min(range.s.r + 6, range.e.r); r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = ws[XLSX.utils.encode_cell({ r, c })];
            const val = cell && cell.v != null ? String(cell.v).trim().toUpperCase() : '';
            if (val === 'SKU') skuCol = c;
            if (val === 'STOCK_FLEX' || val === 'EN MI DEPOSITO' || val === 'EN MI DEPÓSITO') stockCol = c;
        }
    }
    // Planilla con stock por depósito: se escribe en el de Coronel Gil (nunca en Asicha u otro).
    try {
        const colGil = colStockGilMl(XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }));
        if (colGil >= 0) stockCol = colGil;
    } catch (err) {
        Swal.fire({ title: 'Error', html: err.userMessage || err.message, icon: 'error' });
        return;
    }

    const cellEdits = {}; // "H270" -> 5
    let changed = 0;
    const found = new Set();
    for (let r = range.s.r; r <= range.e.r; r++) {
        const skuCell = ws[XLSX.utils.encode_cell({ r, c: skuCol })];
        if (!skuCell || skuCell.v == null) continue;
        const key = normSkuMl(skuCell.v);
        if (!updates.has(key)) continue;
        cellEdits[XLSX.utils.encode_cell({ r, c: stockCol })] = updates.get(key);
        changed++;
        found.add(key);
    }

    const notFound = [];
    updates.forEach((_, k) => { if (!found.has(k)) notFound.push(k); });

    if (changed === 0) {
        Swal.fire('Sin coincidencias', 'Ninguno de los SKU que completaste aparece en la planilla de ML.' + (notFound.length ? '<br><br>' + notFound.join(', ') : ''), 'warning');
        return;
    }

    // 2) Editamos los bytes: descomprimimos el zip, tocamos solo el XML de la hoja, y re-comprimimos.
    let outBytes;
    try {
        const zipObj = fflate.unzipSync(new Uint8Array(mlArrayBufferNuevos));
        const path = rutaHojaXml(zipObj, sheetName) || 'xl/worksheets/sheet' + (wb.SheetNames.indexOf(sheetName) + 1) + '.xml';
        if (!zipObj[path]) throw new Error('No se encontró el XML de la hoja ' + sheetName);

        const dec = new TextDecoder();
        const enc = new TextEncoder();
        let xml = dec.decode(zipObj[path]);

        const cellXml = (addr, attrs, v) => {
            const a = String(attrs || '').replace(/\s*t="[^"]*"/, ''); // se fuerza tipo numérico
            return '<c r="' + addr + '"' + a + ' t="n"><v>' + v + '</v></c>';
        };

        Object.keys(cellEdits).forEach(addr => {
            const v = cellEdits[addr];
            const selfRe = new RegExp('<c r="' + addr + '"([^>]*?)/>');
            if (selfRe.test(xml)) {
                xml = xml.replace(selfRe, (m, a) => cellXml(addr, a, v));
            } else {
                const fullRe = new RegExp('<c r="' + addr + '"([^>]*?)>[\\s\\S]*?</c>');
                xml = xml.replace(fullRe, (m, a) => cellXml(addr, a, v));
            }
        });

        zipObj[path] = enc.encode(xml);
        outBytes = fflate.zipSync(zipObj);
    } catch (err) {
        console.error('Error editando la planilla de ML', err);
        Swal.fire('Error', 'No se pudo generar la planilla: ' + err.message, 'error');
        return;
    }

    const blob = new Blob([outBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const dateStr = new Date().toISOString().split('T')[0];
    a.download = `Planilla_ML_stock_${dateStr}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    Swal.fire({
        title: '¡Planilla lista!',
        html: `Se actualizó el stock de <strong>${changed}</strong> publicación/es en la planilla de ML (los SKU que editaste).` +
              `<br>El resto queda con su stock actual, así ML no lo modifica.` +
              (notFound.length ? `<br><br><span style="color:#f59e0b;">No se encontraron en la planilla:</span> ${notFound.join(', ')}` : '') +
              `<br><br>Subila a Mercado Libre tal cual (Modificar publicaciones → subir Excel).`,
        icon: 'success',
        confirmButtonColor: '#3b82f6'
    });
});

// --- DYNAMIC PARSER FOR MERCADO LIBRE FULL STOCK ---
function parseMlFullStock(rows) {
    const map = {};
    
    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
        const row = rows[i];
        if (row && row.some(cell => typeof cell === 'string' && cell.trim().toUpperCase() === 'SKU')) {
            headerRowIdx = i;
            break;
        }
    }
    
    if (headerRowIdx === -1) {
        throw new Error("No se encontró la fila con el campo 'SKU' en el archivo de Full.");
    }

    const firstHeaderRow = rows[headerRowIdx].map(h => String(h).trim().toUpperCase());
    const idxSku = firstHeaderRow.indexOf('SKU');

    // El nuevo formato de Mercado Libre (hoja "Resumen") reparte el encabezado
    // real en varias filas: la fila principal + subtítulos por grupo de columnas
    // (ej. "Unidades en Full" / "Aptas para vender"). Mientras la columna de SKU
    // venga vacía seguimos acumulando esas filas al encabezado combinado.
    const headerRows = [rows[headerRowIdx]];
    let dataStartIdx = headerRowIdx + 1;
    while (dataStartIdx < rows.length && headerRows.length < 5) {
        const row = rows[dataStartIdx];
        const skuCell = row ? row[idxSku] : undefined;
        const isBlankSku = skuCell === undefined || skuCell === null || String(skuCell).trim() === '';
        if (!isBlankSku) break;
        headerRows.push(row);
        dataStartIdx++;
    }

    const numCols = Math.max(...headerRows.map(r => r.length));
    const headers = [];
    for (let c = 0; c < numCols; c++) {
        headers.push(headerRows.map(r => (r[c] !== undefined && r[c] !== null) ? String(r[c]).trim() : '').join(' ').trim().toUpperCase());
    }

    // Busca por etapas: primero frases específicas (evita falsos positivos como
    // "Publicaciones con este PRODUCTO" o "...más vendidos EN FULL...") y sólo
    // si no aparecen recurre a coincidencias más genéricas.
    const findHeaderIdx = (stages) => {
        for (const stage of stages) {
            const idx = headers.findIndex(h => stage.exact ? h === stage.value : h.includes(stage.value));
            if (idx !== -1) return idx;
        }
        return -1;
    };

    const idxTitle = findHeaderIdx([
        { value: 'PRODUCTO', exact: true },
        { value: 'TÍTULO' },
        { value: 'TITLE' },
        { value: 'PRODUCTO' }
    ]);
    const idxPending = findHeaderIdx([
        { value: 'PENDIENTES DE INGRESO' },
        { value: 'PENDIENTES' },
        { value: 'CAMINO' }
    ]);
    const idxFull = findHeaderIdx([
        { value: 'APTAS PARA VENDER' },
        { value: 'APTAS' },
        { value: 'EN FULL' }
    ]);
    const idxSales = findHeaderIdx([
        { value: 'VENDIDAS' },
        { value: 'VENTAS ÚLTIMOS 30 DÍAS (U.)' },
        { value: 'VENTAS ÚLTIMOS 30 DÍAS' },
        { value: 'VENTAS (U.)' }
    ]);

    if (idxSku === -1 || idxFull === -1) {
        throw new Error("El archivo de Full no tiene las columnas requeridas (SKU y Unidades en Full).");
    }

    for (let i = dataStartIdx; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;
        
        let skuVal = row[idxSku];
        if (skuVal === undefined || skuVal === null || skuVal === '') continue;
        
        let sku = skuVal.toString().trim().toUpperCase();
        if (sku === "SKU" || sku.includes("PENDIENTES DE INGRESO") || sku.includes("ESTADO DE TU STOCK")) continue;
        
        const title = idxTitle !== -1 && row[idxTitle] ? String(row[idxTitle]).trim() : 'Neumático Full';
        const pendingVal = idxPending !== -1 ? parseInt(row[idxPending], 10) || 0 : 0;
        const fullVal = idxFull !== -1 ? parseInt(row[idxFull], 10) || 0 : 0;
        const salesVal = idxSales !== -1 ? parseInt(row[idxSales], 10) || 0 : 0;
        
        if (map.hasOwnProperty(sku)) {
            map[sku].aptas += fullVal;
            map[sku].pendientes += pendingVal;
            map[sku].ventas30 += salesVal;
        } else {
            map[sku] = {
                sku: sku,
                title: title,
                aptas: fullVal,
                pendientes: pendingVal,
                ventas30: salesVal
            };
        }
    }
    return map;
}

// --- LOGICA DE SUGERENCIA DE ENVIOS A MERCADO LIBRE FULL ---
const fileMlFullStock = document.getElementById('file-ml-full-stock');
const fileSysRestock = document.getElementById('file-sys-restock');
const statusMlFullStock = document.getElementById('status-ml-full-stock');
const statusSysRestock = document.getElementById('status-sys-restock');
const btnProcessRestock = document.getElementById('btn-process-restock');
const resultsPanelRestock = document.getElementById('results-panel-restock');
const resultsBodyRestock = document.getElementById('results-body-restock');
const btnDownloadRestock = document.getElementById('btn-download-restock');

const dropZoneMlFullStock = document.getElementById('drop-zone-ml-full-stock');
const dropZoneSysRestock = document.getElementById('drop-zone-sys-restock');

const inputRestockTarget = document.getElementById('restock-target');
const checkboxRestockUseSales = document.getElementById('restock-use-sales');

const statsRestock = {
    models: document.getElementById('stat-restock-models'),
    total: document.getElementById('stat-restock-total'),
    gil: document.getElementById('stat-restock-gil'),
    marzo: document.getElementById('stat-restock-marzo'),
    marzo1435: document.getElementById('stat-restock-marzo1435')
};

let dataMlFullStock = null;
let dataSysRestock = null;
let finalResultsRestock = [];

// Ventas AR de Mercado Libre (opcional): si se carga, reemplaza la columna de ventas del Excel de Full.
const fileMlFullVentas = document.getElementById('file-ml-full-ventas');
const statusMlFullVentas = document.getElementById('status-ml-full-ventas');
const dropZoneMlFullVentas = document.getElementById('drop-zone-ml-full-ventas');
const inputRestockSalesDays = document.getElementById('restock-sales-days');
const thRestockVentas = document.getElementById('th-restock-ventas');
let dataVentasFull = null;          // { rows: [{sku, units, ts}], maxTs, minTs, descartadas, fileName }
let ventasFullLabel = 'Ventas (30 días)';

const MESES_ES_FULL = { enero: 0, febrero: 1, marzo: 2, abril: 3, mayo: 4, junio: 5, julio: 6,
    agosto: 7, septiembre: 8, setiembre: 8, octubre: 9, noviembre: 10, diciembre: 11 };

// "3 de noviembre de 2026 11:01 hs." -> timestamp. Acepta también fechas de Excel (serial/Date).
function parseFechaVentaML(v) {
    if (v instanceof Date) return v.getTime();
    if (typeof v === 'number' && v > 20000) return Math.round((v - 25569) * 86400000);
    const m = String(v || '').toLowerCase().match(/(\d{1,2})\s+de\s+([a-záéíóú]+)\s+de\s+(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (!m || !(m[2] in MESES_ES_FULL)) return null;
    return new Date(+m[3], MESES_ES_FULL[m[2]], +m[1], +(m[4] || 0), +(m[5] || 0)).getTime();
}

// Lee el export "Ventas AR". Las columnas se detectan por nombre (ML cambia el layout entre descargas).
// Se descartan ventas canceladas/devueltas para no inflar la demanda con ventas que no se concretaron.
function parseVentasFullAR(rows) {
    let h = -1;
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
        const cols = (rows[i] || []).map(c => String(c === undefined || c === null ? '' : c).trim().toLowerCase());
        if (cols.includes('sku') && cols.includes('unidades')) { h = i; break; }
    }
    if (h === -1) {
        throw new Error("El Excel no tiene las columnas 'SKU' y 'Unidades'. Subí el reporte 'Ventas AR' de Mercado Libre.");
    }
    const cols = rows[h].map(c => String(c === undefined || c === null ? '' : c).trim().toLowerCase());
    const iSku = cols.indexOf('sku');
    const iUni = cols.indexOf('unidades');
    const iEstado = cols.indexOf('estado');
    const iFecha = cols.findIndex(c => c.startsWith('fecha de venta'));

    const out = { rows: [], maxTs: null, minTs: null, descartadas: 0 };
    for (let i = h + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;
        const sku = String(row[iSku] === undefined || row[iSku] === null ? '' : row[iSku]).trim().toUpperCase();
        if (!sku) continue;
        const estado = iEstado !== -1 ? String(row[iEstado] || '') : '';
        if (/cancela|devoluci/i.test(estado)) { out.descartadas++; continue; }
        const units = parseFloat(row[iUni]) || 0;
        if (units <= 0) continue;
        const ts = iFecha !== -1 ? parseFechaVentaML(row[iFecha]) : null;
        if (ts !== null) {
            if (out.maxTs === null || ts > out.maxTs) out.maxTs = ts;
            if (out.minTs === null || ts < out.minTs) out.minTs = ts;
        }
        out.rows.push({ sku, units, ts });
    }
    return out;
}

function diasVentasFull() {
    return Math.max(1, parseInt(inputRestockSalesDays.value, 10) || 30);
}

// Unidades por SKU dentro de los últimos N días (contados desde la venta más reciente del archivo).
function ventasPorSkuFull(dias) {
    if (!dataVentasFull) return null;
    const cutoff = dataVentasFull.maxTs !== null ? dataVentasFull.maxTs - dias * 86400000 : null;
    const map = {};
    dataVentasFull.rows.forEach(r => {
        if (cutoff !== null && r.ts !== null && r.ts < cutoff) return;
        map[r.sku] = (map[r.sku] || 0) + r.units;
    });
    return map;
}

function refrescarEstadoVentasFull() {
    if (!dataVentasFull) return;
    const dias = diasVentasFull();
    const map = ventasPorSkuFull(dias);
    const skus = Object.keys(map).length;
    const uds = Object.values(map).reduce((a, b) => a + b, 0);
    const fmtF = (ts) => new Date(ts).toLocaleDateString('es-AR');
    let rango = '';
    if (dataVentasFull.maxTs !== null) {
        const desde = Math.max(dataVentasFull.minTs, dataVentasFull.maxTs - dias * 86400000);
        rango = ` · ${fmtF(desde)} al ${fmtF(dataVentasFull.maxTs)}`;
    }
    statusMlFullVentas.textContent = `✓ ${dataVentasFull.fileName} — ${skus} SKUs, ${Math.round(uds)} uds. en ${dias} días${rango}`
        + (dataVentasFull.descartadas ? ` (${dataVentasFull.descartadas} canceladas/devueltas descartadas)` : '');
    statusMlFullVentas.classList.add('uploaded');
}

function handleVentasFullFile(input) {
    const file = input.files[0];
    if (!file) return;
    statusMlFullVentas.textContent = 'Procesando...';
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
            dataVentasFull = parseVentasFullAR(rows);
            dataVentasFull.fileName = file.name;
            refrescarEstadoVentasFull();
        } catch (error) {
            dataVentasFull = null;
            statusMlFullVentas.textContent = 'Error al leer';
            statusMlFullVentas.classList.remove('uploaded');
            Swal.fire('Error', 'No se pudo leer el Excel de ventas: ' + error.message, 'error');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

dropZoneMlFullVentas.addEventListener('click', () => fileMlFullVentas.click());
dropZoneMlFullVentas.addEventListener('dragover', (e) => { e.preventDefault(); dropZoneMlFullVentas.classList.add('drag-over'); });
dropZoneMlFullVentas.addEventListener('dragleave', () => dropZoneMlFullVentas.classList.remove('drag-over'));
dropZoneMlFullVentas.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZoneMlFullVentas.classList.remove('drag-over');
    if (e.dataTransfer.files.length) {
        fileMlFullVentas.files = e.dataTransfer.files;
        handleVentasFullFile(fileMlFullVentas);
    }
});
fileMlFullVentas.addEventListener('change', () => handleVentasFullFile(fileMlFullVentas));
inputRestockSalesDays.addEventListener('input', refrescarEstadoVentasFull);

function setupDropZoneRestock(dropZone, fileInput, statusElement, type) {
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            handleFileSelectRestock(fileInput, statusElement, type);
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length === 0) return;
        handleFileSelectRestock(fileInput, statusElement, type);
    });
}

setupDropZoneRestock(dropZoneMlFullStock, fileMlFullStock, statusMlFullStock, 'ml-full');
setupDropZoneRestock(dropZoneSysRestock, fileSysRestock, statusSysRestock, 'sys-restock');

function handleFileSelectRestock(input, statusElement, type) {
    if (input.files.length === 0) return;
    
    const file = input.files[0];
    statusElement.textContent = file.name;
    statusElement.classList.add('uploaded');
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            let targetSheetName = workbook.SheetNames[0];
            if (type === 'ml-full') {
                if (workbook.SheetNames.includes('Resumen')) {
                    targetSheetName = 'Resumen';
                }
            }
            
            const worksheet = workbook.Sheets[targetSheetName];
            
            if (type === 'ml-full' && worksheet && worksheet['!ref']) {
                const r = XLSX.utils.decode_range(worksheet['!ref']);
                r.s.r = 0;
                r.s.c = 0;
                worksheet['!ref'] = XLSX.utils.encode_range(r);
            }
            
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            if (type === 'ml-full') {
                dataMlFullStock = parseMlFullStock(json);
            } else if (type === 'sys-restock') {
                dataSysRestock = parseSysNuevos(json);
            }
            
            checkReadyRestock();
        } catch (error) {
            Swal.fire('Error', 'No se pudo leer el archivo Excel/CSV: ' + error.message, 'error');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function checkReadyRestock() {
    if (dataMlFullStock && dataSysRestock) {
        btnProcessRestock.disabled = false;
    }
}

btnProcessRestock.addEventListener('click', () => {
    btnProcessRestock.disabled = true;
    btnProcessRestock.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Procesando...';
    
    setTimeout(() => {
        calcularEnvíosRestock();
        renderTableRestock();
        
        resultsPanelRestock.classList.remove('hidden');
        
        btnProcessRestock.innerHTML = 'Planificar y Calcular Envíos';
        btnProcessRestock.disabled = false;
        
        Swal.fire({
            title: '¡Planificación Completada!',
            text: `Se generaron sugerencias para ${finalResultsRestock.length} modelos.`,
            icon: 'success',
            confirmButtonColor: '#3b82f6'
        });
    }, 500);
});

// Extrae la medida de cubierta de una descripción/título.
// Soporta "195/60R15", "215 65 R16", "175/65R14", etc. Devuelve "195/60R15" o null.
function medidaDeDescRestock(desc) {
    const d = String(desc || '').toUpperCase();
    const m = d.match(/(\d{3})\s*[\/\s]\s*(\d{2})\s*(?:Z\s*)?R?\s*(\d{2})/);
    return m ? `${m[1]}/${m[2]}R${m[3]}` : null;
}

function calcularEnvíosRestock() {
    finalResultsRestock = [];

    const targetBase = parseInt(inputRestockTarget.value, 10) || 4;
    const useSales = checkboxRestockUseSales.checked;
    
    // Obtener la unión de todas las SKUs de Full y del Sistema
    const allSkus = new Set([
        ...Object.keys(dataMlFullStock),
        ...Object.keys(dataSysRestock)
    ]);
    
    let countModels = 0;

    // Si se cargó el Excel Ventas AR, las ventas salen de ahí (últimos N días) y no del Excel de Full.
    const diasVentas = diasVentasFull();
    const ventasAR = ventasPorSkuFull(diasVentas);
    ventasFullLabel = ventasAR ? `Ventas (${diasVentas} días, ML)` : 'Ventas (30 días)';
    if (thRestockVentas) thRestockVentas.textContent = ventasAR ? `Ventas (${diasVentas}d · ML)` : 'Ventas (30d)';

    allSkus.forEach(sku => {
        const mlItem = dataMlFullStock[sku] || { aptas: 0, pendientes: 0, ventas30: 0, title: '' };
        const mlAptas = mlItem.aptas;
        const mlPendientes = mlItem.pendientes;
        const mlVentas = ventasAR ? (ventasAR[String(sku).trim().toUpperCase()] || 0) : mlItem.ventas30;
        const mlTitle = mlItem.title;
        
        const mlTotal = mlAptas + mlPendientes;
        
        // Determinar stock objetivo
        const target = useSales ? Math.max(targetBase, mlVentas) : targetBase;
        const deficit = Math.max(0, target - mlTotal);
        
        let skuForSystem = sku;
        let multiplier = 1;
        
        const kitMatchPrefix = sku.match(/^KITX(\d+)-(.*)/);
        const kitMatchSuffix = sku.match(/^(.*)X(\d+)$/);
        
        if (kitMatchPrefix) {
            multiplier = parseInt(kitMatchPrefix[1], 10);
            skuForSystem = kitMatchPrefix[2];
        } else if (kitMatchSuffix) {
            multiplier = parseInt(kitMatchSuffix[2], 10);
            skuForSystem = kitMatchSuffix[1];
        }
        
        let sysStockGil = 0;
        let sysStockMarzo = 0;
        let sysStockMarzo1435 = 0;
        let sysDesc = '';
        
        if (dataSysRestock.hasOwnProperty(skuForSystem)) {
            const sysItem = dataSysRestock[skuForSystem];
            sysStockGil = sysItem.gilStock || 0;
            sysStockMarzo = sysItem.marzoStock || 0;
            sysStockMarzo1435 = sysItem.marzo1435Stock || 0;
            sysDesc = sysItem.descripcion || '';
        }
        
        // Ajustar stock local para kits
        const availGilKits = Math.floor(sysStockGil / multiplier);
        const availMarzoKits = Math.floor(sysStockMarzo / multiplier);
        const availMarzo1435Kits = Math.floor(sysStockMarzo1435 / multiplier);
        const availTotalKits = availGilKits + availMarzoKits + availMarzo1435Kits;
        
        const toSendKits = Math.min(deficit, availTotalKits);
        
        let sendGil = 0;
        let sendMarzo = 0;
        let sendMarzo1435 = 0;
        let sugerencia = '';
        let badgeClass = '';
        
        if (toSendKits === 0) {
            if (deficit === 0) {
                sugerencia = "Sin déficit (Suficiente stock)";
                badgeClass = "bg-info";
            } else {
                sugerencia = "Sin stock local disponible";
                badgeClass = "bg-danger";
            }
        } else {
            badgeClass = "bg-success";
            
            if (availGilKits >= toSendKits) {
                sendGil = toSendKits;
                sendMarzo = 0;
                sendMarzo1435 = 0;
            } else if (availMarzoKits >= toSendKits) {
                sendGil = 0;
                sendMarzo = toSendKits;
                sendMarzo1435 = 0;
            } else if (availMarzo1435Kits >= toSendKits) {
                sendGil = 0;
                sendMarzo = 0;
                sendMarzo1435 = toSendKits;
            } else {
                // Distribute: Gil first, then Marzo, then Marzo 1435
                sendGil = availGilKits;
                let remaining = toSendKits - sendGil;
                if (availMarzoKits >= remaining) {
                    sendMarzo = remaining;
                    sendMarzo1435 = 0;
                } else {
                    sendMarzo = availMarzoKits;
                    remaining -= sendMarzo;
                    sendMarzo1435 = Math.min(availMarzo1435Kits, remaining);
                }
            }
            
            const unitName = multiplier > 1 ? 'kit(s)' : 'u.';
            const parts = [];
            if (sendGil > 0) parts.push(`${sendGil} de Gil`);
            if (sendMarzo > 0) parts.push(`${sendMarzo} de Santiago Marzo`);
            if (sendMarzo1435 > 0) parts.push(`${sendMarzo1435} de Santiago Marzo 1435`);
            
            sugerencia = `Enviar ${toSendKits} ${unitName} (Mandar ${parts.join(' y ')})`;
            
            if (toSendKits < deficit) {
                sugerencia += ` - Faltan ${deficit - toSendKits} ${unitName} por falta de stock`;
                badgeClass = "bg-warning";
            }
        }
        
        finalResultsRestock.push({
            SKU: sku,
            Title: mlTitle || sysDesc || 'Sin detalle',
            Ventas: mlVentas,
            Aptas: mlAptas,
            Pendientes: mlPendientes,
            LocalGil: sysStockGil,
            LocalMarzo: sysStockMarzo,
            LocalMarzo1435: sysStockMarzo1435,
            Multiplier: multiplier,
            Deficit: deficit,
            
            // Guardar valores recomendados originales como referencia
            RecGil: sendGil,
            RecMarzo: sendMarzo,
            RecMarzo1435: sendMarzo1435,
            RecSugerencia: sugerencia,
            RecBadgeClass: badgeClass,
            
            // Selección del usuario (comienza vacía/null)
            ToSend: null,
            SendGil: 0,
            SendMarzo: 0,
            SendMarzo1435: 0,
            Sugerencia: sugerencia,
            _badgeClass: badgeClass
        });
        countModels++;
    });

    // Stock combinado de los 3 depósitos por ítem, y total por medida de cubierta.
    const stockPorMedida = {};
    finalResultsRestock.forEach(item => {
        item.StockDepositos = item.LocalGil + item.LocalMarzo + item.LocalMarzo1435;
        item.Medida = medidaDeDescRestock(item.Title);
        if (item.Medida) {
            stockPorMedida[item.Medida] = (stockPorMedida[item.Medida] || 0) + item.StockDepositos;
        }
    });
    finalResultsRestock.forEach(item => {
        item.StockMedida = item.Medida ? stockPorMedida[item.Medida] : item.StockDepositos;
    });

    recalcularEstadisticasGlobalesRestock();
}
 
function renderTableRestock() {
    resultsBodyRestock.innerHTML = '';
    
    // Ordenar de mayor a menor según el stock local total
    finalResultsRestock.sort((a, b) => (b.LocalGil + b.LocalMarzo + b.LocalMarzo1435) - (a.LocalGil + a.LocalMarzo + a.LocalMarzo1435));
    
    finalResultsRestock.forEach(item => {
        const tr = document.createElement('tr');
        
        const localDisplay = `Gil: ${item.LocalGil} | Marzo: ${item.LocalMarzo} | 1435: ${item.LocalMarzo1435}`;
        const mlDisplay = item.Multiplier > 1 
            ? `${item.Aptas} <span style="font-size: 0.8rem; color: var(--text-secondary);">(x${item.Multiplier})</span>`
            : item.Aptas;
            
        // Mostrar vacío si ToSend es null
        const displayVal = item.ToSend === null ? '' : item.ToSend;
            
        tr.innerHTML = `
            <td><strong>${item.SKU}</strong></td>
            <td><span style="font-size: 0.85rem; color: var(--text-secondary);">${item.Title}</span></td>
            <td style="text-align: center;">${item.Ventas}</td>
            <td style="text-align: center;">${mlDisplay}</td>
            <td style="text-align: center;">${item.Pendientes}</td>
            <td style="text-align: center; font-size: 0.9rem; color: var(--text-secondary);">${localDisplay}</td>
            <td style="text-align: center;">
                <input type="number" min="0" value="${displayVal}" placeholder="0" class="restock-qty-input" data-sku="${item.SKU}" style="width: 70px; padding: 0.25rem; border-radius: 6px; border: 1px solid var(--border-color); background: var(--bg-primary); color: var(--text-primary); text-align: center; font-family: inherit; font-weight: bold; border-color: rgba(255,255,255,0.25);">
            </td>
            <td><span class="status-badge ${item._badgeClass}" id="badge-${item.SKU.replace(/[^a-zA-Z0-9]/g, '_')}">${item.Sugerencia}</span></td>
        `;
        
        resultsBodyRestock.appendChild(tr);
    });
}

// Registrar listener delegado para cambios manuales en la cantidad a enviar
resultsBodyRestock.addEventListener('input', (e) => {
    if (e.target.classList.contains('restock-qty-input')) {
        const sku = e.target.dataset.sku;
        const val = parseInt(e.target.value, 10);
        const newVal = isNaN(val) || val < 0 ? null : val;
        
        actualizarCantidadSugerida(sku, newVal);
    }
});

function actualizarCantidadSugerida(sku, newVal) {
    const item = finalResultsRestock.find(x => x.SKU === sku);
    if (!item) return;
    
    item.ToSend = newVal;
    
    const multiplier = item.Multiplier;
    const availGilKits = Math.floor(item.LocalGil / multiplier);
    const availMarzoKits = Math.floor(item.LocalMarzo / multiplier);
    const availMarzo1435Kits = Math.floor(item.LocalMarzo1435 / multiplier);
    const totalAvailKits = availGilKits + availMarzoKits + availMarzo1435Kits;
    
    let sendGil = 0;
    let sendMarzo = 0;
    let sendMarzo1435 = 0;
    let sugerencia = '';
    let badgeClass = '';
    
    if (newVal === null) {
        // Si está vacío, restaurar recomendación original
        sendGil = item.RecGil;
        sendMarzo = item.RecMarzo;
        sendMarzo1435 = item.RecMarzo1435;
        sugerencia = item.RecSugerencia;
        badgeClass = item.RecBadgeClass;
    } else if (newVal === 0) {
        sugerencia = "No enviar";
        badgeClass = "bg-info";
    } else {
        badgeClass = "bg-success";
        
        if (availGilKits >= newVal) {
            sendGil = newVal;
            sendMarzo = 0;
            sendMarzo1435 = 0;
        } else if (availMarzoKits >= newVal) {
            sendGil = 0;
            sendMarzo = newVal;
            sendMarzo1435 = 0;
        } else if (availMarzo1435Kits >= newVal) {
            sendGil = 0;
            sendMarzo = 0;
            sendMarzo1435 = newVal;
        } else {
            // Distribute: Gil first, then Marzo, then Marzo 1435
            sendGil = availGilKits;
            let remaining = newVal - sendGil;
            if (availMarzoKits >= remaining) {
                sendMarzo = remaining;
                sendMarzo1435 = 0;
            } else {
                sendMarzo = availMarzoKits;
                remaining -= sendMarzo;
                sendMarzo1435 = Math.min(availMarzo1435Kits, remaining);
            }
        }
        
        const unitName = multiplier > 1 ? 'kit(s)' : 'u.';
        const parts = [];
        if (sendGil > 0) parts.push(`${sendGil} de Gil`);
        if (sendMarzo > 0) parts.push(`${sendMarzo} de Santiago Marzo`);
        if (sendMarzo1435 > 0) parts.push(`${sendMarzo1435} de Santiago Marzo 1435`);
        
        if (parts.length > 0) {
            sugerencia = `Enviar ${newVal} ${unitName} (Mandar ${parts.join(' y ')})`;
        } else {
            sugerencia = `Enviar ${newVal} ${unitName} (Falta stock local)`;
        }
        
        const totalSentKits = sendGil + sendMarzo + sendMarzo1435;
        if (totalSentKits < newVal) {
            sugerencia += ` - Faltan ${newVal - totalSentKits} ${unitName} localmente`;
            badgeClass = "bg-warning";
        }
    }
    
    if (newVal === null) {
        item.SendGil = 0;
        item.SendMarzo = 0;
        item.SendMarzo1435 = 0;
        item.Sugerencia = item.RecSugerencia;
        item._badgeClass = item.RecBadgeClass;
    } else {
        item.SendGil = sendGil;
        item.SendMarzo = sendMarzo;
        item.SendMarzo1435 = sendMarzo1435;
        item.Sugerencia = sugerencia;
        item._badgeClass = badgeClass;
    }
    
    // Actualizar el DOM de la fila directamente
    const cleanSkuId = sku.replace(/[^a-zA-Z0-9]/g, '_');
    const badgeEl = document.getElementById(`badge-${cleanSkuId}`);
    if (badgeEl) {
        badgeEl.textContent = item.Sugerencia;
        badgeEl.className = `status-badge ${item._badgeClass}`;
    }
    
    // Recalcular estadísticas globales
    recalcularEstadisticasGlobalesRestock();
}

function recalcularEstadisticasGlobalesRestock() {
    let totalSend = 0;
    let totalGil = 0;
    let totalMarzo = 0;
    let totalMarzo1435 = 0;
    let countModels = 0;
    
    finalResultsRestock.forEach(item => {
        if (item.ToSend > 0) {
            totalSend += item.ToSend * item.Multiplier;
            totalGil += item.SendGil * item.Multiplier;
            totalMarzo += item.SendMarzo * item.Multiplier;
            totalMarzo1435 += item.SendMarzo1435 * item.Multiplier;
            countModels++;
        }
    });
    
    statsRestock.models.textContent = countModels;
    statsRestock.total.textContent = totalSend;
    statsRestock.gil.textContent = totalGil;
    statsRestock.marzo.textContent = totalMarzo;
    statsRestock.marzo1435.textContent = totalMarzo1435;
}

btnDownloadRestock.addEventListener('click', () => {
    const positiveResults = finalResultsRestock.filter(item => item.ToSend > 0);
    
    if (positiveResults.length === 0) {
        Swal.fire({
            title: 'Sin elementos',
            text: 'No has ingresado ninguna unidad a enviar todavía.',
            icon: 'warning',
            confirmButtonColor: '#3b82f6'
        });
        return;
    }
    
    const excelData = positiveResults.map(item => {
        return {
            'SKU': item.SKU,
            'Detalle del Producto': item.Title,
            'Medida': item.Medida || '',
            [ventasFullLabel]: item.Ventas,
            'Stock en Full (Apto)': item.Aptas,
            'En camino (Pendiente)': item.Pendientes,
            'Stock Depósitos (Gil+Marzo+1435)': item.StockDepositos,
            'Stock Total de la Medida': item.StockMedida,
            'Cantidad a Enviar': item.ToSend,
            'Distribución de Origen': item.Sugerencia
        };
    });

    const ws = XLSX.utils.json_to_sheet(excelData);

    const wscols = [
        {wch: 15}, // SKU
        {wch: 45}, // Detalle
        {wch: 14}, // Medida
        {wch: 15}, // Ventas
        {wch: 20}, // Apto
        {wch: 20}, // Camino
        {wch: 28}, // Stock Depósitos (combinado)
        {wch: 22}, // Stock Total de la Medida
        {wch: 20}, // Cantidad a Enviar
        {wch: 55}  // Sugerencia
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Plan de Envío Full");
    
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Plan_Envio_Full_${dateStr}.xlsx`);
});

// =============================================
// --- MÓDULO PEDIDOS (Reposición de Stock) ---
// =============================================

(function() {
    const dropPriceList   = document.getElementById('pedidos-drop-pricelist');
    const dropOportun     = document.getElementById('pedidos-drop-oportunidades');
    const dropSales   = document.getElementById('pedidos-drop-sales');
    const dropSalesMl = document.getElementById('pedidos-drop-sales-ml');
    const dropStock   = document.getElementById('pedidos-drop-stock');
    const filePriceList    = document.getElementById('pedidos-file-pricelist');
    const fileOportun      = document.getElementById('pedidos-file-oportunidades');
    const fileSales   = document.getElementById('pedidos-file-sales');
    const fileSalesMl = document.getElementById('pedidos-file-sales-ml');
    const fileStock   = document.getElementById('pedidos-file-stock');
    const statusPriceList  = document.getElementById('pedidos-status-pricelist');
    const statusOportun    = document.getElementById('pedidos-status-oportunidades');
    const statusSales = document.getElementById('pedidos-status-sales');
    const statusSalesMl = document.getElementById('pedidos-status-sales-ml');
    const statusStock = document.getElementById('pedidos-status-stock');
    const mesesInput  = document.getElementById('pedidos-meses');
    const resultsPanel= document.getElementById('pedidos-results-panel');
    const resultsBody = document.getElementById('pedidos-results-body');
    const btnDownload = document.getElementById('pedidos-btn-download');
    const descListaInputs  = ['pedidos-desc-l1','pedidos-desc-l2','pedidos-desc-l3','pedidos-desc-l4','pedidos-desc-l5'].map(id => document.getElementById(id));
    const descOportInputs  = ['pedidos-desc-o1','pedidos-desc-o2'].map(id => document.getElementById(id));

    let salesMap = null;   // Map: sku -> { sku, descripcion, unidades }
    let salesMLMap = null; // Map: sku -> { sku, descripcion, unidades }  (export ML, opcional, complementa al reporte de ventas)
    let stockMap = null;   // Map: sku -> { sku, descripcion, cgil, sMarzo, sMarzo1435, full, enCamino, costo }
    let priceListMap = new Map();   // sku -> { sku, marca, medida, modelo, precioBase, fuente }
    let offersBySku   = new Map();  // sku -> { precioBase, medida, modelo, fuente }  (Oportunidades)
    let salesPeriodMonths = null;   // meses cubiertos por el reporte de ventas cargado (calculado de las fechas)
    let pedidosResults = [];

    // --- HELPERS ---
    function norm(s) {
        return String(s === undefined || s === null ? '' : s).toLowerCase().trim().normalize('NFD').replace(/[̀-ͯ]/g, '');
    }

    function cleanSkuP(raw) {
        if (raw === undefined || raw === null) return '';
        let s = String(raw).trim();
        if (s.startsWith("'")) s = s.substring(1);
        s = s.trim().split('.')[0].split(',')[0];
        return s;
    }

    function descuentoCombinado(inputs) {
        let f = 1;
        inputs.forEach(inp => {
            const v = inp ? parseFloat(inp.value) : NaN;
            f *= (isNaN(v) || v <= 0) ? 1 : v;
        });
        return f;
    }

    // --- MEDIDA: parseo para poder ordenar numéricamente (205/60R16, 11R22.5, 7.50-16...) ---
    function parseMedidaOrden(medida) {
        if (!medida) return null;
        const s = String(medida);
        const m = s.match(/(\d+(?:\.\d+)?)\s*\/?\s*(\d+(?:\.\d+)?)?\s*R\s*(\d+(?:\.\d+)?)/i) || s.match(/(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)/);
        if (!m) return null;
        if (m.length === 4) return { ancho: parseFloat(m[1]) || 0, perfil: parseFloat(m[2]) || 0, llanta: parseFloat(m[3]) || 0 };
        return { ancho: parseFloat(m[1]) || 0, perfil: 0, llanta: parseFloat(m[2]) || 0 };
    }
    function compararMedida(a, b) {
        const pa = parseMedidaOrden(a), pb = parseMedidaOrden(b);
        if (!pa && !pb) return 0;
        if (!pa) return 1;
        if (!pb) return -1;
        return (pa.ancho - pb.ancho) || (pa.perfil - pb.perfil) || (pa.llanta - pb.llanta);
    }

    // --- KITS: un kit son N unidades del mismo neumático base ---
    // "kitx2-60700" -> { sku:"60700", mult:2 }   ·   "60700X2" -> { sku:"60700", mult:2 }
    // (mismos formatos que manejan los otros módulos de la app)
    function resolverKit(rawSku) {
        const s = String(rawSku || '');
        let m = s.match(/^kitx(\d+)-(.+)$/i);
        if (m) return { sku: cleanSkuP(m[2]), mult: parseInt(m[1], 10) || 1 };
        m = s.match(/^(.+)x(\d+)$/i);
        if (m) return { sku: cleanSkuP(m[1]), mult: parseInt(m[2], 10) || 1 };
        return { sku: s, mult: 1 };
    }

    function fmt(n) {
        return n.toLocaleString('es-AR', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    }

    // --- FILTRO: SOLO NEUMÁTICOS ---
    // Se necesitan las dos señales por separado, porque ninguna cubre todo el catálogo:
    // "Kit De 2 Neumáticos Kumho Es31 P 88 T" no trae medida, y
    // "Scorpion Zero 255/55r18 All Season" no trae la palabra.
    // La categoría del reporte de stock no sirve: hay neumáticos cargados en ACCESORIOS y en SIN CATEGORIA.
    const RX_NEUM_PALABRA = /neum|cubierta/i;
    const RX_NEUM_MEDIDA = [
        /\d{3}\s*\/\s*\d{2}\s*(z\s*)?r?\s*\d{2}(\.\d)?/i,   // 195/65R15 · 175/65 R15 · 215/75R17.5 · Lt285/75r16
        /\d{3}\s+\d{2}\s+(z\s*)?r\s*\d{2}/i,                // 205 55 R16 · 195 75 R16c · 195 55 Z R16
        /\d{3}\s*r\s*\d{2}c?\b/i,                           // 195 R14 · 195r15c  (carga, sin perfil)
        /\d{2}\s*x\s*\d{1,2}(\.\d+)?\s*(lt\s*)?r\s*\d{2}/i, // 31X10.50 R15LT · 31X10.50 LT R15  (la R evita "5w40 X 4l")
        /\d\.\d{2}\s*r\s*\d{2}/i,                           // 5.00r12 · 5.00 R12
        /\d\.\d{2}\s*-\s*\d{2}\b/                           // 7.50-16            (convencional)
    ];

    function esNeumatico() {
        for (const txt of arguments) {
            if (!txt) continue;
            const s = String(txt);
            if (RX_NEUM_PALABRA.test(s)) return true;
            if (RX_NEUM_MEDIDA.some(rx => rx.test(s))) return true;
        }
        return false;
    }

    // --- REDONDEO DE COMPRA ---
    // Nunca menos de 4 unidades; por debajo de 10 se redondea al par superior;
    // de 10 en adelante, al múltiplo de 10 superior.
    const MIN_PEDIDO = 4;

    function redondearPedido(n) {
        if (n <= 0) return 0;
        if (n <= MIN_PEDIDO) return MIN_PEDIDO;
        if (n < 10) return n % 2 === 0 ? n : n + 1;
        return Math.ceil(n / 10) * 10;
    }

    // --- MEDIDA + MODELO desde una descripción de neumático ---
    // Se usa para cruzar ofertas cuyo SKU de proveedor no coincide con el del stock.
    // Ej: "205/60R16 FM601 92V FIREMAX" -> { medida:"205/60R16", modelo:"FM601" }
    const RX_MODELO = /\b(FM\d{3}|KT\d{3,4}|ES\d{2}|HS\d{2}|TH\d{3}|TA\d{2}|TE\d{3}|TR\d{3}|HP\d{2}|PS\d{2,3}|PC\d{2}|KC\d{2}|MT\d{2}|DH\d{2}|DS[A-Z]?\d{2}|DU\w\d{2}|DL\d{2}|W\d{2}|T\d{2}|G\d{3})\b/i;
    function medidaModelo(desc) {
        const d = String(desc || '').toUpperCase();
        const m = d.match(/(\d{3})\s*\/\s*(\d{2})\s*(?:Z\s*)?R?\s*(\d{2})/);
        const medida = m ? `${m[1]}/${m[2]}R${m[3]}` : null;
        const mod = d.match(RX_MODELO);
        return { medida, modelo: mod ? mod[1].toUpperCase() : null };
    }

    // --- DIBUJO: lisa (calle/highway) vs taco (todoterreno AT/MT) ---
    // Una misma medida puede tener los dos tipos (ej. 265/65R17 lisa y con tacos) y no son
    // intercambiables. Se detecta por la descripción. Si dice HT/HP (Highway) es lisa aunque
    // el nombre suene offroad (ej. "Wrangler Fortitude HT").
    const RX_TACO = /(\bA\/?T\b|\bM\/?T\b|\bR\/?T\b|ALL\s*TERRAIN|\bMUD\b|\bATR\b|\bAT\d{1,2}\b|\bMT\d{1,2}\b|\bKL7\d\b|\bFM501\b|\bFM523\b|\bG01[58]\b|GEOLANDAR|WRANGLER|OPEN\s*COUNTRY\s*A|ROAD\s*VENTURE|\bTR292\b)/i;
    const RX_HIGHWAY = /(\bH\/?T\b|\bH\/?P\b)/i;
    function dibujoDeDesc(desc) {
        const d = String(desc || '');
        if (!d) return '';
        return (RX_TACO.test(d) && !RX_HIGHWAY.test(d)) ? 'Taco' : 'Lisa';
    }

    // El pedido es para el proveedor Guerrini, que solo vende estas marcas: no se recomienda
    // comprar ninguna otra (Pirelli, Continental, Nexen, etc.), aunque falte stock.
    const RX_GUERRINI = /\b(KUMHO|KHUMO|FIREMAX|TRIANGLE|KPATOS|DOUBLE\s*STAR|MARSHALL?)\b/i;
    function esMarcaGuerrini(marca, desc) {
        return RX_GUERRINI.test(String(marca || '')) || RX_GUERRINI.test(String(desc || ''));
    }

    // Etiqueta corta del origen de la oferta a partir del nombre de archivo ("...GRUPOA..." -> "Grupo A").
    function etiquetaFuente(fuente) {
        const f = String(fuente || '').toUpperCase();
        const m = f.match(/GRUPO\s*([A-Z])/);
        if (m) return 'Grupo ' + m[1];
        return String(fuente || '').replace(/\.[^.]+$/, '');
    }

    // --- DROP ZONE SETUP ---
    function setupPedidosZone(dropZone, fileInput, onFile) {
        dropZone.addEventListener('click', () => fileInput.click());
        dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
        dropZone.addEventListener('drop', e => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');
            if (e.dataTransfer.files.length) onFile(e.dataTransfer.files[0]);
        });
        fileInput.addEventListener('change', () => {
            if (fileInput.files.length) onFile(fileInput.files[0]);
        });
    }

    // Las zonas de lista de precios y oportunidades aceptan varios archivos a la vez.
    function setupMultiZone(dropZone, fileInput, onFiles) {
        dropZone.addEventListener('click', () => fileInput.click());
        dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
        dropZone.addEventListener('drop', e => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');
            if (e.dataTransfer.files.length) onFiles(e.dataTransfer.files);
        });
        fileInput.addEventListener('change', () => {
            if (fileInput.files.length) onFiles(fileInput.files);
        });
    }

    setupPedidosZone(dropSales, fileSales, handleSalesFile);
    setupMultiZone(dropSalesMl, fileSalesMl, handleSalesMLFiles);
    setupPedidosZone(dropStock, fileStock, handleStockFile);
    setupMultiZone(dropPriceList, filePriceList, handlePriceListFiles);
    setupMultiZone(dropOportun, fileOportun, handleOportunidadesFiles);
    descListaInputs.concat(descOportInputs).forEach(inp => {
        if (!inp) return;
        inp.addEventListener('input', () => { if (priceListMap.size > 0) calcAndRender(); });
    });

    // --- SKU de proveedor: puede venir con separador de miles ("60.698") -> normalizar a "60698" ---
    function cleanSkuProveedor(raw) {
        let s = String(raw === undefined || raw === null ? '' : raw).trim();
        if (s.startsWith("'")) s = s.substring(1);
        if (/^\d{1,3}(\.\d{3})+$/.test(s)) return s.replace(/\./g, '');   // 60.698 -> 60698
        return cleanSkuP(s);
    }

    // Convierte un valor de celda "Fecha" (string ISO, Date o serial de Excel) a Date.
    function parseFechaCell(v) {
        if (v === undefined || v === null || v === '') return null;
        if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
        if (typeof v === 'number') {
            const d = new Date(Math.round((v - 25569) * 86400 * 1000));
            return isNaN(d.getTime()) ? null : d;
        }
        const s = String(v).trim();
        const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
        const d = new Date(s);
        return isNaN(d.getTime()) ? null : d;
    }

    // --- PARSE VENTAS (reporte_ventas: Fecha, Tipo, Código, Descripción, Categoría, Canal, Cantidad, Subtotal) ---
    // Solo se toman las filas cuya Categoría contiene "radial" (neumáticos de auto/camioneta/camión).
    function handleSalesFile(file) {
        statusSales.textContent = 'Procesando...';
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                const wsName = wb.SheetNames.find(n => norm(n).includes('datos')) || wb.SheetNames[0];
                const ws = wb.Sheets[wsName];
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

                let headerIdx = 0, fechaIdx = 0, skuIdx = 2, descIdx = 3, catIdx = 4, cantIdx = 6;
                for (let i = 0; i < Math.min(rows.length, 10); i++) {
                    const row = rows[i];
                    if (!row || !Array.isArray(row)) continue;
                    const cols = Array.from(row).map(norm);
                    const tSku  = cols.findIndex(c => c.includes('codigo'));
                    const tCat  = cols.findIndex(c => c.includes('categoria'));
                    const tCant = cols.findIndex(c => c.includes('cantidad'));
                    if (tSku !== -1 && tCat !== -1 && tCant !== -1) {
                        headerIdx = i; skuIdx = tSku; catIdx = tCat; cantIdx = tCant;
                        const tFecha = cols.findIndex(c => c.includes('fecha'));
                        if (tFecha !== -1) fechaIdx = tFecha;
                        const tDesc = cols.findIndex(c => c.includes('descripcion'));
                        if (tDesc !== -1) descIdx = tDesc;
                        break;
                    }
                }

                salesMap = new Map();
                let totalUnits = 0, minFecha = null, maxFecha = null, filas = 0;
                for (let i = headerIdx + 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row || row.length === 0) continue;
                    const categoria = String(row[catIdx] || '');
                    if (!/radial/i.test(categoria)) continue;
                    const rawSku = cleanSkuP(row[skuIdx]);
                    if (!rawSku) continue;
                    const cant = parseFloat(row[cantIdx]) || 0;
                    if (cant === 0) continue;
                    // Un kit se contabiliza como (unidades × N) del neumático base.
                    const { sku, mult } = resolverKit(rawSku);
                    const units = cant * mult;
                    const desc  = String(row[descIdx] !== undefined ? row[descIdx] : '').trim();
                    totalUnits += units;
                    filas++;
                    const fecha = parseFechaCell(row[fechaIdx]);
                    if (fecha) { if (!minFecha || fecha < minFecha) minFecha = fecha; if (!maxFecha || fecha > maxFecha) maxFecha = fecha; }
                    const ex = salesMap.get(sku);
                    if (ex) { ex.unidades += units; if (!ex.descripcion && desc) ex.descripcion = desc; }
                    else salesMap.set(sku, { sku, descripcion: desc, unidades: units });
                }
                // Netear devoluciones: una fila negativa no puede dejar a un SKU con ventas negativas.
                salesMap.forEach(v => { v.unidades = Math.max(0, v.unidades); });

                if (minFecha && maxFecha) {
                    const dias = Math.max(1, Math.round((maxFecha - minFecha) / 86400000) + 1);
                    salesPeriodMonths = Math.max(0.5, dias / 30.4368);
                } else {
                    salesPeriodMonths = 2;
                }

                dropSales.classList.add('drag-over');
                statusSales.textContent = `✓ ${file.name} — ${salesMap.size} SKUs, ${fmt(totalUnits)} uds. (${salesPeriodMonths.toFixed(1)}m)`;
                statusSales.classList.add('uploaded');
                tryRenderPedidos();
            } catch(err) {
                console.error(err);
                Swal.fire('Error', 'No se pudo leer el reporte de ventas: ' + err.message, 'error');
                statusSales.textContent = 'Error al leer';
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // --- PARSE VENTAS MERCADO LIBRE (opcional, complementa al reporte de ventas) ---
    // El layout de columnas de este export cambia entre descargas de ML (agrega/mueve columnas);
    // se detectan por nombre, no por posición fija. Se descartan filas con Estado de cancelación o
    // devolución para no inflar la demanda real con ventas que no se concretaron.
    const RX_ESTADO_DESCARTAR = /cancela|devoluci/i;
    // Acepta varias cuentas de ML a la vez (ej. 2 cuentas). Cada archivo es una cuenta distinta:
    // las ventas de un mismo SKU se SUMAN entre cuentas para reflejar la demanda total en ML.
    // Se reinicia con cada tanda: arrastrá los 2 archivos juntos (o seleccionalos en un solo paso).
    function handleSalesMLFiles(fileList) {
        const files = Array.from(fileList);
        statusSalesMl.textContent = 'Procesando...';
        salesMLMap = new Map();
        let done = 0, cuentas = 0, totalUnits = 0, descartadas = 0, errores = 0;

        files.forEach(file => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

                    let headerIdx = 5, skuIdx = 22, descIdx = 24, unitsIdx = 6, estadoIdx = 2;
                    for (let i = 0; i < Math.min(rows.length, 15); i++) {
                        const row = rows[i];
                        if (!row || !Array.isArray(row)) continue;
                        const cols = Array.from(row).map(norm);
                        const tSku   = cols.findIndex(c => c === 'sku');
                        const tUnits = cols.findIndex(c => c === 'unidades');
                        if (tSku !== -1 && tUnits !== -1) {
                            headerIdx = i; skuIdx = tSku; unitsIdx = tUnits;
                            const tEstado = cols.findIndex(c => c === 'estado');
                            if (tEstado !== -1) estadoIdx = tEstado;
                            const tDesc = cols.findIndex(c => c.includes('titulo de la publicacion'));
                            if (tDesc !== -1) descIdx = tDesc;
                            break;
                        }
                    }

                    for (let i = headerIdx + 1; i < rows.length; i++) {
                        const row = rows[i];
                        if (!row || row.length === 0) continue;
                        const estado = String(row[estadoIdx] !== undefined ? row[estadoIdx] : '');
                        if (RX_ESTADO_DESCARTAR.test(estado)) { descartadas++; continue; }
                        const rawSku = cleanSkuP(row[skuIdx]);
                        if (!rawSku) continue;
                        const soldUnits = parseFloat(row[unitsIdx]) || 0;
                        if (soldUnits <= 0) continue;
                        const desc  = String(row[descIdx] !== undefined ? row[descIdx] : '').trim();
                        // Este export no tiene columna de categoría (a diferencia del reporte de ventas):
                        // se filtra por título para no mezclar baterías, filtros, etc. de la misma cuenta de ML.
                        if (!esNeumatico(desc)) continue;
                        const { sku, mult } = resolverKit(rawSku);
                        const units = soldUnits * mult;
                        totalUnits += units;
                        const ex = salesMLMap.get(sku);
                        if (ex) { ex.unidades += units; if (!ex.descripcion && desc) ex.descripcion = desc; }
                        else salesMLMap.set(sku, { sku, descripcion: desc, unidades: units });
                    }
                    cuentas++;
                } catch(err) {
                    console.error('Ventas ML ' + file.name, err);
                    errores++;
                } finally {
                    done++;
                    if (done === files.length) finalizeSalesML(cuentas, totalUnits, descartadas, errores);
                }
            };
            reader.onerror = function() { done++; errores++; if (done === files.length) finalizeSalesML(cuentas, totalUnits, descartadas, errores); };
            reader.readAsArrayBuffer(file);
        });
    }

    function finalizeSalesML(cuentas, totalUnits, descartadas, errores) {
        if (cuentas === 0 && errores > 0) {
            Swal.fire('Error', 'No se pudo leer el/los Excel de ventas de Mercado Libre.', 'error');
            statusSalesMl.textContent = 'Error al leer';
            return;
        }
        dropSalesMl.classList.add('drag-over');
        statusSalesMl.textContent = `✓ ${cuentas} cuenta/s — ${salesMLMap.size} SKUs, ${fmt(totalUnits)} uds.`
            + (descartadas ? ` (${descartadas} canceladas/devueltas descartadas)` : '')
            + (errores ? ` (${errores} archivo/s con error)` : '');
        statusSalesMl.classList.add('uploaded');
        tryRenderPedidos();
    }

    // --- PARSE STOCK ---
    function handleStockFile(file) {
        statusStock.textContent = 'Procesando...';
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

                let headerIdx = 0, skuIdx = 1, descIdx = 2;
                let cgilIdx = 5, sMarzoIdx = 6, sMarzo1435Idx = 7, fullIdx = 8, caminoIdx = -1, costIdx = 9;

                for (let i = 0; i < Math.min(rows.length, 15); i++) {
                    const row = rows[i];
                    if (!row || !Array.isArray(row)) continue;
                    const cols = Array.from(row).map(c => String(c || '').toLowerCase().trim());
                    const tSku = cols.findIndex(c => c === 'sku' || c === 'código' || c === 'codigo');
                    if (tSku !== -1) {
                        headerIdx = i; skuIdx = tSku;
                        const tDesc  = cols.findIndex(c => c && (c.includes('descripción') || c.includes('descripcion') || c.includes('detalle') || c.includes('nombre') || c.includes('producto')));
                        if (tDesc !== -1) descIdx = tDesc;
                        const tCgil  = cols.findIndex(c => c && (c.includes('coronel') || c.includes('cgil') || c.includes('c. gil')));
                        if (tCgil !== -1) cgilIdx = tCgil;
                        const tSm    = cols.findIndex(c => c && c.includes('santiago marzo') && !c.includes('1435'));
                        if (tSm !== -1) sMarzoIdx = tSm;
                        const tSm14  = cols.findIndex(c => c && (c.includes('santiago marzo 1435') || c.includes('1435')));
                        if (tSm14 !== -1) sMarzo1435Idx = tSm14;
                        const tFull  = cols.findIndex(c => c && (c.includes('full') || c.includes('ml full')));
                        if (tFull !== -1) fullIdx = tFull;
                        // "En camino" = mercadería que ya está por recibirse (col J en el stock_todo).
                        const tCamino = cols.findIndex(c => c && (c.includes('en camino') || c.includes('camino') || c.includes('recibir') || c.includes('transito') || c.includes('tránsito') || c.includes('pendiente')));
                        if (tCamino !== -1) caminoIdx = tCamino;
                        const tCost  = cols.findIndex(c => c && (c.includes('costo') || c === 'cost'));
                        if (tCost !== -1) costIdx = tCost;
                        break;
                    }
                }

                stockMap = new Map();
                for (let i = headerIdx + 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row || row.length === 0) continue;
                    const sku = cleanSkuP(row[skuIdx]);
                    if (!sku) continue;
                    const desc       = String(row[descIdx] !== undefined ? row[descIdx] : '').trim();
                    const cgil       = parseInt(row[cgilIdx])       || 0;
                    const sMarzo     = parseInt(row[sMarzoIdx])     || 0;
                    const sMarzo1435 = parseInt(row[sMarzo1435Idx]) || 0;
                    const full       = parseInt(row[fullIdx])       || 0;
                    const enCamino   = caminoIdx !== -1 ? (parseInt(row[caminoIdx]) || 0) : 0;
                    const costo      = parseFloat(row[costIdx])     || 0;
                    const ex = stockMap.get(sku);
                    if (ex) { ex.cgil += cgil; ex.sMarzo += sMarzo; ex.sMarzo1435 += sMarzo1435; ex.full += full; ex.enCamino += enCamino; }
                    else stockMap.set(sku, { sku, descripcion: desc, cgil, sMarzo, sMarzo1435, full, enCamino, costo });
                }

                statusStock.textContent = `✓ ${file.name} — ${stockMap.size} SKUs en stock.`;
                statusStock.classList.add('uploaded');
                tryRenderPedidos();
            } catch(err) {
                console.error(err);
                Swal.fire('Error', 'No se pudo leer el archivo de stock: ' + err.message, 'error');
                statusStock.textContent = 'Error al leer';
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // La marca sale del nombre de la hoja, que a veces trae un sufijo del proveedor ("DOUBLESTAR_PCR").
    // Se saca ese sufijo (si está; algunos meses puede no venir) y se deja en formato lindo: "Doublestar".
    function limpiarMarcaHoja(sheetName) {
        let s = String(sheetName || '').trim();
        s = s.replace(/[_\s-]*pcr\s*$/i, '').trim();                          // quita "_PCR" del final si existe
        return s.replace(/\S+/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());  // Title Case
    }

    // --- PARSE LISTA DE PRECIOS (uno o varios excels, 1 hoja por marca) ---
    // Columnas por hoja: Codigo, Sección, Perfil, Tipo, Llanta, Diseño, Telas, PRECIO.
    // La marca sale del nombre de la hoja. Se ignoran las hojas sin ese encabezado (portada, hojas de costos internas, etc.)
    function handlePriceListFiles(fileList) {
        const files = Array.from(fileList);
        statusPriceList.textContent = 'Procesando...';
        let done = 0, totalNuevas = 0, hojasOk = 0, errores = 0;

        files.forEach(file => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                    wb.SheetNames.forEach(sheetName => {
                        const ws = wb.Sheets[sheetName];
                        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

                        let hIdx = -1, codIdx = -1, seccIdx = -1, perfilIdx = -1, llantaIdx = -1, disenoIdx = -1, precioIdx = -1;
                        for (let i = 0; i < Math.min(rows.length, 12); i++) {
                            const row = rows[i];
                            if (!row || !Array.isArray(row)) continue;
                            const cols = Array.from(row).map(norm);
                            const tCod = cols.findIndex(c => c.includes('codigo'));
                            const tPre = cols.findIndex(c => c.startsWith('precio'));
                            if (tCod !== -1 && tPre !== -1) {
                                hIdx = i; codIdx = tCod; precioIdx = tPre;
                                // Algunas hojas (ej. Doublestar) usan otros nombres: SIZE (medida junta "165/70"),
                                // RIM (llanta) y PTTN (diseño) en vez de Sección/Llanta/Diseño.
                                seccIdx   = cols.findIndex(c => c.startsWith('secc') || c.startsWith('size') || c.startsWith('medida'));
                                perfilIdx = cols.findIndex(c => c.startsWith('perfil'));
                                llantaIdx = cols.findIndex(c => c.startsWith('llant') || c.startsWith('rim') || c.startsWith('rodado'));
                                disenoIdx = cols.findIndex(c => c.startsWith('diseno') || c.startsWith('pttn') || c.startsWith('dibujo'));
                                break;
                            }
                        }
                        if (hIdx === -1) return;   // hoja sin lista de precios (portada, hoja de costos, etc.)

                        const marca = limpiarMarcaHoja(sheetName);
                        let nuevas = 0;
                        for (let i = hIdx + 1; i < rows.length; i++) {
                            const row = rows[i];
                            if (!row || row.length === 0) continue;
                            const sku = cleanSkuProveedor(row[codIdx]);
                            if (!sku || !/\d/.test(sku)) continue;
                            const precioBase = parseFloat(row[precioIdx]) || 0;
                            if (precioBase <= 0) continue;
                            // La sección puede venir sola ("165") o junta con el perfil en la columna SIZE ("165/70").
                            let secc = '', perfil = '';
                            if (seccIdx !== -1) {
                                const raw = String(row[seccIdx] !== undefined ? row[seccIdx] : '').trim();
                                if (raw.includes('/')) {
                                    const parts = raw.split('/');
                                    secc   = parts[0].trim().split('.')[0];
                                    perfil = (parts[1] || '').trim().split('.')[0];
                                } else {
                                    secc = raw.split('.')[0];
                                }
                            }
                            if (!perfil && perfilIdx !== -1) perfil = String(row[perfilIdx] !== undefined ? row[perfilIdx] : '').trim().split('.')[0];
                            const llanta = llantaIdx !== -1 ? String(row[llantaIdx] !== undefined ? row[llantaIdx] : '').trim() : '';
                            const modelo = disenoIdx !== -1 ? String(row[disenoIdx] !== undefined ? row[disenoIdx] : '').trim().toUpperCase() : '';
                            let medida = null;
                            if (secc && llanta) medida = perfil ? `${secc}/${perfil}R${llanta}` : `${secc}R${llanta}`;
                            priceListMap.set(sku, { sku, marca, medida, modelo, precioBase, fuente: file.name });
                            nuevas++;
                        }
                        if (nuevas > 0) { totalNuevas += nuevas; hojasOk++; }
                    });
                } catch(err) {
                    console.error('Lista de precios ' + file.name, err);
                    errores++;
                } finally {
                    done++;
                    if (done === files.length) finalizePriceList(totalNuevas, hojasOk, errores);
                }
            };
            reader.onerror = function() { done++; errores++; if (done === files.length) finalizePriceList(totalNuevas, hojasOk, errores); };
            reader.readAsArrayBuffer(file);
        });
    }

    function finalizePriceList(nuevas, hojasOk, errores) {
        const msg = `✓ ${priceListMap.size} SKUs en ${hojasOk} hoja/s de marca` + (errores ? ` (${errores} archivo/s con error)` : '');
        statusPriceList.textContent = msg;
        statusPriceList.classList.add('uploaded');
        dropPriceList.classList.add('drag-over');
        tryRenderPedidos();
    }

    // --- PARSE OPORTUNIDADES (una o varias listas del proveedor, formato Guerrini GRUPOA/GRUPOB) ---
    // El precio de la columna OPORTUNIDADES es el precio de lista del ítem en oferta; el descuento
    // configurado ("Descuento Oportunidades") se aplica en calcAndRender, no acá.
    function handleOportunidadesFiles(fileList) {
        const files = Array.from(fileList);
        statusOportun.textContent = 'Procesando...';
        let done = 0, totalNuevas = 0, errores = 0;

        files.forEach(file => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

                    let hIdx = -1, artIdx = 0, diamIdx = -1, seccIdx = -1, llanIdx = -1, disenoIdx = -1, precioIdx = -1;
                    for (let i = 0; i < Math.min(rows.length, 20); i++) {
                        const row = rows[i];
                        if (!row || !Array.isArray(row)) continue;
                        const cols = Array.from(row).map(norm);
                        const tArt = cols.findIndex(c => c.includes('articulo'));
                        const tPre = cols.findIndex(c => c.includes('oportunidad'));
                        if (tArt !== -1 && tPre !== -1) {
                            hIdx = i; artIdx = tArt; precioIdx = tPre;
                            diamIdx   = cols.findIndex(c => c.startsWith('diam'));
                            seccIdx   = cols.findIndex(c => c.startsWith('secc'));
                            llanIdx   = cols.findIndex(c => c.startsWith('llan'));
                            disenoIdx = cols.findIndex(c => c.startsWith('diseno'));
                            break;
                        }
                    }
                    if (hIdx === -1) throw new Error('no se encontró el encabezado (Articulo / OPORTUNIDADES)');

                    let nuevas = 0;
                    for (let i = hIdx + 1; i < rows.length; i++) {
                        const row = rows[i];
                        if (!row || row.length === 0) continue;
                        const sku = cleanSkuProveedor(row[artIdx]);
                        if (!sku || !/\d/.test(sku)) continue;
                        const precioBase = parseFloat(row[precioIdx]) || 0;
                        if (precioBase <= 0) continue;
                        const diam = String(row[diamIdx] !== undefined ? row[diamIdx] : '').trim().split('.')[0];
                        const secc = String(row[seccIdx] !== undefined ? row[seccIdx] : '').trim().split('.')[0];
                        const llan = String(row[llanIdx] !== undefined ? row[llanIdx] : '').trim().split('.')[0];
                        const modelo = String(row[disenoIdx] !== undefined ? row[disenoIdx] : '').trim().toUpperCase();
                        const medida = (diam && secc && llan) ? `${diam}/${secc}R${llan}` : null;
                        offersBySku.set(sku, { precioBase, medida, modelo, fuente: file.name });
                        nuevas++;
                    }
                    totalNuevas += nuevas;
                } catch(err) {
                    console.error('Oportunidad ' + file.name, err);
                    errores++;
                } finally {
                    done++;
                    if (done === files.length) finalizeOportunidades(totalNuevas, errores);
                }
            };
            reader.onerror = function() { done++; errores++; if (done === files.length) finalizeOportunidades(totalNuevas, errores); };
            reader.readAsArrayBuffer(file);
        });
    }

    function finalizeOportunidades(nuevas, errores) {
        const msg = `✓ ${offersBySku.size} oportunidades cargadas` + (errores ? ` (${errores} archivo/s con error)` : '');
        statusOportun.textContent = msg;
        statusOportun.classList.add('uploaded');
        dropOportun.classList.add('drag-over');
        tryRenderPedidos();
    }

    // --- CALCULAR Y RENDERIZAR ---
    // Universo de filas = todos los SKUs de la lista de precios, más los SKUs de ventas/stock
    // que sean neumáticos aunque no estén en ninguna lista de precios cargada.
    function tryRenderPedidos() {
        if (priceListMap.size === 0 && !salesMap && !salesMLMap && !stockMap) return;
        calcAndRender();
    }

    // Ventas final por SKU = la mayor cantidad entre el reporte de ventas y el export de ML,
    // por si alguno de los dos no capturó todas las ventas de ese código.
    function ventasCombinadas(sku) {
        const a = salesMap   ? salesMap.get(sku)   : null;
        const b = salesMLMap ? salesMLMap.get(sku) : null;
        if (!a && !b) return null;
        const unidadesA = a ? a.unidades : 0;
        const unidadesB = b ? b.unidades : 0;
        const ganador = unidadesB > unidadesA ? b : a;
        return { unidades: Math.max(unidadesA, unidadesB), descripcion: (ganador && ganador.descripcion) || (a && a.descripcion) || (b && b.descripcion) || '' };
    }

    function calcAndRender() {
        const M = parseFloat(mesesInput.value) || 2;
        const monthsSales = salesPeriodMonths || 2;
        const descLista = descuentoCombinado(descListaInputs);
        const descOport = descuentoCombinado(descOportInputs);

        const allSkus = new Set(priceListMap.keys());
        if (salesMap) salesMap.forEach((_, sku) => allSkus.add(sku));
        if (salesMLMap) salesMLMap.forEach((_, sku) => allSkus.add(sku));
        if (stockMap) stockMap.forEach((v, sku) => { if (esNeumatico(v.descripcion)) allSkus.add(sku); });

        // Guerrini a veces lista la misma cubierta (misma medida+marca+modelo) bajo dos códigos de
        // artículo distintos. Se agrupan por esa identidad para no duplicar la fila; dentro del grupo
        // se suman las cantidades (ventas/stock pueden estar repartidas entre los dos códigos viejos)
        // y se elige un solo precio representante.
        const groups = new Map();   // key -> array de candidatos
        allSkus.forEach(sku => {
            const priceRec = priceListMap.get(sku) || null;
            const oport    = offersBySku.get(sku)  || null;
            const stock    = stockMap ? stockMap.get(sku) : null;
            // Ventas = la mayor cantidad entre el reporte de ventas y el export de ML para ese SKU.
            const sale     = ventasCombinadas(sku);

            let desc = (stock && stock.descripcion) || (sale && sale.descripcion) || '';
            let medida = priceRec && priceRec.medida;
            let modelo = priceRec && priceRec.modelo;
            if (!medida || !modelo) {
                const mm = medidaModelo(desc);
                medida = medida || mm.medida;
                modelo = modelo || mm.modelo;
            }
            const marca = (priceRec && priceRec.marca) || '';
            // Sin descripción propia (típico de SKUs que solo están en la lista de precios, sin ventas/stock):
            // se arma con medida + modelo + marca en vez de dejarla vacía.
            if (!desc) desc = [medida, modelo, marca].filter(Boolean).join(' ') || '—';

            const precioLista       = (priceRec && priceRec.precioBase > 0) ? priceRec.precioBase * descLista : 0;
            const precioOportunidad = (oport && oport.precioBase > 0) ? oport.precioBase * descOport : 0;
            const enOferta = precioOportunidad > 0;
            const costoStock = stock ? stock.costo : 0;

            const sucursales = stock ? (stock.cgil + stock.sMarzo + stock.sMarzo1435) : 0;
            const full       = stock ? stock.full : 0;
            const enCamino   = stock ? (stock.enCamino || 0) : 0;
            const totalStock = sucursales + full;
            const totalSold  = sale ? sale.unidades : 0;

            // Ruido: SKU que no está en ninguna lista de precios (no se puede pedir), no tuvo ventas
            // y no tiene stock en ningún lado -> no aporta nada, no vale la pena mostrarlo.
            if (!priceRec && !sale && totalStock === 0) return;

            // La identidad del grupo se toma de la lista de precios (medida+marca+modelo tal cual la
            // cargó Guerrini), no de la medida/modelo detectados por fallback, para no agrupar por error.
            const groupKey = (priceRec && priceRec.marca && priceRec.medida && priceRec.modelo)
                ? `${priceRec.medida}|${priceRec.marca}|${priceRec.modelo}`.toUpperCase()
                : `sku:${sku}`;

            const candidate = {
                sku, desc, marca, medida, modelo,
                totalSold, sucursales, full, enCamino, totalStock,
                hasStock: !!stock, hasSales: !!sale,
                precioLista, precioOportunidad, enOferta, costoStock,
                ofertaFuente: oport ? etiquetaFuente(oport.fuente) : ''
            };
            if (!groups.has(groupKey)) groups.set(groupKey, []);
            groups.get(groupKey).push(candidate);
        });

        pedidosResults = [];
        groups.forEach(list => {
            const totalSold   = list.reduce((s, r) => s + r.totalSold, 0);
            const sucursales  = list.reduce((s, r) => s + r.sucursales, 0);
            const full        = list.reduce((s, r) => s + r.full, 0);
            const enCamino    = list.reduce((s, r) => s + (r.enCamino || 0), 0);
            const totalStock  = sucursales + full;
            const hasStock    = list.some(r => r.hasStock);
            const hasSales    = list.some(r => r.hasSales);

            // Precio representante del grupo: la oportunidad más barata si hay alguna; si no,
            // el precio de lista más barato; si ninguno tiene precio, queda sin precio (fallback a costo de stock).
            const conOferta = list.filter(r => r.enOferta);
            const conLista  = list.filter(r => r.precioLista > 0);
            let winner;
            if (conOferta.length) winner = conOferta.reduce((a, b) => a.precioOportunidad <= b.precioOportunidad ? a : b);
            else if (conLista.length) winner = conLista.reduce((a, b) => a.precioLista <= b.precioLista ? a : b);
            else winner = list[0];

            const costoStockFallback = Math.max(0, ...list.map(r => r.costoStock));
            const enOferta = !!winner.enOferta;
            const precioLista = winner.precioLista;
            const precioOportunidad = winner.precioOportunidad;
            // Prioridad de precio: oportunidad (reemplaza) > lista de precios > costo del reporte de stock (fallback).
            const costo = enOferta ? precioOportunidad : (precioLista > 0 ? precioLista : costoStockFallback);
            const fuentePrecio = enOferta ? 'Oportunidad' : (precioLista > 0 ? 'Lista de Precios' : (costoStockFallback > 0 ? 'Costo de stock' : ''));
            const skusAlt = list.filter(r => r !== winner).map(r => r.sku);

            const monthly     = totalSold / monthsSales;
            // Se descuenta lo que ya está "en camino" (por recibir): no hace falta volver a pedirlo.
            const needed      = hasStock ? Math.max(0, Math.ceil(monthly * M - totalStock - enCamino)) : 0;
            // Solo se recomienda comprar marcas que vende Guerrini (el proveedor de este pedido).
            const esGuerrini  = esMarcaGuerrini(winner.marca, winner.desc);
            const recommended = (hasStock && esGuerrini) ? redondearPedido(needed) : 0;
            const investment  = recommended * costo;

            pedidosResults.push({
                sku: winner.sku, skusAlt, desc: winner.desc, marca: winner.marca, medida: winner.medida, modelo: winner.modelo,
                dibujo: dibujoDeDesc(winner.desc),
                totalSold, sucursales, full, enCamino, totalStock, hasStock, hasSales,
                needed, recommended, costo, costoStock: costoStockFallback, precioLista, precioOportunidad, investment,
                enOferta, ofertaFuente: winner.ofertaFuente, fuentePrecio
            });
        });

        // Orden principal: por medida (ancho/perfil/llanta numérico); dentro de la misma medida, por marca y modelo.
        pedidosResults.sort((a, b) =>
            compararMedida(a.medida, b.medida) ||
            a.marca.localeCompare(b.marca) ||
            String(a.modelo || '').localeCompare(String(b.modelo || '')) ||
            a.sku.localeCompare(b.sku)
        );

        // Stats
        const toOrder     = pedidosResults.filter(r => r.recommended > 0);
        const totalUnits  = toOrder.reduce((s, r) => s + r.recommended, 0);
        const totalInv    = toOrder.reduce((s, r) => s + r.investment, 0);
        const totalVentas = pedidosResults.reduce((s, r) => s + r.totalSold, 0);
        const enOfertaCount = pedidosResults.filter(r => r.enOferta).length;

        document.getElementById('pedidos-stat-skus').textContent      = pedidosResults.length;
        document.getElementById('pedidos-stat-ventas').textContent    = fmt(totalVentas);
        document.getElementById('pedidos-stat-unidades').textContent  = fmt(totalUnits);
        document.getElementById('pedidos-stat-inversion').textContent = '$' + fmt(totalInv);
        const statOferta = document.getElementById('pedidos-stat-oferta');
        if (statOferta) statOferta.textContent = fmt(enOfertaCount);
        const statPeriodo = document.getElementById('pedidos-stat-periodo');
        if (statPeriodo) statPeriodo.textContent = salesMap ? `${monthsSales.toFixed(1)} m` : '—';

        // Tabla
        resultsBody.innerHTML = '';
        pedidosResults.forEach(item => {
            const tr = document.createElement('tr');
            if (item.enOferta) tr.style.borderLeft = '3px solid #a855f7';
            else if (item.recommended > 0) tr.style.borderLeft = '3px solid #22c55e';

            let ofertaCell;
            if (item.enOferta) {
                const ahorro = item.precioLista > 0 ? Math.round((item.precioOportunidad / item.precioLista - 1) * 100) : null;
                const tip = (ahorro !== null ? `Precio de oportunidad vs. lista: ${ahorro >= 0 ? '+' : ''}${ahorro}% — ` : 'Precio de oportunidad — ') + (item.ofertaFuente || 'proveedor') + '. Reemplaza al precio de lista.';
                ofertaCell = `<span title="${tip}" style="display:inline-block; padding:2px 8px; border-radius:10px; background:rgba(168,85,247,0.15); color:#a855f7; font-weight:600; font-size:0.8rem;"><i class="fa-solid fa-tags"></i> En oferta</span>`;
            } else {
                ofertaCell = '<span style="color:var(--text-secondary);">—</span>';
            }

            const skuTip = item.skusAlt && item.skusAlt.length ? `También listado con el/los código/s: ${item.skusAlt.join(', ')} (se unificó por ser la misma medida/marca/modelo).` : '';

            tr.innerHTML = `
                <td><strong title="${skuTip}">${item.sku}${item.skusAlt && item.skusAlt.length ? ' <i class="fa-solid fa-circle-info" style="font-size:0.7rem; opacity:0.6;"></i>' : ''}</strong></td>
                <td style="text-align:center;">${item.medida || '—'}</td>
                <td>${item.marca || '—'}</td>
                <td>${item.modelo || '—'}</td>
                <td style="text-align:center;">${item.dibujo === 'Taco'
                    ? '<span style="display:inline-block; padding:2px 8px; border-radius:10px; background:rgba(245,158,11,0.15); color:#f59e0b; font-weight:600; font-size:0.78rem;">Taco</span>'
                    : (item.dibujo === 'Lisa' ? '<span style="color:var(--text-secondary); font-size:0.8rem;">Lisa</span>' : '—')}</td>
                <td style="font-size:0.82rem; color:var(--text-secondary);">${item.desc}</td>
                <td style="text-align:center;">${item.hasSales ? item.totalSold : '—'}</td>
                <td style="text-align:center;">${item.hasStock ? item.sucursales : '—'}</td>
                <td style="text-align:center;">${item.hasStock ? item.full : '—'}</td>
                <td style="text-align:center;" title="Mercadería en camino (ya por recibirse). Se descuenta del pedido.">${item.hasStock && item.enCamino > 0 ? `<span style="color:#38bdf8; font-weight:600;">${item.enCamino}</span>` : (item.hasStock ? '0' : '—')}</td>
                <td style="text-align:center;" title="${item.recommended > 0 ? `Necesidad calculada: ${item.needed} — redondeado a ${item.recommended}` : ''}"><strong style="color:${item.recommended > 0 ? '#22c55e' : 'inherit'};">${item.recommended || '—'}</strong></td>
                <td style="text-align:center;">${ofertaCell}</td>
                <td style="text-align:center;">${item.costo > 0 ? '$' + fmt(item.costo) : '—'}</td>
                <td style="text-align:center; font-size:0.78rem; color:var(--text-secondary);">${item.fuentePrecio || '—'}</td>
                <td style="text-align:center;">${item.investment > 0 ? '$' + fmt(item.investment) : '—'}</td>
            `;
            resultsBody.appendChild(tr);
        });

        resultsPanel.style.display = '';
    }

    // Recalcular al cambiar meses
    mesesInput.addEventListener('change', tryRenderPedidos);
    mesesInput.addEventListener('input',  tryRenderPedidos);

    // --- EXPORTAR EXCEL ---
    btnDownload.addEventListener('click', () => {
        if (pedidosResults.length === 0) {
            Swal.fire('Sin datos', 'Cargá al menos la lista de precios para generar el Excel.', 'info');
            return;
        }
        const M = parseFloat(mesesInput.value) || 2;
        const data = pedidosResults.map(item => ({
            'SKU': item.sku,
            'Medida': item.medida || '',
            'Marca': item.marca || '',
            'Modelo': item.modelo || '',
            'Dibujo': item.dibujo || '',
            'Descripción': item.desc,
            'Ventas': item.hasSales ? item.totalSold : '',
            'Total Sucursales': item.hasStock ? item.sucursales : '',
            'Stock ML Full': item.hasStock ? item.full : '',
            'En Camino (por recibir)': item.hasStock ? item.enCamino : '',
            'Cantidad a Comprar': item.recommended || '',
            // Precio efectivo para sumar directo: si está en oferta muestra el precio de oportunidad,
            // si no, el de lista. Así una sola columna sirve para el total.
            'Precio Lista (con descuento)': (item.enOferta ? item.precioOportunidad : item.precioLista) || '',
            'En Oferta': item.enOferta ? `Sí (${item.ofertaFuente})` : '',
            'Costo Usado (ARS)': item.costo || '',
            'Precio Usado': item.fuentePrecio || '',
            'Otros Códigos (misma medida/marca/modelo)': (item.skusAlt && item.skusAlt.length) ? item.skusAlt.join(', ') : '',
            'Inversión (ARS)': item.investment || ''
        }));
        // xlsx-js-style (si cargó) permite colorear filas; si no está disponible se exporta sin color.
        const XLib = (typeof XLSXStyle !== 'undefined' && XLSXStyle) ? XLSXStyle : XLSX;
        const ws = XLib.utils.json_to_sheet(data);
        ws['!cols'] = [
            {wch:12},{wch:16},{wch:14},{wch:14},{wch:8},{wch:45},{wch:10},{wch:16},{wch:12},{wch:18},{wch:16},{wch:20},{wch:18},{wch:16},{wch:16},{wch:32},{wch:16}
        ];

        // Bandas de color por grupo de medida: todas las filas de la misma medida comparten color,
        // alternando entre 2 tonos para que se note dónde empieza el grupo siguiente.
        if (XLib === XLSXStyle) {
            const FILL_COLORS = ['FFFFFF', 'DDEBF7'];
            const numCols = Object.keys(data[0] || {}).length;
            let lastMedida = null, colorIdx = -1;
            pedidosResults.forEach((item, i) => {
                const medidaKey = item.medida || ' sin-medida';
                if (medidaKey !== lastMedida) { colorIdx = (colorIdx + 1) % FILL_COLORS.length; lastMedida = medidaKey; }
                for (let c = 0; c < numCols; c++) {
                    const addr = XLib.utils.encode_cell({ r: i + 1, c });
                    if (ws[addr]) ws[addr].s = { fill: { fgColor: { rgb: FILL_COLORS[colorIdx] } } };
                }
            });
        }

        const wbOut = XLib.utils.book_new();
        XLib.utils.book_append_sheet(wbOut, ws, 'Lista de Precios');
        const dateStr = new Date().toISOString().split('T')[0];
        XLib.writeFile(wbOut, `Lista_de_Precios_Neumaticos_${M}m_${dateStr}.xlsx`);
    });
})();

// =============================================
// --- MÓDULO ANÁLISIS DE PRECIOS ML ---
// =============================================

(function() {
    const IVA = 1.21;

    // Elementos
    const dropPub    = document.getElementById('drop-zone-precios-pub');
    const dropVentas = document.getElementById('drop-zone-precios-ventas');
    const dropStock  = document.getElementById('drop-zone-precios-stock');
    const filePub    = document.getElementById('file-precios-pub');
    const fileVentas = document.getElementById('file-precios-ventas');
    const fileStock  = document.getElementById('file-precios-stock');
    const statusPub    = document.getElementById('status-precios-pub');
    const statusVentas = document.getElementById('status-precios-ventas');
    const statusStock  = document.getElementById('status-precios-stock');
    const btnProcess  = document.getElementById('btn-process-precios');
    const btnDownload = document.getElementById('btn-download-precios');
    const resultsPanel = document.getElementById('results-panel-precios');
    const tbody = document.getElementById('results-body-precios');
    const inputMargenMin = document.getElementById('precios-margen-min');
    const inputComision = document.getElementById('precios-comision');
    const chkSoloNeumaticos = document.getElementById('precios-solo-neumaticos');
    const inputBuscar = document.getElementById('precios-buscar');
    const chkSoloProblemas = document.getElementById('precios-solo-problemas');
    const selectOrden = document.getElementById('precios-orden');

    let dataPub = null;      // array de publicaciones parseadas
    let dataVentas = null;   // map MLA -> unidades vendidas último mes
    let dataStock = null;    // map SKU -> costo s/IVA
    let finalResults = [];   // resultado combinado

    // --- Helpers de formato ---
    const fmtMoney = (n) => (n === null || n === undefined || isNaN(n)) ? '—'
        : '$' + Math.round(n).toLocaleString('es-AR');
    const fmtPct = (n) => (n === null || n === undefined || isNaN(n)) ? '—'
        : (n * 100).toFixed(1).replace('.', ',') + '%';
    // Acorta el texto de cuotas de ML para la tabla (mantiene el original en el Excel)
    const fmtCuotas = (c) => {
        if (!c) return '—';
        const low = c.toLowerCase();
        if (low.includes('banco')) return '<span style="opacity:0.6;">Solo bancos</span>';
        const m = c.match(/(\d+)\s*sin inter/i);
        if (m) return `<span style="color:var(--success, #22c55e); font-weight:600;">${m[1]} s/interés</span>`;
        const h = c.match(/hasta\s*(\d+)/i);
        if (h) return `Hasta ${h[1]}`;
        return c;
    };
    const toNum = (v) => {
        if (v === null || v === undefined || v === '') return null;
        if (typeof v === 'number') return v;
        const s = String(v).replace(/[^0-9,.-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.');
        const n = parseFloat(s);
        return isNaN(n) ? null : n;
    };

    // --- Detección de columnas por nombre ---
    function findHeaderRow(rows, mustHave) {
        for (let i = 0; i < Math.min(rows.length, 20); i++) {
            const up = (rows[i] || []).map(c => String(c).trim().toUpperCase());
            if (mustHave.every(m => up.some(h => h.includes(m)))) return i;
        }
        return -1;
    }
    function colFinder(headerRow) {
        const up = headerRow.map(c => String(c).trim().toUpperCase());
        return (...needles) => {
            for (const nd of needles) {
                const idx = up.findIndex(h => h.includes(nd));
                if (idx !== -1) return idx;
            }
            return -1;
        };
    }

    // --- Parser: export de Publicaciones enriquecido ---
    const WRONG_FILE_MSG = "Este no parece el archivo correcto. 'Precios ML' necesita el export que incluye las columnas 'Opción de venta', 'Recibimos (neto)', 'Costo c/IVA' y 'Cuotas' (el que antes se llamaba 'publicaciones_Nuevos_Neumaticos_...'). El export plano de precios/stock (con columnas ITEM_ID, PRICE, STOCK_FLEX) no sirve para este análisis.";

    // Flag: true cuando el archivo cargado es el formato plano (editor masivo),
    // que no trae Recibimos/Comisión/Opción/Cuotas y requiere estimación.
    let pubFormatoPlano = false;

    function parsePublicaciones(rows) {
        // 1) Formato enriquecido (SKU + MLA, o SKU + Recibimos)
        let hIdx = findHeaderRow(rows, ['SKU', 'MLA']);
        if (hIdx === -1) hIdx = findHeaderRow(rows, ['SKU', 'RECIBIMOS']);
        if (hIdx !== -1) { pubFormatoPlano = false; return parsePubEnriquecido(rows, hIdx); }

        // 2) Formato plano / editor masivo (ITEM_ID + PRICE)
        const pIdx = findHeaderRow(rows, ['ITEM_ID', 'PRICE']);
        if (pIdx !== -1) { pubFormatoPlano = true; return parsePubPlano(rows, pIdx); }

        throw new Error(WRONG_FILE_MSG);
    }

    function parsePubEnriquecido(rows, hIdx) {
        const find = colFinder(rows[hIdx]);
        const cSku    = find('SKU');
        const cMla    = find('MLA');
        const cOpcion = find('OPCIÓN DE VENTA', 'OPCION DE VENTA', 'OPCIÓN', 'OPCION');
        const cDesc   = find('DESCRIPCIÓN', 'DESCRIPCION', 'TÍTULO', 'TITULO');
        const cPrecio = find('PRECIO (COMPRADOR)', 'PRECIO COMPRADOR', 'PRECIO');
        const cRecib  = find('RECIBIMOS');
        const cCosto  = find('COSTO C/IVA', 'COSTO');
        const cVentas = find('CANT. VENTAS', 'CANT VENTAS', 'VENTAS');
        const cCuotas = find('CUOTAS');
        const cStock  = find('STOCK');
        // OJO: el encabezado de stock es "Stock depósito (Físico (todos menos Full))"
        // y contiene la palabra "Full", así que la columna Full se busca por match EXACTO.
        const headerUp = rows[hIdx].map(c => String(c).trim().toUpperCase());
        let cFull = headerUp.indexOf('FULL');
        if (cFull === -1) cFull = headerUp.findIndex(h => h === 'FULL' || h === 'EN FULL');
        if (cMla === -1 || cPrecio === -1 || cRecib === -1 || cOpcion === -1) {
            throw new Error(WRONG_FILE_MSG);
        }
        const out = [];
        for (let i = hIdx + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            const mla = row[cMla] ? String(row[cMla]).trim() : '';
            if (!mla) continue;
            const sku = (cSku !== -1 && row[cSku] !== undefined && row[cSku] !== null) ? String(row[cSku]).trim() : '';
            out.push({
                sku,
                mla,
                opcion: cOpcion !== -1 ? String(row[cOpcion] || '').trim() : '',
                cuotas: cCuotas !== -1 ? String(row[cCuotas] || '').trim() : '',
                desc: cDesc !== -1 ? String(row[cDesc] || '').trim() : '',
                precio: toNum(row[cPrecio]),
                recibimos: toNum(row[cRecib]),
                costo: cCosto !== -1 ? toNum(row[cCosto]) : null,
                ventas: cVentas !== -1 ? (toNum(row[cVentas]) || 0) : 0,
                stock: cStock !== -1 ? (toNum(row[cStock]) || 0) : 0,
                full: cFull !== -1 ? String(row[cFull] || '').trim() : ''
            });
        }
        return out;
    }

    // Export plano del editor masivo (columnas ITEM_ID, SKU, TITLE, STOCK_FLEX,
    // STOCK_FULL, PRICE, SALE_PRICE). No trae Recibimos/Opción/Cuotas: el Recibimos
    // se estima luego con el % de comisión configurable.
    function parsePubPlano(rows, hIdx) {
        const headerUp = rows[hIdx].map(c => String(c).trim().toUpperCase());
        const cMla   = headerUp.indexOf('ITEM_ID');
        const cSku   = headerUp.indexOf('SKU');
        const cTitle = headerUp.indexOf('TITLE');
        const cFlex  = headerUp.indexOf('STOCK_FLEX');
        const cFullN = headerUp.indexOf('STOCK_FULL');
        const cPrice = headerUp.indexOf('PRICE');
        const cSale  = headerUp.indexOf('SALE_PRICE');
        if (cMla === -1 || cPrice === -1) throw new Error(WRONG_FILE_MSG);
        const out = [];
        for (let i = hIdx + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            const mla = row[cMla] ? String(row[cMla]).trim() : '';
            if (!mla.toUpperCase().startsWith('MLA')) continue; // saltea filas de ayuda/encabezado
            const sku = (cSku !== -1 && row[cSku] != null) ? String(row[cSku]).trim() : '';
            const sale = cSale !== -1 ? toNum(row[cSale]) : null;
            const price = toNum(row[cPrice]);
            const precio = (sale != null && sale > 0) ? sale : price; // precio de venta real (promo si hay)
            const fullN = cFullN !== -1 ? toNum(row[cFullN]) : null;
            out.push({
                sku,
                mla,
                opcion: '',
                cuotas: '',
                desc: cTitle !== -1 ? String(row[cTitle] || '').trim() : '',
                precio,
                recibimos: null,          // se estima en calcular() con la comisión
                costo: null,              // se completa desde el stock del sistema
                ventas: 0,
                stock: cFlex !== -1 ? (toNum(row[cFlex]) || 0) : 0,
                full: (fullN != null && fullN > 0) ? 'Sí' : 'No'
            });
        }
        return out;
    }

    // --- Parser: Ventas AR -> unidades del último mes por MLA ---
    function parseVentasAR(rows) {
        const hIdx = findHeaderRow(rows, ['SKU', 'PUBLICACIÓN']);
        const hIdx2 = hIdx !== -1 ? hIdx : findHeaderRow(rows, ['SKU', 'PUBLICACION']);
        if (hIdx2 === -1) throw new Error("El reporte de Ventas no tiene columna '# de publicación'.");
        const find = colFinder(rows[hIdx2]);
        const cMla = find('# DE PUBLICACIÓN', '# DE PUBLICACION', 'PUBLICACIÓN', 'PUBLICACION');
        const cUni = find('UNIDADES');
        const cEstado = find('ESTADO');
        const map = {};
        for (let i = hIdx2 + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row) continue;
            const mla = cMla !== -1 && row[cMla] ? String(row[cMla]).trim() : '';
            if (!mla || !mla.toUpperCase().startsWith('MLA')) continue;
            const estado = cEstado !== -1 ? String(row[cEstado] || '').toLowerCase() : '';
            if (estado.includes('cancel') || estado.includes('devol') || estado.includes('reembols')) continue;
            const u = cUni !== -1 ? (toNum(row[cUni]) || 0) : 0;
            map[mla] = (map[mla] || 0) + u;
        }
        return map;
    }

    // --- Parser: stock_todo -> costo s/IVA por SKU ---
    function parseStockTodo(rows) {
        const hIdx = findHeaderRow(rows, ['SKU', 'COSTO']);
        if (hIdx === -1) throw new Error("El stock del sistema no tiene columna 'SKU' y 'Costo'.");
        const find = colFinder(rows[hIdx]);
        const cSku = find('SKU');
        const cCosto = find('COSTO S/IVA', 'COSTO');
        const map = {};
        for (let i = hIdx + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row) continue;
            const sku = cSku !== -1 && row[cSku] ? String(row[cSku]).trim() : '';
            if (!sku) continue;
            const costo = cCosto !== -1 ? toNum(row[cCosto]) : null;
            if (costo !== null) map[sku] = costo;
        }
        return map;
    }

    // --- Setup drop zones ---
    function setupDropZone(dropZone, fileInput, statusElement, type) {
        dropZone.addEventListener('click', () => fileInput.click());
        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');
            if (e.dataTransfer.files.length) {
                fileInput.files = e.dataTransfer.files;
                handleFile(fileInput, statusElement, type);
            }
        });
        fileInput.addEventListener('change', () => {
            if (fileInput.files.length === 0) return;
            handleFile(fileInput, statusElement, type);
        });
    }

    function handleFile(input, statusElement, type) {
        const file = input.files[0];
        if (!file) return;
        statusElement.textContent = file.name;
        statusElement.classList.add('uploaded');
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const ws = workbook.Sheets[workbook.SheetNames.includes('Publicaciones') && type === 'pub'
                    ? 'Publicaciones' : workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                if (type === 'pub')         dataPub = parsePublicaciones(json);
                else if (type === 'ventas') dataVentas = parseVentasAR(json);
                else if (type === 'stock')  dataStock = parseStockTodo(json);
                checkReady();
            } catch (err) {
                Swal.fire('Error', 'No se pudo leer el archivo: ' + err.message, 'error');
                console.error(err);
                statusElement.textContent = 'Error al leer';
                statusElement.classList.remove('uploaded');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function checkReady() {
        btnProcess.disabled = !dataPub;
    }

    setupDropZone(dropPub, filePub, statusPub, 'pub');
    setupDropZone(dropVentas, fileVentas, statusVentas, 'ventas');
    setupDropZone(dropStock, fileStock, statusStock, 'stock');

    // --- Cálculo ---
    function calcular() {
        const comisionPct = (parseFloat(inputComision.value) || 0) / 100;
        finalResults = dataPub.map(p => {
            let costo = p.costo;
            // Completar costo desde stock_todo (s/IVA -> c/IVA) si falta
            if ((!costo || costo === 0) && dataStock && dataStock[p.sku] != null) {
                costo = dataStock[p.sku] * IVA;
            }
            // Si el archivo no trae "Recibimos" (formato plano), se estima restando
            // la comisión configurada al precio de venta.
            let recibimos = p.recibimos;
            let recibimosEstimado = false;
            if (recibimos == null && p.precio != null && comisionPct > 0) {
                recibimos = p.precio * (1 - comisionPct);
                recibimosEstimado = true;
            }
            const margen = (recibimos != null && costo != null && costo > 0) ? recibimos - costo : null;
            const margenPct = (margen != null && costo > 0) ? margen / costo : null;
            const ventasMes = dataVentas ? (dataVentas[p.mla] ?? null) : null;
            return { ...p, costo, recibimos, recibimosEstimado, margen, margenPct, ventasMes };
        });
    }

    // --- Clasificación de problemas ---
    function getFlag(item) {
        const min = (parseFloat(inputMargenMin.value) || 0) / 100;
        if (item.margen == null) return 'sindato';
        if (item.margen < 0) return 'perdida';
        if (item.margenPct != null && item.margenPct < min) return 'bajo';
        return 'ok';
    }

    // --- Agrupa por SKU y ordena grupos + filas según el criterio elegido ---
    // Devuelve un array plano de { r, firstOfGroup, groupIndex } manteniendo
    // todos los MLA de un mismo SKU juntos (el SKU se muestra una sola vez).
    function groupAndOrder(rows, orden) {
        const nz = (v) => (v == null ? Infinity : v);
        const nzDesc = (v) => (v == null ? -Infinity : v);

        // Orden dentro de cada grupo (mismo SKU)
        const withinCmp = (a, b) => {
            switch (orden) {
                case 'ventas-desc': return nzDesc(b.ventas) - nzDesc(a.ventas);
                case 'precio-desc': return nzDesc(b.precio) - nzDesc(a.precio);
                case 'margen-asc':  return nz(a.margen) - nz(b.margen);
                default:            return nz(a.margenPct) - nz(b.margenPct);
            }
        };

        const groups = new Map();
        rows.forEach(r => {
            const key = (r.sku && r.sku !== '0') ? r.sku : '__' + r.mla;
            if (!groups.has(key)) groups.set(key, []);
            groups.get(key).push(r);
        });

        const arr = [...groups.entries()].map(([key, rs]) => {
            rs.sort(withinCmp);
            let rep;
            switch (orden) {
                case 'ventas-desc': rep = rs.reduce((s, r) => s + (r.ventas || 0), 0); break;
                case 'stock-desc':  rep = Math.max(...rs.map(r => nzDesc(r.stock))); break;
                case 'precio-desc': rep = Math.max(...rs.map(r => nzDesc(r.precio))); break;
                case 'margen-asc':  rep = Math.min(...rs.map(r => nz(r.margen))); break;
                case 'sku':         rep = key; break;
                default:            rep = Math.min(...rs.map(r => nz(r.margenPct)));
            }
            return { key, rs, rep };
        });

        arr.sort((a, b) => {
            if (orden === 'sku') return String(a.rep).localeCompare(String(b.rep));
            if (orden === 'ventas-desc' || orden === 'stock-desc' || orden === 'precio-desc') return b.rep - a.rep;
            return a.rep - b.rep; // margen / margenPct ascendente (peor primero)
        });

        const out = [];
        arr.forEach((g, gi) => {
            g.rs.forEach((r, ri) => out.push({ r, firstOfGroup: ri === 0, groupIndex: gi }));
        });
        return out;
    }

    // --- Render ---
    function render() {
        const soloNeumaticos = chkSoloNeumaticos.checked;
        const soloProblemas = chkSoloProblemas.checked;
        const q = inputBuscar.value.trim().toLowerCase();

        let rows = finalResults.slice();

        if (soloNeumaticos) {
            rows = rows.filter(r => r.sku && r.sku !== '0' && r.costo != null && r.costo > 0);
        }
        if (q) {
            rows = rows.filter(r =>
                r.sku.toLowerCase().includes(q) ||
                r.mla.toLowerCase().includes(q) ||
                r.desc.toLowerCase().includes(q));
        }
        if (soloProblemas) {
            rows = rows.filter(r => { const f = getFlag(r); return f === 'perdida' || f === 'bajo'; });
        }

        const orden = selectOrden.value;
        const grouped = groupAndOrder(rows, orden);

        const flagColor = { perdida: 'var(--danger, #ef4444)', bajo: 'var(--warning, #f59e0b)', ok: 'var(--success, #22c55e)', sindato: 'var(--text-secondary)' };
        tbody.innerHTML = grouped.map(({ r, firstOfGroup, groupIndex }) => {
            const f = getFlag(r);
            const problemBg = f === 'perdida' ? 'rgba(239,68,68,0.12)'
                : f === 'bajo' ? 'rgba(245,158,11,0.12)' : '';
            // Franja alterna por grupo de SKU (solo si la fila no tiene color de problema)
            const groupBg = problemBg || (groupIndex % 2 === 1 ? 'rgba(255,255,255,0.03)' : '');
            const rowStyle = `${groupBg ? `background:${groupBg};` : ''}${firstOfGroup && groupIndex > 0 ? 'border-top:3px solid var(--accent);' : ''}`;
            const margenColor = flagColor[f];
            return `<tr style="${rowStyle}">
                <td style="font-weight:700;">${firstOfGroup ? (r.sku || '—') : ''}</td>
                <td style="font-size:0.8rem;">${r.mla}</td>
                <td>${r.opcion || '—'}</td>
                <td style="font-size:0.8rem;">${fmtCuotas(r.cuotas)}</td>
                <td style="font-size:0.85rem;">${r.desc || '—'}</td>
                <td style="text-align:right;">${fmtMoney(r.precio)}</td>
                <td style="text-align:right;" title="${r.recibimosEstimado ? 'Estimado con la comisión configurada' : ''}">${r.recibimosEstimado && r.recibimos != null ? '~' : ''}${fmtMoney(r.recibimos)}</td>
                <td style="text-align:right;">${fmtMoney(r.costo)}</td>
                <td style="text-align:right; color:${margenColor}; font-weight:600;">${fmtMoney(r.margen)}</td>
                <td style="text-align:center; color:${margenColor}; font-weight:600;">${fmtPct(r.margenPct)}</td>
                <td style="text-align:center;">${r.ventas ?? '—'}</td>
                <td style="text-align:center;">${r.ventasMes ?? '—'}</td>
                <td style="text-align:center;">${firstOfGroup ? (r.stock ?? '—') : ''}</td>
                <td style="text-align:center;">${r.full || '—'}</td>
            </tr>`;
        }).join('');

        // Estadísticas (sobre el universo filtrado por "solo neumáticos", no por búsqueda)
        let universe = finalResults.slice();
        if (soloNeumaticos) universe = universe.filter(r => r.sku && r.sku !== '0' && r.costo != null && r.costo > 0);
        const conMargen = universe.filter(r => r.margen != null);
        const perdida = conMargen.filter(r => r.margen < 0).length;
        const bajo = conMargen.filter(r => getFlag(r) === 'bajo').length;
        const margenProm = conMargen.length
            ? conMargen.reduce((s, r) => s + r.margenPct, 0) / conMargen.length : null;
        document.getElementById('stat-precios-total').textContent = universe.length;
        document.getElementById('stat-precios-perdida').textContent = perdida;
        document.getElementById('stat-precios-bajo').textContent = bajo;
        document.getElementById('stat-precios-margen-prom').textContent = margenProm != null ? fmtPct(margenProm) : '—';
    }

    // --- Eventos ---
    btnProcess.addEventListener('click', () => {
        btnProcess.disabled = true;
        btnProcess.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Analizando...';
        setTimeout(() => {
            calcular();
            render();
            resultsPanel.classList.remove('hidden');
            btnProcess.innerHTML = 'Analizar Precios';
            btnProcess.disabled = false;
            const estimado = pubFormatoPlano;
            Swal.fire({
                title: '¡Análisis completado!',
                html: `Se analizaron <b>${finalResults.length}</b> publicaciones.` +
                    (estimado ? `<br><br><span style="color:#b45309;">⚠️ Este archivo no trae Recibimos ni Opción/Cuotas. El <b>Recibimos y el margen son estimados</b> con la comisión configurada (${(parseFloat(inputComision.value)||0)}%)${dataStock ? '' : ', y hace falta subir el <b>Stock del Sistema</b> para el costo'}.</span>` : ''),
                icon: 'success',
                confirmButtonColor: '#3b82f6'
            });
        }, 300);
    });

    [inputBuscar, chkSoloProblemas, selectOrden, inputMargenMin, chkSoloNeumaticos].forEach(el => {
        const ev = (el.type === 'text' || el.type === 'number') ? 'input' : 'change';
        el.addEventListener(ev, () => { if (finalResults.length) render(); });
    });
    // La comisión cambia el Recibimos estimado: hay que recalcular, no solo re-renderizar.
    inputComision.addEventListener('input', () => { if (dataPub) { calcular(); render(); } });

    // --- Exportar a Excel ---
    btnDownload.addEventListener('click', () => {
        const soloNeumaticos = chkSoloNeumaticos.checked;
        let rows = finalResults.slice();
        if (soloNeumaticos) rows = rows.filter(r => r.sku && r.sku !== '0' && r.costo != null && r.costo > 0);
        if (rows.length === 0) {
            Swal.fire({ title: 'Sin datos', text: 'No hay publicaciones para exportar.', icon: 'warning', confirmButtonColor: '#3b82f6' });
            return;
        }
        // Agrupa por SKU igual que la vista (SKU una sola vez por grupo)
        const orden = selectOrden.value;
        const grouped = groupAndOrder(rows, orden);

        const excelData = grouped.map(({ r, firstOfGroup }) => ({
            'SKU': firstOfGroup ? r.sku : '',
            'MLA': r.mla,
            'Opción de venta': r.opcion,
            'Cuotas s/interés': r.cuotas,
            'Descripción': r.desc,
            'Precio comprador': r.precio,
            'Recibimos (neto)': r.recibimos,
            'Costo c/IVA': r.costo,
            'Margen $': r.margen,
            'Margen %': r.margenPct != null ? +(r.margenPct * 100).toFixed(1) : null,
            'Cant. ventas': r.ventas,
            'Ventas último mes': r.ventasMes,
            'Stock': firstOfGroup ? r.stock : '',
            'Full': r.full
        }));

        const XLib = window.XLSXStyle || XLSX;
        const ws = XLib.utils.json_to_sheet(excelData);
        ws['!cols'] = [
            { wch: 16 }, { wch: 16 }, { wch: 12 }, { wch: 22 }, { wch: 45 },
            { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 10 },
            { wch: 12 }, { wch: 16 }, { wch: 8 }, { wch: 6 }
        ];
        // Colorea: rojo/amarillo si hay problema; si no, franja alterna por grupo de SKU.
        if (window.XLSXStyle) {
            const numCols = Object.keys(excelData[0]).length;
            const GROUP_BANDS = ['FFFFFF', 'DDEBF7'];
            grouped.forEach(({ r, groupIndex }, i) => {
                const f = getFlag(r);
                let rgb = null;
                if (f === 'perdida') rgb = 'FFC7CE';
                else if (f === 'bajo') rgb = 'FFEB9C';
                else rgb = GROUP_BANDS[groupIndex % 2];
                for (let c = 0; c < numCols; c++) {
                    const addr = XLib.utils.encode_cell({ r: i + 1, c });
                    if (ws[addr]) ws[addr].s = { fill: { fgColor: { rgb } } };
                }
            });
        }
        const wb = XLib.utils.book_new();
        XLib.utils.book_append_sheet(wb, ws, 'Análisis de Precios');
        const dateStr = new Date().toISOString().split('T')[0];
        XLib.writeFile(wb, `Analisis_Precios_ML_${dateStr}.xlsx`);
    });
})();

// =============================================
// --- MÓDULO COMPETENCIA ML ---
// =============================================
// Por cada SKU nuestro guarda los links/MLA de los competidores con su precio
// (en localStorage, con export/import a Excel) y lo compara contra nuestra
// publicación más barata: ¿podemos bajar para ganarle sin perder el margen mínimo?
(function () {
    const STORE_KEY = 'competenciaML.v1';
    const IVA = 1.21;

    const $ = (id) => document.getElementById(id);
    const dropPub = $('drop-zone-comp-pub');
    const dropStock = $('drop-zone-comp-stock');
    if (!dropPub) return;
    const filePub = $('file-comp-pub');
    const fileStock = $('file-comp-stock');
    const statusPub = $('status-comp-pub');
    const statusStock = $('status-comp-stock');
    const btnProcess = $('btn-process-comp');
    const resultsPanel = $('results-panel-comp');
    const tbody = $('results-body-comp');
    const inBuscar = $('comp-buscar');
    const selVista = $('comp-vista');
    const selOrden = $('comp-orden');
    const chkSoloStock = $('comp-solo-stock');

    let pubs = null;    // [{sku, baseSku, kitN, mla, title, precio, precioLista, stock, full}]
    let costos = null;  // { SKU: {costo, desc} }
    // Reglas fijas: ganarle al rival por $1.000 y que el neto (lo que recibimos) quede por encima del costo c/IVA.
    const DIFERENCIA_RIVAL = 1000;
    // Comisión y envío de referencia (costos_ml.js, sacado del reporte Ventas AR)
    const CML = window.COSTOS_ML || { comisionDefault: 0.156, envioDefault: 18495, comisionPorMla: {}, envioPorSku: {}, envioPorMedida: {}, envioPorRodado: {} };
    let items = [];     // un elemento por SKU (base de la tabla)

    // ---------- helpers ----------
    const esc = (s) => String(s == null ? '' : s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const fmt$ = (n) => (n == null || isNaN(n)) ? '—' : (n < -0.5 ? '-$' : '$') + Math.abs(Math.round(n)).toLocaleString('es-AR');
    const fmtPct = (n) => (n == null || isNaN(n)) ? '—' : (n * 100).toFixed(1).replace('.', ',') + '%';
    const toNum = (v) => {
        if (v === null || v === undefined || v === '' || v === '-') return null;
        if (typeof v === 'number') return v;
        let s = String(v).replace(/[^0-9,.-]/g, '');
        if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');   // 98.000,50
        else s = s.replace(/[.,](?=\d{3}(\D|$))/g, '');                        // 98.000 / 98,000
        const n = parseFloat(s);
        return isNaN(n) ? null : n;
    };
    const normSku = (s) => String(s == null ? '' : s).replace(/^['"]+|['"]+$/g, '').trim().toUpperCase();
    const hoy = () => new Date().toISOString().slice(0, 10);
    const fmtFecha = (iso) => iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) : '';

    // Link o MLA → { mla, url }
    function parseRef(input) {
        const raw = String(input || '').trim();
        if (!raw) return null;
        const m = raw.match(/MLA[-_ ]?(\d{6,})/i) || raw.match(/^(\d{8,})$/);
        const mla = m ? 'MLA' + m[1] : '';
        let url = /^https?:\/\//i.test(raw) ? raw.split('#')[0] : '';
        if (!url && mla) url = 'https://articulo.mercadolibre.com.ar/MLA-' + mla.slice(3);
        if (!mla && !url) return null;
        return { mla, url };
    }

    // ---------- almacenamiento de competidores ----------
    function loadStore() {
        try { return JSON.parse(localStorage.getItem(STORE_KEY)) || {}; } catch (e) { return {}; }
    }
    let store = loadStore();
    function saveStore() {
        try { localStorage.setItem(STORE_KEY, JSON.stringify(store)); }
        catch (e) { Swal.fire('Atención', 'No se pudieron guardar los competidores en este navegador. Exportalos a Excel para no perderlos.', 'warning'); }
    }
    const compsOf = (sku) => store[sku] || [];
    function upsertComp(sku, comp) {
        const list = store[sku] || (store[sku] = []);
        const same = list.find(c => (comp.mla && c.mla === comp.mla) || (!comp.mla && c.url === comp.url));
        if (same) Object.assign(same, comp, { id: same.id });
        else list.push(Object.assign({ id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6) }, comp));
    }

    // ---------- parsers ----------
    function findHeader(rows, needles) {
        for (let i = 0; i < Math.min(rows.length, 25); i++) {
            const up = (rows[i] || []).map(c => String(c).trim().toUpperCase());
            if (needles.every(n => up.some(h => h === n || h.includes(n)))) return i;
        }
        return -1;
    }

    // Excel de "Modificar publicaciones" de ML (ITEM_ID, SKU, TITLE, PRICE, SALE_PRICE, STOCK...)
    function parsePublicaciones(rows) {
        const h = findHeader(rows, ['ITEM_ID', 'SKU', 'PRICE']);
        if (h === -1) throw new Error('No encontré las columnas ITEM_ID, SKU y PRICE. Bajá el Excel de Modificar publicaciones incluyendo los precios.');
        const up = rows[h].map(c => String(c).trim().toUpperCase());
        const c = (name) => up.indexOf(name);
        const cMla = c('ITEM_ID'), cSku = c('SKU'), cTitle = c('TITLE'), cPrice = c('PRICE'),
            cSale = c('SALE_PRICE'), cTotal = c('TOTAL_STOCK_ALL_STORES'), cFlex = c('STOCK_FLEX'), cFull = c('STOCK_FULL');
        const out = [];
        for (let i = h + 1; i < rows.length; i++) {
            const r = rows[i];
            if (!r) continue;
            const mla = String(r[cMla] || '').trim();
            if (!/^MLA\d+/i.test(mla)) continue;
            const sku = normSku(r[cSku]);
            if (!sku) continue;
            const lista = toNum(r[cPrice]);
            const promo = cSale !== -1 ? toNum(r[cSale]) : null;
            const precio = (promo != null && promo > 0) ? promo : lista;
            if (precio == null || precio <= 0) continue;
            const kit = sku.match(/^KITX(\d+)-(.+)$/);
            const stockCol = cTotal !== -1 ? cTotal : cFlex;
            out.push({
                sku, mla,
                baseSku: kit ? kit[2] : sku,
                kitN: kit ? parseInt(kit[1], 10) : 1,
                title: cTitle !== -1 ? String(r[cTitle] || '').trim() : '',
                precio, precioLista: lista, enPromo: promo != null && promo > 0 && promo < lista,
                stock: stockCol !== -1 ? (toNum(r[stockCol]) || 0) : 0,
                full: cFull !== -1 ? (toNum(r[cFull]) || 0) : 0
            });
        }
        if (!out.length) throw new Error('El archivo no tiene publicaciones con SKU y precio.');
        return out;
    }

    // stock_todo: SKU → costo y descripción
    function parseStockTodo(rows) {
        const h = findHeader(rows, ['SKU', 'COSTO']);
        if (h === -1) throw new Error("El stock del sistema no tiene columnas 'SKU' y 'Costo'.");
        const up = rows[h].map(c => String(c).trim().toUpperCase());
        const cSku = up.indexOf('SKU');
        let cCosto = up.findIndex(x => x.startsWith('COSTO S/IVA'));
        if (cCosto === -1) cCosto = up.findIndex(x => x.startsWith('COSTO'));
        const cDesc = up.findIndex(x => x.startsWith('DESCRIP'));
        const map = {};
        for (let i = h + 1; i < rows.length; i++) {
            const r = rows[i];
            if (!r) continue;
            const sku = normSku(r[cSku]);
            const costo = toNum(r[cCosto]);
            if (!sku || costo == null || costo <= 0) continue;
            map[sku] = { costo, desc: cDesc !== -1 ? String(r[cDesc] || '').trim() : '' };
        }
        return map;
    }

    // ---------- carga de archivos ----------
    function setupDrop(zone, input, status, type) {
        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault(); zone.classList.remove('drag-over');
            if (e.dataTransfer.files.length) readFile(e.dataTransfer.files[0], status, type);
        });
        input.addEventListener('change', () => { if (input.files.length) readFile(input.files[0], status, type); input.value = ''; });
    }
    function readFile(file, status, type) {
        status.textContent = file.name;
        status.classList.add('uploaded');
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                const name = type === 'pub' && wb.SheetNames.includes('Publicaciones') ? 'Publicaciones' : wb.SheetNames[0];
                const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
                if (type === 'pub') pubs = parsePublicaciones(rows);
                else costos = parseStockTodo(rows);
                guardarDatos(type, file.name);
                status.textContent = `${file.name} · cargado ahora`;
                btnProcess.disabled = !(pubs && costos);
                if (items.length) { construir(); render(); }
            } catch (err) {
                console.error(err);
                status.textContent = 'Error al leer';
                status.classList.remove('uploaded');
                if (type === 'pub') pubs = null; else costos = null;
                btnProcess.disabled = true;
                Swal.fire('Error', err.message || 'No se pudo leer el archivo.', 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }
    setupDrop(dropPub, filePub, statusPub, 'pub');
    setupDrop(dropStock, fileStock, statusStock, 'stock');

    // La página recuerda los últimos archivos subidos (ya procesados), así no hay que
    // volver a subirlos cada vez. Se reemplazan al subir uno nuevo.
    const DATA_KEY = 'competenciaML.datos.v1';
    let datosGuardados = {};   // { pub: {nombre, fecha}, stock: {nombre, fecha} }
    const fmtFechaHora = (iso) => {
        if (!iso) return '';
        const d = new Date(iso);
        return d.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit' }) + ' ' +
            d.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });
    };
    const diasDesde = (iso) => iso ? Math.floor((Date.now() - new Date(iso).getTime()) / 86400000) : null;
    function guardarDatos(type, nombre) {
        const meta = { nombre, fecha: new Date().toISOString() };
        datosGuardados[type] = meta;
        try {
            const d = JSON.parse(localStorage.getItem(DATA_KEY) || '{}');
            d[type] = { ...meta, datos: type === 'pub' ? pubs : costos };
            localStorage.setItem(DATA_KEY, JSON.stringify(d));
        } catch (e) { console.warn('No se pudieron guardar los archivos en el navegador', e); }
    }
    function restaurarDatos() {
        try {
            const d = JSON.parse(localStorage.getItem(DATA_KEY) || '{}');
            [['pub', statusPub], ['stock', statusStock]].forEach(([type, status]) => {
                if (!d[type] || !d[type].datos) return;
                if (type === 'pub') pubs = d[type].datos; else costos = d[type].datos;
                datosGuardados[type] = { nombre: d[type].nombre, fecha: d[type].fecha };
                status.textContent = `${d[type].nombre} · cargado el ${fmtFechaHora(d[type].fecha)}`;
                status.classList.add('uploaded');
            });
        } catch (e) { console.warn('No se pudieron leer los archivos guardados', e); }
        btnProcess.disabled = !(pubs && costos);
    }
    // Aviso si los archivos guardados ya tienen varios días
    function pintarAvisoDatos() {
        const el = $('comp-aviso-datos');
        const viejos = [['stock', 'el <b>Stock del Sistema</b> (costos)'], ['pub', 'las <b>Publicaciones</b> (nuestros precios)']]
            .map(([t, txt]) => ({ t, txt, d: diasDesde(datosGuardados[t] && datosGuardados[t].fecha) }))
            .filter(x => x.d != null && x.d >= 2);
        el.hidden = !viejos.length;
        el.innerHTML = viejos.length
            ? '<i class="fa-solid fa-clock-rotate-left"></i><span>' + viejos.map(x => `${x.txt} es de hace ${x.d} días`).join(' y ') +
              '. Subí el archivo actualizado arriba para que los cálculos estén al día.</span>'
            : '';
    }
    // ---------- Comisión y envío de referencia ----------
    // Medida "195/55R16" a partir de un título o descripción (acepta 195/55r16, 205 65 R16, 205/45zr17...)
    function medidaDe(txt) {
        const m = String(txt || '').match(/\b(\d{3})\s*[\/\s]\s*(\d{2})\s*(?:Z?R|-)?\s*(\d{2})(?!\d)/i);
        return m ? { medida: `${m[1]}/${m[2]}R${m[3]}`, rodado: 'R' + m[3] } : null;
    }
    const nVentas = (n) => `${n} venta${n === 1 ? '' : 's'}`;
    // Comisión de nuestra publicación y envío por neumático: real si hubo ventas, si no aproximado.
    function costosVenta(sku, mla, texto) {
        const out = { com: CML.comisionDefault, comFuente: 'aproximada (promedio de publicaciones sin cuotas)', comReal: false,
                      envio: CML.envioDefault, envioFuente: 'aproximado (promedio general)', envioReal: false };
        const c = CML.comisionPorMla[mla];
        if (c) Object.assign(out, { com: c[0], comFuente: `real de esta publicación (${nVentas(c[1])})`, comReal: true });
        const md = medidaDe(texto);
        const e = CML.envioPorSku[sku];
        const em = md && CML.envioPorMedida[md.medida];
        const er = md && CML.envioPorRodado[md.rodado];
        if (e) Object.assign(out, { envio: e[0], envioFuente: `real del SKU (${nVentas(e[1])})`, envioReal: true });
        else if (em) Object.assign(out, { envio: em[0], envioFuente: `aproximado por la medida ${md.medida} (${nVentas(em[1])})` });
        else if (er) Object.assign(out, { envio: er[0], envioFuente: `aproximado por el rodado ${md.rodado} (${nVentas(er[1])})` });
        return out;
    }

    // ---------- armado por SKU ----------
    function construir() {
        const bySku = new Map();
        pubs.forEach(p => {
            if (!bySku.has(p.baseSku)) bySku.set(p.baseSku, []);
            bySku.get(p.baseSku).push(p);
        });
        items = [...bySku.entries()].map(([sku, ps]) => {
            const simples = ps.filter(p => p.kitN === 1);
            const pool = simples.length ? simples : ps;
            // Nuestra mejor oferta: la publicación con menor precio por unidad
            const mejor = pool.reduce((a, b) => (b.precio / b.kitN < a.precio / a.kitN ? b : a));
            const info = costos[sku];
            const cv = costosVenta(sku, mejor.mla.toUpperCase(), (info && info.desc) + ' ' + mejor.title);
            return {
                ...cv,
                sku,
                desc: (info && info.desc) || mejor.title,
                title: mejor.title,
                mla: mejor.mla,
                precio: mejor.precio / mejor.kitN,
                enPromo: mejor.enPromo,
                esKit: mejor.kitN > 1,
                nPubs: ps.length,
                mlas: ps.map(p => p.mla),
                stock: Math.max(...ps.map(p => p.stock)) + Math.max(...ps.map(p => p.full)),
                costoBase: info ? info.costo : null,
                ventas: (CML.ventasPorSku || {})[sku] || 0
            };
        });
    }

    // Cálculo con la configuración actual + competidores guardados
    function evaluar(it) {
        const com = Math.min(Math.max(it.com, 0), 0.95);
        const envio = it.envio;
        const recib = (p) => p * (1 - com) - envio;
        const dif = DIFERENCIA_RIVAL;
        const costo = it.costoBase != null ? it.costoBase * IVA : null;
        const recibimos = recib(it.precio);
        const margen = costo != null ? recibimos - costo : null;
        const margenPct = costo ? margen / costo : null;
        // precio más bajo posible: a ese precio lo que recibimos es igual al costo
        const piso = costo != null ? (costo + envio) / (1 - com) : null;
        const comps = compsOf(it.sku);
        const conPrecio = comps.filter(c => c.precio > 0);
        const rival = conPrecio.length ? conPrecio.reduce((a, b) => (b.precio < a.precio ? b : a)) : null;
        const r = { costo, com, envio, recibimos, margen, margenPct, piso, comps, rival, estado: 'sin' };
        if (!rival) { r.estado = comps.length ? 'sinprecio' : 'sin'; return r; }
        if (it.precio <= rival.precio) {
            r.estado = 'ganamos';
            const techo = rival.precio - dif;
            if (techo > it.precio + 1) r.podriaSubir = techo;
            return r;
        }
        r.objetivo = Math.max(rival.precio - dif, 0);
        r.recibObjetivo = recib(r.objetivo);
        r.margenObjetivo = costo != null ? r.recibObjetivo - costo : null;
        r.margenObjetivoPct = costo ? r.margenObjetivo / costo : null;
        if (costo == null) r.estado = 'sincosto';
        else r.estado = r.recibObjetivo > costo ? 'bajar' : 'no';
        return r;
    }

    // ---------- render ----------
    const ORDEN = { bajar: 0, no: 1, sincosto: 2, sinprecio: 3, ganamos: 4, sin: 5 };

    // Cuánto más caros estamos que el competidor más barato (0 si ya somos los más baratos)
    const difPct = ({ it, ev }) => ev.rival && it.precio > ev.rival.precio ? (it.precio - ev.rival.precio) / ev.rival.precio : 0;
    // Impacto: lo que se vende × cuánto nos ganan (sin ventas cuenta como 1 para no quedar en cero)
    const impacto = (f) => (f.it.ventas + 1) * (difPct(f) || 0.0001) + f.it.stock / 1e6;
    function comparador(orden) {
        const nz = (v) => (v == null ? Infinity : v);
        switch (orden) {
            case 'dif': return (a, b) => difPct(b) - difPct(a) || b.it.ventas - a.it.ventas;
            case 'ventas': return (a, b) => b.it.ventas - a.it.ventas || b.it.stock - a.it.stock;
            case 'stock': return (a, b) => b.it.stock - a.it.stock;
            case 'margen': return (a, b) => nz(a.ev.margenPct) - nz(b.ev.margenPct);
            case 'sku': return (a, b) => a.it.sku.localeCompare(b.it.sku, 'es', { numeric: true });
            default: return (a, b) => (ORDEN[a.ev.estado] - ORDEN[b.ev.estado]) || (impacto(b) - impacto(a));
        }
    }

    function filasVisibles() {
        const q = inBuscar.value.trim().toLowerCase();
        const vista = selVista.value;
        return items
            .map(it => ({ it, ev: evaluar(it) }))
            .filter(({ it, ev }) => {
                if (chkSoloStock.checked && !(it.stock > 0)) return false;
                if (q && !(it.sku.toLowerCase().includes(q) || it.desc.toLowerCase().includes(q) ||
                    it.title.toLowerCase().includes(q) || it.mlas.some(m => m.toLowerCase().includes(q)) ||
                    ev.comps.some(c => (c.mla || '').toLowerCase().includes(q) || (c.vendedor || '').toLowerCase().includes(q)))) return false;
                switch (vista) {
                    case 'con': return ev.comps.length > 0;
                    case 'sin': return ev.comps.length === 0;
                    case 'bajar': return ev.estado === 'bajar';
                    case 'no': return ev.estado === 'no';
                    case 'ganamos': return ev.estado === 'ganamos';
                    default: return true;
                }
            })
            .sort(comparador(selOrden.value));
    }

    function sugerenciaHtml(it, ev) {
        switch (ev.estado) {
            case 'bajar':
                return `<span class="status-badge bg-success">Bajar a ${fmt$(ev.objetivo)}</span>
                    <div class="comp-sub">Nos gana por ${fmt$(it.precio - ev.rival.precio)}. A ${fmt$(ev.objetivo)} recibiríamos ${fmt$(ev.recibObjetivo)} neto, ${fmt$(ev.margenObjetivo)} más que el costo (${fmtPct(ev.margenObjetivoPct)})</div>`;
            case 'no':
                return `<span class="status-badge bg-danger">No conviene</span>
                    <div class="comp-sub">Para ganarle habría que ir a ${fmt$(ev.objetivo)}: recibiríamos ${fmt$(ev.recibObjetivo)} neto, ${fmt$(-ev.margenObjetivo)} menos que el costo. Lo mínimo sin perder es ${fmt$(ev.piso)}</div>`;
            case 'ganamos':
                return `<span class="status-badge bg-warning">Ya somos más baratos</span>
                    <div class="comp-sub">Por ${fmt$(ev.rival.precio - it.precio)}${ev.podriaSubir ? `. Se podría subir hasta ${fmt$(ev.podriaSubir)} y seguir ganando` : ''}</div>`;
            case 'sincosto': return `<span class="status-badge bg-info">Sin costo</span><div class="comp-sub">El SKU no está en el stock del sistema</div>`;
            case 'sinprecio': return `<span class="status-badge bg-info">Falta precio</span><div class="comp-sub">Cargá el precio del competidor</div>`;
            default: return `<span class="comp-muted">Sin competidores</span>`;
        }
    }

    const expandidos = new Set();   // SKU con la lista de competidores desplegada
    const MAX_CHIPS = 3;
    function compsHtml(it, ev) {
        const orden = ev.comps.slice().sort((a, b) => (a.precio > 0 ? a.precio : Infinity) - (b.precio > 0 ? b.precio : Infinity));
        const abierto = expandidos.has(it.sku);
        const visibles = abierto ? orden : orden.slice(0, MAX_CHIPS);
        const ocultos = orden.length - visibles.length;
        const chips = visibles.map(c => {
            const best = ev.rival && c.id === ev.rival.id;
            // Una sola línea: [🤖] vendedor (link a la publicación) [FULL] · precio · ✕
            const nombre = c.vendedor || c.mla || 'Publicación';
            const info = `${c.vendedor ? c.vendedor + ' · ' : ''}${c.mla || c.url}${c.auto ? ' · encontrado automáticamente' : ''}${c.fecha ? ' · precio del ' + fmtFecha(c.fecha) : ''}`;
            return `<div class="comp-chip${best ? ' best' : ''}" title="${esc(info)}">
                ${c.auto ? '<i class="fa-solid fa-robot comp-auto"></i>' : '<i class="fa-solid fa-user-pen comp-manual"></i>'}
                <a class="comp-name" href="${esc(c.url)}" target="_blank" rel="noopener noreferrer">${esc(nombre)}</a>
                ${c.full ? '<span class="comp-mini" title="Envío Full">FULL</span>' : ''}
                ${c.precioAnterior > 0 && c.precio > 0 && c.precioAnterior !== c.precio
                    ? (c.precio < c.precioAnterior
                        ? `<i class="fa-solid fa-arrow-down comp-baja" title="Bajó: antes ${fmt$(c.precioAnterior)}"></i>`
                        : `<i class="fa-solid fa-arrow-up comp-sube" title="Subió: antes ${fmt$(c.precioAnterior)}"></i>`)
                    : ''}
                <button class="comp-price" data-act="edit" data-sku="${esc(it.sku)}" data-id="${esc(c.id)}" title="Editar precio / datos">${c.precio > 0 ? fmt$(c.precio) : 'sin precio'}</button>
                <button class="comp-del" data-act="del" data-sku="${esc(it.sku)}" data-id="${esc(c.id)}" title="Quitar competidor"><i class="fa-solid fa-xmark"></i></button>
            </div>`;
        }).join('');
        const mas = ocultos > 0
            ? `<button class="comp-more" data-act="more" data-sku="${esc(it.sku)}">+${ocultos} más</button>`
            : (abierto && orden.length > MAX_CHIPS ? `<button class="comp-more" data-act="more" data-sku="${esc(it.sku)}">ver menos</button>` : '');
        return `${chips}${mas}<button class="comp-add" data-act="add" data-sku="${esc(it.sku)}"><i class="fa-solid fa-plus"></i> Competidor</button>`;
    }

    function render() {
        pintarAvisoDatos();
        const filas = filasVisibles();
        tbody.innerHTML = filas.map(({ it, ev }) => {
            const mColor = ev.margen == null ? '' : ev.margen < 0 ? 'text-danger' : 'text-success';
            const pubUrl = 'https://articulo.mercadolibre.com.ar/MLA-' + it.mla.replace(/^MLA/i, '');
            // Verde: podemos bajar y ganarle. Rojo: para ganarle (o ya hoy) el neto queda debajo del costo.
            const color = ev.estado === 'bajar' ? 'comp-verde'
                : (ev.estado === 'no' || (ev.margen != null && ev.margen < 0)) ? 'comp-rojo' : '';
            return `<tr class="comp-row comp-${ev.estado} ${color}">
                <td style="font-weight:700; white-space:nowrap;">${esc(it.sku)}</td>
                <td class="comp-prod">${esc(it.desc)}<div class="comp-sub">${it.nPubs} ${it.nPubs === 1 ? 'publicación' : 'publicaciones'}</div></td>
                <td style="text-align:center;">${it.stock}<div class="comp-sub" title="Neumáticos vendidos en Mercado Libre en los últimos ${CML.dias || 90} días">${it.ventas} vend.</div></td>
                <td style="text-align:right;">${fmt$(ev.costo)}</td>
                <td style="text-align:right; white-space:nowrap;">
                    <strong>${fmt$(it.precio)}</strong>${it.enPromo ? ' <span class="comp-tag">promo</span>' : ''}${it.esKit ? ' <span class="comp-tag">kit c/u</span>' : ''}
                    <div class="comp-sub"><a href="${pubUrl}" target="_blank" rel="noopener noreferrer">${esc(it.mla)}</a></div>
                </td>
                <td style="text-align:right; white-space:nowrap;">${fmt$(ev.recibimos)}
                    <div class="comp-sub" title="Comisión: ${esc(it.comFuente)}&#10;Envío: ${esc(it.envioFuente)}">com ${fmtPct(ev.com)}${it.comReal ? '' : '*'} · envío ${fmt$(ev.envio)}${it.envioReal ? '' : '*'}</div></td>
                <td style="text-align:right;" class="${mColor}"><strong>${fmt$(ev.margen)}</strong><div class="comp-sub">${fmtPct(ev.margenPct)}</div></td>
                <td class="comp-cell">${compsHtml(it, ev)}</td>
                <td style="text-align:right;">${fmt$(ev.piso)}</td>
                <td class="comp-sug">${sugerenciaHtml(it, ev)}</td>
            </tr>`;
        }).join('') || `<tr><td colspan="10" style="text-align:center; color:var(--text-secondary);">No hay filas para este filtro.</td></tr>`;

        // Estadísticas sobre todo el universo (respeta "solo con stock")
        const all = items.filter(it => !chkSoloStock.checked || it.stock > 0).map(it => evaluar(it));
        $('stat-comp-con').textContent = all.filter(e => e.comps.length).length;
        $('stat-comp-bajar').textContent = all.filter(e => e.estado === 'bajar').length;
        $('stat-comp-no').textContent = all.filter(e => e.estado === 'no').length;
        $('stat-comp-ganamos').textContent = all.filter(e => e.estado === 'ganamos').length;
    }

    // ---------- alta / edición de competidores ----------
    async function editarComp(sku, comp) {
        const it = items.find(x => x.sku === sku);
        const { value } = await Swal.fire({
            title: comp ? 'Editar competidor' : 'Agregar competidor',
            html: `<div class="comp-form">
                <div class="comp-form-prod">${esc(sku)} · ${esc(it ? it.desc : '')}${it ? `<br>Nuestro precio: <b>${fmt$(it.precio)}</b>` : ''}</div>
                <label>Link de la publicación o MLA</label>
                <input id="cf-ref" class="swal2-input" placeholder="https://articulo.mercadolibre.com.ar/MLA-... o MLA123456789" value="${esc(comp ? (comp.url || comp.mla) : '')}">
                <label>Precio del competidor ($)</label>
                <input id="cf-precio" class="swal2-input" inputmode="decimal" placeholder="98000" value="${comp && comp.precio ? Math.round(comp.precio) : ''}">
                <label>Vendedor (opcional)</label>
                <input id="cf-vend" class="swal2-input" placeholder="Nombre del vendedor" value="${esc(comp ? comp.vendedor || '' : '')}">
            </div>`,
            focusConfirm: false,
            showCancelButton: true,
            confirmButtonText: 'Guardar',
            cancelButtonText: 'Cancelar',
            confirmButtonColor: '#3b82f6',
            didOpen: () => { const el = document.getElementById(comp ? 'cf-precio' : 'cf-ref'); if (el) { el.focus(); el.select(); } },
            preConfirm: () => {
                const ref = parseRef(document.getElementById('cf-ref').value);
                if (!ref) { Swal.showValidationMessage('Pegá el link de la publicación o su número MLA'); return false; }
                const precio = toNum(document.getElementById('cf-precio').value);
                if (precio != null && precio < 0) { Swal.showValidationMessage('Precio inválido'); return false; }
                return { ...ref, precio: precio || null, vendedor: document.getElementById('cf-vend').value.trim() };
            }
        });
        if (!value) return;
        if (comp) {
            // si cambió el MLA, se reemplaza el registro
            const list = store[sku] || [];
            const idx = list.findIndex(c => c.id === comp.id);
            const prevPrecio = comp.precio;
            const upd = { ...comp, ...value, fecha: value.precio !== prevPrecio ? hoy() : comp.fecha };
            if (idx !== -1) list[idx] = upd;
        } else {
            upsertComp(sku, { ...value, fecha: hoy() });
        }
        saveStore();
        render();
    }

    tbody.addEventListener('click', async (e) => {
        const btn = e.target.closest('button[data-act]');
        if (!btn) return;
        const sku = btn.dataset.sku;
        const act = btn.dataset.act;
        const comp = (store[sku] || []).find(c => c.id === btn.dataset.id);
        if (act === 'add') editarComp(sku, null);
        else if (act === 'more') { expandidos.has(sku) ? expandidos.delete(sku) : expandidos.add(sku); render(); }
        else if (act === 'edit' && comp) editarComp(sku, comp);
        else if (act === 'del' && comp) {
            const ok = await Swal.fire({
                title: '¿Quitar competidor?',
                html: `${esc(comp.mla || comp.url)} de ${esc(sku)}` + (comp.auto && comp.mla ? '<br><small>No se va a volver a agregar en las próximas búsquedas automáticas.</small>' : ''),
                icon: 'question',
                showCancelButton: true, confirmButtonText: 'Quitar', cancelButtonText: 'Cancelar', confirmButtonColor: '#ef4444'
            });
            if (!ok.isConfirmed) return;
            store[sku] = (store[sku] || []).filter(c => c.id !== comp.id);
            if (!store[sku].length) delete store[sku];
            if (comp.auto && comp.mla) ignorar(sku, comp.mla);
            saveStore();
            render();
        }
    });

    // ---------- búsqueda automática en Mercado Libre (API oficial, vía /api/ml/*) ----------
    // Competidores que el usuario quitó: la búsqueda automática no los vuelve a agregar.
    // Nuestras otras cuentas de ML: nunca cuentan como competencia (la cuenta conectada se excluye sola).
    const VENDEDORES_PROPIOS = ['JUANCMARGARIASRL'];
    const IGNORE_KEY = 'competenciaML.ignorados.v1';
    let ignorados = (() => { try { return JSON.parse(localStorage.getItem(IGNORE_KEY)) || {}; } catch (e) { return {}; } })();
    function ignorar(sku, mla) {
        (ignorados[sku] = ignorados[sku] || []).includes(mla) || ignorados[sku].push(mla);
        try { localStorage.setItem(IGNORE_KEY, JSON.stringify(ignorados)); } catch (e) { /* sin almacenamiento */ }
    }
    const estaIgnorado = (sku, mla) => (ignorados[sku] || []).includes(mla);

    const mlEstado = $('comp-ml-estado');
    const btnMlConectar = $('btn-comp-ml-conectar');
    const btnMlBuscar = $('btn-comp-ml-buscar');
    const mlProg = $('comp-ml-progreso');
    const mlProgFill = $('comp-ml-barra-fill');
    const mlProgTxt = $('comp-ml-progreso-txt');
    let mlConectado = false;
    let mlBuscando = false;

    function pintarEstadoMl(st) {
        mlConectado = !!(st && st.conectado);
        btnMlConectar.hidden = mlConectado;
        btnMlBuscar.disabled = !mlConectado || mlBuscando;
        if (!st) {
            mlEstado.innerHTML = 'La búsqueda automática funciona en la página publicada (<b>analizador-stock-ml.vercel.app</b>), no en una copia local.';
            btnMlConectar.hidden = true;
        } else if (mlConectado) {
            mlEstado.innerHTML = `Conectado a Mercado Libre${st.nickname ? ` como <b>${esc(st.nickname)}</b>` : ''}` +
                (typeof ultima !== 'undefined' && ultima && ultima.fecha ? ` · precios actualizados el <b>${fmtFechaHora(ultima.fecha)}</b> (se actualizan solos una vez por día)` : '') +
                ` · <a href="#" id="comp-ml-salir">desconectar</a>`;
        } else if (st.configurado === false) {
            mlEstado.innerHTML = `Falta configurar la app de Mercado Libre en Vercel (ML_CLIENT_ID y ML_CLIENT_SECRET). URL de retorno: <code>${esc(st.redirectUri || '')}</code>`;
            btnMlConectar.hidden = true;
        } else {
            mlEstado.textContent = 'No conectado: conectá la cuenta de Mercado Libre para buscar competidores automáticamente.';
        }
    }
    async function consultarEstadoMl() {
        try {
            const r = await fetch('/api/ml/status', { credentials: 'same-origin', cache: 'no-store' });
            const ct = r.headers.get('content-type') || '';
            pintarEstadoMl(r.ok && ct.includes('json') ? await r.json() : null);
        } catch (e) { pintarEstadoMl(null); }
    }
    consultarEstadoMl();

    mlEstado.addEventListener('click', async (e) => {
        if (e.target.id !== 'comp-ml-salir') return;
        e.preventDefault();
        await fetch('/api/ml/logout', { method: 'POST', credentials: 'same-origin' }).catch(() => {});
        consultarEstadoMl();
    });

    // El login se hace en una ventana emergente para no perder los Excel ya cargados.
    btnMlConectar.addEventListener('click', () => {
        const w = window.open('/api/ml/login', 'ml-login', 'width=620,height=760');
        if (!w) location.href = '/api/ml/login';
    });
    window.addEventListener('message', (e) => {
        if (e.origin !== location.origin || !e.data || e.data.tipo !== 'ml-login') return;
        if (e.data.ok) {
            // Recién conectado y con datos cargados: arranca la búsqueda sola.
            consultarEstadoMl().then(() => {
                if (items.length) buscarCompetidoresMl();
                else Swal.fire({ title: 'Mercado Libre conectado', icon: 'success', timer: 1600, showConfirmButton: false });
            });
        }
        else Swal.fire('No se pudo conectar', e.data.mensaje || '', 'error');
    });

    const TAMANO_LOTE = 10;
    // Última actualización de precios y lo que cambió en ella
    const ULTIMA_KEY = 'competenciaML.ultima.v1';
    let ultima = (() => { try { return JSON.parse(localStorage.getItem(ULTIMA_KEY)) || null; } catch (e) { return null; } })();
    const HORAS_ACTUALIZACION = 20;   // "una vez por día"
    const tocaActualizar = () => !ultima || !ultima.fecha || (Date.now() - new Date(ultima.fecha).getTime()) > HORAS_ACTUALIZACION * 3600000;

    btnMlBuscar.addEventListener('click', () => buscarCompetidoresMl());
    async function buscarCompetidoresMl(opciones = {}) {
        if (!items.length || mlBuscando || !mlConectado) return;
        const automatica = !!opciones.automatica;
        const lista = items.filter(it => !chkSoloStock.checked || it.stock > 0);
        const lotes = [];
        for (let i = 0; i < lista.length; i += TAMANO_LOTE) lotes.push(lista.slice(i, i + TAMANO_LOTE));

        mlBuscando = true;
        btnMlBuscar.disabled = true;
        mlProg.hidden = false;
        const resumen = { conComp: 0, nuevos: 0, sinCatalogo: [], sinComp: 0, actualizados: 0, errores: 0, cambios: [], hayAnterior: !!(ultima && ultima.fecha), noActualizables: 0 };
        let hechos = 0;
        let abortado = null;
        for (const lote of lotes) {
            mlProgTxt.textContent = `${automatica ? 'Actualización diaria: ' : ''}Buscando en Mercado Libre... ${hechos} de ${lista.length} SKU`;
            mlProgFill.style.width = `${Math.round(hechos / lista.length * 100)}%`;
            // competidores cargados a mano en este lote: se les actualiza el precio
            const seguir = lote.flatMap(it => compsOf(it.sku).filter(c => !c.auto && c.mla).map(c => c.mla));
            let data;
            try {
                const r = await fetch('/api/ml/competidores', {
                    method: 'POST', credentials: 'same-origin',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ skus: lote.map(it => ({ sku: it.sku, mlas: it.mlas })), seguir })
                });
                data = await r.json().catch(() => ({}));
                if (r.status === 401) { abortado = data.error || 'Se perdió la conexión con Mercado Libre.'; break; }
                if (!r.ok) throw new Error(data.error || `Error ${r.status}`);
            } catch (err) {
                console.error(err);
                resumen.errores += lote.length;
                hechos += lote.length;
                continue;
            }
            aplicarResultadosMl(lote, data, resumen);
            hechos += lote.length;
            saveStore();
            render();
        }
        mlBuscando = false;
        mlProg.hidden = true;
        consultarEstadoMl();
        if (abortado) { Swal.fire('Búsqueda interrumpida', abortado, 'warning'); return; }
        if (resumen.errores < lista.length) {
            ultima = { fecha: new Date().toISOString(), cambios: resumen.cambios, anterior: ultima ? ultima.fecha : null };
            try { localStorage.setItem(ULTIMA_KEY, JSON.stringify(ultima)); } catch (e) { /* sin almacenamiento */ }
            pintarCambios();
            consultarEstadoMl();
        }
        if (automatica) {
            const c = resumen.cambios;
            const n = (t) => c.filter(x => x.tipo === t).length;
            Swal.fire({
                toast: true, position: 'top-end', timer: 7000, showConfirmButton: false, icon: 'success',
                title: 'Precios de la competencia actualizados',
                text: resumen.hayAnterior ? `${n('bajo')} bajaron · ${n('subio')} subieron · ${n('nuevo')} nuevos · ${n('salio')} dejaron de vender` : ''
            });
            return;
        }
        Swal.fire({
            title: 'Búsqueda terminada',
            html: `Se encontraron competidores para <b>${resumen.conComp}</b> SKU (${resumen.nuevos} publicaciones nuevas).` +
                (resumen.actualizados ? `<br>Se actualizó el precio de <b>${resumen.actualizados}</b> competidores cargados a mano.` : '') +
                (resumen.noActualizables ? `<br><span style="color:#b45309;">${resumen.noActualizables} competidores cargados a mano no están en el catálogo de ML: su precio no se puede actualizar solo (editalo a mano si cambió).</span>` : '') +
                (resumen.sinComp ? `<br>${resumen.sinComp} SKU están en catálogo pero sin otros vendedores.` : '') +
                (resumen.sinCatalogo.length ? `<br><span style="color:#b45309;">${resumen.sinCatalogo.length} SKU no están en el catálogo de ML: a esos hay que cargarles los competidores a mano.</span>` : '') +
                (resumen.errores ? `<br><span style="color:#ef4444;">${resumen.errores} SKU no se pudieron consultar (probá de nuevo).</span>` : ''),
            icon: resumen.errores ? 'warning' : 'success', confirmButtonColor: '#3b82f6'
        });
    }

    function aplicarResultadosMl(lote, data, resumen) {
        const res = data.resultados || {};
        const seg = data.seguidos || {};
        lote.forEach(it => {
            const r = res[it.sku];
            if (!r) return;
            const manuales = (store[it.sku] || []).filter(c => !c.auto);
            // Actualiza precio de los cargados a mano
            const anotar = (tipo, c, antes, ahora) => resumen.cambios.push({
                tipo, sku: it.sku, desc: it.desc, vendedor: c.vendedor || '', mla: c.mla, url: c.url,
                antes: antes || null, ahora: ahora || null, nuestro: it.precio
            });
            // ML no deja consultar publicaciones ajenas una por una, pero si el competidor
            // manual está en el catálogo, su precio actual viene en la búsqueda por catálogo.
            const delCatalogo = new Map((r.competidores || []).map(c => [c.mla, c]));
            manuales.forEach(c => {
                const s = (c.mla && delCatalogo.get(c.mla)) || (c.mla && seg[c.mla]);
                delete c.precioAnterior;
                if (s && s.precio > 0) {
                    if (s.precio !== c.precio) {
                        c.fecha = hoy();
                        if (c.precio > 0) { c.precioAnterior = c.precio; anotar(s.precio < c.precio ? 'bajo' : 'subio', c, c.precio, s.precio); }
                    }
                    c.precio = s.precio;
                    if (!c.vendedor && s.vendedor) c.vendedor = s.vendedor;
                    resumen.actualizados++;
                } else if (c.mla) resumen.noActualizables++;
            });
            // Los automáticos se reemplazan por los de esta búsqueda
            const previos = new Map((store[it.sku] || []).filter(c => c.auto).map(c => [c.mla, c]));
            const manualesMla = new Set(manuales.map(c => c.mla));
            const nuevos = (r.competidores || [])
                .filter(c => !manualesMla.has(c.mla) && !estaIgnorado(it.sku, c.mla) &&
                    !VENDEDORES_PROPIOS.includes(String(c.vendedor || '').toUpperCase()))
                .map(c => {
                    const prev = previos.get(c.mla);
                    if (!prev) resumen.nuevos++;
                    const comp = {
                        id: prev ? prev.id : Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
                        mla: c.mla,
                        url: 'https://articulo.mercadolibre.com.ar/MLA-' + c.mla.replace(/^MLA/, ''),
                        precio: c.precio,
                        vendedor: c.vendedor || '',
                        full: !!c.full,
                        auto: true,
                        fecha: prev && prev.precio === c.precio ? prev.fecha : hoy()
                    };
                    // Solo se registran cambios si ya había una búsqueda anterior para comparar
                    if (resumen.hayAnterior) {
                        if (prev && prev.precio > 0 && prev.precio !== c.precio) {
                            comp.precioAnterior = prev.precio;
                            anotar(c.precio < prev.precio ? 'bajo' : 'subio', comp, prev.precio, c.precio);
                        } else if (!prev && previos.size) anotar('nuevo', comp, null, c.precio);
                    }
                    return comp;
                });
            if (resumen.hayAnterior) {
                const ahora = new Set(nuevos.map(c => c.mla));
                previos.forEach((c, mla) => { if (!ahora.has(mla) && !estaIgnorado(it.sku, mla)) anotar('salio', c, c.precio, null); });
            }
            const todos = manuales.concat(nuevos);
            if (todos.length) store[it.sku] = todos; else delete store[it.sku];
            if (nuevos.length) resumen.conComp++;
            else if (r.motivo === 'sin_catalogo' || r.motivo === 'no_encontrada') resumen.sinCatalogo.push(it.sku);
            else resumen.sinComp++;
        });
    }

    // Panel "Cambios en la última actualización"
    function pintarCambios() {
        const box = $('comp-cambios');
        if (!ultima || !ultima.fecha || !ultima.anterior) { box.hidden = true; return; }
        const c = ultima.cambios || [];
        const n = (t) => c.filter(x => x.tipo === t).length;
        box.hidden = false;
        $('comp-cambios-resumen').innerHTML =
            `<i class="fa-solid fa-chart-line"></i> Cambios de la competencia entre el ${fmtFechaHora(ultima.anterior)} y el ${fmtFechaHora(ultima.fecha)}: ` +
            `<span class="cb-baja">↓ ${n('bajo')} bajaron</span> · <span class="cb-sube">↑ ${n('subio')} subieron</span> · ` +
            `<span>＋ ${n('nuevo')} nuevos</span> · <span>− ${n('salio')} dejaron de vender</span>`;
        if (!c.length) { $('comp-cambios-body').innerHTML = '<p class="comp-muted">No hubo cambios de precio.</p>'; return; }
        const ordenTipo = { bajo: 0, nuevo: 1, subio: 2, salio: 3 };
        const variacion = (x) => x.antes && x.ahora ? (x.ahora - x.antes) / x.antes : 0;
        const filas = c.slice().sort((a, b) => ordenTipo[a.tipo] - ordenTipo[b.tipo] || variacion(a) - variacion(b)).slice(0, 150);
        const TIPO = { bajo: '↓ Bajó', subio: '↑ Subió', nuevo: '＋ Nuevo', salio: '− Dejó de vender' };
        $('comp-cambios-body').innerHTML = `<table class="comp-cambios-tabla"><thead><tr>
                <th>Cambio</th><th>SKU</th><th>Producto</th><th>Vendedor</th><th style="text-align:right;">Antes</th>
                <th style="text-align:right;">Ahora</th><th style="text-align:right;">Variación</th><th>¿Nos gana?</th></tr></thead><tbody>` +
            filas.map(x => {
                const nosGana = x.ahora && x.nuestro > x.ahora;
                return `<tr class="cb-${x.tipo}">
                    <td>${TIPO[x.tipo]}</td><td>${esc(x.sku)}</td><td>${esc(x.desc)}</td>
                    <td><a href="${esc(x.url)}" target="_blank" rel="noopener noreferrer">${esc(x.vendedor || x.mla)}</a></td>
                    <td style="text-align:right;">${fmt$(x.antes)}</td><td style="text-align:right;">${fmt$(x.ahora)}</td>
                    <td style="text-align:right;">${x.antes && x.ahora ? fmtPct(variacion(x)) : '—'}</td>
                    <td>${x.ahora ? (nosGana ? '<b class="text-danger">Sí</b>' : 'No') : '—'}</td></tr>`;
            }).join('') + '</tbody></table>' + (c.length > filas.length ? `<p class="comp-muted">Se muestran ${filas.length} de ${c.length} cambios.</p>` : '');
    }

    // ---------- botones ----------
    btnProcess.addEventListener('click', () => {
        construir();
        render();
        resultsPanel.classList.remove('hidden');
        // Con Mercado Libre conectado, los competidores se buscan solos (sin tocar el botón).
        if (mlConectado) {
            $('results-panel-comp').scrollIntoView({ behavior: 'smooth', block: 'start' });
            buscarCompetidoresMl();
            return;
        }
        const sinCosto = items.filter(it => it.costoBase == null).length;
        const cargados = Object.keys(store).length;
        Swal.fire({
            title: '¡Listo!',
            html: `Se armaron <b>${items.length}</b> SKU a partir de ${pubs.length} publicaciones.` +
                (sinCosto ? `<br><span style="color:#b45309;">${sinCosto} SKU no están en el stock del sistema (sin costo).</span>` : '') +
                `<br>Comisión real en <b>${items.filter(i => i.comReal).length}</b> SKU y envío real en <b>${items.filter(i => i.envioReal).length}</b>; el resto es aproximado (marcado con *).` +
                `<br><br>${cargados ? `Hay competidores guardados para <b>${cargados}</b> SKU.` : 'Todavía no hay competidores cargados: usá el botón <b>+ Competidor</b> en cada fila o <b>Importar competidores</b>.'}`,
            icon: 'success', confirmButtonColor: '#3b82f6'
        });
    });

    [inBuscar, selVista, chkSoloStock, selOrden].forEach(el => {
        const ev = (el.type === 'text' || el.type === 'number') ? 'input' : 'change';
        el.addEventListener(ev, () => { if (items.length) render(); });
    });

    // Exportar competidores (respaldo / pasar a otra PC)
    $('btn-comp-export').addEventListener('click', () => {
        const rows = [];
        Object.keys(store).sort().forEach(sku => {
            const it = items.find(x => x.sku === sku);
            store[sku].forEach(c => rows.push({
                'SKU': sku, 'Producto': it ? it.desc : '', 'Link': c.url, 'MLA': c.mla,
                'Vendedor': c.vendedor || '', 'Precio': c.precio || '', 'Fecha precio': c.fecha || ''
            }));
        });
        if (!rows.length) { Swal.fire({ title: 'Sin competidores', text: 'Todavía no cargaste ningún competidor.', icon: 'info', confirmButtonColor: '#3b82f6' }); return; }
        const ws = XLSX.utils.json_to_sheet(rows);
        ws['!cols'] = [{ wch: 14 }, { wch: 40 }, { wch: 60 }, { wch: 16 }, { wch: 20 }, { wch: 12 }, { wch: 12 }];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Competidores');
        XLSX.writeFile(wb, `Competidores_ML_${hoy()}.xlsx`);
    });

    // Importar competidores: columnas SKU + (Link o MLA) [+ Vendedor, Precio, Fecha precio]
    const fileImport = $('file-comp-import');
    $('btn-comp-import').addEventListener('click', () => fileImport.click());
    fileImport.addEventListener('change', () => {
        const file = fileImport.files[0];
        fileImport.value = '';
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
                const h = findHeader(rows, ['SKU']);
                if (h === -1) throw new Error("El Excel necesita una columna 'SKU' y otra 'Link' o 'MLA'.");
                const up = rows[h].map(c => String(c).trim().toUpperCase());
                const cSku = up.indexOf('SKU');
                const cLink = up.findIndex(x => x.startsWith('LINK') || x === 'URL');
                const cMla = up.findIndex(x => x === 'MLA' || x.startsWith('MLA'));
                const cVend = up.findIndex(x => x.startsWith('VENDEDOR'));
                const cPrecio = up.findIndex(x => x.startsWith('PRECIO'));
                const cFecha = up.findIndex(x => x.startsWith('FECHA'));
                if (cLink === -1 && cMla === -1) throw new Error("El Excel necesita una columna 'Link' o 'MLA'.");
                let n = 0, skus = new Set();
                for (let i = h + 1; i < rows.length; i++) {
                    const r = rows[i];
                    const sku = normSku(r[cSku]);
                    const ref = parseRef((cLink !== -1 && r[cLink]) ? r[cLink] : (cMla !== -1 ? r[cMla] : ''));
                    if (!sku || !ref) continue;
                    const precio = cPrecio !== -1 ? toNum(r[cPrecio]) : null;
                    const fecha = cFecha !== -1 && /^\d{4}-\d{2}-\d{2}$/.test(String(r[cFecha])) ? String(r[cFecha]) : hoy();
                    upsertComp(sku, { ...ref, precio: precio || null, vendedor: cVend !== -1 ? String(r[cVend] || '').trim() : '', fecha });
                    n++; skus.add(sku);
                }
                saveStore();
                if (items.length) render();
                Swal.fire({ title: 'Importado', html: `Se cargaron <b>${n}</b> competidores para <b>${skus.size}</b> SKU.`, icon: 'success', confirmButtonColor: '#3b82f6' });
            } catch (err) {
                console.error(err);
                Swal.fire('Error', err.message || 'No se pudo leer el archivo.', 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    });

    // Descargar el análisis completo: Excel con 3 hojas (Resumen, Por SKU, Competidores)
    $('btn-download-comp').addEventListener('click', () => {
        const filas = filasVisibles();
        if (!filas.length) { Swal.fire({ title: 'Sin datos', text: 'No hay filas para exportar.', icon: 'warning', confirmButtonColor: '#3b82f6' }); return; }
        const XLib = window.XLSXStyle || XLSX;
        const conEstilo = !!window.XLSXStyle;
        const ESTADO = { bajar: 'Podemos bajar', no: 'No conviene bajar', ganamos: 'Ya somos más baratos', sincosto: 'Sin costo', sinprecio: 'Falta precio del competidor', sin: 'Sin competidores' };
        const COLOR = { bajar: 'C6EFCE', no: 'FFC7CE', ganamos: 'FFF2CC' };
        const MONEDA = '"$" #,##0;[Red]-"$" #,##0';
        const PCT = '0.0%;[Red]-0.0%';
        const urlMla = (mla) => 'https://articulo.mercadolibre.com.ar/MLA-' + String(mla).replace(/^MLA/i, '');
        const r0 = (n) => (n == null || isNaN(n)) ? null : Math.round(n);
        const colorFila = (ev) => (ev.margen != null && ev.margen < 0) ? COLOR.no : COLOR[ev.estado];

        // Arma una hoja a partir de columnas {titulo, ancho, fmt?, link?} y filas de valores
        function hoja(columnas, filasDatos, colores) {
            const aoa = [columnas.map(c => c.titulo), ...filasDatos.map(f => f.map(v => (v && typeof v === 'object') ? v.v : v))];
            const ws = XLib.utils.aoa_to_sheet(aoa);
            ws['!cols'] = columnas.map(c => ({ wch: c.ancho }));
            ws['!autofilter'] = { ref: XLib.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: filasDatos.length, c: columnas.length - 1 } }) };
            filasDatos.forEach((f, i) => f.forEach((v, c) => {
                const cell = ws[XLib.utils.encode_cell({ r: i + 1, c })];
                if (!cell) return;
                if (columnas[c].fmt && typeof cell.v === 'number') cell.z = columnas[c].fmt;
                if (v && typeof v === 'object' && v.link) cell.l = { Target: v.link, Tooltip: 'Abrir en Mercado Libre' };
                if (conEstilo) {
                    const s = { border: { bottom: { style: 'thin', color: { rgb: 'D9D9D9' } } }, alignment: { vertical: 'center' } };
                    if (colores && colores[i]) s.fill = { fgColor: { rgb: colores[i] } };
                    if (columnas[c].fmt) s.numFmt = columnas[c].fmt;
                    if (v && typeof v === 'object' && v.link) s.font = { color: { rgb: '0563C1' }, underline: true };
                    if (v && typeof v === 'object' && v.bold) s.font = Object.assign(s.font || {}, { bold: true });
                    cell.s = s;
                }
            }));
            if (conEstilo) columnas.forEach((c, i) => {
                const cell = ws[XLib.utils.encode_cell({ r: 0, c: i })];
                if (cell) cell.s = {
                    font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1F3864' } },
                    alignment: { vertical: 'center', horizontal: 'center', wrapText: true },
                    border: { bottom: { style: 'medium', color: { rgb: '000000' } } }
                };
            });
            ws['!rows'] = [{ hpt: 32 }];
            return ws;
        }

        // ---- Hoja "Por SKU" ----
        const colSku = [
            { titulo: 'Estado', ancho: 20 }, { titulo: 'SKU', ancho: 16 }, { titulo: 'Producto', ancho: 42 },
            { titulo: 'Stock', ancho: 7 }, { titulo: 'Vendidos 90 días', ancho: 9 }, { titulo: 'Costo c/IVA', ancho: 13, fmt: MONEDA },
            { titulo: 'Nuestro precio', ancho: 14, fmt: MONEDA }, { titulo: 'Nuestra publicación', ancho: 17 },
            { titulo: 'Comisión ML', ancho: 11, fmt: PCT }, { titulo: 'Envío por unidad', ancho: 13, fmt: MONEDA },
            { titulo: 'Recibimos (neto)', ancho: 14, fmt: MONEDA }, { titulo: 'Ganancia actual', ancho: 14, fmt: MONEDA },
            { titulo: 'Ganancia actual %', ancho: 11, fmt: PCT },
            { titulo: 'Competidores', ancho: 8 }, { titulo: 'Nuestra posición', ancho: 11 },
            { titulo: 'Competidor más barato', ancho: 24 }, { titulo: 'Su precio', ancho: 13, fmt: MONEDA },
            { titulo: 'Diferencia vs. él', ancho: 13, fmt: MONEDA }, { titulo: 'Diferencia %', ancho: 10, fmt: PCT },
            { titulo: 'Precio sugerido', ancho: 14, fmt: MONEDA }, { titulo: 'Neto con precio sugerido', ancho: 14, fmt: MONEDA },
            { titulo: 'Ganancia con precio sugerido', ancho: 14, fmt: MONEDA },
            { titulo: 'Mínimo sin perder', ancho: 14, fmt: MONEDA },
            { titulo: 'Origen comisión', ancho: 34 }, { titulo: 'Origen envío', ancho: 34 }
        ];
        const datosSku = filas.map(({ it, ev }) => {
            const conPrecio = ev.comps.filter(c => c.precio > 0);
            const posicion = conPrecio.length ? `${conPrecio.filter(c => c.precio < it.precio).length + 1} de ${conPrecio.length + 1}` : '';
            const baja = ev.estado === 'bajar';
            return [
                { v: ESTADO[ev.estado] || '', bold: true }, it.sku, it.desc, it.stock, it.ventas, r0(ev.costo),
                r0(it.precio), { v: it.mla, link: urlMla(it.mla) },
                ev.com, r0(ev.envio), r0(ev.recibimos), r0(ev.margen), ev.margenPct,
                ev.comps.length, posicion,
                ev.rival ? { v: ev.rival.vendedor || ev.rival.mla || 'Ver publicación', link: ev.rival.url } : '',
                ev.rival ? r0(ev.rival.precio) : null,
                ev.rival ? r0(it.precio - ev.rival.precio) : null,
                ev.rival ? (it.precio - ev.rival.precio) / ev.rival.precio : null,
                baja ? r0(ev.objetivo) : null, baja ? r0(ev.recibObjetivo) : null, baja ? r0(ev.margenObjetivo) : null,
                r0(ev.piso), it.comFuente, it.envioFuente
            ];
        });
        const wsSku = hoja(colSku, datosSku, filas.map(({ ev }) => colorFila(ev)));

        // ---- Hoja "Competidores" (un renglón por competidor) ----
        const colComp = [
            { titulo: 'SKU', ancho: 16 }, { titulo: 'Producto', ancho: 40 }, { titulo: 'Nuestro precio', ancho: 14, fmt: MONEDA },
            { titulo: 'Vendedor', ancho: 26 }, { titulo: 'Publicación', ancho: 17 }, { titulo: 'Su precio', ancho: 13, fmt: MONEDA },
            { titulo: 'Diferencia (nosotros − él)', ancho: 15, fmt: MONEDA }, { titulo: 'Diferencia %', ancho: 11, fmt: PCT },
            { titulo: '¿Más barato que nosotros?', ancho: 13 }, { titulo: 'Full', ancho: 6 },
            { titulo: 'Cómo se cargó', ancho: 14 }, { titulo: 'Fecha del precio', ancho: 13 }
        ];
        const datosComp = [], coloresComp = [];
        filas.forEach(({ it, ev }) => {
            ev.comps.slice().sort((a, b) => (a.precio || Infinity) - (b.precio || Infinity)).forEach(c => {
                const dif = c.precio > 0 ? it.precio - c.precio : null;
                datosComp.push([
                    it.sku, it.desc, r0(it.precio), c.vendedor || '', { v: c.mla || 'Link', link: c.url },
                    c.precio > 0 ? r0(c.precio) : null, r0(dif), dif != null ? dif / c.precio : null,
                    dif == null ? '' : dif > 0 ? 'Sí' : 'No', c.full ? 'Sí' : '',
                    c.auto ? 'Automático' : 'Manual', c.fecha || ''
                ]);
                coloresComp.push(dif != null && dif > 0 ? 'FCE4D6' : null);
            });
        });
        const wsComp = hoja(colComp, datosComp, coloresComp);

        // ---- Hoja "Resumen" ----
        const cuenta = (e) => filas.filter(f => f.ev.estado === e).length;
        const bajar = filas.filter(f => f.ev.estado === 'bajar');
        const resumen = [
            ['Competencia en Mercado Libre'],
            [`Generado el ${new Date().toLocaleString('es-AR')}`],
            [`Filtro aplicado: ${selVista.options[selVista.selectedIndex].text}${chkSoloStock.checked ? ' · solo con stock' : ''}${inBuscar.value.trim() ? ` · búsqueda "${inBuscar.value.trim()}"` : ''}`],
            [],
            ['Indicador', 'Cantidad'],
            ['SKU analizados', filas.length],
            ['SKU con competidores', filas.filter(f => f.ev.comps.length).length],
            ['Competidores en total', datosComp.length],
            ['Nos ganan y PODEMOS bajar', bajar.length],
            ['Nos ganan y NO conviene bajar', cuenta('no')],
            ['Ya somos los más baratos', cuenta('ganamos')],
            ['Hoy vendemos debajo del costo', filas.filter(f => f.ev.margen != null && f.ev.margen < 0).length],
            ['Sin competidores cargados', cuenta('sin')],
            [],
            ['Cómo se calcula'],
            ['Recibimos = precio × (1 − comisión ML) − envío por unidad. Comisión y envío reales de nuestras ventas; si no hubo ventas, aproximados por medida/rodado.'],
            ['"Podemos bajar": poniéndonos $1.000 debajo del competidor más barato, lo que recibimos sigue siendo mayor que el costo c/IVA.'],
            ['Colores: verde = podemos bajar · rojo = no conviene / debajo del costo · amarillo = ya somos los más baratos.']
        ];
        const wsRes = XLib.utils.aoa_to_sheet(resumen);
        wsRes['!cols'] = [{ wch: 36 }, { wch: 12 }];
        wsRes['!merges'] = [0, 1, 2, 14, 15, 16, 17].map(r => ({ s: { r, c: 0 }, e: { r, c: 5 } }));
        if (conEstilo) {
            const set = (r, c, s) => { const a = XLib.utils.encode_cell({ r, c }); if (wsRes[a]) wsRes[a].s = s; };
            set(0, 0, { font: { bold: true, sz: 16, color: { rgb: '1F3864' } } });
            set(1, 0, { font: { italic: true, color: { rgb: '595959' } } });
            set(2, 0, { font: { italic: true, color: { rgb: '595959' } } });
            [0, 1].forEach(c => set(4, c, { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1F3864' } } }));
            set(8, 0, { font: { bold: true }, fill: { fgColor: { rgb: COLOR.bajar } } }); set(8, 1, { font: { bold: true }, fill: { fgColor: { rgb: COLOR.bajar } } });
            set(9, 0, { fill: { fgColor: { rgb: COLOR.no } } }); set(9, 1, { fill: { fgColor: { rgb: COLOR.no } } });
            set(10, 0, { fill: { fgColor: { rgb: COLOR.ganamos } } }); set(10, 1, { fill: { fgColor: { rgb: COLOR.ganamos } } });
            set(14, 0, { font: { bold: true, color: { rgb: '1F3864' } } });
        }

        const wb = XLib.utils.book_new();
        XLib.utils.book_append_sheet(wb, wsRes, 'Resumen');
        XLib.utils.book_append_sheet(wb, wsSku, 'Por SKU');
        XLib.utils.book_append_sheet(wb, wsComp, 'Competidores');
        XLib.writeFile(wb, `Competencia_ML_${hoy()}.xlsx`);
    });
    // ---------- arranque ----------
    // Con archivos guardados se arma la tabla sola; si ML está conectado y pasó un día
    // desde la última actualización, los precios de la competencia se actualizan solos.
    restaurarDatos();
    if (pubs && costos) {
        construir();
        render();
        resultsPanel.classList.remove('hidden');
    }
    pintarCambios();
    consultarEstadoMl().then(() => {
        if (mlConectado && items.length && tocaActualizar()) buscarCompetidoresMl({ automatica: true });
    });
})();

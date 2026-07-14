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
            if (type === 'ml' || type === 'sys') {
                handleFileSelect(fileInput, statusElement, type);
            } else {
                handleFileSelectRentabilidad(fileInput, statusElement, type);
            }
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length === 0) return;
        if (type === 'ml' || type === 'sys') {
            handleFileSelect(fileInput, statusElement, type);
        } else {
            handleFileSelectRentabilidad(fileInput, statusElement, type);
        }
    });
}

setupDropZone(dropZoneMl, fileMl, statusMl, 'ml');
setupDropZone(dropZoneSys, fileSys, statusSys, 'sys');

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
// parsing logic
function parseMl(rows) {
    const map = {};
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        
        let sku = row[4];
        let stock = row[7];
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
        
        if (row.length < 8) continue;
        
        let sku = row[1];
        let desc = row[2];
        let stockSantiago = parseInt(row[5], 10) || 0;
        let stockSantiago1435 = parseInt(row[6], 10) || 0;
        let stockFull = parseInt(row[7], 10) || 0;
        let totalStock = stockSantiago + stockSantiago1435 + stockFull;
        
        if (sku !== undefined && sku !== null && sku !== '') {
            sku = sku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            if (sku === "SKU") continue;
            
            const descVal = desc ? desc.toString().trim() : 'Neumático Sistema';
            
            if (map.hasOwnProperty(sku)) {
                map[sku].stock += totalStock;
                map[sku].marzoStock += stockSantiago;
                map[sku].marzo1435Stock += stockSantiago1435;
                map[sku].fullStock += stockFull;
            } else {
                map[sku] = {
                    stock: totalStock,
                    descripcion: descVal,
                    marzoStock: stockSantiago,
                    marzo1435Stock: stockSantiago1435,
                    fullStock: stockFull
                };
            }
        }
    }
    return map;
}

btnProcess.addEventListener('click', () => {
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

btnDownload.addEventListener('click', () => {
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
            } else if (type === 'sys') {
                dataSysNuevos = parseSysNuevos(json);
            }
            
            checkReadyNuevos();
        } catch (error) {
            Swal.fire('Error', 'No se pudo leer el archivo Excel/CSV.', 'error');
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
            
        tr.innerHTML = `
            <td><strong>${item.SKU}</strong></td>
            <td><span style="font-size: 0.85rem; color: var(--text-secondary);">${item.Title || 'Sin detalle'}</span></td>
            <td>${mlDisplay}</td>
            <td>${item['Stock Sistema']}</td>
            <td><span class="status-badge ${item._badgeClass}">${item.Motivo}</span></td>
        `;
        
        resultsBodyNuevos.appendChild(tr);
    });
}

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

// --- NEW RENTABILIDAD LOGIC ---
const fileGuerrini = document.getElementById('file-guerrini');
const fileVentas = document.getElementById('file-ventas');
const statusGuerrini = document.getElementById('status-guerrini');
const statusVentas = document.getElementById('status-ventas');
const btnProcessRentabilidad = document.getElementById('btn-process-rentabilidad');
const resultsPanelRentabilidad = document.getElementById('results-panel-rentabilidad');
const resultsBodyRentabilidad = document.getElementById('results-body-rentabilidad');
const btnDownloadRentabilidad = document.getElementById('btn-download-rentabilidad');

const statGanancia = document.getElementById('stat-ganancia');
const statPerdida = document.getElementById('stat-perdida');

const dropZoneGuerrini = document.getElementById('drop-zone-guerrini');
const dropZoneVentas = document.getElementById('drop-zone-ventas');

let dataGuerrini = {};
let dataVentas = [];
let finalResultsRentabilidad = [];

setupDropZone(dropZoneGuerrini, fileGuerrini, statusGuerrini, 'guerrini');
setupDropZone(dropZoneVentas, fileVentas, statusVentas, 'ventas');

function handleFileSelectRentabilidad(input, statusElement, type) {
    if (input.files.length === 0) return;
    
    const file = input.files[0];
    statusElement.textContent = file.name;
    statusElement.classList.add('uploaded');
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            if (type === 'guerrini') {
                dataGuerrini = {}; // Reset
                // Parse all sheets in Guerrini list
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                    parseGuerrini(json);
                });
            } else if (type === 'ventas') {
                // Ventas ML report typically uses the first active sheet
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                dataVentas = parseVentas(json);
            }
            
            if (Object.keys(dataGuerrini).length > 0 && dataVentas.length > 0) {
                btnProcessRentabilidad.disabled = false;
            }
        } catch (error) {
            Swal.fire('Error', 'No se pudo leer el archivo Excel.', 'error');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseGuerrini(rows) {
    // Col A (0) = SKU, Col H (7) = Precio Base
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 8) continue;
        
        let sku = row[0];
        let price = row[7];
        
        if (sku !== undefined && sku !== null && sku !== '') {
            sku = sku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            // Handle possibility of string formatted price (e.g., "$ 1,000.00")
            if(typeof price === 'string') {
               price = parseFloat(price.replace(/[^0-9,-]+/g, '').replace(',', '.'));
            } else {
               price = parseFloat(price);
            }
            
            if (!isNaN(price) && price > 0) {
                // Apple discounts: 35%, 25%, 8% => price * 0.65 * 0.75 * 0.92 = price * 0.4485
                let finalCost = price * 0.4485;
                dataGuerrini[sku] = finalCost;
            }
        }
    }
}

function parseVentas(rows) {
    const list = [];
    // Start reading from row 6 (index 5)
    // Col V (21) = SKU, Col S (18) = Sold Price, Col G (6) = Quantity
    for (let i = 5; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 25) continue;
        
        let dateRaw = row[1];    // B
        let paramSku = row[21]; // V
        let soldPrice = row[18]; // S
        let quantity = row[6];   // G
        let title = row[24];     // Y
        
        if (paramSku !== undefined && paramSku !== null && paramSku !== '') {
            let sku = paramSku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            
            if(typeof soldPrice === 'string') {
               soldPrice = parseFloat(soldPrice.replace(/[^0-9,-]+/g, '').replace(',', '.'));
            } else {
               soldPrice = parseFloat(soldPrice);
            }
            
            if(typeof quantity === 'string') {
               quantity = parseFloat(quantity.replace(/[^0-9,-]+/g, '').replace(',', '.'));
            } else {
               quantity = parseFloat(quantity);
            }
            
            // Prevent division by zero
            if (isNaN(quantity) || quantity <= 0) {
               quantity = 1;
            }
            
            if (!isNaN(soldPrice)) {
                // Divide the total sold price by the quantity to get the unit sold price
                let unitSoldPrice = soldPrice / quantity;
                
                // Ignorar ventas con montos muy bajos, suelen ser devoluciones
                if (unitSoldPrice > 1000) {
                    list.push({ 
                        rowIndex: i, 
                        dateRaw: dateRaw, 
                        sku: sku, 
                        title: title || 'Sin detalle', 
                        soldPrice: unitSoldPrice, 
                        quantity: quantity 
                    });
                }
            }
        }
    }
    return list;
}

btnProcessRentabilidad.addEventListener('click', () => {
    btnProcessRentabilidad.disabled = true;
    btnProcessRentabilidad.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Procesando...';
    
    setTimeout(() => {
        analizarRentabilidad();
        renderTableRentabilidad();
        
        resultsPanelRentabilidad.classList.remove('hidden');
        
        btnProcessRentabilidad.innerHTML = '<i class="fa-solid fa-bolt"></i> Analizar Rentabilidad';
        btnProcessRentabilidad.disabled = false;
        
        Swal.fire({
            title: '¡Análisis de Rentabilidad Completado!',
            text: `Se analizaron ${finalResultsRentabilidad.length} ventas.`,
            icon: 'success',
            confirmButtonColor: '#3b82f6'
        });
    }, 500); 
});

function formatter(val) {
    return new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS' }).format(val);
}

function analizarRentabilidad() {
    finalResultsRentabilidad = [];
    let gananciaCount = 0;
    let perdidaCount = 0;

    dataVentas.forEach(venta => {
        // Find if base SKU is present (or support Kit variations just in case, but usually exact SKU match is fine for sales)
        let cost = dataGuerrini[venta.sku];
        
        // Let's also verify kit parsing like we did for stock, just in case they sell kits.
        const kitMatchPrefix = venta.sku.match(/^KITX(\d+)-(.*)/);
        const kitMatchSuffix = venta.sku.match(/^(.*)X(\d+)$/);
        
        if (cost === undefined) {
            if (kitMatchPrefix) {
                let multiplier = parseInt(kitMatchPrefix[1], 10);
                let realSku = kitMatchPrefix[2];
                if(dataGuerrini[realSku] !== undefined) {
                   cost = dataGuerrini[realSku] * multiplier;
                }
            } else if (kitMatchSuffix) {
                let multiplier = parseInt(kitMatchSuffix[2], 10);
                let realSku = kitMatchSuffix[1];
                if(dataGuerrini[realSku] !== undefined) {
                   cost = dataGuerrini[realSku] * multiplier;
                }
            }
        }

        if (cost !== undefined) {
            const unitProfit = venta.soldPrice - cost;
            const totalDiferencia = unitProfit * venta.quantity;
            
            const estado = totalDiferencia > 0 ? 'Ganancia' : 'Pérdida';
            const badgeClass = totalDiferencia > 0 ? 'bg-success' : 'bg-danger';
            
            if (totalDiferencia > 0) gananciaCount++;
            else perdidaCount++;

            finalResultsRentabilidad.push({
                OriginalIndex: venta.rowIndex,
                Fecha: venta.dateRaw,
                SKU: venta.sku,
                Title: venta.title,
                CostoFinal: cost,
                PrecioVendido: venta.soldPrice,
                CantidadVendida: venta.quantity,
                Resultado: totalDiferencia,
                Estado: estado,
                _badgeClass: badgeClass
            });
        }
    });

    statGanancia.textContent = gananciaCount;
    statPerdida.textContent = perdidaCount;
}

function renderTableRentabilidad() {
    resultsBodyRentabilidad.innerHTML = '';
    
    // El reporte de ML originalmente suele venir ordenado por fecha de forma descendente.
    // Usamos el OriginalIndex para garantizar un orden cronológico exacto al del Excel madre,
    // garantizando que las fechas queden correlativas en orden.
    finalResultsRentabilidad.sort((a, b) => a.OriginalIndex - b.OriginalIndex);
    
    finalResultsRentabilidad.forEach(item => {
        const tr = document.createElement('tr');
        
        let displayDate = item.Fecha ? item.Fecha.toString() : '';
        // Si Excel envió la fecha como código serial (muy frecuente)
        if (typeof item.Fecha === 'number') {
            displayDate = new Date((item.Fecha - (25567 + 2)) * 86400 * 1000).toLocaleDateString('es-AR');
        } else if (displayDate.length > 10 && displayDate.includes('T')) {
            displayDate = displayDate.split('T')[0]; // Para fechas ISO reducidas
        } else if (displayDate.length > 15) {
            displayDate = displayDate.substring(0, 10); // Reducirlo para no ocupar mucho
        }
        
        tr.innerHTML = `
            <td><span style="font-size: 0.8rem; color: var(--text-secondary); white-space: nowrap;">${displayDate}</span></td>
            <td><strong>${item.SKU}</strong></td>
            <td><span style="font-size: 0.85rem; color: var(--text-secondary);">${item.Title}</span></td>
            <td>${formatter(item.CostoFinal)}</td>
            <td>${formatter(item.PrecioVendido)}</td>
            <td><strong>x${item.CantidadVendida}</strong></td>
            <td class="${item.Resultado > 0 ? 'text-success' : 'text-danger'}"><strong>${item.Resultado > 0 ? '+' : ''}${formatter(item.Resultado)}</strong></td>
            <td><span class="status-badge ${item._badgeClass}">${item.Estado}</span></td>
        `;
        
        resultsBodyRentabilidad.appendChild(tr);
    });
}

btnDownloadRentabilidad.addEventListener('click', () => {
    const excelData = finalResultsRentabilidad.map(item => {
        return {
            'Fecha Venta': item.Fecha,
            'SKU': item.SKU,
            'Detalle del Producto': item.Title,
            'Costo Guerrini (Por Artículo Publicado) ($)': item.CostoFinal,
            'Precio de Venta ML (Por Artículo Publicado) ($)': item.PrecioVendido,
            'Unidades de este SKU Vendidas': item.CantidadVendida,
            'Ganancia/Pérdida Total del Renglón ($)': item.Resultado,
            'Estado': item.Estado
        };
    });
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    
    const wscols = [
        {wch: 15}, // Fecha
        {wch: 25}, // SKU
        {wch: 45}, // Detalle
        {wch: 20}, // Costo
        {wch: 20}, // Venta
        {wch: 20}, // Cantidad
        {wch: 25}, // Resultado
        {wch: 15}  // Estado
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rentabilidad");
    
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Analisis_Rentabilidad_${dateStr}.xlsx`);
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
    
    const headers = rows[headerRowIdx].map(h => String(h).trim().toUpperCase());
    
    const idxSku = headers.indexOf('SKU');
    const idxTitle = headers.findIndex(h => h.includes('PRODUCTO') || h.includes('TÍTULO') || h.includes('TITLE'));
    const idxPending = headers.findIndex(h => h.includes('CAMINO') || h.includes('PENDIENTES'));
    const idxFull = headers.findIndex(h => h.includes('EN FULL') || h.includes('APTAS'));
    const idxSales = headers.findIndex(h => h.includes('VENTAS ÚLTIMOS 30 DÍAS (U.)') || h.includes('VENTAS ÚLTIMOS 30 DÍAS') || h.includes('VENTAS (U.)'));
    
    if (idxSku === -1 || idxFull === -1) {
        throw new Error("El archivo de Full no tiene las columnas requeridas (SKU y Unidades en Full).");
    }
    
    for (let i = headerRowIdx + 1; i < rows.length; i++) {
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
    marzo: document.getElementById('stat-restock-marzo'),
    marzo1435: document.getElementById('stat-restock-marzo1435')
};

let dataMlFullStock = null;
let dataSysRestock = null;
let finalResultsRestock = [];

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
    
    allSkus.forEach(sku => {
        const mlItem = dataMlFullStock[sku] || { aptas: 0, pendientes: 0, ventas30: 0, title: '' };
        const mlAptas = mlItem.aptas;
        const mlPendientes = mlItem.pendientes;
        const mlVentas = mlItem.ventas30;
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
        
        let sysStockMarzo = 0;
        let sysStockMarzo1435 = 0;
        let sysDesc = '';
        
        if (dataSysRestock.hasOwnProperty(skuForSystem)) {
            const sysItem = dataSysRestock[skuForSystem];
            sysStockMarzo = sysItem.marzoStock || 0;
            sysStockMarzo1435 = sysItem.marzo1435Stock || 0;
            sysDesc = sysItem.descripcion || '';
        }
        
        // Ajustar stock local para kits
        const availMarzoKits = Math.floor(sysStockMarzo / multiplier);
        const availMarzo1435Kits = Math.floor(sysStockMarzo1435 / multiplier);
        const availTotalKits = availMarzoKits + availMarzo1435Kits;
        
        const toSendKits = Math.min(deficit, availTotalKits);
        
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
            
            if (availMarzoKits >= toSendKits) {
                sendMarzo = toSendKits;
                sendMarzo1435 = 0;
            } else if (availMarzo1435Kits >= toSendKits) {
                sendMarzo = 0;
                sendMarzo1435 = toSendKits;
            } else {
                // Distribute: Marzo first, then Marzo 1435
                sendMarzo = availMarzoKits;
                sendMarzo1435 = Math.min(availMarzo1435Kits, toSendKits - sendMarzo);
            }
            
            const unitName = multiplier > 1 ? 'kit(s)' : 'u.';
            const parts = [];
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
            LocalMarzo: sysStockMarzo,
            LocalMarzo1435: sysStockMarzo1435,
            Multiplier: multiplier,
            Deficit: deficit,
            
            // Guardar valores recomendados originales como referencia
            RecMarzo: sendMarzo,
            RecMarzo1435: sendMarzo1435,
            RecSugerencia: sugerencia,
            RecBadgeClass: badgeClass,
            
            // Selección del usuario (comienza vacía/null)
            ToSend: null,
            SendMarzo: 0,
            SendMarzo1435: 0,
            Sugerencia: sugerencia,
            _badgeClass: badgeClass
        });
        countModels++;
    });
    
    recalcularEstadisticasGlobalesRestock();
}
 
function renderTableRestock() {
    resultsBodyRestock.innerHTML = '';
    
    // Ordenar de mayor a menor según el stock local total
    finalResultsRestock.sort((a, b) => (b.LocalMarzo + b.LocalMarzo1435) - (a.LocalMarzo + a.LocalMarzo1435));
    
    finalResultsRestock.forEach(item => {
        const tr = document.createElement('tr');
        
        const localDisplay = `Marzo: ${item.LocalMarzo} | 1435: ${item.LocalMarzo1435}`;
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
    const availMarzoKits = Math.floor(item.LocalMarzo / multiplier);
    const availMarzo1435Kits = Math.floor(item.LocalMarzo1435 / multiplier);
    const totalAvailKits = availMarzoKits + availMarzo1435Kits;
    
    let sendMarzo = 0;
    let sendMarzo1435 = 0;
    let sugerencia = '';
    let badgeClass = '';
    
    if (newVal === null) {
        // Si está vacío, restaurar recomendación original
        sendMarzo = item.RecMarzo;
        sendMarzo1435 = item.RecMarzo1435;
        sugerencia = item.RecSugerencia;
        badgeClass = item.RecBadgeClass;
    } else if (newVal === 0) {
        sugerencia = "No enviar";
        badgeClass = "bg-info";
    } else {
        badgeClass = "bg-success";
        
        if (availMarzoKits >= newVal) {
            sendMarzo = newVal;
            sendMarzo1435 = 0;
        } else if (availMarzo1435Kits >= newVal) {
            sendMarzo = 0;
            sendMarzo1435 = newVal;
        } else {
            // Distribute: Marzo first, then Marzo 1435
            sendMarzo = availMarzoKits;
            sendMarzo1435 = Math.min(availMarzo1435Kits, newVal - sendMarzo);
        }
        
        const unitName = multiplier > 1 ? 'kit(s)' : 'u.';
        const parts = [];
        if (sendMarzo > 0) parts.push(`${sendMarzo} de Santiago Marzo`);
        if (sendMarzo1435 > 0) parts.push(`${sendMarzo1435} de Santiago Marzo 1435`);
        
        if (parts.length > 0) {
            sugerencia = `Enviar ${newVal} ${unitName} (Mandar ${parts.join(' y ')})`;
        } else {
            sugerencia = `Enviar ${newVal} ${unitName} (Falta stock local)`;
        }
        
        const totalSentKits = sendMarzo + sendMarzo1435;
        if (totalSentKits < newVal) {
            sugerencia += ` - Faltan ${newVal - totalSentKits} ${unitName} localmente`;
            badgeClass = "bg-warning";
        }
    }
    
    if (newVal === null) {
        item.SendMarzo = 0;
        item.SendMarzo1435 = 0;
        item.Sugerencia = item.RecSugerencia;
        item._badgeClass = item.RecBadgeClass;
    } else {
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
    let totalMarzo = 0;
    let totalMarzo1435 = 0;
    let countModels = 0;
    
    finalResultsRestock.forEach(item => {
        if (item.ToSend > 0) {
            totalSend += item.ToSend * item.Multiplier;
            totalMarzo += item.SendMarzo * item.Multiplier;
            totalMarzo1435 += item.SendMarzo1435 * item.Multiplier;
            countModels++;
        }
    });
    
    statsRestock.models.textContent = countModels;
    statsRestock.total.textContent = totalSend;
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
            'Ventas (30 días)': item.Ventas,
            'Stock en Full (Apto)': item.Aptas,
            'En camino (Pendiente)': item.Pendientes,
            'Stock Santiago Marzo': item.LocalMarzo,
            'Stock Santiago Marzo 1435': item.LocalMarzo1435,
            'Cantidad a Enviar': item.ToSend,
            'Distribución de Origen': item.Sugerencia
        };
    });
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    
    const wscols = [
        {wch: 15}, // SKU
        {wch: 45}, // Detalle
        {wch: 15}, // Ventas
        {wch: 20}, // Apto
        {wch: 20}, // Camino
        {wch: 20}, // Marzo
        {wch: 25}, // Marzo 1435
        {wch: 20}, // Cantidad a Enviar
        {wch: 55}  // Sugerencia
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Plan de Envío Full");
    
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Plan_Envio_Full_${dateStr}.xlsx`);
});

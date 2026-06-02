// --- FUNCIÓN PARA FITOSANITARIOS ---
async function bajarExcelFitoOficial(selector) {
    if (typeof ExcelJS === 'undefined') { alert("Error: Librería no cargada."); return; }
    try {
        const año = selector ? selector.value : '2023-2024';
        const datos = JSON.parse(localStorage.getItem('fito_' + año)) || [];
        if (datos.length === 0) { alert("No hay datos."); return; }
        
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Tratamientos');
        const cabeceras = ['Fecha', 'Parada', 'Producto', 'Dosis', 'Registro', 'Plazo', 'Operario', 'Carnet', 'Equipo', 'ROMA'];
        const headerRow = ws.addRow(cabeceras);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D5A27' } };
        ws.columns = cabeceras.map(() => ({ width: 15 }));
        datos.forEach(r => {
            const row = ws.addRow([r.fecha, r.parada, r.producto, r.dosis, r.registro, r.plazo, r.operario, r.carnet, r.equipo, r.roma]);
            row.eachCell((cell) => { cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; });
        });
        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, `Fitosanitarios_${año}.xlsx`);
    } catch (e) { alert("Error: " + e.message); }
}

// --- FUNCIÓN PARA REGISTRO (PARADAS) ---
async function bajarExcelRegistroOficial(selector) {
    if (typeof ExcelJS === 'undefined') { alert("Error: Librería no cargada."); return; }
    try {
        const año = selector ? selector.value : '2023-2024';
        const datos = JSON.parse(localStorage.getItem('datos_' + año)) || [];
        if (datos.length === 0) { alert("No hay datos."); return; }
        
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Registros');
        const cabeceras = ['Fecha', 'Parada', 'Kilos'];
        const headerRow = ws.addRow(cabeceras);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D5A27' } };
        ws.columns = cabeceras.map(() => ({ width: 15 }));
        datos.forEach(r => {
            const row = ws.addRow([r.fecha, r.parada, r.kilos]);
            row.eachCell((cell) => { cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; });
        });
        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, `Registros_${año}.xlsx`);
    } catch (e) { alert("Error: " + e.message); }
}
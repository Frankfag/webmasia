async function descargarExcelTotalesPremium(selectorAño, selectorFiltro, modoVista) {
    let registros = JSON.parse(localStorage.getItem('cosecha_oliva')) || [];
    let añoActivo = selectorAño ? selectorAño.value : '2025-2026';
    let pActiva = selectorFiltro ? selectorFiltro.value : 'TODAS';
    
    let filtrados = registros.filter(function(r) { 
        return r && r.fecha && (pActiva === 'TODAS' || r.parada === pActiva) && r.campaña === añoActivo; 
    });

    if (filtrados.length === 0) return alert("No hay datos en este filtro para exportar.");

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Resumen_Balance');
    worksheet.views = [{ showGridLines: true }];

    worksheet.getRow(1).height = 18;
    worksheet.getRow(2).height = 28;

    worksheet.mergeCells('B2:D2');
    let titulo = worksheet.getCell('B2');
    titulo.value = "MASÍA CENTENARIA";
    titulo.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF2C3E2F' } };

    worksheet.mergeCells('B3:D3');
    let subtitulo = worksheet.getCell('B3');
    subtitulo.value = 'Informe Analítico de Cosecha — Campaña ' + añoActivo + ' (Filtro: ' + pActiva + ')';
    subtitulo.font = { name: 'Segoe UI', size: 10.5, italic: true, color: { argb: 'FF555555' } };
    worksheet.getRow(3).height = 20;

    const filaEncabezado = worksheet.getRow(5);
    filaEncabezado.values = ['Fecha Balance', 'Parada / Concepto', 'Kilos Netos'];
    filaEncabezado.height = 26;
    filaEncabezado.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E2F' } };
        cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = { bottom: { style: 'medium', color: { argb: 'FFB58D3D' } } };
    });

    let totalKilos = 0;
    let filtradosOrdenados = [...filtrados].sort((a, b) => new Date(b.fecha.split('/').reverse().join('-')) - new Date(a.fecha.split('/').reverse().join('-')));

    if (modoVista === 'detallada') {
        filtradosOrdenados.forEach((r, idx) => {
            let k = parseFloat(r.kilos) || 0;
            totalKilos += k;
            let row = worksheet.addRow([r.fecha, r.parada + ' (Viaje #' + (filtradosOrdenados.length - idx) + ')', k]);
            row.height = 20;
            row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
            row.getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };
            row.getCell(3).alignment = { vertical: 'middle', horizontal: 'right' };
            row.getCell(3).numFmt = '#,##0" kg"';
            row.eachCell(c => {
                c.font = { name: 'Segoe UI', size: 10 };
                c.border = { bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } } };
            });
        });
    } else {
        let agrupados = {};
        filtradosOrdenados.forEach(function(reg) {
            let k = parseFloat(reg.kilos) || 0;
            totalKilos += k;
            let clave = reg.fecha + "-" + reg.parada;
            if (!agrupados[clave]) {
                agrupados[clave] = { fecha: reg.fecha, parada: reg.parada, total: k, veces: 1 };
            } else {
                agrupados[clave].total += k;
                agrupados[clave].veces += 1;
            }
        });

        Object.values(agrupados).forEach(function(inf) {
            let row = worksheet.addRow([inf.fecha, inf.parada + ' (' + inf.veces + ' pesajes)', inf.total]);
            row.height = 20;
            row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
            row.getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };
            row.getCell(3).alignment = { vertical: 'middle', horizontal: 'right' };
            row.getCell(3).numFmt = '#,##0" kg"';
            row.eachCell(c => {
                c.font = { name: 'Segoe UI', size: 10 };
                c.border = { bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } } };
            });
        });
    }

    worksheet.addRow([]);
    let filaTotal = worksheet.getRow(worksheet.lastRow.number + 1);
    filaTotal.values = ['TOTAL KILOS BALANCE GENERAL:', '', totalKilos];
    filaTotal.height = 26;
    worksheet.mergeCells(`A${filaTotal.number}:B${filaTotal.number}`);
    
    [1, 3].forEach(colIdx => {
        let c = filaTotal.getCell(colIdx);
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E2F' } };
        c.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        c.border = { top: { style: 'medium', color: { argb: 'FFB58D3D' } } };
    });
    filaTotal.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
    filaTotal.getCell(3).alignment = { vertical: 'middle', horizontal: 'right' };
    filaTotal.getCell(3).numFmt = '#,##0" kg"';

    worksheet.columns.forEach(column => {
        let maxLen = 0;
        column.eachCell({ includeEmpty: true }, cell => {
            let valLen = cell.value ? cell.value.toString().length : 0;
            if (valLen > maxLen) maxLen = valLen;
        });
        column.width = maxLen < 16 ? 16 : maxLen + 3;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, 'CUE_Informe_General_' + añoActivo + '.xlsx');
}
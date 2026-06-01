async function descargarExcelParadasPremium(selectorAño) {
    let registros = JSON.parse(localStorage.getItem('cosecha_oliva')) || [];
    let añoActivo = selectorAño ? selectorAño.value : '2025-2026';
    let filtrados = registros.filter(function(r) { return r && r.campaña === añoActivo; });

    if (filtrados.length === 0) return alert("No hay datos en esta campaña para exportar.");

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Pesajes_Diarios');
    worksheet.views = [{ showGridLines: true }];

    worksheet.getRow(1).height = 18;
    worksheet.getRow(2).height = 28;

    worksheet.mergeCells('B2:D2');
    let titulo = worksheet.getCell('B2');
    titulo.value = "MASÍA CENTENARIA";
    titulo.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF2C3E2F' } };
    titulo.alignment = { vertical: 'middle' };

    worksheet.mergeCells('B3:D3');
    let subtitulo = worksheet.getCell('B3');
    subtitulo.value = 'Listado Oficial de Paradas — Campaña ' + añoActivo;
    subtitulo.font = { name: 'Segoe UI', size: 10.5, italic: true, color: { argb: 'FF555555' } };
    worksheet.getRow(3).height = 20;

    const filaEncabezado = worksheet.getRow(5);
    filaEncabezado.values = ['Fecha Recogida', 'Parada / Parcela', 'Kilos Pesados'];
    filaEncabezado.height = 26;

    filaEncabezado.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E2F' } };
        cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = { bottom: { style: 'medium', color: { argb: 'FFB58D3D' } } };
    });

    let totalKilos = 0;
    let filtradosOrdenados = [...filtrados].sort((a, b) => new Date(b.fecha.split('/').reverse().join('-')) - new Date(a.fecha.split('/').reverse().join('-')));

    filtradosOrdenados.forEach(r => {
        let kilosNum = parseFloat(r.kilos) || 0;
        totalKilos += kilosNum;

        let nuevaFila = worksheet.addRow([r.fecha, r.parada, kilosNum]);
        nuevaFila.height = 20;

        nuevaFila.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
        nuevaFila.getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };
        nuevaFila.getCell(3).alignment = { vertical: 'middle', horizontal: 'right' };
        nuevaFila.getCell(3).numFmt = '#,##0" kg"';

        nuevaFila.eachCell((cell) => {
            cell.font = { name: 'Segoe UI', size: 10 };
            cell.border = { bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } } };
        });
    });

    worksheet.addRow([]); 
    let filaTotal = worksheet.getRow(worksheet.lastRow.number + 1);
    filaTotal.values = ['TOTAL KILOS COSECHADOS:', '', totalKilos];
    filaTotal.height = 26;
    worksheet.mergeCells(`A${filaTotal.number}:B${filaTotal.number}`);

    [1, 3].forEach(colIdx => {
        let cell = filaTotal.getCell(colIdx);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E2F' } };
        cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.border = { top: { style: 'medium', color: { argb: 'FFB58D3D' } } };
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
    saveAs(blob, 'CUE_Pesajes_Paradas_' + añoActivo + '.xlsx');
}
document.addEventListener('DOMContentLoaded', function() {
    var selectorAño = document.getElementById('selectorAño');
    var formularioFito = document.getElementById('form-fito');
    var cuerpoTabla = document.querySelector('.tabla-datos tbody');

    if (selectorAño) {
        selectorAño.innerHTML = "";
        for (var i = 2024; i <= 2050; i++) {
            var opt = document.createElement("option");
            opt.value = i + "-" + (i + 1); opt.textContent = i + "-" + (i + 1);
            selectorAño.appendChild(opt);
        }
        selectorAño.value = localStorage.getItem('campaña_activa') || '2025-2026';
        if(document.getElementById('añoDinamico')) document.getElementById('añoDinamico').textContent = "REGISTRO DE TRATAMIENTOS DE CAMPO — " + selectorAño.value;

        selectorAño.addEventListener('change', function(e) {
            localStorage.setItem('campaña_activa', e.target.value);
            if(document.getElementById('añoDinamico')) document.getElementById('añoDinamico').textContent = "REGISTRO DE TRATAMIENTOS DE CAMPO — " + e.target.value;
            cargarTratamientos();
        });
    }

    function cargarTratamientos() {
        if (!cuerpoTabla || !selectorAño) return; cuerpoTabla.innerHTML = '';
        var listado = JSON.parse(localStorage.getItem('fito_tratamientos')) || [];
        var listadoOrdenado = [...listado].reverse();

        listadoOrdenado.forEach(function(t) {
            if (t.campaña === selectorAño.value) {
                var tr = document.createElement('tr');
                tr.innerHTML = '<td>' + t.fecha + '</td>' +
                               '<td><b>' + t.parada + '</b><br><small style="color:#7f8c8d;">Sup: ' + (t.superficie || '---') + ' Ha</small></td>' +
                               '<td><b>' + t.producto + '</b><br><small style="color:#b58d3d;">' + t.plaga + '</small></td>' +
                               '<td><small>Reg: ' + t.registro + '<br>Dosis: ' + t.dosis + '<br>Caldo: ' + t.litros + 'L</small></td>' +
                               '<td><span style="color:#c0392b; font-weight:bold;">' + (t.plazo || '0') + ' días</span><br><small style="font-size:11px;color:#7f8c8d;">ROMA: ' + (t.roma || '---') + '</small></td>' +
                               '<td><button class="btn-del" data-id="' + t.id + '" style="background:#e74c3c;color:white;border:none;padding:5px 10px;border-radius:4px;cursor:pointer;font-weight:bold;">🗑</button></td>';
                
                tr.querySelector('.btn-del').addEventListener('click', function() {
                    if(confirm("¿Deseas eliminar este registro fitosanitario oficial?")) {
                        var todos = JSON.parse(localStorage.getItem('fito_tratamientos')) || [];
                        var filtrados = todos.filter(item => item.id !== t.id);
                        localStorage.setItem('fito_tratamientos', JSON.stringify(filtrados));
                        cargarTratamientos();
                    }
                });
                cuerpoTabla.appendChild(tr);
            }
        });
    }

    if (formularioFito) {
        formularioFito.addEventListener('submit', function(e) {
            e.preventDefault();
            if (!selectorAño) return false;
            var f = document.getElementById('fito-fecha').value; if (!f) return false;
            var partes = f.split('-');
            var fEuro = partes[2] + '/' + partes[1] + '/' + partes[0];
            
            var nuevoTratamiento = { 
                id: Date.now(), 
                fecha: fEuro, 
                parada: document.getElementById('fito-parada').value, 
                superficie: document.getElementById('fito-superficie').value,
                producto: document.getElementById('fito-producto').value, 
                registro: document.getElementById('fito-registro').value, 
                dosis: document.getElementById('fito-dosis').value, 
                litros: document.getElementById('fito-litros').value, 
                plaga: document.getElementById('fito-plaga').value, 
                operario: document.getElementById('fito-operario').value, 
                carnet: document.getElementById('fito-carnet').value,
                maquinaria: document.getElementById('fito-maquinaria').value, 
                roma: document.getElementById('fito-roma').value,
                plazo: document.getElementById('fito-plazo').value,
                campaña: selectorAño.value 
            };
            
            var listado = JSON.parse(localStorage.getItem('fito_tratamientos')) || [];
            listado.push(nuevoTratamiento); localStorage.setItem('fito_tratamientos', JSON.stringify(listado));
            alert("¡Tratamiento Oficial guardado en el Cuaderno!"); formularioFito.reset(); cargarTratamientos();
            return false;
        });
    }

    window.bajarExcelFitoOficial = async function(selectorAño) {
        let listado = JSON.parse(localStorage.getItem('fito_tratamientos')) || [];
        let añoActivo = selectorAño ? selectorAño.value : '2025-2026';
        let filtrados = listado.filter(r => r && r.campaña === añoActivo);

        if (filtrados.length === 0) return alert("No hay tratamientos para exportar.");

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('CUE_Fitosanitarios');
        worksheet.views = [{ showGridLines: true }];

        worksheet.mergeCells('A2:I2');
        let title = worksheet.getCell('A2');
        title.value = "MASÍA CENTENARIA - REGISTRO OFICIAL DE FITOSANITARIOS (CUE)";
        title.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FF2C3E2F' } };
        worksheet.getRow(2).height = 25;

        const headers = ['Fecha', 'Parada / Parcela', 'Superficie (Ha)', 'Producto Comercial', 'Nº Registro', 'Dosis', 'Vol. Caldo', 'Plaga/Motivo', 'Plazo Seguridad'];
        const filaHeader = worksheet.getRow(4);
        filaHeader.values = headers;
        filaHeader.height = 24;
        filaHeader.eachCell(c => {
            c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E2F' } };
            c.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
            c.alignment = { vertical: 'middle', horizontal: 'center' };
        });

        [...filtrados].reverse().forEach(r => {
            let row = worksheet.addRow([r.fecha, r.parada, parseFloat(r.superficie) || 0, r.producto, r.registro, r.dosis, (parseFloat(r.litros) || 0) + " L", r.plaga, (r.plazo || 0) + " días"]);
            row.height = 20;
            row.eachCell(c => { c.font = { name: 'Segoe UI', size: 9.5 }; c.border = { bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } } }; });
        });

        worksheet.columns.forEach(col => { col.width = 16; });

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), 'CUE_Fitosanitarios_Oficial_' + añoActivo + '.xlsx');
    };

    cargarTratamientos();
});
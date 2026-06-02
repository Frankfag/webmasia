document.addEventListener('DOMContentLoaded', () => {
    // 1. Detectar el nombre de la página actual (ej: "fitosanitario.html")
    const urlActual = window.location.pathname.split("/").pop();
    
    // 2. Buscar todos los enlaces del menú
    const enlaces = document.querySelectorAll('.menu-enlaces a');
    
    enlaces.forEach(enlace => {
        // 3. Si el href del enlace coincide con la página actual, le ponemos la clase
        if (enlace.getAttribute('href') === urlActual) {
            enlace.classList.add('activo');
        }
    });
});







document.addEventListener('DOMContentLoaded', function() {
    var selectorAño = document.getElementById('selectorAño');
    var formularioCosecha = document.querySelector('.formulario-cosecha');
    var cuerpoTabla = document.querySelector('.tabla-datos tbody');
    var selectorFiltro = document.getElementById('filtro-parada');
    var tarjetaSumaKilos = document.getElementById('total-kilos-suma');
    var modoVista = "agrupada"; 

    if (selectorAño) {
        selectorAño.innerHTML = "";
        for (var i = 2024; i <= 2050; i++) {
            var opt = document.createElement("option");
            opt.value = i + "-" + (i + 1); opt.textContent = i + "-" + (i + 1);
            selectorAño.appendChild(opt);
        }
        selectorAño.value = localStorage.getItem('campaña_activa') || '2025-2026';
        actualizarTextoCabecera();

        selectorAño.addEventListener('change', function(e) {
            localStorage.setItem('campaña_activa', e.target.value);
            actualizarTextoCabecera();
            arrancarPrograma();
        });
    }

    function actualizarTextoCabecera() {
        if(document.getElementById('añoDinamico') && selectorAño) {
            document.getElementById('añoDinamico').textContent = (tarjetaSumaKilos ? "INFORME GENERAL DE KILOS — " : "REGISTRO DE COSECHA DIARIA — ") + selectorAño.value;
        }
    }

    function arrancarPrograma() {
        if (!cuerpoTabla) return;
        procesarInformes();
    }

    // Configuración de los botones de cambio de vista en el informe global
    var btnAgrupada = document.getElementById('btn-vista-agrupada');
    var btnDetallada = document.getElementById('btn-vista-detallada');
    var btnFiltrar = document.getElementById('btn-filtrar');
    var btnReiniciar = document.getElementById('btn-reiniciar');

    if (btnAgrupada && btnDetallada) {
        btnAgrupada.addEventListener('click', function() { modoVista = "agrupada"; procesarInformes(); });
        btnDetallada.addEventListener('click', function() { modoVista = "detallada"; procesarInformes(); });
    }
    if (btnFiltrar) { btnFiltrar.addEventListener('click', procesarInformes); }
    if (btnReiniciar) {
        btnReiniciar.addEventListener('click', function() {
            if(confirm("¿Seguro que quieres borrar la memoria local de pesajes?")) {
                localStorage.removeItem('cosecha_oliva');
                procesarInformes();
            }
        });
    }

    if (formularioCosecha && !tarjetaSumaKilos && selectorAño) {
        formularioCosecha.addEventListener('submit', function(e) {
            e.preventDefault();
            var f = document.getElementById('fecha').value; if (!f) return;
            var partes = f.split('-');
            
            // CORREGIDO: Ajuste para evitar comas y montar formato europeo real
            var fEuro = partes[2] + '/' + partes[1] + '/' + partes[0];
            
            var nuevo = { 
                id: Date.now(), 
                fecha: fEuro, 
                parada: document.getElementById('parada').value, 
                kilos: document.getElementById('kilos').value, 
                campaña: selectorAño.value 
            };
            
            var registros = JSON.parse(localStorage.getItem('cosecha_oliva')) || [];
            registros.push(nuevo); 
            localStorage.setItem('cosecha_oliva', JSON.stringify(registros));
            
            alert("¡Pesaje guardado!"); 
            formularioCosecha.reset(); 
            arrancarPrograma();
        });
    }

    function procesarInformes() {
        if (!cuerpoTabla || !selectorAño) return; cuerpoTabla.innerHTML = '';
        var registros = JSON.parse(localStorage.getItem('cosecha_oliva')) || [];
        var pActiva = selectorFiltro ? selectorFiltro.value : 'TODAS';
        
        var filtrados = registros.filter(function(r) { 
            return r && r.fecha && (pActiva === 'TODAS' || r.parada === pActiva) && r.campaña === selectorAño.value; 
        });
        
        var totalKilos = 0, maxKilos = 0;
        var filtradosOrdenados = [...filtrados].reverse();

        if (modoVista === "detallada" || !tarjetaSumaKilos) {
            // Vista Detallada individual por viaje
            filtradosOrdenados.forEach(function(reg) {
                var k = parseFloat(reg.kilos) || 0; totalKilos += k; if (k > maxKilos) maxKilos = k;
                var fila = document.createElement('tr');
                var numViaje = filtrados.indexOf(reg) + 1;
                
                fila.innerHTML = '<td>' + reg.fecha + '</td><td>' + reg.parada + ' <small style="color:#b58d3d;font-weight:bold;">(Viaje #' + numViaje + ')</small></td><td>' + k.toLocaleString('es-ES') + ' kg</td><td><button class="btn-del" style="background:#e74c3c;color:white;border:none;padding:5px 10px;border-radius:4px;cursor:pointer;font-weight:bold;">🗑</button></td>';
                
                fila.querySelector('.btn-del').addEventListener('click', function() {
                    if (confirm("¿Eliminar viaje seleccionado?")) {
                        var todos = JSON.parse(localStorage.getItem('cosecha_oliva')) || [];
                        var limpios = todos.filter(function(r) { return r.id !== reg.id; });
                        localStorage.setItem('cosecha_oliva', JSON.stringify(limpios)); 
                        procesarInformes();
                    }
                });
                cuerpoTabla.appendChild(fila);
            });
        } else {
            // Vista Agrupada Consolidada
            var agrupados = {};
            filtradosOrdenados.forEach(function(reg) {
                var k = parseFloat(reg.kilos) || 0; totalKilos += k; if (k > maxKilos) maxKilos = k;
                var clave = reg.fecha + "-" + reg.parada;
                if (!agrupados[clave]) agrupados[clave] = { fecha: reg.fecha, parada: reg.parada, total: k, veces: 1 };
                else { agrupados[clave].total += k; agrupados[clave].veces += 1; }
            });
            Object.values(agrupados).forEach(function(inf) {
                var fila = document.createElement('tr');
                fila.innerHTML = '<td>' + inf.fecha + '</td><td>' + inf.parada + ' <small style="color:#7f8c8d;">(' + inf.veces + ' pesajes)</small></td><td>' + inf.total.toLocaleString('es-ES') + ' kg</td><td style="color:#7f8c8d;font-style:italic;font-size:12px;">Consolidado</td>';
                cuerpoTabla.appendChild(fila);
            });
        }
        
        if (tarjetaSumaKilos) tarjetaSumaKilos.textContent = totalKilos.toLocaleString('es-ES') + ' kg';
        if (document.getElementById('mayor-entrega')) document.getElementById('mayor-entrega').textContent = maxKilos.toLocaleString('es-ES') + ' kg';
        if (document.getElementById('total-registros')) document.getElementById('total-registros').textContent = filtrados.length + " pesajes";
    }

    arrancarPrograma();
});


function cambiarTituloParada(valor) {
    const titulo = document.querySelector('h4');
    if (titulo) {
        titulo.textContent = (valor === "TODAS" || valor === "") ? "TOTAL SELECCIONADO" : "TOTAL: " + valor.toUpperCase();
    }
}



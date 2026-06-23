// --- ESTADO DE LA APLICACIÓN ---
const MAX_CAPACIDAD = 28;
let totalTarimas = 0;
let datosTrailer = []; // Array en memoria para exportar a Excel

// --- REFERENCIAS AL DOM ---
const inputCaja = document.getElementById('inputCaja');
const inputMaterial = document.getElementById('inputMaterial');
const inputCantidad = document.getElementById('inputCantidad');
const mapaTrailer = document.getElementById('mapa-trailer');
const contadorTarimas = document.getElementById('contadorTarimas');

// --- NUEVA FUNCIÓN: ENCARGADA DE DIBUJAR TODO EL TRÁILER ---
function renderizarTrailer() {
    // 1. Limpiar por completo el contenedor visual para volverlo a armar
    mapaTrailer.innerHTML = '';
    
    // 2. Sincronizar el contador global con el tamaño real del arreglo
    totalTarimas = datosTrailer.length;
    contadorTarimas.innerText = totalTarimas;
    
    // 3. Recorrer el arreglo de datos y fabricar los cuadritos en el Grid
    datosTrailer.forEach((tarima, index) => {
        // Corrección automática de posición (por si borramos una intermedia, los números se reajustan de 1 a N)
        tarima.Posición_Tráiler = index + 1;
        
        // Crear el elemento visual
        const nuevaTarima = document.createElement('div');
        nuevaTarima.className = 'tarima';
        nuevaTarima.innerHTML = `
            <button class="btn-eliminar-tarima" title="Eliminar esta tarima">×</button>
            <strong>${tarima.Número_Material}</strong>
            <span>${tarima.Cantidad_Piezas} pcs</span>
        `;
        
        // ASIGNAR EVENTO DE ELIMINACIÓN A LA "X"
        const btnEliminar = nuevaTarima.querySelector('.btn-eliminar-tarima');
        btnEliminar.addEventListener('click', function() {
            // Borramos el elemento del arreglo usando su índice actual
            datosTrailer.splice(index, 1);
            
            // ¡Magia! Volvemos a renderizar para que la pantalla se actualice sola
            renderizarTrailer();
        });
        
        // Agregar al mapa del tráiler
        mapaTrailer.appendChild(nuevaTarima);
    });
}

// --- LÓGICA DE ESCÁNER (AUTO-FOCUS) ---

// Escanear Caja de trailer o ingresar
inputCaja.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
        if(inputCaja.value.trim() !== '') {
            inputMaterial.focus();
        }
    }
});

// Escanear Material 
inputMaterial.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
        if(inputMaterial.value.trim() !== '') {
            inputCantidad.focus();
        }
    }
});



// Escanear Cantidad: Procesa, Dibuja y Regresa a Material
inputCantidad.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault();

        const material = inputMaterial.value.trim().toUpperCase();
        const cantidad = parseInt(inputCantidad.value.trim());
        

        // Validar campos vacíos
        if (!material || isNaN(cantidad) || cantidad <= 0){
            alert("Cantidad invalida")
            inputCantidad.focus;
            return;

        } ;

       // Validar límite físico usando el tamaño real del array
        if (datosTrailer.length >= MAX_CAPACIDAD) {
            alert(`¡ALTO! La caja del tráiler está llena (${MAX_CAPACIDAD} tarimas máximo).`);
            inputMaterial.value = '';
            inputCantidad.value = '';
            inputMaterial.focus();
            return;
        }

        // Guardar los datos limpios
        datosTrailer.push({
            "Posición_Tráiler": datosTrailer.length + 1,
            "Número_Material": material,
            "Cantidad_Piezas": parseInt(cantidad)
        });

        // Llamar a nuestra función de dibujo
        renderizarTrailer();

        // Limpiar para el siguiente ciclo del escáner
        inputMaterial.value = '';
        inputCantidad.value = '';
        inputMaterial.focus();
    }
});

// --- LÓGICA DE EXPORTACIÓN Excel ---
document.getElementById('btnExportar').addEventListener('click', function() {
    if (datosTrailer.length === 0) {
        alert("No hay datos para exportar. Escanea al menos una tarima.");
        return;
    }

    // 1. Generar nombre de archivo dinámico
    const numCaja = inputCaja.value.trim().toUpperCase() || 'SINNUM';
    const hoy = new Date();
    const anio = hoy.getFullYear();
    const mes = String(hoy.getMonth() + 1).padStart(2, '0');
    const dia = String(hoy.getDate()).padStart(2, '0');
    const nombreSugerido = `Caja-${numCaja}-${anio}-${mes}-${dia}.xlsx`;

    try {
        // --- HOJA 1: MAPA VISUAL ---
        // Construimos una matriz de filas y columnas
        const matrizMapa = [];
        
        // Cabecera visual
        matrizMapa.push(["=== FRENTE DEL TRÁILER ===", "=== FRENTE DEL TRÁILER ==="]);

        // Crear la estructura vacía de 14 filas x 2 columnas
        for (let i = 0; i < 14; i++) {
            matrizMapa.push([`(Vacío)`, `(Vacío)`]);
        }

        // Llenar las coordenadas exactas con los datos escaneados
        datosTrailer.forEach(tarima => {
            const index = tarima.Posición_Tráiler - 1; // Índice de 0 a 27
            
            // Cálculos matemáticos simples para saber la celda exacta
            const filaExcel = Math.floor(index / 2) + 1; // +1 porque la fila 0 es el Frente
            const columnaExcel = index % 2; // 0 (Izquierda) o 1 (Derecha)
            
            // Formatear el texto de cada cuadrito en el Excel
            matrizMapa[filaExcel][columnaExcel] = `[T-${tarima.Posición_Tráiler}] ${tarima.Número_Material} | ${tarima.Cantidad_Piezas} pcs`;
        });

        // Puertas
        matrizMapa.push(["=== PUERTAS ===", "=== PUERTAS ==="]);

        // Convertir la matriz a hoja de Excel
        const hojaMapa = XLSX.utils.aoa_to_sheet(matrizMapa);
        
        // Ajustar el ancho de las 2 columnas para que el texto no se corte
        hojaMapa['!cols'] = [{ wch: 35 }, { wch: 35 }];


        // --- HOJA 2: DATOS TABULARES ---
        const hojaDatos = XLSX.utils.json_to_sheet(datosTrailer);


        // --- ARMAR EL LIBRO DE EXCEL CON LAS DOS HOJAS ---
        const libro = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(libro, hojaMapa, "Mapa_Visual");
        XLSX.utils.book_append_sheet(libro, hojaDatos, "Datos_ETL");

        // --- DESCARGAR EL ARCHIVO NATIVAMENTE ---
        XLSX.writeFile(libro, nombreSugerido);

    } catch (error) {
        console.error("Error crítico al generar Excel:", error);
        alert("Hubo un problema al generar el archivo. Presiona F12 para ver el error.");
    }
});

// --- BOTÓN CANCELAR / LIMPIAR ---
document.getElementById('btnCancelar').addEventListener('click', function() {
    if(confirm("¿Estás seguro de que quieres borrar todo el progreso actual?")) {
        // Reiniciar variables
        totalTarimas = 0;
        datosTrailer = [];
        
        // Limpiar UI
        contadorTarimas.innerText = '0';
        mapaTrailer.innerHTML = ''; // Borra todos los cuadritos
        inputCaja.value = '';
        inputMaterial.value = '';
        inputCantidad.value = '';
        
        inputCaja.focus();
    }
});
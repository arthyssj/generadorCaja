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

// --- LÓGICA DE ESCÁNER (AUTO-FOCUS) ---

// 1. Escanear Caja -> Salta a Material
inputCaja.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
        if(inputCaja.value.trim() !== '') {
            inputMaterial.focus();
        }
    }
});

// 2. Escanear Material -> Salta a Cantidad
inputMaterial.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
        if(inputMaterial.value.trim() !== '') {
            inputCantidad.focus();
        }
    }
});

// 3. Escanear Cantidad -> Procesa, Dibuja y Regresa a Material
inputCantidad.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault();

        const material = inputMaterial.value.trim().toUpperCase();
        const cantidad = inputCantidad.value.trim();

        // Validar campos vacíos
        if (!material || !cantidad) return;

        // Validar límite físico
        if (totalTarimas >= MAX_CAPACIDAD) {
            alert("¡ALTO! La caja del tráiler está llena (28 tarimas máximo).");
            inputMaterial.value = '';
            inputCantidad.value = '';
            inputMaterial.focus();
            return;
        }

        // --- PROCESAMIENTO DE DATOS ---
        totalTarimas++;

        // Guardar para el Excel (Ideal para Power Query)
        datosTrailer.push({
            "Posición_Tráiler": totalTarimas,
            "Número_Material": material,
            "Cantidad_Piezas": parseInt(cantidad)
        });

        // --- ACTUALIZACIÓN VISUAL ---
        // Crear el div de la tarima
        const nuevaTarima = document.createElement('div');
        nuevaTarima.className = 'tarima';
        nuevaTarima.innerHTML = `
            <strong>${material}</strong>
            <span>${cantidad} pcs</span>
        `;

        // CSS Grid lo acomodará automáticamente en zigzag
        mapaTrailer.appendChild(nuevaTarima);

        // Actualizar el número verde
        contadorTarimas.innerText = totalTarimas;

        // Limpiar para el siguiente ciclo del escáner
        inputMaterial.value = '';
        inputCantidad.value = '';
        inputMaterial.focus();
    }
});

// --- LÓGICA DE EXPORTACIÓN (EXCEL LOCAL) ---
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
        // --- HOJA 1: MAPA VISUAL (Para imprimir / Operadores) ---
        // Construimos una matriz de filas y columnas
        const matrizMapa = [];
        
        // Fila 1: Cabecera visual
        matrizMapa.push(["=== FRENTE DEL TRÁILER ===", "=== FRENTE DEL TRÁILER ==="]);

        // Crear la estructura vacía de 13 filas x 2 columnas
        for (let i = 0; i < 14; i++) {
            matrizMapa.push([`(Vacío)`, `(Vacío)`]);
        }

        // Llenar las coordenadas exactas con los datos escaneados
        datosTrailer.forEach(tarima => {
            const index = tarima.Posición_Tráiler - 1; // Índice de 0 a 25
            
            // Cálculos matemáticos simples para saber la celda exacta
            const filaExcel = Math.floor(index / 2) + 1; // +1 porque la fila 0 es el Frente
            const columnaExcel = index % 2; // 0 (Izquierda) o 1 (Derecha)
            
            // Formatear el texto de cada cuadrito en el Excel
            matrizMapa[filaExcel][columnaExcel] = `[T-${tarima.Posición_Tráiler}] ${tarima.Número_Material} | ${tarima.Cantidad_Piezas} pcs`;
        });

        // Fila Final: Puertas
        matrizMapa.push(["=== PUERTAS ===", "=== PUERTAS ==="]);

        // Convertir la matriz a hoja de Excel
        const hojaMapa = XLSX.utils.aoa_to_sheet(matrizMapa);
        
        // Ajustar el ancho de las 2 columnas para que el texto no se corte
        hojaMapa['!cols'] = [{ wch: 35 }, { wch: 35 }];


        // --- HOJA 2: DATOS TABULARES (Para Power Query) ---
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
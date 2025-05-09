/*
 * Copyright 2025 [Tu Nombre o Nombre de tu Sitio/Empresa]. Todos los derechos reservados.
 * Script para la Calculadora de Materiales Tablayeso.
 * Maneja la lógica de agregar ítems, calcular materiales y generar reportes.
 * Implementa el criterio de cálculo v2.0, con nombres específicos para Durock Calibre 20 y lógica de tornillos de 1".
 * Implementa lógica de cálculo de paneles con acumuladores fraccionarios/redondeados para áreas pequeñas/grandes (según criterio de imagen).
 * Implementa selección de tipo de panel por cara de muro y para cielos.
 * Implementa Múltiples Entradas de Medida (Segmentos) para Muros Y Cielos.
 * Ajusta el orden de las entradas.
 * Implementa cálculo de Angular de Lámina para cielos basado en el perímetro completo de segmentos.
 * Agrega opción para descontar metraje de Angular en cielos.
 * Agrega campo para Área de Trabajo en el cálculo general y lo incluye en reportes.
 * Añade manejo básico de errores en el cálculo.
 *
 * --- NUEVAS FUNCIONALIDADES ---
 * Implementa cálculo de Metraje (Área para Muros/Cielos, Lineal para Cenefas) con regla "menor a 1 = 1".
 * Permite agregar ítems de tipo Cenefa.
 * Añade selector para el tipo de Poste en ítems de tipo Muro.
 * Muestra metraje calculado por ítem automáticamente (dentro del bloque del ítem).
 * Muestra totales de metrajes (Muro Área, Cielo Área, Cenefa Lineal) en el resumen final.
 * Incluye metrajes en reportes PDF y Excel.
 * CORRECCIÓN: Lógica de cálculo de postes ajustada.
 * --- FIN NUEVAS FUNCIONALIDADES ---
 */

document.addEventListener('DOMContentLoaded', () => {
    const itemsContainer = document.getElementById('items-container');
    const addItemBtn = document.getElementById('add-item-btn');
    const calculateBtn = document.getElementById('calculate-btn');
    const resultsContent = document.getElementById('results-content');
    const downloadOptionsDiv = document.querySelector('.download-options');
    const generatePdfBtn = document.getElementById('generate-pdf-btn');
    const generateExcelBtn = document.getElementById('generate-excel-btn');

    let itemCounter = 0; // To give unique IDs to item blocks

    // Variables to store the last calculated state (needed for PDF/Excel)
    let lastCalculatedTotalMaterials = {}; // Stores final rounded totals for all materials
    let lastCalculatedTotalMetrajes = {}; // --- NUEVA VARIABLE para almacenar los totales de metraje ---
    let lastCalculatedItemsSpecs = []; // Specs of items included in calculation
    let lastErrorMessages = []; // Store errors as an array of strings
    let lastCalculatedWorkArea = ''; // --- Variable para almacenar el Área de Trabajo calculada ---


    // --- Constants ---
    const PANEL_RENDIMIENTO_M2 = 2.98; // m2 por panel (rendimiento estándar)
    const SMALL_AREA_THRESHOLD_M2 = 1.5; // Umbral para considerar un área "pequeña" (en m2 por cara/área total del SEGMENTO)

    // Definición de tipos de panel permitidos (deben coincidir con las opciones en el HTML)
    const PANEL_TYPES = [
        "Normal",
        "Resistente a la Humedad",
        "Resistente al Fuego",
        "Alta Resistencia",
        "Exterior" // Asociado comúnmente con Durock, pero aplicable si se usa ese tipo de panel en yeso especial
    ];

    // --- Definición de tipos de Poste (NUEVO) ---
    const POSTE_TYPES = [
        "Poste 2 1/2\" x 8' cal 26", // 2.44m
        "Poste 2 1/2\" x 10' cal 26", // 3.05m
        "Poste 2 1/2\" x 12' cal 26", // 3.66m
        "Poste 2 1/2\" x 8' cal 20",  // 2.44m
        "Poste 2 1/2\" x 10' cal 20"  // 3.05m
        // Considerar añadir los de 1 5/8 si son relevantes
        // "Poste 1 5/8\" x 8' cal 26",
        // "Poste 1 5/8\" x 10\' cal 26",
        // "Poste 1 5/8\" x 12\' cal 26"
    ];
    // --- Fin Definición de tipos de Poste ---


     // --- Helper Function for Rounding Up Final Units (Applies per item material quantity, EXCEPT panels in accumulators) ---
    const roundUpFinalUnit = (num) => Math.ceil(num);

    // --- Helper Function to get display name for item type ---
    const getItemTypeName = (typeValue) => {
        switch (typeValue) {
            case 'muro': return 'Muro';
            case 'cielo': return 'Cielo Falso';
            case 'cenefa': return 'Cenefa'; // --- NUEVO tipo Cenefa ---
            default: return 'Ítem Desconocido';
        }
    };

     // Helper to map item type internal value to a more descriptive name for inputs
     const getItemTypeDescription = (typeValue) => {
         switch (typeValue) {
             case 'muro': return 'Muro';
             case 'cielo': return 'Cielo Falso';
             case 'cenefa': return 'Cenefa'; // --- NUEVO tipo Cenefa ---
             default: return 'Ítem';
         }
     };


    // --- Helper Function to get the unit for a given material name ---
    // Mantiene la lógica existente, asegúrate de que los nombres de los materiales de Cenefa coincidan si son diferentes.
    const getMaterialUnit = (materialName) => {
         // Map specific names to units based on the new criterion
        // Material names can now include panel types, e.g., "Paneles de Normal"
        if (materialName.startsWith('Paneles de ')) return 'Und'; // Handle all panel types

        switch (materialName) {
            case 'Postes': return 'Und'; // Genérico, debería usarse el nombre específico ahora
            case 'Poste 2 1/2" x 8\' cal 26': return 'Und'; // Específico por tipo
            case 'Poste 2 1/2" x 10\' cal 26': return 'Und';
            case 'Poste 2 1/2" x 12\' cal 26': return 'Und';
            case 'Poste 2 1/2" x 8\' cal 20': return 'Und';
            case 'Poste 2 1/2" x 10\' cal 20': return 'Und';
             // Agregar tipos de poste 1 5/8 si se añaden a la lista de constantes
            // case 'Poste 1 5/8" x 8\' cal 26': return 'Und';
            // case 'Poste 1 5/8" x 10\' cal 26': return 'Und';
            // case 'Poste 1 5/8" x 12\' cal 26': return 'Und';

            case 'Canales': return 'Und'; // Genérico, debería usarse el nombre específico ahora
            case 'Canales Calibre 20': return 'Und'; // Asumiendo que solo hay un tipo de canal calibre 20 por ahora
            case 'Pasta': return 'Caja';
            case 'Cinta de Papel': return 'm';
            case 'Lija Grano 120': return 'Pliego';
            case 'Clavos con Roldana': return 'Und';
            case 'Fulminantes': return 'Und';
            case 'Tornillos de 1" punta fina': return 'Und';
            case 'Tornillos de 1/2" punta fina': return 'Und';
            case 'Canal Listón': return 'Und';
            case 'Canal Soporte': return 'Und';
            case 'Angular de Lámina': return 'Und';
            case 'Tornillos de 1" punta broca': return 'Und';
            case 'Tornillos de 1/2" punta broca': return 'Und';
            case 'Patas': return 'Und';
            case 'Canal Listón (para cuelgue)': return 'Und';
            case 'Basecoat': return 'Saco'; // Associated with Durock-like panels
            case 'Cinta malla': return 'm'; // Associated with Durock-like panels
             // Materiales específicos de Cenefa si los hubiera y no estuvieran ya listados
            default: return 'Und'; // Default unit if not specified
        }
    };

    // Helper function to get the associated finishing materials based on panel type
    const getFinishingMaterials = (panelType) => {
         const finishing = {};
         // Associate finishing materials based on the panel type name or a category derived from it
         if (panelType === 'Normal' || panelType === 'Resistente a la Humedad' || panelType === 'Resistente al Fuego' || panelType === 'Alta Resistencia') {
             finishing['Pasta'] = 0;
             finishing['Cinta de Papel'] = 0;
             finishing['Lija Grano 120'] = 0;
             finishing['Tornillos de 1" punta fina'] = 0; // Yeso type screws
             finishing['Tornillos de 1/2" punta fina'] = 0; // Yeso type screws for structure
         } else if (panelType === 'Exterior') { // Assuming 'Exterior' implies Durock or similar
             finishing['Basecoat'] = 0;
             finishing['Cinta malla'] = 0;
             finishing['Tornillos de 1" punta broca'] = 0; // Durock type screws
             finishing['Tornillos de 1/2" punta broca'] = 0; // Durock type screws for structure
         }
         return finishing;
    };

    // --- Function to Populate Panel Type Selects ---
    const populatePanelTypes = (selectElement, selectedValue = 'Normal') => {
        selectElement.innerHTML = ''; // Clear existing options
        PANEL_TYPES.forEach(type => {
            const option = document.createElement('option');
            option.value = type;
            option.textContent = type;
            if (type === selectedValue) {
                option.selected = true;
            }
            selectElement.appendChild(option);
        });
    };

    // --- Function to Populate Poste Type Selects (NUEVO) ---
     const populatePosteTypes = (selectElement, selectedValue = "Poste 2 1/2\" x 8' cal 26") => {
         selectElement.innerHTML = ''; // Clear existing options
         POSTE_TYPES.forEach(type => {
             const option = document.createElement('option');
             option.value = type;
             option.textContent = type;
             if (type === selectedValue) {
                 option.selected = true;
             }
             selectElement.appendChild(option);
         });
     };
    // --- Fin Function to Populate Poste Type Selects ---


    // --- Helper Function to update the summary details displayed within a segment block ---
    // Mantiene la lógica existente para Muros/Cielos
    const updateSegmentItemSummary = (segmentBlock) => {
        // Find the parent item block from the segment block
        const itemBlock = segmentBlock.closest('.item-block');
        if (!itemBlock) {
            console.error("Could not find parent item block for segment.");
            return; // Exit if parent not found
        }

        const segmentSummaryDiv = segmentBlock.querySelector('.segment-item-summary');
        if (!segmentSummaryDiv) {
             console.error("Could not find segment summary div.");
             return; // Exit if summary div not found
        }

        // Read parent item options
        const type = itemBlock.querySelector('.item-structure-type').value;
        const itemNumber = itemBlock.dataset.itemId.split('-')[1]; // Get item number from item ID like 'item-1' -> '1'

        let summaryText = `${getItemTypeDescription(type)} #${itemNumber} - `;

        if (type === 'muro') {
            const facesInput = itemBlock.querySelector('.item-faces');
            const faces = facesInput && !facesInput.closest('.hidden') ? parseInt(facesInput.value) : 1;

            const cara1PanelSelect = itemBlock.querySelector('.item-cara1-panel-type');
            const cara1PanelType = cara1PanelSelect && !cara1PanelSelect.closest('.hidden') ? cara1PanelSelect.value : 'N/A';

            const cara2PanelSelect = itemBlock.querySelector('.item-cara2-panel-type');
            const cara2PanelType = (faces === 2 && cara2PanelSelect && !cara2PanelSelect.closest('.hidden')) ? cara2PanelSelect.value : 'N/A';

            const postSpacingInput = itemBlock.querySelector('.item-post-spacing');
            const postSpacing = postSpacingInput && !postSpacingInput.closest('.hidden') ? parseFloat(postSpacingInput.value) : 'N/A';

            const postTypeSelect = itemBlock.querySelector('.item-poste-type'); // --- LEE el tipo de poste (NUEVO) ---
            const postType = postTypeSelect && !postTypeSelect.closest('.hidden') ? postTypeSelect.value : 'N/A'; // --- Obtiene el valor (NUEVO) ---


            const isDoubleStructureInput = itemBlock.querySelector('.item-double-structure');
            const isDoubleStructure = isDoubleStructureInput && !isDoubleStructureInput.closest('.hidden') ? isDoubleStructureInput.checked : false;

            summaryText += `${faces} Cara${faces > 1 ? 's' : ''}, Panel C1: ${cara1PanelType}`;
            if (faces === 2) summaryText += `, Panel C2: ${cara2PanelType}`;
            if (postSpacing !== 'N/A' && !isNaN(postSpacing)) summaryText += `, Esp: ${postSpacing.toFixed(2)}m`;
            if (postType !== 'N/A') summaryText += `, Poste: ${postType}`; // --- Añade tipo de poste al resumen (NUEVO) ---
            if (isDoubleStructure) summaryText += `, Doble Estructura`;

        } else if (type === 'cielo') {
            const cieloPanelSelect = itemBlock.querySelector('.item-cielo-panel-type');
            const cieloPanelType = cieloPanelSelect && !cieloPanelSelect.closest('.hidden') ? cieloPanelSelect.value : 'N/A';

            const plenumInput = itemBlock.querySelector('.item-plenum');
            const plenum = plenumInput && !plenumInput.closest('.hidden') ? parseFloat(plenumInput.value) : 'N/A';

            const angularDeductionInput = itemBlock.querySelector('.item-angular-deduction');
             // Read the deduction value to display it in the segment summary if needed
            const angularDeduction = angularDeductionInput && !angularDeductionInput.closest('.hidden') ? parseFloat(angularDeductionInput.value) : 'N/A';


            summaryText += `Panel: ${cieloPanelType}`;
            if (plenum !== 'N/A' && !isNaN(plenum)) summaryText += `, Pleno: ${plenum.toFixed(2)}m`;
            // Include angular deduction in the segment summary if desired
            if (angularDeduction !== 'N/A' && !isNaN(angularDeduction) && angularDeduction > 0) summaryText += `, Desc. Ang: ${angularDeduction.toFixed(2)}m`;

         } else if (type === 'cenefa') { // --- NUEVO: Manejo de Cenefa en el resumen del segmento (aunque cenefa no tiene segmentos, el summary div existe) ---
             // Para cenefa, no hay resumen a nivel de segmento, solo a nivel de ítem.
             // Podríamos ocultar el div o dejarlo vacío. Dejémoslo vacío por ahora.
              summaryText = ''; // Borra el texto por defecto para cenefa en el segmento

         } else {
             summaryText += "Configuración Desconocida"; // Fallback for unknown type
         }

        // Update the text content of the dedicated summary div within the segment
        segmentSummaryDiv.textContent = summaryText;
    };


     // --- Function to Create a Muro Segment Input Block ---
    const createMuroSegmentBlock = (itemId, segmentNumber) => {
        const segmentHtml = `
            <div class="muro-segment" data-segment-id="${itemId}-mseg-${segmentNumber}">
                 <div class="segment-header-line"> <h4>Segmento ${segmentNumber}</h4>
                    <div class="segment-item-summary"></div>
                    <div class="segment-metraje-display"></div>
                    </div>
                 <button type="button" class="remove-segment-btn">X</button> <div class="input-group">
                    <label for="mwidth-${itemId}-mseg-${segmentNumber}">Ancho (m):</label>
                    <input type="number" class="item-width" id="mwidth-${itemId}-mseg-${segmentNumber}" step="0.01" min="0" value="3.0">
                </div>
                <div class="input-group">
                    <label for="mheight-${itemId}-mseg-${segmentNumber}">Alto (m):</label>
                    <input type="number" class="item-height" id="mheight-${itemId}-mseg-${segmentNumber}" step="0.01" min="0" value="2.4">
                </div>
            </div>
        `;
        const newElement = document.createElement('div');
        newElement.innerHTML = segmentHtml.trim();
        const segmentBlock = newElement.firstChild;

        // Add remove listener
        const removeButton = segmentBlock.querySelector('.remove-segment-btn');
         removeButton.addEventListener('click', () => {
            const segmentsContainer = segmentBlock.closest('.segments-list'); // Correct selector
            if (segmentsContainer.querySelectorAll('.muro-segment').length > 1) {
                 segmentBlock.remove();
                 // Re-number segments visually after removal
                 segmentsContainer.querySelectorAll('.muro-segment h4').forEach((h4, index) => {
                    h4.textContent = `Segmento ${index + 1}`;
                 });
                 // --- Recalcula y muestra el metraje del ítem después de eliminar un segmento ---
                 const itemBlock = segmentsContainer.closest('.item-block');
                 calculateAndDisplayItemMetraje(itemBlock); // Llama a la función de metraje del ítem
                 // --- Fin Recálculo Metraje Ítem ---

                 // Clear results and hide download buttons after removal
                 resultsContent.innerHTML = '<p>Segmento eliminado. Recalcula los materiales totales.</p>';
                 downloadOptionsDiv.classList.add('hidden');
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedTotalMetrajes = {}; // Limpia metrajes totales almacenados
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
                 lastCalculatedWorkArea = ''; // Clear stored data on item removal
            } else {
                 alert("Un muro debe tener al menos un segmento.");
            }
         });

        // --- Agregar listeners a los inputs de dimensión del segmento para actualizar el metraje del ítem (NUEVO) ---
         const widthInput = segmentBlock.querySelector('.item-width');
         const heightInput = segmentBlock.querySelector('.item-height');
         const itemBlock = segmentBlock.closest('.item-block'); // Obtiene el bloque del ítem padre

         const updateMetraje = () => {
             calculateAndDisplayItemMetraje(itemBlock); // Llama a la función de metraje del ítem padre
         };

         widthInput.addEventListener('input', updateMetraje);
         heightInput.addEventListener('input', updateMetraje);
        // --- Fin Agregar listeners ---


        return segmentBlock;
    };

     // --- Function to Create a Cielo Segment Input Block ---
     const createCieloSegmentBlock = (itemId, segmentNumber) => {
         const segmentHtml = `
            <div class="cielo-segment" data-segment-id="${itemId}-cseg-${segmentNumber}">
                 <div class="segment-header-line"> <h4>Segmento ${segmentNumber}</h4>
                    <div class="segment-item-summary"></div>
                    <div class="segment-metraje-display"></div>
                    </div>
                 <button type="button" class="remove-segment-btn">X</button> <div class="input-group">
                    <label for="cwidth-${itemId}-cseg-${segmentNumber}">Ancho (m):</label>
                    <input type="number" class="item-width" id="cwidth-${itemId}-cseg-${segmentNumber}" step="0.01" min="0" value="3.0">
                </div>
                <div class="input-group">
                    <label for="clength-${itemId}-cseg-${segmentNumber}">Largo (m):</label>
                    <input type="number" class="item-length" id="clength-${itemId}-cseg-${segmentNumber}" step="0.01" min="0" value="4.0">
                </div>
            </div>
        `;
        const newElement = document.createElement('div');
        newElement.innerHTML = segmentHtml.trim();
        const segmentBlock = newElement.firstChild;

         // Add remove listener
         const removeButton = segmentBlock.querySelector('.remove-segment-btn');
         removeButton.addEventListener('click', () => {
             const segmentsContainer = segmentBlock.closest('.segments-list'); // Correct selector
             if (segmentsContainer.querySelectorAll('.cielo-segment').length > 1) {
                  segmentBlock.remove();
                  // Re-number segments visually after removal
                  segmentsContainer.querySelectorAll('.cielo-segment h4').forEach((h4, index) => {
                     h4.textContent = `Segmento ${index + 1}`;
                  });
                  // --- Recalcula y muestra el metraje del ítem después de eliminar un segmento ---
                  const itemBlock = segmentsContainer.closest('.item-block');
                  calculateAndDisplayItemMetraje(itemBlock); // Llama a la función de metraje del ítem
                  // --- Fin Recálculo Metraje Ítem ---


                  // Clear results and hide download buttons after removal
                  resultsContent.innerHTML = '<p>Segmento de cielo agregado. Recalcula los materiales totales.</p>';
                  downloadOptionsDiv.classList.add('hidden');
                  lastCalculatedTotalMaterials = {};
                  lastCalculatedTotalMetrajes = {}; // Limpia metrajes totales almacenados
                  lastCalculatedItemsSpecs = [];
                  lastErrorMessages = [];
                  lastCalculatedWorkArea = ''; // Clear stored data on adding/removing segment
             } else {
                  alert("Un cielo falso debe tener al menos un segmento.");
             }
         });

         // --- Agregar listeners a los inputs de dimensión del segmento para actualizar el metraje del ítem (NUEVO) ---
         const widthInput = segmentBlock.querySelector('.item-width');
         const lengthInput = segmentBlock.querySelector('.item-length'); // Usa .item-length para cielo
         const itemBlock = segmentBlock.closest('.item-block'); // Obtiene el bloque del ítem padre

         const updateMetraje = () => {
              calculateAndDisplayItemMetraje(itemBlock); // Llama a la función de metraje del ítem padre
         };

         widthInput.addEventListener('input', updateMetraje);
         lengthInput.addEventListener('input', updateMetraje);
         // --- Fin Agregar listeners ---

         return segmentBlock;
     };

    // --- Function to Create a Cenefa Input Block (NUEVO) ---
    // Cenefas no tendrán segmentos en esta lógica
    const createCenefaBlock = (itemId) => {
         const cenefaHtml = `
             <div class="cenefa-inputs" data-item-id="${itemId}">
                 <div class="input-group">
                    <label for="cenefa-length-${itemId}">Largo Total (m):</label>
                    <input type="number" class="item-length" id="cenefa-length-${itemId}" step="0.01" min="0" value="5.0">
                 </div>
                 <div class="input-group">
                    <label for="cenefa-faces-${itemId}">Nº de Caras:</label>
                    <input type="number" class="item-faces" id="cenefa-faces-${itemId}" step="1" min="1" value="2">
                 </div>
                 <div class="input-group">
                    <label for="cenefa-panel-type-${itemId}">Tipo de Panel:</label>
                    <select class="item-cielo-panel-type" id="cenefa-panel-type-${itemId}"></select> </div>
             </div>
         `;
        const newElement = document.createElement('div');
        newElement.innerHTML = cenefaHtml.trim();
        const cenefaBlock = newElement.firstChild; // Get the actual div element

        // --- Agregar listeners a los inputs de la cenefa para actualizar el metraje del ítem (NUEVO) ---
        const lengthInput = cenefaBlock.querySelector('.item-length');
        const facesInput = cenefaBlock.querySelector('.item-faces'); // Reutiliza la clase .item-faces para cantidad de caras
        const itemBlock = cenefaBlock.closest('.item-block'); // Obtiene el bloque del ítem padre

        const updateMetraje = () => {
             calculateAndDisplayItemMetraje(itemBlock); // Llama a la función de metraje del ítem padre
        };

        if (lengthInput) lengthInput.addEventListener('input', updateMetraje);
        if (facesInput) facesInput.addEventListener('input', updateMetraje);
        // El listener para el panel type se añade más abajo en el listener del select principal del ítem,
        // ya que updateItemInputVisibility recrea los inputs de cenefa a veces.
        // --- Fin Agregar listeners ---


         return cenefaBlock;
    };
    // --- Fin Function to Create a Cenefa Input Block ---


     // --- Function to Update Input Visibility WITHIN an Item Block ---
    // Modificada para incluir el tipo Cenefa y el selector de poste
    const updateItemInputVisibility = (itemBlock) => {
        const structureTypeSelect = itemBlock.querySelector('.item-structure-type');
        const type = structureTypeSelect.value;

        // Common input groups (some are hidden/shown based on type)
        const facesInputGroup = itemBlock.querySelector('.item-faces-input'); // Usado para Muro (caras) y Cenefa (cantidad caras)
        const muroPanelTypesDiv = itemBlock.querySelector('.muro-panel-types');
        const cieloPanelTypeDiv = itemBlock.querySelector('.cielo-panel-type'); // Usado para Cielo y Cenefa panel type select
        const postSpacingInputGroup = itemBlock.querySelector('.item-post-spacing-input');
        const postTypeInputGroup = itemBlock.querySelector('.item-poste-type-input'); // --- NUEVO: Input Group para Tipo de Poste ---
        const doubleStructureInputGroup = itemBlock.querySelector('.item-double-structure-input');
        const plenumInputGroup = itemBlock.querySelector('.item-plenum-input');
        const angularDeductionInputGroup = itemBlock.querySelector('.item-angular-deduction-input');

        // Type-specific dimension/segment/item-level input containers
        const muroSegmentsContainer = itemBlock.querySelector('.muro-segments');
        const cieloSegmentsContainer = itemBlock.querySelector('.cielo-segments');
        const cenefaInputsContainer = itemBlock.querySelector('.cenefa-inputs'); // --- NUEVO: Contenedor de Inputs para Cenefa ---


        // Reset visibility for ALL type-specific input groups within this block
        // Oculta todos los contenedores principales específicos de tipo
        if (muroPanelTypesDiv) muroPanelTypesDiv.classList.add('hidden');
        if (cieloPanelTypeDiv) cieloPanelTypeDiv.classList.add('hidden'); // Oculta el contenedor compartido (cielo/cenefa panel type)
        if (postSpacingInputGroup) postSpacingInputGroup.classList.add('hidden');
        if (postTypeInputGroup) postTypeInputGroup.classList.add('hidden');
        if (doubleStructureInputGroup) doubleStructureInputGroup.classList.add('hidden');
        if (plenumInputGroup) plenumInputGroup.classList.add('hidden');
        if (muroSegmentsContainer) muroSegmentsContainer.classList.add('hidden');
        if (cieloSegmentsContainer) cieloSegmentsContainer.classList.add('hidden');
        if (angularDeductionInputGroup) angularDeductionInputGroup.classList.add('hidden');

         // Oculta el input de caras que es compartido por Muro y Cenefa. Se mostrará condicionalmente después.
         if (facesInputGroup) facesInputGroup.classList.add('hidden');

         // Oculta el contenedor de inputs de Cenefa
         if (cenefaInputsContainer) cenefaInputsContainer.classList.add('hidden');


         // --- Oculta el div de resumen de metraje por segmento si no aplica (Cenefa no tiene segmentos) ---
         // Los divs de metraje por segmento están dentro de los contenedores de segmentos.
         // Ya se ocultan si el contenedor principal de segmentos está oculto.
         // Pero el div de metraje por ítem (item-metraje-display) está directamente en el bloque del ítem padre.
         const itemMetrajeDisplay = itemBlock.querySelector('.item-metraje-display');
         if (itemMetrajeDisplay) {
              itemMetrajeDisplay.classList.remove('hidden'); // Muestra el div de metraje por ítem por defecto
         }


        // Set visibility based on selected type for THIS block
        if (type === 'muro') {
            if (facesInputGroup) facesInputGroup.classList.remove('hidden'); // Muestra el input de caras para muro
            if (muroPanelTypesDiv) muroPanelTypesDiv.classList.remove('hidden'); // Show wall panel type selectors
            if (postSpacingInputGroup) postSpacingInputGroup.classList.remove('hidden'); // Post spacing applies to walls
            if (postTypeInputGroup) postTypeInputGroup.classList.remove('hidden'); // --- MUESTRA el input de tipo de poste (NUEVO) ---
            if (doubleStructureInputGroup) doubleStructureInputGroup.classList.remove('hidden'); // Double structure applies to walls
            if (muroSegmentsContainer) muroSegmentsContainer.classList.remove('hidden'); // Show muro segments container

             // Update visibility of face-specific panel type selectors based on faces input
             const facesInput = itemBlock.querySelector('.item-faces');
             const cara2PanelTypeGroup = itemBlock.querySelector('.cara-2-panel-type-group');

             if (cara2PanelTypeGroup) { // Asegura que el grupo exista
                 if (facesInput && parseInt(facesInput.value) === 2) {
                     cara2PanelTypeGroup.classList.remove('hidden');
                 } else {
                     cara2PanelTypeGroup.classList.add('hidden');
                 }
             }


        } else if (type === 'cielo') {
            if (cieloPanelTypeDiv) cieloPanelTypeDiv.classList.remove('hidden'); // Show ceiling panel type selector (contenedor compartido)
            if (plenumInputGroup) plenumInputGroup.classList.remove('hidden');
            if (cieloSegmentsContainer) cieloSegmentsContainer.classList.remove('hidden'); // Show cielo segments container
            if (angularDeductionInputGroup) angularDeductionInputGroup.classList.remove('hidden'); // Muestra el input de descuento angular para cielos

         } else if (type === 'cenefa') { // --- NUEVO: Lógica para Cenefa ---
             if (cenefaInputsContainer) cenefaInputsContainer.classList.remove('hidden'); // Muestra los inputs de Cenefa
             if (facesInputGroup) facesInputGroup.classList.remove('hidden'); // Muestra el input de cantidad de caras (reutiliza la clase .item-faces)
             if (cieloPanelTypeDiv) cieloPanelTypeDiv.classList.remove('hidden'); // Muestra el selector de tipo de panel (reutiliza la clase .item-cielo-panel-type)
             // Para cenefas, el resumen del metraje se muestra a nivel de ítem, no de segmento.
             // Los divs .segment-metraje-display ya están ocultos porque los contenedores de segmentos están ocultos.
             // El div .item-metraje-display está visible por defecto al inicio de la función.

        }
        // No need for 'else' as all are hidden by default initially
    };

    // --- Function to update the main item header summary (kept for consistency, but not used for detailed summary anymore) ---
    // The detailed summary is now handled within each segment block.
    // Modificada para incluir Cenefa
    const updateItemHeaderSummary = (itemBlock) => {
         const itemHeader = itemBlock.querySelector('h3');
         const type = itemBlock.querySelector('.item-structure-type').value;
         const itemNumber = itemBlock.dataset.itemId.split('-')[1]; // Get item number from item ID

         // Update the main header to just show Type and Number
         itemHeader.textContent = `${getItemTypeDescription(type)} #${itemNumber}`;
     };


    // --- Function to Create an Item Input Block ---
    // Modificada para incluir el tipo Cenefa y el selector de poste
    const createItemBlock = () => {
        itemCounter++;
        const itemId = `item-${itemCounter}`;

        // Restructured HTML template
        const itemHtml = `
            <div class="item-block" data-item-id="${itemId}">
                <h3>${getItemTypeDescription('muro')} #${itemCounter}</h3> <button class="remove-item-btn">Eliminar</button>

                <div class="item-metraje-display">Metraje calculado para este ítem: -</div>
                <div class="input-group">
                    <label for="type-${itemId}">Tipo de Estructura:</label>
                    <select class="item-structure-type" id="type-${itemId}">
                        <option value="muro">Muro</option>
                        <option value="cielo">Cielo Falso</option>
                        <option value="cenefa">Cenefa</option> </select>
                </div>

                <div class="input-group item-faces-input">
                    <label for="faces-${itemId}">Nº de Caras:</label> <input type="number" class="item-faces" id="faces-${itemId}" step="1" min="1" value="1"> </div>

                <div class="muro-panel-types">
                    <div class="input-group cara-1-panel-type-group">
                        <label for="cara1-panel-type-${itemId}">Panel Cara 1:</label>
                        <select class="item-cara1-panel-type" id="cara1-panel-type-${itemId}"></select>
                    </div>
                    <div class="input-group cara-2-panel-type-group hidden">
                        <label for="cara2-panel-type-${itemId}">Panel Cara 2:</label>
                        <select class="item-cara2-panel-type" id="cara2-panel-type-${itemId}"></select>
                    </div>
                </div>

                <div class="input-group item-post-spacing-input">
                    <label for="post-spacing-${itemId}">Espaciamiento Postes (m):</label>
                    <input type="number" class="item-post-spacing" id="post-spacing-${itemId}" step="0.01" min="0.1" value="0.40">
                </div>

                 <div class="input-group item-poste-type-input">
                     <label for="poste-type-${itemId}">Tipo de Poste:</label>
                     <select class="item-poste-type" id="poste-type-${itemId}"></select>
                 </div>
                <div class="input-group item-double-structure-input">
                    <label for="double-structure-${itemId}">Estructura Doble:</label>
                    <input type="checkbox" class="item-double-structure" id="double-structure-${itemId}">
                </div>

                <div class="input-group cielo-panel-type"> <label for="cielo-panel-type-${itemId}">Tipo de Panel:</label> <select class="item-cielo-panel-type" id="cielo-panel-type-${itemId}"></select>
                </div>

                <div class="input-group item-plenum-input hidden">
                    <label for="plenum-${itemId}">Pleno del Cielo (m):</label>
                    <input type="number" class="item-plenum" id="plenum-${itemId}" step="0.01" min="0" value="0.5">
                </div>

                <div class="input-group item-angular-deduction-input hidden">
                    <label for="angular-deduction-${itemId}">Metros a descontar de Angular:</label>
                    <input type="number" class="item-angular-deduction" id="angular-deduction-${itemId}" step="0.01" min="0" value="0">
                </div>

                <div class="cenefa-inputs hidden">
                      </div>
                <div class="muro-segments">
                    <h4>Segmentos del Muro:</h4>
                     <div class="segments-list">
                         </div>
                    <button type="button" class="add-segment-btn">Agregar Segmento de Muro</button>
                </div>

                 <div class="cielo-segments hidden">
                    <h4>Segmentos del Cielo Falso:</h4>
                     <div class="segments-list">
                         </div>
                    <button type="button" class="add-segment-btn">Agregar Segmento de Cielo</button>
                 </div>

            </div>
        `;

        const newElement = document.createElement('div');
        newElement.innerHTML = itemHtml.trim();
        const itemBlock = newElement.firstChild; // Get the actual div element

        itemsContainer.appendChild(itemBlock);

        // Actualiza el encabezado principal del ítem al crearlo
        updateItemHeaderSummary(itemBlock);


        // Add an initial segment based on the DEFAULT type ('muro')
        const muroSegmentsListContainer = itemBlock.querySelector('.muro-segments .segments-list');
        if (muroSegmentsListContainer) {
             const initialSegment = createMuroSegmentBlock(itemId, 1);
             muroSegmentsListContainer.appendChild(initialSegment);
             // --- Llama a la función de resumen y metraje para el segmento inicial de muro ---
             updateSegmentItemSummary(initialSegment); // Update segment summary
             // El cálculo de metraje del item se hace después del event listener change del select principal
             // o en los listeners de los inputs de segmento, y al final de createItemBlock.
             // calculateAndDisplayItemMetraje(itemBlock); // NO llamar aquí para el segmento inicial
             // --- Fin llamado ---


             // Add listener for "Agregar Segmento" button for muro
             const addMuroSegmentBtn = itemBlock.querySelector('.muro-segments .add-segment-btn');
             if (addMuroSegmentBtn) { // Asegura que el botón exista
                 addMuroSegmentBtn.addEventListener('click', () => {
                     const currentSegments = muroSegmentsListContainer.querySelectorAll('.muro-segment').length;
                     const newSegment = createMuroSegmentBlock(itemId, currentSegments + 1);
                     muroSegmentsListContainer.appendChild(newSegment);
                     // --- Llama a la función de resumen y metraje para el nuevo segmento de muro ---
                     updateSegmentItemSummary(newSegment); // Update segment summary
                     calculateAndDisplayItemMetraje(itemBlock); // Calculate and display item metraje
                     // --- Fin llamado ---

                     // Clear results and hide download buttons after adding a segment
                     resultsContent.innerHTML = '<p>Segmento de muro agregado. Recalcula los materiales y metrajes totales.</p>'; // Mensaje actualizado
                     downloadOptionsDiv.classList.add('hidden');
                     lastCalculatedTotalMaterials = {};
                     lastCalculatedTotalMetrajes = {}; // Limpia metrajes totales almacenados
                     lastCalculatedItemsSpecs = [];
                     lastErrorMessages = [];
                     lastCalculatedWorkArea = ''; // Clear stored data on adding/removing segment
                 });
             }
        }

         // Add listener for "Agregar Segmento" button for cielo (initially hidden)
         const cieloSegmentsListContainer = itemBlock.querySelector('.cielo-segments .segments-list');
          if (cieloSegmentsListContainer) {
              // No creamos un segmento inicial de cielo aquí porque el item por defecto es muro.
              // El segmento inicial de cielo se creará cuando se cambie el tipo a 'cielo'.
             const addCieloSegmentBtn = itemBlock.querySelector('.cielo-segments .add-segment-btn');
             if (addCieloSegmentBtn) { // Asegura que el botón exista
                 addCieloSegmentBtn.addEventListener('click', () => {
                     const currentSegments = cieloSegmentsListContainer.querySelectorAll('.cielo-segment').length;
                     const newSegment = createCieloSegmentBlock(itemId, currentSegments + 1);
                     cieloSegmentsListContainer.appendChild(newSegment);
                     // --- Llama a la función de resumen y metraje para el nuevo segmento de cielo ---
                     updateSegmentItemSummary(newSegment); // Update segment summary
                     calculateAndDisplayItemMetraje(itemBlock); // Calculate and display item metraje
                     // --- Fin llamado ---


                     // Clear results and hide download buttons after adding a segment
                     resultsContent.innerHTML = '<p>Segmento de cielo agregado. Recalcula los materiales totales.</p>'; // Mensaje actualizado
                     downloadOptionsDiv.classList.add('hidden');
                     lastCalculatedTotalMaterials = {};
                     lastCalculatedTotalMetrajes = {}; // Limpia metrajes totales almacenados
                     lastCalculatedItemsSpecs = [];
                     lastErrorMessages = [];
                     lastCalculatedWorkArea = ''; // Clear stored data on adding/removing segment
                 });
             }
          }

         // --- NUEVO: Lógica para crear los inputs de Cenefa cuando se selecciona el tipo ---
         // Ya la manejamos en el listener del select principal item-structure-type.
         // Solo necesitamos asegurarnos de que el contenedor .cenefa-inputs esté en el HTML base, lo cual ya hicimos.

        // Populate panel type selects in the new block
        const cara1PanelSelect = itemBlock.querySelector('.item-cara1-panel-type');
        const cara2PanelSelect = itemBlock.querySelector('.item-cara2-panel-type');
        const cieloPanelSelect = itemBlock.querySelector('.item-cielo-panel-type'); // Reutilizado para Cenefa panel type
        if(cara1PanelSelect) populatePanelTypes(cara1PanelSelect);
        if(cara2PanelSelect) populatePanelTypes(cara2PanelSelect);
        // Popula el selector reutilizado para cielo/cenefa SOLO si existe (podría no existir si la estructura HTML es diferente)
        // Pero con el html actual siempre existe.
        if(cieloPanelSelect) populatePanelTypes(cieloPanelSelect);


        // --- Populate Poste type select (NUEVO) ---
        const posteTypeSelect = itemBlock.querySelector('.item-poste-type');
        if(posteTypeSelect) populatePosteTypes(posteTypeSelect);
        // --- Fin Populate Poste type select ---


        // Add event listener to the new select element IN THIS BLOCK
        const structureTypeSelect = itemBlock.querySelector('.item-structure-type');
        if (structureTypeSelect) { // Asegura que el select exista
            structureTypeSelect.addEventListener('change', (event) => {
                const selectedType = event.target.value;
                const itemId = itemBlock.dataset.itemId; // Obtiene el item ID

                // Actualizamos el título principal del ítem para reflejar el tipo y número.
                updateItemHeaderSummary(itemBlock); // Llama a la función que ahora solo actualiza el h3 con tipo y número.

                // Clear existing segments when changing type
                const muroSegmentsList = itemBlock.querySelector('.muro-segments .segments-list');
                const cieloSegmentsList = itemBlock.querySelector('.cielo-segments .segments-list');
                const cenefaInputsContainer = itemBlock.querySelector('.cenefa-inputs');


                if (selectedType === 'muro') {
                    // Clear cielo segments and add a muro segment if needed
                    if (cieloSegmentsList) cieloSegmentsList.innerHTML = '';
                    // Clear cenefa inputs if they exist
                     if (cenefaInputsContainer) cenefaInputsContainer.innerHTML = '';


                    if (muroSegmentsList && muroSegmentsList.querySelectorAll('.muro-segment').length === 0) {
                         const newSegment = createMuroSegmentBlock(itemId, 1);
                         muroSegmentsList.appendChild(newSegment);
                         updateSegmentItemSummary(newSegment); // Update summary for newly created segment
                         // calculateAndDisplayItemMetraje(itemBlock); // Se llama después de updateItemInputVisibility
                    }
                } else if (selectedType === 'cielo') {
                     // Clear muro segments and add a cielo segment if needed
                     if (muroSegmentsList) muroSegmentsList.innerHTML = '';
                     // Clear cenefa inputs if they exist
                      if (cenefaInputsContainer) cenefaInputsContainer.innerHTML = '';


                     if (cieloSegmentsList && cieloSegmentsList.querySelectorAll('.cielo-segment').length === 0) {
                         const newSegment = createCieloSegmentBlock(itemId, 1);
                         cieloSegmentsList.appendChild(newSegment);
                         updateSegmentItemSummary(newSegment); // Update summary for newly created segment
                         // calculateAndDisplayItemMetraje(itemBlock); // Se llama después de updateItemInputVisibility
                     }
                } else if (selectedType === 'cenefa') { // --- NUEVO: Lógica al cambiar a Cenefa ---
                    // Clear muro and cielo segments
                    if (muroSegmentsList) muroSegmentsList.innerHTML = '';
                    if (cieloSegmentsList) cieloSegmentsList.innerHTML = '';

                     // Add cenefa inputs if they don't exist yet (check by looking for a key input inside the container)
                     const currentCenefaLengthInput = cenefaInputsContainer ? cenefaInputsContainer.querySelector('.item-length') : null;
                     if (cenefaInputsContainer && !currentCenefaLengthInput) { // Asegura que el contenedor exista y no tenga ya inputs
                         const cenefaBlock = createCenefaBlock(itemId);
                         cenefaInputsContainer.appendChild(cenefaBlock);
                         // Populate panel type select for cenefa (it's inside the created cenefaBlock now)
                         const cenefaPanelSelect = cenefaInputsContainer.querySelector('.item-cielo-panel-type');
                         if(cenefaPanelSelect) populatePanelTypes(cenefaPanelSelect); // Reutiliza la función

                         // Establece el valor por defecto de caras a 2 para cenefas al crearlas si el input existe
                         const facesInput = itemBlock.querySelector('.item-faces'); // Ya existe un input-group .item-faces-input compartido
                         // No necesitamos crear un nuevo input de caras dentro de createCenefaBlock si ya tenemos uno global.
                         // La lógica de createCenefaBlock fue simplificada para reflejar esto.
                         // Asegurémonos de que el input .item-faces global tenga el valor 2 cuando se cambia a cenefa.
                          if (facesInput) {
                               facesInput.value = 2;
                               // Dispara un evento 'input' o 'change' para que el listener de metraje se active
                               const event = new Event('input'); // o 'change'
                               facesInput.dispatchEvent(event);
                           }
                         // Dispara un evento 'input' para el Largo de Cenefa si existe, para calcular el metraje inicial
                         const cenefaLengthInput = cenefaInputsContainer.querySelector('.item-length');
                         if (cenefaLengthInput) {
                              const event = new Event('input');
                              cenefaLengthInput.dispatchEvent(event);
                         }
                     }
                }

                updateItemInputVisibility(itemBlock); // Update visibility after changing structure/clearing segments

                // --- Llama a la función de resumen para *todos* los segmentos existentes del ítem después de cambiar el tipo ---
                // Esto asegura que si cambias de muro a cielo, los segmentos existentes (ahora de cielo) actualicen su resumen.
                // Y también actualiza el metraje del ítem completo.
                itemBlock.querySelectorAll('.muro-segment, .cielo-segment').forEach(segBlock => {
                    updateSegmentItemSummary(segBlock);
                    // No necesitamos llamar a calculateAndDisplayItemMetraje desde aquí para cada segmento,
                    // la llamada global para el ítem es suficiente.
                });
                calculateAndDisplayItemMetraje(itemBlock); // Calculate and display item metraje after type change
                // --- Fin llamado ---


                // Clear results and hide download buttons on type change
                 resultsContent.innerHTML = '<p>Tipo de ítem cambiado. Recalcula los materiales y metrajes totales.</p>'; // Mensaje actualizado
                 downloadOptionsDiv.classList.add('hidden');
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedTotalMetrajes = {}; // Limpia metrajes totales almacenados
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
                 lastCalculatedWorkArea = ''; // Clear stored data on type change
            });
        }


         // --- Agrega event listeners a inputs relevantes del ítem para actualizar el resumen en CADA SEGMENTO Y EL METRAJE DEL ÍTEM ---
         // Modificado para incluir el tipo de poste y los inputs de cenefa.
         // NOTA: Los listeners para los inputs de segmento (.item-width, .item-height, .item-length)
         // y los inputs de cenefa (dentro de .cenefa-inputs .item-length, .item-faces)
         // se añaden DENTRO de createMuroSegmentBlock, createCieloSegmentBlock, y createCenefaBlock respectivamente.
         // Estos listeners solo necesitan llamar a calculateAndDisplayItemMetraje para el ítem padre.

         // Los listeners aquí son para los inputs a nivel de ITEM que afectan el resumen del segmento (Muro/Cielo)
         // o el metraje del ítem (todos los tipos).
         const relevantInputs = itemBlock.querySelectorAll(
             // Inputs que afectan el resumen del segmento (Muro/Cielo) Y el metraje del ítem (Muro/Cielo/Cenefa via type change effects)
             '.item-structure-type, .item-faces, .item-cara1-panel-type, .item-cara2-panel-type, ' +
             '.item-post-spacing, .item-poste-type, .item-double-structure, .item-cielo-panel-type, .item-plenum, .item-angular-deduction'
             // Inputs de Cenefa: .item-length, .item-faces (ya cubierto por el selector principal de .item-faces)
             // .item-cielo-panel-type (ya cubierto por el selector compartido)
         );

         relevantInputs.forEach(input => {
             // Determina el tipo de evento apropiado: 'input' para campos de texto/número, 'change' para selects y checkboxes.
             const eventType = (input.tagName === 'SELECT' || input.type === 'checkbox' || input.type === 'number') ? 'change' : 'input'; // Añadir 'number' a change event type


             input.addEventListener(eventType, () => {
                 // Cuando un input relevante cambia, actualiza el resumen en TODOS los segmentos de este ítem
                 // (si aplica) y recalcula y muestra el metraje del ítem.
                 itemBlock.querySelectorAll('.muro-segment, .cielo-segment').forEach(segBlock => {
                     updateSegmentItemSummary(segBlock);
                 });
                  // Si el input que cambió es el de 'faces', también necesitamos actualizar la visibilidad de los paneles de la Cara 2 para muros.
                 if (input.classList.contains('item-faces')) {
                      updateItemInputVisibility(itemBlock); // Esto ya llama a updateSegmentItemSummary dentro si el tipo es muro.
                 }
                 // También actualiza el encabezado principal si es relevante (solo tipo y número)
                 updateItemHeaderSummary(itemBlock);

                 // --- Recalcula y muestra el metraje del ítem completo en tiempo real con cada cambio relevante ---
                 calculateAndDisplayItemMetraje(itemBlock);
                 // --- Fin Recálculo Metraje Ítem ---

             });
         });
         // --- Fin Agregación de Event Listeners para Segmentos y Metraje de Ítem ---


        // Add event listener to the new remove button
        const removeButton = itemBlock.querySelector('.remove-item-btn');
        if (removeButton) { // Asegura que el botón exista
            removeButton.addEventListener('click', () => {
                itemBlock.remove(); // Remove the block from the DOM
                // Clear results and hide download buttons after removal for immediate feedback
                 resultsContent.innerHTML = '<p>Ítem eliminado. Recalcula los materiales y metrajes totales.</p>'; // Mensaje actualizado
                 downloadOptionsDiv.classList.add('hidden'); // Hide download options
                 // Also reset stored data on item removal
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedTotalMetrajes = {}; // Limpia metrajes totales almacenados
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
                 lastCalculatedWorkArea = ''; // Clear stored data on item removal
                 // Re-evaluate if calculate button should be disabled (if no items left)
                 toggleCalculateButtonState();
            });
        }


        // Set initial visibility for the inputs in the new block (defaults to muro)
        // Esto también llama a updateSegmentItemSummary para los segmentos iniciales si el tipo es muro.
        updateItemInputVisibility(itemBlock);

        // --- Calcula y muestra el metraje inicial para el ítem recién creado (por defecto Muro con 1 segmento) ---
        calculateAndDisplayItemMetraje(itemBlock);
        // --- Fin Cálculo Metraje Inicial ---

        // Re-evaluate if calculate button should be enabled (since an item was added)
        toggleCalculateButtonState();

        return itemBlock; // Return the created element
    };

     // --- NUEVA FUNCIÓN: Calcular y Mostrar Metraje por Ítem (en tiempo real) ---
     const calculateAndDisplayItemMetraje = (itemBlock) => {
         const type = itemBlock.querySelector('.item-structure-type').value;
         const itemMetrajeDisplay = itemBlock.querySelector('.item-metraje-display');
         let metrajeValue = 0; // Usaremos esta variable para el valor calculado
         let metrajeUnit = ''; // Unidad del metraje (m² o m)

         // Limpia el contenido anterior
          if (itemMetrajeDisplay) { // Asegura que el div exista antes de intentar actualizarlo
             itemMetrajeDisplay.textContent = 'Calculando metraje...';
             itemMetrajeDisplay.style.color = 'initial'; // Reset color
          } else {
              console.warn("Metraje display div not found for item", itemBlock.dataset.itemId);
              return; // Sale si no encuentra el div
          }


         try {
             if (type === 'muro') {
                 let totalMuroMetrajeArea = 0;
                 const segmentBlocks = itemBlock.querySelectorAll('.muro-segment');
                  if (segmentBlocks.length === 0 && itemBlock.querySelector('.muro-segments:not(.hidden)')) {
                     itemMetrajeDisplay.textContent = 'Metraje: Agrega segmentos válidos.';
                     itemMetrajeDisplay.style.color = 'orange';
                     // Oculta metrajes de segmentos si no hay segmentos válidos
                     itemBlock.querySelectorAll('.segment-metraje-display').forEach(div => div.textContent = '');
                     return; // No hay segmentos para calcular
                 }
                 segmentBlocks.forEach(segBlock => {
                     const segmentWidth = parseFloat(segBlock.querySelector('.item-width').value);
                     const segmentHeight = parseFloat(segBlock.querySelector('.item-height').value);

                     // Aplica la regla: si la dimensión es menor a 1, usa 1.
                     // Si la dimensión real es 0 o inválida (NaN), el ajustado también debería ser 0 o causar un error.
                     const adjustedWidth = isNaN(segmentWidth) || segmentWidth <= 0 ? 0 : Math.max(1, segmentWidth);
                     const adjustedHeight = isNaN(segmentHeight) || segmentHeight <= 0 ? 0 : Math.max(1, segmentHeight);

                     // Calcula el área del segmento usando las dimensiones ajustadas
                      const segmentMetrajeArea = adjustedWidth * adjustedHeight;

                     totalMuroMetrajeArea += segmentMetrajeArea;

                     // Opcional: Mostrar metraje por segmento (si el div existe)
                     const segmentMetrajeDisplayDiv = segBlock.querySelector('.segment-metraje-display');
                     if (segmentMetrajeDisplayDiv) {
                          if (isNaN(segmentMetrajeArea)) {
                              segmentMetrajeDisplayDiv.textContent = 'Metraje Seg: Error';
                              segmentMetrajeDisplayDiv.style.color = 'red';
                          } else {
                               segmentMetrajeDisplayDiv.textContent = `Metraje Seg: ${segmentMetrajeArea.toFixed(2)} m²`;
                               segmentMetrajeDisplayDiv.style.color = 'inherit'; // Color por defecto
                          }
                     }
                 });
                 metrajeValue = totalMuroMetrajeArea;
                 metrajeUnit = 'm²';

             } else if (type === 'cielo') {
                 let totalCieloMetrajeArea = 0;
                 const segmentBlocks = itemBlock.querySelectorAll('.cielo-segment');
                  if (segmentBlocks.length === 0 && itemBlock.querySelector('.cielo-segments:not(.hidden)')) {
                      itemMetrajeDisplay.textContent = 'Metraje: Agrega segmentos válidos.';
                      itemMetrajeDisplay.style.color = 'orange';
                       // Oculta metrajes de segmentos si no hay segmentos válidos
                      itemBlock.querySelectorAll('.segment-metraje-display').forEach(div => div.textContent = '');
                      return; // No hay segmentos para calcular
                  }
                 segmentBlocks.forEach(segBlock => {
                     const segmentWidth = parseFloat(segBlock.querySelector('.item-width').value); // Ancho para cielo
                     const segmentLength = parseFloat(segBlock.querySelector('.item-length').value); // Largo para cielo

                     // Aplica la regla: si la dimensión es menor a 1, usa 1.
                      const adjustedWidth = isNaN(segmentWidth) || segmentWidth <= 0 ? 0 : Math.max(1, segmentWidth);
                      const adjustedLength = isNaN(segmentLength) || segmentLength <= 0 ? 0 : Math.max(1, segmentLength);

                      // Calcula el área del segmento usando las dimensiones ajustadas
                      const segmentMetrajeArea = adjustedWidth * adjustedLength;

                     totalCieloMetrajeArea += segmentMetrajeArea;

                      // Opcional: Mostrar metraje por segmento (si el div existe)
                     const segmentMetrajeDisplayDiv = segBlock.querySelector('.segment-metraje-display');
                     if (segmentMetrajeDisplayDiv) {
                         if (isNaN(segmentMetrajeArea)) {
                              segmentMetrajeDisplayDiv.textContent = 'Metraje Seg: Error';
                              segmentMetrajeDisplayDiv.style.color = 'red';
                         } else {
                              segmentMetrajeDisplayDiv.textContent = `Metraje Seg: ${segmentMetrajeArea.toFixed(2)} m²`;
                              segmentMetrajeDisplayDiv.style.color = 'inherit'; // Color por defecto
                         }
                     }
                 });
                 metrajeValue = totalCieloMetrajeArea;
                 metrajeUnit = 'm²';

             } else if (type === 'cenefa') { // --- NUEVO: Cálculo de metraje para Cenefa ---
                 const cenefaInputsContainer = itemBlock.querySelector('.cenefa-inputs');
                 const lengthInput = cenefaInputsContainer ? cenefaInputsContainer.querySelector('.item-length') : null;
                 const facesInput = itemBlock.querySelector('.item-faces'); // Usa la clase compartida .item-faces

                 const cenefaLength = parseFloat(lengthInput ? lengthInput.value : NaN);
                 const cenefaFaces = parseInt(facesInput ? facesInput.value : NaN);

                  if (isNaN(cenefaLength) || cenefaLength <= 0 || isNaN(cenefaFaces) || cenefaFaces <= 0) {
                      itemMetrajeDisplay.textContent = 'Metraje: Ingresa Largo y Nº Caras válidos (> 0).';
                      itemMetrajeDisplay.style.color = 'orange';
                       // Oculta los divs de metraje por segmento ya que cenefa no tiene segmentos
                      itemBlock.querySelectorAll('.segment-metraje-display').forEach(div => div.textContent = '');
                      return; // Inputs inválidos
                  }

                  // Aplica la regla: si el largo es menor a 1, usa 1. Las caras no se ajustan con esta regla.
                  const adjustedLength = Math.max(1, cenefaLength);

                  metrajeValue = adjustedLength * cenefaFaces;
                  metrajeUnit = 'm'; // Metraje lineal para cenefa

                 // Oculta los divs de metraje por segmento ya que cenefa no tiene segmentos
                 itemBlock.querySelectorAll('.segment-metraje-display').forEach(div => div.textContent = '');


             } else {
                 itemMetrajeDisplay.textContent = 'Metraje: Tipo desconocido.';
                 itemMetrajeDisplay.style.color = 'red';
                  // Oculta metrajes de segmentos
                 itemBlock.querySelectorAll('.segment-metraje-display').forEach(div => div.textContent = '');
                 return; // Tipo desconocido
             }

             // Muestra el metraje calculado para el ítem
             if (!isNaN(metrajeValue)) {
                  itemMetrajeDisplay.textContent = `Metraje Total Ítem: ${metrajeValue.toFixed(2)} ${metrajeUnit}`;
                  itemMetrajeDisplay.style.color = 'inherit'; // Reset color to default
             } else {
                  itemMetrajeDisplay.textContent = 'Metraje: Error en cálculo.'; // Debería ser manejado por el catch, pero seguridad
                  itemMetrajeDisplay.style.color = 'red';
                   // Oculta metrajes de segmentos
                 itemBlock.querySelectorAll('.segment-metraje-display').forEach(div => div.textContent = '');
             }


         } catch (error) {
             console.error(`Error calculando metraje para ítem ${itemBlock.dataset.itemId}:`, error);
             itemMetrajeDisplay.textContent = 'Metraje: Error en cálculo.';
             itemMetrajeDisplay.style.color = 'red';
              // Oculta metrajes de segmentos
             itemBlock.querySelectorAll('.segment-metraje-display').forEach(div => div.textContent = '');
         }
     };
     // --- Fin NUEVA FUNCIÓN: Calcular y Mostrar Metraje por Ítem ---


    // --- Main Calculation Function for ALL Items ---
    // Modificada para incluir cálculo y acumulación de metrajes totales.
    const calculateMaterials = () => {
        console.log("Iniciando cálculo de materiales y metrajes...");
        const itemBlocks = itemsContainer.querySelectorAll('.item-block');

        // --- Lee el valor del nuevo input de Área de Trabajo ---
        const workAreaInput = document.getElementById('work-area');
        const workArea = workAreaInput ? workAreaInput.value.trim() : ''; // Lee el valor y quita espacios al inicio/fin
        console.log(`Área de Trabajo: "${workArea}"`);
        // --- Fin Lectura ---


        // --- Accumulators for Panels (per panel type) based on Image Logic ---
        let panelAccumulators = {};
         PANEL_TYPES.forEach(type => {
            panelAccumulators[type] = {
                suma_fraccionaria_pequenas: 0.0,
                suma_redondeada_otros: 0
            };
        });
         console.log("Acumuladores de paneles inicializados:", panelAccumulators);

        // --- Accumulator for ALL other materials (rounded per item and summed) ---
        let otherMaterialsTotal = {};

        // --- NUEVOS Acumuladores para Metrajes Totales (NUEVO) ---
        let totalMuroMetrajeAreaSum = 0;
        let totalCieloMetrajeAreaSum = 0;
        let totalCenefaMetrajeLinearSum = 0;
        // --- Fin NUEVOS Acumuladores ---


        let currentCalculatedItemsSpecs = []; // Array to store specs of validly calculated items
        let currentErrorMessages = []; // Use an array to collect validation error messages

        // --- Almacena el valor del Área de Trabajo con los resultados actuales ---
        let currentCalculatedWorkArea = workArea;

        // Clear previous results and hide download buttons initially
        resultsContent.innerHTML = '';
        downloadOptionsDiv.classList.add('hidden');

        if (itemBlocks.length === 0) {
            console.log("No hay ítems para calcular.");
            resultsContent.innerHTML = '<p style="color: orange; text-align: center; font-style: italic;">Por favor, agrega al menos un Muro, Cielo o Cenefa para calcular.</p>'; // Actualizado el texto
             // Store empty results
             lastCalculatedTotalMaterials = {};
             lastCalculatedTotalMetrajes = {}; // Limpia metrajes totales almacenados
             lastCalculatedItemsSpecs = [];
             lastErrorMessages = ['No hay ítems agregados para calcular.'];
             lastCalculatedWorkArea = ''; // Limpia el área de trabajo almacenada si no hay cálculo
            return;
        }

        console.log(`Procesando ${itemBlocks.length} ítems.`);
        // Iterate through each item block and calculate its materials
        itemBlocks.forEach(itemBlock => {
             // --- Manejo de Errores a Nivel de Ítem ---
             try {
                const itemNumber = itemBlock.dataset.itemId.split('-')[1]; // Get item number from item ID

                const type = itemBlock.querySelector('.item-structure-type').value;
                const itemId = itemBlock.dataset.itemId; // Get the unique item ID

                console.log(`Procesando Ítem #${itemNumber} (ID: ${itemId}): Tipo=${type}`);

                 // Basic Validation for Each Item
                 let itemSpecificErrors = [];
                 let itemValidatedSpecs = { // Store specs for all segments (valid or not for materials) and item details
                     id: itemId,
                     number: parseInt(itemNumber), // Store number as integer
                     type: type,
                     segments: [] // Array to store ALL segments (muro or cielo) with their details and metraje
                     // Specs type-specific added later
                 };

                 // --- Variables para almacenar el metraje calculado para ESTE ítem (NUEVO - se calculará explícitamente después) ---
                 let itemMetrajeArea = 0; // Para Muros y Cielos
                 let itemMetrajeLinear = 0; // Para Cenefas
                 // --- Fin Variables Metraje Ítem ---


                // Object to hold calculated *other* materials for THIS single item (initial floats)
                // Initialize finishing based on the primary panel type for the item (Cara 1 for Muro, Cielo type for Cielo, Cenefa type for Cenefa)
                let itemOtherMaterialsFloat = getFinishingMaterials(type === 'muro' ? (itemBlock.querySelector('.item-cara1-panel-type')?.value || null) :
                                                                      type === 'cielo' ? (itemBlock.querySelector('.item-cielo-panel-type')?.value || null) :
                                                                      type === 'cenefa' ? (itemBlock.querySelector('.cenefa-inputs .item-cielo-panel-type')?.value || null) : null);


                // --- Calculation Logic for the CURRENT Item ---

                if (type === 'muro') {
                    // Get muro-specific values
                    const facesInput = itemBlock.querySelector('.item-faces');
                    const faces = facesInput && !facesInput.closest('.hidden') ? parseInt(facesInput.value) : NaN; // Read faces here
                    const postSpacingInput = itemBlock.querySelector('.item-post-spacing');
                    const postSpacing = postSpacingInput && !postSpacingInput.closest('.hidden') ? parseFloat(postSpacingInput.value) : NaN; // Read spacing here
                     const postTypeSelect = itemBlock.querySelector('.item-poste-type'); // --- LEE el tipo de poste (NUEVO) ---
                     const postType = postTypeSelect && !postTypeSelect.closest('.hidden') ? postTypeSelect.value : null; // --- Obtiene el valor (NUEVO) ---

                    const isDoubleStructureInput = itemBlock.querySelector('.item-double-structure');
                    const isDoubleStructure = isDoubleStructureInput && !isDoubleStructureInput.closest('.hidden') ? isDoubleStructureInput.checked : false;

                     const cara1PanelTypeSelect = itemBlock.querySelector('.item-cara1-panel-type');
                     const cara1PanelType = cara1PanelTypeSelect && !cara1PanelTypeSelect.closest('.hidden') ? cara1PanelTypeSelect.value : null;

                     const cara2PanelTypeSelect = itemBlock.querySelector('.item-cara2-panel-type');
                     // Only read if faces is 2, selector is visible, and the value is not null/empty
                     const cara2PanelType = (faces === 2 && cara2PanelTypeSelect && !cara2PanelTypeSelect.closest('.hidden') && cara2PanelTypeSelect.value) ? cara2PanelTypeSelect.value : null;


                    const segmentBlocks = itemBlock.querySelectorAll('.muro-segment');
                    let totalMuroAreaForPanelsFinishing = 0; // Area using actual dimensions for materials/finishing (sum of valid segments)
                    let totalMuroWidthForStructure = 0; // Width using actual dimensions for structure (sum of valid segments)
                    let hasValidSegmentForMaterials = false; // Flag to check if at least one segment is valid FOR MATERIAL calculation

                     itemValidatedSpecs.segments = []; // Reset segments array for this item

                     if (segmentBlocks.length === 0) {
                         itemSpecificErrors.push('Muro debe tener al menos un segmento de medida.');
                     } else {
                         segmentBlocks.forEach((segBlock, index) => {
                             const segmentWidth = parseFloat(segBlock.querySelector('.item-width').value); // Use .item-width class
                             const segmentHeight = parseFloat(segBlock.querySelector('.item-height').value); // Use .item-height class
                             const segmentNumber = index + 1;

                             const isSegmentValidForMaterials = !isNaN(segmentWidth) && segmentWidth > 0 && !isNaN(segmentHeight) && segmentHeight > 0;

                             // If segment dimensions are valid for material calculation
                             if (isSegmentValidForMaterials) {
                                  hasValidSegmentForMaterials = true; // Mark as valid if at least one segment is valid
                                  const segmentArea = segmentWidth * segmentHeight;
                                  totalMuroAreaForPanelsFinishing += segmentArea; // Sum area for panels/finishing (using actual dims)
                                  totalMuroWidthForStructure += segmentWidth; // Sum width for structure (using actual dims)

                                   // --- Panel Calculation for THIS Muro Segment (using Image Logic) ---
                                   // Panel calculation is based on the area of *each segment* (using actual area)
                                   const panelTypeFace1 = cara1PanelType;
                                   const panelesFloatFace1 = segmentArea / PANEL_RENDIMIENTO_M2;

                                   if (segmentArea < SMALL_AREA_THRESHOLD_M2 && segmentArea > 0) {
                                       panelAccumulators[panelTypeFace1].suma_fraccionaria_pequenas += panelesFloatFace1;
                                       console.log(`Muro #${itemNumber} Cara 1 Segmento ${segmentNumber} (${panelTypeFace1}): Área pequeña (${segmentArea.toFixed(2)}m2). Sumando fraccional (${panelesFloatFace1.toFixed(2)}) a acumulador.`);
                                   } else if (segmentArea >= SMALL_AREA_THRESHOLD_M2) {
                                       const panelesRoundedFace1 = roundUpFinalUnit(panelesFloatFace1);
                                       panelAccumulators[panelTypeFace1].suma_redondeada_otros += panelesRoundedFace1;
                                        console.log(`Muro #${itemNumber} Cara 1 Segmento ${segmentNumber} (${panelTypeFace1}): Área grande (${segmentArea.toFixed(2)}m2). Sumando redondeado (${panelesRoundedFace1}) a acumulador.`);
                                   }


                                   if (faces === 2 && cara2PanelType) { // Only calculate for Face 2 if 2 faces are selected and type exists
                                       const panelTypeFace2 = cara2PanelType;
                                       const panelesFloatFace2 = segmentArea / PANEL_RENDIMIENTO_M2;

                                       if (segmentArea < SMALL_AREA_THRESHOLD_M2 && segmentArea > 0) {
                                           panelAccumulators[panelTypeFace2].suma_fraccionaria_pequenas += panelesFloatFace2;
                                           console.log(`Muro #${itemNumber} Cara 2 Segmento ${segmentNumber} (${panelTypeFace2}): Área pequeña (${segmentArea.toFixed(2)}m2). Sumando fraccional (${panelesFloatFace2.toFixed(2)}) a acumulador.`);
                                       } else if (segmentArea >= SMALL_AREA_THRESHOLD_M2) {
                                           const panelesRoundedFace2 = roundUpFinalUnit(panelesFloatFace2);
                                           panelAccumulators[panelTypeFace2].suma_redondeada_otros += panelesRoundedFace2;
                                           console.log(`Muro #${itemNumber} Cara 2 Segmento ${segmentNumber} (${panelTypeFace2}): Área grande (${segmentArea.toFixed(2)}m2). Sumando redondeado (${panelesRoundedFace2}) a acumulador.`);
                                       }
                                   }
                                   // Note: Segment details for itemValidatedSpecs are pushed outside this if block
                                   // to include ALL segments, valid or not for materials.

                               } else {
                                   // Add validation error for invalid segment dimensions for materials
                                   itemSpecificErrors.push(`Segmento ${segmentNumber}: Dimensiones inválidas para cálculo de materiales (Ancho y Alto deben ser > 0)`); // Mensaje más específico
                               }


                                // --- Cálculo de Metraje para ESTE Segmento de Muro (NUEVO) ---
                               // Aplica la regla de "menor a 1 se convierte en 1" para el metraje.
                               // Calcula el metraje para CADA segmento ingresado, independientemente de si es válido para materiales.
                               const metrajeWidth = isNaN(segmentWidth) || segmentWidth <= 0 ? 0 : Math.max(1, segmentWidth);
                               const metrajeHeight = isNaN(segmentHeight) || segmentHeight <= 0 ? 0 : Math.max(1, segmentHeight);
                               const segmentMetrajeArea = metrajeWidth * metrajeHeight;
                               // No acumulamos aquí totalMuroMetrajeAreaForMetraje; se hará explícitamente después.

                               // Store segment specs for report, regardless of material validity
                               itemValidatedSpecs.segments.push({
                                   number: segmentNumber,
                                   width: isNaN(segmentWidth) ? 0 : segmentWidth, // Store raw input, use 0 if NaN for consistency
                                   height: isNaN(segmentHeight) ? 0 : segmentHeight, // Store raw input
                                   area: isSegmentValidForMaterials ? segmentWidth * segmentHeight : 0, // Store actual area only if valid for materials
                                   metrajeArea: segmentMetrajeArea, // Store metraje area for this segment
                                   isValidForMaterials: isSegmentValidForMaterials // Flag if segment is valid for material calculation
                               });


                         }); // End of segmentBlocks.forEach

                         // Add validation specific to Muros AFTER processing segments
                         if (!hasValidSegmentForMaterials && itemBlock.querySelectorAll('.muro-segment').length > 0) {
                             itemSpecificErrors.push('Muro debe tener al menos un segmento de medida válido (> 0 en Ancho y Alto) para calcular materiales.'); // Mensaje más específico
                         }
                         if (isNaN(faces) || (faces !== 1 && faces !== 2)) itemSpecificErrors.push('Nº Caras inválido (debe ser 1 o 2)');
                         if (isNaN(postSpacing) || postSpacing <= 0) itemSpecificErrors.push('Espaciamiento Postes inválido (debe ser > 0)');
                         if (!cara1PanelType || !PANEL_TYPES.includes(cara1PanelType)) itemSpecificErrors.push('Tipo de Panel Cara 1 inválido.');
                         // Check cara2PanelType only if faces is 2
                         if (faces === 2 && (!cara2PanelType || !PANEL_TYPES.includes(cara2PanelType))) itemSpecificErrors.push('Tipo de Panel Cara 2 inválido para 2 caras.');
                         // --- Validación para Tipo de Poste (NUEVO) ---
                         if (!postType || !POSTE_TYPES.includes(postType)) itemSpecificErrors.push('Tipo de Poste inválido.');
                         // --- Fin Validación Tipo de Poste ---

                         // Store total calculated values for muros (actual dimensions used for materials)
                         itemValidatedSpecs.faces = faces; // Store faces
                         itemValidatedSpecs.cara1PanelType = cara1PanelType; // Store cara1 panel type
                         itemValidatedSpecs.cara2PanelType = cara2PanelType; // Store cara2 panel type
                         itemValidatedSpecs.postSpacing = postSpacing; // Store post spacing
                         itemValidatedSpecs.postType = postType; // --- Store poste type (NUEVO) ---
                         itemValidatedSpecs.isDoubleStructure = isDoubleStructure; // Store double structure
                         // totalMuroArea and totalMuroWidth are already calculated from valid segments above
                         itemValidatedSpecs.totalMuroArea = totalMuroAreaForPanelsFinishing;
                         itemValidatedSpecs.totalMuroWidth = totalMuroWidthForStructure;

                         // --- Calculate Item Metraje Area for Muro (NUEVO - Explicitly) ---
                         // Suma el metraje de CADA segmento ingresado para obtener el metraje total del ÍTEM (usando la regla menor a 1 = 1)
                         let itemMuroMetrajeAreaCalc = 0;
                         itemValidatedSpecs.segments.forEach(seg => { // Iterate over all segments stored in specs
                              itemMuroMetrajeAreaCalc += seg.metrajeArea; // Sum their metraje areas
                         });
                         itemValidatedSpecs.metrajeArea = itemMuroMetrajeAreaCalc; // Store the calculated metraje area for the item
                         itemMetrajeArea = itemMuroMetrajeAreaCalc; // Assign to item-level variable for global sum
                         // --- Fin Calculate Item Metraje Area ---


                     } // End if segmentBlocks.length > 0


                    // --- Other Materials Calculation for THIS Muro Item (Structure, Finishing, Screws) ---
                    // Only proceed with material calculation if there are no validation errors *so far* for this item and at least one valid segment for materials
                    if (itemSpecificErrors.length === 0 && hasValidSegmentForMaterials) { // Use hasValidSegmentForMaterials flag
                        // Structure calculation is based on the *total accumulated width* of all *valid* segments (using actual width)
                        const totalMuroWidthForStructureCalc = itemValidatedSpecs.totalMuroWidth; // Use the sum of valid widths
                        const totalMuroAreaForPanelsFinishingCalc = itemValidatedSpecs.totalMuroArea; // Use the sum of valid areas

                        // Postes (based on total valid width and SELECTED POSTE TYPE)
                        let postesFloat;
                        if (totalMuroWidthForStructureCalc > 0 && postSpacing > 0) {
                           // --- Lógica de cálculo de postes usando la fórmula original ---
                           postesFloat = (totalMuroWidthForStructureCalc / postSpacing) + 1; // Calculate number of intervals and add 1
                           // --- FIN Lógica Original ---
                        } else {
                           postesFloat = 0;
                        }
                        if (isDoubleStructure) postesFloat *= 2;
                         // --- Store Postes quantity under the specific Poste Type name (NUEVO) ---
                         if (postType) { // Ensure postType was read validly
                             itemOtherMaterialsFloat[postType] = postesFloat; // Use the specific type name as the key
                             // Remove the old generic 'Postes' key if it exists from the initial finishing materials logic
                             delete itemOtherMaterialsFloat['Postes'];
                         }


                        // Canales (based on total valid width)
                        let canalesFloat;
                         if (totalMuroWidthForStructureCalc > 0) {
                           // Need 2 channels (top/bottom) along the total width. Standard length 3.05m.
                            canalesFloat = (totalMuroWidthForStructureCalc * 2) / 3.05;
                        } else {
                            canalesFloat = 0;
                        }
                        if (isDoubleStructure) canalesFloat *= 2;
                        // Determine Canales type based on panel type of Cara 1
                         itemOtherMaterialsFloat[cara1PanelType === 'Exterior' ? 'Canales Calibre 20' : 'Canales'] = canalesFloat;


                         // Acabado y Tornillos de Panel (based on total accumulated area of ALL *valid* segments)
                        const primaryPanelTypeForFinishing = cara1PanelType; // Use Cara 1 type for finishing logic

                        if (primaryPanelTypeForFinishing === 'Normal' || primaryPanelTypeForFinishing === 'Resistente a la Humedad' || primaryPanelTypeForFinishing === 'Resistente al Fuego' || primaryPanelTypeForFinishing === 'Alta Resistencia') {
                            // Pasta (Yeso) - Calculation based on Total Accumulated Area
                            itemOtherMaterialsFloat['Pasta'] = totalMuroAreaForPanelsFinishingCalc > 0 ? totalMuroAreaForPanelsFinishingCalc / 22 : 0; // Area / Rendimiento (22 m2/caja)

                            // Cinta de Papel - Calculation based on Total Accumulated Area
                            itemOtherMaterialsFloat['Cinta de Papel'] = totalMuroAreaForPanelsFinishingCalc > 0 ? totalMuroAreaForPanelsFinishingCalc * 1 : 0; // 1 meter per m2

                            // Lija Grano 120 - Calculation based on Total Accumulated Area
                            itemOtherMaterialsFloat['Lija Grano 120'] = totalMuroAreaForPanelsFinishingCalc > 0 ? (totalMuroAreaForPanelsFinishingCalc / PANEL_RENDIMIENTO_M2) / 2 : 0;


                            // Tornillos de 1" punta fina (for Yeso type panels on this item)
                            itemOtherMaterialsFloat['Tornillos de 1" punta fina'] = totalMuroAreaForPanelsFinishingCalc > 0 ? (totalMuroAreaForPanelsFinishingCalc / PANEL_RENDIMIENTO_M2) * 40 : 0; // 40 screws per panel (approx)


                        } else if (primaryPanelTypeForFinishing === 'Exterior') { // Assuming 'Exterior' implies Durock or similar
                             // Basecoat (Durock) - Calculation based on Total Accumulated Area
                            itemOtherMaterialsFloat['Basecoat'] = totalMuroAreaForPanelsFinishingCalc > 0 ? totalMuroAreaForPanelsFinishingCalc / 8 : 0; // Area / Rendimiento (8 m2/saco)

                            // Cinta malla (Durock) - Calculation based on Total Accumulated Area
                             itemOtherMaterialsFloat['Cinta malla'] = totalMuroAreaForPanelsFinishingCalc > 0 ? totalMuroAreaForPanelsFinishingCalc * 1 : 0; // 1 meter per m2

                            // Tornillos de 1" punta broca (for Durock type panels on this item)
                             itemOtherMaterialsFloat['Tornillos de 1" punta broca'] = totalMuroAreaForPanelsFinishingCalc > 0 ? (totalMuroAreaForPanelsFinishingCalc / PANEL_RENDIMIENTO_M2) * 40 : 0; // 40 screws per panel (approx)
                        }


                         // --- Calculate Tornillos/Clavos Based on Rounded Component Counts for THIS Item ---
                         // These are calculated per item based on rounded structure counts for THIS item.
                         // Only calculate if there is width from valid segments
                        if (totalMuroWidthForStructureCalc > 0) {
                            // Use the specific poste type name to get its calculated quantity
                            let roundedPostes = roundUpFinalUnit(itemOtherMaterialsFloat[postType] || 0); // Use the specific type name
                            let roundedCanales = roundUpFinalUnit(itemOtherMaterialsFloat[cara1PanelType === 'Exterior' ? 'Canales Calibre 20' : 'Canales'] || 0); // Use the specific type name

                            // Clavos con Roldana (8 per Canal) - Use rounded Canales for this item
                             itemOtherMaterialsFloat['Clavos con Roldana'] = roundedCanales * 8;
                             itemOtherMaterialsFloat['Fulminantes'] = itemOtherMaterialsFloat['Clavos con Roldana']; // Igual cantidad

                             // Tornillos 1/2" (4 per Poste) - Use rounded Postes for this item
                             // Determine screw type based on panel type of Cara 1
                            itemOtherMaterialsFloat[cara1PanelType === 'Exterior' ? 'Tornillos de 1/2" punta broca' : 'Tornillos de 1/2" punta fina'] = roundedPostes * 4;
                        }


                    } // End if itemSpecificErrors.length === 0 && hasValidSegmentForMaterials


                } else if (type === 'cielo') {
                     // Get cielo-specific values
                    const plenumInput = itemBlock.querySelector('.item-plenum');
                    const plenum = plenumInput && !plenumInput.closest('.hidden') ? parseFloat(plenumInput.value) : NaN;
                    const angularDeductionInput = itemBlock.querySelector('.item-angular-deduction');
                    const angularDeduction = angularDeductionInput && !angularDeductionInput.closest('.hidden') ? parseFloat(angularDeductionInput.value) : NaN;
                    const cieloPanelTypeSelect = itemBlock.querySelector('.item-cielo-panel-type');
                    const cieloPanelType = cieloPanelTypeSelect && !cieloPanelTypeSelect.closest('.hidden') ? cieloPanelTypeSelect.value : null;


                     const segmentBlocks = itemBlock.querySelectorAll('.cielo-segment');
                     let totalCieloAreaForPanelsFinishing = 0; // Total area from all segments (using actual dims) (sum of valid segments)
                     let totalCieloPerimeterForAngular = 0; // Sum of FULL perimeters (using actual dims) for angular calculation (sum of valid segments)
                     let hasValidSegmentForMaterials = false; // Flag to check if at least one segment is valid FOR MATERIAL calculation


                     itemValidatedSpecs.segments = []; // Reset segments array for this item

                     if (segmentBlocks.length === 0) {
                         itemSpecificErrors.push('Cielo Falso debe tener al menos un segmento de medida.');
                     } else {
                         segmentBlocks.forEach((segBlock, index) => {
                              const segmentWidth = parseFloat(segBlock.querySelector('.item-width').value); // Use .item-width class
                              const segmentLength = parseFloat(segBlock.querySelector('.item-length').value); // Use .item-length class
                              const segmentNumber = index + 1;

                              const isSegmentValidForMaterials = !isNaN(segmentWidth) && segmentWidth > 0 && !isNaN(segmentLength) && segmentLength > 0;


                              // If segment dimensions are valid for material calculation
                              if (isSegmentValidForMaterials) {
                                   hasValidSegmentForMaterials = true; // Mark as valid if at least one segment is valid
                                   const segmentArea = segmentWidth * segmentLength;
                                   totalCieloAreaForPanelsFinishing += segmentArea; // Sum area (using actual dims)
                                   // --- CÁLCULO DEL PERÍMETRO COMPLETO POR SEGMENTO (usando actual dims) ---
                                   totalCieloPerimeterForAngular += 2 * (segmentWidth + segmentLength); // Sum of full perimeters for angular calculation
                                   // --- FIN CÁLCULO ---

                                   // --- Panel Calculation for THIS Cielo Segment (using Image Logic) ---
                                   // Panel calculation is based on the area of *each segment* (using actual area)
                                   const panelTypeCielo = cieloPanelType;
                                   const panelesFloatCielo = segmentArea / PANEL_RENDIMIENTO_M2;

                                   if (segmentArea < SMALL_AREA_THRESHOLD_M2 && segmentArea > 0) {
                                       panelAccumulators[panelTypeCielo].suma_fraccionaria_pequenas += panelesFloatCielo;
                                       console.log(`Cielo #${itemNumber} Segmento ${segmentNumber} (${panelTypeCielo}): Área pequeña (${segmentArea.toFixed(2)}m2). Sumando fraccional (${panelesFloatCielo.toFixed(2)}) a acumulador.`);
                                   } else if (segmentArea >= SMALL_AREA_THRESHOLD_M2) {
                                        const panelesRoundedCielo = roundUpFinalUnit(panelesFloatCielo);
                                        panelAccumulators[panelTypeCielo].suma_redondeada_otros += panelesRoundedCielo;
                                        console.log(`Cielo #${itemNumber} Segmento ${segmentNumber} (${panelTypeCielo}): Área grande (${segmentArea.toFixed(2)}m2). Sumando redondeado (${panelesRoundedCielo}) a acumulador.`);
                                   }
                                   // Note: Segment details for itemValidatedSpecs are pushed outside this if block
                                   // to include ALL segments, valid or not for materials.

                              } else {
                                  itemSpecificErrors.push(`Segmento ${segmentNumber}: Dimensiones inválidas para cálculo de materiales (Ancho y Largo deben ser > 0)`); // Mensaje más específico
                              }


                               // --- Cálculo de Metraje para ESTE Segmento de Cielo (NUEVO) ---
                               // Aplica la regla de "menor a 1 se convierte en 1" para el metraje.
                               // Calcula el metraje para CADA segmento ingresado, independientemente de si es válido para materiales.
                               const metrajeWidth = isNaN(segmentWidth) || segmentWidth <= 0 ? 0 : Math.max(1, segmentWidth);
                               const metrajeLength = isNaN(segmentLength) || segmentLength <= 0 ? 0 : Math.max(1, segmentLength);
                               const segmentMetrajeArea = metrajeWidth * metrajeLength;
                               // No acumulamos aquí totalCieloMetrajeAreaForMetraje; se hará explícitamente después.


                              // Store segment specs for report, regardless of material validity
                              itemValidatedSpecs.segments.push({
                                  number: segmentNumber,
                                  width: isNaN(segmentWidth) ? 0 : segmentWidth, // Store raw input
                                  length: isNaN(segmentLength) ? 0 : segmentLength, // Store raw input
                                  area: isSegmentValidForMaterials ? segmentWidth * segmentLength : 0, // Store actual area only if valid for materials
                                  metrajeArea: segmentMetrajeArea, // Store metraje area for this segment
                                  isValidForMaterials: isSegmentValidForMaterials // Flag if segment is valid for material calculation
                              });


                         }); // End of segmentBlocks.forEach

                         // Add validation specific to Cielos AFTER processing segments
                         if (!hasValidSegmentForMaterials && itemBlock.querySelectorAll('.cielo-segment').length > 0) {
                              itemSpecificErrors.push('Cielo Falso debe tener al menos un segmento de medida válido (> 0 en Ancho y Largo) para calcular materiales.'); // Mensaje actualizado
                         }
                          // Plenum validation only if visible and required (check plenum input existence and value)
                         const plenumInput = itemBlock.querySelector('.item-plenum'); // Re-get input as it might be hidden
                          if (itemBlock.querySelector('.item-plenum-input') && !itemBlock.querySelector('.item-plenum-input').classList.contains('hidden') && (isNaN(plenum) || plenum < 0)) {
                              itemSpecificErrors.push('Pleno inválido (debe ser >= 0)');
                          }
                         if (!cieloPanelType || !PANEL_TYPES.includes(cieloPanelType)) itemSpecificErrors.push('Tipo de Panel de Cielo inválido.');
                          // --- Agrega validación para el descuento angular ---
                          const angularDeductionInput = itemBlock.querySelector('.item-angular-deduction'); // Re-get input
                          if (itemBlock.querySelector('.item-angular-deduction-input') && !itemBlock.querySelector('.item-angular-deduction-input').classList.contains('hidden') && (isNaN(angularDeduction) || angularDeduction < 0)) {
                               itemSpecificErrors.push('Metros a descontar de Angular inválido (debe ser >= 0).');
                          }
                          // --- Fin agregación ---


                         // Store validated specs for cielos (actual dimensions used for materials)
                         itemValidatedSpecs.cieloPanelType = cieloPanelType; // Store cielo panel type
                         itemValidatedSpecs.plenum = plenum; // Store plenum
                         itemValidatedSpecs.angularDeduction = angularDeduction; // Store the parsed deduction value
                         // totalCieloArea and totalCieloPerimeterSum are calculated from valid segments above
                         itemValidatedSpecs.totalCieloArea = totalCieloAreaForPanelsFinishing;
                         itemValidatedSpecs.totalCieloPerimeterSum = totalCieloPerimeterForAngular;


                         // --- Calculate Item Metraje Area for Cielo (NUEVO - Explicitly) ---
                         // Suma el metraje de CADA segmento ingresado para obtener el metraje total del ÍTEM (usando la regla menor a 1 = 1)
                         let itemCieloMetrajeAreaCalc = 0;
                         itemValidatedSpecs.segments.forEach(seg => { // Iterate over all segments stored in specs
                              itemCieloMetrajeAreaCalc += seg.metrajeArea; // Sum their metraje areas
                         });
                         itemValidatedSpecs.metrajeArea = itemCieloMetrajeAreaCalc; // Store the calculated metraje area for the item
                         itemMetrajeArea = itemCieloMetrajeAreaCalc; // Assign to item-level variable for global sum
                         // --- Fin Calculate Item Metraje Area ---


                     } // End if segmentBlocks.length > 0 for cielo


                     // --- Other Materials Calculation for THIS Cielo Item (Structure, Finishing, Screws) ---
                     // Only proceed with material calculation if there are no validation errors *so far* for this item and at least one valid segment for materials
                    if (itemSpecificErrors.length === 0 && hasValidSegmentForMaterials) { // Use hasValidSegmentForMaterials flag
                        const totalCieloAreaForPanelsFinishingCalc = itemValidatedSpecs.totalCieloArea; // Use sum of valid areas
                        const totalCieloPerimeterForAngularCalc = itemValidatedSpecs.totalCieloPerimeterSum; // Use sum of valid perimeters

                         // Canal Listón (based on total valid area)
                         let canalListonFloat = 0;
                         if (totalCieloAreaForPanelsFinishingCalc > 0) {
                             // Formula seems to be: total area / spacing (0.40) / length of piece (3.66)
                             canalListonFloat = (totalCieloAreaForPanelsFinishingCalc / 0.40) / 3.66;
                         }
                         itemOtherMaterialsFloat['Canal Listón'] = canalListonFloat;


                         // Canal Soporte (based on total valid area)
                         let canalSoporteFloat = 0;
                         if (totalCieloAreaForPanelsFinishingCalc > 0) {
                             // Formula seems to be: total area / spacing (0.90) / length of piece (3.66)
                            canalSoporteFloat = (totalCieloAreaForPanelsFinishingCalc / 0.90) / 3.66;
                         }
                         itemOtherMaterialsFloat['Canal Soporte'] = canalSoporteFloat;


                         // Angular de Lámina (based on sum of full perimeters of segments MINUS deduction - using total valid perimeter sum)
                         let angularLaminaFloat = 0;
                         if (totalCieloPerimeterForAngularCalc > 0) {
                             // --- Aplica el descuento antes de dividir ---
                             let adjustedPerimeter = totalCieloPerimeterForAngularCalc - (isNaN(angularDeduction) ? 0 : angularDeduction); // Subtract deduction, default to 0 if NaN
                             if (adjustedPerimeter < 0) adjustedPerimeter = 0; // Ensure it doesn't go below 0
                             // --- Fin aplicación de descuento ---
                             angularLaminaFloat = adjustedPerimeter / 2.44;
                         }
                         itemOtherMaterialsFloat['Angular de Lámina'] = angularLaminaFloat;


                        // Patas (Soportes) - Calculation based on Canal Soporte quantity for THIS item (using total valid area for Canal Soporte calc)
                         // Calculate base Patas quantity first, then use its rounded value later for Cuelgue/Screws
                        let patasFloat = (itemOtherMaterialsFloat['Canal Soporte'] || 0) * 4; // 4 patas per Canal Soporte (approx)
                        itemOtherMaterialsFloat['Patas'] = patasFloat; // Store float value


                         // Canal Listón (para cuelgue) - Represents the number of 3.66m profiles needed to cut the hangers
                         // Calculation based on rounded Patas quantity and Plenum for THIS item
                         let roundedPatasForCuelgue = roundUpFinalUnit(patasFloat || 0); // Round patas float value
                         if (plenum > 0 && roundedPatasForCuelgue > 0 && !isNaN(plenum)) {
                            // Number of cuelgues needed = roundedPatasForCuelgue. Each cuelgue has length = plenum. Total length = roundedPatasForCuelgue * plenum. Convert to 3.66m pieces.
                            itemOtherMaterialsFloat['Canal Listón (para cuelgue)'] = (roundedPatasForCuelgue * plenum) / 3.66;
                         } else {
                             itemOtherMaterialsFloat['Canal Listón (para cuelgue)'] = 0;
                         }


                         // Tornillos 1" punta broca (for Cielo panels - assuming Yeso type uses broca for structure attachment to ceiling)
                         // Base this on the estimated *total* panel count for the item's total valid area.
                         itemOtherMaterialsFloat['Tornillos de 1" punta broca'] = totalCieloAreaForPanelsFinishingCalc > 0 ? (totalCieloAreaForPanelsFinishingCalc / PANEL_RENDIMIENTO_M2) * 40 : 0; // 40 screws per panel (approx)


                         // Acabado (Pasta, Cinta de Papel, Lija / Basecoat, Cinta malla) - Same calculation as Muros, based on total accumulated valid panel area
                         const primaryPanelTypeForFinishing = cieloPanelType; // Use Cielo type for finishing logic

                         if (primaryPanelTypeForFinishing === 'Normal' || primaryPanelTypeForFinishing === 'Resistente a la Humedad' || primaryPanelTypeForFininshing === 'Resistente al Fuego' || primaryPanelTypeForFinishing === 'Alta Resistencia') {
                             // Pasta (Yeso) - Calculation based on Area (total item valid area)
                             itemOtherMaterialsFloat['Pasta'] = totalCieloAreaForPanelsFinishingCalc > 0 ? totalCieloAreaForPanelsFinishingCalc / 22 : 0;

                             // Cinta de Papel - Calculation based on Area
                             itemOtherMaterialsFloat['Cinta de Papel'] = totalCieloAreaForPanelsFinishingCalc > 0 ? totalCieloAreaForPanelsFinishingCalc * 1 : 0;

                             // Lija Grano 120 - Calculation based on Area (estimated panels / 2)
                             itemOtherMaterialsFloat['Lija Grano 120'] = totalCieloAreaForPanelsFinishingCalc > 0 ? (totalCieloAreaForPanelsFinishingCalc / PANEL_RENDIMIENTO_M2) / 2 : 0;

                         } else if (primaryPanelTypeForFinishing === 'Exterior') { // Assuming 'Exterior' implies Durock or similar
                              // Basecoat (Durock) - Calculation based on Area
                             itemOtherMaterialsFloat['Basecoat'] = totalCieloAreaForPanelsFinishingCalc > 0 ? totalCieloAreaForPanelsFinishingCalc / 8 : 0;

                             // Cinta malla (Durock) - Calculation based on Area
                             itemOtherMaterialsFloat['Cinta malla'] = totalCieloAreaForPanelsFinishingCalc > 0 ? totalCieloAreaForPanelsFinishingCalc * 1 : 0;
                         }

                         // --- Calculate Tornillos/Clavos Based on Rounded Component Counts for THIS Item ---
                         // These are calculated per item based on rounded structure counts for THIS item.
                         // Only calculate if there is area/structure components
                         if (totalCieloAreaForPanelsFinishingCalc > 0 || totalCieloPerimeterForAngularCalc > 0) {
                             let roundedAngularLamina = roundUpFinalUnit(itemOtherMaterialsFloat['Angular de Lámina'] || 0); // Use the calculated value AFTER deduction
                             let roundedCanalSoporte = roundUpFinalUnit(itemOtherMaterialsFloat['Canal Soporte'] || 0);
                             let roundedCanalListon = roundUpFinalUnit(itemOtherMaterialsFloat['Canal Listón'] || 0);
                              // Patas calculation depends on Canal Soporte, round its calculated float value *for this item*
                             let roundedPatas = roundUpFinalUnit(itemOtherMaterialsFloat['Patas'] || 0); // Use the previously stored float value

                             // Clavos con Roldana (5 per Angular + 8 per Canal Soporte) - Use rounded counts for this item
                             itemOtherMaterialsFloat['Clavos con Roldana'] = (roundedAngularLamina * 5) + (roundedCanalSoporte * 8);
                             itemOtherMaterialsFloat['Fulminantes'] = itemOtherMaterialsFloat['Clavos con Roldana']; // Igual cantidad

                             // Tornillos 1/2" punta fina (12 per Canal Listón + 2 per Pata) - Use rounded counts for this item
                             itemOtherMaterialsFloat['Tornillos de 1/2" punta fina'] = (roundedCanalListon * 12) + (roundedPatas * 2); // Cielo structure uses punta fina

                         }


                     } // End if itemSpecificErrors.length === 0 && hasValidSegmentForMaterials

                } else if (type === 'cenefa') { // --- NUEVO: Cálculo para Cenefa ---
                     const cenefaInputsContainer = itemBlock.querySelector('.cenefa-inputs');
                     const lengthInput = cenefaInputsContainer ? cenefaInputsContainer.querySelector('.item-length') : null;
                     const facesInput = itemBlock.querySelector('.item-faces'); // Usa la clase compartida .item-faces
                     const panelTypeSelect = cenefaInputsContainer ? cenefaInputsContainer.querySelector('.item-cielo-panel-type') : null; // Reutiliza la clase para el selector


                    const cenefaLength = parseFloat(lengthInput ? lengthInput.value : NaN);
                    const cenefaFaces = parseInt(facesInput ? facesInput.value : NaN);
                    const cenefaPanelType = panelTypeSelect ? panelTypeSelect.value : null;

                    // Validación específica de Cenefa
                    if (isNaN(cenefaLength) || cenefaLength <= 0) itemSpecificErrors.push('Largo Total inválido (debe ser > 0).'); // Añadir validación explícita del largo
                    if (isNaN(cenefaFaces) || cenefaFaces <= 0) itemSpecificErrors.push('Nº de Caras inválido (debe ser > 0).'); // Añadir validación explícita de caras
                    if (!panelTypeSelect || !PANEL_TYPES.includes(cenefaPanelType)) itemSpecificErrors.push('Tipo de Panel de Cenefa inválido.');


                    // --- Calculate Item Metraje Linear for Cenefa (NUEVO - Explicitly) ---
                    let itemCenefaMetrajeLinearCalc = 0;
                    if (!isNaN(cenefaLength) && cenefaLength > 0 && !isNaN(cenefaFaces) && cenefaFaces > 0) {
                       // Aplica la regla: si el largo es menor a 1, usa 1. Caras no se ajustan.
                        const adjustedLengthForMetraje = Math.max(1, cenefaLength);
                       itemCenefaMetrajeLinearCalc = adjustedLengthForMetraje * cenefaFaces;
                    }
                    itemValidatedSpecs.metrajeLinear = itemCenefaMetrajeLinearCalc; // Store the calculated metraje linear
                    itemMetrajeLinear = itemCenefaMetrajeLinearCalc; // Assign to item-level variable for global sum
                    // --- Fin Calculate Item Metraje Linear ---


                    // Store validated specs for cenefas
                    itemValidatedSpecs.cenefaLength = cenefaLength; // Store actual length
                    itemValidatedSpecs.cenefaFaces = cenefaFaces; // Store actual faces
                    itemValidatedSpecs.cenefaPanelType = cenefaPanelType; // Store panel type


                    // --- Other Materials Calculation for THIS Cenefa Item ---
                    // Assuming Cenefas use Angular de Lámina for structure and finishing based on panel type.
                     // The length for Angular is the total length * number of faces.
                    // Only proceed with material calculation if there are no validation errors *so far* for this item
                    if (itemSpecificErrors.length === 0) { // Check errors before material calculation for cenefa
                         // Angular de Lámina (based on actual total length * faces)
                         const totalCenefaLinearLength = cenefaLength * cenefaFaces;
                         itemOtherMaterialsFloat['Angular de Lámina'] = totalCenefaLinearLength > 0 ? totalCenefaLinearLength / 2.44 : 0; // Divide by length of piece

                         // Acabado (Pasta, Cinta de Papel, Lija / Basecoat, Cinta malla) - Based on estimated *panel area* for this item.
                         // How to estimate panel area for a cenefa? Assuming the area is the total linear length * the height of the panel piece used.
                         // A standard panel is 1.22m wide. If the cenefa height is less than 1.22m, you use part of a panel.
                         // Let's assume for simplification the panel area for finishing is the total linear length * 0.30m (a typical cenefa height).
                         // This is an assumption and might need clarification from the user. Using a placeholder height of 0.3m.
                          const assumedCenefaHeightForFinishing = 0.30; // Assumption!
                          const estimatedCenefaAreaForFinishing = totalCenefaLinearLength * assumedCenefaHeightForFinishing;


                         const primaryPanelTypeForFinishing = cenefaPanelType; // Use Cenefa type for finishing logic

                         if (primaryPanelTypeForFinishing === 'Normal' || primaryPanelTypeForFinishing === 'Resistente a la Humedad' || primaryPanelTypeForFininshing === 'Resistente al Fuego' || primaryPanelTypeForFinishing === 'Alta Resistencia') {
                             // Pasta (Yeso) - Calculation based on Estimated Area
                             itemOtherMaterialsFloat['Pasta'] = estimatedCenefaAreaForFinishing > 0 ? estimatedCenefaAreaForFinishing / 22 : 0;

                             // Cinta de Papel - Calculation based on Estimated Area
                             itemOtherMaterialsFloat['Cinta de Papel'] = estimatedCenefaAreaForFinishing > 0 ? estimatedCenefaAreaForFinishing * 1 : 0;

                             // Lija Grano 120 - Calculation based on Estimated Area (estimated panels / 2)
                             itemOtherMaterialsFloat['Lija Grano 120'] = estimatedCenefaAreaForFinishing > 0 ? (estimatedCenefaAreaForFinishing / PANEL_RENDIMIENTO_M2) / 2 : 0;

                             // Tornillos de 1" punta fina (for Yeso type panels on this item)
                              itemOtherMaterialsFloat['Tornillos de 1" punta fina'] = estimatedCenefaAreaForFinishing > 0 ? (estimatedCenefaAreaForFinishing / PANEL_RENDIMIENTO_M2) * 40 : 0; // 40 screws per panel (approx)


                         } else if (primaryPanelTypeForFinishing === 'Exterior') { // Assuming 'Exterior' implies Durock or similar
                              // Basecoat (Durock) - Calculation based on Estimated Area
                             itemOtherMaterialsFloat['Basecoat'] = estimatedCenefaAreaForFinishing > 0 ? estimatedCenefaAreaForFinishing / 8 : 0;

                             // Cinta malla (Durock) - Calculation based on Estimated Area
                             itemOtherMaterialsFloat['Cinta malla'] = estimatedCenefaAreaForFinishing > 0 ? estimatedCenefaAreaForFinishing * 1 : 0;
                         }

                         // Tornillos para estructura (e.g., 1/2" to attach angular to ceiling/wall)
                         // Need to estimate based on Angular quantity. Assume X screws per Angular piece.
                         // Let's assume 5 tornillos 1/2" punta broca per Angular piece (similar to angular attachment in cielos).
                         let roundedAngularLamina = roundUpFinalUnit(itemOtherMaterialsFloat['Angular de Lámina'] || 0);
                         // Assuming Cenefas attached to structure (wall/ceiling) would use broca screws.
                         itemOtherMaterialsFloat['Tornillos de 1/2" punta broca'] = roundedAngularLamina * 5; // Use broca for attaching to structure

                    }


                } else {
                    // Unknown type (shouldn't happen with validation)
                     itemSpecificErrors.push('Tipo de estructura desconocido.');
                }

                console.log(`Ítem #${itemNumber}: Errores de validación - ${itemSpecificErrors.length}`);

                // If item has errors, add to global error list and skip calculation for this item
                if (itemSpecificErrors.length > 0) {
                     const errorTitle = `${getItemTypeName(type)} #${itemNumber}`;
                     currentErrorMessages.push(`Error en ${errorTitle}: ${itemSpecificErrors.join(', ')}. Revisa los inputs del ítem.`); // Mensaje más amigable
                     console.warn(`Item inválido o incompleto: ${errorTitle}. Errores: ${itemSpecificErrors.join(', ')}. Este ítem no se incluirá en el cálculo total.`);
                     // Do NOT add to currentCalculatedItemsSpecs if there are errors
                     return; // Skip calculation and summing for this invalid item
                }

                 // Store the validated specs for this item (segments already pushed for muro/cielo)
                 // Only push if there were no errors for this item
                 currentCalculatedItemsSpecs.push(itemValidatedSpecs);

                 // --- Acumular Metrajes Totales (NUEVO) ---
                 // Usa itemMetrajeArea o itemMetrajeLinear que ya fueron calculados explícitamente para el ítem validado.
                 if (type === 'muro') {
                      totalMuroMetrajeAreaSum += itemMetrajeArea;
                 } else if (type === 'cielo') {
                      totalCieloMetrajeAreaSum += itemMetrajeArea;
                 } else if (type === 'cenefa') {
                     totalCenefaMetrajeLinearSum += itemMetrajeLinear;
                 }
                 // --- Fin Acumular Metrajes Totales ---


                console.log(`Ítem #${itemNumber}: Otros materiales calculados (float) antes de redondear individualmente:`, itemOtherMaterialsFloat);

                 // --- Round Up *Other* Material Quantities for THIS Item and Sum to Total ---
                // Panels are handled separately in accumulators.
                // This block executes only for valid items (no errors)
                for (const material in itemOtherMaterialsFloat) {
                    if (itemOtherMaterialsFloat.hasOwnProperty(material)) {
                        const floatQuantity = itemOtherMaterialsFloat[material];
                        // Ensure the value is a valid number before rounding and summing
                        // Only sum positive quantities or quantities calculated based on rounded components (which could be 0 if components are 0)
                        if (!isNaN(floatQuantity)) { // Sum all valid numbers, even 0 if it results from calc
                             const roundedQuantity = roundUpFinalUnit(floatQuantity);
                             // Sum this rounded quantity from the current item to the overall total of OTHER materials
                             otherMaterialsTotal[material] = (otherMaterialsTotal[material] || 0) + roundedQuantity;
                        }
                    }
                }
                 console.log(`Ítem #${itemNumber}: Otros materiales redondeados y sumados a total. Total parcial otros materiales:`, otherMaterialsTotal);

             } catch (error) {
                 // --- Captura y reporta errores inesperados durante el procesamiento de un ítem ---
                 const itemIdentifier = itemBlock.dataset.itemId ? `#${itemBlock.dataset.itemId.split('-')[1]}` : '(ID desconocido)';
                 const itemType = itemBlock.querySelector('.item-structure-type') ? getItemTypeName(itemBlock.querySelector('.item-structure-type').value) : 'Desconocido';
                 const errorMessage = `Error inesperado procesando Ítem ${itemType} ${itemIdentifier}: ${error.message}. Revisa la consola para más detalles.`; // Mensaje más amigable
                 currentErrorMessages.push(errorMessage);
                 console.error(errorMessage, error);
                 // Continúa con el siguiente ítem si hay un error en este, pero no lo incluyas en los specs calculados ni en los totales de metraje.
             }


        }); // End of itemBlocks.forEach

        console.log("Fin del procesamiento de ítems.");
         console.log("Acumuladores de paneles finales (fraccional/redondeado):", panelAccumulators);
        console.log("Errores totales encontrados:", currentErrorMessages);
        console.log("Totales de Metraje acumulados (NUEVO):", {totalMuroMetrajeAreaSum, totalCieloMetrajeAreaSum, totalCenefaMetrajeLinearSum});


        // --- Final Calculation of Panels from Accumulators ---
        let finalPanelTotals = {};
         for (const type in panelAccumulators) {
             if (panelAccumulators.hasOwnProperty(type)) {
                 const acc = panelAccumulators[type];
                 // Apply ceiling only to the fractional sum before adding to the rounded sum
                 const totalPanelsForType = roundUpFinalUnit(acc.suma_fraccionaria_pequenas) + acc.suma_redondeada_otros;

                 if (totalPanelsForType > 0) {
                      // Store final rounded panel total with descriptive name
                      finalPanelTotals[`Paneles de ${type}`] = totalPanelsForType;
                 }
             }
         }
         console.log("Totales finales de paneles (redondeo aplicado a fraccional + suma de redondeados):", finalPanelTotals);

        // --- Combine Final Panels with Other Materials Total ---
        let finalTotalMaterials = { ...finalPanelTotals, ...otherMaterialsTotal };
        console.log("Total final de materiales combinados (Paneles + Otros):", finalTotalMaterials);

        // --- Store Final Metraje Totales (NUEVO) ---
        const finalTotalMetrajes = {
             'Muro Área Total Metraje (m²)' : totalMuroMetrajeAreaSum,
             'Cielo Área Total Metraje (m²)' : totalCieloMetrajeAreaSum,
             'Cenefa Lineal Total Metraje (m)' : totalCenefaMetrajeLinearSum
        };
        console.log("Totales finales de Metraje para mostrar y almacenar:", finalTotalMetrajes);
        // --- Fin Store Metraje Totales ---


        // --- Display Results ---
        // Si hay errores de validación o errores inesperados, solo muestra los errores.
        if (currentErrorMessages.length > 0) {
            console.log("Mostrando mensajes de error.");
            // Display errors first if any
             resultsContent.innerHTML = '<div class="error-message"><h2>Errores Encontrados:</h2>' +
                                        currentErrorMessages.map(msg => `<p>${msg}</p>`).join('') +
                                        '<p>Por favor, corrige los errores indicados y vuelve a calcular.</p></div>';
             // Clear/reset previous results and hide download buttons
             downloadOptionsDiv.classList.add('hidden');
             lastCalculatedTotalMaterials = {}; // No hay totales válidos para descargar
             lastCalculatedTotalMetrajes = {}; // Limpia metrajes totales almacenados
             lastCalculatedItemsSpecs = []; // No hay specs válidos para descargar
             lastErrorMessages = currentErrorMessages; // Store errors for potential future handling
             lastCalculatedWorkArea = ''; // Limpia el área de trabajo almacenada si hay errores
            return; // Stop here if there are errors
        }

        // Si no hay errores y hay ítems calculados válidamente, muestra los resultados.
        // Verificamos si hay itemspecs válidos O si hay metrajes totales positivos.
        if (currentCalculatedItemsSpecs.length > 0 ||
            finalTotalMetrajes['Muro Área Total Metraje (m²)'] > 0 ||
            finalTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0 ||
            finalTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {
            console.log("No se encontraron errores de validación. Generando resultados HTML.");
            // If no errors, proceed to display results and store them

            let resultsHtml = '<div class="report-header">';
            resultsHtml += '<h2>Resumen de Materiales y Metrajes</h2>'; // Título actualizado
            // --- Muestra el Área de Trabajo en el resumen HTML ---
            if (currentCalculatedWorkArea) {
                 resultsHtml += `<p><strong>Área de Trabajo:</strong> <span>${currentCalculatedWorkArea}</span></p>`;
            }
            // --- Fin Mostrar Área de Trabajo ---
            resultsHtml += `<p>Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}</p>`; // Format date for Spanish
            resultsHtml += '</div>';
            resultsHtml += '<hr>';

            // Display individual item summaries for valid items
            // Mostramos el detalle de ítems SÓLO si hay ítems válidos calculados (itemValidatedSpecs.length > 0)
            if (currentCalculatedItemsSpecs.length > 0) {
                console.log("Generando resumen de ítems calculados.");
                resultsHtml += '<h3>Detalle de Ítems Calculados:</h3>';
                currentCalculatedItemsSpecs.forEach(item => {
                    resultsHtml += `<div class="item-summary">`;
                    // Usamos la información almacenada en itemValidatedSpecs para el resumen del ítem
                    resultsHtml += `<h4>${getItemTypeName(item.type)} #${item.number}</h4>`;
                    resultsHtml += `<p><strong>Tipo:</strong> <span>${getItemTypeName(item.type)}</span></p>`;

                     // --- Muestra el Metraje calculado para este ítem en el resumen principal (NUEVO) ---
                     if (item.type === 'muro' && !isNaN(item.metrajeArea)) {
                         resultsHtml += `<p><strong>Metraje (Área):</strong> <span>${item.metrajeArea.toFixed(2)} m²</span></p>`;
                     } else if (item.type === 'cielo' && !isNaN(item.metrajeArea)) {
                         resultsHtml += `<p><strong>Metraje (Área):</strong> <span>${item.metrajeArea.toFixed(2)} m²</span></p>`;
                     } else if (item.type === 'cenefa' && !isNaN(item.metrajeLinear)) {
                         resultsHtml += `<p><strong>Metraje (Lineal):</strong> <span>${item.metrajeLinear.toFixed(2)} m</span></p>`;
                     }
                    // --- Fin Mostrar Metraje Ítem ---


                    if (item.type === 'muro') {
                        if (!isNaN(item.faces)) resultsHtml += `<p><strong>Nº Caras:</strong> <span>${item.faces}</span></p>`;
                        if (item.cara1PanelType) resultsHtml += `<p><strong>Panel Cara 1:</strong> <span>${item.cara1PanelType}</span></p>`;
                        if (item.faces === 2 && item.cara2PanelType) resultsHtml += `<p><strong>Panel Cara 2:</strong> <span>${item.cara2PanelType}</span></p>`;
                        if (!isNaN(item.postSpacing)) resultsHtml += `<p><strong>Espaciamiento Postes:</strong> <span>${item.postSpacing.toFixed(2)} m</span></p>`;
                         if (item.postType) resultsHtml += `<p><strong>Tipo de Poste:</strong> <span>${item.postType}</span></p>`; // --- Muestra tipo de poste (NUEVO) ---
                        resultsHtml += `<p><strong>Estructura Doble:</strong> <span>${item.isDoubleStructure ? 'Sí' : 'No'}</span></p>`;

                        // Segmentos - Ahora mostramos TODOS los segmentos ingresados, indicando si fueron válidos para materiales y su metraje
                         resultsHtml += `<p><strong>Segmentos:</strong></p>`;
                        if (item.segments && item.segments.length > 0) {
                            item.segments.forEach(seg => {
                                // Mostrar dimensiones reales (input) y metraje calculado para el segmento
                                 let segmentLine = `- Segmento ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.height.toFixed(2)}m (Real)`;
                                 if (!seg.isValidForMaterials) {
                                     segmentLine += ' (Inválido para Materiales)'; // Indica si no se usó para materiales
                                 }
                                 if (!isNaN(seg.metrajeArea)) {
                                      segmentLine += ` - Metraje: ${seg.metrajeArea.toFixed(2)} m²`; // Muestra metraje por segmento
                                 }
                                resultsHtml += `<p style="margin-left: 20px;">${segmentLine}</p>`;
                            });
                             // Mostrar totales calculados (solo si son válidos para materiales)
                             if (item.totalMuroArea > 0) { // totalMuroArea ya solo suma de válidos
                                  resultsHtml += `<p style="margin-left: 20px;"><strong>Área Total Segmentos (Usada para Materiales):</strong> ${item.totalMuroArea.toFixed(2)} m²</p>`; // Aclara que es área usada para materiales
                             }
                             if (item.totalMuroWidth > 0) { // totalMuroWidth ya solo suma de válidos
                                 resultsHtml += `<p style="margin-left: 20px;"><strong>Ancho Total Segmentos (Usado para Materiales):</strong> ${item.totalMuroWidth.toFixed(2)} m</p>`; // Aclara que es ancho usado para materiales
                             }


                        } else {
                             resultsHtml += `<p style="margin-left: 20px;">- Sin segmentos ingresados o válidos</p>`; // Mensaje actualizado
                        }


                    } else if (item.type === 'cielo') {
                        if (item.cieloPanelType) resultsHtml += `<p><strong>Tipo de Panel:</strong> <span>${item.cieloPanelType}</span></p>`;
                        if (!isNaN(item.plenum)) resultsHtml += `<p><strong>Pleno:</strong> <span>${item.plenum.toFixed(2)} m</span></p>`;
                        // --- Agrega el Descuento Angular al Resumen HTML ---
                        if (!isNaN(item.angularDeduction) && item.angularDeduction > 0) { // Muestra solo si es > 0 para claridad
                             resultsHtml += `<p><strong>Descuento Angular:</strong> <span>${item.angularDeduction.toFixed(2)} m</span></p>`;
                        }
                        // --- Fin Agregación ---


                         // Segmentos - Ahora mostramos TODOS los segmentos ingresados, indicando si fueron válidos para materiales y su metraje
                         resultsHtml += `<p><strong>Segmentos:</strong></p>`;
                         if (item.segments && item.segments.length > 0) {
                            item.segments.forEach(seg => {
                                // Mostrar dimensiones reales (input) y metraje calculado para el segmento
                                 let segmentLine = `- Seg ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.length.toFixed(2)}m (Real)`;
                                 if (!seg.isValidForMaterials) {
                                     segmentLine += ' (Inválido para Materiales)'; // Indica si no se usó para materiales
                                 }
                                  if (!isNaN(seg.metrajeArea)) {
                                      segmentLine += ` - Metraje: ${seg.metrajeArea.toFixed(2)} m²`; // Muestra metraje por segmento
                                  }
                                resultsHtml += `<p style="margin-left: 20px;">${segmentLine}</p>`;
                            });
                             // Mostrar totales calculados (solo si son válidos para materiales)
                             if (item.totalCieloArea > 0) { // totalCieloArea ya solo suma de válidos
                                resultsHtml += `<p style="margin-left: 20px;"><strong>Área Total Segmentos (Usada para Materiales):</strong> ${item.totalCieloArea.toFixed(2)} m²</p>`; // Aclara que es área usada para materiales
                            }
                             if (item.totalCieloPerimeterSum > 0) { // totalCieloPerimeterSum ya solo suma de válidos
                                resultsHtml += `<p style="margin-left: 20px;"><strong>Suma Perímetros Segmentos (Usada para Materiales):</strong> ${item.totalCieloPerimeterSum.toFixed(2)} m</p>`; // Aclara que es perímetro usado para materiales
                             }
                        } else {
                            resultsHtml += `<p style="margin-left: 20px;">- Sin segmentos ingresados o válidos</p>`; // Mensaje actualizado
                            finalY += itemSummaryLineHeight; // Ensure space even if no segments listed
                        }
                    } else if (item.type === 'cenefa') { // --- NUEVO: Resumen de Cenefa ---
                        if (!isNaN(item.cenefaLength)) resultsHtml += `<p><strong>Largo Total (Real):</strong> <span>${item.cenefaLength.toFixed(2)} m</span></p>`; // Aclara que es largo real
                        if (!isNaN(item.cenefaFaces)) resultsHtml += `<p><strong>Nº de Caras:</strong> <span>${item.cenefaFaces}</span></p>`;
                        if (item.cenefaPanelType) resultsHtml += `<p><strong>Tipo de Panel:</strong> <span>${item.cenefaPanelType}</span></p>`;
                        // Nota: El metraje lineal ya se mostró al inicio del resumen del ítem.
                    }
                    resultsHtml += `</div>`;
                });
                resultsHtml += '<hr>';
            } else {
                 console.log("No hay ítems válidos para mostrar en el detalle de ítems calculados.");
                 // Mostrar un mensaje si no hay ítems válidos calculados pero sí hay totales de metraje (ej: solo cenefas válidas)
                 if (finalTotalMetrajes['Muro Área Total Metraje (m²)'] > 0 ||
                     finalTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0 ||
                     finalTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {
                      resultsHtml += '<p>No hay ítems de Muro o Cielo con dimensiones válidas para el cálculo detallado de materiales, pero se calcularon metrajes totales.</p>';
                      resultsHtml += '<hr>';
                 }
            }


            // --- NUEVA SECCIÓN: Totales de Metraje ---
            console.log("Generando sección de totales de metraje.");
             // Solo añade la sección si hay algún metraje total positivo
             if (finalTotalMetrajes['Muro Área Total Metraje (m²)'] > 0 ||
                 finalTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0 ||
                 finalTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {

                  resultsHtml += '<h3>Totales de Metraje:</h3>';
                  resultsHtml += '<ul>';
                  // Asegura que solo se muestren los metrajes que tienen un valor positivo
                  if (finalTotalMetrajes['Muro Área Total Metraje (m²)'] > 0) {
                      resultsHtml += `<li><strong>Muro Área Total:</strong> ${finalTotalMetrajes['Muro Área Total Metraje (m²)'].toFixed(2)} m²</li>`;
                  }
                   if (finalTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0) {
                      resultsHtml += `<li><strong>Cielo Área Total:</strong> ${finalTotalMetrajes['Cielo Área Total Metraje (m²)'].toFixed(2)} m²</li>`;
                   }
                   if (finalTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {
                       resultsHtml += `<li><strong>Cenefa Lineal Total:</strong> ${finalTotalMetrajes['Cenefa Lineal Total Metraje (m)'].toFixed(2)} m</li>`;
                   }
                  resultsHtml += '</ul>';
                  resultsHtml += '<hr>';
             } else {
                  console.log("No hay totales de metraje positivos para añadir sección.");
             }
            // --- Fin NUEVA SECCIÓN Metraje ---


            // Tabla de Totales de Materiales
             // Solo añade la tabla si hay materiales con cantidad > 0
             if (Object.keys(finalTotalMaterials).some(material => finalTotalMaterials[material] > 0)) {
                 console.log("Generando tabla de materiales totales.");
                 resultsHtml += '<h3>Totales de Materiales (Cantidades a Comprar):</h3>'; // Título actualizado

                 // finalTotalMaterials now holds the combined totals
                 console.log("Total final de materiales calculados (combinados):", finalTotalMaterials);

                 // Sort materials alphabetically for consistent display
                 const sortedMaterials = Object.keys(finalTotalMaterials).sort();

                resultsHtml += '<table><thead><tr><th>Material</th><th>Cantidad</th><th>Unidad</th></tr></thead><tbody>';

                sortedMaterials.forEach(material => {
                    const cantidad = finalTotalMaterials[material];
                    const unidad = getMaterialUnit(material); // Get unit using the helper function
                    // Display the material name
                     // Solo muestra los materiales con cantidad mayor a 0
                     if (cantidad > 0) {
                         resultsHtml += `<tr><td>${material}</td><td>${cantidad}</td><td>${unidad}</td></tr>`;
                     }
                });

                 // Si no hay materiales con cantidad > 0, muestra un mensaje en la tabla
                 if (sortedMaterials.every(material => finalTotalMaterials[material] <= 0)) {
                     resultsHtml += `<tr><td colspan="3" style="text-align: center;">No se calcularon cantidades positivas de materiales.</td></tr>`;
                 }


                resultsHtml += '</tbody></table>';
                downloadOptionsDiv.classList.remove('hidden'); // Show download options

             } else {
                  console.log("No hay materiales con cantidad positiva para añadir tabla.");
                  // Si no hay materiales positivos ni metrajes positivos, el mensaje inicial de error ya se mostró.
                  // Si hay metrajes positivos pero no materiales, se mostró la sección de metrajes y este bloque no añade la tabla.
             }


            // Append the generated HTML to the results content area
            resultsContent.innerHTML = resultsHtml; // Use = to replace previous content

            console.log("Resultados HTML generados y añadidos al DOM.");

            // --- Almacena el estado completo de los resultados, incluyendo Área de Trabajo y Metrajes Totales ---
            lastCalculatedTotalMaterials = finalTotalMaterials;
            lastCalculatedTotalMetrajes = finalTotalMetrajes; // --- ALMACENA Metrajes Totales (NUEVO) ---
            lastCalculatedItemsSpecs = currentCalculatedItemsSpecs;
            lastErrorMessages = []; // Si llegamos aquí, no hay errores de validación que impidan mostrar resultados
            lastCalculatedWorkArea = currentCalculatedWorkArea; // Almacena el área de trabajo leído

             console.log("Estado de resultados almacenado para descarga.");

        } else {
             console.log("No hay ítems válidos calculados ni metrajes totales positivos.");
             // Esto podría pasar si no hay ítems válidos o si los cálculos resultan en 0 para todos los materiales y metrajes.
             resultsContent.innerHTML = '<p>No se pudieron calcular materiales ni metrajes con las dimensiones ingresadas. Revisa los valores y si hay errores de validación.</p>'; // Mensaje actualizado
             downloadOptionsDiv.classList.add('hidden'); // Hide download options
             lastCalculatedTotalMaterials = {}; // Limpia los totales almacenados
             lastCalculatedTotalMetrajes = {}; // Limpia metrajes totales almacenados
             lastCalculatedItemsSpecs = []; // Limpia los specs almacenados
             // lastErrorMessages ya contiene los errores si los hubo.
             lastCalculatedWorkArea = ''; // Limpia el área de trabajo almacenada
        }

    };


    // --- PDF Generation Function ---
    // Modificada para incluir los metrajes
    const generatePDF = () => {
        console.log("Iniciando generación de PDF...");
       // Ensure there are calculated results to download
       // Ahora verificamos también si hay metrajes O materiales calculados
       if (Object.keys(lastCalculatedTotalMaterials).length === 0 &&
           (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] <= 0 &&
            lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] <= 0 &&
            lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] <= 0) ||
           lastCalculatedItemsSpecs.length === 0) { // Aunque haya metrajes totales, necesitamos ítems válidos para el detalle de ítems en el reporte.
           // Si solo hay metrajes totales (ej: ítems inválidos para materiales pero válidos para metraje), se generará un PDF solo con encabezado y totales de metraje.
           // Vamos a permitir generar PDF si hay itemspecs VÁLIDOS (incluso si materiales dan 0) O si hay metrajes totales positivos.
            if (lastCalculatedItemsSpecs.length === 0 &&
               (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] <= 0 &&
                lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] <= 0 &&
                lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] <= 0))
           {
               console.warn("No hay resultados calculados (materiales o metrajes) para generar el PDF.");
               alert("Por favor, realiza un cálculo válido antes de generar el PDF.");
               return;
           }
       }


       // Initialize jsPDF
       const { jsPDF } = window.jspdf;
       const doc = new jsPDF();

       // Define colors in RGB from CSS variables (using approximations based on common web colors)
       const primaryOliveRGB = [85, 107, 47]; // #556B2F
       const secondaryOliveRGB = [128, 128, 0]; // #808000
       const darkGrayRGB = [51, 51, 51]; // #333
       const mediumGrayRGB = [102, 102, 102]; // #666
       const lightGrayRGB = [224, 224, 224]; // #e0e0e0
       const extraLightGrayRGB = [248, 248, 248]; // #f8f8f8


       // --- Add Header ---
       doc.setFontSize(18);
       doc.setTextColor(primaryOliveRGB[0], primaryOliveRGB[1], primaryOliveRGB[2]);
       doc.setFont("helvetica", "bold"); // Use a standard font or include custom fonts
       doc.text("Resumen de Materiales y Metrajes Tablayeso", 14, 22); // Título actualizado

       doc.setFontSize(10);
       doc.setTextColor(mediumGrayRGB[0], mediumGrayRGB[1], mediumGrayRGB[2]);
        doc.setFont("helvetica", "normal");
       doc.text(`Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}`, 14, 28);

        // --- Agrega el Área de Trabajo al Encabezado del PDF ---
        let currentY = 35; // Posición inicial después de la fecha
        if (lastCalculatedWorkArea) {
             doc.text(`Área de Trabajo: ${lastCalculatedWorkArea}`, 14, currentY);
             currentY += 7; // Deja un poco de espacio si se muestra el área de trabajo
        }
        // Set starting Y position for the next content block
        let finalY = currentY;
        // --- Fin Agregar Área de Trabajo ---


       // --- Add Item Summaries ---
       // Solo añade la sección si hay ítems válidos calculados (itemValidatedSpecs.length > 0)
       if (lastCalculatedItemsSpecs.length > 0) {
            console.log("Añadiendo resumen de ítems al PDF.");
            doc.setFontSize(14);
            doc.setTextColor(secondaryOliveRGB[0], secondaryOliveRGB[1], secondaryOliveRGB[2]);
           doc.setFont("helvetica", "bold");
            doc.text("Detalle de Ítems Calculados:", 14, finalY + 10);
            finalY += 15; // Move Y below the title

           const itemSummaryLineHeight = 5; // Space between summary lines within an item
           const itemBlockSpacing = 8; // Space between different item summaries

            lastCalculatedItemsSpecs.forEach(item => {
                // Add item title (using type and number from specs)
                doc.setFontSize(10);
                doc.setTextColor(primaryOliveRGB[0], primaryOliveRGB[1], primaryOliveRGB[2]);
                doc.setFont("helvetica", "bold");
                doc.text(`${getItemTypeName(item.type)} #${item.number}:`, 14, finalY + itemSummaryLineHeight);
                finalY += itemSummaryLineHeight * 1.5; // Move down after the title

                // Add general item details (indented)
                doc.setFontSize(9);
                doc.setTextColor(darkGrayRGB[0], darkGrayRGB[1], darkGrayRGB[2]);
                doc.setFont("helvetica", "normal");

                doc.text(`Tipo: ${getItemTypeName(item.type)}`, 20, finalY + itemSummaryLineHeight);
                finalY += itemSummaryLineHeight;

                 // --- Muestra el Metraje calculado para este ítem en el resumen del PDF (NUEVO) ---
                 if (item.type === 'muro' && !isNaN(item.metrajeArea)) {
                      doc.text(`Metraje (Área): ${item.metrajeArea.toFixed(2)} m²`, 20, finalY + itemSummaryLineHeight);
                      finalY += itemSummaryLineHeight;
                 } else if (item.type === 'cielo' && !isNaN(item.metrajeArea)) {
                      doc.text(`Metraje (Área): ${item.metrajeArea.toFixed(2)} m²`, 20, finalY + itemSummaryLineHeight);
                      finalY += itemSummaryLineHeight;
                 } else if (item.type === 'cenefa' && !isNaN(item.metrajeLinear)) {
                      doc.text(`Metraje (Lineal): ${item.metrajeLinear.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                      finalY += itemSummaryLineHeight;
                 }
                 // --- Fin Mostrar Metraje Ítem en PDF ---


                // Add type-specific details
                if (item.type === 'muro') {
                     if (!isNaN(item.faces)) {
                          doc.text(`Nº Caras: ${item.faces}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                     if (item.cara1PanelType) {
                          doc.text(`Panel Cara 1: ${item.cara1PanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                     if (item.faces === 2 && item.cara2PanelType) {
                          doc.text(`Panel Cara 2: ${item.cara2PanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                      if (!isNaN(item.postSpacing)) {
                         doc.text(`Espaciamiento Postes: ${item.postSpacing.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                     if (item.postType) { // --- Muestra tipo de poste en PDF (NUEVO) ---
                         doc.text(`Tipo de Poste: ${item.postType}`, 20, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                     }
                      doc.text(`Estructura Doble: ${item.isDoubleStructure ? 'Sí' : 'No'}`, 20, finalY + itemSummaryLineHeight);
                       finalY += itemSummaryLineHeight;

                     doc.text(`Segmentos:`, 20, finalY + itemSummaryLineHeight); // Etiqueta simple, detalles abajo
                     finalY += itemSummaryLineHeight;
                      if (item.segments && item.segments.length > 0) {
                         item.segments.forEach(seg => {
                             // Mostrar dimensiones reales (input) y metraje calculado para el segmento
                              let segmentLine = `- Seg ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.height.toFixed(2)}m (Real)`;
                              if (!seg.isValidForMaterials) {
                                  segmentLine += ' (Inválido para Materiales)'; // Indica si no se usó para materiales
                              }
                              if (!isNaN(seg.metrajeArea)) {
                                   segmentLine += ` - Metraje: ${seg.metrajeArea.toFixed(2)} m²`; // Muestra metraje por segmento
                              }
                             doc.text(segmentLine, 25, finalY + itemSummaryLineHeight); // Usar texto en lugar de p para PDF
                             finalY += itemSummaryLineHeight;
                         });
                          // Mostrar totales calculados (solo si son válidos para materiales)
                           if (item.totalMuroArea > 0) { // totalMuroArea ya solo suma de válidos
                                doc.text(`- Área Total Segmentos (Usada para Materiales): ${item.totalMuroArea.toFixed(2)} m²`, 25, finalY + itemSummaryLineHeight); // Aclara que es área usada para materiales
                                finalY += itemSummaryLineHeight;
                           }
                            if (item.totalMuroWidth > 0) { // totalMuroWidth ya solo suma de válidos
                                doc.text(`- Ancho Total Segmentos (Usado para Materiales): ${item.totalMuroWidth.toFixed(2)} m`, 25, finalY + itemSummaryLineHeight); // Aclara que es ancho usado para materiales
                                finalY += itemSummaryLineHeight;
                           }


                      } else {
                           doc.text(`- Sin segmentos ingresados o válidos`, 25, finalY + itemSummaryLineHeight); // Mensaje actualizado
                           finalY += itemSummaryLineHeight;
                      }


                } else if (item.type === 'cielo') {
                     if (item.cieloPanelType) {
                          doc.text(`Tipo de Panel: ${item.cieloPanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                    if (!isNaN(item.plenum)) {
                        doc.text(`Pleno: ${item.plenum.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                    }
                    // --- Agrega el Descuento Angular al Resumen del PDF ---
                    if (!isNaN(item.angularDeduction) && item.angularDeduction > 0) { // Muestra solo si es > 0 para claridad
                        doc.text(`Descuento Angular: ${item.angularDeduction.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                        finalY += itemSummaryLineHeight;
                    }
                    // --- Fin Agregación ---

                     doc.text(`Segmentos:`, 20, finalY + itemSummaryLineHeight); // Etiqueta simple
                     finalY += itemSummaryLineHeight;
                     if (item.segments && item.segments.length > 0) {
                         item.segments.forEach(seg => {
                             // Mostrar dimensiones reales (input) y metraje calculado para el segmento
                             let segmentLine = `- Seg ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.length.toFixed(2)}m (Real)`;
                              if (!seg.isValidForMaterials) {
                                   segmentLine += ' (Inválido para Materiales)'; // Indica si no se usó para materiales
                              }
                              if (!isNaN(seg.metrajeArea)) {
                                   segmentLine += ` - Metraje: ${seg.metrajeArea.toFixed(2)} m²`; // Muestra metraje por segmento
                              }
                             doc.text(segmentLine, 25, finalY + itemSummaryLineHeight); // Usar texto
                             finalY += itemSummaryLineHeight;
                         });
                         // Mostrar totales calculados (solo si son válidos para materiales)
                          if (item.totalCieloArea > 0) { // totalCieloArea ya solo suma de válidos
                             doc.text(`- Área Total Segmentos (Usada para Materiales): ${item.totalCieloArea.toFixed(2)} m²`, 25, finalY + itemSummaryLineHeight); // Aclara que es área usada para materiales
                             finalY += itemSummaryLineHeight;
                         }
                          if (item.totalCieloPerimeterSum > 0) { // totalCieloPerimeterSum ya solo suma de válidos
                             doc.text(`- Suma Perímetros Segmentos (Usada para Materiales): ${item.totalCieloPerimeterSum.toFixed(2)} m`, 25, finalY + itemSummaryLineHeight); // Aclara que es perímetro usado para materiales
                             finalY += itemSummaryLineHeight;
                         }

                     } else {
                         doc.text(`- Sin segmentos ingresados o válidos`, 25, finalY + itemSummaryLineHeight); // Mensaje actualizado
                         finalY += itemSummaryLineHeight;
                     }
                } else if (item.type === 'cenefa') { // --- NUEVO: Resumen de Cenefa en PDF ---
                     if (!isNaN(item.cenefaLength)) {
                         doc.text(`Largo Total (Real): ${item.cenefaLength.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight); // Etiqueta actualizada
                         finalY += itemSummaryLineHeight;
                     }
                     if (!isNaN(item.cenefaFaces)) {
                         doc.text(`Nº de Caras: ${item.cenefaFaces}`, 20, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                     }
                     if (item.cenefaPanelType) {
                         doc.text(`Tipo de Panel: ${item.cenefaPanelType}`, 20, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                     }
                     // Nota: El metraje lineal ya se mostró al inicio del resumen del ítem.
                }
                finalY += itemBlockSpacing; // Add space after each item summary block
            });
            finalY += 5; // Add space before the next title
       } else {
            console.log("No hay ítems válidos calculados para añadir resumen al PDF.");
            // Si no hay ítems válidos para el detalle, pero hay metrajes totales positivos, dejamos un espacio antes de la sección de totales.
             if (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] > 0 ||
                 lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0 ||
                 lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {
                  finalY += 10; // Añade un poco más de espacio si no hay detalle de ítems pero sí totales
             }
       }

        // --- NUEVA SECCIÓN: Totales de Metraje en PDF ---
        console.log("Añadiendo sección de totales de metraje al PDF.");
        // Solo añade la sección si hay algún metraje total positivo
        if (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] > 0 ||
            lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0 ||
            lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {

             doc.setFontSize(14);
             doc.setTextColor(secondaryOliveRGB[0], secondaryOliveRGB[1], secondaryOliveRGB[2]);
             doc.setFont("helvetica", "bold");
             doc.text("Totales de Metraje:", 14, finalY + 10);
             finalY += 15; // Move Y below the title

             doc.setFontSize(9);
             doc.setTextColor(darkGrayRGB[0], darkGrayRGB[1], darkGrayRGB[2]);
             doc.setFont("helvetica", "normal");

             const metrajeLineHeight = 5; // Espacio entre líneas de metraje
              if (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] > 0) {
                  doc.text(`Muro Área Total: ${lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'].toFixed(2)} m²`, 20, finalY + metrajeLineHeight);
                  finalY += metrajeLineHeight;
              }
               if (lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0) {
                  doc.text(`Cielo Área Total: ${lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'].toFixed(2)} m²`, 20, finalY + metrajeLineHeight);
                   finalY += metrajeLineHeight;
               }
               if (lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {
                   doc.text(`Cenefa Lineal Total: ${lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'].toFixed(2)} m`, 20, finalY + metrajeLineHeight);
                   finalY += metrajeLineHeight;
               }
             finalY += 8; // Espacio después de la sección de metraje
        } else {
             console.log("No hay totales de metraje positivos para añadir sección al PDF.");
        }
       // --- Fin NUEVA SECCIÓN Metraje en PDF ---


       // --- Add Total Materials Table ---
        // Solo añade la tabla si hay materiales con cantidad > 0
        if (Object.keys(lastCalculatedTotalMaterials).some(material => lastCalculatedTotalMaterials[material] > 0)) {
            console.log("Añadiendo tabla de materiales totales al PDF.");
             doc.setFontSize(14);
             doc.setTextColor(secondaryOliveRGB[0], secondaryOliveRGB[1], secondaryOliveRGB[2]);
             doc.setFont("helvetica", "bold");
             doc.text("Totales de Materiales (Cantidades a Comprar):", 14, finalY + 10); // Título actualizado
             finalY += 15; // Move Y below the title

            const tableColumn = ["Material", "Cantidad", "Unidad"];
            const tableRows = [];

            // Prepare data for the table, only including materials with quantity > 0
            const sortedMaterials = Object.keys(lastCalculatedTotalMaterials).sort();
            sortedMaterials.forEach(material => {
                const cantidad = lastCalculatedTotalMaterials[material];
                if (cantidad > 0) { // Solo añade si la cantidad es positiva
                     const unidad = getMaterialUnit(material); // Get unit using the helper function
                    // Use the material name directly from the key
                     tableRows.push([material, cantidad, unidad]);
                }
            });

             // Add the table using jspdf-autotable
             doc.autoTable({
                 head: [tableColumn],
                 body: tableRows,
                 startY: finalY, // Start position below the last content
                 theme: 'plain', // Start with a plain theme to apply custom styles
                 headStyles: {
                     fillColor: lightGrayRGB,
                     textColor: darkGrayRGB,
                     fontStyle: 'bold',
                     halign: 'center', // Horizontal alignment
                     valign: 'middle', // Vertical alignment
                     lineWidth: 0.1,
                     lineColor: lightGrayRGB,
                     fontSize: 10 // Match HTML table header font size
                 },
                 bodyStyles: {
                     textColor: darkGrayRGB,
                     lineWidth: 0.1,
                     lineColor: lightGrayRGB,
                     fontSize: 9 // Match HTML table body font size
                 },
                  alternateRowStyles: { // Styling for alternate rows
                     fillColor: extraLightGrayRGB,
                 },
                  // Specific column styles (Cantidad column is the second one, index 1)
                 columnStyles: {
                     1: {
                         halign: 'right', // Align quantity to the right
                         fontStyle: 'bold',
                         textColor: primaryOliveRGB // For quantity text color
                     },
                      2: { // Unit column
                         halign: 'center' // Align unit to the center or left as preferred
                     }
                 },
                 margin: { top: 10, right: 14, bottom: 14, left: 14 }, // Add margin
                  didDrawPage: function (data) {
                    // Optional: Add page number or footer here
                    doc.setFontSize(8);
                    doc.setTextColor(mediumGrayRGB[0], mediumGrayRGB[1], mediumGrayRGB[2]);
                    // Using `doc.internal.pageSize.getWidth()` to center the footer roughly
                    const footerText = '© 2025 [PROPUL] - Calculadora de Materiales Tablayeso v2.0'; // Replaced placeholder
                    const textWidth = doc.getStringUnitWidth(footerText) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                    const centerX = (doc.internal.pageSize.getWidth() - textWidth) / 2;
                    doc.text(footerText, centerX, doc.internal.pageSize.height - 10);

                    // Add simple page number
                    const pageNumberText = `Página ${data.pageNumber}`;
                    const pageNumberWidth = doc.getStringUnitWidth(pageNumberText) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                    const pageNumberX = doc.internal.pageSize.getWidth() - data.settings.margin.right - pageNumberWidth;
                    doc.text(pageNumberText, pageNumberX, doc.internal.pageSize.height - 10);
                 }
             });

             // Update finalY after the table
            finalY = doc.autoTable.previous.finalY;

            console.log("PDF generado.");
            // --- Save the PDF ---
            doc.save(`Calculo_Materiales_${new Date().toLocaleDateString('es-ES').replace(/\//g, '-')}.pdf`); // Filename with date

       } else {
             console.log("No hay materiales con cantidad positiva para añadir tabla al PDF.");
             // Si no hay materiales positivos pero sí hubo metrajes, igual guardamos el PDF con el resumen y los metrajes.
             // Si tampoco hubo metrajes, la función ya habría salido al inicio.
              console.log("PDF generado (solo resumen de ítems y metrajes si aplica).");
             doc.save(`Calculo_Materiales_${new Date().toLocaleDateString('es-ES').replace(/\//g, '-')}.pdf`); // Guarda el PDF aunque no haya tabla de materiales
       }

   };


// --- Excel Generation Function ---
// Modificada para incluir los metrajes
const generateExcel = () => {
    console.log("Iniciando generación de Excel...");
    // Ensure there are calculated results to download (materiales o metrajes)
    if (Object.keys(lastCalculatedTotalMaterials).length === 0 &&
        (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] <= 0 &&
         lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] <= 0 &&
         lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] <= 0) ||
        lastCalculatedItemsSpecs.length === 0) { // Aunque haya metrajes totales, necesitamos ítems válidos para el detalle de ítems en el reporte.
         // Si solo hay metrajes totales (ej: ítems inválidos para materiales pero válidos para metraje), se generará un Excel solo con encabezado y totales de metraje.
         // Vamos a permitir generar Excel si hay itemspecs VÁLIDOS (incluso si materiales dan 0) O si hay metrajes totales positivos.
          if (lastCalculatedItemsSpecs.length === 0 &&
             (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] <= 0 &&
              lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] <= 0 &&
              lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] <= 0))
          {
              console.warn("No hay resultados calculados (materiales o metrajes) para generar el Excel.");
              alert("Por favor, realiza un cálculo válido antes de generar el Excel.");
              return;
          }
       }

    // Assumes you have loaded the xlsx library
   if (typeof XLSX === 'undefined') {
        console.error("La librería xlsx no está cargada.");
        alert("Error al generar Excel: Librería xlsx no encontrada.");
        return;
   }


   // Data for the Excel sheet
   let sheetData = [];

   // Add Header
   sheetData.push(["Calculadora de Materiales y Metrajes Tablayeso"]); // Título actualizado
   sheetData.push([`Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}`]);
   // --- Agrega el Área de Trabajo al Encabezado del Excel ---
   if (lastCalculatedWorkArea) {
       sheetData.push([`Área de Trabajo: ${lastCalculatedWorkArea}`]);
   }
   // --- Fin Agregar Área de Trabajo ---
   sheetData.push([]); // Fila en blanco para espaciar


   // Add Item Summaries
   // Solo añade la sección si hay ítems válidos calculados (itemValidatedSpecs.length > 0)
   if (lastCalculatedItemsSpecs.length > 0) {
        console.log("Añadiendo resumen de ítems al Excel.");
       sheetData.push(["Detalle de Ítems Calculados:"]);
       // --- ENCABEZADOS DE LA TABLA DE DETALLE DE ÍTEMS ---
       // Incluye columnas para Muro/Cielo/Cenefa específicos y las columnas de metraje por ítem
       sheetData.push([
           "Tipo Item", "Número Item", "Detalle/Dimensiones",
           "Nº Caras (Muro/Cenefa)", // Etiqueta actualizada
           "Panel Cara 1 (Muro)", // Etiqueta más específica
           "Panel Cara 2 (Muro)", // Etiqueta más específica
           "Tipo Panel (Cielo/Cenefa)", // Etiqueta más específica
           "Espaciamiento Postes (m)", "Tipo de Poste", // --- NUEVA columna Tipo de Poste ---
           "Estructura Doble",
           "Pleno (m)", "Metros Descuento Angular (m)",
           "Largo Total (Cenefa) (m)", // --- NUEVA columna para Largo Cenefa ---
           "Suma Perímetros Segmentos (Usada para Materiales) (m)", // Etiqueta actualizada
           "Ancho Total (Usado para Materiales) (m)", // Etiqueta actualizada
           "Área Total (Usada para Materiales) (m²)", // Etiqueta actualizada
           "Metraje (Área Muro/Cielo) (m²)", // --- NUEVA columna Metraje Área por Ítem ---
           "Metraje (Lineal Cenefa) (m)" // --- NUEVA columna Metraje Lineal por Ítem ---
       ]);
       // --- FIN ENCABEZADOS ---


       lastCalculatedItemsSpecs.forEach(item => {
            if (item.type === 'muro') {
                // --- Detalles comunes a este ítem Muro (configuración) ---
                // Aseguramos que el array tenga el tamaño correcto, llenando las columnas de Cielo/Cenefa con vacío.
                const muroRowBase = [
                    getItemTypeName(item.type), // 0: Tipo Item
                    item.number,                 // 1: Número Item
                    '',                          // 2: Placeholder para Detalle/Dimensiones
                    !isNaN(item.faces) ? item.faces : '', // 3: Nº Caras (Muro/Cenefa)
                    item.cara1PanelType ? item.cara1PanelType : '', // 4: Panel Cara 1 (Muro)
                    item.faces === 2 && item.cara2PanelType ? item.cara2PanelType : '', // 5: Panel Cara 2 (Muro)
                    '',                          // 6: Tipo Panel (Cielo/Cenefa) (vacío para muro)
                    !isNaN(item.postSpacing) ? item.postSpacing.toFixed(2) : '', // 7: Espaciamiento Postes
                    item.postType ? item.postType : '', // 8: Tipo de Poste (NUEVO)
                    item.isDoubleStructure ? 'Sí' : 'No', // 9: Estructura Doble
                    '', '',                     // 10, 11: Pleno, Descuento Angular (vacío para muro)
                    ''                           // 12: Largo Total (Cenefa) (vacío para muro)
                    // Las columnas de Totales Reales y Metrajes por Ítem empiezan después (13, 14, 15, 16, 17)
                ];

                // Fila principal con opciones y totales del ítem
                const muroSummaryRow = [...muroRowBase]; // Copia los detalles comunes
                muroSummaryRow[2] = 'Opciones:'; // Etiqueta en la columna de detalle
                // --- Agrega los Totales Reales del Muro y Metraje del Ítem ---
                muroSummaryRow.push(
                   '',                                                               // 13: Suma Perímetros Segmentos (Real) (vacío para Muro)
                   !isNaN(item.totalMuroWidth) ? item.totalMuroWidth.toFixed(2) : '', // 14: Ancho Total (Muro Real)
                   !isNaN(item.totalMuroArea) ? item.totalMuroArea.toFixed(2) : '',     // 15: Área Total (Real)
                   !isNaN(item.metrajeArea) ? item.metrajeArea.toFixed(2) : '',        // 16: Metraje (Área Muro/Cielo) (NUEVO)
                   ''                                                              // 17: Metraje (Lineal Cenefa) (vacío para muro)
                );
                // --- Fin Totales Reales y Metraje Ítem Muro ---
                sheetData.push(muroSummaryRow);

                // Fila que etiqueta la sección de Segmentos
                 const muroSegmentsLabelRow = [...muroRowBase];
                 muroSegmentsLabelRow[2] = 'Segmentos:'; // Etiqueta "Segmentos:"
                 // Agrega celdas vacías para las columnas de totales reales y metrajes
                 muroSegmentsLabelRow.push('', '', '', '', ''); // 13, 14, 15, 16, 17
                 sheetData.push(muroSegmentsLabelRow);


                 if (item.segments && item.segments.length > 0) {
                     item.segments.forEach(seg => {
                          // Fila para cada Segmento individual
                          const segmentRow = [...muroRowBase]; // Copia los detalles comunes para esta fila
                          let segmentDetails = `- Seg ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.height.toFixed(2)}m (Real)`;
                          if (!seg.isValidForMaterials) {
                              segmentDetails += ' (Inválido para Materiales)';
                          }
                          if (!isNaN(seg.metrajeArea)) {
                              segmentDetails += ` - Metraje: ${seg.metrajeArea.toFixed(2)} m²`;
                          }
                          segmentRow[2] = segmentDetails; // Dimensiones reales y metraje del segmento
                          // Agrega celdas vacías para las columnas de totales reales y metrajes (repetir vacíos)
                          segmentRow.push('', '', '', '', ''); // 13, 14, 15, 16, 17
                          sheetData.push(segmentRow);

                          // --- NUEVO: Fila para el metraje calculado por segmento (opcional, si queremos mostrarlo en Excel) ---
                          // Puede ser demasiados detalles, el resumen por ítem puede ser suficiente.
                          // Si se quiere, se podría añadir otra fila con "Metraje Seg X: Y m²" en la columna 2.
                          // Por ahora, dejemos el metraje a nivel de ítem para simplificar el reporte en Excel.
                          // Si el usuario lo pide, podemos añadirlo.

                     });
                  } else {
                       // Fila para "Sin segmentos ingresados o válidos"
                       const noSegmentsRow = [...muroRowBase];
                       noSegmentsRow[2] = `- Sin segmentos ingresados o válidos`; // Mensaje actualizado
                        // Agrega celdas vacías para las columnas de totales reales y metrajes
                        noSegmentsRow.push('', '', '', '', ''); // 13, 14, 15, 16, 17
                        sheetData.push(noSegmentsRow);
                  }


            } else if (item.type === 'cielo') {
                // --- Detalles comunes a este ítem Cielo (configuración) ---
                // Aseguramos que el array tenga el tamaño correcto, llenando las columnas de Muro/Cenefa con vacío.
                const cieloRowBase = [
                    getItemTypeName(item.type), // 0: Tipo Item
                    item.number,                 // 1: Número Item
                    '',                          // 2: Placeholder para Detalle/Dimensiones
                    '',                          // 3: Nº Caras (vacío para cielo)
                    '', '',                     // 4, 5: Panel Cara 1, Cara 2 (vacío para cielo)
                    item.cieloPanelType ? item.cieloPanelType : '', // 6: Tipo Panel (Cielo/Cenefa)
                    '',                          // 7: Espaciamiento Postes (vacío para cielo)
                    '',                          // 8: Tipo de Poste (vacío para cielo)
                    ''                           // 9: Estructura Doble (vacío para cielo)
                    // Las columnas de Pleno, Descuento Angular, Largo Cenefa, Totales Reales y Metrajes por Ítem empiezan después (10, 11, 12, 13, 14, 15, 16, 17)
                ];

                 // Fila principal con opciones y totales del ítem
                const cieloSummaryRow = [...cieloRowBase]; // Copia los detalles comunes
                cieloSummaryRow[2] = 'Opciones:'; // Etiqueta en la columna de detalle
                // Agrega Pleno y Descuento Angular
                cieloSummaryRow.push(
                     !isNaN(item.plenum) ? item.plenum.toFixed(2) : '', // 10: Pleno
                     !isNaN(item.angularDeduction) ? item.angularDeduction.toFixed(2) : '', // 11: Metros Descuento Angular
                     '' // 12: Largo Total (Cenefa) (vacío para cielo)
                 );
                // --- Agrega los Totales Reales del Cielo y Metraje del Ítem ---
                cieloSummaryRow.push(
                   !isNaN(item.totalCieloPerimeterSum) ? item.totalCieloPerimeterSum.toFixed(2) : '', // 13: Suma Perímetros Segmentos (Real)
                   '',                                                              // 14: Ancho Total (Muro Real) - Vacío para Cielo
                   !isNaN(item.totalCieloArea) ? item.totalCieloArea.toFixed(2) : '', // 15: Área Total (Real)
                   !isNaN(item.metrajeArea) ? item.metrajeArea.toFixed(2) : '',        // 16: Metraje (Área Muro/Cielo) (NUEVO)
                   ''                                                              // 17: Metraje (Lineal Cenefa) (vacío para cielo)
                );
                // --- Fin Totales Reales y Metraje Ítem Cielo ---
                sheetData.push(cieloSummaryRow);

                // Fila que etiqueta la sección de Segmentos
                 const cieloSegmentsLabelRow = [...cieloRowBase];
                 cieloSegmentsLabelRow[2] = 'Segmentos:'; // Etiqueta "Segmentos:"
                  // Agrega celdas vacías para Pleno, Descuento Angular, Largo Cenefa, totales reales y metrajes
                  cieloSegmentsLabelRow.push('', '', '', '', '', '', '', ''); // 10 al 17
                  sheetData.push(cieloSegmentsLabelRow);


                  if (item.segments && item.segments.length > 0) {
                      item.segments.forEach(seg => {
                          // Fila para cada Segmento individual
                          const segmentRow = [...cieloRowBase]; // Copia los detalles comunes para esta fila
                          let segmentDetails = `Seg ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.length.toFixed(2)}m (Real)`;
                           if (!seg.isValidForMaterials) {
                                segmentDetails += ' (Inválido para Materiales)';
                           }
                           if (!isNaN(seg.metrajeArea)) {
                                segmentDetails += ` - Metraje: ${seg.metrajeArea.toFixed(2)} m²`;
                           }
                           segmentRow[2] = segmentDetails; // Dimensiones reales y metraje del segmento
                           // Agrega celdas vacías para Pleno, Descuento Angular, Largo Cenefa, totales reales y metrajes (repetir vacíos)
                           segmentRow.push('', '', '', '', '', '', '', ''); // 10 al 17
                           sheetData.push(segmentRow);
                      });
                   } else {
                       // Fila para "Sin segmentos ingresados o válidos"
                       const noSegmentsRow = [...cieloRowBase];
                       noSegmentsRow[2] = `- Sin segmentos ingresados o válidos`; // Mensaje actualizado
                        // Agrega celdas vacías para Pleno, Descuento Angular, Largo Cenefa, totales reales y metrajes
                        noSegmentsRow.push('', '', '', '', '', '', '', ''); // 10 al 17
                        sheetData.push(noSegmentsRow);
                   }

            } else if (item.type === 'cenefa') { // --- NUEVO: Resumen de Cenefa en Excel ---
                // Aseguramos que el array tenga el tamaño correcto, llenando las columnas de Muro/Cielo con vacío.
                 const cenefaRowBase = [
                     getItemTypeName(item.type), // 0: Tipo Item
                     item.number,                 // 1: Número Item
                     'Opciones:',                 // 2: Etiqueta para opciones
                     !isNaN(item.cenefaFaces) ? item.cenefaFaces : '', // 3: Nº Caras (Muro/Cenefa)
                     '', '',                      // 4, 5: Panel Cara 1, Cara 2 (vacío para cenefa)
                     item.cenefaPanelType ? item.cenefaPanelType : '', // 6: Tipo Panel (Cielo/Cenefa)
                     '', '',                      // 7, 8: Espaciamiento Postes, Tipo de Poste (vacío para cenefa)
                     ''                            // 9: Estructura Doble (vacío para cenefa)
                     // Las columnas de Pleno, Descuento Angular, Largo Cenefa, Totales Reales y Metrajes por Ítem empiezan después (10, 11, 12, 13, 14, 15, 16, 17)
                 ];

                 const cenefaSummaryRow = [...cenefaRowBase]; // Copia los detalles comunes
                 // Agrega Pleno, Descuento Angular (vacíos), Largo Cenefa, Totales Reales (vacíos), Metraje Área (vacío), Metraje Lineal.
                 cenefaSummaryRow.push(
                     '', // 10: Pleno (vacío)
                     '', // 11: Descuento Angular (vacío)
                     !isNaN(item.cenefaLength) ? item.cenefaLength.toFixed(2) : '', // 12: Largo Total (Cenefa) (Real)
                     '', // 13: Suma Perímetros Segmentos (Real) (vacío)
                     '', // 14: Ancho Total (Muro Real) (vacío)
                     '', // 15: Área Total (Real) (vacío)
                     '', // 16: Metraje (Área Muro/Cielo) (vacío)
                     !isNaN(item.metrajeLinear) ? item.metrajeLinear.toFixed(2) : '' // 17: Metraje (Lineal Cenefa) (NUEVO)
                 );
                 sheetData.push(cenefaSummaryRow);

                 // Cenefas no tienen segmentos, así que no hay sección de segmentos detallada.
            }
       });
       sheetData.push([]); // Fila en blanco para espaciar

   } else {
        console.log("No hay ítems válidos para mostrar en el detalle de ítems calculados en Excel.");
        // Mostrar un mensaje si no hay ítems válidos calculados pero sí hay totales de metraje (ej: solo cenefas válidas)
        if (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] > 0 ||
            lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0 ||
            lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {
             sheetData.push(['']); // Fila en blanco
             sheetData.push(['No hay ítems de Muro o Cielo con dimensiones válidas para el cálculo detallado de materiales, pero se calcularon metrajes totales.']);
             sheetData.push(['']); // Fila en blanco
        }
   }


    // --- NUEVA SECCIÓN: Totales de Metraje en Excel ---
    console.log("Añadiendo sección de totales de metraje al Excel.");
    // Solo añade la sección si hay algún metraje total positivo
    if (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] > 0 ||
        lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0 ||
        lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {

         sheetData.push(["Totales de Metraje:"]);
         // Agrega las filas de metraje total solo si son positivas
         if (lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'] > 0) {
             sheetData.push(['Muro Área Total', lastCalculatedTotalMetrajes['Muro Área Total Metraje (m²)'].toFixed(2), 'm²']);
         }
         if (lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'] > 0) {
              sheetData.push(['Cielo Área Total', lastCalculatedTotalMetrajes['Cielo Área Total Metraje (m²)'].toFixed(2), 'm²']);
         }
          if (lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'] > 0) {
              sheetData.push(['Cenefa Lineal Total', lastCalculatedTotalMetrajes['Cenefa Lineal Total Metraje (m)'].toFixed(2), 'm']);
          }
         sheetData.push([]); // Fila en blanco para espaciar
    } else {
         console.log("No hay totales de metraje positivos para añadir sección al Excel.");
    }
    // --- Fin NUEVA SECCIÓN Metraje en Excel ---


   // Tabla de Totales de Materiales
    // Solo añade la tabla si hay materiales con cantidad > 0
    if (Object.keys(lastCalculatedTotalMaterials).some(material => lastCalculatedTotalMaterials[material] > 0)) {
        console.log("Generando tabla de materiales totales al Excel.");
         sheetData.push(["Totales de Materiales (Cantidades a Comprar):"]); // Título actualizado
         sheetData.push(["Material", "Cantidad", "Unidad"]);

         const sortedMaterials = Object.keys(lastCalculatedTotalMaterials).sort();
         sortedMaterials.forEach(material => {
             const cantidad = lastCalculatedTotalMaterials[material];
             if (cantidad > 0) { // Solo añade si la cantidad es positiva
                 const unidad = getMaterialUnit(material);
                 // Usa el nombre del material directamente de la clave
                  sheetData.push([material, cantidad, unidad]);
             }
         });
         sheetData.push([]); // Fila en blanco al final de la tabla
    } else {
         console.log("No hay materiales con cantidad positiva para añadir tabla al Excel.");
    }


   // Crea un libro y una hoja de Excel
   const wb = XLSX.utils.book_new();
   // `aoa_to_sheet` convierte un array de arrays (sheetData) a una hoja
   const ws = XLSX.utils.aoa_to_sheet(sheetData);

   // Opcional: Agregar estilos básicos (encabezados en negrita, etc.)
   // XLSX.js básico tiene limitaciones de estilo. Para estilos avanzados, se necesita otra librería.
   // Aquí no se agregan estilos complejos para mantener la simplicidad.
   // Sin embargo, podemos ajustar anchos de columna estimados.
    const col_widths = [];
    // Asegura que haya suficientes entradas en col_widths para todas las columnas antes de estimar
    const maxCols = Math.max(...sheetData.map(row => row.length));
    for(let i = 0; i < maxCols; i++) {
        col_widths[i] = 0; // Initialize widths to 0
    }

    sheetData.forEach(row => {
        row.forEach((cell, colIndex) => {
            // Estimar ancho basado en la longitud del texto. Ajustar factor según se vea en Excel.
            const cellLength = cell ? cell.toString().length : 0;
            // Si la columna ya tiene un ancho, usa el máximo entre el actual y el necesario para esta celda.
            // Usa un factor de 1.2 para estimar el ancho real en Excel
            col_widths[colIndex] = Math.max(col_widths[colIndex] || 0, cellLength * 1.2); // +2 for padding? Maybe not needed with factor
        });
    });
    // Algunos ajustes manuales para columnas específicas si es necesario.
    // Por ejemplo, la columna de "Detalle/Dimensiones" (índice 2) podría necesitar más ancho.
    if (col_widths[2]) col_widths[2] = Math.max(col_widths[2], 30); // Ajusta 30 si es necesario
     // Columna de Materiales en la tabla de totales (puede variar su índice dependiendo de si hay sección de detalle de ítems)
     // Para la tabla de materiales totales, las columnas son [Material, Cantidad, Unidad]. Material es índice 0 dentro de esa tabla.
     // Pero en el sheetData general, si hay detalle de ítems, la tabla de materiales empieza después.
     // Es más seguro ajustar anchos para columnas por su contenido típico más que por índice fijo si la estructura varía.
     // Sin embargo, si la tabla de detalle de ítems tiene un número fijo de columnas (18 en nuestro caso),
     // la tabla de materiales (3 columnas) empezaría después.
     // El índice de la columna Materiales sería 18. Cantidad 19, Unidad 20.
     // Ajustar ancho para la columna Materiales (índice 18 si la tabla de detalle de ítems tiene 18 cols + 3 cols de tabla de materiales)
     // Vamos a asumir que la tabla de detalle de ítems tiene 18 columnas (0 a 17)
      const MATERIAL_COLUMN_INDEX_IN_SHEET = 18; // Assuming detail table has 18 columns
      if (col_widths[MATERIAL_COLUMN_INDEX_IN_SHEET]) col_widths[MATERIAL_COLUMN_INDEX_IN_SHEET] = Math.max(col_widths[MATERIAL_COLUMN_INDEX_IN_SHEET], 30);


    // Asigna los anchos estimados a la hoja
    ws['!cols'] = col_widths.map(w => ({ wch: w }));


   // Agrega la hoja al libro
   XLSX.utils.book_append_sheet(wb, ws, "CalculoMateriales"); // "CalculoMateriales" es el nombre de la hoja

   // Genera y guarda el archivo Excel
   XLSX.writeFile(wb, `Calculo_Materiales_${new Date().toLocaleDateString('es-ES').replace(/\//g, '-')}.xlsx`); // Nombre del archivo con fecha
   console.log("Excel generado.");
};


// --- Event Listeners ---
// Asigna las funciones a los eventos de los botones
if (addItemBtn) addItemBtn.addEventListener('click', createItemBlock); // Botón "Agregar Muro, Cielo o Cenefa"
if (calculateBtn) calculateBtn.addEventListener('click', calculateMaterials); // Botón "Calcular Materiales y Metrajes"
if (generatePdfBtn) generatePdfBtn.addEventListener('click', generatePDF); // Botón "Generar PDF"
if (generateExcelBtn) generateExcelBtn.addEventListener('click', generateExcel); // Botón "Generar Excel"

// --- Configuración Inicial ---
// Agrega un ítem (Muro por defecto) al cargar la página
createItemBlock();
// Establece el estado inicial del botón de cálculo (habilitado si hay ítems)
toggleCalculateButtonState();


}); // Fin del evento DOMContentLoaded. Asegura que el script se ejecuta después de que la página esté completamente cargada.

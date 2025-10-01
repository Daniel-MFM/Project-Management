<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Reddit+Sans+Condensed:wght@400;700&display=swap" rel="stylesheet">

  <style>
     /* Base styles */
     body { padding: 10px; font-family: 'Reddit Sans Condensed', sans-serif; font-size: 14px; }
     body, p, span, a, button, h3, div, table, td { font-family: 'Reddit Sans Condensed', sans-serif; }
     h3 { font-size: 15px; font-weight: bold; margin-top: 15px; margin-bottom: 8px; border-bottom: 1px solid #ccc; padding-bottom: 4px; }

     /* --- Button Layout & Styling --- */
     .view-buttons, .report-buttons, .format-buttons, .quote-buttons { padding: 0 2px; }

     /* Base styles for ALL section buttons */
     .view-buttons button, .report-buttons button, .format-buttons button, .quote-buttons button {
       display: flex; align-items: center; justify-content: center;
       text-transform: uppercase;
       border: 1px solid transparent;
       box-sizing: border-box;
       padding: 7px 5px;
       width: 95%;
       margin: 8px auto;
       cursor: pointer;
     }
     /* Specific layout class for the summary button */
     .summary-button-layout {
       margin-bottom: 12px;
       font-size: 16px;
       font-weight: bold;
     }
     /* Container for button pairs */
     .button-pair {
       display: flex; justify-content: space-between;
       margin: 8px auto;
       width: 95%;
     }
     /* Styling for buttons within pairs */
     .button-pair button {
       width: 48%; margin: 0; font-size: 13px; padding: 7px 5px;
     }

     /* Cog button styling */
    .view-button-container { display: flex; align-items: center; width: 95%; margin: 8px auto; }
    .view-button-container button.blue { flex-grow: 1; margin: 0; }
    .cog-button {
      width: 30px !important;
      min-width: 30px !important;
      height: 30px;
      padding: 5px !important;
      margin-left: 8px !important;
      font-size: 16px;
      line-height: 1;
      border: 1px solid #ccc !important;
      background-color: #f8f9fa;
      color: #3c4043;
      cursor: pointer;
    }
    .cog-button:hover { background-color: #e9ecef; }

     /* Info section styling */
     .info-section { margin-top: 15px; }
     .info-section p.planta-line { margin: 5px 0 10px 0; font-size: 13px; line-height: 1.5; word-wrap: break-word; }
     .info-section p.planta-line .label { font-weight: bold; }
     .info-section p.planta-line a { text-decoration: none; color: #4285f4; cursor: pointer; }
     .info-section p.planta-line a:hover { text-decoration: underline; }
     .info-section table { width: 100%; border-collapse: collapse; margin-top: 5px; }
     .info-section td { padding: 4px 0; vertical-align: top; font-size: 13px; line-height: 1.4; }
     .info-section .label-cell { font-weight: bold; width: 65px; padding-right: 8px; white-space: nowrap; }
     .info-section .data-cell { word-wrap: break-word; }

     /* Loader and status/error styling */
     .loader, .error-message, #pdf-status, #format-status, #quote-status {
         font-size: 12px;
         margin-top: 5px;
         padding-left: 5px;
         min-height: 1em;
     }
     .loader { font-style: italic; color: #888; display: none; }
     .error-message { color: red; font-weight: bold; display: none; }
     #pdf-status a, #quote-status a { color: #188038; font-weight: bold; }
     #pdf-status span, #format-status span, #quote-status span {
         display: block; margin-top: 3px;
     }
     .success { color: #188038; }
     .warning { color: #dd8300; }
     .error { color: red; }

  </style>
</head>
<body>
  <h3>VISUALIZAÇÕES</h3>
  <div class="view-buttons">
    <div class="view-button-container">
      <button class="blue summary-button-layout" onclick="callViewFunction('SetViewResumo')">Sumário</button>
      <button class="cog-button" data-view-name="Resumo" title="Personalizar Vista Sumário">⚙️</button>
    </div>
    <div class="button-pair">
      <div class="view-button-container" style="width:48%">
        <button class="blue" onclick="callViewFunction('SetViewDefault')">Normal</button>
        <button class="cog-button" data-view-name="Default" title="Personalizar Vista Normal">⚙️</button>
      </div>
      <div class="view-button-container" style="width:48%">
        <button class="blue" onclick="callViewFunction('SetViewBudget')">Orçamento</button>
        <button class="cog-button" data-view-name="Budget" title="Personalizar Vista Orçamento">⚙️</button>
      </div>
    </div>
    <div class="button-pair">
      <div class="view-button-container" style="width:48%">
        <button class="blue" onclick="callViewFunction('SetViewClient')">Cliente</button>
        <button class="cog-button" data-view-name="Client" title="Personalizar Vista Cliente">⚙️</button>
      </div>
      <div class="view-button-container" style="width:48%">
        <button class="blue" onclick="callViewFunction('setViewLayout')">Layout</button>
        <button class="cog-button" data-view-name="Layout" title="Personalizar Vista Layout">⚙️</button>
      </div>
    </div>
  </div>

  <h3>ORÇAMENTO</h3>
  <div class="quote-buttons">
      <button onclick="triggerQuotePreparation()">Preparar Orçamento</button>
      <button onclick="triggerQuotePdfGeneration()">Gerar PDF Orçamento</button>
      <div id="quote-status"></div>
  </div>
  <h3>RELATÓRIOS</h3>
  <div class="report-buttons">
      <button onclick="triggerSummaryPdfGeneration()">Gerar PDF Sumário</button>
      <div id="pdf-status"></div>
  </div>

  <h3>FORMATAÇÃO</h3>
  <div class="format-buttons">
      <button onclick="triggerUppercaseConversion()">Texto Maiúsculas (Folha Atual)</button>
      <div id="format-status"></div>
  </div>

  <h3>INFORMAÇÃO</h3>
  <div class="info-section" id="info-section-container">
     <p class="planta-line"><span class="label">Planta Geral:</span> <a id="floorplan-link" data-field="floorplanUrl" href="#" target="_blank">A carregar link...</a></p>
     <table>
       <tbody>
         <tr><td class="label-cell">Cliente:</td><td class="data-cell"><span id="client-name" data-field="clientName">A carregar...</span></td></tr>
         <tr><td class="label-cell">Morada:</td><td class="data-cell"><span id="client-address" data-field="clientAddress">A carregar...</span></td></tr>
         <tr><td class="label-cell">NIF:</td><td class="data-cell"><span id="client-nif" data-field="clientNif">A carregar...</span></td></tr>
         <tr><td class="label-cell">Email:</td><td class="data-cell"><span id="client-email" data-field="clientEmail">A carregar...</span></td></tr>
       </tbody>
     </table>
     <div id="info-loader" class="loader">A carregar informação...</div>
     <div id="info-error" class="error-message"></div>
  </div>

  <div id="action-loader" class="loader" style="margin-top: 10px;">A processar...</div>

  <script>
    const DEBUG_MODE = false;

    // --- Utility Functions ---

    /** Shows the main processing loader. */
    function showLoader(loaderId = 'action-loader') {
      const loader = document.getElementById(loaderId);
      if (loader) loader.style.display = 'block';
    }

    /** Hides all loaders. */
    function hideLoaders() {
      ['info-loader', 'action-loader'].forEach(id => {
        const loader = document.getElementById(id);
        if (loader) loader.style.display = 'none';
      });
    }

    /** Clears all status message areas. */
    function clearAllStatusMessages() {
      ['info-error', 'pdf-status', 'format-status', 'quote-status'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
          el.innerHTML = '';
          if (id === 'info-error') el.style.display = 'none';
        }
      });
    }

    /** Displays a status message in a specific status area. */
    function showStatusMessage(elementId, result) {
        hideLoaders();
        const statusArea = document.getElementById(elementId);
        if (!statusArea || !result || typeof result !== 'string') return;

        let statusClass = 'success';
        let message = result;

        if (result.toLowerCase().startsWith('http')) { // PDF link
            const linkText = (elementId === 'pdf-status') ? 'PDF Sumário Gerado! Clique para abrir.' : 'PDF Orçamento Gerado! Clique para abrir.';
            statusArea.innerHTML = `<a href="${result}" target="_blank">${linkText}</a>`;
            return;
        } else if (result.toLowerCase().startsWith('erro:') || result.toLowerCase().startsWith('erro ao')) { // Error
            statusClass = 'error';
            message = result;
        } else if (result.includes("Nenhum texto encontrado") || result.includes("Nenhuma linha de dados encontrada")) { // Warning/Info
             statusClass = 'warning';
        }

        statusArea.innerHTML = `<span class="${statusClass}">${message}</span>`;
    }

    /** Handles Apps Script call failures. */
    function onFailure(error, statusElementId = null) {
      if (DEBUG_MODE) console.error("Apps Script call failed:", error);
      hideLoaders();
      const message = 'Ocorreu um erro: ' + (error.message || 'Erro desconhecido.');
      if (statusElementId) {
          const statusArea = document.getElementById(statusElementId);
          if (statusArea) {
              statusArea.innerHTML = `<span class="error">${message}</span>`;
          } else {
              // Fallback if status area ID is invalid
              console.error(`Status area #${statusElementId} not found.`);
              alert(message);
          }
      } else {
        alert(message);
      }
    }

    /** Clears info placeholders and optionally sets error state. */
    function clearInfoPlaceholders(isError = false) {
       const errorText = isError ? 'Erro' : '...';
       const floorplanLink = document.getElementById('floorplan-link');

       floorplanLink.innerText = isError ? 'Erro' : 'A carregar...';
       floorplanLink.href = '#';

       document.getElementById('client-name').innerText = errorText;
       document.getElementById('client-address').innerText = errorText;
       document.getElementById('client-nif').innerText = errorText;
       document.getElementById('client-email').innerText = errorText;
    }

    // --- Generic & Specific Action Callers ---

    /** Generic caller for view switching functions. */
    function callViewFunction(functionName) {
      if (DEBUG_MODE) console.log(`Calling view function: ${functionName}`);
      showLoader('action-loader');
      clearAllStatusMessages();
      google.script.run
        .withSuccessHandler(hideLoaders)
        .withFailureHandler(onFailure)
        [functionName]();
    }

    /** Caller for functions that return status/URL, updates a specific element. */
    function callWithStatusUpdate(functionName, statusElementId, statusVerb) {
        if (DEBUG_MODE) console.log(`Calling status function: ${functionName}`);
        showLoader('action-loader');
        clearAllStatusMessages();
        const statusArea = document.getElementById(statusElementId);
        if (statusArea) statusArea.innerHTML = `<span class="loader">${statusVerb}...</span>`;

        google.script.run
          .withSuccessHandler(result => showStatusMessage(statusElementId, result))
          .withFailureHandler(error => onFailure(error, statusElementId))
          [functionName]();
    }

    function triggerSummaryPdfGeneration() {
        callWithStatusUpdate('generateSummaryPdf', 'pdf-status', 'A gerar PDF Sumário');
    }
    function triggerQuotePdfGeneration() {
        callWithStatusUpdate('generateQuotePdf', 'quote-status', 'A gerar PDF Orçamento');
    }
    function triggerUppercaseConversion() {
        callWithStatusUpdate('convertSheetToUppercase', 'format-status', 'A converter texto');
    }
    function triggerQuotePreparation() {
        callWithStatusUpdate('prepareQuoteSheet', 'quote-status', 'A preparar orçamento');
    }

    // --- Info Section Handlers ---

    /** Initiates loading of project info on sidebar load. */
    function loadInitialData() {
      showLoader('info-loader');
      clearAllStatusMessages();
      clearInfoPlaceholders();

      google.script.run
        .withSuccessHandler(updateInfoSection)
        .withFailureHandler(onInfoFailure)
        .getProjectInfo();
    }

    /** Updates the DOM with received project info. */
    function updateInfoSection(projectInfo) {
      hideLoaders();

      if (projectInfo && !projectInfo.error) {
        const floorplanLink = document.getElementById('floorplan-link');
        floorplanLink.href = projectInfo.floorplanUrl || '#';
        floorplanLink.innerText = (projectInfo.floorplanUrl && projectInfo.floorplanUrl !== '#') ? 'Abrir Documento' : 'N/A';

        document.getElementById('client-name').innerText = projectInfo.clientName || 'N/A';
        document.getElementById('client-address').innerText = projectInfo.clientAddress || 'N/A';
        document.getElementById('client-nif').innerText = projectInfo.clientNif || 'N/A';
        document.getElementById('client-email').innerText = projectInfo.clientEmail || 'N/A';
      } else {
         onInfoFailure(projectInfo || { message: 'Dados recebidos inválidos do script.' });
      }
    }

    /** Handles failure during project info loading. */
    function onInfoFailure(error) {
      hideLoaders();
      const errorDiv = document.getElementById('info-error');
      errorDiv.innerText = 'Erro Info: ' + (error.message || 'Detalhes indisponíveis.');
      errorDiv.style.display = 'block';
      clearInfoPlaceholders(true);
    }

    /** Opens the Edit Dialog on double-click of an info field. */
    function openEditDialog(event) {
        // Find the closest element with a data-field attribute
        const targetElement = event.target.closest('[data-field]');
        if (!targetElement) return;

        const fieldName = targetElement.dataset.field;
        // Use innerText for <span>s and href for <a> (floorplan link)
        const currentValue = targetElement.tagName === 'A' ? targetElement.href : targetElement.innerText;

        if (fieldName && currentValue !== 'N/A' && currentValue !== 'A carregar...' && currentValue !== 'Erro') {
            if (DEBUG_MODE) console.log(`Opening edit dialog for ${fieldName}`);
            showLoader('action-loader');
            google.script.run
                .withSuccessHandler(hideLoaders)
                .withFailureHandler(onFailure)
                .showEditInfoFieldModal(fieldName, currentValue === 'Abrir Documento' ? '' : currentValue); // Pass empty string if link text is just "Abrir Documento"
        }
    }

    // --- Initialization and Listeners ---

     window.addEventListener('load', loadInitialData);

    // Attach double-click listeners for editing info fields
    document.querySelectorAll('#info-section-container [data-field]').forEach(element => {
        element.addEventListener('dblclick', openEditDialog);
    });

    // Add event listeners to all cog buttons to open the MODAL dialog
    document.querySelectorAll('.cog-button').forEach(button => {
      button.addEventListener('click', event => {
        const viewName = event.currentTarget.dataset.viewName;
        showLoader('action-loader');
        google.script.run
          .withSuccessHandler(hideLoaders)
          .withFailureHandler(onFailure)
          .showColumnChooserModal(viewName);
      });
    });
  </script>
</body>
</html>

/**
 * @OnlyCurrentDoc
 */

// --- Configuration ---
const CONFIG = {
  SHEET_METADATA: "METADATA (NAO MEXER)",
  SHEET_MAIN_PANEL: "PAINEL PRINCIPAL",
  SHEET_QUOTE: "ORÇAMENTO",
  PDF_FOLDER_NAME: "Relatorios PDF Gerados",
  HEADER_ROW: 1,
  FONT_FAMILY: "Reddit Sans Condensed",
  VIEW_STATE_PREFIX: "customViewState_",

  // Mapping client-side field names to D2:D6 cells in the METADATA sheet
  INFO_CELL_MAP: {
    clientName: "D2",
    clientAddress: "D3",
    clientNif: "D4",
    clientEmail: "D5",
    floorplanUrl: "D6"
  },

  // Quote Sheet Specifics
  QUOTE_CELLS: {
    LAST_NUMBER: "E7", // Cell on METADATA sheet holding the last quote number
    TARGET_NUMBER: "I2", // Cell on QUOTE sheet for the quote number display
    TARGET_DATE: "I3", // Cell on QUOTE sheet for the date
    TARGET_NAME: "F4", // Cell on QUOTE sheet for Client Name
    TARGET_ADDR1: "F5", // Cell on QUOTE sheet for Address Line 1
    TARGET_ADDR2: "F6", // Cell on QUOTE sheet for Address Line 2
    TARGET_NIF: "F7", // Cell on QUOTE sheet for NIF
  },

  // Column definitions for predefined views (using column letters)
  VIEW_COLUMN_MAP: {
    Resumo: ["A", "B", "C", "D", "E", "F", "H", "N", "P", "Q", "U", "V"],
    Default: ["A", "B", "D", "E", "F", "H", "I", "L", "Q", "U", "V"],
    Budget: ["A", "B", "C", "D", "E", "F", "G", "I", "J", "L", "M", "N", "O", "P"],
    Client: ["A", "B", "C", "D", "E", "F", "H", "P"],
    Layout: ["A", "B", "C", "D"]
  }
};

// --- Global Styles for Rich Text Header Formatting ---
const HEADER_STYLE_PRIMARY = SpreadsheetApp.newTextStyle()
  .setBold(true)
  .setFontSize(11)
  .setFontFamily(CONFIG.FONT_FAMILY)
  .build();

const HEADER_STYLE_SECONDARY = SpreadsheetApp.newTextStyle()
  .setBold(false)
  .setFontSize(9)
  .setFontFamily(CONFIG.FONT_FAMILY)
  .build();

// --- UI Management ---

/**
 * Adds a custom menu "Painel de Controlo" when the spreadsheet opens.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Painel de Controlo")
      .addItem("Mostrar Painel", "showSidebar")
      .addToUi();
  } catch (e) {
    Logger.log(`Error adding custom menu: ${e.message}`);
  }
}

/**
 * Displays the main sidebar UI from 'Sidebar.html'.
 */
function showSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Painel de Controlo')
      .setWidth(240);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.log(`Error in showSidebar: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Não foi possível abrir o Painel de Controlo. Erro: ${error.message}`);
  }
}

/**
 * Displays a modal dialog asking the user to open the main sidebar. (Currently unused).
 */
function showSidebarPromptDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('OpenSidebarPrompt')
      .setWidth(300)
      .setHeight(120);
    SpreadsheetApp.getUi().showModalDialog(html, 'Abrir Painel de Controlo?');
  } catch (e) {
    Logger.log(`Error showing prompt dialog: ${e.message}`);
    SpreadsheetApp.getUi().alert(`Erro ao tentar abrir o diálogo do painel: ${e.message}`);
  }
}

/**
 * Shows the Column Chooser modal dialog for a given view name.
 * @param {string} viewName The name of the view context (e.g., "Default", "Resumo").
 */
function showColumnChooserModal(viewName) {
  if (!viewName) {
    SpreadsheetApp.getUi().alert("Erro: Nome da vista não especificado para o personalizador de colunas.");
    return;
  }
  try {
    const template = HtmlService.createTemplateFromFile("ColumnChooserDialog");
    template.viewName = viewName;

    const htmlOutput = template.evaluate()
      .setWidth(380)
      .setHeight(580);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Personalizar Vista: ${viewName}`);
  } catch (e) {
    Logger.log(`Error in showColumnChooserModal: ${e.message}`);
    SpreadsheetApp.getUi().alert(`Erro ao tentar mostrar o diálogo de personalização de colunas: ${e.message}`);
  }
}

/**
 * Shows a modal dialog to edit a specific project information field.
 * @param {string} fieldName The key name of the field to be edited (e.g., "clientName").
 * @param {string} currentValue The current value of the field.
 */
function showEditInfoFieldModal(fieldName, currentValue) {
  if (!fieldName || !CONFIG.INFO_CELL_MAP.hasOwnProperty(fieldName) || typeof currentValue === 'undefined') {
    SpreadsheetApp.getUi().alert(`Erro: Não é possível editar o campo '${fieldName}'.`);
    return;
  }
  try {
    const template = HtmlService.createTemplateFromFile("EditInfoFieldDialog");
    template.fieldName = fieldName;
    template.currentValue = currentValue;

    let dialogTitle = "Editar Campo";
    switch (fieldName) {
      case "clientName":
        dialogTitle = "Editar Nome do Cliente";
        break;
      case "clientAddress":
        dialogTitle = "Editar Morada do Cliente";
        break;
      case "clientNif":
        dialogTitle = "Editar NIF do Cliente";
        break;
      case "clientEmail":
        dialogTitle = "Editar Email do Cliente";
        break;
      case "floorplanUrl":
        dialogTitle = "Editar Link da Planta";
        break;
    }

    const htmlOutput = template.evaluate()
      .setWidth(350)
      .setHeight(280);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
  } catch (e) {
    Logger.log(`Error in showEditInfoFieldModal: ${e.message}`);
    SpreadsheetApp.getUi().alert(`Erro ao tentar mostrar o diálogo de edição: ${e.message}`);
  }
}

// --- Project Data Retrieval & Update ---

/**
 * Gets project info from the metadata sheet for display in the sidebar.
 * @return {object} Object with project info or an error object.
 */
function getProjectInfo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG.SHEET_METADATA);
    if (!configSheet) {
      throw new Error(`A folha '${CONFIG.SHEET_METADATA}' não foi encontrada.`);
    }

    // Read values D2:D6 (Mapped to client info fields)
    const rangeNotation = Object.values(CONFIG.INFO_CELL_MAP).join(",");
    const infoValues = configSheet.getRangeList(Object.values(CONFIG.INFO_CELL_MAP)).getRanges().map(r => r.getValue());

    const projectInfo = {
      clientName: infoValues[0] || 'N/A',
      clientAddress: infoValues[1] || 'N/A',
      clientNif: infoValues[2] || 'N/A',
      clientEmail: infoValues[3] || 'N/A',
      floorplanUrl: infoValues[4] || '#'
    };

    return projectInfo;
  } catch (error) {
    Logger.log(`getProjectInfo: Error: ${error.message}`);
    return {
      error: true,
      message: `Erro ao ler '${CONFIG.SHEET_METADATA}': ${error.message}`
    };
  }
}

/**
 * Updates a specific project information field in the metadata sheet.
 * @param {string} fieldName The key name of the field to update.
 * @param {string} newValue The new value for the field.
 * @return {Object} An object indicating success or failure, including a message.
 */
function updateProjectInfoField(fieldName, newValue) {
  try {
    if (!CONFIG.INFO_CELL_MAP.hasOwnProperty(fieldName)) {
      throw new Error("Nome de campo inválido fornecido para atualização.");
    }

    const cellNotation = CONFIG.INFO_CELL_MAP[fieldName];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG.SHEET_METADATA);

    if (!configSheet) {
      throw new Error(`A folha de configuração '${CONFIG.SHEET_METADATA}' não foi encontrada.`);
    }

    configSheet.getRange(cellNotation).setValue(newValue);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: `Campo atualizado com sucesso para: '${newValue}'.`,
      updatedField: fieldName,
      value: newValue
    };

  } catch (error) {
    Logger.log(`Error in updateProjectInfoField: ${error.message}`);
    return {
      success: false,
      message: error.message
    };
  }
}

// --- View Switching and Customization ---

/**
 * Converts column letter(s) (e.g., "A", "AA") to its 1-based numeric index.
 * @param {string} letter The column letter(s).
 * @return {number} The 1-based column index (e.g., A=1, AA=27).
 */
function columnLetterToNumber(letter) {
  let num = 0;
  const letters = letter.toUpperCase();
  for (let i = 0; i < letters.length; i++) {
    num = num * 26 + (letters.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return num;
}

/**
 * Applies a view by showing/hiding columns on the main data sheet.
 * @param {Array<string>} allowedColumnsArray Array of column letters to show.
 * @return {string} A success message.
 */
function setColumnView(allowedColumnsArray) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_MAIN_PANEL) {
    SpreadsheetApp.getUi().alert(`As funções de visualização só podem ser usadas na folha "${CONFIG.SHEET_MAIN_PANEL}".`);
    return "N/A"; // Return early, alert handles user notification
  }
  try {
    const maxCols = sheet.getMaxColumns();
    const allowedNumsSet = new Set(allowedColumnsArray.map(columnLetterToNumber));

    // Reset: Show all columns initially
    sheet.showColumns(1, maxCols);
    SpreadsheetApp.flush();

    // Hide unallowed columns in contiguous blocks
    let startHideRange = null;
    for (let i = 1; i <= maxCols; i++) {
      if (!allowedNumsSet.has(i)) {
        if (startHideRange === null) {
          startHideRange = i;
        }
      } else {
        if (startHideRange !== null) {
          sheet.hideColumns(startHideRange, i - startHideRange);
          startHideRange = null;
        }
      }
    }
    // Handle remaining tail end
    if (startHideRange !== null) {
      sheet.hideColumns(startHideRange, maxCols - startHideRange + 1);
    }

    SpreadsheetApp.flush();
    return "Vista de colunas aplicada com sucesso.";
  } catch (error) {
    Logger.log(`Error in setColumnView: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Erro ao definir a visão: ${error.message}`);
    return `Erro: ${error.message}`;
  }
}

// Dynamic View Switching functions
function SetViewResumo() {
  return setColumnView(CONFIG.VIEW_COLUMN_MAP.Resumo);
}
function SetViewDefault() {
  return setColumnView(CONFIG.VIEW_COLUMN_MAP.Default);
}
function SetViewBudget() {
  return setColumnView(CONFIG.VIEW_COLUMN_MAP.Budget);
}
function SetViewClient() {
  return setColumnView(CONFIG.VIEW_COLUMN_MAP.Client);
}
function setViewLayout() {
  return setColumnView(CONFIG.VIEW_COLUMN_MAP.Layout);
}

/**
 * Gets all header texts from the HEADER_ROW of the main data sheet.
 * @param {string} viewName The name of the view (for logging).
 * @return {Array<string>} An array of header strings.
 */
function getViewHeaders(viewName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MAIN_PANEL);
    if (!sheet) {
      throw new Error(`Sheet '${CONFIG.SHEET_MAIN_PANEL}' not found.`);
    }

    const maxCols = sheet.getMaxColumns();
    const headerValues = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, maxCols).getValues()[0];

    const filteredHeaders = headerValues.filter(header => typeof header === 'string' && header.trim() !== '');

    return filteredHeaders;
  } catch (error) {
    Logger.log(`Error in getViewHeaders: ${error.message}`);
    throw new Error(`Erro ao obter cabeçalhos para '${viewName}': ${error.message}`);
  }
}

/**
 * Gets the current visibility state of all columns in the main data sheet.
 * @param {string} viewName The name of the view (for logging).
 * @return {Object<string, boolean>} Object mapping header names to visibility state.
 */
function getCurrentColumnVisibility(viewName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MAIN_PANEL);
    if (!sheet) {
      throw new Error(`Sheet '${CONFIG.SHEET_MAIN_PANEL}' not found.`);
    }

    const maxCols = sheet.getMaxColumns();
    const headerRowValues = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, maxCols).getValues()[0];
    const visibilityStates = {};

    for (let i = 0; i < maxCols; i++) {
      const headerText = (typeof headerRowValues[i] === 'string') ? headerRowValues[i].trim() : '';
      if (headerText !== '') {
        const columnIndex = i + 1;
        visibilityStates[headerText] = !sheet.isColumnHiddenByUser(columnIndex);
      }
    }
    return visibilityStates;
  } catch (error) {
    Logger.log(`Error in getCurrentColumnVisibility: ${error.message}`);
    throw new Error(`Erro ao obter o estado de visibilidade das colunas para '${viewName}': ${error.message}`);
  }
}

/**
 * Applies column visibility settings to the main data sheet.
 * @param {string} viewName The name of the view being applied (for logging).
 * @param {Object} visibilitySettings Object mapping header names to boolean visibility state.
 * @return {string} A success message.
 */
function applyColumnVisibility(viewName, visibilitySettings) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MAIN_PANEL);
    if (!sheet) {
      throw new Error(`Sheet '${CONFIG.SHEET_MAIN_PANEL}' not found.`);
    }

    const maxCols = sheet.getMaxColumns();
    const headerRowValues = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, maxCols).getValues()[0];

    sheet.showColumns(1, maxCols);
    SpreadsheetApp.flush();

    for (let i = 0; i < maxCols; i++) {
      const columnIndex = i + 1;
      const headerText = (typeof headerRowValues[i] === 'string') ? headerRowValues[i].trim() : '';

      if (headerText === '') continue;

      const isVisible = visibilitySettings.hasOwnProperty(headerText) && visibilitySettings[headerText] === true;

      if (!isVisible) {
        sheet.hideColumns(columnIndex);
      }
    }

    SpreadsheetApp.flush();
    return `Visibilidade das colunas atualizada para a visão '${viewName}'.`;
  } catch (error) {
    Logger.log(`Error in applyColumnVisibility: ${error.message}`);
    throw new Error(`Erro ao aplicar visibilidade de colunas para '${viewName}': ${error.message}`);
  }
}

/**
 * Applies filters to the main data sheet based on the provided settings.
 * @param {string} viewName The name of the view being filtered (for logging).
 * @param {Object} filterSettings Object mapping header names to filter values (exact match).
 * @return {string} A success or informational message.
 */
function applyViewFilters(viewName, filterSettings) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MAIN_PANEL);
    if (!sheet) {
      throw new Error(`Sheet '${CONFIG.SHEET_MAIN_PANEL}' not found.`);
    }

    // Remove existing filter
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }

    // If no filters are provided, we are done
    if (!filterSettings || Object.keys(filterSettings).length === 0) {
      return "Filtros removidos ou nenhum filtro aplicado.";
    }

    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getMaxColumns()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => {
      if (typeof h === 'string' && h.trim() !== '') {
        headerMap[h.trim()] = i + 1; // 1-based index
      }
    });

    const columnCriteria = {};

    for (const headerName in filterSettings) {
      if (filterSettings.hasOwnProperty(headerName)) {
        const filterValue = filterSettings[headerName];
        const columnIndex = headerMap[headerName];

        if (columnIndex && typeof filterValue === 'string' && filterValue !== '') {
          const criterion = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(filterValue).build();
          columnCriteria[columnIndex] = criterion;
        }
      }
    }

    if (Object.keys(columnCriteria).length > 0) {
      // The filter range should exclude the header row if filtering by value (which it does)
      const filterRange = sheet.getRange(CONFIG.HEADER_ROW, 1, sheet.getMaxRows() - CONFIG.HEADER_ROW + 1, sheet.getMaxColumns());
      const newFilter = filterRange.createFilter();

      for (const colIdxStr in columnCriteria) {
        if (columnCriteria.hasOwnProperty(colIdxStr)) {
          const colIdx = parseInt(colIdxStr);
          newFilter.setColumnFilterCriteria(colIdx, columnCriteria[colIdx]);
        }
      }
      return `Filtros aplicados com sucesso para a visão '${viewName}'.`;
    } else {
      return "Nenhum critério de filtro válido foi aplicado.";
    }

  } catch (error) {
    Logger.log(`Error in applyViewFilters: ${error.message}`);
    throw new Error(`Erro ao aplicar filtros para a visão '${viewName}': ${error.message}`);
  }
}

/**
 * Retrieves unique, sorted, non-empty string values from a specified column in the main data sheet.
 * @param {string} headerName The text of the header for the column to process.
 * @param {string} viewName For logging/consistency.
 * @return {Array<string>} A sorted array of unique string values.
 */
function getColumnUniqueValues(headerName, viewName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MAIN_PANEL);
    if (!sheet) {
      throw new Error(`Sheet '${CONFIG.SHEET_MAIN_PANEL}' not found.`);
    }

    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getMaxColumns()).getValues()[0];
    let columnIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      if (typeof headers[i] === 'string' && headers[i].trim() === headerName) {
        columnIndex = i + 1;
        break;
      }
    }

    if (columnIndex === -1) {
      throw new Error(`Header '${headerName}' not found in sheet.`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) {
      return [];
    }

    const columnValues = sheet.getRange(CONFIG.HEADER_ROW + 1, columnIndex, lastRow - CONFIG.HEADER_ROW, 1).getDisplayValues();
    const uniqueValues = new Set();

    columnValues.forEach(row => {
      const cellValue = row[0];
      if (typeof cellValue === 'string' && cellValue.trim() !== '') {
        uniqueValues.add(cellValue.trim());
      }
    });

    return Array.from(uniqueValues).sort();
  } catch (error) {
    Logger.log(`Error in getColumnUniqueValues: ${error.message}`);
    throw new Error(`Erro ao obter valores únicos para a coluna '${headerName}': ${error.message}`);
  }
}

// --- Custom View State Management ---

/**
 * Saves a custom view state (column visibility and filters) to Document Properties.
 * @param {string} viewName The base view name (e.g., "Default", "Resumo").
 * @param {string} stateName The user-defined name for this custom state.
 * @param {Object} settings An object containing 'columnVisibility' and 'filterSettings'.
 * @return {string} A success message.
 */
function saveCustomViewState(viewName, stateName, settings) {
  if (!viewName || !stateName || !settings) {
    throw new Error("Argumentos inválidos. Nome da vista, nome do estado e configurações são obrigatórios.");
  }
  try {
    const properties = PropertiesService.getDocumentProperties();
    const propertyKey = CONFIG.VIEW_STATE_PREFIX + viewName + "_" + stateName;
    properties.setProperty(propertyKey, JSON.stringify(settings));
    return `Estado '${stateName}' guardado com sucesso para a vista '${viewName}'.`;
  } catch (error) {
    Logger.log(`Error in saveCustomViewState: ${error.message}`);
    throw new Error(`Erro ao guardar o estado da vista '${stateName}': ${error.message}`);
  }
}

/**
 * Retrieves all saved custom view states for a given base view name.
 * @param {string} viewName The base view name (e.g., "Default").
 * @return {Array<Object>} An array of objects, each like { name: stateName, settings: parsedSettings }.
 */
function getCustomViewStates(viewName) {
  if (!viewName) {
    throw new Error("Nome da vista é obrigatório para obter os estados guardados.");
  }
  try {
    const properties = PropertiesService.getDocumentProperties().getProperties();
    const savedStates = [];
    const keyPrefixToSearch = CONFIG.VIEW_STATE_PREFIX + viewName + "_";

    for (const key in properties) {
      if (key.startsWith(keyPrefixToSearch)) {
        try {
          const stateName = key.substring(keyPrefixToSearch.length);
          const parsedSettings = JSON.parse(properties[key]);
          savedStates.push({
            name: stateName,
            settings: parsedSettings
          });
        } catch (parseError) {
          Logger.log(`Error parsing settings for key '${key}': ${parseError.message}. Skipping.`);
        }
      }
    }
    return savedStates;
  } catch (error) {
    Logger.log(`Error in getCustomViewStates: ${error.message}`);
    throw new Error(`Erro ao obter os estados da vista para '${viewName}': ${error.message}`);
  }
}

/**
 * Deletes a specific custom view state from Document Properties.
 * @param {string} viewName The base view name.
 * @param {string} stateName The user-defined name of the state to delete.
 * @return {string} A success message.
 */
function deleteCustomViewState(viewName, stateName) {
  if (!viewName || !stateName) {
    throw new Error("Nome da vista e nome do estado são obrigatórios para apagar.");
  }
  try {
    const properties = PropertiesService.getDocumentProperties();
    const propertyKey = CONFIG.VIEW_STATE_PREFIX + viewName + "_" + stateName;
    properties.deleteProperty(propertyKey);
    return `Estado '${stateName}' da vista '${viewName}' apagado com sucesso.`;
  } catch (error) {
    Logger.log(`Error in deleteCustomViewState: ${error.message}`);
    throw new Error(`Erro ao apagar o estado da vista '${stateName}': ${error.message}`);
  }
}

/**
 * Applies a previously saved custom view state (column visibility and filters).
 * @param {string} viewName The base view name.
 * @param {string} stateName The user-defined name of the state to apply.
 * @return {string} A success message.
 */
function applyCustomViewState(viewName, stateName) {
  if (!viewName || !stateName) {
    throw new Error("Nome da vista e nome do estado são obrigatórios para aplicar.");
  }
  try {
    const properties = PropertiesService.getDocumentProperties();
    const propertyKey = CONFIG.VIEW_STATE_PREFIX + viewName + "_" + stateName;
    const stringifiedSettings = properties.getProperty(propertyKey);

    if (!stringifiedSettings) {
      throw new Error(`Estado guardado '${stateName}' para a vista '${viewName}' não encontrado.`);
    }

    const parsedSettings = JSON.parse(stringifiedSettings);
    if (!parsedSettings.columnVisibility || !parsedSettings.filterSettings) {
      throw new Error(`Configurações para o estado '${stateName}' estão malformadas.`);
    }

    applyColumnVisibility(viewName, parsedSettings.columnVisibility);
    applyViewFilters(viewName, parsedSettings.filterSettings);

    return `Estado '${stateName}' da vista '${viewName}' aplicado com sucesso.`;

  } catch (error) {
    Logger.log(`Error in applyCustomViewState: ${error.message}`);
    throw new Error(`Erro ao aplicar o estado da vista '${stateName}': ${error.message}`);
  }
}

// --- Document and Quote Generation ---

/**
 * Reads last quote number, increments it, updates metadata sheet,
 * populates quote sheet with new number, date, and client info.
 * @return {string} Success or error message.
 */
function prepareQuoteSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG.SHEET_METADATA);
    const quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTE);

    if (!configSheet) throw new Error(`A folha de configuração '${CONFIG.SHEET_METADATA}' não foi encontrada.`);
    if (!quoteSheet) throw new Error(`A folha de orçamento '${CONFIG.SHEET_QUOTE}' não foi encontrada.`);

    // Increment Quote Number
    const lastQuoteNumCell = configSheet.getRange(CONFIG.QUOTE_CELLS.LAST_NUMBER);
    const lastQuoteNum = lastQuoteNumCell.getValue();
    if (typeof lastQuoteNum !== 'number' || !Number.isInteger(lastQuoteNum)) {
      throw new Error("O valor do último número de orçamento não é um número inteiro válido.");
    }
    const newQuoteNum = lastQuoteNum + 1;
    lastQuoteNumCell.setValue(newQuoteNum);

    // Prepare Display Values
    const currentYear = new Date().getFullYear();
    const displayQuoteNum = `${newQuoteNum}/${currentYear}`;
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

    // Get Client Info (D2:D4 from metadata)
    const clientInfoRange = configSheet.getRangeList([CONFIG.INFO_CELL_MAP.clientName, CONFIG.INFO_CELL_MAP.clientAddress, CONFIG.INFO_CELL_MAP.clientNif]);
    const [clientName, clientAddr1, clientNif] = clientInfoRange.getRanges().map(r => r.getValue() || '');

    // Update Quote Sheet
    quoteSheet.getRange(CONFIG.QUOTE_CELLS.TARGET_NUMBER).setValue(displayQuoteNum);
    quoteSheet.getRange(CONFIG.QUOTE_CELLS.TARGET_DATE).setValue(currentDate);
    quoteSheet.getRange(CONFIG.QUOTE_CELLS.TARGET_NAME).setValue(clientName);
    quoteSheet.getRange(CONFIG.QUOTE_CELLS.TARGET_ADDR1).setValue(clientAddr1);
    quoteSheet.getRange(CONFIG.QUOTE_CELLS.TARGET_ADDR2).setValue(""); // Address Line 2 assumed empty unless explicitly mapped
    quoteSheet.getRange(CONFIG.QUOTE_CELLS.TARGET_NIF).setValue(clientNif);

    SpreadsheetApp.flush();
    return `Orçamento #${displayQuoteNum} preparado com dados do cliente.`;

  } catch (error) {
    Logger.log(`Error in prepareQuoteSheet: ${error.message}`);
    return `Erro ao preparar orçamento: ${error.message}`;
  }
}

/**
 * Generates a PDF of the "Summary" view of the main panel sheet.
 * @return {string} The URL of the saved PDF file in Google Drive, or an error message string.
 */
function generateSummaryPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tempSheet = null;
  try {
    const sourceSheet = ss.getSheetByName(CONFIG.SHEET_MAIN_PANEL);
    if (!sourceSheet) throw new Error(`A folha '${CONFIG.SHEET_MAIN_PANEL}' não foi encontrada.`);

    const targetColumns = ["A", "B", "C", "D", "E", "F", "H", "N", "P", "Q"];
    const targetIndices = targetColumns.map(letter => sourceSheet.getRange(`${letter}1`).getColumn() - 1); // 0-based indices

    const allDataRange = sourceSheet.getDataRange();
    const allValues = allDataRange.getDisplayValues();
    if (allValues.length === 0) throw new Error("A folha ativa está vazia.");

    const headers = allValues[0];
    const filteredHeaders = targetIndices.map(index => headers[index]);
    const filteredData = [];

    for (let i = 1; i < allValues.length; i++) {
      const row = allValues[i];
      // Only include rows where the first column (POS, index 0) is not empty.
      if (row[0] && String(row[0]).trim() !== "") {
        const filteredRow = targetIndices.map(index => (index < row.length ? row[index] : ""));
        filteredData.push(filteredRow);
      }
    }

    if (filteredData.length === 0) throw new Error("Não foram encontrados dados válidos para incluir no PDF.");

    // Create and format temporary sheet
    const tempSheetName = `TempPDF_${new Date().getTime()}`;
    tempSheet = ss.insertSheet(tempSheetName);
    const headerRange = tempSheet.getRange(1, 1, 1, filteredHeaders.length);
    const dataRange = tempSheet.getRange(2, 1, filteredData.length, filteredHeaders.length);

    headerRange.setValues([filteredHeaders]);
    dataRange.setValues(filteredData);

    headerRange
      .setBackground("#4285F4")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontFamily(CONFIG.FONT_FAMILY)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    dataRange
      .setFontFamily(CONFIG.FONT_FAMILY)
      .setVerticalAlignment("top");

    for (let j = 1; j <= filteredHeaders.length; j++) {
      tempSheet.autoResizeColumn(j);
    }
    SpreadsheetApp.flush();

    // Export PDF
    const pdfUrl = buildPdfExportUrl(ss.getId(), tempSheet.getSheetId());
    const pdfFile = fetchAndSavePdf(pdfUrl, `Sumario_${sourceSheet.getName()}`, CONFIG.PDF_FOLDER_NAME);

    return pdfFile.getUrl();

  } catch (error) {
    Logger.log(`Error in generateSummaryPdf: ${error.message}`);
    return `Erro: ${error.message}`;
  } finally {
    // Clean up temporary sheet
    if (tempSheet) {
      try {
        ss.deleteSheet(tempSheet);
      } catch (e) {
        Logger.log(`Error deleting temporary sheet: ${e.message}`);
      }
    }
  }
}

/**
 * Generates a PDF of the Quote sheet.
 * @return {string} The URL of the saved PDF file in Google Drive, or an error message string.
 */
function generateQuotePdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const quoteSheet = ss.getSheetByName(CONFIG.SHEET_QUOTE);
    if (!quoteSheet) throw new Error(`A folha de orçamento '${CONFIG.SHEET_QUOTE}' não foi encontrada.`);

    const quoteNumber = quoteSheet.getRange(CONFIG.QUOTE_CELLS.TARGET_NUMBER).getDisplayValue().replace("/", "-");
    const clientName = quoteSheet.getRange(CONFIG.QUOTE_CELLS.TARGET_NAME).getDisplayValue();

    // Build URL for quote sheet (A4 Portrait, no gridlines, no fit-to-width)
    const pdfUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
      `format=pdf&gid=${quoteSheet.getSheetId()}&size=a4&portrait=true&fitw=false&scale=1&` +
      `top_margin=0.75&bottom_margin=0.75&left_margin=0.70&right_margin=0.70&` +
      `gridlines=false&printtitle=false&sheetnames=false&fzr=false`;

    const pdfFile = fetchAndSavePdf(pdfUrl, `Orcamento_${quoteNumber}_${clientName}`, CONFIG.PDF_FOLDER_NAME);

    return pdfFile.getUrl();
  } catch (error) {
    Logger.log(`Error in generateQuotePdf: ${error.message}`);
    return `Erro ao gerar PDF do orçamento: ${error.message}`;
  }
}

/**
 * Helper to construct the PDF export URL for a temporary summary sheet.
 * @param {string} ssId Spreadsheet ID.
 * @param {number} sheetId GID of the sheet.
 * @return {string} The PDF export URL.
 */
function buildPdfExportUrl(ssId, sheetId) {
  return `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
    `format=pdf&gid=${sheetId}&size=a4&portrait=false&fitw=true&scale=4&` +
    `top_margin=0.50&bottom_margin=0.50&left_margin=0.50&right_margin=0.50&` +
    `gridlines=true&printtitle=false&sheetnames=false&fzr=false&` +
    `horizontal_alignment=CENTER&vertical_alignment=TOP`;
}

/**
 * Helper to fetch the PDF blob and save it to Drive.
 * @param {string} pdfUrl The generated PDF export URL.
 * @param {string} baseFileName The base name for the PDF file (before date suffix).
 * @param {string} folderName The name of the Drive folder to save to.
 * @return {GoogleAppsScript.Drive.File} The created Drive file object.
 */
function fetchAndSavePdf(pdfUrl, baseFileName, folderName) {
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(pdfUrl, {
    headers: {
      'Authorization': `Bearer ${token}`
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Erro (${response.getResponseCode()}) ao gerar o PDF a partir do Google Sheets.`);
  }

  const blob = response.getBlob();
  let pdfFileName = `${baseFileName}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.pdf`;
  pdfFileName = pdfFileName.replace(/[/\\?%*:|"<>]/g, '-'); // Sanitize filename
  blob.setName(pdfFileName);

  const folders = DriveApp.getFoldersByName(folderName);
  const pdfFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  // Remove existing file with the same name
  const existingFiles = pdfFolder.getFilesByName(pdfFileName);
  if (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  return pdfFolder.createFile(blob);
}

// --- Formatting and Triggers ---

/**
 * Converts all text content in the active sheet (excluding header) to uppercase.
 * @return {string} A success or error message.
 */
function convertSheetToUppercase() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const headerRowsToSkip = CONFIG.HEADER_ROW;
    const dataRange = sheet.getDataRange();

    if (dataRange.getNumRows() <= headerRowsToSkip) {
      return "Nenhuma linha de dados encontrada abaixo do cabeçalho para converter.";
    }

    const startRow = headerRowsToSkip + 1;
    const numRows = dataRange.getNumRows() - headerRowsToSkip;
    const numCols = dataRange.getNumColumns();

    const dataOnlyRange = sheet.getRange(startRow, 1, numRows, numCols);
    const values = dataOnlyRange.getValues();
    let changed = false;

    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        if (typeof values[i][j] === 'string') {
          const originalValue = values[i][j];
          const upperValue = originalValue.toUpperCase();
          if (originalValue !== upperValue) {
            values[i][j] = upperValue;
            changed = true;
          }
        }
      }
    }

    if (changed) {
      dataOnlyRange.setValues(values);
      return "Texto da folha convertido para maiúsculas!";
    } else {
      return "Nenhum texto encontrado que necessite de conversão.";
    }

  } catch (error) {
    Logger.log(`Error in convertSheetToUppercase: ${error.message}`);
    return `Erro ao converter para maiúsculas: ${error.message}`;
  }
}

/**
 * Iterates through all columns in the HEADER_ROW of the target sheet and
 * applies the two-line rich text format if the cell contains a newline '\n'.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 */
function fixHeaderFormatting(sheet) {
  try {
    const maxCols = sheet.getMaxColumns();
    const headerRange = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, maxCols);
    const headerValues = headerRange.getValues()[0];

    for (let col = 1; col <= maxCols; col++) {
      const headerText = headerValues[col - 1];

      if (typeof headerText === 'string' && headerText.includes('\n')) {
        const targetRange = sheet.getRange(CONFIG.HEADER_ROW, col);
        const parts = headerText.split('\n', 2);
        const line1 = parts[0];

        const richText = SpreadsheetApp.newRichTextValue()
          .setText(headerText)
          .setTextStyle(0, line1.length, HEADER_STYLE_PRIMARY)
          .setTextStyle(line1.length + 1, headerText.length, HEADER_STYLE_SECONDARY)
          .build();

        targetRange.setRichTextValue(richText);
      }
    }
    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log(`Error in fixHeaderFormatting: ${error.message}`);
  }
}

/**
 * Handles automatic formatting and logic upon spreadsheet edits.
 * 1. Formats two-line headers across the entire row.
 * 2. Conditionally bolds the edited row based on Column H status.
 * 3. Adds a warning note to Column R based on Q/R date comparison.
 * @param {Object} e The event object passed by the onEdit trigger.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const {
    range,
    range: {
      rowStart: row,
      columnStart: col
    },
    value,
    oldValue
  } = e;
  const sheet = range.getSheet();

  // Only run advanced logic on the main data sheet
  if (sheet.getName() !== CONFIG.SHEET_MAIN_PANEL) return;

  // --- Logic 1: Two-Line Header Formatting Fix (Runs on every edit) ---
  fixHeaderFormatting(sheet);

  // Skip data-dependent logic if editing the header row
  if (row <= CONFIG.HEADER_ROW) return;

  // --- Logic 2: Conditional Row Bolding (Column H) ---
  const TARGET_COL_H = 8;
  if (col === TARGET_COL_H && range.getNumRows() === 1 && range.getNumColumns() === 1) {
    if (value === oldValue) return;
    try {
      const triggerWords = ["Aprovado", "Por Definir | M", "Por Desenhar | M", "Por Orçamentar | M", "Por Aprovar | C", "Por Aprovar | M", "Por Levantar"];
      const shouldBold = triggerWords.includes(range.getValue());
      const lastColumn = sheet.getLastColumn();
      const targetWeight = shouldBold ? "bold" : "normal";

      // Apply format only if it's changing
      if (sheet.getRange(row, 1).getFontWeight() !== targetWeight) {
        sheet.getRange(row, 1, 1, lastColumn).setFontWeight(targetWeight);
      }
    } catch (error) {
      Logger.log(`Error in onEdit (Bold Logic, Row ${row}): ${error.message}`);
    }
  }

  // --- Logic 3: Date Comparison Note (Columns Q & R) ---
  const TARGET_COL_Q = 17;
  const TARGET_COL_R = 18;
  if (col === TARGET_COL_Q || col === TARGET_COL_R) {
    if (value === oldValue) return;
    try {
      const cellQ = sheet.getRange(row, TARGET_COL_Q);
      const cellR = sheet.getRange(row, TARGET_COL_R);
      const dateQ = cellQ.getValue();
      const dateR = cellR.getValue();
      let noteMsg = "";

      if (dateQ instanceof Date && dateR instanceof Date) {
        const dayDifference = Math.floor((dateQ.getTime() - dateR.getTime()) / (1000 * 60 * 60 * 24));
        if (dayDifference < 0) {
          noteMsg = `⚠️ Atrasado por ${Math.abs(dayDifference)} dias!`;
        }
      }

      if (cellR.getNote() !== noteMsg) {
        cellR.setNote(noteMsg);
      }
    } catch (error) {
      Logger.log(`Error in onEdit (Date Logic, Row ${row}): ${error.message}`);
    }
  }
}

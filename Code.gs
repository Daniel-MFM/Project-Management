/**
 * @OnlyCurrentDoc
 */

// --- Configuration ---
const CONFIG_SHEET_NAME = "METADATA (NAO MEXER)"; // Name of the sheet holding project info
const PDF_FOLDER_NAME = "Relatorios PDF Gerados"; // Folder name in Google Drive to save PDFs
const HEADER_ROW = 1; // Row number containing headers to format
const FONT_FAMILY = "Reddit Sans Condensed"; // Default font for formatting
const TARGET_DATA_SHEET_NAME = "PAINEL PRINCIPAL"; // Name of the sheet the view functions should affect
const QUOTE_SHEET_NAME = "ORÇAMENTO"; // Name of the quote sheet
const LAST_QUOTE_NUMBER_CELL = "E7"; // Cell on CONFIG_SHEET_NAME holding the last quote number (e.g., 28)
const QUOTE_NUMBER_TARGET_CELL = "I2"; // Cell on QUOTE_SHEET_NAME for the quote number (e.g., 29/2025)
const QUOTE_DATE_TARGET_CELL = "I3"; // Cell on QUOTE_SHEET_NAME for the date
const QUOTE_CLIENT_NAME_CELL = "F4"; // Cell on QUOTE_SHEET_NAME for Client Name
const QUOTE_CLIENT_ADDR1_CELL = "F5"; // Cell on QUOTE_SHEET_NAME for Address Line 1
const QUOTE_CLIENT_ADDR2_CELL = "F6"; // Cell on QUOTE_SHEET_NAME for Address Line 2
const QUOTE_CLIENT_NIF_CELL = "F7"; // Cell on QUOTE_SHEET_NAME for NIF
const VIEW_STATE_PREFIX = "customViewState_"; // Prefix for storing custom view states in PropertiesService

// Map of field names (expected from sidebar/client-side) to cell notations in CONFIG_SHEET_NAME
const PROJECT_INFO_CELL_MAP = {
  "clientName": "D2",
  "clientAddress": "D3",
  "clientNif": "D4",
  "clientEmail": "D5",
  "floorplanUrl": "D6"
};
// --- End Configuration ---

// --- Global Text Styles (for onEdit header formatting) ---
const BOLD_STYLE_11 = SpreadsheetApp.newTextStyle()
  .setBold(true)
  .setFontSize(11)
  .setFontFamily(FONT_FAMILY)
  .build();
const REGULAR_STYLE_9 = SpreadsheetApp.newTextStyle()
  .setBold(false)
  .setFontSize(9)
  .setFontFamily(FONT_FAMILY)
  .build();
// --- End Styles ---


// --- Menu, Dialog, and Sidebar ---

/**
 * Adds custom menu to show the sidebar directly when the spreadsheet opens.
 * Runs automatically via the onOpen simple trigger.
 */
function onOpen() {
  // Add menu item to directly show the sidebar
  try {
      SpreadsheetApp.getUi()
        .createMenu("Painel de Controlo") // UI: Control Panel Menu Name (PT-PT)
        .addItem("Mostrar Painel", "showSidebar") // UI: Menu item text (PT-PT)
        .addToUi();
      Logger.log("Menu 'Painel de Controlo' added (calls showSidebar).");
  } catch (e) {
      Logger.log("Error adding custom menu: " + e);
  }
}

/**
 * Shows a modal dialog to edit a specific project information field.
 * @param {string} fieldName The name of the field to be edited (e.g., "clientName").
 * @param {string} currentValue The current value of the field.
 */
function showEditInfoFieldModal(fieldName, currentValue) {
  try {
    if (!fieldName || typeof currentValue === 'undefined') { // currentValue can be empty string, so check for undefined
      Logger.log("showEditInfoFieldModal: fieldName or currentValue is missing. fieldName: " + fieldName + ", currentValue: " + currentValue);
      SpreadsheetApp.getUi().alert("Erro: Informação incompleta para editar o campo.");
      return;
    }
    
    // Optional: Validate fieldName against PROJECT_INFO_CELL_MAP keys if desired server-side,
    // though client-side should ideally only send valid fieldNames.
    if (!PROJECT_INFO_CELL_MAP.hasOwnProperty(fieldName)) {
        Logger.log("showEditInfoFieldModal: Invalid fieldName received: " + fieldName);
        SpreadsheetApp.getUi().alert("Erro: Tentativa de editar um campo desconhecido: " + fieldName);
        return;
    }

    Logger.log("showEditInfoFieldModal: Called for fieldName = '" + fieldName + "', currentValue = '" + currentValue + "'");

    var template = HtmlService.createTemplateFromFile("EditInfoFieldDialog");
    template.fieldName = fieldName;
    template.currentValue = currentValue;

    // Attempt to create a more descriptive title from the fieldName
    // This is a simple example; a more robust mapping might be needed for better PT-PT titles
    var dialogTitle = "Editar " + fieldName;
    if (fieldName === "clientName") dialogTitle = "Editar Nome do Cliente";
    else if (fieldName === "clientAddress") dialogTitle = "Editar Morada do Cliente";
    else if (fieldName === "clientNif") dialogTitle = "Editar NIF do Cliente";
    else if (fieldName === "clientEmail") dialogTitle = "Editar Email do Cliente";
    else if (fieldName === "floorplanUrl") dialogTitle = "Editar Link da Planta";


    var htmlOutput = template.evaluate()
        .setWidth(350)  // Adjusted width
        .setHeight(280); // Adjusted height
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
    Logger.log("showEditInfoFieldModal: Dialog displayed for fieldName = '" + fieldName + "'");

  } catch (e) {
    Logger.log("Error in showEditInfoFieldModal: " + e.message + " (Stack: " + e.stack + ")");
    SpreadsheetApp.getUi().alert("Erro ao tentar mostrar o diálogo de edição: " + e.message);
  }
}

/**
 * Updates a specific project information field in the CONFIG_SHEET_NAME.
 * @param {string} fieldName The name of the field to update (must be a key in PROJECT_INFO_CELL_MAP).
 * @param {string} newValue The new value for the field.
 * @return {Object} An object indicating success or failure, including a message.
 */
function updateProjectInfoField(fieldName, newValue) {
  Logger.log("updateProjectInfoField: Called with fieldName='" + fieldName + "', newValue='" + newValue + "'");
  try {
    if (!PROJECT_INFO_CELL_MAP.hasOwnProperty(fieldName)) {
      Logger.log("updateProjectInfoField: Invalid field name provided: " + fieldName);
      throw new Error("Nome de campo inválido fornecido para atualização.");
    }

    const cellNotation = PROJECT_INFO_CELL_MAP[fieldName];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!configSheet) {
      Logger.log("updateProjectInfoField: Configuration sheet '" + CONFIG_SHEET_NAME + "' not found.");
      throw new Error("A folha de configuração '" + CONFIG_SHEET_NAME + "' não foi encontrada.");
    }

    var cell = configSheet.getRange(cellNotation);
    cell.setValue(newValue);
    SpreadsheetApp.flush(); // Apply the change immediately

    Logger.log("updateProjectInfoField: Field '" + fieldName + "' (Cell: " + cellNotation + ") updated successfully to: " + newValue);
    return { 
      success: true, 
      message: "'" + fieldName + "' atualizado com sucesso para: '" + newValue + "'.", // Simple name for now, can map to PT-PT later
      updatedField: fieldName, 
      value: newValue 
    };

  } catch (error) {
    Logger.log("Error in updateProjectInfoField: " + error.message + " (Stack: " + error.stack + ")");
    // Re-throw with a user-friendly message, potentially including the original if it's already user-friendly
    if (error.message.startsWith("Erro ao") || error.message.startsWith("Nome de campo inválido")) { 
        throw error;
    }
    throw new Error("Erro ao atualizar o campo '" + fieldName + "': " + error.message);
  }
}

/**
 * Shows the Column Chooser modal dialog for a given view name.
 * @param {string} viewName The name of the view context for the dialog.
 */
function showColumnChooserModal(viewName) {
  try {
    if (!viewName) {
      Logger.log("showColumnChooserModal: viewName is undefined or empty. Cannot show dialog without a view context.");
      SpreadsheetApp.getUi().alert("Erro: Nome da vista não especificado para o personalizador de colunas.");
      return;
    }
    Logger.log("showColumnChooserModal: Called for viewName = '" + viewName + "'");

    var template = HtmlService.createTemplateFromFile("ColumnChooserDialog");
    template.viewName = viewName; // Pass viewName to the template for client-side access

    var htmlOutput = template.evaluate()
        .setWidth(380)  // Adjusted width
        .setHeight(580); // Adjusted height, might need further tuning based on content
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Personalizar Vista: " + viewName);
    Logger.log("showColumnChooserModal: Dialog displayed for viewName = '" + viewName + "'");

  } catch (e) {
    Logger.log("Error in showColumnChooserModal: " + e.message + " (Stack: " + e.stack + ")");
    SpreadsheetApp.getUi().alert("Erro ao tentar mostrar o diálogo de personalização de colunas: " + e.message);
  }
}

// --- Custom View State Management ---

/**
 * Saves a custom view state (column visibility and filters) to Document Properties.
 * @param {string} viewName The base view name (e.g., "Default", "Resumo").
 * @param {string} stateName The user-defined name for this custom state.
 * @param {Object} settings An object containing 'columnVisibility' and 'filterSettings'.
 * @return {string} A success message.
 * @throws {Error} If settings cannot be saved.
 */
function saveCustomViewState(viewName, stateName, settings) {
  Logger.log("saveCustomViewState: Called for viewName='" + viewName + "', stateName='" + stateName + "', settings:", settings);
  if (!viewName || !stateName || !settings || typeof settings.columnVisibility === 'undefined' || typeof settings.filterSettings === 'undefined') {
    Logger.log("saveCustomViewState: Invalid arguments. View name, state name, and settings (with columnVisibility and filterSettings) are required.");
    throw new Error("Argumentos inválidos. Nome da vista, nome do estado e configurações (com visibilidade de coluna e filtros) são obrigatórios.");
  }
  try {
    var properties = PropertiesService.getDocumentProperties();
    var propertyKey = VIEW_STATE_PREFIX + viewName + "_" + stateName;
    var stringifiedSettings = JSON.stringify(settings);

    properties.setProperty(propertyKey, stringifiedSettings);
    Logger.log("saveCustomViewState: State '" + stateName + "' for view '" + viewName + "' saved successfully. Key: " + propertyKey);
    return "Estado '" + stateName + "' guardado com sucesso para a vista '" + viewName + "'.";
  } catch (error) {
    Logger.log("Error in saveCustomViewState: " + error.message + " (Stack: " + error.stack + ")");
    throw new Error("Erro ao guardar o estado da vista '" + stateName + "': " + error.message);
  }
}

/**
 * Retrieves all saved custom view states for a given base view name.
 * @param {string} viewName The base view name (e.g., "Default").
 * @return {Array<Object>} An array of objects, each like { name: stateName, settings: parsedSettings }.
 * @throws {Error} If states cannot be retrieved.
 */
function getCustomViewStates(viewName) {
  Logger.log("getCustomViewStates: Called for viewName='" + viewName + "'");
  if (!viewName) {
    Logger.log("getCustomViewStates: View name is required.");
    throw new Error("Nome da vista é obrigatório para obter os estados guardados.");
  }
  try {
    var properties = PropertiesService.getDocumentProperties().getProperties();
    var savedStates = [];
    var keyPrefixToSearch = VIEW_STATE_PREFIX + viewName + "_";

    for (var key in properties) {
      if (properties.hasOwnProperty(key) && key.startsWith(keyPrefixToSearch)) {
        try {
          var stateName = key.substring(keyPrefixToSearch.length);
          var stringifiedSettings = properties[key];
          var parsedSettings = JSON.parse(stringifiedSettings);
          savedStates.push({ name: stateName, settings: parsedSettings });
          Logger.log("getCustomViewStates: Found state: name='" + stateName + "' for view='" + viewName + "'");
        } catch (parseError) {
          Logger.log("Error parsing settings for key '" + key + "': " + parseError.message + ". Skipping this state.");
          // Optionally, could delete the malformed property here.
        }
      }
    }
    Logger.log("getCustomViewStates: Returning " + savedStates.length + " states for view '" + viewName + "'.");
    return savedStates;
  } catch (error) {
    Logger.log("Error in getCustomViewStates: " + error.message + " (Stack: " + error.stack + ")");
    throw new Error("Erro ao obter os estados da vista para '" + viewName + "': " + error.message);
  }
}

/**
 * Deletes a specific custom view state from Document Properties.
 * @param {string} viewName The base view name.
 * @param {string} stateName The user-defined name of the state to delete.
 * @return {string} A success message.
 * @throws {Error} If the state cannot be deleted.
 */
function deleteCustomViewState(viewName, stateName) {
  Logger.log("deleteCustomViewState: Called for viewName='" + viewName + "', stateName='" + stateName + "'");
   if (!viewName || !stateName) {
    Logger.log("deleteCustomViewState: View name and state name are required.");
    throw new Error("Nome da vista e nome do estado são obrigatórios para apagar.");
  }
  try {
    var properties = PropertiesService.getDocumentProperties();
    var propertyKey = VIEW_STATE_PREFIX + viewName + "_" + stateName;

    properties.deleteProperty(propertyKey);
    Logger.log("deleteCustomViewState: State '" + stateName + "' for view '" + viewName + "' deleted successfully. Key: " + propertyKey);
    return "Estado '" + stateName + "' da vista '" + viewName + "' apagado com sucesso.";
  } catch (error) {
    Logger.log("Error in deleteCustomViewState: " + error.message + " (Stack: " + error.stack + ")");
    throw new Error("Erro ao apagar o estado da vista '" + stateName + "': " + error.message);
  }
}

/**
 * Applies a previously saved custom view state (column visibility and filters).
 * @param {string} viewName The base view name.
 * @param {string} stateName The user-defined name of the state to apply.
 * @return {string} A success message.
 * @throws {Error} If the state is not found or cannot be applied.
 */
function applyCustomViewState(viewName, stateName) {
  Logger.log("applyCustomViewState: Called for viewName='" + viewName + "', stateName='" + stateName + "'");
  if (!viewName || !stateName) {
    Logger.log("applyCustomViewState: View name and state name are required.");
    throw new Error("Nome da vista e nome do estado são obrigatórios para aplicar.");
  }
  try {
    var properties = PropertiesService.getDocumentProperties();
    var propertyKey = VIEW_STATE_PREFIX + viewName + "_" + stateName;
    var stringifiedSettings = properties.getProperty(propertyKey);

    if (!stringifiedSettings) {
      Logger.log("applyCustomViewState: State '" + stateName + "' for view '" + viewName + "' not found. Key: " + propertyKey);
      throw new Error("Estado guardado '" + stateName + "' para a vista '" + viewName + "' não encontrado.");
    }

    var parsedSettings;
    try {
      parsedSettings = JSON.parse(stringifiedSettings);
    } catch (parseError) {
      Logger.log("Error parsing settings for state '" + stateName + "' (key '" + propertyKey + "'): " + parseError.message);
      throw new Error("Erro ao ler as configurações do estado guardado '" + stateName + "'.");
    }
    
    if (!parsedSettings.columnVisibility || !parsedSettings.filterSettings) {
        Logger.log("applyCustomViewState: Settings object for state '" + stateName + "' is malformed. Missing columnVisibility or filterSettings.");
        throw new Error("Configurações para o estado '" + stateName + "' estão malformadas.");
    }

    Logger.log("applyCustomViewState: Applying column visibility for state '" + stateName + "'. Settings:", parsedSettings.columnVisibility);
    applyColumnVisibility(viewName, parsedSettings.columnVisibility); // viewName for logging inside this function

    Logger.log("applyCustomViewState: Applying filters for state '" + stateName + "'. Settings:", parsedSettings.filterSettings);
    applyViewFilters(viewName, parsedSettings.filterSettings); // viewName for logging inside this function

    Logger.log("applyCustomViewState: State '" + stateName + "' for view '" + viewName + "' applied successfully.");
    return "Estado '" + stateName + "' da vista '" + viewName + "' aplicado com sucesso.";

  } catch (error) { // Catches errors from this function and from applyColumnVisibility/applyViewFilters
    Logger.log("Error in applyCustomViewState: " + error.message + " (Stack: " + error.stack + ")");
    // Re-throw with a user-friendly message, potentially including the original if it's already user-friendly
    if (error.message.startsWith("Erro ao")) { // If it's one of our custom Portuguese errors
        throw error;
    }
    throw new Error("Erro ao aplicar o estado da vista '" + stateName + "': " + error.message);
  }
}

/**
 * Applies filters to the TARGET_DATA_SHEET_NAME based on the provided settings.
 * @param {string} viewName The name of the view being filtered (for logging).
 * @param {Object} filterSettings An object where keys are header names and values are the string values to filter by (exact match).
 * @return {string} A success or informational message.
 * @throws {Error} If the sheet is not found or filters cannot be applied.
 */
function applyViewFilters(viewName, filterSettings) {
  Logger.log("applyViewFilters: Called for view '" + viewName + "' with settings: " + JSON.stringify(filterSettings));
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(TARGET_DATA_SHEET_NAME);
    if (!sheet) {
      throw new Error("Sheet '" + TARGET_DATA_SHEET_NAME + "' not found.");
    }

    // Remove existing filter, if any
    var existingFilter = sheet.getFilter();
    if (existingFilter) {
      Logger.log("applyViewFilters: Removing existing filter from sheet '" + TARGET_DATA_SHEET_NAME + "'.");
      existingFilter.remove();
    }

    // If filterSettings is empty or null, we just wanted to clear filters.
    if (!filterSettings || Object.keys(filterSettings).length === 0) {
      Logger.log("applyViewFilters: No filter settings provided or settings are empty. Filters cleared.");
      return "Filtros removidos ou nenhum filtro aplicado."; // UI: Success/Info (PT-PT)
    }

    var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getMaxColumns()).getValues()[0];
    var headerMap = {}; // To store headerName: columnIndex
    for (var i = 0; i < headers.length; i++) {
      if (typeof headers[i] === 'string' && headers[i].trim() !== '') {
        headerMap[headers[i].trim()] = i + 1; // 1-based index
      }
    }

    var columnCriteria = {}; // Store FilterCriterion objects, keyed by 1-based column index

    for (var headerName in filterSettings) {
      if (filterSettings.hasOwnProperty(headerName)) {
        var filterValue = filterSettings[headerName];
        if (headerMap.hasOwnProperty(headerName)) {
          var columnIndex = headerMap[headerName];
          if (typeof filterValue === 'string' && filterValue !== '') { // Ensure there's a value to filter by
            Logger.log("applyViewFilters: Creating filter for column '" + headerName + "' (index " + columnIndex + ") with value '" + filterValue + "'.");
            var criterion = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(filterValue).build();
            columnCriteria[columnIndex] = criterion;
          } else {
            Logger.log("applyViewFilters: Skipping filter for column '" + headerName + "' due to empty filter value.");
          }
        } else {
          Logger.log("applyViewFilters: Warning - Header '" + headerName + "' from filterSettings not found in sheet. Skipping this filter.");
        }
      }
    }

    if (Object.keys(columnCriteria).length > 0) {
      Logger.log("applyViewFilters: Applying new filter criteria to sheet.");
      var filterRange = sheet.getRange(HEADER_ROW, 1, sheet.getMaxRows() - HEADER_ROW, sheet.getMaxColumns());
      var newFilter = filterRange.createFilter();

      for (var colIdxStr in columnCriteria) {
        if (columnCriteria.hasOwnProperty(colIdxStr)) {
          var colIdx = parseInt(colIdxStr);
           Logger.log("applyViewFilters: Setting filter for column index " + colIdx);
          newFilter.setColumnFilterCriteria(colIdx, columnCriteria[colIdx]);
        }
      }
      Logger.log("applyViewFilters: Filter applied successfully for view '" + viewName + "'.");
      return "Filtros aplicados com sucesso para a visão '" + viewName + "'."; // UI: Success (PT-PT)
    } else {
      Logger.log("applyViewFilters: No valid criteria were built. No new filter applied.");
      return "Nenhum critério de filtro válido foi aplicado."; // UI: Info (PT-PT)
    }

  } catch (error) {
    Logger.log("Error in applyViewFilters for view '" + viewName + "': " + error.message + " (Stack: " + error.stack + ")");
    throw new Error("Erro ao aplicar filtros para a visão '" + viewName + "': " + error.message); // UI: Error (PT-PT)
  }
}

/**
 * Retrieves unique, sorted, non-empty string values from a specified column in TARGET_DATA_SHEET_NAME.
 * @param {string} headerName The text of the header for the column to process.
 * @param {string} viewName For logging/consistency, name of the view this might be related to.
 * @return {Array<string>} A sorted array of unique string values from the column.
 * @throws {Error} If sheet or header is not found, or if column data cannot be read.
 */
function getColumnUniqueValues(headerName, viewName) {
  Logger.log("getColumnUniqueValues: Called for header '" + headerName + "', viewName '" + viewName + "'. Operates on TARGET_DATA_SHEET_NAME.");
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(TARGET_DATA_SHEET_NAME);
    if (!sheet) {
      throw new Error("Sheet '" + TARGET_DATA_SHEET_NAME + "' not found.");
    }

    var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getMaxColumns()).getValues()[0];
    var columnIndex = -1;
    for (var i = 0; i < headers.length; i++) {
      if (typeof headers[i] === 'string' && headers[i].trim() === headerName) {
        columnIndex = i + 1; // 1-based index
        break;
      }
    }

    if (columnIndex === -1) {
      throw new Error("Header '" + headerName + "' not found in sheet '" + TARGET_DATA_SHEET_NAME + "'.");
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROW) {
      Logger.log("getColumnUniqueValues: No data rows found below header for column '" + headerName + "'.");
      return []; // No data rows
    }

    var columnValues = sheet.getRange(HEADER_ROW + 1, columnIndex, lastRow - HEADER_ROW, 1).getDisplayValues();
    var uniqueValues = new Set();

    columnValues.forEach(function(row) {
      var cellValue = row[0];
      if (typeof cellValue === 'string' && cellValue.trim() !== '') {
        uniqueValues.add(cellValue.trim());
      }
    });

    var sortedUniqueValues = Array.from(uniqueValues).sort();
    Logger.log("getColumnUniqueValues: Found " + sortedUniqueValues.length + " unique values for header '" + headerName + "': " + JSON.stringify(sortedUniqueValues));
    return sortedUniqueValues;

  } catch (error) {
    Logger.log("Error in getColumnUniqueValues for header '" + headerName + "': " + error.message + " (Stack: " + error.stack + ")");
    // It's important to re-throw the error so the client-side (Sidebar.html) can catch it with .withFailureHandler
    throw new Error("Erro ao obter valores únicos para a coluna '" + headerName + "': " + error.message); // UI: Error (PT-PT)
  }
}

/**
 * Gets the current visibility state of all columns in the TARGET_DATA_SHEET_NAME.
 * Intended for the Column Chooser feature to know the current state.
 * @param {string} viewName The name of the view (primarily for logging/consistency).
 * @return {Object} An object where keys are header names and values are booleans (true if visible, false if hidden).
 * @throws {Error} If the sheet is not found or headers cannot be read.
 */
function getCurrentColumnVisibility(viewName) {
  Logger.log("getCurrentColumnVisibility: Called for viewName = '" + viewName + "' (operates on TARGET_DATA_SHEET_NAME)");
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // This function currently assumes views always target the TARGET_DATA_SHEET_NAME.
    // Could be made more flexible if different views target different sheets.
    var sheet = ss.getSheetByName(TARGET_DATA_SHEET_NAME);
    if (!sheet) {
      throw new Error("Sheet '" + TARGET_DATA_SHEET_NAME + "' not found.");
    }

    var maxCols = sheet.getMaxColumns();
    var headerRowValues = sheet.getRange(HEADER_ROW, 1, 1, maxCols).getValues()[0];
    var visibilityStates = {};

    for (var i = 0; i < maxCols; i++) {
      var headerText = (typeof headerRowValues[i] === 'string') ? headerRowValues[i].trim() : '';
      if (headerText !== '') { // Only consider columns with actual header text
        var columnIndex = i + 1; // 1-based index
        visibilityStates[headerText] = !sheet.isColumnHiddenByUser(columnIndex);
      }
    }

    Logger.log("getCurrentColumnVisibility: Returning states: " + JSON.stringify(visibilityStates));
    return visibilityStates;

  } catch (error) {
    Logger.log("Error in getCurrentColumnVisibility: " + error.message + " (Stack: " + error.stack + ")");
    throw new Error("Erro ao obter o estado de visibilidade das colunas para '" + viewName + "': " + error.message); // UI: Error (PT-PT)
  }
}

/**
 * Displays a modal dialog asking the user to open the main sidebar.
 * NOTE: This function is currently NOT called by default.
 */
function showSidebarPromptDialog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('OpenSidebarPrompt')
        .setWidth(300)
        .setHeight(120);
    SpreadsheetApp.getUi().showModalDialog(html, 'Abrir Painel de Controlo?'); // UI: Dialog Title (PT-PT)
    Logger.log("Dialog 'Abrir Painel' displayed via manual call.");
  } catch (e) {
    Logger.log("Error showing prompt dialog: " + e);
    SpreadsheetApp.getUi().alert("Erro ao tentar abrir o diálogo do painel: " + e.message); // UI: Alert text (PT-PT)
  }
}

/**
 * Shows the main sidebar panel containing the custom UI.
 * Reads content from the 'Sidebar.html' file.
 */
function showSidebar() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('Sidebar')
        .setTitle('Painel de Controlo') // UI: Sidebar Title (PT-PT)
        .setWidth(240);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.log("Error in showSidebar: " + error);
    SpreadsheetApp.getUi().alert("Não foi possível abrir o Painel de Controlo. Verifique o ficheiro 'Sidebar.html' no editor de scripts. Erro: " + error.message); // UI: Alert text (PT-PT)
  }
}

// --- View Switching Functions ---

function SetViewDefault() { setView(["A", "B", "D", "E", "F", "H", "I", "L", "Q", "U", "V"]); }
function SetViewBudget() { setView(["A", "B", "C", "D", "E", "F", "G", "I", "J", "L", "M", "N", "O", "P"]); }
function SetViewClient() { setView(["A", "B", "C", "D", "E", "F", "H", "P"]); }
function SetViewResumo() { setView(["A", "B", "C", "D", "E", "F", "H", "N", "P", "Q", "U", "V"]); }
function setViewLayout() { setView(["A", "B", "C", "D"]); }

/**
 * Applies a view by showing/hiding columns on the TARGET_DATA_SHEET_NAME.
 * @param {Array<string>} allowedColumnsArray Array of column letters to show.
 */
function setView(allowedColumnsArray) {
 try {
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   var sheetName = sheet.getName();
   if (sheetName !== TARGET_DATA_SHEET_NAME) {
       SpreadsheetApp.getUi().alert('As funções de visualização só podem ser usadas na folha "' + TARGET_DATA_SHEET_NAME + '".'); // UI: Alert (PT-PT)
       return;
   }
   var maxCols = sheet.getMaxColumns();
   var allowedNumsSet = new Set(allowedColumnsArray.map(columnLetterToNumber));
   // SpreadsheetApp.flush(); // Removed: Attempting to reduce flushes. The showColumns below might make this redundant.
   try { sheet.showColumns(1, maxCols); } catch(uiError) { Logger.log("Minor error showing all columns: " + uiError); }
   SpreadsheetApp.flush(); // Ensure all columns are shown before attempting to hide specific ones. This is important for the logic below.
   var startHideRange = null;
   for (var i = 1; i <= maxCols; i++) {
     if (!allowedNumsSet.has(i)) {
       if (startHideRange === null) { startHideRange = i; }
     } else {
       if (startHideRange !== null) {
         sheet.hideColumns(startHideRange, i - startHideRange);
         startHideRange = null;
       }
     }
   }
   if (startHideRange !== null) {
     sheet.hideColumns(startHideRange, maxCols - startHideRange + 1);
   }
   SpreadsheetApp.flush(); // Ensure all hide operations are completed before the function exits.
 } catch (error) {
   Logger.log("Error in setView: " + error);
   SpreadsheetApp.getUi().alert("Erro ao definir a visão: " + error.message); // UI: Alert text (PT-PT)
 }
}

/**
 * Converts column letter(s) (e.g., "A", "Z", "AA") to its 1-based numeric index.
 * @param {String} letter The column letter(s).
 * @return {Number} The 1-based column index (e.g., A=1, Z=26, AA=27).
 */
function columnLetterToNumber(letter) {
  var num = 0;
  var letters = letter.toUpperCase();
  for (var i = 0; i < letters.length; i++) {
    num = num * 26 + (letters.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return num;
}

// --- Project Information Function ---

/**
 * Gets project info from the metadata sheet.
 * @return {object} Object with project info or an error object.
 */
function getProjectInfo() {
  Logger.log("getProjectInfo: Starting data fetch...");
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (!configSheet) { throw new Error("A folha '" + CONFIG_SHEET_NAME + "' não foi encontrada."); }
    var rangeNotation = "D2:D6";
    var infoValues = configSheet.getRange(rangeNotation).getValues();
    if (!Array.isArray(infoValues) || infoValues.length < 5 || !Array.isArray(infoValues[0])) {
        throw new Error("Estrutura de dados inesperada lida da folha de configuração.");
    }
    const projectInfo = {
      clientName: infoValues[0][0] || 'N/A', clientAddress: infoValues[1][0] || 'N/A',
      clientNif: infoValues[2][0] || 'N/A', clientEmail: infoValues[3][0] || 'N/A',
      floorplanUrl: infoValues[4][0] || '#'
    };
    Logger.log("getProjectInfo: Data processed. Returning: " + JSON.stringify(projectInfo));
    return projectInfo;
  } catch (error) {
    Logger.log("getProjectInfo: CATCH block error: " + error);
    return { error: true, message: "Erro ao ler '" + CONFIG_SHEET_NAME + "': " + error.message };
  }
}

// --- Quote Preparation Function --- ADDED

/**
 * Reads last quote number, increments it, updates metadata sheet,
 * populates quote sheet with new number, date, and client info.
 * @return {string} Success or error message.
 */
function prepareQuoteSheet() {
    Logger.log("prepareQuoteSheet: Starting...");
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
        var quoteSheet = ss.getSheetByName(QUOTE_SHEET_NAME);

        // Validate sheets
        if (!configSheet) { throw new Error("A folha de configuração '" + CONFIG_SHEET_NAME + "' não foi encontrada."); }
        if (!quoteSheet) { throw new Error("A folha de orçamento '" + QUOTE_SHEET_NAME + "' não foi encontrada."); }

        // --- Get and Increment Quote Number ---
        var lastQuoteNumCell = configSheet.getRange(LAST_QUOTE_NUMBER_CELL);
        var lastQuoteNum = lastQuoteNumCell.getValue();
        if (typeof lastQuoteNum !== 'number' || !Number.isInteger(lastQuoteNum)) {
            throw new Error("O valor na célula " + LAST_QUOTE_NUMBER_CELL + " da folha " + CONFIG_SHEET_NAME + " não é um número inteiro válido.");
        }
        var newQuoteNum = lastQuoteNum + 1;
        lastQuoteNumCell.setValue(newQuoteNum); // Update the number in metadata
        Logger.log("prepareQuoteSheet: Quote number incremented to " + newQuoteNum);

        // Format for display (e.g., 29/2025)
        var currentYear = new Date().getFullYear();
        var displayQuoteNum = newQuoteNum + "/" + currentYear;
        var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

        // --- Get Client Info ---
        var clientInfoRange = configSheet.getRange("D2:D4"); // Read Name, Address, NIF
        var clientInfoValues = clientInfoRange.getValues();
        var clientName = clientInfoValues[0][0] || '';
        var clientAddr1 = clientInfoValues[1][0] || ''; // Assuming address might be split or just one line needed
        var clientAddr2 = ""; // Assuming second address line might be needed - adjust if necessary
        var clientNif = clientInfoValues[2][0] || '';

        // --- Update Quote Sheet ---
        // Set values without changing formatting
        quoteSheet.getRange(QUOTE_NUMBER_TARGET_CELL).setValue(displayQuoteNum);
        quoteSheet.getRange(QUOTE_DATE_TARGET_CELL).setValue(currentDate);
        quoteSheet.getRange(QUOTE_CLIENT_NAME_CELL).setValue(clientName);
        quoteSheet.getRange(QUOTE_CLIENT_ADDR1_CELL).setValue(clientAddr1);
        quoteSheet.getRange(QUOTE_CLIENT_ADDR2_CELL).setValue(clientAddr2); // Set second address line if needed
        quoteSheet.getRange(QUOTE_CLIENT_NIF_CELL).setValue(clientNif);
        Logger.log("prepareQuoteSheet: Quote sheet updated with number, date, and client info.");

        SpreadsheetApp.flush(); // Apply changes
        return "Orçamento #" + displayQuoteNum + " preparado com dados do cliente."; // UI: Success (PT-PT)

    } catch (error) {
        Logger.log("Error in prepareQuoteSheet: " + error);
        return "Erro ao preparar orçamento: " + error.message; // UI: Error (PT-PT)
    }
}

// --- Column Chooser Backend Functions ---

/**
 * Gets all header texts from the HEADER_ROW of the TARGET_DATA_SHEET_NAME.
 * This function is intended for the Column Chooser feature.
 * @param {string} viewName The name of the view (currently unused, but for future flexibility).
 * @return {Array<string>} An array of header strings.
 * @throws {Error} If the target sheet is not found or headers cannot be read.
 */
function getViewHeaders(viewName) {
  Logger.log("getViewHeaders: Called for viewName = '" + viewName + "' (currently operates on TARGET_DATA_SHEET_NAME)");
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // This function currently assumes views always target the TARGET_DATA_SHEET_NAME.
    // Could be made more flexible if different views target different sheets.
    var sheet = ss.getSheetByName(TARGET_DATA_SHEET_NAME);
    if (!sheet) {
      throw new Error("Sheet '" + TARGET_DATA_SHEET_NAME + "' not found.");
    }

    var headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getMaxColumns());
    var headerValues = headerRange.getValues()[0]; // Get the first (and only) row of values

    // Filter out empty cells/values from the header row
    var filteredHeaders = headerValues.filter(function(header) {
      return typeof header === 'string' && header.trim() !== '';
    });

    Logger.log("getViewHeaders: Found headers: " + JSON.stringify(filteredHeaders));
    return filteredHeaders;

  } catch (error) {
    Logger.log("Error in getViewHeaders: " + error.message + " (Stack: " + error.stack + ")");
    throw new Error("Erro ao obter cabeçalhos para '" + viewName + "': " + error.message); // UI: Error (PT-PT)
  }
}

/**
 * Applies column visibility settings to the TARGET_DATA_SHEET_NAME.
 * This function is intended for the Column Chooser feature.
 * @param {string} viewName The name of the view being applied (for logging).
 * @param {Object} columnVisibilitySettings An object where keys are header names and values are booleans (true to show, false to hide).
 * @return {string} A success message or throws an error.
 * @throws {Error} If the sheet is not found or settings cannot be applied.
 */
function applyColumnVisibility(viewName, columnVisibilitySettings) {
  Logger.log("applyColumnVisibility: Called for viewName = '" + viewName + "' with settings: " + JSON.stringify(columnVisibilitySettings));
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // This function currently assumes views always target the TARGET_DATA_SHEET_NAME.
    var sheet = ss.getSheetByName(TARGET_DATA_SHEET_NAME);
    if (!sheet) {
      throw new Error("Sheet '" + TARGET_DATA_SHEET_NAME + "' not found.");
    }

    var maxCols = sheet.getMaxColumns();
    var headerRowValues = sheet.getRange(HEADER_ROW, 1, 1, maxCols).getValues()[0];

    // First, show all columns to reset previous states.
    Logger.log("applyColumnVisibility: Showing all columns (1 to " + maxCols + ") initially.");
    sheet.showColumns(1, maxCols);
    SpreadsheetApp.flush(); // Ensure columns are shown before specific hiding.

    // Iterate through all columns and hide based on settings.
    for (var i = 0; i < maxCols; i++) {
      var columnIndex = i + 1; // 1-based index
      var headerText = (typeof headerRowValues[i] === 'string') ? headerRowValues[i].trim() : '';

      if (headerText === '') { // Skip empty header cells (or treat as columns to hide by default)
        // Logger.log("applyColumnVisibility: Column " + columnIndex + " has empty header. Hiding by default.");
        // sheet.hideColumns(columnIndex); // Decided to only hide if explicitly false or not in settings.
        continue;
      }

      if (columnVisibilitySettings.hasOwnProperty(headerText)) {
        if (columnVisibilitySettings[headerText] === false) {
          Logger.log("applyColumnVisibility: Hiding column '" + headerText + "' (index " + columnIndex + ") as per settings.");
          sheet.hideColumns(columnIndex);
        } else {
          // Logger.log("applyColumnVisibility: Showing column '" + headerText + "' (index " + columnIndex + ") as per settings.");
          // Column should be visible, and we already showed all columns.
        }
      } else {
        // Default behavior: Hide columns not explicitly mentioned in settings.
        Logger.log("applyColumnVisibility: Hiding column '" + headerText + "' (index " + columnIndex + ") as it's not in visibility settings (default hide).");
        sheet.hideColumns(columnIndex);
      }
    }

    SpreadsheetApp.flush(); // Ensure all hide operations are completed.
    Logger.log("applyColumnVisibility: Successfully applied column visibility for view '" + viewName + "'.");
    return "Visibilidade das colunas atualizada para a visão '" + viewName + "'."; // UI: Success (PT-PT)

  } catch (error) {
    Logger.log("Error in applyColumnVisibility: " + error.message + " (Stack: " + error.stack + ")");
    throw new Error("Erro ao aplicar visibilidade de colunas para '" + viewName + "': " + error.message); // UI: Error (PT-PT)
  }
}

// --- Automatic Edit Handling ---
/**
 * Handles edits made to the spreadsheet. Runs automatically via the onEdit simple trigger.
 * Logic includes:
 * 1. Formatting two-line headers in HEADER_ROW (Bold/11pt, Regular/9pt).
 * 2. Bolding the entire row if a specific value is entered in Column H (below header).
 * 3. Adding a warning note to Column R if the date in Column Q is earlier than Column R (below header).
 * @param {Object} e The event object passed by the onEdit trigger.
 */
function onEdit(e) {
  // Exit if the event object or range is missing.
  if (!e || !e.range) { return; }

  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();
  var value = e.value; // The new value entered
  var oldValue = e.oldValue; // The value before the edit

  // --- Logic 0: Two-Line Header Formatting ---
  // This logic is intended to run only on the target data sheet.
  if (sheet.getName() !== TARGET_DATA_SHEET_NAME) {
    // Logger.log("onEdit Header Formatting: Edit on sheet '" + sheet.getName() + "', skipping as it's not '" + TARGET_DATA_SHEET_NAME + "'.");
    return;
  }
  // Check if the edit was in the designated header row and is a single cell.
  if (row === HEADER_ROW && range.getNumRows() === 1 && range.getNumColumns() === 1) {
    // Avoid infinite loop: Check if value actually changed. Rich text edits might re-trigger onEdit.
    // We compare simple string value; complex rich text comparison is harder.
    if (value === oldValue) {
       // Logger.log("Header cell edit detected, but value didn't change. Skipping format.");
       return;
    }
    try {
        // Check if the new value is a string and contains a newline.
        if (typeof value === 'string' && value.includes('\n')) {
            const parts = value.split('\n', 2);
            const line1 = parts[0];
            const line2 = parts.length > 1 ? parts[1] : '';

            // Build the Rich Text value using globally defined styles.
            const richText = SpreadsheetApp.newRichTextValue()
              .setText(value)
              .setTextStyle(0, line1.length, BOLD_STYLE_11)
              .setTextStyle(line1.length + 1, value.length, REGULAR_STYLE_9)
              .build();

            // Set the rich text value back to the cell.
            range.setRichTextValue(richText);
            Logger.log("Applied two-line header format to " + range.getA1Notation());
            // Important: Return here to prevent other logic running on header edit
            return;
        } else if (typeof value === 'string') {
             // Optional: If it's a single line, maybe reset to default bold?
             // This might interfere if user *wants* non-bold single line headers.
             // range.setFontWeight("bold").setFontSize(11).setFontFamily(FONT_FAMILY).setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setFontSize(11).setFontFamily(FONT_FAMILY).build());
        }
    } catch (error) {
        Logger.log("Error in onEdit (Header Formatting, Row " + row + "): " + error);
    }
    // If header formatting was applied or attempted, exit onEdit for this event.
    return;
  } // End Header Formatting Logic


  // --- Logic 1: Bold Row based on Column H status (Only run if NOT header row) ---
  // Ensure this logic only runs on the target data sheet.
  if (sheet.getName() !== TARGET_DATA_SHEET_NAME) {
    // Logger.log("onEdit Bold Row: Edit on sheet '" + sheet.getName() + "', skipping as it's not '" + TARGET_DATA_SHEET_NAME + "'.");
    return;
  }
  var targetColumnH = 8; // Column H is the 8th column.
  // Run only if a single cell in Column H was edited AND it's below the header row.
  if (row > HEADER_ROW && col === targetColumnH && range.getNumRows() === 1 && range.getNumColumns() === 1) {
    // Check if value changed to avoid unnecessary formatting on minor edits
    if (value === oldValue) return;
    try {
      var cellValueH = range.getValue(); // Use getValue() for consistency
      var triggerWords = ["Aprovado", "Por Definir | M", "Por Desenhar | M", "Por Orçamentar | M", "Por Aprovar | C", "Por Aprovar | M", "Por Levantar"];
      var shouldBold = triggerWords.includes(cellValueH);
      var lastColumn = sheet.getLastColumn();
      var currentWeight = sheet.getRange(row, 1).getFontWeight();
      var targetWeight = shouldBold ? "bold" : "normal";
      if (currentWeight !== targetWeight) {
         sheet.getRange(row, 1, 1, lastColumn).setFontWeight(targetWeight);
      }
    } catch (error) { Logger.log("Error in onEdit (Bold Logic, Row " + row + "): " + error); }
     // Decide if editing column H should stop further onEdit checks (like Q/R)
     // return; // Uncomment this line if an edit in H should *not* also trigger the Q/R check in the same event.
     // Both "Bold Row" and "Date Comparison" logic can run for the same edit if the conditions for both are met.
  } // End Bold Row Logic


  // --- Logic 2: Date Comparison Note (Columns Q & R) (Only run if NOT header row) ---
  // Ensure this logic only runs on the target data sheet.
  if (sheet.getName() !== TARGET_DATA_SHEET_NAME) {
    // Logger.log("onEdit Date Comparison: Edit on sheet '" + sheet.getName() + "', skipping as it's not '" + TARGET_DATA_SHEET_NAME + "'.");
    return;
  }
  var targetColumnQ = 17; // Column Q is the 17th column.
  var targetColumnR = 18; // Column R is the 18th column.
  // Run only if editing Q or R in row 2 or below
  if (row > HEADER_ROW && (col === targetColumnQ || col === targetColumnR)) {
    // Check if value changed
     if (value === oldValue) return;
    try {
      var cellQ = sheet.getRange(row, targetColumnQ);
      var cellR = sheet.getRange(row, targetColumnR);
      var dateQ = cellQ.getValue();
      var dateR = cellR.getValue();
      var noteCell = cellR;
      var noteMsg = "";

      if (dateQ instanceof Date && dateR instanceof Date) {
        var timeDifference = dateQ.getTime() - dateR.getTime();
        var dayDifference = Math.floor(timeDifference / (1000 * 60 * 60 * 24));
        if (dayDifference < 0) {
          noteMsg = "⚠️ Atrasado por " + Math.abs(dayDifference) + " dias!"; // UI: Note text (PT-PT)
        }
      }
      if (noteCell.getNote() !== noteMsg) {
         noteCell.setNote(noteMsg);
      }
    } catch (error) { Logger.log("Error in onEdit (Date Logic, Row " + row + "): " + error); }
  } // End Date Logic
} // End of onEdit function


// --- PDF Generation Function ---
/**
 * Generates a PDF file of the "Summary" view (specific columns) of the active sheet.
 * Applies basic formatting (header color, font) to a temporary sheet used for export.
 * Saves the generated PDF to a designated folder in Google Drive.
 * @return {string} The URL of the saved PDF file in Google Drive, or an error message string starting with "Erro:".
 */
function generateSummaryPdf() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = null; // Reference to the temporary sheet for cleanup.
  try {
    var sourceSheet = ss.getSheetByName(TARGET_DATA_SHEET_NAME);
    if (!sourceSheet) { throw new Error("A folha '" + TARGET_DATA_SHEET_NAME + "' não foi encontrada."); }
    var sourceSheetName = sourceSheet.getName();
    var ssId = ss.getId(); // Spreadsheet ID needed for export URL.
    var targetSheetFont = "Reddit Sans Condensed"; // Font to apply in temp sheet.
    var headerBackgroundColor = "#4285F4"; // Header background color (e.g., standard Google Blue).
    var headerFontColor = "#FFFFFF"; // Header font color (White).

    // 1. Define target columns for the PDF content.
    var targetColumns = ["A", "B", "C", "D", "E", "F", "H", "N", "P", "Q"];
    // Get 0-based indices for these columns using the helper function.
    var targetIndices = getColumnIndices(sourceSheet, targetColumns);

    // 2. Read all data from the source sheet (using display values).
    var allDataRange = sourceSheet.getDataRange();
    var allValues = allDataRange.getDisplayValues();
    if (allValues.length === 0) { throw new Error("A folha ativa está vazia."); } // Error message in PT-PT

    // 3. Filter headers and data rows based on target columns.
    var headers = allValues[0]; // First row is headers.
    var filteredHeaders = targetIndices.map(index => headers[index]); // Extract target headers.
    var filteredData = [];
    // Iterate through data rows (starting from index 1).
    for (var i = 1; i < allValues.length; i++) {
      var row = allValues[i];
      // Include row only if the first column (POS) is not empty.
      if (row[0] && row[0].toString().trim() !== "") {
         // Extract data for target columns, handling potential short rows.
         var filteredRow = targetIndices.map(index => (index < row.length ? row[index] : ""));
         filteredData.push(filteredRow);
      }
    }
    // Check if any valid data rows were found.
    if (filteredData.length === 0) { throw new Error("Não foram encontrados dados válidos para incluir no PDF."); } // Error message in PT-PT

    // 4. Create a temporary sheet for PDF export.
    var tempSheetName = "TempPDF_" + new Date().getTime(); // Unique temporary name.
    tempSheet = ss.insertSheet(tempSheetName);

    // 5. Populate and format the temporary sheet.
    var headerRange = tempSheet.getRange(1, 1, 1, filteredHeaders.length);
    headerRange.setValues([filteredHeaders]); // Write headers.
    var dataRange = tempSheet.getRange(2, 1, filteredData.length, filteredHeaders.length);
    dataRange.setValues(filteredData); // Write data.

    // Apply formatting to headers.
    headerRange.setBackground(headerBackgroundColor)
      .setFontColor(headerFontColor)
      .setFontWeight("bold")
      .setFontFamily(targetSheetFont)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    // Apply formatting to data cells.
    dataRange.setFontFamily(targetSheetFont)
      .setVerticalAlignment("top"); // Align top for potentially wrapped text.

    // Auto-resize columns based on content.
    for (var j = 1; j <= filteredHeaders.length; j++) {
      tempSheet.autoResizeColumn(j);
    }
    SpreadsheetApp.flush(); // Ensure formatting is applied before export.

    // 6. Construct the PDF export URL.
    var sheetId = tempSheet.getSheetId(); // Get GID of the temporary sheet.
    // URL parameters define PDF options (A4, landscape, fit to width, margins, gridlines).
    var pdfUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
                 `format=pdf&gid=${sheetId}&size=a4&portrait=false&fitw=true&scale=4&` +
                 `top_margin=0.50&bottom_margin=0.50&left_margin=0.50&right_margin=0.50&` +
                 `gridlines=true&printtitle=false&sheetnames=false&fzr=false&` +
                 `horizontal_alignment=CENTER&vertical_alignment=TOP`;

    // 7. Fetch the PDF content using UrlFetchApp with OAuth token.
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(pdfUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true // Handle non-200 responses manually.
    });
    var responseCode = response.getResponseCode();
    // Check if PDF generation was successful.
    if (responseCode !== 200) {
      Logger.log("Error fetching PDF URL. Code: " + responseCode + ". Response: " + response.getContentText());
      throw new Error("Erro (" + responseCode + ") ao gerar o PDF a partir do Google Sheets."); // Error message in PT-PT
    }
    var blob = response.getBlob(); // Get PDF content as a Blob.

    // 8. Save the PDF Blob to Google Drive.
    var pdfFileName = `Sumario_${sourceSheetName}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.pdf`;
    blob.setName(pdfFileName);

    // Find or create the designated folder in Drive.
    var folders = DriveApp.getFoldersByName(PDF_FOLDER_NAME);
    var pdfFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PDF_FOLDER_NAME);

    // Remove existing file with the same name in the target folder to avoid duplicates.
    var existingFiles = pdfFolder.getFilesByName(pdfFileName);
    if (existingFiles.hasNext()) {
        existingFiles.next().setTrashed(true); // Move old file to trash.
    }
    // Create the new PDF file.
    var pdfFile = pdfFolder.createFile(blob);
    var fileUrl = pdfFile.getUrl(); // Get the shareable URL of the new file.
    Logger.log("PDF generated and saved: " + fileUrl);

    // 9. Return the URL of the saved PDF file.
    return fileUrl;

  } catch (error) {
    Logger.log("Error in generateSummaryPdf: " + error + " (Stack: " + error.stack + ")");
    // Return an error message string to be displayed in the sidebar.
    return "Erro: " + error.message; // Error message in PT-PT
  } finally {
    // 10. Clean up: Delete the temporary sheet regardless of success or failure.
    if (tempSheet) {
      try { ss.deleteSheet(tempSheet); } catch (e) { Logger.log("Error deleting temporary sheet '" + (tempSheet ? tempSheet.getName() : 'undefined') + "': " + e); }
    }
  }
}

function generateQuotePdf() { // ADDED
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var quoteSheet = ss.getSheetByName(QUOTE_SHEET_NAME);
    if (!quoteSheet) { throw new Error("A folha de orçamento '" + QUOTE_SHEET_NAME + "' não foi encontrada."); }

    var ssId = ss.getId();
    var sheetId = quoteSheet.getSheetId();
    var quoteNumber = quoteSheet.getRange(QUOTE_NUMBER_TARGET_CELL).getDisplayValue().replace("/", "-"); // Get quote number for filename
    var clientName = quoteSheet.getRange(QUOTE_CLIENT_NAME_CELL).getDisplayValue();

    // PDF Export URL - A4 Portrait, actual size (no fit-to-width), gridlines off (common for quotes)
    var pdfUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
                 `format=pdf&` +
                 `gid=${sheetId}&` +
                 `size=a4&` +                   // A4 Size
                 `portrait=true&` +             // Portrait orientation
                 `fitw=false&` +                // Do NOT fit to width (use actual size/scaling)
                 `scale=1&` +                   // Scale 1 = 100% ? (Might need adjustment) - Use sheet's print settings if possible? No direct way.
                 `top_margin=0.75&bottom_margin=0.75&left_margin=0.70&right_margin=0.70&` + // Standard margins
                 `gridlines=false&` +           // Usually no gridlines on quotes
                 `printtitle=false&sheetnames=false&fzr=false&` + // Standard options
                 `horizontal_alignment=CENTER&vertical_alignment=TOP`;

    // Fetch PDF Blob and Save to Drive
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(pdfUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var responseCode = response.getResponseCode();
    if (responseCode !== 200) { throw new Error("Erro (" + responseCode + ") ao gerar o PDF do orçamento."); }

    var blob = response.getBlob();
    var pdfFileName = `Orcamento_${quoteNumber}_${clientName}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.pdf`;
    pdfFileName = pdfFileName.replace(/[/\\?%*:|"<>]/g, '-'); // Sanitize filename
    blob.setName(pdfFileName);

    // Find or create the target folder
    var folders = DriveApp.getFoldersByName(PDF_FOLDER_NAME);
    var pdfFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PDF_FOLDER_NAME);

    // Remove existing file with same name
    var existingFiles = pdfFolder.getFilesByName(pdfFileName);
    if (existingFiles.hasNext()) { existingFiles.next().setTrashed(true); }

    var pdfFile = pdfFolder.createFile(blob);
    var fileUrl = pdfFile.getUrl();
    Logger.log("PDF Orçamento gerado e guardado: " + fileUrl);

    return fileUrl; // Return URL to sidebar

  } catch (error) {
    Logger.log("Error in generateQuotePdf: " + error + " (Stack: " + error.stack + ")");
    return "Erro ao gerar PDF do orçamento: " + error.message; // UI: Error (PT-PT)
  }
  // No temporary sheet needed for this version as we export the existing sheet
}

// --- Uppercase Function ---
/**
 * Converts all text content in the active sheet (from row 2 downwards) to uppercase.
 * Skips non-string values.
 * @return {string} A success or error message for the sidebar.
 */
function convertSheetToUppercase() {
    Logger.log("convertSheetToUppercase: Starting conversion...");
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var headerRowsToSkip = 1; // Skip the first row (header)
        var dataRange = sheet.getDataRange();
        var startRow = headerRowsToSkip + 1; // Data starts at row 2 (1-based index)
        var numRows = dataRange.getNumRows() - headerRowsToSkip;
        var numCols = dataRange.getNumColumns();

        // Check if there are any data rows to process
        if (numRows <= 0) {
            Logger.log("convertSheetToUppercase: No data rows found below header.");
            return "Nenhuma linha de dados encontrada abaixo do cabeçalho para converter."; // UI: Message (PT-PT)
        }

        // Get the range containing only the data rows (excluding header)
        var dataOnlyRange = sheet.getRange(startRow, 1, numRows, numCols);
        var values = dataOnlyRange.getValues();
        var changed = false; // Flag to track if any changes were made

        // Loop through data rows and columns
        for (var i = 0; i < values.length; i++) {
            for (var j = 0; j < values[i].length; j++) {
                // Check if the cell value is a string
                if (typeof values[i][j] === 'string') {
                    var originalValue = values[i][j];
                    var upperValue = originalValue.toUpperCase();
                    // Only update if the value actually changes
                    if (originalValue !== upperValue) {
                        values[i][j] = upperValue;
                        changed = true;
                    }
                }
            }
        }

        // Write the modified values back only if changes were made
        if (changed) {
            dataOnlyRange.setValues(values);
            Logger.log("convertSheetToUppercase: Conversion complete. Changes applied.");
            return "Texto da folha convertido para maiúsculas!"; // UI: Success message (PT-PT)
        } else {
            Logger.log("convertSheetToUppercase: No text found requiring conversion.");
            return "Nenhum texto encontrado que necessite de conversão."; // UI: No changes message (PT-PT)
        }

    } catch (error) {
        Logger.log("Error in convertSheetToUppercase: " + error);
        return "Erro ao converter para maiúsculas: " + error.message; // UI: Error message (PT-PT)
    }
}


// --- Helper Functions ---
/**
 * Helper function to get 0-based column indices from an array of column letters.
 * @param {Sheet} sheet The Google Sheet object to get ranges from.
 * @param {Array<string>} columnLetters Array of column letters (e.g., ["A", "C", "F"]).
 * @return {Array<number>} Array of corresponding 0-based column indices (e.g., [0, 2, 5]).
 * @throws {Error} If the sheet object is invalid or a column letter is invalid.
 */
function getColumnIndices(sheet, columnLetters) {
  // Validate the sheet input.
  if (!sheet || typeof sheet.getRange !== 'function') {
    throw new Error("Invalid sheet object provided to getColumnIndices.");
  }
  // Map each letter to its 0-based index.
  return columnLetters.map(letter => {
    try {
      // getColumn() returns 1-based index, subtract 1 for 0-based.
      return sheet.getRange(letter + "1").getColumn() - 1;
    } catch(e) {
      Logger.log("Error getting index for column letter '" + letter + "': " + e);
      throw new Error("Letra de coluna inválida: " + letter); // Error message in PT-PT
    }
  });
}

/*
==================================================================================
 CHANGELOG
==================================================================================
 * 2025-04-25 (Gemini): Initial creation with sidebar, view switching, info panel.
 * 2025-04-25 (Gemini): Added PDF generation function (`generateSummaryPdf`).
 * 2025-04-25 (Gemini): Added dialog prompt logic (`showSidebarPromptDialog`, `OpenSidebarPrompt.html`).
 * 2025-04-25 (Gemini): Reverted `onOpen` to only create menu, calling `showSidebar` directly from menu item.
                       Removed test alert from `onOpen`. Dialog function (`showSidebarPromptDialog`) kept but unused by default.
 * 2025-04-25 (Gemini): Standardized comments to English, UI strings to PT-PT. Cleaned up code structure. Added this changelog.
 * 2025-04-25 (Gemini): Re-integrated two-line header formatting into `onEdit`. Added `convertSheetToUppercase` function.
 * 2025-04-25 (Gemini): Restricted `setView` functions to only run on `TARGET_DATA_SHEET_NAME`.
==================================================================================
*/

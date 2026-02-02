import { QueryColumn, SelectedEntity, JoinedEntity } from '../types';

/**
 * Gets the formatted value for export from a data row
 */
export const getExportValue = (
  row: any,
  col: QueryColumn,
  entities: (SelectedEntity | JoinedEntity)[]
): string => {
  // Resolve the object holding the value (traverse alias/joins)
  let targetObj = row;

  if (col.entityAlias !== 'main') {
    const path: string[] = [];
    let currentAlias = col.entityAlias;

    // Reconstruct path from alias back to main
    while (currentAlias !== 'main') {
      const entity = entities.find(e => e.alias === currentAlias) as JoinedEntity;
      if (!entity) break;
      path.unshift(entity.navigationPropertyName);
      currentAlias = entity.parentAlias;
    }

    // Traverse
    for (const prop of path) {
      if (targetObj && targetObj[prop]) {
        targetObj = targetObj[prop];
      } else {
        targetObj = null;
        break;
      }
    }
  }

  if (!targetObj) return '';

  const resolveRawValue = (source: any) => {
    let val = source?.[col.attribute];
    if ((val === undefined || val === null) && col.logicalName) {
      const lookupKey = `_${col.logicalName}_value`;
      if (lookupKey !== col.attribute) {
        val = source?.[lookupKey];
      }
    }
    return val;
  };

  const resolveFormattedValue = (source: any) => {
    const formattedKey = `${col.attribute}@OData.Community.Display.V1.FormattedValue`;
    if (source && source[formattedKey] !== undefined) {
      return source[formattedKey];
    }
    if (col.logicalName) {
      const lookupFormattedKey = `_${col.logicalName}_value@OData.Community.Display.V1.FormattedValue`;
      if (lookupFormattedKey !== formattedKey && source && source[lookupFormattedKey] !== undefined) {
        return source[lookupFormattedKey];
      }
    }
    return undefined;
  };

  // Handle arrays (from OneToMany relationships)
  if (Array.isArray(targetObj)) {
    const values = targetObj.map(item => {
      const formatted = resolveFormattedValue(item);
      if (formatted !== undefined) return formatted;
      return resolveRawValue(item);
    }).filter(v => v !== null && v !== undefined);
    return values.join('; ');
  }

  // Single object value extraction
  const formatted = resolveFormattedValue(targetObj);
  if (formatted !== undefined) {
    return String(formatted);
  }

  const val = resolveRawValue(targetObj);
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') return JSON.stringify(val);
  return String(val);
};

/**
 * Escapes a value for CSV format (handles commas, quotes, newlines)
 */
const escapeCSV = (value: string): string => {
  if (value.includes(',') || value.includes('"') || value.includes('\n') || value.includes('\r')) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
};

/**
 * Generates CSV content from query results
 */
export const generateCSV = (
  data: any[],
  columns: QueryColumn[],
  entities: (SelectedEntity | JoinedEntity)[]
): string => {
  // Header row
  const headers = columns.map(col => escapeCSV(col.displayName));
  const lines = [headers.join(',')];

  // Data rows
  for (const row of data) {
    const values = columns.map(col => escapeCSV(getExportValue(row, col, entities)));
    lines.push(values.join(','));
  }

  return lines.join('\r\n');
};

/**
 * Checks if a column name suggests it contains phone number data
 */
const isPhoneColumn = (columnName: string): boolean => {
  const lower = columnName.toLowerCase();
  return lower.includes('phone') || lower.includes('telephone') || lower.includes('mobile') || lower.includes('fax');
};

/**
 * Determines if a value should be treated as a number in Excel.
 * Excludes long numeric strings (phone numbers, IDs) to prevent scientific notation.
 */
const shouldBeNumber = (value: string, columnName: string): boolean => {
  if (value === '' || value.includes(' ') || value.includes('-') || value.includes('+')) {
    return false;
  }
  if (isNaN(Number(value))) {
    return false;
  }
  // Don't treat long numbers as numeric (phone numbers, IDs, etc.)
  // Scientific notation kicks in for numbers > 11 digits
  if (value.replace(/[^0-9]/g, '').length > 11) {
    return false;
  }
  // Check column name for phone-related fields
  if (isPhoneColumn(columnName)) {
    return false;
  }
  return true;
};

/**
 * Generates Excel-compatible XML (SpreadsheetML) content from query results.
 * This format is natively supported by Excel and preserves column widths and formatting.
 */
export const generateExcelXML = (
  data: any[],
  columns: QueryColumn[],
  entities: (SelectedEntity | JoinedEntity)[],
  sheetName: string = 'Query Results',
  baseUrl?: string
): string => {
  const escapeXML = (str: string): string => {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  };

  // Build column definitions with auto-width (default 120 pixels)
  const columnDefs = columns.map(() => '<Column ss:AutoFitWidth="1" ss:Width="120"/>').join('\n      ');

  // Build header row
  const headerCells = columns.map(col =>
    `<Cell ss:StyleID="Header"><Data ss:Type="String">${escapeXML(col.displayName)}</Data></Cell>`
  ).join('\n          ');

  // Build data rows
  const dataRows = data.map(row => {
    const cells = columns.map(col => {
      const value = getExportValue(row, col, entities);

      // Check if this is a lookup field that could have a hyperlink
      let hyperlink: string | null = null;
      if (baseUrl && col.attribute.startsWith('_') && col.attribute.endsWith('_value')) {
        // This is a lookup field - try to build a hyperlink
        const lookupId = row[col.attribute];
        if (lookupId) {
          // Try to get the entity type from the lookup metadata
          const entityTypeKey = `${col.attribute}@Microsoft.Dynamics.CRM.lookuplogicalname`;
          const entityType = row[entityTypeKey];
          if (entityType) {
            hyperlink = `${baseUrl}/main.aspx?etn=${entityType}&id=${lookupId}&pagetype=entityrecord`;
          }
        }
      }

      // Also check for primary ID columns (first column is usually the record link)
      if (baseUrl && col.attribute.endsWith('id') && entities.length > 0) {
        const mainEntity = entities.find(e => e.alias === 'main' || e.alias === col.entityAlias);
        if (mainEntity && col.attribute === mainEntity.primaryIdAttribute) {
          const recordId = row[col.attribute];
          if (recordId) {
            hyperlink = `${baseUrl}/main.aspx?etn=${mainEntity.logicalName}&id=${recordId}&pagetype=entityrecord`;
          }
        }
      }

      // Determine data type - avoid scientific notation for phone numbers and long IDs
      const isNumeric = shouldBeNumber(value, col.displayName);
      const dataType = isNumeric ? 'Number' : 'String';

      if (hyperlink) {
        return `<Cell ss:HRef="${escapeXML(hyperlink)}"><Data ss:Type="${dataType}">${escapeXML(value)}</Data></Cell>`;
      }
      return `<Cell><Data ss:Type="${dataType}">${escapeXML(value)}</Data></Cell>`;
    }).join('\n          ');
    return `        <Row>\n          ${cells}\n        </Row>`;
  }).join('\n');

  return `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal">
      <Alignment ss:Vertical="Bottom"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Color="#000000"/>
    </Style>
    <Style ss:ID="Header">
      <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/>
      <Interior ss:Color="#0078D4" ss:Pattern="Solid"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#005A9E"/>
      </Borders>
    </Style>
  </Styles>
  <Worksheet ss:Name="${escapeXML(sheetName)}">
    <Table ss:DefaultRowHeight="15">
      ${columnDefs}
      <Row ss:StyleID="Default">
          ${headerCells}
        </Row>
${dataRows}
    </Table>
  </Worksheet>
</Workbook>`;
};

/**
 * Triggers a file download in the browser
 */
export const downloadFile = (content: string, filename: string, mimeType: string): void => {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);

  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
};

/**
 * Exports data to CSV file
 */
export const exportToCSV = (
  data: any[],
  columns: QueryColumn[],
  entities: (SelectedEntity | JoinedEntity)[],
  filename: string = 'export'
): void => {
  const csv = generateCSV(data, columns, entities);
  downloadFile(csv, `${filename}.csv`, 'text/csv;charset=utf-8;');
};

/**
 * Exports data to Excel XML file (SpreadsheetML format).
 * Uses .xls extension which opens directly in Excel. User may see a format warning
 * but clicking "Yes" opens the file correctly with all formatting preserved.
 */
export const exportToExcel = (
  data: any[],
  columns: QueryColumn[],
  entities: (SelectedEntity | JoinedEntity)[],
  filename: string = 'export',
  baseUrl?: string
): void => {
  const xml = generateExcelXML(data, columns, entities, filename, baseUrl);
  downloadFile(xml, `${filename}.xls`, 'application/vnd.ms-excel');
};

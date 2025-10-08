import type * as Excel from '@mui/x-internal-exceljs-fork';
import { RefObject } from '@mui/x-internals/types';
import {
  GridRowId,
  GridColDef,
  GridApi,
  ValueOptions,
  GRID_DATE_COL_DEF,
  GRID_DATETIME_COL_DEF,
  GridValidRowModel,
} from '@mui/x-data-grid-pro';
import {
  GridStateColDef,
  GridSingleSelectColDef,
  isObject,
  isSingleSelectColDef,
  gridHasColSpanSelector,
} from '@mui/x-data-grid/internals';
import { warnOnce } from '@mui/x-internals/warning';
import { ColumnsStylesInterface, GridExcelExportOptions } from '../gridExcelExportInterface';
import { GridPrivateApiPremium } from '../../../../models/gridApiPremium';
import {
  addColumnGroupingHeaders,
  addSerializedRowToWorksheet,
  createValueOptionsSheetIfNeeded,
  getExcelJs,
  SerializedColumns,
  SerializedRow,
  ValueOptionsData,
} from './utils';

export type { ExcelExportInitEvent } from './utils';

const getFormattedValueOptions = (
  colDef: GridSingleSelectColDef,
  row: GridValidRowModel,
  valueOptions: ValueOptions[],
  api: GridApi,
  callback: (value: any, index: number) => void,
) => {
  if (!colDef.valueOptions) {
    return;
  }
  const valueFormatter = colDef.valueFormatter;

  for (let i = 0; i < valueOptions.length; i += 1) {
    const option = valueOptions[i];
    let value: any;
    if (valueFormatter) {
      if (typeof option === 'object') {
        value = option.label;
      } else {
        value = String(colDef.valueFormatter!(option as never, row, colDef, { current: api }));
      }
    } else {
      value = typeof option === 'object' ? option.label : option;
    }
    callback(value, i);
  }
};

const commaRegex = /,/g;
const commaReplacement = 'CHAR(44)';

function getSerializedCellValue(
  row: any,
  column: GridStateColDef,
  colIndex: number,
  id: GridRowId,
  apiRef: RefObject<GridPrivateApiPremium>,
  defaultValueOptionsFormulae: { [field: string]: { address: string } },
  options: Pick<BuildExcelOptions, 'escapeFormulas'>,
  dataValidation: SerializedRow['dataValidation'],
  mergedCells: SerializedRow['mergedCells'],
  hasColSpan: boolean,
) {
  const api = apiRef.current;
  let cellValue: any;

  const colSpanInfo = hasColSpan ? api.unstable_getCellColSpanInfo(id, colIndex) : undefined;
  if (colSpanInfo && colSpanInfo.spannedByColSpan) {
    return undefined;
  }
  if (colSpanInfo && colSpanInfo.cellProps.colSpan > 1) {
    mergedCells.push({
      leftIndex: colIndex + 1,
      rightIndex: colIndex + colSpanInfo.cellProps.colSpan,
    });
  }

  switch (column.type) {
    case 'singleSelect': {
      const castColumn = column as GridSingleSelectColDef;
      if (typeof castColumn.valueOptions === 'function') {
        const valueOptions = castColumn.valueOptions({
          id,
          row,
          field: column.field,
        });

        let formulae: string = '"';
        getFormattedValueOptions(castColumn, row, valueOptions, api, (value, index) => {
          const formatted = value.toString().replace(commaRegex, commaReplacement);
          formulae += formatted;
          if (index < valueOptions.length - 1) {
            formulae += ',';
          }
        });
        formulae += '"';

        dataValidation[castColumn.field] = {
          type: 'list',
          allowBlank: true,
          formulae: [formulae],
        };
      } else {
        const address = defaultValueOptionsFormulae[column.field].address;
        dataValidation[castColumn.field] = {
          type: 'list',
          allowBlank: true,
          formulae: [address],
        };
      }

      const formattedValue = api.getRowFormattedValue(row, castColumn);
      if (process.env.NODE_ENV !== 'production') {
        if (String(formattedValue) === '[object Object]') {
          warnOnce([
            'MUI X: When the value of a field is an object or a `renderCell` is provided, the Excel export might not display the value correctly.',
            'You can provide a `valueFormatter` with a string representation to be used.',
          ]);
        }
      }
      if (isObject<{ label: any }>(formattedValue)) {
        return formattedValue?.label;
      }
      return formattedValue as any;
    }
    case 'boolean':
    case 'number':
      cellValue = api.getRowValue(row, column);
      break;
    case 'date':
    case 'dateTime': {
      const value = api.getRowValue(row, column) as Date;
      if (!value) {
        break;
      }
      const utcDate = new Date(
        Date.UTC(
          value.getFullYear(),
          value.getMonth(),
          value.getDate(),
          value.getHours(),
          value.getMinutes(),
          value.getSeconds(),
        ),
      );
      return utcDate;
    }
    case 'actions':
      break;
    default:
      cellValue = api.getRowFormattedValue(row, column);
      if (process.env.NODE_ENV !== 'production') {
        if (String(cellValue) === '[object Object]') {
          warnOnce([
            'MUI X: When the value of a field is an object or a `renderCell` is provided, the Excel export might not display the value correctly.',
            'You can provide a `valueFormatter` with a string representation to be used.',
          ]);
        }
      }
      break;
  }

  if (typeof cellValue === 'string' && options.escapeFormulas) {
    if (['=', '+', '-', '@', '\t', '\r'].includes(cellValue[0])) {
      cellValue = `'${cellValue}`;
    }
  }

  return cellValue;
}

/**
 * FIXME: This function mutates the colspan info, but colspan info assumes that the columns
 * passed to it are always consistent. In this case, the exported columns may differ from the
 * actual rendered columns.
 * The caller of this function MUST call `resetColSpan()` before and after usage.
 */
export const serializeRowUnsafe = (
  id: GridRowId,
  columns: GridStateColDef[],
  apiRef: RefObject<GridPrivateApiPremium>,
  defaultValueOptionsFormulae: { [field: string]: { address: string } },
  options: Pick<BuildExcelOptions, 'escapeFormulas'>,
): SerializedRow | SerializedRow[] => {
  const dataValidation: SerializedRow['dataValidation'] = {};
  const mergedCells: SerializedRow['mergedCells'] = [];

  const row = apiRef.current.getRow(id);
  const rowNode = apiRef.current.getRowNode(id);
  if (!row || !rowNode) {
    throw new Error(`No row with id #${id} found`);
  }
  const outlineLevel = rowNode.depth;
  const hasColSpan = gridHasColSpanSelector(apiRef);

  if (hasColSpan) {
    apiRef.current.calculateColSpan(id, 0, columns.length, columns);
  }

  // Split columns into parent and child
  const parentColumns = columns.filter((col) => !col.isExportChildColumn);
  const childColumns = columns.filter((col) => col.isExportChildColumn);

  // Parent row: only parent columns
  const parentSerializedRow: SerializedRow['row'] = {};
  parentColumns.forEach((column, colIndex) => {
    parentSerializedRow[column.field] = getSerializedCellValue(
      row,
      column,
      colIndex,
      id,
      apiRef,
      defaultValueOptionsFormulae,
      options,
      dataValidation,
      mergedCells,
      hasColSpan,
    );
  });

  /**
   * New approach: Flatten parent-child structure
   * For each child row, duplicate the parent data and add child data on the same level
   *
   * Example:
   * | Parent 1 | Parent 2 | Child 1 | Child 2 |
   * | data1    | data2    | child1  | child2  |
   * | data1    | data2    | child3  | child4  |
   * | data1    | data2    | child5  | child6  |
   */
  if (childColumns.length > 0) {
    const rows: SerializedRow[] = [];

    // Find the array field used by the child columns (e.g. "verksamhetsInnehall")
    const arrayField = childColumns[0].field.split('.')[0];
    const childArray = Array.isArray(row[arrayField]) ? row[arrayField] : [];

    // If there are no child rows, just export the parent row
    if (childArray.length === 0) {
      const parentRow: SerializedRow['row'] = {};

      // Add parent columns
      parentColumns.forEach((col) => {
        parentRow[col.field] = getSerializedCellValue(
          row,
          col,
          columns.indexOf(col),
          id,
          apiRef,
          defaultValueOptionsFormulae,
          options,
          dataValidation,
          mergedCells,
          hasColSpan,
        );
      });

      // Add empty child columns
      childColumns.forEach((col) => {
        parentRow[col.field] = '';
      });

      return {
        row: parentRow,
        dataValidation,
        outlineLevel,
        mergedCells,
        borderStyle: {
          bottom: { style: 'thin' },
        },
      };
    }

    // For each child row, create a complete row with parent + child data
    for (let i = 0; i < childArray.length; i += 1) {
      const combinedRow: SerializedRow['row'] = {};
      const childRowData = childArray[i];

      // Add parent columns (same data for each child row)
      parentColumns.forEach((col) => {
        combinedRow[col.field] = getSerializedCellValue(
          row,
          col,
          columns.indexOf(col),
          id,
          apiRef,
          defaultValueOptionsFormulae,
          options,
          dataValidation,
          mergedCells,
          hasColSpan,
        );
      });

      // Add child columns (specific data for this child row)
      childColumns.forEach((col) => {
        // If the field is dot-notated, use only the last part for the child row
        const fieldParts = col.field.split('.');
        const childField = fieldParts[fieldParts.length - 1];
        const patchedCol = { ...col, field: childField };

        const cellValue = getSerializedCellValue(
          childRowData,
          patchedCol,
          columns.indexOf(col),
          id,
          apiRef,
          defaultValueOptionsFormulae,
          options,
          dataValidation,
          mergedCells,
          hasColSpan,
        );
        combinedRow[col.field] = cellValue ?? '';
      });

      rows.push({
        row: combinedRow,
        dataValidation: {},
        outlineLevel: rowNode.depth,
        mergedCells: [],
      });
    }

    return rows;
  }

  // Regular row: all columns
  const serializedRow: SerializedRow['row'] = {};
  columns.forEach((column, colIndex) => {
    serializedRow[column.field] = getSerializedCellValue(
      row,
      column,
      colIndex,
      id,
      apiRef,
      defaultValueOptionsFormulae,
      options,
      dataValidation,
      mergedCells,
      hasColSpan,
    );
  });

  return {
    row: serializedRow,
    dataValidation,
    outlineLevel,
    mergedCells,
  };
};

const defaultColumnsStyles = {
  [GRID_DATE_COL_DEF.type as string]: { numFmt: 'dd.mm.yyyy' },
  [GRID_DATETIME_COL_DEF.type as string]: { numFmt: 'dd.mm.yyyy hh:mm' },
};

export const serializeColumn = (column: GridColDef, columnsStyles: ColumnsStylesInterface) => {
  const { field, type } = column;

  return {
    key: field,
    headerText: column.headerName ?? column.field,
    // Excel width must stay between 0 and 255 (https://support.microsoft.com/en-us/office/change-the-column-width-and-row-height-72f5e3cc-994d-43e8-ae58-9774a0905f46)
    // From the example of column width behavior (https://docs.microsoft.com/en-US/office/troubleshoot/excel/determine-column-widths#example-of-column-width-behavior)
    // a value of 10 corresponds to 75px. This is an approximation, because column width depends on the font-size
    width: Math.min(255, column.width ? column.width / 7.5 : 8.43),
    style: { ...(type && defaultColumnsStyles?.[type]), ...columnsStyles?.[field] },
  };
};

export function serializeColumns(
  columns: GridStateColDef[],
  styles: ColumnsStylesInterface,
): SerializedColumns {
  return columns.map((column) => serializeColumn(column, styles));
}

export async function getDataForValueOptionsSheet(
  columns: GridStateColDef[],
  valueOptionsSheetName: string,
  api: GridPrivateApiPremium,
): Promise<ValueOptionsData> {
  // Creates a temp worksheet to obtain the column letters
  const excelJS = await getExcelJs();
  const workbook: Excel.Workbook = new excelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');

  const record: Record<string, { values: (string | number)[]; address: string }> = {};
  const worksheetColumns: typeof worksheet.columns = [];

  for (let i = 0; i < columns.length; i += 1) {
    const column = columns[i];
    const isCandidateColumn = isSingleSelectColDef(column) && Array.isArray(column.valueOptions);
    if (!isCandidateColumn) {
      continue;
    }

    worksheetColumns.push({ key: column.field });
    worksheet.columns = worksheetColumns;

    const header = column.headerName ?? column.field;
    const values: any[] = [header];
    getFormattedValueOptions(
      column,
      {},
      column.valueOptions as Array<ValueOptions>,
      api,
      (value) => {
        values.push(value);
      },
    );

    const letter = worksheet.getColumn(column.field).letter;
    const address = `${valueOptionsSheetName}!$${letter}$2:$${letter}$${values.length}`;

    record[column.field] = { values, address };
  }

  return record;
}
interface BuildExcelOptions
  extends
    Pick<GridExcelExportOptions, 'exceljsPreProcess' | 'exceljsPostProcess'>,
    Pick<
      Required<GridExcelExportOptions>,
      'valueOptionsSheetName' | 'includeHeaders' | 'includeColumnGroupsHeaders' | 'escapeFormulas'
    > {
  columns: GridStateColDef[];
  rowIds: GridRowId[];
  columnsStyles?: ColumnsStylesInterface;
  customSerializeColumns?: typeof serializeColumns;
}

export async function buildExcel(
  options: BuildExcelOptions,
  apiRef: RefObject<GridPrivateApiPremium>,
): Promise<Excel.Workbook> {
  const {
    columns,
    rowIds,
    includeHeaders,
    includeColumnGroupsHeaders,
    valueOptionsSheetName = 'Options',
    exceljsPreProcess,
    exceljsPostProcess,
    customSerializeColumns,
    columnsStyles = {},
  } = options;

  const excelJS = await getExcelJs();
  const workbook: Excel.Workbook = new excelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');

  const serializedColumns = customSerializeColumns
    ? customSerializeColumns(columns, columnsStyles)
    : serializeColumns(columns, columnsStyles);
  worksheet.columns = serializedColumns;

  if (exceljsPreProcess) {
    await exceljsPreProcess({
      workbook,
      worksheet,
    });
  }

  if (includeColumnGroupsHeaders) {
    const columnGroupPaths = columns.reduce<Record<string, string[]>>((acc, column) => {
      acc[column.field] = apiRef.current.getColumnGroupPath(column.field);
      return acc;
    }, {});

    addColumnGroupingHeaders(
      worksheet,
      serializedColumns,
      columnGroupPaths,
      apiRef.current.getAllGroupDetails(),
    );
  }

  if (includeHeaders) {
    // Add headers for ALL columns (parent + child)
    worksheet.addRow(columns.map((column) => column.headerName ?? column.field));
  }

  const valueOptionsData = await getDataForValueOptionsSheet(
    columns,
    valueOptionsSheetName,
    apiRef.current,
  );
  createValueOptionsSheetIfNeeded(valueOptionsData, valueOptionsSheetName, workbook);

  apiRef.current.resetColSpan();
  rowIds.forEach((id) => {
    const serialized = serializeRowUnsafe(id, columns, apiRef, valueOptionsData, options);
    if (Array.isArray(serialized)) {
      serialized.forEach((row) => addSerializedRowToWorksheet(row, worksheet));
    } else {
      addSerializedRowToWorksheet(serialized, worksheet);
    }
  });
  apiRef.current.resetColSpan();

  if (exceljsPostProcess) {
    await exceljsPostProcess({
      workbook,
      worksheet,
    });
  }

  return workbook;
}

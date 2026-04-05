export type MemoryMode = 'constant-memory' | 'low-memory' | 'standard';

export interface XlsxStyleDefinition {
  bold?: boolean;
  italic?: boolean;
  fontColor?: string;
  backgroundColor?: string;
  fontSize?: number;
  border?: 'thin' | 'none';
  align?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  textWrap?: boolean;
  numFormat?: string;
}

export interface XlsxTheme {
  styles?: Record<string, XlsxStyleDefinition>;
}

export interface XlsxColumn {
  key: string;
  title: string;
  width: number;
  dataStyle?: string;
  headerStyle?: string;
  successLabelList?: string[];
  dangerLabelList?: string[];
}

export interface XlsxEmptyCell {
  type: 'empty';
  styleList?: string[];
}

export interface XlsxStringCell {
  type: 'string';
  value?: string | null;
  styleList?: string[];
}

export interface XlsxNumberCell {
  type: 'number';
  value: number;
  styleList?: string[];
}

export interface XlsxFormulaCell {
  type: 'formula';
  formula: string;
  styleList?: string[];
}

export interface XlsxLinkCell {
  type: 'link';
  url: string;
  label?: string;
  styleList?: string[];
}

export interface XlsxDateCell {
  type: 'date';
  value?: string | null;
  styleList?: string[];
}

export type XlsxCellObject =
  | XlsxEmptyCell
  | XlsxStringCell
  | XlsxNumberCell
  | XlsxFormulaCell
  | XlsxLinkCell
  | XlsxDateCell;

export type XlsxCellValue = string | number | boolean | null | XlsxCellObject;

export interface XlsxRow {
  values: Record<string, XlsxCellValue>;
}

export interface XlsxMergeRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  value?: string | null;
  styleList?: string[];
}

export interface XlsxConditionalFormat {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  formula: string;
  style: string;
}

export interface XlsxSheet {
  sheetName: string;
  headerRowIndex?: number;
  headerRowHeight?: number;
  freezeHeaderRow?: boolean;
  autoFilter?: boolean;
  columns: XlsxColumn[];
  rows: XlsxRow[];
  mergeRangeList?: XlsxMergeRange[];
  conditionalFormatList?: XlsxConditionalFormat[];
}

export interface Workbook {
  sheetList: XlsxSheet[];
}

export interface BufferTarget {
  kind: 'buffer';
}

export interface FileTarget {
  kind: 'file';
  path: string;
}

export type RenderTarget = BufferTarget | FileTarget;

export interface RenderWorkbookParams {
  workbook: Workbook;
  theme: XlsxTheme;
  target: RenderTarget;
  memoryMode?: MemoryMode;
  tempDir?: string;
}

export interface RenderWorkbookFileResult {
  kind: 'file';
  path: string;
  bytes: number;
}

export declare function renderWorkbook(
  params: RenderWorkbookParams,
): Promise<Buffer | RenderWorkbookFileResult>;

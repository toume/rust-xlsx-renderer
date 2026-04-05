# rust-xlsx-renderer

[![CI](https://github.com/toume/rust-xlsx-renderer/actions/workflows/ci.yml/badge.svg)](https://github.com/toume/rust-xlsx-renderer/actions/workflows/ci.yml)

`rust-xlsx-renderer` is a native Node.js package for generating `.xlsx` files with
[`rust_xlsxwriter`](https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/) and
[`napi-rs`](https://napi.rs/).

License: Apache-2.0.

Repository: [toume/rust-xlsx-renderer](https://github.com/toume/rust-xlsx-renderer).

## Installation

For consumers:

```bash
npm install rust-xlsx-renderer
```

For local development in this repository:

```bash
npm install
npm run build
```

## Trust Boundary

This package assumes the application chooses safe filesystem locations.

- Do not pass untrusted user input directly into `target.path`.
- Do not pass untrusted user input directly into `tempDir`.
- Prefer `target: { kind: 'file' }` for large exports and let the application own the output directory.
- Use `target: { kind: 'buffer' }` only for bounded exports that are safe to hold fully in memory.

## Usage

```js
const { renderWorkbook } = require('rust-xlsx-renderer');

const workbook = {
  sheetList: [
    {
      sheetName: 'Report',
      columns: [
        { key: 'name', title: 'Name', width: 24 },
        { key: 'amount', title: 'Amount', width: 18, dataStyle: 'currency' },
        {
          key: 'status',
          title: 'Status',
          width: 18,
          successLabelList: ['Paid'],
          dangerLabelList: ['Overdue'],
        },
      ],
      rows: [
        {
          values: {
            name: 'Alice',
            amount: 1200.5,
            status: {
              type: 'string',
              value: 'Overdue',
              styleList: ['dangerText'],
            },
          },
        },
      ],
    },
  ],
};

const theme = {
  styles: {
    header: {
      bold: true,
      backgroundColor: '#D9E2F3',
    },
    currency: {
      numFormat: '#,##0.00 [$EUR]',
      align: 'right',
    },
    success: {
      fontColor: '#00875A',
      backgroundColor: '#E3FCEF',
    },
    danger: {
      fontColor: '#DE350B',
      backgroundColor: '#FFEBE6',
    },
    dangerText: {
      fontColor: '#DE350B',
    },
  },
};

await renderWorkbook({
  workbook,
  theme,
  target: { kind: 'file', path: '/tmp/report.xlsx' },
  memoryMode: 'constant-memory',
});
```

To render to memory instead:

```js
const buffer = await renderWorkbook({
  workbook,
  theme,
  target: { kind: 'buffer' },
});
```

## Public API

The package exposes a single async function:

```ts
renderWorkbook(params: RenderWorkbookParams): Promise<Buffer | RenderWorkbookFileResult>
```

### `RenderWorkbookParams`

```ts
type RenderWorkbookParams = {
  workbook: Workbook;
  theme: XlsxTheme;
  target: { kind: 'buffer' } | { kind: 'file'; path: string };
  memoryMode?: 'constant-memory' | 'low-memory' | 'standard';
  tempDir?: string;
};
```

### Workbook shape

```ts
type Workbook = {
  sheetList: XlsxSheet[];
};

type XlsxSheet = {
  sheetName: string;
  headerRowIndex?: number;
  headerRowHeight?: number;
  freezeHeaderRow?: boolean;
  autoFilter?: boolean;
  columns: XlsxColumn[];
  rows: XlsxRow[];
  mergeRangeList?: XlsxMergeRange[];
  conditionalFormatList?: XlsxConditionalFormat[];
};

type XlsxColumn = {
  key: string;
  title: string;
  width: number;
  dataStyle?: string;
  headerStyle?: string;
  successLabelList?: string[];
  dangerLabelList?: string[];
};

type XlsxRow = {
  values: Record<string, XlsxCellValue>;
};
```

### Cell values

Primitive values are written directly:

- `string`
- `number`
- `boolean`
- `null`

Structured cell objects are also supported:

```ts
type XlsxCellObject =
  | { type: 'empty'; styleList?: string[] }
  | { type: 'string'; value?: string | null; styleList?: string[] }
  | { type: 'number'; value: number; styleList?: string[] }
  | { type: 'formula'; formula: string; styleList?: string[] }
  | { type: 'link'; url: string; label?: string; styleList?: string[] }
  | { type: 'date'; value?: string | null; styleList?: string[] };
```

`type: 'date'` expects an ISO-like string. If the date cannot be parsed, the raw string is written.

### Theme shape

```ts
type XlsxTheme = {
  styles?: Record<string, XlsxStyleDefinition>;
};

type XlsxStyleDefinition = {
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
};
```

Supported style properties:

- `bold`
- `italic`
- `fontColor`: 6-digit RGB, with or without `#`
- `backgroundColor`: 6-digit RGB, with or without `#`
- `fontSize`
- `border`: `thin` or `none`
- `align`: `left`, `center`, `right`
- `verticalAlign`: `top`, `middle`, `bottom`
- `textWrap`
- `numFormat`: Excel number format string

Unknown style names and unsupported style values are ignored. Missing `header` and `cell`
styles fall back to an empty Excel format.

### Built-in conditional styling

If a column declares `successLabelList` or `dangerLabelList`, the renderer applies formula-based
conditional formats using the theme styles named `success` and `danger`.

If either style is missing when required, rendering fails with `INVALID_CONDITIONAL_STYLE`.

### Manual ranges and conditional formats

```ts
type XlsxMergeRange = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  value?: string | null;
  styleList?: string[];
};

type XlsxConditionalFormat = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  formula: string;
  style: string;
};
```

## Memory modes

- `constant-memory`: preferred for large sequential exports
- `low-memory`: useful when repeated strings would make `constant-memory` outputs too large
- `standard`: fallback for exports that need the most flexibility

For large HTTP downloads, prefer `target: { kind: 'file' }` and stream the file afterwards.
`target: { kind: 'buffer' }` builds the entire workbook in memory before returning it.

## Error handling

Native errors are serialized with stable codes in the message payload.

Typical codes:

- `INVALID_PAYLOAD`
- `INVALID_MEMORY_MODE`
- `INVALID_TARGET`
- `INVALID_CELL`
- `INVALID_CELL_TYPE`
- `INVALID_LINK`
- `INVALID_NUMBER`
- `INVALID_CONDITIONAL_STYLE`
- `IO_ERROR`
- `XLSX_WRITE_FAILED`

## Quality checks

```bash
npm run format:check
npm run test:rust
npm run test:ts
npm run pack:dry-run
```

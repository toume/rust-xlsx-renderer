'use strict';

let nativeBinding;

function getNativeBinding() {
  if (!nativeBinding) {
    nativeBinding = require('./index.js');
  }

  return nativeBinding;
}

async function renderWorkbookWithBinding(binding, params) {
  if (!params || !params.workbook || !params.theme || !params.target || !params.target.kind) {
    throw new TypeError('renderWorkbook expects a payload with workbook, theme, and target.kind.');
  }

  const payload = JSON.stringify(params);
  if (params.target.kind === 'buffer') {
    return binding.renderWorkbookToBuffer(payload);
  }

  if (params.target.kind === 'file') {
    return binding.renderWorkbookToFile(payload);
  }

  throw new TypeError(`Unsupported XLSX target kind: ${params.target.kind}`);
}

async function renderWorkbook(params) {
  return renderWorkbookWithBinding(getNativeBinding(), params);
}

module.exports = {
  renderWorkbook,
  __private: {
    getNativeBinding,
    renderWorkbookWithBinding,
  },
};

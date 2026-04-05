const { __private } = require('./wrapper.js') as {
  __private: {
    renderWorkbookWithBinding: (
      binding: {
        renderWorkbookToBuffer: (payload: string) => Promise<Buffer>;
        renderWorkbookToFile: (payload: string) => Promise<{ kind: 'file'; path: string; bytes: number }>;
      },
      params: unknown,
    ) => Promise<unknown>;
  };
};

describe('rust-xlsx-renderer wrapper', () => {
  const workbook = {
    sheetList: [
      {
        sheetName: 'Test',
        columns: [{ key: 'name', title: 'Name', width: 20 }],
        rows: [{ values: { name: 'Alice' } }],
      },
    ],
  };
  const theme = { styles: {} };

  it('routes buffer targets to renderWorkbookToBuffer', async () => {
    const renderWorkbookToBuffer = jest.fn().mockResolvedValue(Buffer.from('xlsx'));
    const renderWorkbookToFile = jest.fn();

    const result = await __private.renderWorkbookWithBinding(
      { renderWorkbookToBuffer, renderWorkbookToFile },
      {
        workbook,
        theme,
        target: { kind: 'buffer' },
        memoryMode: 'constant-memory',
      },
    );

    expect(Buffer.isBuffer(result)).toBe(true);
    expect(renderWorkbookToBuffer).toHaveBeenCalledWith(
      JSON.stringify({
        workbook,
        theme,
        target: { kind: 'buffer' },
        memoryMode: 'constant-memory',
      }),
    );
    expect(renderWorkbookToFile).not.toHaveBeenCalled();
  });

  it('routes file targets to renderWorkbookToFile', async () => {
    const renderWorkbookToBuffer = jest.fn();
    const renderWorkbookToFile = jest.fn().mockResolvedValue({
      kind: 'file',
      path: '/tmp/export.xlsx',
      bytes: 42,
    });

    const result = await __private.renderWorkbookWithBinding(
      { renderWorkbookToBuffer, renderWorkbookToFile },
      {
        workbook,
        theme,
        target: { kind: 'file', path: '/tmp/export.xlsx' },
        memoryMode: 'low-memory',
      },
    );

    expect(result).toEqual({
      kind: 'file',
      path: '/tmp/export.xlsx',
      bytes: 42,
    });
    expect(renderWorkbookToFile).toHaveBeenCalledWith(
      JSON.stringify({
        workbook,
        theme,
        target: { kind: 'file', path: '/tmp/export.xlsx' },
        memoryMode: 'low-memory',
      }),
    );
    expect(renderWorkbookToBuffer).not.toHaveBeenCalled();
  });

  it('rejects invalid payloads before calling the native binding', async () => {
    const renderWorkbookToBuffer = jest.fn();
    const renderWorkbookToFile = jest.fn();

    await expect(
      __private.renderWorkbookWithBinding(
        { renderWorkbookToBuffer, renderWorkbookToFile },
        { target: { kind: 'buffer' } },
      ),
    ).rejects.toThrow('renderWorkbook expects a payload with workbook, theme, and target.kind.');

    expect(renderWorkbookToBuffer).not.toHaveBeenCalled();
    expect(renderWorkbookToFile).not.toHaveBeenCalled();
  });

  it('rejects payloads without theme before calling the native binding', async () => {
    const renderWorkbookToBuffer = jest.fn();
    const renderWorkbookToFile = jest.fn();

    await expect(
      __private.renderWorkbookWithBinding(
        { renderWorkbookToBuffer, renderWorkbookToFile },
        {
          workbook,
          target: { kind: 'buffer' },
        },
      ),
    ).rejects.toThrow('renderWorkbook expects a payload with workbook, theme, and target.kind.');

    expect(renderWorkbookToBuffer).not.toHaveBeenCalled();
    expect(renderWorkbookToFile).not.toHaveBeenCalled();
  });

  it('rejects unsupported targets before calling the native binding', async () => {
    const renderWorkbookToBuffer = jest.fn();
    const renderWorkbookToFile = jest.fn();

    await expect(
      __private.renderWorkbookWithBinding(
        { renderWorkbookToBuffer, renderWorkbookToFile },
        {
          workbook,
          theme,
          target: { kind: 'stream' },
        },
      ),
    ).rejects.toThrow('Unsupported XLSX target kind');

    expect(renderWorkbookToBuffer).not.toHaveBeenCalled();
    expect(renderWorkbookToFile).not.toHaveBeenCalled();
  });
});

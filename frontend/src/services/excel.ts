/* global Office, Excel */
declare const Excel: any;
declare const Office: any;

export function parseAddress(address: string): { sheet: string; cell: string } {
  if (address.includes('!')) {
    const idx = address.indexOf('!');
    return {
      sheet: address.slice(0, idx).replace(/^'|'$/g, ''),
      cell: address.slice(idx + 1),
    };
  }
  return { sheet: '', cell: address };
}

export async function highlightCell(address: string, color: string): Promise<void> {
  const { sheet, cell } = parseAddress(address);
  await Excel.run(async (ctx) => {
    const ws = sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();
    const range = ws.getRange(cell);
    range.format.fill.color = color;
    await ctx.sync();
  });
}

export async function navigateTo(address: string): Promise<void> {
  const { sheet, cell } = parseAddress(address);
  await Excel.run(async (ctx) => {
    const ws = sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();
    ws.activate();
    const range = ws.getRange(cell);
    range.select();
    await ctx.sync();
  });
}

export async function clearAllHighlights(): Promise<void> {
  await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items');
    await ctx.sync();
    for (const ws of sheets.items) {
      const used = ws.getUsedRange(true);
      used.load('address');
      await ctx.sync();
      try {
        used.format.fill.clear();
        await ctx.sync();
      } catch {
        // empty sheet — skip
      }
    }
  });
}

export async function getSelectedCell(): Promise<string> {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load(['address', 'worksheet']);
    const ws = range.worksheet;
    ws.load('name');
    await ctx.sync();
    const addr = range.address;
    const sheetName = ws.name;
    const cellPart = addr.includes('!') ? addr.split('!')[1] : addr;
    return `${sheetName}!${cellPart}`;
  });
}

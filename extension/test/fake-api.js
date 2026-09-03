// An in-memory Sheets API: getGrid(ranges) answers from a cell table and records
// every range it was asked for, so tests can assert what an audit fetched.
// cells: { 'Sheet': { row: { col: {v, d, f} } } }  (d defaults to String(v), f to '')
function makeFakeApi(S, meta, cells) {
  const byId = {};
  for (const s of meta.sheets) byId[s.properties.title] = s.properties.sheetId;
  const calls = [];
  return {
    calls,
    getGrid(ranges) {
      calls.push(ranges.slice());
      const sheets = [];
      for (const rs of ranges) {
        const bang = rs.lastIndexOf('!');
        const sheetName = rs.substring(1, bang - 1).split("''").join("'");
        const r = S.a1ToRect(rs.substring(bang + 1), sheetName);
        const rowData = [];
        for (let row = r.r1; row <= r.r2; row++) {
          const values = [];
          for (let col = r.c1; col <= r.c2; col++) {
            const c = cells[sheetName] && cells[sheetName][row] && cells[sheetName][row][col];
            if (!c) { values.push({}); continue; }
            const v = {};
            if (typeof c.v === 'number') v.effectiveValue = { numberValue: c.v };
            else if (typeof c.v === 'boolean') v.effectiveValue = { boolValue: c.v };
            else if (c.v !== '' && c.v !== undefined) v.effectiveValue = { stringValue: String(c.v) };
            v.formattedValue = c.d !== undefined ? c.d : (c.v === undefined ? '' : String(c.v));
            if (c.f) v.userEnteredValue = { formulaValue: c.f };
            else if (v.effectiveValue) v.userEnteredValue = v.effectiveValue;
            values.push(v);
          }
          rowData.push({ values });
        }
        sheets.push({ properties: { sheetId: byId[sheetName] }, data: [{ startRow: r.r1 - 1, startColumn: r.c1 - 1, rowData }] });
      }
      return Promise.resolve({ sheets });
    }
  };
}
module.exports = { makeFakeApi };

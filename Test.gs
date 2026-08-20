/**
 * Test.gs — assert-based self-check for the parts that do not need a document.
 *
 * In Apps Script: run `runTests()` and read the execution log.
 * On a workstation:  node run_tests.js
 */

function assertEq(actual, expected, label) {
  var a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) throw new Error(label + '\n  expected: ' + e + '\n  actual:   ' + a);
}

function refKeys(formula) {
  return extractRefs(formula).map(function (r) {
    return (r.kind === 'ref' ? '' : r.kind + ':') + (r.sheet ? r.sheet + '!' : '') + r.a1;
  });
}

function runTests() {
  var t = [];

  // --- column arithmetic -------------------------------------------------
  t.push(function () {
    assertEq(colToNum('A'), 1, 'colToNum A');
    assertEq(colToNum('Z'), 26, 'colToNum Z');
    assertEq(colToNum('AA'), 27, 'colToNum AA');
    assertEq(numToCol(1), 'A', 'numToCol 1');
    assertEq(numToCol(27), 'AA', 'numToCol 27');
    assertEq(numToCol(colToNum('ZZZ')), 'ZZZ', 'numToCol round trip');
  });

  // --- the cases a regex gets wrong --------------------------------------
  t.push(function () {
    assertEq(refKeys('=IF(A1="B2",C3,D4)'), ['A1', 'C3', 'D4'],
      'a cell address inside a string literal is not a reference');
  });
  t.push(function () {
    assertEq(refKeys("='O''Connor'!A1+'My Sheet'!B2"), ["O'Connor!A1", 'My Sheet!B2'],
      'quoted sheet names, with a doubled quote escape');
  });
  t.push(function () {
    assertEq(refKeys('=LOG10(A1)'), ['A1'],
      'LOG10 is a valid A1 address but the "(" makes it a function');
    assertEq(refKeys('=LOG10+1'), ['LOG10'],
      'without the "(" the same text really is cell LOG10');
  });
  t.push(function () {
    assertEq(refKeys('=SUM(TaxRate,A1)'), ['name:TaxRate', 'A1'],
      'a bare identifier is a named-range candidate');
  });
  t.push(function () {
    assertEq(refKeys('=SUM($A$1:$B$10)'), ['A1:B10'], 'absolute markers are stripped');
    assertEq(refKeys('=SUM(A:A)+SUM(3:3)'), ['A:A', '3:3'], 'whole column and whole row');
  });
  t.push(function () {
    assertEq(refKeys('=SUM(Sheet1!A1:Sheet1!B2)'), ['Sheet1!A1:B2'],
      'a repeated sheet qualifier on the right side of a range');
  });
  t.push(function () {
    assertEq(refKeys('=A1&"quote"" inside"&B2'), ['A1', 'B2'],
      'an escaped quote does not end the string early');
  });
  t.push(function () {
    assertEq(refKeys('=1E10+A1'), ['A1'], 'scientific notation is a number, not a reference');
    assertEq(refKeys('=IF(TRUE,A1,B1)'), ['A1', 'B1'], 'TRUE/FALSE are not named ranges');
  });
  t.push(function () {
    var refs = extractRefs('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/KEY/edit","Data!A1:B5")');
    assertEq(refs.length, 1, 'IMPORTRANGE produces one row');
    assertEq(refs[0].kind, 'external', 'IMPORTRANGE is external');
    assertEq(refs[0].a1, 'Data!A1:B5', 'IMPORTRANGE carries its range argument');
  });
  t.push(function () {
    assertEq(refKeys('=A1+'), ['A1'], 'a truncated formula does not throw');
    assertEq(refKeys(''), [], 'an empty formula yields nothing');
  });

  // --- argument splitting -------------------------------------------------
  t.push(function () {
    assertEq(splitCallArgs('=OFFSET(A1,1,2)', 0).args, ['A1', '1', '2'], 'plain arguments');
    assertEq(splitCallArgs('=OFFSET(A1,MAX(1,2),3)', 0).args, ['A1', 'MAX(1,2)', '3'],
      'a comma inside a nested call is not an argument separator');
    assertEq(splitCallArgs('=INDIRECT("A,1"&B2)', 0).args, ['"A,1"&B2'],
      'a comma inside a string literal is not an argument separator');
  });
  t.push(function () {
    var calls = findCalls('=OFFSET(A1,1,0)+INDEX(B1:B9,2)', ['OFFSET', 'INDEX', 'INDIRECT']);
    assertEq(calls.length, 2, 'both dynamic calls are found');
    assertEq(calls[0].text, 'OFFSET(A1,1,0)', 'the exact call text is preserved');
    assertEq(calls[1].text, 'INDEX(B1:B9,2)', 'the second call text is preserved');
  });
  t.push(function () {
    var calls = findCalls('=SUM(OFFSET(A1,1,0))', ['OFFSET']);
    assertEq(calls.length, 1, 'a nested dynamic call is found');
    assertEq(calls[0].text, 'OFFSET(A1,1,0)', 'nested call text stops at its own bracket');
  });

  // --- rectangles ---------------------------------------------------------
  t.push(function () {
    assertEq(a1ToRect('B3', 'S'), { sheetName: 'S', r1: 3, c1: 2, r2: 3, c2: 2 }, 'single cell');
    assertEq(a1ToRect('B3:D9', 'S'), { sheetName: 'S', r1: 3, c1: 2, r2: 9, c2: 4 }, 'range');
    assertEq(a1ToRect('D9:B3', 'S'), { sheetName: 'S', r1: 3, c1: 2, r2: 9, c2: 4 },
      'a reversed range is normalised');
    assertEq(a1ToRect('C:C', 'S').r2, MAX_ROW, 'whole column reaches the last row');
    assertEq(a1ToRect('5:5', 'S').c2, MAX_COL, 'whole row reaches the last column');
    assertEq(a1ToRect('not a ref', 'S'), null, 'garbage is rejected');
  });
  t.push(function () {
    assertEq(rectToA1(a1ToRect('B3', 'S')), 'B3', 'round trip single cell');
    assertEq(rectToA1(a1ToRect('B3:D9', 'S')), 'B3:D9', 'round trip range');
    assertEq(rectToA1(a1ToRect('C:C', 'S')), 'C:C', 'round trip whole column');
    assertEq(rectToA1(a1ToRect('5:5', 'S')), '5:5', 'round trip whole row');
  });
  t.push(function () {
    var r = a1ToRect('B2:D4', 'S');
    assertEq(rectContains(r, 'S', 3, 3), true, 'inside');
    assertEq(rectContains(r, 'S', 1, 3), false, 'above');
    assertEq(rectContains(r, 'Other', 3, 3), false, 'right cell, wrong sheet');
    assertEq(rectsOverlap(a1ToRect('A1:C3', 'S'), a1ToRect('C3:E5', 'S')), true, 'touching corner overlaps');
    assertEq(rectsOverlap(a1ToRect('A1:B2', 'S'), a1ToRect('C3:D4', 'S')), false, 'disjoint');
  });

  // --- the arithmetic used for OFFSET/INDEX arguments ---------------------
  t.push(function () {
    assertEq(arithmetic('1+2*3'), 7, 'precedence');
    assertEq(arithmetic('(1+2)*3'), 9, 'brackets');
    assertEq(arithmetic('-2+5'), 3, 'leading unary minus');
    assertEq(arithmetic('10/4'), 2.5, 'division');
    assertEq(arithmetic('1/0'), null, 'division by zero is unresolvable, not Infinity');
    assertEq(arithmetic('1+'), null, 'a dangling operator is rejected');
    assertEq(arithmetic('(1+2'), null, 'an unclosed bracket is rejected');
  });
  t.push(function () {
    assertEq(splitTopLevel('"a"&B1&"c"', '&'), ['"a"', 'B1', '"c"'], 'concatenation splits');
    assertEq(splitTopLevel('"a&b"&C1', '&'), ['"a&b"', 'C1'],
      'an ampersand inside a string is not a separator');
  });

  // --- regressions found in review ---------------------------------------
  t.push(function () {
    assertEq(refKeys('=LET(rate,A1,rate*2)'), ['A1'],
      'a LET-bound name is local, not a workbook named range');
    assertEq(refKeys('=LAMBDA(x,x+A1)(B1)'), ['A1', 'B1'],
      'a LAMBDA parameter is local too');
    assertEq(refKeys('=LET(rate,A1,rate*Growth)'), ['A1', 'name:Growth'],
      'a genuine named range beside a LET binding still comes through');
  });

  t.push(function () {
    var computed = extractRefs('=IMPORTRANGE(A1,B1)');
    assertEq(computed[0].url, '',
      'a computed IMPORTRANGE argument is not presented as a URL');
    assertEq(computed.filter(function (r) { return r.kind === 'ref'; })
      .map(function (r) { return r.a1; }), ['A1', 'B1'],
      'its arguments are still reported as ordinary precedents');
  });

  t.push(function () {
    assertEq(refKeys('=A1+a1'), ['A1'], 'A1 addresses are case-insensitive when deduplicating');
  });

  t.push(function () {
    assertEq(arithmetic('2*-3'), -6, 'unary minus binds tighter than multiplication');
    assertEq(arithmetic('2--1'), 3, 'a subtraction followed by a unary minus');
    assertEq(arithmetic('2/-2'), -1, 'a unary minus after a division');
    assertEq(arithmetic('-2*-3'), 6, 'unary signs on both operands');
  });

  t.push(function () {
    var noValues = { get: function () { return null; } };
    var nested = resolveDynamicTargets('=OFFSET(OFFSET(A1,1,0),1,0)', 'S', noValues, {});
    assertEq(rectToA1(nested[0].rect), 'A3',
      'a nested OFFSET base is resolved, not approximated by its first reference');
    var bogus = resolveDynamicTargets('=OFFSET(SOMEFN(A1),1,0)', 'S', noValues, {});
    assertEq(bogus[0].rect, null,
      'an expression that merely contains a reference is not treated as that reference');
    var outside = resolveDynamicTargets('=INDEX(A1:A2,3)', 'S', noValues, {});
    assertEq(outside[0].rect, null, 'an index past the range is #REF!, not the next cell');
    assertEq(rectToA1(resolveDynamicTargets('=INDEX(A1:B3,,2)', 'S', noValues, {})[0].rect), 'B1:B3',
      'an omitted INDEX row means the whole column of the range');
  });

  t.push(function () {
    assertEq(refKeys("='Sheet 1'!A1:'Sheet 1'!B10"), ['Sheet 1!A1:B10'],
      'a quoted qualifier repeated on the right of a range stays one range');
    assertEq(refKeys('=1E+10*A1'), ['A1'],
      'scientific notation with a sign is one number, not the name "1E"');
    assertEq(refKeys('=1E-3+B2'), ['B2'], 'the same with a negative exponent');
  });

  t.push(function () {
    var noValues = { get: function () { return null; } };
    assertEq(rectToA1(resolveDynamicTargets('=OFFSET(A1,,2)', 'S', noValues, {})[0].rect), 'C1',
      'an omitted OFFSET argument is zero');
    assertEq(rectToA1(resolveDynamicTargets('=INDEX(A1:E1,3)', 'S', noValues, {})[0].rect), 'C1',
      'a lone INDEX argument on a single-row range selects a column');
    assertEq(rectToA1(resolveDynamicTargets('=INDEX(A1:E1,0)', 'S', noValues, {})[0].rect), 'A1:E1',
      'index 0 on a single-row range is the whole row');
  });

  t.push(function () {
    var refs = extractRefs('=OFFSET(Inputs!B3,1,0)');
    assertEq(refs[0].raw, 'Inputs!B3',
      'raw keeps the sheet qualifier, or substituting it into the argument text fails');
    assertEq(extractRefs("='My Sheet'!A1")[0].raw, "'My Sheet'!A1",
      'a quoted qualifier is kept verbatim');
    assertEq(extractRefs('=A1')[0].raw, 'A1', 'an unqualified reference is unchanged');
  });

  t.push(function () {
    assertEq(replaceRefToken('A1+A10', 'A1', '5'), '5+A10',
      'substituting A1 must not corrupt A10');
    assertEq(replaceRefToken('A10+A1', 'A1', '5'), 'A10+5',
      'the same, with the longer reference first');
    assertEq(replaceRefToken('A1', 'A1', '5'), '5', 'a lone reference is replaced');
  });

  t.push(function () {
    // A stub value source: dynamic resolution must not need a live document
    // when every argument is a literal.
    var noValues = { get: function () { return null; } };
    var names = {};

    var offset = resolveDynamicTargets('=OFFSET(B2,1,2)', 'S', noValues, names);
    assertEq(offset.length, 1, 'one OFFSET call');
    assertEq(rectToA1(offset[0].rect), 'D3',
      'OFFSET(B2,1,2) is one row down and two columns right');

    var sized = resolveDynamicTargets('=OFFSET(B2,0,0,3,2)', 'S', noValues, names);
    assertEq(rectToA1(sized[0].rect), 'B2:C4', 'height and width size the result');

    var idx2d = resolveDynamicTargets('=INDEX(A1:C10,2)', 'S', noValues, names);
    assertEq(rectToA1(idx2d[0].rect), 'A2:C2',
      'INDEX over a 2-D range with no column argument is the whole row');

    var idx1d = resolveDynamicTargets('=INDEX(A1:A10,3)', 'S', noValues, names);
    assertEq(rectToA1(idx1d[0].rect), 'A3',
      'INDEX over a single column collapses to one cell');

    var idxCol = resolveDynamicTargets('=INDEX(A1:C10,0,2)', 'S', noValues, names);
    assertEq(rectToA1(idxCol[0].rect), 'B1:B10', 'row 0 means the whole column');

    var indirect = resolveDynamicTargets('=INDIRECT("Data!B7")', 'S', noValues, names);
    assertEq(indirect[0].rect.sheetName, 'Data', 'INDIRECT resolves the sheet');
    assertEq(rectToA1(indirect[0].rect), 'B7', 'INDIRECT resolves the address');

    var hard = resolveDynamicTargets('=OFFSET(A1,MATCH(1,B:B,0),0)', 'S', noValues, names);
    assertEq(hard[0].rect, null, 'an unsupported argument stays unresolved');
    assertEq(hard[0].reason.length > 0, true, 'and carries a reason');
  });

  t.push(function () {
    var noValues = { get: function () { return null; } };
    var named = { 'SALES': { name: 'Sales', rect: rect('Data', 2, 3, 20, 3) } };
    var viaName = resolveDynamicTargets('=INDEX(Sales,4)', 'S', noValues, named);
    assertEq(viaName[0].rect.sheetName, 'Data', 'a named range resolves to its own sheet');
    assertEq(rectToA1(viaName[0].rect), 'C5', 'INDEX(Sales,4) is the 4th row of Sales');
  });

  var failures = [];
  for (var i = 0; i < t.length; i++) {
    try { t[i](); } catch (e) { failures.push(e.message); }
  }

  var summary = (t.length - failures.length) + '/' + t.length + ' test groups passed';
  if (failures.length) {
    summary += '\n\nFAILURES:\n' + failures.join('\n\n');
    Logger.log(summary);
    throw new Error(summary);
  }
  Logger.log(summary);
  return summary;
}

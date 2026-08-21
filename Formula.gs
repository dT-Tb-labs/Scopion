/**
 * Formula.gs — pure formula text parsing. No SpreadsheetApp calls live here,
 * so every function in this file is testable from Test.gs without a document.
 *
 * Port note: the Excel original found precedents with Range.NavigateArrow, which
 * Apps Script has no equivalent of. Everything a precedent walk used to give us
 * has to come out of the formula string instead, which is what this file does.
 */

var ERROR_LITERALS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#ERROR!'];
var MAX_COL = 18278; // Google Sheets column limit (ZZZ)
var MAX_ROW = 10000000;

var TOK = {
  STRING: 'STRING',
  SHEET: 'SHEET',     // a sheet qualifier, i.e. the "Sheet1!" part
  REF: 'REF',         // an A1 reference or range
  FUNC: 'FUNC',       // an identifier immediately followed by "("
  NAME: 'NAME',       // a bare identifier: named-range candidate
  NUMBER: 'NUMBER',
  PUNCT: 'PUNCT'      // ( ) , ; and operators
};

function isDigit(c) { return c >= '0' && c <= '9'; }
function isAlpha(c) { return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'); }
function isIdentChar(c) { return isAlpha(c) || isDigit(c) || c === '_' || c === '.' || c === '$'; }

function colToNum(letters) {
  var n = 0;
  for (var i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
  }
  return n;
}

function numToCol(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Parse one side of a reference ("$A$1", "A", "1") into {col, row}, either of
 * which may be null for a whole-column ("A") or whole-row ("1") reference.
 * Returns null when the text is not a reference at all.
 */
function parseRefPart(text) {
  var m = /^\$?([A-Za-z]{1,3})?\$?([0-9]{1,8})?$/.exec(text);
  if (!m || (!m[1] && !m[2])) return null;
  var col = m[1] ? colToNum(m[1]) : null;
  var row = m[2] ? parseInt(m[2], 10) : null;
  if (col !== null && (col < 1 || col > MAX_COL)) return null;
  if (row !== null && (row < 1 || row > MAX_ROW)) return null;
  return { col: col, row: row };
}

/** True when text is a self-contained A1 reference such as "B7" or "$B$7". */
function isFullCellRef(text) {
  var p = parseRefPart(text);
  return !!(p && p.col !== null && p.row !== null);
}

/**
 * Tokenize a formula. The leading "=" is optional.
 *
 * A character scanner rather than a regex, because the same characters mean
 * different things depending on the state we are in:
 *   =IF(A1="B2",...)        the "B2" inside the string literal is not a reference
 *   ='O''Connor'!A1         the doubled quote is one literal quote, not a close
 *   =LOG10(A1)              LOG10 is a valid A1 address; only the "(" says otherwise
 *   =SUM(TaxRate)           TaxRate is a named range, not an address
 */
function tokenize(formula) {
  var tokens = [];
  if (!formula) return tokens;
  var s = String(formula);
  if (s.charAt(0) === '=') s = s.substring(1);
  var i = 0;

  while (i < s.length) {
    var c = s.charAt(i);

    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue; }

    // String literal, "" is an escaped quote.
    if (c === '"') {
      var str = '';
      i++;
      while (i < s.length) {
        if (s.charAt(i) === '"') {
          if (s.charAt(i + 1) === '"') { str += '"'; i += 2; continue; }
          i++;
          break;
        }
        str += s.charAt(i); i++;
      }
      tokens.push({ type: TOK.STRING, value: str });
      continue;
    }

    // Quoted sheet name, '' is an escaped quote. Only a qualifier when "!" follows.
    if (c === "'") {
      var name = '';
      var j = i + 1;
      while (j < s.length) {
        if (s.charAt(j) === "'") {
          if (s.charAt(j + 1) === "'") { name += "'"; j += 2; continue; }
          j++;
          break;
        }
        name += s.charAt(j); j++;
      }
      if (s.charAt(j) === '!') {
        tokens.push({ type: TOK.SHEET, value: name, raw: s.substring(i, j) });
        i = j + 1;
      } else {
        // A stray quoted run that is not a sheet qualifier. Skip it rather than
        // letting its contents be scanned as references.
        tokens.push({ type: TOK.STRING, value: name });
        i = j;
      }
      continue;
    }

    // Error literals are constants, not references. Left to the identifier
    // scanner, "#N/A" becomes the named-range candidates N and A and
    // "#DIV/0!" becomes DIV — and a workbook that really defines such a name
    // would then show a precedent the formula never reads. Matched by literal
    // comparison rather than a regex: the pattern needs backslashes, and those
    // do not survive being moved between environments intact.
    if (c === '#') {
      var head = s.substring(i, i + 8).toUpperCase();
      var hit = null;
      for (var e = 0; e < ERROR_LITERALS.length; e++) {
        if (head.indexOf(ERROR_LITERALS[e]) === 0) { hit = ERROR_LITERALS[e]; break; }
      }
      if (hit) {
        i += hit.length;
        // "#REF!A1" is what a formula pointing at a deleted sheet looks like:
        // the address after it is meaningless, and reading it as a same-sheet
        // reference would invent a precedent.
        if (hit === '#REF!') {
          while (i < s.length && isIdentChar(s.charAt(i))) i++;
        }
        tokens.push({ type: TOK.NUMBER, value: hit });
        continue;
      }
    }

    if (isIdentChar(c)) {
      var start = i;
      var ident = '';
      while (i < s.length && isIdentChar(s.charAt(i))) { ident += s.charAt(i); i++; }

      // 1E+10 stops at the sign above; pull the exponent back in so it stays a
      // number instead of becoming the named-range candidate "1E".
      if (/^[0-9.]+[eE]$/.test(ident) && /[-+]/.test(s.charAt(i)) && isDigit(s.charAt(i + 1))) {
        ident += s.charAt(i); i++;
        while (i < s.length && isDigit(s.charAt(i))) { ident += s.charAt(i); i++; }
      }

      // Sheet qualifier: Sheet1!A1
      if (s.charAt(i) === '!') {
        tokens.push({ type: TOK.SHEET, value: ident, raw: ident });
        i++;
        continue;
      }

      // Function call: the "(" is what distinguishes LOG10( from cell LOG10.
      if (s.charAt(i) === '(') {
        tokens.push({ type: TOK.FUNC, value: ident.toUpperCase(), start: start });
        continue; // leave "(" for the punct branch so bracket depth still counts it
      }

      // Range: A1:B2, A:A, 1:1. The right side may carry its own sheet qualifier
      // (Sheet1!A1:Sheet1!B2), which we drop — a range cannot span two sheets.
      if (s.charAt(i) === ':') {
        var k = i + 1;
        var rightIdent = '';
        // The right side may repeat the sheet qualifier, quoted or not
        // ('Sheet 1'!A1:'Sheet 1'!B10). A range cannot span two sheets, so the
        // repeat is skipped rather than parsed.
        if (s.charAt(k) === "'") {
          var q = k + 1;
          while (q < s.length) {
            if (s.charAt(q) === "'") {
              if (s.charAt(q + 1) === "'") { q += 2; continue; }
              q++;
              break;
            }
            q++;
          }
          if (s.charAt(q) === '!') k = q + 1;
        }
        while (k < s.length && isIdentChar(s.charAt(k))) { rightIdent += s.charAt(k); k++; }
        if (s.charAt(k) === '!') {
          k++;
          rightIdent = '';
          while (k < s.length && isIdentChar(s.charAt(k))) { rightIdent += s.charAt(k); k++; }
        }
        if (parseRefPart(ident) && parseRefPart(rightIdent)) {
          tokens.push({ type: TOK.REF, value: ident + ':' + rightIdent, start: start });
          i = k;
          continue;
        }
        // Not a range after all (e.g. a named range before a ":" typo). Fall through.
      }

      if (/^[0-9.]+$/.test(ident) || /^[0-9.]+[eE][+-]?[0-9]+$/.test(ident)) {
        tokens.push({ type: TOK.NUMBER, value: ident });
      } else if (isFullCellRef(ident)) {
        tokens.push({ type: TOK.REF, value: ident, start: start });
      } else if (/^(TRUE|FALSE)$/i.test(ident)) {
        tokens.push({ type: TOK.NUMBER, value: ident });
      } else {
        tokens.push({ type: TOK.NAME, value: ident, start: start });
      }
      continue;
    }

    tokens.push({ type: TOK.PUNCT, value: c });
    i++;
  }

  return tokens;
}

/**
 * Extract every reference a formula reads, as {sheet, a1, kind, raw}.
 * `sheet` is null when the reference is unqualified (same sheet as the formula).
 * `kind` is 'ref' | 'name' | 'external'.
 *
 * Dynamic targets (OFFSET/INDEX/INDIRECT) are NOT resolved here — Trace.gs does
 * that, because resolving them needs to read cell values.
 */
function extractRefs(formula) {
  var tokens = tokenize(formula);
  var locals = boundNames(formula);
  var out = [];
  var pendingSheet = null;
  var pendingSheetRaw = '';

  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];

    if (t.type === TOK.SHEET) {
      pendingSheet = t.value;
      pendingSheetRaw = t.raw;
      continue;
    }

    if (t.type === TOK.FUNC && t.value === 'IMPORTRANGE') {
      // IMPORTRANGE("<url or key>", "<Sheet!Range>") is the only cross-file read
      // Google Sheets has. It cannot be navigated into, so it is recorded as an
      // external row carrying its own arguments.
      // Only a quoted literal may be shown as the source. A computed argument
      // (=IMPORTRANGE(A1,B1)) must not be presented as a URL called "A1"; its
      // cells are still reported as ordinary precedents by the loop below.
      var call = splitCallArgs(formula, t.start);
      out.push({
        sheet: null,
        a1: stringLiteral(call.args[1]),
        kind: 'external',
        raw: call.args.length ? 'IMPORTRANGE(' + call.args.join(', ') + ')' : 'IMPORTRANGE()',
        url: stringLiteral(call.args[0]),
        pos: t.start
      });
      pendingSheet = null;
      continue;
    }

    if (t.type === TOK.REF) {
      out.push({
        sheet: pendingSheet,
        a1: t.value.replace(/\$/g, ''),
        kind: 'ref',
        raw: (pendingSheet !== null ? pendingSheetRaw + '!' : '') + t.value,
        pos: t.start
      });
      pendingSheet = null;
      continue;
    }

    if (t.type === TOK.NAME) {
      // A sheet-qualified bare identifier is not a named range; drop the qualifier.
      if (pendingSheet === null && !locals[t.value.toUpperCase()]) {
        out.push({ sheet: null, a1: t.value, kind: 'name', raw: t.value, pos: t.start });
      }
      pendingSheet = null;
      continue;
    }

    if (t.type !== TOK.PUNCT || (t.value !== '(' && t.value !== ')')) pendingSheet = null;
  }

  return dedupeRefs(out);
}

/**
 * Names bound by LET/LAMBDA inside this formula. They look exactly like named
 * ranges, so without this a workbook name colliding with a LET variable is
 * reported as a precedent that the formula never reads.
 *
 * Scope is deliberately ignored: a name bound anywhere in the formula is treated
 * as local throughout it. Erring towards "local" only ever drops a name that a
 * shadowed outer scope would have resolved, which is the quieter mistake.
 */
function boundNames(formula) {
  var bound = {};
  var calls = findCalls(formula, ['LET', 'LAMBDA']);
  for (var i = 0; i < calls.length; i++) {
    var args = calls[i].args;
    var limit = args.length - 1; // the last argument is the body
    for (var a = 0; a < limit; a++) {
      if (calls[i].name === 'LET' && a % 2 !== 0) continue; // LET alternates name, value
      var name = String(args[a]).trim();
      if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) bound[name.toUpperCase()] = true;
    }
  }
  return bound;
}

/** The text of a quoted string literal, or '' when the argument is computed. */
function stringLiteral(arg) {
  var text = String(arg === undefined ? '' : arg).trim();
  var m = /^"((?:[^"]|"")*)"$/.exec(text);
  return m ? m[1].split('""').join('"') : '';
}

function dedupeRefs(refs) {
  var seen = {};
  var out = [];
  for (var i = 0; i < refs.length; i++) {
    var key = refs[i].kind + '|' + (refs[i].sheet || '') + '|' + refs[i].a1.toUpperCase();
    if (seen[key]) continue;
    seen[key] = true;
    out.push(refs[i]);
  }
  return out;
}

/**
 * Split the argument list of the call whose FUNC token sits at `funcIndex`,
 * returning the raw source text of each top-level argument. Bracket- and
 * string-aware, so nested calls and commas inside literals stay intact.
 */
function splitCallArgs(formula, funcIndex) {
  var s = String(formula);
  if (s.charAt(0) === '=') s = s.substring(1);
  var open = s.indexOf('(', funcIndex);
  if (open < 0) return { args: [], end: -1 };

  var args = [];
  var cur = '';
  var depth = 1;
  var i = open + 1;
  var inStr = false;
  var inQuote = false;

  while (i < s.length) {
    var c = s.charAt(i);
    if (inStr) {
      cur += c;
      if (c === '"') {
        if (s.charAt(i + 1) === '"') { cur += '"'; i += 2; continue; }
        inStr = false;
      }
      i++;
      continue;
    }
    if (inQuote) {
      cur += c;
      if (c === "'") {
        if (s.charAt(i + 1) === "'") { cur += "'"; i += 2; continue; }
        inQuote = false;
      }
      i++;
      continue;
    }
    if (c === '"') { inStr = true; cur += c; i++; continue; }
    if (c === "'") { inQuote = true; cur += c; i++; continue; }
    if (c === '(') { depth++; cur += c; i++; continue; }
    if (c === ')') {
      depth--;
      if (depth === 0) { args.push(cur.trim()); i++; break; }
      cur += c; i++; continue;
    }
    if ((c === ',' || c === ';') && depth === 1) { args.push(cur.trim()); cur = ''; i++; continue; }
    cur += c;
    i++;
  }

  if (depth > 0) args.push(cur.trim()); // unbalanced formula: keep what we have
  return { args: args, end: i };
}

/**
 * Find every call to one of `names` in a formula, returning
 * {name, args, text, index} for each. Used for OFFSET/INDEX/INDIRECT.
 */
function findCalls(formula, names) {
  var upper = {};
  for (var n = 0; n < names.length; n++) upper[names[n].toUpperCase()] = true;

  var tokens = tokenize(formula);
  var found = [];
  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    if (t.type !== TOK.FUNC || !upper[t.value]) continue;
    var split = splitCallArgs(formula, t.start);
    if (split.end < 0) continue;
    var body = String(formula);
    if (body.charAt(0) === '=') body = body.substring(1);
    found.push({
      name: t.value,
      args: split.args,
      text: body.substring(t.start, split.end),
      index: t.start
    });
  }
  return found;
}

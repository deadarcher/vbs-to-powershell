/**
 * VBScript -> PowerShell converter core.
 *
 * Author: Brian Vitko (RFF - https://getrff.com)
 * Part of the free tools at https://getrff.com/tools/
 *
 * DESIGNED FROM MEASURED DATA, not guesswork: profile.py over 202 real sysadmin scripts (22,995
 * code lines) ranked what actually appears. The rules below are ordered by that ranking, and
 * anything rare enough to be risky is FLAGGED rather than guessed at.
 *
 * THE RULE THAT MATTERS MOST: never emit PowerShell we are not confident in. These scripts run as
 * SYSTEM across fleets, so a silently-wrong conversion is worse than an honest refusal. Everything
 * unconvertible comes back as the original line plus a reason. A competitor that hands you
 * confident, plausible, subtly-wrong PowerShell is the failure mode this exists to avoid.
 *
 * Every output line is classified so the conversion rate can be measured honestly:
 *   converted   - a rule matched and produced PowerShell
 *   blank       - whitespace
 *   comment     - a comment, mechanically re-marked
 *   flagged     - RECOGNISED but not safely convertible; emitted with a reason
 *   unknown     - no rule matched; emitted untouched with a reason
 */

// ── string protection ────────────────────────────────────────────────────────
// Every rule below is a regex over code. Without this, a rule would happily rewrite the inside of
// a string literal - turning  WScript.Echo "Set x = 1"  into nonsense. VBScript escapes a quote by
// doubling it, so the literal pattern has to allow "".
const STR = "STR";

function protectStrings(line) {
  const strings = [];
  const code = line.replace(/"(?:[^"]|"")*"/g, (m) => {
    strings.push(m);
    return `${STR}${strings.length - 1}`;
  });
  return { code, strings };
}

function restoreStrings(code, strings) {
  return code.replace(new RegExp(`${STR}(\\d+)\\u0002`, "g"), (_, i) => strings[Number(i)]);
}

/**
 * Restore for PowerShell OUTPUT. A VBScript literal is emitted as a double-quoted PowerShell
 * string, where `$` interpolates and a backtick escapes - so `"=COUNTIFS($B$2:$B$9)"` came out
 * reading three variables that do not exist. Both are escaped here and only here; the plain
 * restore above rebuilds VBScript text (colon splitting, canonicalisation) and must not.
 */
function restoreStringsPs(code, strings) {
  return code.replace(new RegExp(`${STR}(\\d+)\\u0002`, "g"),
    (_, i) => strings[Number(i)].replace(/`/g, "``").replace(/\$/g, "`$"));
}

// ── line joining ─────────────────────────────────────────────────────────────
/**
 * VBScript continues a statement with a trailing underscore. Half the files in the corpus use it
 * (762 occurrences), so this has to happen before any rule runs - otherwise every continued
 * statement is scored as two unconvertible fragments and the measurement is garbage.
 */
/** Toggle so the colon-splitting pass can be isolated when measuring; a change that moves the
 *  number needs to be attributable to ONE thing. */
const SPLIT_COLONS = globalThis.process?.env?.NO_COLON_SPLIT !== "1";

export function joinContinuations(src) {
  // NORMALISE LINE ENDINGS FIRST, and do it by collapsing every CR rather than with a \r?\n split.
  // Real .vbs files in the wild turn up with \r\r\n (a CRLF file converted to CRLF a second time)
  // and occasionally bare \r. A \r?\n split leaves the stray carriage return attached to the line,
  // and `$` in a regex does not match before \r - so every comment in those files failed its
  // comment test and fell through to the rule table as an unconvertible line. It cost ~5 points of
  // measured conversion rate, all of it on file-header comment blocks.
  // `\r+\n?` as ONE break, not `\r\n?`.
  //
  // Files with \r\r\n endings (a CRLF file converted to CRLF twice) are common in this corpus, and
  // `/\r\n?/` matches the first \r ALONE and then \r\n separately - turning one line break into two
  // and injecting a blank line between every real line. That silently broke every line
  // continuation, because the line after a trailing `_` was the phantom blank rather than the real
  // next line, so the statement was joined to nothing and emitted a dangling operator.
  //
  // An earlier pass "fixed" the same root cause by patching its comment-detection symptom. Symptom
  // gone, cause still there - which is why continuations kept failing for a completely different
  // -looking reason.
  const raw = src.replace(/\r+\n?|\n/g, "\n").split("\n");
  const out = [];
  let buf = null;
  let startLine = 0;

  for (let i = 0; i < raw.length; i++) {
    const line = raw[i];
    // `_` at end of line, NOT preceded by a word character. The old test required whitespace before
    // the underscore, but real scripts write `"text" &_` with no space and VBScript accepts it - so
    // every one of those statements stayed split, and the fragment `... + _` reached the output.
    // It was the single largest first-failure on the expanded corpus.
    // The word-boundary guard is what stops a trailing identifier like `my_var_` being mistaken
    // for a continuation.
    const isCont = /(?<!\w)_\s*$/.test(line) && !/^\s*'/.test(line);
    if (buf === null) {
      startLine = i + 1;
      buf = line;
    } else {
      buf += " " + line.trim();
    }
    if (isCont) {
      buf = buf.replace(/(?<!\w)_\s*$/, "");
    } else {
      out.push({ n: startLine, text: buf });
      buf = null;
    }
  }
  if (buf !== null) out.push({ n: startLine, text: buf });
  return SPLIT_COLONS ? out.flatMap(splitColonStatements) : out;
}

/**
 * VBScript separates statements on one line with a colon: `x = 1 : y = 2 : Set z = Nothing`.
 *
 * Handled by SPLITTING into separate logical lines here, so every rule downstream keeps seeing one
 * statement per line. Without it the whole thing reached the rule table as a single unrecognised
 * statement and emitted `$x = $MyInvocation.MyCommand.Path : $y = $x`, which PowerShell rejects -
 * it was the second-largest cause of parse failure on the 396-script corpus.
 *
 * Splitting is done on the STRING-PROTECTED form so a colon inside a literal is never a separator:
 * "C:\temp", "winmgmts:{...}" and "LDAP://x" all contain one. Parenthesis depth is tracked too, so
 * a colon inside a call argument does not split either.
 */
function splitColonStatements(entry) {
  const { code } = protectStrings(entry.text);
  if (!code.includes(":")) return [entry];

  // DO NOT split a line that opens a block or carries an inline Then. In VBScript
  // `If cond Then a : b` puts BOTH statements inside the If, so splitting silently changes the
  // control flow - and where the rule opens a brace, it also leaves the file unbalanced.
  // Splitting these cost 4 points of parse-clean before this guard existed.
  // ONLY an inline If is unsafe to split. `If cond Then a : b` puts BOTH statements inside the If,
  // so splitting changes the control flow.
  //
  // Everything else splits correctly and MUST be allowed to: `For i = 1 To 10 : x = i : Next` and
  // `Sub Foo() : body : End Sub` are loop and procedure bodies written on one line, and blocking
  // them left the trailing statements inside the For header, emitting
  // `for (...; ... : If $fs.FileExists(...) Then Exit For) {`. Guarding on the mere presence of
  // `Then` was the over-correction - it blocked `For ... : If ... Then Exit For`, which is fine to
  // split, because the If is a whole statement rather than a header with a body hanging off it.
  if (/^\s*(If|ElseIf)\b/i.test(code) && /\bThen\b/i.test(code)) return [entry];
  // `With x : .a = 1 : End With` on one line DOES split now that With blocks are expanded rather
  // than refused; the guard that kept it whole predates expandWith.

  // STRIP THE TRAILING COMMENT BEFORE LOOKING FOR SEPARATORS. Comments are full of colons -
  // "note:", "TODO:", "http://..." - and splitting inside one produces a fragment of English prose
  // that no rule matches. Protecting string literals is not enough; this cost 4.8 points of
  // parse-clean and showed up as a jump in unrecognised lines, not as an obvious break.
  const q = code.indexOf("'");
  const comment = q === -1 ? "" : code.slice(q);
  const bare = q === -1 ? code : code.slice(0, q);
  if (!bare.includes(":")) return [entry];

  const cuts = [];
  let depth = 0;
  for (let i = 0; i < bare.length; i++) {
    const ch = bare[i];
    if (ch === "(" || ch === "[") depth++;
    else if (ch === ")" || ch === "]") depth--;
    else if (ch === ":" && depth === 0) cuts.push(i);
  }
  if (!cuts.length) return [entry];

  // The protected form is a different LENGTH from the original (a literal becomes a short
  // placeholder), so the original cannot be sliced at these offsets. Split the protected string,
  // then restore each piece independently.
  const { strings } = protectStrings(entry.text);
  const parts = [];
  let prev = 0;
  for (const c of [...cuts, bare.length]) {
    const piece = restoreStrings(bare.slice(prev, c), strings).trim();
    if (piece) parts.push(piece);
    prev = c + 1;
  }
  const indent = (entry.text.match(/^\s*/) || [""])[0];
  // The trailing comment rides along with the LAST statement, which is where it sat in the source.
  if (comment) parts[parts.length - 1] += " " + restoreStrings(comment, strings);
  // ONE statement and a colon means the colon was TRAILING - `Sub Create(path, fs):` - which
  // VBScript tolerates. Left in place it reached the Sub rule and came out as `($path, $fs):) {`.
  if (parts.length < 2) return parts.length === 1 && parts[0] !== entry.text.trim()
    ? [{ n: entry.n, text: indent + parts[0] }] : [entry];

  return parts.map((p) => ({ n: entry.n, text: indent + p }));
}

// ── identifier canonicalisation ──────────────────────────────────────────────
/**
 * VBScript identifiers are CASE-INSENSITIVE. Every name table in this file is a case-sensitive Set,
 * and half the lookups lowercase both sides while the other half do not - so `stringContains = ...`
 * inside `Function StringContains` missed the return-by-name check and the function never returned,
 * and a Sub declared `Hello` but called as `hello` was registered as a variable by the catch-all and
 * emitted as `$null = $hello`, a call that silently does nothing. Both parse. Both are wrong.
 *
 * Fixed ONCE, here, before any other pass sees the text: every identifier is rewritten to the
 * spelling used where it was declared (or, for names never declared, its first use). Downstream
 * every table can stay case-sensitive because it only ever sees one spelling.
 *
 * String literals and comments are never touched - the ORIGINAL line is kept as `orig` so a TODO
 * still shows the reader exactly what they wrote. Members after a `.` are left alone: those belong
 * to COM objects, not this script.
 */
export function canonicaliseIdentifiers(lines) {
  const canon = new Map();                              // lower -> spelling to use everywhere
  // A DECLARED name may shadow an intrinsic - `Function Add(first, second)` makes `second` a
  // parameter, and the body's `Second` must follow it, not the date function. So a declaration
  // skips only keywords; a first USE also skips intrinsics and vb-constants, which are not
  // variables and keep whatever spelling the rules match on.
  const declare = (name, fromUse = false) => {
    if (!name) return;
    const l = name.toLowerCase();
    if (VB_KEYWORDS.has(l) || /^STR\d+$/.test(name)) return;
    if (fromUse && (VB_BUILTINS.has(l) || /^vb[A-Z]/.test(name))) return;
    if (!canon.has(l)) canon.set(l, name);
  };
  const declareList = (s) => s.split(",").forEach((p) => {
    const m = p.trim().match(/^([A-Za-z_]\w*)/);
    if (m) declare(m[1]);
  });
  // Code only: strings protected, trailing comment cut, Rem lines dropped.
  const codeOf = (text) => {
    if (/^\s*(?:'|Rem\b)/i.test(text)) return "";
    const { code } = protectStrings(text);
    const q = code.indexOf("'");
    return q === -1 ? code : code.slice(0, q);
  };
  const stripped = lines.map(({ text }) => codeOf(text));

  // Declarations FIRST, in file order, so a declared spelling always beats a use that precedes it.
  for (const code of stripped) {
    let m;
    if ((m = code.match(/^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?(?:Function|Sub)\s+([A-Za-z_]\w*)\s*\(?([^)]*)\)?/i))) {
      declare(m[1]);
      if (m[2]) declareList(m[2].replace(/\b(ByVal|ByRef)\b/gi, ""));
    }
    if ((m = code.match(/^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?Property\s+(?:Get|Let|Set)\s+([A-Za-z_]\w*)\s*\(?([^)]*)\)?/i))) {
      declare(m[1]);
      if (m[2]) declareList(m[2].replace(/\b(ByVal|ByRef)\b/gi, ""));
    }
    if ((m = code.match(/^\s*Class\s+([A-Za-z_]\w*)/i))) declare(m[1]);
    if ((m = code.match(/^\s*(?:Public|Private)?\s*(?:Dim|ReDim)\s+(?:Preserve\s+)?(.+)$/i)))
      declareList(m[1].replace(/\([^)]*\)/g, ""));
    if ((m = code.match(/^\s*(?:Public|Private)?\s*Const\s+(.+)$/i)))
      m[1].split(",").forEach((part) => declareList(part.split("=")[0]));
    if ((m = code.match(/^\s*(?:Public|Private)\s+([A-Za-z_]\w*(?:\s*,\s*[A-Za-z_]\w*)*)\s*$/i))) declareList(m[1]);
    if ((m = code.match(/^\s*For\s+Each\s+([A-Za-z_]\w*)/i))) declare(m[1]);
    if ((m = code.match(/^\s*For\s+([A-Za-z_]\w*)\s*=/i))) declare(m[1]);
  }
  // Then everything else at first use (scripts without Option Explicit declare nothing).
  // The lookbehind excludes members (`.Name`), sigils, and the control byte that opens a string
  // placeholder.
  const IDENT = /(?<![\w$.])([A-Za-z_]\w*)/g;
  const isHexLiteral = (s, off, name) => s[off - 1] === "&" && /^[HhOo][0-9A-Fa-f]+$/.test(name);
  for (const code of stripped)
    for (const m of code.matchAll(IDENT)) {
      if (isHexLiteral(code, m.index, m[1])) continue;        // `&H1F` is a number
      declare(m[1], true);
    }

  let changed = false;
  const out = lines.map((entry) => {
    const { text } = entry;
    if (/^\s*(?:'|Rem\b)/i.test(text) || !text.trim()) return entry;
    const { code, strings } = protectStrings(text);
    const q = code.indexOf("'");
    const head = q === -1 ? code : code.slice(0, q);
    const tail = q === -1 ? "" : code.slice(q);
    const fixed = head.replace(IDENT, (m, name, off, s) => {
      if (isHexLiteral(s, off, name)) return m;
      const c = canon.get(name.toLowerCase());
      return c && c !== name ? c : m;
    });
    if (fixed === head) return entry;
    changed = true;
    return { ...entry, text: restoreStrings(fixed + tail, strings), orig: text };
  });
  return changed ? out : lines;
}

// ── With blocks ──────────────────────────────────────────────────────────────
/**
 * `With obj ... .Member ... End With` has no PowerShell form, and it was REFUSED - every line inside
 * a With block came out as a TODO, in a construct that half the scripts of the 2000s use for
 * their RegExp and FileSystemObject setup. Expanded here, before any rule runs, into the VBScript
 * it stands for: a simple target (`With re`, `With objShell.Environment`) is written back in front
 * of each leading `.Member`; an expression target (`With CreateObject(...)`) is captured once into
 * an alias with `Set`, so it is evaluated once as VBScript evaluates it. Nesting is a stack.
 */
export function expandWith(lines) {
  const stack = [];
  let alias = 0, changed = false;
  const out = lines.map((entry) => {
    const { text } = entry;
    if (/^\s*(?:'|Rem\b)/i.test(text) || !text.trim()) return entry;
    const { code, strings } = protectStrings(text);
    const q = code.indexOf("'");
    const head = q === -1 ? code : code.slice(0, q);
    const tail = q === -1 ? "" : code.slice(q);
    const indent = (head.match(/^\s*/) || [""])[0];
    let m;
    if ((m = head.match(/^\s*With\s+(.+?)\s*$/i))) {
      changed = true;
      const target = restoreStrings(m[1], strings);
      if (/^[A-Za-z_][\w.]*$/.test(target)) {
        stack.push(target);
        return { ...entry, text: `${indent}' With ${target}`, orig: text };
      }
      const name = `__with${++alias}`;
      stack.push(name);
      return { ...entry, text: `${indent}Set ${name} = ${target}`, orig: text };
    }
    if (/^\s*End\s+With\s*$/i.test(head)) {
      changed = true;
      stack.pop();
      return { ...entry, text: `${indent}' End With`, orig: text };
    }
    if (!stack.length) return entry;
    const target = stack[stack.length - 1];
    // A leading-dot member: `.Name` not preceded by anything that would make it a normal member
    // access (a word, a closing bracket, a string) or a decimal point.
    const fixed = head.replace(/(?<![\w)\].])\.([A-Za-z_]\w*)/g, `${target}.$1`);
    if (fixed === head) return entry;
    changed = true;
    return { ...entry, text: restoreStrings(fixed + tail, strings), orig: entry.orig ?? text };
  });
  return changed ? out : lines;
}

// ── identifier discovery ─────────────────────────────────────────────────────
const VB_KEYWORDS = new Set(
  ("if then else elseif end select case for each next to step do while until loop wend function sub " +
   "dim redim preserve set const option explicit call exit on error resume goto class get let " +
   "public private with and or not xor mod is nothing true false null empty new byval byref as in " +
   "wscript err me erase lbound ubound").split(" ")
  // `property` is deliberately NOT here: it is only reserved inside a Class statement, and real
  // scripts use it as a variable (`For Each property In ...`). The Property rules match by regex.
);

/**
 * Collects the names that are VARIABLES, so `$` is added to those and only those. Function and Sub
 * names are collected separately and deliberately excluded - prefixing a call site would produce
 * PowerShell that parses and does the wrong thing, which is exactly what this tool must not do.
 */
export function discoverNames(lines) {
  const vars = new Set();
  const funcs = new Set();
  /** Loop counters. A class-level `Dim i, j` makes these class members in VBScript, but PowerShell
   *  cannot write `for ($this.i = 0; ...)` - so they are declared as members and referenced as
   *  plain locals, which is what the loop needs and costs nothing elsewhere. */
  const loopVars = new Set();
  const classes = new Set();
  /** Methods declared inside a Class. Their names live in the class, not the script. */
  const classFuncs = new Set();
  let classDepth = 0;

  const addList = (s) =>
    s.split(",").forEach((p) => {
      const m = p.trim().match(/^([A-Za-z_]\w*)/);
      if (m && !VB_KEYWORDS.has(m[1].toLowerCase())) vars.add(m[1]);
    });

  for (const { text } of lines) {
    const { code } = protectStrings(text);
    let m;

    if ((m = code.match(/^\s*(?:Public|Private)?\s*(?:Dim|ReDim)\s+(?:Preserve\s+)?(.+)$/i)))
      addList(m[1].replace(/\([^)]*\)/g, ""));
    if ((m = code.match(/^\s*(?:Public|Private)?\s*Const\s+(.+)$/i))) {
      // `Const ForReading = 1, ForWriting = 2` declares TWO names. Splitting on the first "=" and
      // taking the left side found only ForReading, so every later use of ForWriting stayed
      // unsigiled and the emitted PowerShell would not parse. Split on commas first, then on "=".
      m[1].split(",").forEach((part) => addList(part.split("=")[0]));
    }
    if ((m = code.match(/^\s*Set\s+([A-Za-z_]\w*)/i))) vars.add(m[1]);
    if ((m = code.match(/^\s*([A-Za-z_]\w*)\s*=\s*[^=]/))) {
      if (!VB_KEYWORDS.has(m[1].toLowerCase())) vars.add(m[1]);
    }
    if ((m = code.match(/^\s*For\s+Each\s+([A-Za-z_]\w*)/i))) { vars.add(m[1]); loopVars.add(m[1]); }
    if ((m = code.match(/^\s*For\s+([A-Za-z_]\w*)\s*=/i))) { vars.add(m[1]); loopVars.add(m[1]); }
    if (/^\s*Class\s+[A-Za-z_]\w*/i.test(code)) classDepth++;
    if (/^\s*End\s+Class\b/i.test(code)) classDepth = Math.max(0, classDepth - 1);
    if ((m = code.match(/^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?(?:Function|Sub)\s+([A-Za-z_]\w*)\s*\(?([^)]*)\)?/i))) {
      funcs.add(m[1]);
      if (classDepth) classFuncs.add(m[1]);
      if (m[2]) addList(m[2].replace(/\b(ByVal|ByRef)\b/gi, ""));
    }
    // `Private m_lOnBits(30)` and `Public a, b` - member ARRAYS and lists too. The old pattern
    // took one bare name only, so an array member's every use stayed unsigiled.
    if ((m = code.match(/^\s*(?:Public|Private)\s+(?!(?:Sub|Function|Property|Const|Default)\b)(.+)$/i)))
      addList(m[1].replace(/\([^)]*\)/g, ""));
    // A class name is a TYPE. The catch-all below registered it as a variable, so `New StreamBacker`
    // in an expression came out as `new $StreamBacker`.
    if ((m = code.match(/^\s*Class\s+([A-Za-z_]\w*)/i))) classes.add(m[1]);

    // Property Get/Let/Set parameters. The Property statement itself is refused (PowerShell has no
    // 1:1 equivalent), but its BODY is still converted - and without this the parameter name was
    // never registered, so every use inside the body emitted a bare identifier that will not parse.
    // `Public Property Let Enable(bYesNo)` was the single most common remaining failure.
    if ((m = code.match(/^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?Property\s+(?:Get|Let|Set)\s+[A-Za-z_]\w*\s*\(([^)]*)\)/i)))
      if (m[1]) addList(m[1].replace(/\b(ByVal|ByRef)\b/gi, ""));
  }

  /** Names from EXPLICIT declarations only, before the catch-all widens the set. */
  const declared = new Set(vars);
  /** Every bare name seen followed by `(`. */
  const called = new Set();

  // CATCH-ALL. Everything above finds names from explicit declarations, but real scripts are full
  // of identifiers that are never declared: ByRef out-parameters (`objReg.EnumKey(h, k, arrSubKeys)`),
  // globals defined in an included file, and anything in a script without Option Explicit.
  // Those reached the output as bare words, which PowerShell rejects - and between them they were
  // the largest remaining failure class.
  //
  // This is sound because VBScript has no bare-word literals: an identifier in an expression is a
  // variable, a procedure, or an intrinsic. Procedures and intrinsics are both known, so whatever
  // is left is a variable.
  //
  // Identifiers FOLLOWED BY `(` are skipped - those are calls or indexes, and a wrongly sigiled
  // call is a silent behaviour change rather than a visible error.
  for (const { text } of lines) {
    const { code } = protectStrings(text);
    // Strip the string PLACEHOLDERS before scanning. They look exactly like identifiers
    // (`STR0` preceded by a control byte), so the catch-all happily registered them as variables -
    // sigilize then rewrote every placeholder to `$STR0` and restoreStrings could no longer match
    // it, blanking every string literal in the file. Cost 44 points in one go.
    const bare = code.replace(/'[^\n]*$/, "").replace(/STR\d+/g, " 0 ");
    for (const m of bare.matchAll(/(?<![\w$.])([A-Za-z_]\w*)\s*(\()?/g)) {
      const name = m[1];
      if (m[2]) { called.add(name); continue; }          // followed by '(' - a call or an index
      const lower = name.toLowerCase();
      if (VB_KEYWORDS.has(lower) || VB_BUILTINS.has(lower)) continue;
      if (funcs.has(name)) continue;
      if (/^vb[A-Z]/.test(name)) continue;               // intrinsic constant, mapped separately
      vars.add(name);
    }
  }

  // A procedure name is not a variable - EXCEPT when a class METHOD shares its name with a
  // variable the script DECLARES. The stock VBScript MD5 class has methods F, G, H and I, and the
  // scripts that embed it use I as a loop counter at script level; VBScript keeps the two apart by
  // scope, and evicting `I` here left every `Item(I)` in two 1,600-line scripts unsigiled.
  // Inside the class, sigilize is handed the registry MINUS the method names (see convertVbs).
  funcs.forEach((f) => { if (!(classFuncs.has(f) && declared.has(f))) vars.delete(f); });
  // `Name(...)` where Name is nothing this file declares - not a variable, an array (those need a
  // Dim), a procedure, a class or an intrinsic - is a call into an INCLUDED file. Left in VBScript
  // syntax PowerShell reads `Name(a, b)` as one array argument; treated as a call it gets the
  // command form. If the include actually declares an array of that name the emitted call fails
  // loudly at run time, which is the honest outcome for something this file cannot know.
  const unknownCalls = new Set([...called].filter((n) =>
    !vars.has(n) && !funcs.has(n) && !classes.has(n) && !VB_BUILTINS.has(n.toLowerCase())
    && !VB_KEYWORDS.has(n.toLowerCase()) && !/^vb[A-Z]/.test(n)));
  // Class names are deliberately NOT removed from vars. Real scripts pair `Class XmlReader`
  // with a parameter called `xmlReader`, one name in VBScript, and removing it left the parameter
  // bare in every body that used it. The type positions are safe anyway: `New X` is rewritten to
  // `[X]::new()` before sigilize runs, and sigilize never touches a `[Type]` literal.
  return { vars, funcs, loopVars, classes, unknownCalls };
}

/** VBScript intrinsic functions and objects. Excluded from the catch-all above, because sigiling
 *  one of these would turn a working call into a variable reference that is always null. */
const VB_BUILTINS = new Set(
  ("ucase lcase trim ltrim rtrim mid left right len instr instrrev replace split join strreverse " +
   "strcomp space string filter cstr cint clng csng cdbl cbool cbyte cdate chr chrw asc ascw hex oct " +
   "abs int fix sgn sqr sin cos tan atn lenb ascb chrb midb leftb rightb instrb __invoke " +
   "round exp log rnd randomize timer now date time year month day hour minute " +
   "second weekday weekdayname monthname dateadd datediff datepart dateserial datevalue timeserial " +
   "timevalue formatnumber formatpercent formatcurrency formatdatetime isnull isempty isobject " +
   "isnumeric isarray isdate typename vartype ubound lbound array erase createobject getobject " +
   "getref eval execute executeglobal msgbox inputbox err wscript rgb loadpicture scriptengine " +
   "scriptenginebuildversion scriptenginemajorversion scriptengineminorversion cverr numlock " +
   "escape unescape setlocale getlocale").split(" ")
);

/**
 * Decides which KIND of `On Error Resume Next` a script is using, because the two need opposite
 * treatment and the difference is visible in the source.
 *
 *  - "checked"   : Err.Number / Err.Description is read somewhere. The control flow depends on it,
 *                  so an automatic conversion would change behaviour. Refuse and explain.
 *  - "suppress"  : Err is never read. The intent is just "keep going", and
 *                  $ErrorActionPreference = 'SilentlyContinue' is a fair mechanical equivalent for
 *                  the non-terminating errors that VBScript would have skipped past.
 *
 * `isLogonScript` sharpens the advice rather than the conversion: at logon nobody will accept a
 * blocking dialog or a slow script, which is usually the whole reason the line is there.
 */
/**
 * How the script is DEPLOYED changes the right answer, and often the source cannot tell you. A
 * script pushed by PDQ or an RMM is judged on its exit code; the same script run at logon
 * usually is not judged on anything.
 *
 * The same `On Error Resume Next` is benign in a logon script and dangerous in a PDQ deployment,
 * because there the exit code is the whole contract. So this GUESSES, the UI shows the guess, and
 * the user can correct it - a guess the user can see and override beats an assumption baked in.
 */
export const SCRIPT_CONTEXTS = {
  logon: {
    label: "Logon / startup script (GPO)",
    note: "Runs at sign-in, often as the user, with nobody watching. It must never block and must "
        + "not be slow - every second is added to every sign-in.",
  },
  rmm: {
    label: "Deployed by an RMM / PDQ / SCCM / Intune",
    note: "THE EXIT CODE IS THE CONTRACT. The deployment tool decides success or failure from it, "
        + "so anything that swallows an error without affecting the exit code turns a failed "
        + "deployment into a reported success.",
  },
  scheduled: {
    label: "Scheduled task / service",
    note: "Non-interactive, usually SYSTEM. No UI of any kind will be seen, and anything that waits "
        + "for input hangs until the task times out.",
  },
  interactive: {
    label: "Run by hand at a console",
    note: "A human is present, so prompts and dialogs are acceptable and the exit code usually is "
        + "not load-bearing.",
  },
  unknown: { label: "Not sure", note: "" },
};

export function detectContext(src) {
  const has = (rx) => rx.test(src);

  const logon = [
    /CreateObject\s*\(\s*["']WScript\.Network/i, /\.MapNetworkDrive\b/i,
    /\.AddWindowsPrinterConnection\b/i, /\.SetDefaultPrinter\b/i, /NETLOGON/i,
  ].filter(has).length;

  // Deployment tells: it reports a status by exiting with a code, or it installs something.
  const rmm = [
    /WScript\.Quit\s*\(?\s*[1-9]/i,          // a NON-ZERO exit code is the giveaway
    /msiexec/i, /\bsetup\.exe\b/i, /\/quiet\b|\/qn\b|\/silent\b/i,
    /\bUninstallString\b/i, /Win32_Product/i,
  ].filter(has).length;

  const interactive = [/\bMsgBox\b/i, /\bInputBox\b/i, /WScript\.StdIn/i].filter(has).length;

  if (logon >= 2) return "logon";
  if (rmm >= 2) return "rmm";
  if (interactive >= 1 && rmm === 0 && logon === 0) return "interactive";
  if (rmm === 1) return "rmm";
  return "unknown";
}

/**
 * Finds procedures whose parameters are MUTATED in the body.
 *
 * VBScript passes ByRef BY DEFAULT, so `Sub Bump(n)` followed by `n = n + 1` changes the caller's
 * variable. PowerShell passes by value, so the mutation is lost and the caller silently sees the
 * old value - no error, no warning, just a different answer.
 *
 * Doing this properly needs dataflow analysis (which call sites pass a variable rather than an
 * expression, and whether [ref] is safe there), which a line-based converter cannot do. So it is
 * DETECTED and flagged rather than guessed at - a REVIEW comment the reader can act on beats a
 * conversion that looks right and is not.
 */
function findByRefMutations(lines) {
  const out = new Map();           // proc name -> [mutated param names]
  let current = null, params = [], mutated = new Set();

  const flush = () => {
    if (current && mutated.size) out.set(current, [...mutated]);
    current = null; params = []; mutated = new Set();
  };

  for (const { text } of lines) {
    const { code } = protectStrings(text);
    let m;

    if ((m = code.match(/^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?(?:Function|Sub)\s+([A-Za-z_]\w*)\s*\(([^)]*)\)/i))) {
      flush();
      current = m[1];
      // ByVal is explicit opt-out, so those parameters are not at risk.
      params = splitArgs(m[2] || "")
        .map((p) => p.trim())
        .filter((p) => p && !/^ByVal\s/i.test(p))
        .map((p) => p.replace(/^ByRef\s+/i, "").trim());
      continue;
    }
    if (/^\s*End\s+(Function|Sub)\s*$/i.test(code)) { flush(); continue; }

    if (current && (m = code.match(/^\s*(?:Set\s+)?([A-Za-z_]\w*)\s*=\s*[^=]/))) {
      if (params.some((p) => p.toLowerCase() === m[1].toLowerCase())) mutated.add(m[1]);
    }
  }
  flush();
  return out;
}

/**
 * Procedures that RETURN a value, i.e. that assign to their own name.
 *
 * Needed at the point the method is OPENED, because a PowerShell class method with no type
 * annotation is implicitly [void] - and a [void] method containing `return $x` is a hard parse
 * error. Knowing this requires looking ahead into the body, which is why it is a pre-pass.
 */
function findReturningFuncs(lines) {
  const out = new Set();
  let current = null;

  for (const { text } of lines) {
    const { code } = protectStrings(text);
    let m;
    if ((m = code.match(/^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?(?:Function|Sub|Property\s+Get)\s+([A-Za-z_]\w*)/i))) {
      current = m[1];
      continue;
    }
    if (/^\s*End\s+(Function|Sub|Property)\s*$/i.test(code)) { current = null; continue; }
    // `Set Foo = obj` is the SAME return idiom, just for an object, and the optional `Set` was not
    // allowed for. A function returning a Dictionary or a COM handle therefore got no
    // `return $Foo` at all and silently returned nothing - and since PowerShell resolves a bare
    // name as a command, `GetActiveSessionIDs = $dic` then re-invoked the function by name.
    if (current && (m = code.match(/^\s*(?:Set\s+)?([A-Za-z_]\w*)\s*=\s*[^=]/i))) {
      if (m[1].toLowerCase() === current.toLowerCase()) out.add(current);
    }
    // `If x Then Foo = 1` - the return can also be inline, and so can `Case 0  Foo = "ok"`.
    if (current && (m = code.match(/\b(?:Then|Else)\s+(?:Set\s+)?([A-Za-z_]\w*)\s*=\s*[^=]/i))) {
      if (m[1].toLowerCase() === current.toLowerCase()) out.add(current);
    }
    if (current && (m = code.match(/^\s*Case\b.*?\s(?:Set\s+)?([A-Za-z_]\w*)\s*=\s*[^=]/i))) {
      if (m[1].toLowerCase() === current.toLowerCase()) out.add(current);
    }
  }
  return out;
}

function analyseErrorHandling(lines) {
  const src = lines.map((l) => l.text).join("\n");
  const usesOnError = /On\s+Error\s+Resume\s+Next/i.test(src);
  const readsErr = /\bErr\s*\.\s*(Number|Description|Source)\b/i.test(src);

  // Logon-script tells: drive/printer mapping, the network object, or running out of NETLOGON.
  const logonHits = [
    /WScript\.CreateObject\s*\(\s*["']WScript\.Network/i,
    /CreateObject\s*\(\s*["']WScript\.Network/i,
    /\.MapNetworkDrive\b/i,
    /\.AddWindowsPrinterConnection\b/i,
    /\.SetDefaultPrinter\b/i,
    /NETLOGON/i,
    /\bUserName\b.*\bComputerName\b/i,
  ].filter((rx) => rx.test(src)).length;

  return {
    usesOnError,
    mode: !usesOnError ? "none" : readsErr ? "checked" : "suppress",
    isLogonScript: logonHits >= 2,
  };
}

/**
 * Which `On Error Resume Next` regions BRANCH on Err, by source line.
 *
 * The file-level verdict in analyseErrorHandling cannot answer this: one Err.Number check in one
 * function made every OERN in the script "checked", so four regions that only wanted "do not
 * stop" were refused because a fifth was hard. Regions end at `On Error GoTo 0` or at a procedure
 * boundary, because VBScript scopes OERN per procedure.
 *
 * Returns the source line numbers of the checked OERN statements themselves, and of every line
 * inside a checked region, so the rule and the try/catch wrapper agree on where they are.
 */
function findCheckedOernRegions(lines) {
  const regions = [];
  let open = null;
  lines.forEach((l) => {
    const t = (l.text || "").trim();
    if (/^On\s+Error\s+Resume\s+Next\s*$/i.test(t)) {
      if (open) regions.push(open);
      open = { startLine: l.n, lines: [], reads: false };
      return;
    }
    if (/^On\s+Error\s+Goto\s+0\s*$/i.test(t)
     || /^(End\s+(Function|Sub)|Function\s|Sub\s)/i.test(t)) {
      if (open) { open.endLine = l.n; regions.push(open); open = null; }
      return;
    }
    if (!open) return;
    open.lines.push(l.n);
    if (/\bErr\s*\.\s*(Number|Description|Source)\b/i.test(t)) open.reads = true;
  });
  if (open) regions.push(open);

  const checked = regions.filter((r) => r.reads);
  return {
    starts: new Set(checked.map((r) => r.startLine)),
    lines: new Set(checked.flatMap((r) => r.lines)),
    endsChecked: new Set(checked.map((r) => r.endLine).filter((x) => x !== undefined)),
  };
}

export function sigilize(code, vars) {
  if (!vars.size) return code;
  // Longest-first so `objFile` is not partially matched by `objFil`.
  const names = [...vars].sort((a, b) => b.length - a.length).map((n) => n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  const rx = new RegExp(`(?<![\\w$.])(${names.join("|")})\\b`, "g");
  // NEVER sigilise a TYPE LITERAL. A script with a variable called `string` or `int` turned the
  // casts this converter emits into `[$string](...)`, which is "Missing type name after '['" and
  // cascades through the rest of the file. Detected by shape rather than by a name list, because
  // the casts we emit all look the same: `[Name]` immediately followed by a value - `[string](`,
  // `[int]$x`, `[math]::`. An array index like `$arr[i]` is followed by none of those, so real
  // indexing still gets its sigil.
  return indexBrackets(code.replace(rx, (m, name, off, s) =>
    (s[off - 1] === "[" && /^\]\s*(\(|\$|::|\[)/.test(s.slice(off + name.length)) ? m : "$" + name)));
}

/**
 * `$arr(0)` -> `$arr[0]`.
 *
 * VBScript indexes arrays with parentheses, PowerShell with brackets, and getting this wrong was the
 * single biggest source of parse errors in the measured run - it cascades, because one bad paren
 * desynchronises the rest of the expression and the file reports dozens of errors from one cause.
 *
 * Safe to do blindly here ONLY because sigilize has already run: a `$` prefix means the name was
 * found in the variable registry, and function names are deliberately excluded from that registry.
 * So `$name(` cannot be a call - it is an index. Scans for the balanced closing paren rather than
 * regex-matching, so nested calls inside the subscript survive.
 */
/**
 * `MyFunc(a, b)` -> `(MyFunc a b)`.
 *
 * PowerShell calls a function with space-separated arguments, and inside an expression the call has
 * to be parenthesised: `-not hasOption("x")` is a parse error, `-not (hasOption "x")` is not.
 * `MyFunc(a, b)` is also a real trap rather than merely unidiomatic - PowerShell reads it as ONE
 * argument containing an array, so it parses cleanly and silently passes the wrong thing.
 *
 * Only applied to names in the FUNCTION registry, so variables and COM method calls are untouched.
 */
function parenthesizeCalls(code, funcs, methods) {
  if (!funcs || !funcs.size) return code;
  const names = [...funcs].sort((a, b) => b.length - a.length)
    .map((n) => n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  // `:` excluded: `[math]::Cos(0)` is a static member, not a call to rewrite - the day Cos was
  // mapped it came out as `[math]::(Cos 0)`.
  const rx = new RegExp(`(?<![\\w$.:])(${names.join("|")})\\s*\\(`, "g");

  // Scans FORWARD from a moving offset instead of restarting at the beginning each time.
  //
  // Restarting made the rule re-match its own output: having emitted
  // `(GetInstallDir (VbsConcat a b) $sh)`, the `\s*\(` pattern then matched `GetInstallDir (` and
  // took the INNER call's parenthesis as its own argument list - stripping those parens and
  // leaving `$sh)` dangling. Nested calls are common in real scripts, so this corrupted them
  // silently wherever one call was passed to another.
  let from = 0;
  for (let guard = 0; guard < 200; guard++) {
    rx.lastIndex = from;
    const m = rx.exec(code);
    if (!m) break;
    const open = m.index + m[0].length - 1;
    let depth = 0, close = open;
    for (; close < code.length; close++) {
      if (code[close] === "(") depth++;
      else if (code[close] === ")") { depth--; if (depth === 0) break; }
    }
    // Unbalanced: skip this call site rather than abandoning the whole expression.
    if (close >= code.length) { from = m.index + m[0].length; continue; }

    // Each argument is converted in its own right. The scan resumes AFTER this call, so a call
    // nested inside the argument list - `Outer(a, Inner("x", y))` - was never visited and reached
    // the output in VBScript syntax, which PowerShell reads as one array argument.
    const args = splitArgs(code.slice(open + 1, close)).map((a) => a.trim()).filter(Boolean)
      .map((a) => parenthesizeCalls(a, funcs, methods));
    // An argument that is an EXPRESSION has to be parenthesised in command syntax: `(Fact $n - 1)`
    // passes three arguments, `$n`, `-` and `1`, and Fact(5) recursed with n=5 until the stack
    // went. Atoms - a variable, a number, a member path, a string placeholder, or something already
    // wrapped as one group - are left as they are.
    const isOneGroup = (t) => {
      if (t[0] !== "(") return false;
      let d = 0;
      for (let i = 0; i < t.length; i++) {
        if (t[i] === "(") d++;
        else if (t[i] === ")" && --d === 0) return i === t.length - 1;
      }
      return false;
    };
    const wrapped = args.map((a) =>
      /^-?\$?[\w.:]+$/.test(a) || /^STR\d+$/.test(a) || isOneGroup(a) ? a : `(${a})`);
    // A CLASS METHOD needs .NET call syntax, and this is the only place that still has the commas.
    // Emitting `(GetErrorMessage $res)` here and trying to add `$this.` later meant reconstructing
    // the argument list from a space-joined string, which cannot be done correctly - an argument
    // that is itself an expression has spaces in it.
    const call = methods && methods.has(m[1])
      ? `($this.${m[1]}(${args.join(", ")}))`
      : args.length ? `(${m[1]} ${wrapped.join(" ")})` : `(${m[1]})`;
    code = code.slice(0, m.index) + call + code.slice(close + 1);
    // Resume AFTER what was just written, so this call site cannot be reconsidered.
    from = m.index + call.length;
  }
  return code;
}

/**
 * `f(a,,,d)` -> `f(a,$null,$null,d)`.
 *
 * VBScript lets you omit arguments positionally - `OpenTextFile(path,,,-2)` is the idiomatic way to
 * say "default format, default create, unicode" - and PowerShell has no equivalent, so the bare
 * commas are a syntax error. 28 files failed on this one pattern.
 *
 * Run BEFORE string literals are restored, so a comma inside a literal is still a placeholder and
 * cannot be mistaken for an argument separator. Loops because `,,,` needs filling more than once.
 */
function fillEmptyArgs(code) {
  // `@(, X)` is the one-element array the Array() emitter writes on purpose; the unary comma is
  // not an omitted argument. Masked for the duration of the fill.
  code = code.replace(/@\(\s*,/g, "@(");
  let prev;
  do {
    prev = code;
    code = code.replace(/([(,])\s*,/g, "$1 $null,");
  } while (code !== prev);
  return code.replace(/,\s*\)/g, ", $null)").replace(//g, ",");
}

function indexBrackets(code) {
  // `Split(s, ".")(0)` - indexing the RESULT of a call. By the time this runs the call is already
  // `(VbsSplit $s ".")`, so the subscript appears as `)(0)`, which PowerShell reads as an attempt
  // to invoke the result rather than index it.
  code = code.replace(/\)\s*\((\d+|\$\w+)\)/g, ")[$1]");

  let out = "";
  for (let i = 0; i < code.length; i++) {
    // `$name(` OR `$this.name(` - a class member subscript reaches here with the $this prefix
    // already applied, and matching only the bare form left `$this.arr[$i](1)` unconverted.
    const m = /^(\$(?:this\.)?\w+)\(/.exec(code.slice(i));
    if (!m) { out += code[i]; continue; }

    // `$name()` with EMPTY parens is a call, not an index - converting it gave `$DSN[]`, which is
    // MissingArrayIndexExpression. Only a non-empty subscript can be an index.
    if (/^\$(?:this\.)?\w+\(\s*\)/.test(code.slice(i))) { out += code[i]; continue; }

    let depth = 0, j = i + m[0].length - 1;
    for (; j < code.length; j++) {
      if (code[j] === "(") depth++;
      else if (code[j] === ")") { depth--; if (depth === 0) break; }
    }
    if (j >= code.length) { out += code[i]; continue; }     // unbalanced: leave it alone

    out += `${m[1]}[` + indexBrackets(code.slice(i + m[0].length, j)) + "]";
    i = j;

    // CHAINED subscripts: VBScript writes a jagged array as `arr(i)(j)`, so a second `(...)`
    // immediately after the one just converted is another index, not a call on the result.
    // Also reached after a `]` that a previous pass produced, e.g. `$this.arr[$i](1)`.
    while (code[i + 1] === "(") {
      let d2 = 0, k = i + 1;
      for (; k < code.length; k++) {
        if (code[k] === "(") d2++;
        else if (code[k] === ")") { d2--; if (d2 === 0) break; }
      }
      if (k >= code.length) break;
      out += "[" + indexBrackets(code.slice(i + 2, k)) + "]";
      i = k;
    }
  }
  // A subscript on a subscript, `arr(i)(1)(j)`, AFTER the scan above has made the first one a
  // bracket. The chain must START at a `$variable`: a bare `](` is also the CAST `[string]($s)`,
  // and a first version turned every cast in the corpus into `[string][$s]` - 122 files in one
  // run. A second version ran before the scan, so the chain never started with a bracket.
  const chained = /(\$\w+(?:\[[^[\]]*\])+)\s*\((\d+|\$\w+)\)/g;
  for (let guard = 0; guard < 4 && chained.test(out); guard++) {
    chained.lastIndex = 0;
    out = out.replace(chained, "$1[$2]");
  }
  return out;
}

// ── expression translation ───────────────────────────────────────────────────
function translateOperators(code, { isCondition }) {
  // VBScript accepts `=>` and `=<` as spellings of `>=` and `<=`; emitted as written they became
  // `-ceq -gt`, a missing operand.
  let c = code.replace(/=>/g, ">=").replace(/=</g, "<=");

  // VBScript's `Not` binds LOOSER than comparison: `Not LCase(x) = "y"` means `Not (LCase(x) = "y")`.
  // PowerShell's -not binds TIGHTER, so a literal translation reads as `(-not LCase(x)) -eq "y"` -
  // a different question, and usually a parse error too. Parenthesise before anything else runs.
  if (isCondition)
    c = c.replace(/\bNot\s+(.+?)\s*(<>|>=|<=|=|>|<)\s*(.+)$/i, (_, a, op, b) => `Not (${a} ${op} ${b})`);

  // `^` is exponentiation in VBScript and NOT an operator in PowerShell at all, so it is a hard
  // parse error rather than a wrong answer.
  c = c.replace(/([\w$.()[\]]+)\s*\^\s*([\w$.()[\]]+)/g, "[math]::Pow($1, $2)");

  // `\` is integer division, truncating toward zero. PowerShell has no such operator, and `/` is
  // always floating point - so `7 \ 2` must be 2, not 3.5.
  // foldBinary, not a character-class regex: the old pattern could not span a space, so
  // `UBound(a) \ n` - whose left operand is `(VbsUBound $a)` by now - came out as
  // `(VbsIntDiv ((VbsUBound $a) $n))`, one argument.
  c = foldBinary(c, "\\", "VbsIntDiv");
  // Spaces around comparison operators are OPTIONAL in VBScript - `If i=0 Then` is as common as
  // `If i = 0 Then` - and requiring them left a bare `=` in the emitted condition, which PowerShell
  // reads as an assignment and rejects. This was the single largest failure class on the
  // 396-script corpus at 23 files.
  //
  // ORDER MATTERS: the two-character operators must go first, or the `=` rule bites the `=` in
  // `>=` and turns it into `> -eq`.
  // -cne / -ceq, not -ne / -eq.
  //
  // VBScript compares strings CASE-SENSITIVELY by default; PowerShell's -eq is case-INSENSITIVE.
  // `If strUser = "Administrator"` therefore flips from False to True on "administrator" - a guard
  // that silently stops guarding. The case-sensitive operators are correct for strings and behave
  // identically to the insensitive ones on numbers, so they are safe everywhere.
  c = c.replace(/\s*<>\s*/g, " -cne ");
  // The ordered comparisons are unambiguous in EVERY context - VBScript has no shifts or generics,
  // so `<` and `>` only ever compare - and `x = (Year(d) >= 2026)` left them as written. Only the
  // `=` needs the condition flag, because outside a condition it may be the assignment itself.
  c = c.replace(/\s*>=\s*/g, " -ge ").replace(/\s*<=\s*/g, " -le ");
  // `=` too, in every context. This function only ever sees the RIGHT side of an assignment (the
  // rules split on the first `=` before calling it), so any `=` left is a comparison:
  // `d.Add(key, UCase(x) = "INCLUDE")` passes a Boolean. The condition flag no longer gates it.
  c = c.replace(/(?<![-<>!=+*/])=(?!=)/g, " -ceq ");
  c = c.replace(/(?<![-\w])\s*>\s*/g, " -gt ").replace(/(?<![-\w])\s*<\s*/g, " -lt ");
  c = c.replace(/(?<=\w)\s*>\s*/g, " -gt ").replace(/(?<=\w)\s*<\s*/g, " -lt ");
  // VBScript's And/Or/Xor/Not are BITWISE, always. They read as logical only because True is -1,
  // i.e. every bit set, so `-1 And -1` is `-1` is True. PowerShell's `-and` is logical ONLY.
  //
  // Emitting `-and` was silently wrong in the way that matters most: `mybyte -and 0x80` is $true
  // for EVERY nonzero byte, so a UTF-8 decoder built on it takes the multibyte path on plain ASCII
  // and throws on the first character. It parses perfectly, which is why parse-clean never saw it
  // and why this suite was green while the output was dead. Found by a reviewer running a real
  // script, not by any check here.
  //
  // A blanket swap to `-band` is NOT the fix, for two separate reasons:
  //   1. PRECEDENCE. PowerShell puts -band/-bor at COMPARISON precedence, so
  //      `$a -eq 1 -band $b -eq 2` parses as `(($a -eq 1) -band $b) -eq 2`. Folding into a call
  //      parenthesises both operands, which is the only reliable way to keep the grouping.
  //   2. Not. VBScript True is -1 so `Not True` is 0 (False). PowerShell $true is 1, so
  //      `-bnot $true` is -2, which is TRUTHY. Exactly backwards. The helper normalises a bool to
  //      -1/0 first, which makes the bitwise form correct for booleans AND for real bit masks.
  //
  // Folded LOOSEST-FIRST (Xor, then Or, then And) so the resulting nesting matches VBScript's
  // precedence: Not binds tightest, then And, then Or, then Xor.
  // Drop the case-sensitive form where one side is a NUMERIC LITERAL. -ceq and -eq are identical
  // on numbers, so `$depth -ceq 0` is correct but reads as though someone did not know what -ceq
  // means, and that costs credibility on every line of a long script. The case-sensitive operator
  // is kept everywhere else, because that is where it is actually doing work.
  c = c.replace(/-c(eq|ne)\s+(-?\d+(?:\.\d+)?|0x[0-9A-Fa-f]+)\b/g, "-$1 $2");
  c = c.replace(/\b(-?\d+(?:\.\d+)?|0x[0-9A-Fa-f]+)\s+-c(eq|ne)\b/g, "$1 -$2");

  c = foldWordOp(c, "Xor", "VbsXor");
  c = foldWordOp(c, "Or", "VbsOr");
  c = foldWordOp(c, "And", "VbsAnd");
  c = foldUnaryNot(c);
  c = c.replace(/\bMod\b/gi, "%");
  // The (?<!\$) guards are load-bearing. These run in sequence over the SAME string, and the
  // matches are case-insensitive - so `Is Nothing` became `-eq $null`, and then the later
  // `\bNull\b/i` rule matched the `null` it had just produced and emitted `$$null`. A replacement
  // re-matching its own output is silent: it parses as a variable named `$null` dereferenced twice.
  c = c.replace(/\bIs\s+Nothing\b/gi, "-eq $null").replace(/(?<!\$)\bNothing\b/gi, "$null");
  // `a Is b` on two objects is REFERENCE equality, which -eq is not for most .NET objects.
  // No parentheses in the operand classes: `(a Is b)` captured `(a` and `b)` and emitted
  // `ReferenceEquals((a, b))`, one tuple argument.
  c = c.replace(/([\w$.[\]]+)\s+Is\s+([\w$.[\]]+)/g, "[object]::ReferenceEquals($1, $2)");
  // $null goes on the LEFT. With an ARRAY on the left, PowerShell's -eq FILTERS instead of
  // comparing: `$items -eq $null` returns the null elements, so an empty collection is falsy and a
  // populated one is truthy - which happens to give the right answer often enough to hide the bug,
  // and gives the wrong one whenever the array legitimately contains a null. `$null -eq $items`
  // always compares. It reads backwards and it is the documented form for exactly this reason.
  c = c.replace(/(\S+)\s+-c?eq\s+\$null\b/g, "$null -eq $1");
  c = c.replace(/(\S+)\s+-c?ne\s+\$null\b/g, "$null -ne $1");
  c = c.replace(/(?<!\$)\bTrue\b/gi, "$true").replace(/(?<!\$)\bFalse\b/gi, "$false");
  c = c.replace(/(?<!\$)\bEmpty\b/gi, "$null").replace(/(?<!\$)\bNull\b/gi, "$null");
  // VBScript concatenation. Spaces around `&` are OPTIONAL - `strKey&"Value"` is extremely common -
  // and requiring them left the ampersand in the output, where PowerShell reads it as the
  // call operator and rejects the line.
  //
  // `&H1F` is a HEX LITERAL, not concatenation, so it has to be converted rather than split;
  // getting that wrong turns a registry hive constant into a syntax error.
  // Case-INSENSITIVE: `&h3f` is as common as `&H3F`, and matching only the uppercase form left the
  // ampersand to be read as concatenation - `mybyte And &h3f` came out as `$mybyte -and + $h3f`,
  // which is both a syntax error and, had it parsed, the wrong operation.
  c = c.replace(/&H([0-9A-Fa-f]+)&?/gi, "0x$1");

  // `&` becomes VbsConcat, NOT `+`.
  //
  // `+` was the original translation and it is SILENTLY wrong: VBScript's `&` always concatenates
  // as string, so `1 & 2` is "12" where PowerShell's `+` on two numbers gives 3. Parse-clean could
  // never see this - and "fixing" the ampersand to `+` actually RAISED that score while
  // introducing the bug. The conformance suite is what catches it.
  // PAIRWISE, deliberately. Nesting five VbsConcat calls for a five-part string is correct and
  // ugly, and flattening it into one variadic call was tried on 2026-09-21 and REVERTED: foldChain
  // interacts badly with the multi-argument builtins, turning
  // `cbyte("&H" & mid(str, i + 1, 2))` into `(VbsConcat "&H" (mid($str, $null)) $i + 1, 2)` - it
  // took Mid's argument list apart. The flattening is worth redoing, but only after Mid/InStr/Left
  // are parsed into a real call node instead of being pattern-matched, because that is the actual
  // conflict. Readable output is not worth a silently mangled argument list.
  c = foldBinary(c, "&", "VbsConcat");

  // `+` is overloaded in VBScript: numeric when both operands look numeric, string otherwise.
  // PowerShell decides from the LEFT operand only, so `"1" + 1` is 2 in VBScript and "11" here.
  //
  // Only folded when a STRING LITERAL is involved. Folding every `+` would be more faithful and
  // would turn `i = i + 1` into `$i = VbsAdd $i 1` throughout - unreadable output for a case that
  // does not arise, since ordinary arithmetic behaves identically. A literal on one side is the
  // signal that the overload actually matters.
  // Folded whenever an operand is NOT a numeric literal. The first version folded only next to a
  // string LITERAL, and `Split(addr, ":")(1) + 1` - a string from Split, a number from the
  // script - came out as `"17472" + 1`, which PowerShell makes "174721": a client API port that
  // does not exist, returned without complaint. Two numeric literals are left alone; anything
  // else goes through VbsAdd, which decides the way VBScript does.
  {
    const numeric = /^-?\d+(?:\.\d+)?$/;
    let needs = false;
    for (const m of c.matchAll(/\+/g)) {
      const left = (c.slice(0, m.index).match(/([\w$.()[\]]+)\s*$/) || [, ""])[1];
      const right = (c.slice(m.index + 1).match(/^\s*([\w$.()[\]]+)/) || [, ""])[1];
      if (!left) continue;                                        // a sign, not an operator
      if (!(numeric.test(left) && numeric.test(right))) { needs = true; break; }
    }
    if (needs) c = foldBinary(c, "+", "VbsAdd");
  }

  return c.replace(/\s{2,}/g, " ");
}

/**
 * Rewrites `a OP b OP c` into nested calls: `Fn (Fn a b) c`.
 *
 * Left-associative, and it walks OUTWARD from each operator to take whole operands - a regex
 * cannot, because an operand may itself be a call with its own parentheses and commas.
 * Deliberately stops at a comma or an unmatched bracket so an argument list is never swallowed.
 */
function foldBinary(code, op, fn) {
  const isBoundary = (ch) => ch === undefined || ",;".includes(ch);

  for (let guard = 0; guard < 80; guard++) {
    const i = code.indexOf(op);
    if (i === -1) break;

    // left operand: scan back over balanced brackets to the start of the term
    let l = i - 1, depth = 0;
    while (l >= 0) {
      const ch = code[l];
      if (ch === ")" || ch === "]") depth++;
      else if (ch === "(" || ch === "[") { if (depth === 0) break; depth--; }
      else if (depth === 0 && isBoundary(ch)) break;
      l--;
    }
    // right operand: scan forward the same way, stopping before the next operator of any kind
    let r = i + op.length, d2 = 0;
    while (r < code.length) {
      const ch = code[r];
      if (ch === "(" || ch === "[") d2++;
      else if (ch === ")" || ch === "]") { if (d2 === 0) break; d2--; }
      else if (d2 === 0 && (isBoundary(ch) || (ch === "&" && r > i + op.length))) break;
      r++;
    }

    const left = code.slice(l + 1, i).trim();
    const right = code.slice(i + op.length, r).trim();
    if (!left || !right) break;                    // malformed; leave the line alone

    code = code.slice(0, l + 1) + `(${fn} ${wrapArg(left)} ${wrapArg(right)})` + code.slice(r);
  }
  return code;
}

/** Arguments to a PowerShell command need parentheses unless they are a single atom. */
/**
 * Split on TOP-LEVEL occurrences of a VBScript word operator and fold them into helper calls,
 * left-associatively. Called loosest-operator-first so the nesting comes out matching VBScript's
 * precedence (Not > And > Or > Xor).
 *
 * Top-level means "not inside brackets". String literals are already swapped for STRn placeholders
 * by this point, so a word like `Or` sitting inside a message can't be matched here by accident.
 * The word-boundary match is what keeps `Ordinal`, `Android` and `strOrder` out of it.
 */
/**
 * Split a chain of one SYMBOL operator into a single variadic call: `a & b & c` becomes
 * `(VbsConcat a b c)` rather than a three-deep nest.
 *
 * Recurses into bracket groups first, same as foldWordOp, so a chain inside a call argument is
 * flattened at its own level. String literals are placeholders by this point, so an `&` inside a
 * message cannot be matched, and `&H1F` was rewritten to `0x1F` upstream.
 */
function foldChain(code, op, fn) {
  let rebuilt = "", i = 0;
  while (i < code.length) {
    const ch = code[i];
    if (ch === "(" || ch === "[") {
      const close = ch === "(" ? ")" : "]";
      let d = 1, j = i + 1;
      while (j < code.length && d > 0) {
        if (code[j] === ch) d++;
        else if (code[j] === close) d--;
        j++;
      }
      if (d !== 0) { rebuilt += code.slice(i); i = code.length; break; }
      rebuilt += ch + foldChain(code.slice(i + 1, j - 1), op, fn) + close;
      i = j;
    } else { rebuilt += ch; i++; }
  }
  code = rebuilt;

  // A chain STOPS at an argument separator. Without this, `Foo(a & b, c)` chained straight across
  // the comma and emitted `Foo((VbsConcat $a ($b, $null)) $c)` - it ate the next argument. Split on
  // top-level commas, flatten each piece independently, put the commas back.
  {
    const segs = [];
    let d = 0, start = 0;
    for (let k = 0; k < code.length; k++) {
      const ch = code[k];
      if (ch === "(" || ch === "[") d++;
      else if (ch === ")" || ch === "]") d--;
      else if (d === 0 && (ch === "," || ch === ";")) { segs.push(code.slice(start, k), ch); start = k + 1; }
    }
    if (segs.length) {
      segs.push(code.slice(start));
      return segs.map((s) => (s === "," || s === ";" ? s : foldChain(s, op, fn))).join("");
    }
  }

  const cuts = [];
  let depth = 0;
  for (let k = 0; k < code.length; k++) {
    const ch = code[k];
    if (ch === "(" || ch === "[") depth++;
    else if (ch === ")" || ch === "]") depth--;
    else if (depth === 0 && code.startsWith(op, k)) cuts.push(k);
  }
  if (!cuts.length) return code;

  // Everything before the first operator and after the last belongs to the surrounding statement
  // (`$x = ` on the left, a trailing `)` or `,` on the right), so only the span between the first
  // and last operand is replaced.
  const parts = [];
  let prev = 0;
  for (const k of cuts) { parts.push(code.slice(prev, k)); prev = k + op.length; }
  parts.push(code.slice(prev));
  if (parts.some((p) => !p.trim())) return code;         // malformed; leave the line alone

  // Reattach whatever sat outside the chain. The first part may carry a `$x = ` prefix and the
  // last may carry a trailing separator; splitting those off keeps the call itself clean.
  const head = parts[0].match(/^(.*?)(\S+\s*)$/s) || [null, "", parts[0]];
  const tail = parts[parts.length - 1].match(/^(\s*\S+)(.*)$/s) || [null, parts[parts.length - 1], ""];
  const operands = [head[2], ...parts.slice(1, -1), tail[1]].map((p) => wrapArg(p));
  return `${head[1]}(${fn} ${operands.join(" ")})${tail[2]}`;
}

function foldWordOp(code, word, fn) {
  // Recurse into bracket groups FIRST. A depth-0-only scan misses the single most common shape
  // there is - `If (myByte And &H80) Then` - because the whole expression sits inside the
  // condition's own parentheses, so the operator is never at depth 0 and was left untouched.
  let rebuilt = "", i = 0;
  while (i < code.length) {
    const ch = code[i];
    if (ch === "(" || ch === "[") {
      const close = ch === "(" ? ")" : "]";
      let depth = 1, j = i + 1;
      while (j < code.length && depth > 0) {
        if (code[j] === ch) depth++;
        else if (code[j] === close) depth--;
        j++;
      }
      if (depth !== 0) { rebuilt += code.slice(i); i = code.length; break; }  // unbalanced, leave it
      rebuilt += ch + foldWordOp(code.slice(i + 1, j - 1), word, fn) + close;
      i = j;
    } else { rebuilt += ch; i++; }
  }
  code = rebuilt;

  const rx = new RegExp(`\\b${word}\\b`, "gi");
  const cuts = [];
  let depth = 0;
  for (let k = 0; k < code.length; k++) {
    const ch = code[k];
    if (ch === "(" || ch === "[") depth++;
    else if (ch === ")" || ch === "]") depth--;
    else if (depth === 0) {
      rx.lastIndex = k;
      const m = rx.exec(code);
      if (m && m.index === k) { cuts.push([k, k + m[0].length]); k = k + m[0].length - 1; }
    }
  }
  if (!cuts.length) return code;

  const parts = [];
  let prev = 0;
  for (const [s, e] of cuts) { parts.push(code.slice(prev, s)); prev = e; }
  parts.push(code.slice(prev));

  // A bare operator with nothing on one side is not an expression we understand; leaving the line
  // untouched is better than emitting a call with a missing argument.
  if (parts.some((p) => !p.trim())) return code;

  return parts.reduce((acc, p) => `(${fn} ${wrapArg(acc)} ${wrapArg(p)})`);
}

/**
 * `Not` is unary and binds tightest, so it folds after the binary operators have already been
 * turned into calls. Scans right-to-left so `Not Not x` nests correctly.
 */
function foldUnaryNot(code) {
  // Same recursion as foldWordOp, and for the same reason: `If Not (a = 2) Then` puts the operand
  // in brackets, and `If (Not x) Then` puts the whole thing in brackets.
  let rebuilt = "", i = 0;
  while (i < code.length) {
    const ch = code[i];
    if (ch === "(" || ch === "[") {
      const close = ch === "(" ? ")" : "]";
      let d = 1, j = i + 1;
      while (j < code.length && d > 0) {
        if (code[j] === ch) d++;
        else if (code[j] === close) d--;
        j++;
      }
      if (d !== 0) { rebuilt += code.slice(i); i = code.length; break; }
      rebuilt += ch + foldUnaryNot(code.slice(i + 1, j - 1)) + close;
      i = j;
    } else { rebuilt += ch; i++; }
  }
  code = rebuilt;

  const rx = /\bNot\b/gi;
  let out = code, guard = 0;
  while (guard++ < 40) {
    // Rightmost top-level `Not`, so the innermost one is folded first.
    let found = -1, depth = 0;
    for (let i = 0; i < out.length; i++) {
      const ch = out[i];
      if (ch === "(" || ch === "[") depth++;
      else if (ch === ")" || ch === "]") depth--;
      else if (depth === 0) {
        rx.lastIndex = i;
        const m = rx.exec(out);
        if (m && m.index === i) { found = i; i = i + m[0].length - 1; }
      }
    }
    if (found === -1) break;

    // Operand runs to the end of the term: balanced brackets, stopping at a separator.
    let r = found + 3, d2 = 0;
    while (r < out.length) {
      const ch = out[r];
      if (ch === "(" || ch === "[") d2++;
      else if (ch === ")" || ch === "]") { if (d2 === 0) break; d2--; }
      else if (d2 === 0 && ",;".includes(ch)) break;
      r++;
    }
    const operand = out.slice(found + 3, r).trim();
    if (!operand) break;
    out = out.slice(0, found) + `(VbsNot ${wrapArg(operand)})` + out.slice(r);
  }
  return out;
}

/**
 * Rewrite `Name(<anything, brackets balanced>)` by scanning for the matching close paren.
 *
 * The lookbehind is load-bearing in two directions. It stops `.Trim(` in this function's OWN
 * output from being matched again, which would loop forever; and it stops a user-defined
 * `$obj.Len(...)` or a variable named `strTrim` from being mistaken for the builtin.
 */
function replaceCall1(code, name, build) {
  // The `:` in the lookbehind is load-bearing: `Abs` emits `[math]::Abs(`, which would otherwise
  // match itself on the next pass and spin until the guard trips, leaving nested garbage behind.
  if (USER_FUNCS.has(name.toLowerCase())) return code;      // the script's own procedure shadows it
  const rx = new RegExp(`(?<![\\w.$:])${name}\\s*\\(`, "gi");
  let out = code;
  for (let guard = 0; guard < 200; guard++) {
    rx.lastIndex = 0;
    const m = rx.exec(out);
    if (!m) break;

    const open = m.index + m[0].length - 1;
    let depth = 1, j = open + 1;
    while (j < out.length && depth > 0) {
      if (out[j] === "(") depth++;
      else if (out[j] === ")") depth--;
      j++;
    }
    if (depth !== 0) break;                 // unbalanced; leaving the line alone beats mangling it

    out = out.slice(0, m.index) + build(out.slice(open + 1, j - 1)) + out.slice(j);
  }
  return out;
}

/**
 * True when an expression is a single token that can be passed as a bare command argument.
 *
 * Everything else has to be parenthesised, because PowerShell's argument parsing is not expression
 * parsing: `Write-Output [math]::Round(2.6)` prints "[math]::Round" and then 2.6 on separate lines,
 * and `Write-Output "a" + $b` passes three arguments rather than concatenating. Both are silently
 * wrong output, not errors.
 */
/**
 * Discard the value of a call used as a STATEMENT: `$null = <call>`.
 *
 * THE most consequential difference between the two languages, and it is completely silent. In
 * VBScript a procedure returns only what you assign to its name, and WScript.Echo is a side
 * channel. In PowerShell every uncaptured value inside a function joins that function's return
 * value, so a Sub that logs three times and ends with a comparison returns an ARRAY of four
 * things - and an array of four things is TRUTHY.
 *
 * Concretely, in the script this came from: CheckRunasAlreadyLaunchedForSession returned
 * @("Found potential...", $true, ..., $false) and every caller saw "already launched", so
 * LaunchOncePerSession would never launch anything. Nothing errors. Nothing parses wrong. The
 * feature just quietly stops working. A COM method call leaks its return the same way -
 * .CreateFolder() hands back a folder object that lands in the output.
 *
 * Left alone: anything that is already an assignment, a keyword statement, or one of the cmdlets
 * we emit deliberately for their output.
 */
function suppressOutput(call) {
  const t = call.trim();
  if (!t || t.startsWith("#")) return call;
  // Quotes included: an indexed assignment with a string key - `$objEnv["SEE_MASK"] = 1` - is
  // still an assignment, and prefixing it produced `$null = $objEnv["SEE_MASK"] = 1`.
  // `$` included too: a subscript that is itself a variable - `$lines[$lineIdx] = ...` - is still
  // an assignment, and without it the line came out as `$null = $lines[$lineIdx] = ...`.
  if (/^\$?[\w:.\[\]"'$]+\s*=[^=]/.test(t)) return call;             // already captured
  if (/^(return|exit|throw|break|continue|if|foreach|for|while|do|switch|function|try|catch|finally|param)\b/i.test(t)) return call;
  if (/^(Write-Output|Write-Host|Write-Error|Write-Warning|Write-Verbose|\[Console\]::)/.test(t)) return call;
  if (/^\$null\s*=/.test(t)) return call;                            // idempotent: already done
  // ANY brace anywhere on the line. Two narrower versions of this both failed: checking only the
  // first character missed `1 { ... }`, and checking for a TRAILING brace missed the single-line
  // switch clause `0 { $x = "Success" }`, which ends in `}`. A call in statement position never
  // contains a brace in this converter's output, so the broad test costs nothing and stops the
  // guessing.
  if (/[{}]/.test(t)) return call;
  // A leading TYPE LITERAL means a class property or method declaration - `[object] $x = $null` -
  // and `$null = [object] $x = $null` is "Missing ';' or end-of-line in property definition".
  // Two class-heavy scripts went from clean to 85 errors each on exactly this.
  if (/^\[[A-Za-z_][\w.]*\]/.test(t)) return call;
  // The converter's own helpers. They either return nothing or their value is already handled,
  // and `$null = VbsClearErr` on every On Error line is pure noise.
  if (/^Vbs[A-Z]/.test(t)) return call;
  // Cmdlets we emit deliberately. They either return nothing or their output is the point, and
  // `$null = Start-Sleep -Milliseconds 500` is noise on every line that has one.
  if (/^[A-Z][a-z]+-[A-Z]\w+/.test(t)) return call;
  // A call is the only other thing this path emits: `$obj.Method(...)`, `Name(...)`, `Name arg`.
  return call.replace(/^(\s*)/, "$1$null = ");
}

function isAtom(e) {
  const t = e.trim();
  return /^\$?[\w.]+$/.test(t) || /^"(?:[^"`]|`.)*"$/.test(t);
}

function wrapArg(s) {
  const t = s.trim();
  if (/^\$?[\w.]+$/.test(t)) return t;                       // $var, literal, member path
  if (/^STR\d+$/.test(t)) return t;              // a protected string literal
  if (/^\(.*\)$/.test(t)) return t;                          // already parenthesised
  return `(${t})`;
}

const FUNC_MAP = [
  [/\bUCase\s*\(/gi, "(($1).ToUpper()"], [/\bLCase\s*\(/gi, "(($1).ToLower()"],
];

/**
 * Multi-argument VBScript built-ins -> .NET/PowerShell equivalents.
 *
 * Done with balanced-paren scanning rather than regex because the arguments are themselves
 * expressions that contain parentheses and commas: `Replace(Mid(s, 1, 4), "a", "b")` is common and
 * a regex splits it in the wrong place, producing output that parses but does the wrong thing -
 * worse than not converting it at all.
 *
 * NOTE THE OFF-BY-ONE: VBScript's Mid and InStr are 1-based, .NET's Substring and IndexOf are
 * 0-based. Getting this wrong silently shifts every string operation by one character.
 */
const MULTI_ARG = {
  // Through the shim, not inline .NET. Correctness first (VbsMid clamps and returns "" out of
  // range where a bare Substring throws), readability second: `VbsMid $s 2 3` is reviewable,
  // `$s.Substring((2) - 1, 3)` hides whether the off-by-one was handled at all.
  replace: (a) => a.length >= 3 ? `(VbsReplace ${wrapArg(a[0])} ${wrapArg(a[1])} ${wrapArg(a[2])})` : null,
  mid:     (a) => a.length === 3 ? `(VbsMid ${wrapArg(a[0])} ${wrapArg(a[1])} ${wrapArg(a[2])})`
                : a.length === 2 ? `(VbsMid ${wrapArg(a[0])} ${wrapArg(a[1])})` : null,
  left:    (a) => a.length === 2 ? `(VbsLeft ${wrapArg(a[0])} ${wrapArg(a[1])})` : null,
  right:   (a) => a.length === 2 ? `(VbsRight ${wrapArg(a[0])} ${wrapArg(a[1])})` : null,
  // InStr has TWO shapes: InStr(haystack, needle) and InStr(start, haystack, needle [, compare]).
  // Only the 2-arg form was handled, so the 4-arg form - which is what any case-insensitive search
  // uses - fell through unconverted. The compare flag is dropped and noted rather than guessed at:
  // vbTextCompare is case-INSENSITIVE and .IndexOf defaults to case-sensitive, so silently
  // discarding it would invert the test.
  instr:   (a) => a.length >= 2 && a.length <= 4
                ? `(VbsInStr ${a.slice(0, 3).map(wrapArg).join(" ")})` : null,
  split:   (a) => a.length >= 1 ? `(VbsSplit ${a.slice(0, 2).map(wrapArg).join(" ")})` : null,
  join:    (a) => a.length >= 1 ? `(VbsJoin ${a.slice(0, 2).map(wrapArg).join(" ")})` : null,
  ubound:  (a) => a.length >= 1 ? `(VbsUBound ${wrapArg(a[0])})` : null,
  lbound:  () => "0",

  // Conversions belong here, not in the flat replacements: their argument is routinely an
  // expression with its own parentheses - CStr(100 + Minute(v)) - which a [^()]* regex cannot
  // match, so those calls were emitted untouched and PowerShell rejected the bare name.
  cstr:    (a) => a.length === 1 ? `(VbsCStr ${wrapArg(a[0])})` : null,
  cint:    (a) => a.length === 1 ? `(VbsCInt ${wrapArg(a[0])})` : null,
  clng:    (a) => a.length === 1 ? `(VbsCLng ${wrapArg(a[0])})` : null,
  // Passed straight through before, so the output called a function PowerShell does not have.
  // It also has to handle the "&H41" STRING form, which is how hex-decoding routines feed it.
  cbyte:   (a) => a.length === 1 ? `(VbsCByte ${wrapArg(a[0])})` : null,
  cdbl:    (a) => a.length === 1 ? `(VbsCDbl ${wrapArg(a[0])})` : null,
  cbool:   (a) => a.length === 1 ? `(VbsCBool ${wrapArg(a[0])})` : null,

  // Date and time parts. All of these are bare intrinsic calls in VBScript and have no PowerShell
  // equivalent by that name, so leaving them produced an unrecognised command.
  year:    (a) => a.length === 1 ? `([datetime](${a[0]})).Year` : null,
  month:   (a) => a.length === 1 ? `([datetime](${a[0]})).Month` : null,
  day:     (a) => a.length === 1 ? `([datetime](${a[0]})).Day` : null,
  hour:    (a) => a.length === 1 ? `([datetime](${a[0]})).Hour` : null,
  minute:  (a) => a.length === 1 ? `([datetime](${a[0]})).Minute` : null,
  second:  (a) => a.length === 1 ? `([datetime](${a[0]})).Second` : null,
  weekday: (a) => a.length >= 1 ? `((([datetime](${a[0]})).DayOfWeek -as [int]) + 1)` : null,

  // FormatDateTime's second argument is a FORMAT CODE, and the codes mean different things
  // (0 general, 1 long date, 2 short date, 3 long time, 4 short time). Each maps to a distinct
  // .NET format string, so the code has to be read rather than defaulted.
  formatdatetime: (a) => {
    if (a.length === 0) return null;
    const d = `([datetime](${a[0]}))`;
    const fmt = { "1": "D", "2": "d", "3": "T", "4": "t" }[String(a[1] ?? "0").trim()];
    return fmt ? `${d}.ToString("${fmt}")` : `${d}.ToString()`;
  },

  vartype: (a) => a.length === 1 ? `(VbsVarType ${wrapArg(a[0])})` : null,
  // ONE element gets the unary comma: `@(X)` where X is itself an array is that array, not an
  // array holding it, so `Array(Array(1, 2))` lost a level. `@(, X)` is one element whatever X is.
  array:   (a) => (a.length === 1 ? `@(, ${a[0]})` : `@(${a.join(", ")})`),
  // Round belongs here, not in the simple replacements: its first argument is usually an expression
  // containing parentheses - Round((a/b)*100, 1) - which a flat regex splits in the wrong place.
  round:   (a) => a.length === 2 ? `[math]::Round(${a[0]}, ${a[1]})`
                : a.length === 1 ? `[math]::Round(${a[0]})` : null,
  cdate:   (a) => a.length === 1 ? `[datetime](${a[0]})` : null,
  // DateAdd is deliberately absent. Its first argument is an interval CODE ("d", "yyyy", "n") and
  // by this point string literals are placeholders, so the code cannot be read to pick the right
  // .AddDays/.AddMonths method. Emitting a comment mid-expression would corrupt the line, and
  // guessing the interval would be silently wrong - so it is left alone for a human.
  instrrev:(a) => a.length === 2 ? `((${a[0]}).LastIndexOf(${a[1]}) + 1)` : null,
  // Each operand parenthesised. `String(intLen - Len(s), "0")` emitted `("0" * $intLen - (...))`,
  // which PowerShell groups as `("0" * $intLen) - (...)` - a string repeat followed by subtracting
  // a number FROM A STRING. The count is nearly always an expression here, because that is what a
  // pad is, so dropping the parens broke the common case rather than an exotic one.
  string:  (a) => a.length === 2 ? `((${a[1]}) * (${a[0]}))` : null,
  // The interval is decided at RUNTIME, not here. The first version switched on the literal
  // `"s"` / `"d"` at translation time and never matched anything, because string literals are
  // still sentinel placeholders at this point in the pipeline - so every DateDiff in the corpus
  // silently failed to translate and shipped as a bare word. Handing the interval to the helper
  // sidesteps the ordering entirely, and lets it THROW on the calendar intervals (month, quarter,
  // year) whose answer a TimeSpan genuinely cannot give.
  datediff: (a) => (a.length >= 3 ? `(VbsDateDiff ${wrapArg(a[0])} ${wrapArg(a[1])} ${wrapArg(a[2])})` : null),
  dateadd:  (a) => (a.length === 3 ? `(VbsDateAdd ${wrapArg(a[0])} ${wrapArg(a[1])} ${wrapArg(a[2])})` : null),
  datepart: (a) => (a.length >= 2 && a.length <= 4 ? `(VbsDatePart ${wrapArg(a[0])} ${wrapArg(a[1])})` : null),
  dateserial: (a) => (a.length === 3 ? `(VbsDateSerial ${wrapArg(a[0])} ${wrapArg(a[1])} ${wrapArg(a[2])})` : null),
  // A TimeSpan: `DateSerial(...) + TimeSerial(...)`, the idiom this appears in, is then a DateTime.
  timeserial: (a) => (a.length === 3 ? `([timespan]::new(${wrapArg(a[0])}, ${wrapArg(a[1])}, ${wrapArg(a[2])}))` : null),
  // Invocation of a procedure held in a variable (see the GetRef handling in the main loop).
  __invoke: (a) => (a.length >= 1 ? `(& ${wrapArg(a[0])}${a.slice(1).map((x) => " " + wrapArg(x)).join("")})` : null),
  strcomp:  (a) => (a.length === 2 || a.length === 3
    ? `(VbsStrComp ${wrapArg(a[0])} ${wrapArg(a[1])} ${a[2] ? wrapArg(a[2]) : 0})` : null),
  formatnumber: (a) => (a.length >= 1 && a.length <= 5
    ? `(VbsFormatNumber ${wrapArg(a[0])} ${a[1] ? wrapArg(a[1]) : -1})` : null),
  // In an EXPRESSION only - `If MsgBox(...) = vbYes` - the statement form is refused with advice
  // about the deployment context, because a dialog under an RMM hangs the job.
  msgbox:   (a) => (a.length >= 1 && a.length <= 5
    ? `(VbsMsgBox ${wrapArg(a[0])} ${a[1] ? wrapArg(a[1]) : 0} ${a[2] ? wrapArg(a[2]) : '""'})` : null),
  // The second argument is VBScript's "abbreviate" flag. It has to be honoured, not dropped: an
  // RFC 822 date needs "Mon" and "Sep", and GetMonthName has no two-argument overload to fall back
  // on. Passed to a helper so the flag can be an expression rather than a literal.
  weekdayname: (a) => (a.length >= 1 && a.length <= 2
    ? `(VbsWeekdayName ${wrapArg(a[0])} ${a[1] ? wrapArg(a[1]) : "$false"})` : null),
  monthname: (a) => (a.length >= 1 && a.length <= 2
    ? `(VbsMonthName ${wrapArg(a[0])} ${a[1] ? wrapArg(a[1]) : "$false"})` : null),
};

function translateMultiArg(code) {
  const names = Object.keys(MULTI_ARG).filter((n) => !USER_FUNCS.has(n)).join("|");
  if (!names) return code;
  // The `:` is load-bearing and was missing. `round` emits `[math]::Round(`, whose `Round` was
  // still matchable, so each pass re-wrapped its own output and the guard below ran all 50 times:
  // `x = Round(a/b, 2)` came out as fifty-one `[math]::` prefixes and an instant parse failure.
  // replaceCall1 carries the identical guard - this is the same bug in the second of two tables.
  const rx = new RegExp(`(?<![\\w$.:])(${names})\\s*\\(`, "i");
  for (let guard = 0; guard < 50; guard++) {
    const m = rx.exec(code);
    if (!m) break;

    const open = m.index + m[0].length - 1;
    let depth = 0, close = open;
    for (; close < code.length; close++) {
      if (code[close] === "(") depth++;
      else if (code[close] === ")") { depth--; if (depth === 0) break; }
    }
    if (close >= code.length) break;                       // unbalanced; leave the line alone

    const args = splitArgs(code.slice(open + 1, close)).map((a) => translateMultiArg(a.trim()));
    const built = MULTI_ARG[m[1].toLowerCase()](args);
    if (built === null) break;                             // wrong arity: do not guess
    code = code.slice(0, m.index) + built + code.slice(close + 1);
  }
  return code;
}

/** VBScript intrinsic constants. Numeric values, per the MsgBox/VarType/FormatDateTime tables. */
const VBCONST = { vbyes: 6, vbno: 7, vbok: 1, vbcancel: 2, vbabort: 3, vbretry: 4, vbignore: 5,
                  vbokonly: 0, vbokcancel: 1, vbyesnocancel: 3, vbyesno: 4, vbretrycancel: 5,
                  vbcritical: 16, vbquestion: 32, vbexclamation: 48, vbinformation: 64,
                  vbempty: 0, vbnull: 1, vbinteger: 2, vblong: 3, vbstring: 8, vbboolean: 11,
                  vbbinarycompare: 0, vbtextcompare: 1,
                  vbgeneraldate: 0, vblongdate: 1, vbshortdate: 2, vblongtime: 3, vbshorttime: 4,
                  vbobjecterror: -2147221504,
                  vbdate: 7, vbsingle: 4, vbdouble: 5, vbcurrency: 6, vbobject: 9, vberror: 10,
                  vbvariant: 12, vbdataobject: 13, vbdecimal: 14, vbbyte: 17, vbarray: 8192,
                  vbtrue: -1, vbfalse: 0, vbusedefault: -2,
                  // MsgBox modality and default-button flags, and the weekday/first-week tables.
                  vbapplicationmodal: 0, vbsystemmodal: 4096, vbmsgboxhelpbutton: 16384,
                  vbmsgboxsetforeground: 65536, vbmsgboxright: 524288, vbmsgboxrtlreading: 1048576,
                  vbdefaultbutton1: 0, vbdefaultbutton2: 256, vbdefaultbutton3: 512, vbdefaultbutton4: 768,
                  vbabortretryignore: 2,
                  vbsunday: 1, vbmonday: 2, vbtuesday: 3, vbwednesday: 4, vbthursday: 5, vbfriday: 6, vbsaturday: 7,
                  vbusesystem: 0, vbusesystemdayofweek: 0, vbfirstjan1: 1, vbfirstfourdays: 2, vbfirstfullweek: 3 };

function translateBuiltins(code) {
  // Constants FIRST, before any call is translated. FormatDateTime's emitter reads the numeric
  // format code to choose a .NET format string, so it has to see 3 rather than vbLongTime.
  let c = code.replace(/(?<![\w$.])(vb\w+)\b/g, (mm, name) => {      // \w: vbDefaultButton2 has a digit
    const v = VBCONST[name.toLowerCase()];
    return v === undefined ? mm : String(v);
  });
  c = translateMultiArg(c);
  // UBound and LBound are handled in MULTI_ARG, which splits arguments on BALANCED parentheses.
  // The old flat regex used [^)]+ and so broke on a nested call: UBound(Split(s, ",")) matched only
  // `Split(s, ","` and produced `(@(("a,b,c").Count - 1).Split(","))`, which parses and throws.
  // BALANCED-paren scanning, not `[^()]*`. The old patterns could not span a nested call, so
  // `Trim(UTF8Decode(x))` was left completely untranslated and the output called a VBScript
  // builtin that PowerShell does not have. Library-style scripts nest constantly, and the comment
  // directly above records fixing this same bug for UBound/LBound - the single-argument helpers
  // kept the broken form for another ten functions.
  // [string] CAST, not a bare parenthesis. VBScript coerces the argument to a string first, so
  // Trim(Empty) is "" and Len(Empty) is 0. Emitting `($x).Trim()` throws "You cannot call a method
  // on a null-valued expression" the moment the value is Empty - which happens on an empty
  // argument, a function that fell off its last branch, or any uninitialised accumulator.
  // The cast is also what makes Len/UCase work on a number, which VBScript allows.
  c = replaceCall1(c, "Len", (s) => `([string](${s})).Length`);
  c = replaceCall1(c, "UCase", (s) => `([string](${s})).ToUpper()`);
  c = replaceCall1(c, "LCase", (s) => `([string](${s})).ToLower()`);
  c = replaceCall1(c, "Trim", (s) => `([string](${s})).Trim()`);
  c = replaceCall1(c, "LTrim", (s) => `([string](${s})).TrimStart()`);
  c = replaceCall1(c, "RTrim", (s) => `([string](${s})).TrimEnd()`);
  // Conversions go through the shim: VBScript's True is -1 (not 1) and CInt uses BANKER'S rounding,
  // so `[int]` is wrong on both counts.
  c = replaceCall1(c, "StrReverse", (s) => `(VbsStrReverse (${s}))`);
  c = replaceCall1(c, "IsNumeric", (s) => `((${s}) -as [double]) -ne $null`);
  c = replaceCall1(c, "IsObject", (s) => `((${s}) -is [object])`);
  c = replaceCall1(c, "IsNull", (s) => `((${s}) -eq $null)`);
  c = replaceCall1(c, "IsEmpty", (s) => `((${s}) -eq $null)`);
  // WScript host properties. Left alone these emit `wScript.ScriptName`, which PowerShell parses as
  // a property access on an undefined variable and rejects - 19 files failed first on exactly this.
  // Command-line arguments. These were passed through verbatim, so the output called into a
  // WScript host that does not exist and died on the first argument check - which is the very
  // first thing most deployed scripts do.
  //
  // VBScript's collection is 0-based like $args, so the indexes carry over unchanged. Ordered
  // longest-pattern-first: `.Count` and `.Item(n)` have to match before the bare collection.
  c = c.replace(/\bWScript\.Arguments\.Count\b/gi, "$args.Count");
  c = c.replace(/\bWScript\.Arguments\.Item\s*\(\s*([^)]+?)\s*\)/gi, "$args[$1]");
  c = c.replace(/\bWScript\.Arguments\s*\(\s*(\d+)\s*\)/gi, "$args[$1]");
  // `WScript.CreateObject(x)` is the same thing as `CreateObject(x)` for these purposes, and the
  // plain form is already handled. Stripping the prefix here routes it to that rule instead of
  // leaving a bare word. It hid in the ONE function that launches the process, so $objShell stayed
  // $null and nothing was ever launched - the script ran to completion doing nothing.
  c = c.replace(/\bWScript\.CreateObject\s*\(/gi, "CreateObject(");

  // Named arguments. Deployment tools pass `/Key:Value` tokens and scripts read them through
  // WScript.Arguments.Named, which was flagged rather than translated - and since it is usually
  // the FIRST executable line of the main body, nothing after it ran. The helper returns a
  // hashtable-backed object exposing .Exists() and .Item(), so every call site converts unchanged.
  c = c.replace(/\bWScript\.Arguments\.Named\b/gi, "(VbsNamedArgs)");
  c = c.replace(/\bWScript\.Arguments\.Unnamed\b/gi, "(VbsUnnamedArgs)");
  c = c.replace(/\bWScript\.Arguments\b(?!\s*\.\s*(?:Named|Unnamed)\b)/gi, "$args");

  // `.Item(n)` is VBScript's explicit collection accessor and survives assignment: after
  // `Set colArgs = WScript.Arguments`, the script says `colArgs.Item(0)`. That becomes a
  // PowerShell ARRAY, which has no .Item() method, so the line throws at runtime.
  //
  // Only rewritten on a plain variable, never on a member path like `objDoc.Fields.Item(x)` -
  // a real COM collection DOES have .Item() and rewriting it there would break working code. The
  // lookbehind is what draws that line. No `$` in the pattern: this runs BEFORE sigilisation, so
  // requiring one matched nothing at all.
  c = c.replace(/(?<![.\w$])([A-Za-z_]\w*)\.Item\s*\(\s*([^()]+?)\s*\)/g, "$1[$2]");

  // $PSCommandPath, NOT $MyInvocation.MyCommand. Inside a function $MyInvocation describes the
  // FUNCTION, so WScript.ScriptName returned "ContentLog" instead of the script's filename and the
  // log file was written to ContentLogs\ContentLog.log. It is correct at top level and wrong
  // everywhere else, which is the worst way for it to be wrong. $PSCommandPath is script-scoped
  // and means the same thing from anywhere in the file.
  c = c.replace(/\bWScript\.ScriptFullName\b/gi, "$PSCommandPath");
  c = c.replace(/\bWScript\.ScriptName\b/gi, "(Split-Path -Leaf $PSCommandPath)");
  c = c.replace(/\bWScript\.Path\b/gi, "$PSScriptRoot");
  c = c.replace(/\bWScript\.StdOut\.WriteBlankLines\b/gi, "VbsBlankLines");     // 74 corpus lines
  c = c.replace(/\bWScript\.StdOut\.WriteLine\b/gi, "Write-Output");
  c = c.replace(/\bWScript\.StdOut\.Write\b/gi, "Write-Host -NoNewline");
  c = c.replace(/\bWScript\.StdErr\.WriteLine\b/gi, "Write-Error");
  c = c.replace(/\bWScript\.StdErr\.Write\b/gi, "[Console]::Error.Write");
  // The stream OBJECTS themselves (`Set out = WScript.StdOut`): the .NET writers have the same
  // Write/WriteLine members.
  c = c.replace(/\bWScript\.StdOut\b/gi, "[Console]::Out").replace(/\bWScript\.StdErr\b/gi, "[Console]::Error");
  c = c.replace(/\bWScript\.(?:Version|BuildVersion)\b/gi, "($PSVersionTable.PSVersion.ToString())");
  // `CreateObject(...)` INSIDE an expression - as an argument, or with a member chained on it.
  // The statement form `Set x = CreateObject(...)` has its own rule and never reaches here.
  c = replaceCall1(c, "CreateObject", (s) => `(New-Object -ComObject ${s})`);
  c = replaceCall1(c, "IsArray", (s) => `((${s}) -is [array])`);
  c = replaceCall1(c, "IsDate", (s) => `(VbsIsDate (${s}))`);
  c = replaceCall1(c, "DateValue", (s) => `([datetime](${s})).Date`);
  c = replaceCall1(c, "CSng", (s) => `[single](${s})`);
  // StdIn has no PowerShell object equivalent; these are the cmdlets that do the same job.
  // Parens OPTIONAL: VBScript lets you call a parameterless method without them, and the bare-call
  // fallback then appended `()` to the untouched text, reproducing the original verbatim.
  c = c.replace(/\bWScript\.StdIn\.ReadLine\s*(?:\(\s*\))?/gi, "(Read-Host)");
  c = c.replace(/\bWScript\.StdIn\.ReadAll\s*(?:\(\s*\))?/gi, "([Console]::In.ReadToEnd())");
  c = c.replace(/\bWScript\.Interactive\b/gi, "$true");
  c = c.replace(/\bWScript\.FullName\b/gi, "([Diagnostics.Process]::GetCurrentProcess().Path)");

  // ChrW/AscW before Chr/Asc: the shorter name is a prefix of the longer one, so the un-suffixed
  // rule would match first and leave a stray `W(`.
  // `[char][int]`, never a bare `[char]`. VBScript's ChrW takes any numeric, and arithmetic here
  // routinely produces a DOUBLE - `val * 2 ^ 6 + b` goes through [math]::Pow, which returns
  // double. PowerShell refuses to cast double to char ("Cannot convert value 233 to type
  // System.Char"), so any non-ASCII path threw while ASCII sailed through. The converter
  // introduced that divergence itself by choosing Pow; the [int] is what pays for it.
  // Wrapped in their own parens. A bare cast used as a command argument is read as a STRING:
  // `Write-Output [char][int]($v)` prints the literal text "[char][int]" and then 233 on the next
  // line, rather than the character. It only showed up once a case echoed the cast directly -
  // every earlier one passed through VbsConcat, which parenthesises its arguments anyway.
  c = replaceCall1(c, "ChrW", (s) => `([char][int](${s}))`);
  // AscW IS the UTF-16 code unit, so [int][char] is exactly right for it. Asc is NOT: it returns
  // the ANSI code-page byte, and the two only agree below 256. Splitting them costs nothing and
  // removes an assumption that would eventually be wrong on somebody's script.
  c = replaceCall1(c, "AscW", (s) => `([int][char](${s}))`);
  c = replaceCall1(c, "Chr", (s) => `([char][int](${s}))`);
  c = replaceCall1(c, "Asc", (s) => `(VbsAsc (${s}))`);
  // The Err object gets a real backing variable. VBScript's Err carries the LAST error and
  // survives the failing statement, which is what `If Err.Number <> 0 Then` depends on. PowerShell
  // has no equivalent that behaves the same way ($Error is a list and $? is not an error code), so
  // the converter supplies one and the try/catch wrapping below populates it.
  c = c.replace(/\bErr\s*\.\s*Number\b/gi, "$script:VbsErr.Number");
  // Bare `Err` in an expression - `If Err <> 0` - is Err.Number, its default property.
  c = c.replace(/(?<![\w$.])Err\b(?!\s*[.(])/g, "$script:VbsErr.Number");
  c = c.replace(/\bErr\s*\.\s*Description\b/gi, "$script:VbsErr.Description");
  c = c.replace(/\bErr\s*\.\s*Source\b/gi, "$script:VbsErr.Source");

  // Locale and date intrinsics. Left unmapped these were bare commands; now that untranslated
  // intrinsics are flagged they at least failed loudly, but these five have exact equivalents and
  // a flag is a poor substitute for a translation anyone could write.
  //
  // SetLocale/GetLocale change the thread culture, which is what the date formatting around them
  // is usually compensating FOR - the honest translation is the culture swap, and the reader can
  // then decide the whole dance is unnecessary.
  c = replaceCall1(c, "SetLocale",
    (s) => `([Threading.Thread]::CurrentThread.CurrentCulture = [cultureinfo]::GetCultureInfo(${s}))`);
  c = replaceCall1(c, "GetLocale", () => `([Threading.Thread]::CurrentThread.CurrentCulture.LCID)`);
  // WeekdayName and MonthName are handled in MULTI_ARG, NOT here: both take an optional
  // "abbreviate" flag, and a single-argument scanner swallowed `n, True` whole and emitted
  // `((n, $true) - 1) % 7` - subtracting from an array. They also have to honour that flag,
  // since RFC 822 and most log formats want "Mon" and "Sep" rather than the full names.
  c = replaceCall1(c, "Hex", (s) => `[Convert]::ToString(${s}, 16)`);
  c = replaceCall1(c, "Oct", (s) => `[Convert]::ToString(${s}, 8)`);
  c = replaceCall1(c, "Abs", (s) => `[math]::Abs(${s})`);
  c = replaceCall1(c, "Sgn", (s) => `[math]::Sign(${s})`);
  c = replaceCall1(c, "Sqr", (s) => `[math]::Sqrt(${s})`);
  c = replaceCall1(c, "Sin", (s) => `[math]::Sin(${s})`);
  c = replaceCall1(c, "Cos", (s) => `[math]::Cos(${s})`);
  c = replaceCall1(c, "Tan", (s) => `[math]::Tan(${s})`);
  c = replaceCall1(c, "Atn", (s) => `[math]::Atan(${s})`);
  c = replaceCall1(c, "Exp", (s) => `[math]::Exp(${s})`);
  c = replaceCall1(c, "Log", (s) => `[math]::Log(${s})`);
  c = replaceCall1(c, "TypeName", (s) => `(VbsTypeName (${s}))`);
  // Rnd, with or without an (ignored) argument: a double in [0, 1).
  c = replaceCall1(c, "Rnd", () => "(Get-Random -Minimum 0.0 -Maximum 1.0)");   // balanced: Rnd(x + 1) too
  c = c.replace(/(?<![\w$.])Rnd\b(?![\w(])/gi, "(Get-Random -Minimum 0.0 -Maximum 1.0)");
  c = replaceCall1(c, "Int", (s) => `[math]::Floor(${s})`);
  c = replaceCall1(c, "Fix", (s) => `[math]::Truncate(${s})`);
  // The B (byte) variants, on the BSTR model VBScript uses: two bytes per character.
  c = replaceCall1(c, "LenB", (s) => `(2 * ([string](${s})).Length)`);
  c = replaceCall1(c, "AscB", (s) => `([int][char](([string](${s}))[0]) -band 0xFF)`);
  c = replaceCall1(c, "ChrB", (s) => `([char][int](${s}))`);
  c = replaceCall1(c, "Space", (s) => `(" " * (${s}))`);
  // Round is handled by MULTI_ARG above and MUST NOT be repeated here. Two rules for one function
  // is what produced `[math]::[math]::Round(...)`: MULTI_ARG emitted `[math]::Round(`, then these
  // two lines matched their own output - `\b` happily matches straight after `::` - and wrapped it
  // a second time. The comment at MULTI_ARG.round even says Round belongs there "not in the simple
  // replacements", and the simple replacement was left behind anyway. The tell was the double
  // space in `,  2`: MULTI_ARG trims its arguments and this regex did not.


  c = c.replace(/\bNow\s*\(\s*\)/gi, "(Get-Date)").replace(/\bNow\b(?!\s*\()/gi, "(Get-Date)");
  // Bare `Date`, `Time` and `Timer` too - VBScript reserves all three, so a bare word is always
  // the intrinsic (a script cannot assign to them). `DatePart("m", Date)` reached the output with
  // the argument as a bare word. Timer is seconds since midnight.
  // With parentheses only, here. The BARE forms are resolved in `fin` after sigilize has run,
  // because scripts do declare `Dim time` (adegard's timetrigger) and a bare word at this stage
  // cannot tell the variable from the intrinsic. `-` in the lookbehind: `Get-Date` is not `Date`.
  c = c.replace(/(?<![\w$.\-])Date\s*\(\s*\)/gi, "(Get-Date).Date");
  c = c.replace(/(?<![\w$.\-])Time\s*\(\s*\)/gi, "(Get-Date)");
  c = c.replace(/(?<![\w$.\-])Timer\s*\(\s*\)/gi, "((Get-Date).TimeOfDay.TotalSeconds)");
  // `#1/1/1601#` is a date literal.
  c = c.replace(/#(\d[^#]*)#/g, '([datetime]"$1")');
  // `WScript.Arguments.Named("x")` - the helper returns a Dictionary, so index it as one.
  c = c.replace(/\(VbsNamedArgs\)\s*\(/g, "(VbsNamedArgs).Item(");
  c = c.replace(/\bvbCrLf\b/gi, '"`r`n"').replace(/\bvbNewLine\b/gi, '"`r`n"').replace(/\bvbTab\b/gi, '"`t"')
       .replace(/\bvbLf\b/gi, '"`n"').replace(/\bvbCr\b/gi, '"`r"')
       .replace(/\bvbNullString\b/gi, '""').replace(/\bvbNullChar\b/gi, '"`0"')
       .replace(/\bvbBack\b/gi, '"`b"').replace(/\bvbFormFeed\b/gi, '"`f"').replace(/\bvbVerticalTab\b/gi, '"`v"');
  return c;
}

// ── the rule table ───────────────────────────────────────────────────────────
// Ordered roughly by how often the construct appears in the corpus.
const RULES = [
  // `Call Foo(a, b)` is VBScript's explicit-invocation form. `Call` is not a PowerShell keyword,
  // so it emitted a call to a command named "Call". Stripping it leaves the ordinary call, which
  // the rules below already handle. It tends to hide on error paths, which is why it survived
  // three reviews - here it only fired when the client directory was missing.
  { name: "Call", rx: /^\s*Call\s+(.+)$/i,
    to: (m, ctx) => ctx.statement(m[1].replace(/^\((.*)\)$/s, "$1").trim()) },

  { name: "Option Explicit", rx: /^\s*Option\s+Explicit\s*$/i, to: () => "Set-StrictMode -Version 2.0" },

  // `Dim arr(), strLen, i, sT` mixes an array declaration with plain scalars. The old rule ran a
  // single replace over the whole list and emitted the leftovers verbatim, so the scalars arrived
  // as a bare comma-separated expression. Declarations are handled per name now: arrays become an
  // empty array, scalars vanish (PowerShell needs no declaration).
  { name: "Dim", rx: /^\s*Dim\s+(.+)$/i,
    to: (m, ctx) => {
      // Inside a class, `Dim x, y` declares PROPERTIES, not locals - and this rule sits above the
      // class-member rule, so it was swallowing them and leaving every reference inside a method
      // bare. PowerShell then rejected those with MissingThis.
      if (ctx.inClass())
        return splitArgs(m[1]).map((n) => ctx.member(n.trim().replace(/\s*\(.*$/, ""))).join("\n");

      const decls = splitArgs(m[1])
        .map((d) => {
          // TWO-DIMENSIONAL: `Dim a(10, 5)`. VBScript bounds are INCLUSIVE, so both get +1, and
          // PowerShell needs a real rectangular array - `object[,]` - for `$a[$i, $j]` to work.
          // Without this the declaration produced a one-dimensional array and every write to it
          // failed at runtime.
          const two = d.trim().match(/^([A-Za-z_]\w*)\s*\(\s*(\d+)\s*,\s*(\d+)\s*\)$/);
          if (two) {
            return `$${two[1]} = New-Object 'object[,]' ${Number(two[2]) + 1},${Number(two[3]) + 1}`;
          }

          const a = d.trim().match(/^([A-Za-z_]\w*)\s*\(\s*(\d*)\s*\)$/);
          if (!a) {
            // A plain scalar USED to emit nothing, on the reasoning that PowerShell needs no
            // declaration. It does here: we translate `Option Explicit` into
            // `Set-StrictMode -Version 2.0`, and StrictMode THROWS on reading a variable that was
            // never assigned. VBScript's Dim creates it as Empty, and code leans on that -
            // `sR = sR & chrw(b)` accumulating onto an undeclared string is completely ordinary.
            // So Dim has to produce the Empty, which is $null on this side.
            const name = d.trim().match(/^([A-Za-z_]\w*)$/);
            if (!name) return null;
            // Only the accumulators. See findVarsReadBeforeAssign - emitting this for every
            // declaration doubled the line count of every real script.
            return ctx.skipInit(name[1]) ? null : `$${name[1]} = $null`;
          }
          return a[2]
            ? `$${a[1]} = New-Object object[] ${Number(a[2]) + 1}`
            : `$${a[1]} = @()`;                                   // Dim arr() - dynamic array
        })
        .filter(Boolean);
      return decls.length ? decls.join("\n") : null;
    } },

  // The bound expression MUST go through ctx.expr like every other rule here. Interpolating m[2]
  // raw emitted `(arraysize + 1 + 1 - ...)` with no sigil on the variable, which PowerShell reads
  // as a command named `arraysize`. Parses fine, dies at runtime, and under Set-StrictMode takes
  // the array with it.
  // Without Preserve the old contents are discarded, so allocate fresh: the extend form reads the
  // variable first, and when ReDim is its first mention (VBScript lets ReDim declare) a class
  // method fails with "Variable is not assigned in the method".
  { name: "ReDim", rx: /^\s*ReDim\s+(Preserve\s+)?([A-Za-z_]\w*)\s*\(\s*(.+?)\s*\)\s*$/i,
    to: (m, ctx) => {
      if (m[1]) return `$${m[2]} = @($${m[2]}) + @($null) * (${ctx.expr(m[3])} + 1 - @($${m[2]}).Count)`;
      const dims = splitArgs(m[3]).map((d) => `(${ctx.expr(d)} + 1)`);
      return dims.length === 1
        ? `$${m[2]} = @($null) * ${dims[0]}`
        : `$${m[2]} = New-Object 'object[${",".repeat(dims.length - 1)}]' ${dims.join(",")}`;
    } },

  // ONE Const can declare SEVERAL constants: `Const A = "x", B = "y"`. Capturing the first name
  // and then everything after the first `=` produced
  // `Set-Variable -Name A -Value "x", $B = "y" -Option Constant` - which makes A an ARRAY of both
  // values and leaves B an ordinary variable that is not constant at all. Split on top-level
  // commas and emit one statement each.
  { name: "Const", rx: /^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?Const\s+(.+)$/i,
    to: (m, ctx) => {
      const out = splitArgs(m[1])
        .map((d) => {
          const one = d.trim().match(/^([A-Za-z_]\w*)\s*=\s*(.+)$/);
          if (!one) return null;
          // At CLASS level a constant is a member: a class body cannot hold a statement, and a
          // method could not see the Set-Variable anyway. ctx.member registers the name so uses
          // inside methods get `$this.`; the value replaces its `$null`.
          if (ctx.inClass()) return ctx.member(one[1]).replace(/= \$null$/, `= ${ctx.expr(one[2])}`);
          return `Set-Variable -Name ${one[1]} -Value ${ctx.expr(one[2])} -Option Constant`;
        })
        .filter(Boolean);
      return out.length ? out.join("\n") : null;
    } },

  // `WScript.` prefix optional. The prefix is stripped later in translateBuiltins, but the RULES
  // table runs FIRST, so `Set x = WScript.CreateObject(...)` missed this rule, fell through to the
  // generic assignment path, and came out as a bare `CreateObject(...)` call.
  { name: "CreateObject", rx: /^\s*Set\s+([A-Za-z_]\w*)\s*=\s*(?:WScript\.)?CreateObject\s*\(\s*(.+?)\s*\)\s*$/i,
    // `Set Foo = CreateObject(...)` where Foo is the enclosing procedure is the RETURN idiom, and
    // this rule matched before the Set-assignment rule that knows about it. So a method whose only
    // job was to build and return a Dictionary assigned the value and never returned it -
    // "Not all code path returns value within method" on a [object] method.
    to: (m, ctx) => {
      if (ctx.isFunc(m[1])) ctx.noteReturnAssign(m[1]);
      return `$${m[1]} = New-Object -ComObject ${m[2]}`;
    } },

  // Matches the PLACEHOLDER, not the literal: by this point strings have been protected, so a rule
  // whose regex contains "winmgmts:..." can never fire. It has to restore the argument and inspect it.
  // Matches ANY argument, not just a single literal. The real-world form is concatenated -
  // GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") - so a
  // rule requiring one string placeholder never fired, and the line fell through to a bare
  // `GetObject(...)` that PowerShell cannot parse. Largest single cause on the 396-script corpus.
  // NOT when an ExecQuery is chained onto it. This rule sits above the ExecQuery rule and its
  // `.+` is greedy, so `Set c = GetObject("winmgmts:...").ExecQuery(q)` matched HERE first and
  // emitted the namespace string, silently discarding the query.
  { name: "GetObject winmgmts", rx: /^\s*Set\s+([A-Za-z_]\w*)\s*=\s*GetObject\s*\((?!.*\.ExecQuery\s*\()(.+)\)\s*$/i,
    to: (m, ctx) => {
      const arg = ctx.raw(m[2]);
      if (!/winmgmts:/i.test(arg)) return SKIP;

      // An OBJECT PATH binds ONE instance: `winmgmts:...\root\cimv2:Win32_Service.Name='X'`. That
      // is a query, not a namespace, and emitting the namespace string for it was the worst output
      // this converter has produced: the script then read `.State` off a string, never saw
      // "Stopped", and spun forever. Silent, and only in the restart path.
      //
      // Handles both the all-literal form and the concatenated one, which is what a per-item
      // re-query always looks like.
      const path = /:([A-Za-z_]\w*)\.([A-Za-z_]\w*)\s*=\s*'(.*)$/i.exec(arg);
      if (path) {
        const [, cls, key, rest] = path;
        // Two shapes. Concatenated - `...Name='" & objService.Name & "'"` - where `rest` starts
        // with the DOUBLE quote that closes the VBScript literal, not a single one. And
        // all-literal, `...Name='Spooler'`, which only counts when there is no concatenation left
        // in it; otherwise the literal pattern happily matches up to the first `'` and swallows
        // the VBScript expression raw, which is how this first emitted
        // `-Filter "Name='" & objService.Name & "'"`.
        // The value term must be SIMPLE - `objService.Name`, not an arbitrary expression. A lazy
        // `.+?` here ran straight past a nested GetObject on one corpus script and produced an
        // unclosed `+` expression. A key value in a WMI object path is a variable or a member
        // access in every real case; anything else falls through to the namespace handling below,
        // which at least parses.
        const concat = /^"\s*[&+]\s*([A-Za-z_][\w.]*)\s*[&+]\s*"'/.exec(rest);
        const literal = !/[&+]/.test(rest) && /^([^']*)'/.exec(rest);
        if (concat) {
          return `$${m[1]} = Get-CimInstance -ClassName ${cls} `
               + `-Filter ("${key}='" + ${ctx.expr(concat[1])} + "'")`;
        }
        if (literal) {
          return `$${m[1]} = Get-CimInstance -ClassName ${cls} -Filter "${key}='${literal[1]}'"`;
        }
      }

      // Keep the remote computer if the script targeted one: it has to reappear as
      // -ComputerName on the Get-CimInstance calls, and dropping it would silently turn a
      // remote query into a local one.
      const remote = /\\\\"\s*[+&]\s*([A-Za-z_]\w*)/.exec(arg);
      const ns = (/\\(root\\[\w\\]+)/i.exec(arg) || [, "root\\cimv2"])[1];
      return `$${m[1]} = "${ns}"   # WMI namespace`
           + (remote ? `\n# NOTE: this connected to $${remote[1]} - add -ComputerName $${remote[1]} to the Get-CimInstance calls below.` : "")
           + `\n# The ExecQuery calls below become Get-CimInstance.`;
    } },

  // `$x = GetObject("LDAP://...")` is not merely unidiomatic, it does not PARSE: a bare identifier
  // followed by `(` in expression position is a syntax error in PowerShell. It was the single
  // largest first-failure on the expanded corpus. [ADSI] is the real equivalent and converts cleanly,
  // so this is worth translating rather than refusing.
  // `GetObject(, "Excel.Application")` - the running instance, not a new one.
  { name: "GetObject active", rx: /^\s*Set\s+([A-Za-z_]\w*)\s*=\s*GetObject\s*\(\s*,\s*(.+?)\s*\)\s*$/i,
    to: (m, ctx) => `$${m[1]} = [Runtime.InteropServices.Marshal]::GetActiveObject(${ctx.expr(m[2])})` },

  { name: "GetObject ADSI", rx: /^\s*Set\s+([A-Za-z_]\w*)\s*=\s*GetObject\s*\(\s*(.+?)\s*\)\s*$/i,
    to: (m, ctx) => {
      const arg = ctx.raw(m[2]);
      if (/winmgmts:/i.test(arg)) return SKIP;                    // handled by the WMI rule above
      // A VARIABLE holding the path - `GetObject(strGroupPath)` - is the same bind when the script
      // builds WinNT:// or LDAP:// paths anywhere; that is what those variables hold.
      const bareVar = /^[A-Za-z_]\w*$/.test(arg) && ctx.usesAdsi();
      if (!bareVar && !/(WinNT|LDAP|GC|IIS):\/\//i.test(arg)) return SKIP;   // some other moniker; refused below
      const e = ctx.expr(m[2]);
      return `$${m[1]} = [ADSI]${/[\s+]/.test(e.trim()) ? `(${e})` : e}`;
    } },

  // The left of `.ExecQuery` may be a CHAINED GetObject, not just a variable:
  // `Set c = GetObject("winmgmts:root\cimv2").ExecQuery(q)` is as common as the two-line form.
  // Matching only a bare variable meant this fell through to the GetObject rule above, which
  // emitted the NAMESPACE STRING and dropped the query entirely - so the collection became
  // "root\cimv2" and every loop over it read properties off a string. Silent, and it produced an
  // infinite wait loop in one script.
  { name: "WMI ExecQuery", rx: /^\s*Set\s+([A-Za-z_]\w*)\s*=\s*(?:[A-Za-z_]\w*|GetObject\s*\([^()]*\))\.ExecQuery\s*\(\s*(.+?)\s*\)\s*$/i,
    // Only the FIRST argument. ExecQuery takes optional flags - `ExecQuery(q, , 48)` is common for
    // forward-only enumeration - and passing those through produced `-Query "...", , 48`, which is
    // not a parseable argument list.
    // The query goes through ctx.expr. It was passed through RAW, so a query built by
    // concatenation - which is most of them, since they interpolate a process or service name -
    // kept its VBScript `&` and emitted
    // `-Query "Select ... Name = '" & strExe & "'"`. PowerShell reads `&` as the call operator, so
    // the line is a parse error, and a WQL query is exactly the place a name gets interpolated.
    // Wrapped in @(). A VBScript ExecQuery result is always a collection, and scripts read .Count
    // on it. Get-CimInstance returns a bare object for one match and $null for none, and under
    // Set-StrictMode reading .Count on either throws. @() makes the shape match what the original
    // code expects, and costs nothing when there are many matches.
    to: (m, ctx) => {
      const q = ctx.expr(splitArgs(m[2])[0]);       // m[2]: the source side is non-capturing now
      return `$${m[1]} = @(Get-CimInstance -Query ${/[\s+]/.test(q.trim()) ? `(${q})` : q})`;
    } },

  // The argument is PARENTHESISED on purpose. `Write-Output "a" + $b` passes three separate
  // arguments in PowerShell ("a", "+", $b) and prints all three on separate lines - it does not
  // concatenate. That is a silently-wrong conversion, which is the class of bug this tool exists
  // to avoid, so the parens go on whenever the expression is not a single atom.
  { name: "WScript.Echo", rx: /^\s*WScript\.Echo\s*(.*)$/i,
    to: (m, ctx) => {
      // [Console]::WriteLine, NOT Write-Output. WScript.Echo is a SIDE CHANNEL in VBScript: it
      // prints and contributes nothing to what the enclosing function returns. Write-Output is the
      // opposite - it puts the value on the pipeline, so a Sub that echoes three times returns an
      // array of three strings, and every caller testing that result gets a truthy array. Console
      // output is the faithful equivalent and cannot pollute a return value.
      //
      // Wrap unless the argument is a SINGLE ATOM. The old test looked for an operator or a space,
      // which missed `[math]::Round(2.6)` and `[char][int](65)` - no spaces, no operators, but as
      // a bare command argument PowerShell reads the type literal as a string and prints it on its
      // own line followed by the value. Asking "is this one token" instead of "does this look
      // complicated" gets both, and every future emitter shaped like them.
      if (!m[1].trim()) return "Write-Output ''";
      const e = ctx.expr(m[1]);
      // ALWAYS parenthesised. A cast must be followed by a parseable expression, and an
      // untranslated bare word like `Err.Description` is not one - `[string]Err.Description` is a
      // hard parse error where the old `Write-Output Err.Description` at least parsed as a command
      // with an argument. Wrapping keeps a line we could not fully convert from taking the file
      // down with it.
      return `[Console]::WriteLine([string](${e}))`;
    } },

  // The exit code is very often a CONSTANT, not a literal - `WScript.Quit EXIT_FAILURE` is the
  // house style in plenty of vendor-shipped scripts. The old `\d*` only matched digits, so those lines fell
  // through untouched and the output called a WScript host that is not there. On a deployment
  // tool that reads the exit code, that is the single worst line to get wrong.
  { name: "WScript.Quit", rx: /^\s*WScript\.Quit\s*(?:\(\s*(.*?)\s*\)|\s+(.+?))?\s*$/i,
    to: (m, ctx) => {
      const arg = (m[1] ?? m[2] ?? "").trim();
      return `exit ${arg ? ctx.expr(arg) : 0}`;
    } },

  { name: "WScript.Sleep", rx: /^\s*WScript\.Sleep\s*\(?\s*(.+?)\s*\)?\s*$/i,
    to: (m, ctx) => `Start-Sleep -Milliseconds ${ctx.expr(m[1])}` },

  { name: "Set = Nothing", rx: /^\s*Set\s+([A-Za-z_]\w*)\s*=\s*Nothing\s*$/i,
    to: (m, ctx) => { if (ctx.isFunc(m[1])) ctx.noteReturnAssign(m[1]); return `$${m[1]} = $null`; } },

  // `Set x = New Thing` instantiates either a VBScript class defined in the same file or one of the
  // handful of built-in creatable objects. Left alone the catch-all turned the type name into a
  // variable and emitted `New $RegExp`, which is not a statement at all.
  { name: "Set = New", rx: /^\s*Set\s+([A-Za-z_]\w*)\s*=\s*New\s+([A-Za-z_]\w*)\s*$/i,
    to: (m, ctx) => {
      // `Set Make = New T` inside Function Make IS the return. This rule never told closeFunc,
      // so a factory function built its object and returned nothing - silently, for as long as
      // the rule existed. The With-block conformance case is what finally showed it.
      if (ctx.isFunc(m[1])) ctx.noteReturnAssign(m[1]);
      // RegExp is a VBScript intrinsic, not a user class. Kept as the COM object so the
      // .Pattern/.Global/.IgnoreCase members and .Replace semantics stay identical - translating to
      // [regex] would change all three.
      if (/^RegExp$/i.test(m[2])) return `$${m[1]} = New-Object -ComObject VBScript.RegExp`;
      // A class THIS file declares is a type PowerShell can see at parse time. One from an
      // included file is not - `[VBSApp]::new()` is "Unable to find type" in 14 corpus files -
      // so it is created by name at run time, which is when the include has been dot-sourced.
      return ctx.isClass(m[2]) ? `$${m[1]} = [${m[2]}]::new()` : `$${m[1]} = (New-Object -TypeName ${m[2]})`;
    } },

  // `Set Foo = obj` is the FUNCTION-RETURN idiom when Foo is the enclosing function, exactly like
  // `Foo = value`. Function names are kept out of the variable registry so call sites stay intact,
  // so ctx.lhs left this one bare - and a bare name on the left of `=` does not assign in
  // PowerShell, it CALLS. `GetActiveSessionIDs = $dic` re-invoked the function with `=` and the
  // dictionary as arguments, recursing until the stack blew. Worse than a missing return, and it
  // parses perfectly.
  { name: "Set assignment", rx: /^\s*Set\s+(.+?)\s*=\s*(.+)$/i,
    to: (m, ctx) => {
      const lhs = m[1].trim();
      if (ctx.isFunc(lhs)) { ctx.noteReturnAssign(lhs); return `$${lhs} = ${ctx.expr(m[2])}`; }
      return `${ctx.lhs(m[1])} = ${ctx.expr(m[2])}`;
    } },

  // The body goes back through the WHOLE rule table, not just expression translation. Previously it
  // used ctx.stmt(), which does operators and sigils but knows nothing about statements - so
  // `If fs.FileExists(f) Then fs.DeleteFile f` emitted `{ $fs.DeleteFile $f }`, a paren-less method
  // call that PowerShell reads as a command invocation and rejects.
  // `If\b\s*`, not `If\s+`. `If(Is64) Then` with no space is legal VBScript and appears in real
  // vendor code; requiring whitespace dropped the whole statement to an unrecognised-line TODO,
  // which left the guarded body running UNCONDITIONALLY. A tokenizer gap that silently deletes a
  // condition is worse than one that fails to convert.
  { name: "If/Then inline", rx: /^\s*If\b\s*(.+?)\s+Then\s+(.+?)\s*$/i,
    to: (m, ctx) => {
      if (/^\s*$/.test(m[2])) return null;

      // REFUSALS APPLY INSIDE THE BODY TOO. ctx.statement() converts but does not refuse, so a
      // construct we would never translate on its own line - Eval, Execute, reading Err - passed
      // through silently when it was written as `If cond Then GetDir = Eval(...)`. That is the
      // refusal mechanism failing exactly where the code is most surprising.
      // Err.Raise is deliberately NOT in this set. It does not READ Err state, it throws - and the
      // standalone Err.Raise rule below already emits `throw`. Lumping it in here meant
      // `If depth > 0 Then Err.Raise 5` became a TODO while the very same statement on its own
      // line converted cleanly, in the same file, three lines apart. Worse, the refusal silently
      // DROPPED a guard: malformed UTF-8 was then accepted instead of raising. Refusing where the
      // answer is obvious costs more trust than the refusal saves.
      const danger = /\b(Eval|Execute|ExecuteGlobal)\s*\(/i.test(m[2]) ? "Eval/Execute"
                   : /\bErr\s*\.\s*Clear\b/i.test(m[2]) ? "Err object"
                   : null;
      if (danger) {
        return { flag: danger === "Eval/Execute"
          ? "The body of this inline If builds and runs code from a string. Invoke-Expression is the "
          + "literal equivalent but is an injection risk - rewrite the logic explicitly."
          // No longer claims an On Error block exists. It often does not, and naming one that is
          // not there sends the reader looking for something they will never find.
          : "The body of this inline If READS VBScript's Err object, whose state persists after the "
          + "failing statement. PowerShell's $Error works differently - check what this branch "
          + "expects Err to still contain." };
      }

      // Colon-separated bodies all belong inside the If.
      // `If a Then If b Then c` nests; the inner If is a statement of the outer body and used to
      // reach the output as VBScript.
      const stmts = (s) => splitArgs(s.replace(/:/g, ",")).map((x) => x.trim()).filter(Boolean).map((x) => {
        const nested = x.match(/^If\b\s*(.+?)\s+Then\s+(.+)$/i);
        if (nested) return `if (${ctx.cond(nested[1])}) { ${stmts(nested[2]).join("; ")} }`;
        return ctx.statement(x);
      }).filter(Boolean);
      // `If c Then a Else b` on one line. Strings are placeholders here, so the split cannot hit
      // an Else inside a literal. Without this the Else rode along inside the Then body and came
      // out as `[string]("yes" Else WScript.Echo "no")`.
      // `If x Then y End If` on ONE line: VBScript accepts the trailing End If; it is not a statement.
      const branches = m[2].replace(/\s*:?\s*End\s+If\s*$/i, "").split(/\s+Else\s+/i);
      const head = `if (${ctx.cond(m[1])}) { ${stmts(branches[0]).join("; ")} }`;
      return branches.length > 1 ? `${head} else { ${stmts(branches.slice(1).join(" Else ")).join("; ")} }` : head;
    } },

  { name: "If/Then", rx: /^\s*If\b\s*(.+?)\s+Then\s*$/i,
    to: (m, ctx) => { ctx.open("if"); return `if (${ctx.cond(m[1])}) {`; } },

  { name: "ElseIf", rx: /^\s*ElseIf\b\s*(.+?)\s+Then\s*$/i,
    to: (m, ctx) => `} elseif (${ctx.cond(m[1])}) {` },

  { name: "Else", rx: /^\s*Else\s*$/i, to: () => "} else {" },
  { name: "End If", rx: /^\s*End\s+If\s*$/i, to: (m, ctx) => ctx.close("if") },

  { name: "For Each", rx: /^\s*For\s+Each\s+([A-Za-z_]\w*)\s+In\s+(.+?)\s*$/i,
    to: (m, ctx) => { ctx.open("loop"); return `foreach ($${m[1]} in ${ctx.expr(m[2])}) {`; } },

  { name: "For i = a To b", rx: /^\s*For\s+([A-Za-z_]\w*)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+?))?\s*$/i,
    to: (m, ctx) => {
      const step = m[4] ? ctx.expr(m[4]) : "1";
      const cmp = m[4] && m[4].trim().startsWith("-") ? "-ge" : "-le";
      const inc = step === "1" ? `$${m[1]}++` : `$${m[1]} += ${step}`;
      ctx.open("loop");
      return `for ($${m[1]} = ${ctx.expr(m[2])}; $${m[1]} ${cmp} ${ctx.expr(m[3])}; ${inc}) {`;
    } },

  { name: "Next", rx: /^\s*Next\s*$/i, to: (m, ctx) => ctx.close("loop") },

  { name: "Do While", rx: /^\s*Do\s+While\s+(.+?)\s*$/i,
    to: (m, ctx) => { ctx.open("loop"); return `while (${ctx.cond(m[1])}) {`; } },
  { name: "Do Until", rx: /^\s*Do\s+Until\s+(.+?)\s*$/i,
    to: (m, ctx) => { ctx.open("loop"); return `while (-not (${ctx.cond(m[1])})) {`; } },
  // A bare `Do` opens a do/while, and its own block kind so `Loop` knows which form to close with.
  { name: "Do (bare)", rx: /^\s*Do\s*$/i,
    to: (m, ctx) => { ctx.open("do"); return "do {"; } },

  { name: "Loop While/Until", rx: /^\s*Loop\s+(While|Until)\s+(.+?)\s*$/i,
    to: (m, ctx) => { ctx.close(blocksTop(ctx) === "do" ? "do" : "loop");
      return /until/i.test(m[1]) ? `} while (-not (${ctx.cond(m[2])}))` : `} while (${ctx.cond(m[2])})`; } },

  // `Do ... Loop` with no condition is an infinite loop in VBScript. PowerShell's do-block REQUIRES
  // a trailing while/until, so closing it with a bare `}` is a parse error - it has to become
  // `} while ($true)`, and the caller breaks out with Exit Do as before.
  { name: "Loop", rx: /^\s*Loop\s*$/i,
    to: (m, ctx) => (blocksTop(ctx) === "do"
      ? (ctx.close("do"), "} while ($true)")
      : ctx.close("loop")) },
  { name: "While", rx: /^\s*While\s+(.+?)\s*$/i,
    to: (m, ctx) => { ctx.open("loop"); return `while (${ctx.cond(m[1])}) {`; } },
  { name: "Wend", rx: /^\s*Wend\s*$/i, to: (m, ctx) => ctx.close("loop") },

  // Select Case is the one construct that CANNOT be done with independent line rules: PowerShell's
  // switch needs every clause to be `value { ... }`, so each Case has to CLOSE the previous clause
  // before opening its own. VBScript's Case just falls into the next one. This produced 118
  // MissingSwitchStatementClause errors before the block stack existed.
  { name: "Select Case", rx: /^\s*Select\s+Case\s+(.+?)\s*$/i,
    to: (m, ctx) => { ctx.open("switch"); return `switch (${ctx.expr(m[1])}) {`; } },
  { name: "Case Else", rx: /^\s*Case\s+Else\s*$/i,
    to: (m, ctx) => ctx.openCase("default") },
  // `Case "F02" strSrvLet = "I5835437"` - VBScript allows the statement on the Case line with no
  // separator at all. Without this the whole thing was taken as the case LABEL, emitting
  // `"F02" $strSrvLet = "I5835437" {`.
  { name: "Case with inline statement", rx: /^\s*Case\s+(.+?)\s+([A-Za-z_]\w*\s*=\s*[^=].*)$/i,
    to: (m, ctx) => `${ctx.openCase(splitArgs(m[1]).map((v) => ctx.expr(v)))} ${ctx.statement(m[2])} }`
      // The clause is opened and closed on one line, so the stack entry it pushed must come back off.
      + (ctx.closeCaseInline(), "") },

  // A Case label that CONTAINS a comparison is the `Select Case True` idiom - each clause is a
  // boolean expression rather than a value to match. Translated as a value it emitted
  // `$name = $MS_CVE_FILE {`, which is an assignment where a switch clause belongs.
  //
  // PowerShell expresses this as a condition block, and the expression needs condition-mode
  // translation so `=` becomes a comparison rather than staying an assignment.
  { name: "Case (comparison)", rx: /^\s*Case\s+(.+?)\s*$/i,
    to: (m, ctx) => {
      if (!/<=|>=|<>|[<>]|(?<![-<>!=+*/])=(?!=)/.test(m[1])) return SKIP;
      return ctx.openCase([`{ ${ctx.cond(m[1])} }`]);
    } },

  { name: "Case", rx: /^\s*Case\s+(.+?)\s*$/i,
    to: (m, ctx) => ctx.openCase(splitArgs(m[1]).map((v) => ctx.expr(v))) },
  { name: "End Select", rx: /^\s*End\s+Select\s*$/i, to: (m, ctx) => ctx.closeSwitch() },

  // A PowerShell class declares METHODS, not functions - `function X () {` inside a class block is
  // a parse error (MissingMethodParameterList), which was the single most common first failure.
  // VBScript's Class_Initialize is the constructor, and a PowerShell constructor is named for its
  // class, so the class name has to be carried on the block stack to emit it.
  { name: "Function", rx: /^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?Function\s+([A-Za-z_]\w*)\s*\(?\s*(.*?)\s*\)?\s*$/i,
    to: (m, ctx) => { ctx.openFunc(m[1], "func"); return ctx.method(m[1], m[2], "func"); } },
  { name: "Sub", rx: /^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?Sub\s+([A-Za-z_]\w*)\s*\(?\s*(.*?)\s*\)?\s*$/i,
    to: (m, ctx) => { ctx.openFunc(m[1], "sub"); return ctx.method(m[1], m[2], "sub"); } },
  { name: "End Function", rx: /^\s*End\s+Function\s*$/i, to: (m, ctx) => ctx.closeFunc() },
  { name: "End Sub", rx: /^\s*End\s+Sub\s*$/i, to: (m, ctx) => ctx.closeFunc() },

  // `Exit Function` must carry the return value out too, not just jump.
  { name: "Exit Function", rx: /^\s*Exit\s+Function\s*$/i, to: (m, ctx) => ctx.earlyReturn() },

  { name: "Exit Function/Sub", rx: /^\s*Exit\s+(Function|Sub)\s*$/i, to: () => "return" },
  { name: "Exit For/Do", rx: /^\s*Exit\s+(For|Do)\s*$/i, to: () => "break" },

  { name: "Class", rx: /^\s*Class\s+([A-Za-z_]\w*)\s*$/i,
    to: (m, ctx) => { ctx.openClass(m[1]); return `class ${m[1]} {`; } },
  { name: "End Class", rx: /^\s*End\s+Class\s*$/i, to: (m, ctx) => ctx.close("class") },

  { name: "With", rx: /^\s*With\s+(.+?)\s*$/i, to: null,
    flag: "PowerShell has no With block. Repeat the object on each line, or assign it to a short variable first." },
  { name: "End With", rx: /^\s*End\s+With\s*$/i, to: null,
    flag: "Closes a With block, which PowerShell does not have." },

  // `Call obj.Method(a)` - the name may be DOTTED. The old pattern captured only the leading
  // identifier, so `.Method(a` fell into the argument group and the trailing `)` was swallowed by
  // the optional `\)?`, emitting `objFSO .CreateFolder($a` with a space and no closing paren.
  { name: "Call method", rx: /^\s*Call\s+([A-Za-z_][\w.]*\.[A-Za-z_]\w*)\s*\((.*)\)\s*$/i,
    to: (m, ctx) => `${ctx.lhs(m[1])}(${splitArgs(m[2]).map((a) => ctx.expr(a)).join(", ")})` },

  { name: "Call", rx: /^\s*Call\s+([A-Za-z_]\w*)\s*\(?\s*(.*?)\s*\)?\s*$/i,
    to: (m, ctx) => `${m[1]} ${m[2] ? splitArgs(m[2]).map((a) => ctx.expr(a)).join(" ") : ""}`.trim() },

  { name: "Erase", rx: /^\s*Erase\s+([A-Za-z_]\w*)\s*$/i, to: (m) => `$${m[1]} = @()` },

  // Property Get/Let/Set is CONVERTED to a method rather than refused. Refusing it emitted no
  // opening brace while `End Property` still emitted a closing one, so every class with a property
  // came out brace-unbalanced and nothing after it parsed. A method that needs renaming is a far
  // smaller problem than a file that will not parse at all.
  { name: "Property Get", rx: /^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?Property\s+Get\s+([A-Za-z_]\w*)\s*\(?\s*(.*?)\s*\)?\s*$/i,
    // A REAL method, through the same path as Function: it gets its `$Name = $null` initialiser,
    // an [object] return type, and `return $Name` at End Property. Before this it was a bare
    // `Name() {` block whose body assigned `$Name` and returned nothing - every getter in the
    // corpus silently yielded $null.
    to: (m, ctx) => { ctx.openFunc(m[1], "func");
      return `# REVIEW: was Property Get ${m[1]}. It is a METHOD now: callers outside the class must write`
           + ` .${m[1]}() with parentheses - the property form hands back the method object, which is always truthy.\n`
           + ctx.method(m[1], m[2], "func"); } },

  { name: "Property Let/Set", rx: /^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?Property\s+(?:Let|Set)\s+([A-Za-z_]\w*)\s*\(?\s*(.*?)\s*\)?\s*$/i,
    to: (m, ctx) => { ctx.openFunc(m[1], "sub");
      return `# TODO: was Property Let/Set ${m[1]} - PowerShell has no direct property accessor\n`
           + `${ctx.inClass() ? "[void] " : "function "}Set${m[1]}(${paramList(m[2])}) {`; } },

  { name: "End Property", rx: /^\s*End\s+Property\s*$/i, to: (m, ctx) => ctx.closeFunc() },

  // A PowerShell class property is a DECLARATION, not an assignment - `[object] $x = $null`. And
  // inside methods those members must be reached through $this, which ctx.member() records here so
  // the expression pipeline can rewrite references later.
  // Multiple names per declaration: `Private objFSO, bVerbose` is as common as one per line, and
  // matching only the single-name form left the rest unregistered - so references to them inside
  // methods stayed bare and PowerShell rejected them.
  { name: "class member", rx: /^\s*(?:Private|Public|Dim)\s+([A-Za-z_]\w*(?:\s*,\s*[A-Za-z_]\w*)*)\s*$/i,
    to: (m, ctx) => {
      if (!ctx.inClass()) return null;
      return splitArgs(m[1]).map((n) => ctx.member(n.trim())).join("\n");
    } },

  // ── the honest refusals ────────────────────────────────────────────────────
  // TWO DIFFERENT INTENTS, and which one it is decides the right answer:
  //
  //   a) the script BRANCHES on Err.Number ("x if error 2, y if error 3"). Converting that
  //      automatically is dangerous - the control flow depends on it, so it needs real try/catch
  //      with the branches rewritten by hand.
  //   b) the script never looks at Err at all. The intent is simply "do not stop", overwhelmingly
  //      because it is a logon script and nobody will tolerate a dialog box at sign-in.
  //
  // Refusing BOTH was over-cautious: case (b) is most of them, and it has a safe mechanical answer.
  // Which case applies is decided per script in analyseErrorHandling(), not guessed at here.
  { name: "On Error Resume Next", rx: /^\s*On\s+Error\s+Resume\s+Next\s*$/i,
    to: (m, ctx) => ctx.onErrorResumeNext() },

  // Pairs with the statement above: if that one became a preference change, this one restores it.
  // Left as a refusal only when the script reads Err, where the region boundaries matter.
  { name: "On Error GoTo 0", rx: /^\s*On\s+Error\s+GoTo\s+0\s*$/i,
    to: (m, ctx) => ctx.onErrorGoToZero() },

  // Err.Raise has a real equivalent, unlike reading Err. The optional middle arguments are dropped
  // rather than guessed at: VBScript's signature is (number, source, description) and PowerShell's
  // throw takes a message, so only the description is carried across.
  { name: "Err.Raise", rx: /^\s*Err\.Raise\s+(.+)$/i,
    to: (m, ctx) => {
      const a = splitArgs(m[1]).map((x) => x.trim());
      const msg = a[2] || a[1] || a[0];
      return `throw ${msg ? ctx.expr(msg) : '"error"'}`;
    } },

  // Clears the SHIM, not $Error. `$Error.Clear()` emptied PowerShell's error list, which is not
  // what the script is doing: it is resetting the value that `If Err.Number <> 0` reads next.
  { name: "Err.Clear", rx: /^\s*Err\.Clear\s*$/i, to: () => "VbsClearErr" },

  // ASSIGNING to Err is how a script raises its own error without throwing - a SetError helper
  // that records a code and message for the caller to check. This was refused along with reading
  // Err, which predates the shim; the result was a SetError that set nothing, so every error path
  // reported success. On a Direct Connect script that means the console shows OK while no session
  // opened, which is the worst failure shape there is.
  { name: "Err assignment",
    rx: /^\s*Err\.(Number|Description|Source)\s*=\s*(.+)$/i,
    to: (m, ctx) => `$script:VbsErr.${m[1]} = ${ctx.expr(m[2])}` },

  // READS of Number/Description/Source are no longer refused: the shim ($script:VbsErr, set from
  // the exception in every catch the converter emits, cleared by Err.Clear and On Error) IS the
  // persisting state those reads expect, and 156 corpus lines - `errDesc = Err.Description` on
  // the line after the failing call - were TODOs for no reason. Raise and Clear in any position
  // their own rules do not cover are still refused.
  { name: "Err object", rx: /\bErr\.(Raise|Clear)\b/i, to: null,
    flag: "VBScript's Err object has no direct equivalent here. Inside a catch block use $_.Exception; " +
          "Err.Raise becomes throw." },

  { name: "Eval/Execute", rx: /^\s*(Eval|Execute|ExecuteGlobal)\b/i, to: null,
    flag: "Executes a string as code. Invoke-Expression is the literal equivalent but is a known " +
          "injection risk - rewrite the logic explicitly instead." },

  // The SEVERITY of a dialog depends entirely on where the script runs, so the reason is written
  // for the context the user selected rather than hedged for all of them.
  { name: "MsgBox", rx: /^\s*(?:\w+\s*=\s*)?MsgBox\b/i,
    to: (m, ctx) => ({ flag: ctx.dialogAdvice("MsgBox") }) },

  { name: "InputBox", rx: /^\s*\w+\s*=\s*InputBox\b/i,
    to: (m, ctx) => ({ flag: ctx.dialogAdvice("InputBox") }) },


  { name: "ADSI", rx: /["'](WinNT|LDAP):\/\//i, to: null,
    flag: "ADSI path. PowerShell can use [ADSI] directly, but the surrounding idiom usually reads " +
          "better rewritten with the ActiveDirectory module." },

  // Rnd is translated in translateBuiltins (Get-Random over 0..1); Randomize has nothing to do.
  { name: "Randomize", rx: /^\s*Randomize\b/i,
    to: () => "# Randomize: nothing to do, Get-Random is seeded already" },

  // ── generic call fallbacks, LAST ON PURPOSE ────────────────────────────────
  // These match almost anything shaped like a call, so they must sit AFTER every refusal above.
  // Placed earlier, "MsgBox ..." would be silently converted into a call instead of flagged - the
  // precise failure this tool exists to avoid.
  //
  // `obj.Method arg1, arg2`: a call with arguments and NO parentheses. Very common in real VBScript
  // (fout.WriteLine text, shell.RegWrite key, value, type) and missed by the bare-call fallback,
  // which only handled calls with no arguments at all.
  // The lookahead is (?!\s*=), not (?!=). With (?!=) a PROPERTY ASSIGNMENT written with padding -
  // `$obj.Global  =  True` - backtracked into a match: \s+ consumed one space, the next character
  // was another space rather than `=`, the lookahead passed, and the rule emitted
  // `$obj.Global(= $true)`. That parses as a method call and silently does the wrong thing.
  // The registry through WMI's StdRegProv, the way every 2000s inventory script reads it. Its
  // methods hand the value back through an OUT parameter, which a PowerShell method call cannot
  // do: `$reg.GetStringValue($h, $k, $v, $computerName)` runs, returns a status, and leaves
  // $computerName exactly as it was. Invoke-CimMethod returns an object carrying both.
  { name: "StdRegProv method",
    rx: /^\s*(?:(?:Set\s+)?([A-Za-z_]\w*)\s*=\s*)?([A-Za-z_][\w.]*)\.(GetStringValue|GetExpandedStringValue|GetMultiStringValue|GetDWORDValue|GetQWORDValue|GetBinaryValue|EnumKey|EnumValues|SetStringValue|SetExpandedStringValue|SetMultiStringValue|SetDWORDValue|SetQWORDValue|SetBinaryValue|CreateKey|DeleteKey|DeleteValue|CheckAccess)\s*(.*?)\s*$/i,
    to: (m, ctx) => {
      const method = m[3];
      // Strip an OUTER pair of parentheses only when it is one group. An optional `\)?` at the
      // end of the pattern ate the last parenthesis of `format( Array(...))` in the paren-less
      // statement form and left an unclosed call in the output.
      let raw = m[4].trim();
      if (raw.startsWith("(")) {
        let depth = 0, one = false;
        for (let i = 0; i < raw.length; i++) {
          if (raw[i] === "(") depth++;
          else if (raw[i] === ")" && --depth === 0) { one = i === raw.length - 1; break; }
        }
        if (one) raw = raw.slice(1, -1);
      }
      const a = splitArgs(raw).map((x) => x.trim());
      const lower = method.toLowerCase();
      // [args in] -> [named CIM arguments], [out params] -> result properties.
      const shapes = {
        getstringvalue: [["hDefKey", "sSubKeyName", "sValueName"], ["sValue"]],
        getexpandedstringvalue: [["hDefKey", "sSubKeyName", "sValueName"], ["sValue"]],
        getmultistringvalue: [["hDefKey", "sSubKeyName", "sValueName"], ["sValue"]],
        getdwordvalue: [["hDefKey", "sSubKeyName", "sValueName"], ["uValue"]],
        getqwordvalue: [["hDefKey", "sSubKeyName", "sValueName"], ["uValue"]],
        getbinaryvalue: [["hDefKey", "sSubKeyName", "sValueName"], ["uValue"]],
        enumkey: [["hDefKey", "sSubKeyName"], ["sNames"]],
        enumvalues: [["hDefKey", "sSubKeyName"], ["sNames", "Types"]],
        setstringvalue: [["hDefKey", "sSubKeyName", "sValueName", "sValue"], []],
        setexpandedstringvalue: [["hDefKey", "sSubKeyName", "sValueName", "sValue"], []],
        setmultistringvalue: [["hDefKey", "sSubKeyName", "sValueName", "sValue"], []],
        setdwordvalue: [["hDefKey", "sSubKeyName", "sValueName", "uValue"], []],
        setqwordvalue: [["hDefKey", "sSubKeyName", "sValueName", "uValue"], []],
        setbinaryvalue: [["hDefKey", "sSubKeyName", "sValueName", "uValue"], []],
        createkey: [["hDefKey", "sSubKeyName"], []],
        deletekey: [["hDefKey", "sSubKeyName"], []],
        deletevalue: [["hDefKey", "sSubKeyName", "sValueName"], []],
        checkaccess: [["hDefKey", "sSubKeyName", "uRequired"], ["bGranted"]],
      };
      const [ins, outs] = shapes[lower];
      if (a.length < ins.length || a.length > ins.length + outs.length) return SKIP;
      // The hive constant is a 32-bit Long in VBScript (&H80000002 is negative there); the CIM
      // argument is an unsigned 32-bit value.
      const args = ins.map((name, i) => `${name} = ${i === 0 ? `[uint32]([int64](${ctx.expr(a[0])}) -band 0xFFFFFFFF)` : ctx.expr(a[i])}`);
      const lines = [`$__reg = Invoke-CimMethod -Namespace root\\default -ClassName StdRegProv -MethodName ${method} -Arguments @{ ${args.join("; ")} }`];
      outs.forEach((prop, i) => { const out = a[ins.length + i]; if (out) lines.push(`${ctx.lhs(out)} = $__reg.${prop}`); });
      if (m[1]) lines.push(`${ctx.lhs(m[1])} = $__reg.ReturnValue`);
      return lines.join("\n");
    } },

  { name: "method call (no parens)",
    // The object may be INDEXED on the way: `Items(li).Add "x", 5`, `nd(1).setAttribute a, v`.
    rx: /^\s*([A-Za-z_][\w.]*(?:\([^()]*\))?(?:\.[A-Za-z_]\w*(?:\([^()]*\))?)*\.[A-Za-z_]\w*)\s+(?!\s*=)(.+?)\s*$/i,
    to: (m, ctx) => {
      if (/^\s*(If|For|Do|While|Select|Set|Call|Function|Sub|Const|Dim)\b/i.test(m[0])) return SKIP;
      // An omitted argument - `o.Logon "p", , False` - is `$null`, not an empty slot.
      return `${ctx.lhs(m[1])}(${splitArgs(m[2]).map((a) => (a.trim() ? ctx.expr(a) : "$null")).join(", ")})`;
    } },

  // Bare `SubName arg1, arg2` - calling a Sub without Call and without parentheses.
  { name: "sub call (no parens)",
    rx: /^\s*([A-Za-z_]\w*)\s+(?!=)([^=]+?)\s*$/i,
    to: (m, ctx) => {
      if (VB_KEYWORDS.has(m[1].toLowerCase())) return SKIP;
      // `Name ()` - an empty argument list after a space is no arguments at all, and `Name ()`
      // in PowerShell is a command given an empty expression.
      if (/^\(\s*\)$/.test(m[2].trim())) return ctx.isClassMethod(m[1]) ? `$this.${m[1]}()` : `(${m[1]})`;
      const args = splitArgs(m[2]).map((a) => (a.trim() ? ctx.expr(a) : "$null"));
      // A sibling METHOD needs .NET call syntax, and it has to be built HERE, where the commas
      // still exist. Reconstructing it later from the emitted string went wrong the obvious way:
      // the arguments are space-joined for a function call, so `SetIt n * 2` - one argument -
      // came back as three when split on whitespace.
      if (ctx.isClassMethod(m[1])) return `$this.${m[1]}(${args.join(", ")})`;
      return `${m[1]} ${args.join(" ")}`;
    } },
];

/** Sentinel: "this rule does not actually apply, keep looking". Distinct from null, which means
 *  "handled, and it deliberately produces no output" (a bare Dim). */
const SKIP = Symbol("skip");

/** The innermost open block, so a Loop rule can tell a `Do` from a `Do While`. */
const blocksTop = (ctx) => ctx.top();

/** Splits an argument list on commas that are not inside parentheses. */
function splitArgs(s) {
  const out = [];
  let depth = 0, cur = "";
  for (const ch of s) {
    // Brackets nest too: by the time a call is parenthesised its subscripts are `[i, j]`, and
    // splitting on that comma tore `arr[$i, $j]` into two arguments.
    if (ch === "(" || ch === "[") depth++;
    if (ch === ")" || ch === "]") depth--;
    if (ch === "," && depth === 0) { out.push(cur.trim()); cur = ""; continue; }
    cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}

function paramList(s) {
  if (!s || !s.trim()) return "";
  return s.split(",")
    .map((p) => p.trim().replace(/\b(ByVal|ByRef)\s+/gi, "").replace(/\s*\(\s*\)$/, ""))   // `arr()` = array param
    .filter(Boolean)
    .map((p) => "$" + p.replace(/^\$/, ""))
    .join(", ");
}

// ── the converter ────────────────────────────────────────────────────────────
/** @param opts.context One of SCRIPT_CONTEXTS. Omit to auto-detect. */
/**
 * Which Dim'd scalars actually need an `= $null` line.
 *
 * `Dim x` has to emit SOMETHING under Set-StrictMode, because VBScript's Dim creates the variable
 * as Empty and code leans on that. But emitting it for every declaration doubles the line count of
 * any real script - `$x = $null` immediately followed by `$x = 5` is noise on every page, and
 * noisy output is output nobody pastes into their editor.
 *
 * Only the accumulators need it: variables READ before they are first written. The test is the
 * first mention after the declaration - if that is an assignment whose right-hand side does not
 * mention the variable itself, the assignment establishes it and the $null is dead weight.
 *
 * The RHS check is the whole subtlety. `total = total + i` inside a loop is textually an
 * assignment FIRST, but it reads `total` on the right, so it does need initialising.
 *
 * Deliberately whole-file rather than per-scope, and deliberately biased toward keeping the init:
 * a missing initialiser is a runtime crash, while a redundant one is only untidy.
 */
/**
 * A VBScript intrinsic still being CALLED in the emitted PowerShell, or null.
 *
 * Looks for `Name(` and `Name.`, since those are the two shapes an untranslated intrinsic takes.
 * The lookbehind does the real work: `$x`, `.Method`, `-Parameter` and `[math]::Round` are all
 * legitimate and must not match, so `$`, `.`, `-`, `:` and word characters all disqualify.
 */
function leftoverIntrinsic(ps) {
  const t = ps
    .replace(/"(?:[^"`]|`.)*"/g, '""')      // string bodies can say anything
    .replace(/'[^']*'/g, "''")
    .replace(/^\s*#.*$/gm, "")              // our own TODO text names these on purpose
    .replace(/\s#.*$/gm, "");               // and a TRAILING comment - `$sh = $null  #WScript.Shell object` flagged 16 lines
  for (const m of t.matchAll(/(?<![\w$.:\-`])([A-Za-z_]\w*)\s*[(.]/g)) {
    if (VB_BUILTINS.has(m[1].toLowerCase())) return m[1];
  }
  return null;
}

/**
 * Every method name declared inside a `Class ... End Class`.
 *
 * Collected in a PRE-PASS rather than as the class body is processed, because VBScript happily
 * calls a method declared further down the class and a set built as we go would not contain it
 * yet - the first call in the file is usually to the last method in it.
 */
/**
 * Positional-to-named parameter mapping for the WMI methods scripts actually call.
 *
 * Deliberately small and only where the name is unambiguous from the WMI class documentation.
 * Anything not here still emits Arg1/Arg2 with a TODO: inventing a parameter name that binds to
 * nothing is worse than telling the reader to look it up, because Invoke-CimMethod fails loudly on
 * an unknown name but silently ignores the RIGHT value bound to the WRONG one.
 */
const WMI_METHOD_PARAMS = {
  changestartmode: ["StartMode"],
  create: ["CommandLine", "CurrentDirectory", "ProcessStartupInformation"],   // Win32_Process
  terminate: ["Reason"],
  rename: ["NewFileName"],
  setpowerstate: ["PowerState", "Time"],
};

/**
 * Flatten `(VbsConcat (VbsConcat a b) c)` into `(VbsConcat a b c)`.
 *
 * Folding is pairwise, so a five-part string came out four deep - correct and unreadable, and
 * unreadable output is what stops someone pasting it into their editor.
 *
 * This runs on the EMITTED PowerShell, not on the source expression. A previous attempt flattened
 * during folding and had to be reverted: it fought with the multi-argument builtins and took
 * `cbyte("&H" & mid(str, i + 1, 2))` apart. Here the nesting is already balanced and unambiguous,
 * so splicing the inner arguments into the outer call is a structural no-op - the helper is
 * variadic and coerces every argument to [string] either way.
 */
function flattenConcat(s) {
  const TAG = "(VbsConcat ";
  for (let guard = 0; guard < 60; guard++) {
    let changed = false;
    let i = s.indexOf(TAG);
    while (i !== -1) {
      let depth = 0, j = i;
      for (; j < s.length; j++) {
        if (s[j] === "(") depth++;
        else if (s[j] === ")") { depth--; if (depth === 0) break; }
      }
      if (j >= s.length) break;                       // unbalanced; leave it alone

      const inner = s.slice(i + TAG.length, j);
      const args = splitTopLevelSpaces(inner);
      // Only the nested-call arguments are spliced; everything else is carried through untouched.
      if (args.some((a) => a.startsWith(TAG) && a.endsWith(")"))) {
        const flat = args.flatMap((a) =>
          (a.startsWith(TAG) && a.endsWith(")") ? splitTopLevelSpaces(a.slice(TAG.length, -1)) : [a]));
        const rebuilt = `${TAG}${flat.join(" ")})`;
        s = s.slice(0, i) + rebuilt + s.slice(j + 1);
        changed = true;
        i = s.indexOf(TAG, i);                        // re-examine this call site
        continue;
      }
      i = s.indexOf(TAG, i + TAG.length);
    }
    if (!changed) break;
  }
  return s;
}

/** Split on whitespace that is not inside brackets or a string. */
function splitTopLevelSpaces(s) {
  const out = [];
  let depth = 0, inStr = false, cur = "";
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"' && s[i - 1] !== "`") inStr = !inStr;
    if (!inStr) {
      if (ch === "(" || ch === "[") depth++;
      else if (ch === ")" || ch === "]") depth--;
      else if (/\s/.test(ch) && depth === 0) { if (cur) out.push(cur); cur = ""; continue; }
    }
    cur += ch;
  }
  if (cur) out.push(cur);
  return out;
}

function findClassMethods(lines) {
  const out = new Set();
  let depth = 0;
  for (const l of lines) {
    const t = (l.text || "").trim();
    if (/^Class\s+[A-Za-z_]\w*/i.test(t)) { depth++; continue; }
    if (/^End\s+Class\b/i.test(t)) { depth = Math.max(0, depth - 1); continue; }
    if (!depth) continue;
    const m = t.match(/^(?:(?:Public|Private)\s+)?(?:Default\s+)?(?:Sub|Function|Property\s+Get)\s+([A-Za-z_]\w*)/i);
    if (m) out.add(m[1]);
  }
  return out;
}

/**
 * Class member names, declared at class-body level, for the WHOLE file up front. They used to be
 * registered as their declaration was emitted, so a class that lists its methods first and its
 * `Private` fields last (a common style in class libraries) had every member reference in those
 * methods left bare.
 */
function findClassMembers(lines) {
  const out = new Set();
  let depth = 0, inProc = false;
  for (const l of lines) {
    const t = (l.text || "").trim();
    let m;
    if (/^Class\s+[A-Za-z_]\w*/i.test(t)) { depth++; inProc = false; continue; }
    if (/^End\s+Class\b/i.test(t)) { depth = Math.max(0, depth - 1); inProc = false; continue; }
    if (!depth) continue;
    if (/^(?:(?:Public|Private)\s+)?(?:Default\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+/i.test(t)) { inProc = true; continue; }
    if (/^End\s+(?:Sub|Function|Property)\b/i.test(t)) { inProc = false; continue; }
    if (inProc) continue;
    if ((m = t.match(/^(?:Public|Private|Dim)\s+(.+)$/i)))
      for (const part of m[1].split(",")) {
        const mm = part.trim().match(/^([A-Za-z_]\w*)/);
        if (mm && !VB_KEYWORDS.has(mm[1].toLowerCase())) out.add(mm[1]);
      }
  }
  return out;
}

/** Property Get names per file. Read as bare words inside the class (`If Not HasStdStream Then`),
 *  which is a METHOD CALL in PowerShell and needs `$this.HasStdStream()`. */
function findClassGetters(lines) {
  const out = new Set();
  let depth = 0;
  for (const l of lines) {
    const t = (l.text || "").trim();
    if (/^Class\s+[A-Za-z_]\w*/i.test(t)) { depth++; continue; }
    if (/^End\s+Class\b/i.test(t)) { depth = Math.max(0, depth - 1); continue; }
    if (!depth) continue;
    const m = t.match(/^(?:(?:Public|Private)\s+)?(?:Default\s+)?Property\s+Get\s+([A-Za-z_]\w*)/i);
    if (m) out.add(m[1]);
  }
  return out;
}

const EMPTY_SET = new Set();
/**
 * Lower-cased names of the procedures the CURRENT script declares. A script's own `Sub Log(s)`
 * shadows the intrinsic Log, so every intrinsic translation consults this before firing. Set by
 * convertVbs for the file being converted; module-level because the translators are plain
 * functions that predate the symbol table.
 */
let USER_FUNCS = EMPTY_SET;
/** Variables PowerShell defines itself, which this converter emits and must never re-scope. */
const PS_AUTO_VARS = new Set(("this null true false args _ input psitem matches error lastexitcode host " +
  "psscriptroot pscommandpath myinvocation env psversiontable pid home pshome errorview " +
  "erroractionpreference progresspreference " +
  // scope prefixes: `$script:VbsErr` was matched on the word `script` and became `$script:script:`
  "script global local private using function variable").split(" "));

/** Class name -> the name of its `Default` method, for the `(New X)(args)` construction idiom. */
function findDefaultMethods(lines) {
  const out = new Map();
  let cls = null;
  for (const l of lines) {
    const t = (l.text || "").trim();
    let m;
    if ((m = t.match(/^Class\s+([A-Za-z_]\w*)/i))) { cls = m[1]; continue; }
    if (/^End\s+Class\b/i.test(t)) { cls = null; continue; }
    if (cls && (m = t.match(/^(?:(?:Public|Private)\s+)?Default\s+(?:Sub|Function|Property\s+Get)\s+([A-Za-z_]\w*)/i)))
      out.set(cls, m[1]);
  }
  return out;
}

/**
 * Which names are SCRIPT-LEVEL, and which are local to each procedure.
 *
 * `globals` is every name declared or assigned outside any Sub/Function/Property/Class body.
 * `locals` maps a procedure name to its parameters plus anything it Dims, which is exactly what
 * shadows a script-level name inside it. Class-level declarations are members, not globals, and
 * are left to the class handling. Constants are excluded: they are only ever read, and a read
 * resolves through PowerShell's dynamic scoping without help.
 */
export function findProcedureScopes(lines) {
  const globals = new Set();
  const locals = new Map();
  const assigned = new Map();
  const params = new Map();                  // parameters only, a subset of locals
  /** Source line that first made each name a global - for working out WHY a name got `$script:`. */
  const globalAt = new Map();
  let current = null, inClass = 0, lineNo = 0;
  const names = (s) => s.split(",").map((p) => p.trim().match(/^([A-Za-z_]\w*)/)).filter(Boolean).map((m) => m[1]);
  const add = (set, list) => list.forEach((n) => {
    if (VB_KEYWORDS.has(n.toLowerCase())) return;
    if (set === globals && !set.has(n)) globalAt.set(n, lineNo);
    set.add(n);
  });
  for (const { text, n } of lines) {
    lineNo = n;
    if (/^\s*(?:'|Rem\b)/i.test(text)) continue;
    let { code } = protectStrings(text);
    const q = code.indexOf("'");
    if (q !== -1) code = code.slice(0, q);
    let m;
    if (/^\s*Class\s+[A-Za-z_]\w*/i.test(code)) { inClass++; continue; }
    if (/^\s*End\s+Class\b/i.test(code)) { inClass = Math.max(0, inClass - 1); continue; }
    if ((m = code.match(/^\s*(?:(?:Public|Private)\s+)?(?:Default\s+)?(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\s*\(?([^)]*)\)?/i))) {
      current = m[1];
      if (!locals.has(current)) locals.set(current, new Set());
      if (!assigned.has(current)) assigned.set(current, new Set());
      if (!params.has(current)) params.set(current, new Set());
      if (m[2]) {
        add(locals.get(current), names(m[2].replace(/\b(ByVal|ByRef)\b/gi, "")));
        add(params.get(current), names(m[2].replace(/\b(ByVal|ByRef)\b/gi, "")));
      }
      continue;
    }
    if (/^\s*End\s+(Function|Sub|Property)\b/i.test(code)) { current = null; continue; }

    const target = current ? locals.get(current) : inClass ? null : globals;
    if (!target) continue;
    if ((m = code.match(/^\s*(?:Public|Private)?\s*Dim\s+(.+)$/i)))
      add(target, names(m[1].replace(/\([^)]*\)/g, "")));
    // ReDim is NOT a declaration when the name already exists in an enclosing scope: a method
    // that does `ReDim Preserve arr(n)` on a class member resizes the MEMBER. Counting it as a
    // local made thisify skip the name and the method wrote to a variable that does not exist.
    // Constants too. A plain function reads them through dynamic scoping, but a CLASS METHOD sees
    // no script-level variable at all without `$script:` - "Variable is not assigned in the method".
    if ((m = code.match(/^\s*(?:Public|Private)?\s*Const\s+(.+)$/i)))
      m[1].split(",").forEach((part) => add(target, names(part.split("=")[0])));
    // Only Dim makes a local. ReDim resizes whatever the name already is. An ASSIGNMENT inside a
    // procedure is recorded separately: without Option Explicit it creates a procedure-local, which
    // is what tells an unknown name that is written here from one that is only read here.
    if (!current && (m = code.match(/^\s*ReDim\s+(?:Preserve\s+)?(.+)$/i)))
      add(target, names(m[1].replace(/\([^)]*\)/g, "")));
    const sink = current ? assigned.get(current) : target;
    if ((m = code.match(/^\s*Set\s+([A-Za-z_]\w*)\s*=/i))) add(sink, [m[1]]);
    if ((m = code.match(/^\s*([A-Za-z_]\w*)\s*(?:\([^)]*\))?\s*=\s*[^=]/))) add(sink, [m[1]]);
    if ((m = code.match(/^\s*For\s+(?:Each\s+)?([A-Za-z_]\w*)/i))) add(sink, [m[1]]);
    if ((m = code.match(/\bThen\s+(?:Set\s+)?([A-Za-z_]\w*)\s*=\s*[^=]/i))) add(sink, [m[1]]);
  }
  return { globals, locals, assigned, params, globalAt };
}

/** Line numbers inside ANY `On Error Resume Next` region, checked or not, up to `On Error GoTo 0`
 *  or the end of the enclosing procedure. */
function findAllOernLines(lines) {
  const out = new Set();
  let on = false;
  for (const { text, n } of lines) {
    const t = text.trim();
    if (/^On\s+Error\s+Resume\s+Next\b/i.test(t)) { on = true; continue; }
    if (/^On\s+Error\s+GoTo\s+0\b/i.test(t) || /^End\s+(Function|Sub|Property)\b/i.test(t)) { on = false; continue; }
    if (on) out.add(n);
  }
  return out;
}

function findVarsReadBeforeAssign(lines, oernLines) {
  const safe = new Set();
  const text = lines
    .map((l) => (typeof l === "string" ? l : l.text || ""))
    // Strings and comments cannot declare or assign, and an identifier inside either would
    // otherwise count as a "read" and keep the initialiser forever.
    .map((l) => l.replace(/"(?:[^"]|"")*"/g, '""').replace(/'.*$/, ""))
    .join("\n");

  const declared = new Set();
  for (const m of text.matchAll(/^\s*(?:Public|Private)?\s*Dim\s+(.+)$/gim)) {
    for (const d of m[1].split(",")) {
      const n = d.trim().match(/^([A-Za-z_]\w*)$/);
      if (n) declared.add(n[1]);
    }
  }

  for (const name of declared) {
    const mention = new RegExp(`(?<![\\w.])${name}\\b`, "gi");
    let first = null;
    for (const m of text.matchAll(mention)) {
      // Skip the declaration itself.
      const lineStart = text.lastIndexOf("\n", m.index) + 1;
      const line = text.slice(lineStart, text.indexOf("\n", m.index) === -1 ? undefined : text.indexOf("\n", m.index));
      if (/^\s*(?:Public|Private)?\s*Dim\s/i.test(line)) continue;
      // Source line number, recovered from the offset: the joined text has one entry per logical
      // line, so counting newlines before the match indexes straight back into `lines`.
      const lineIdx = text.slice(0, m.index).split("\n").length - 1;
      first = { line, n: lines[lineIdx] ? lines[lineIdx].n : -1 };
      break;
    }
    if (!first) { safe.add(name); continue; }          // never used at all

    // A For loop establishes its own counter, and For Each its own element, so neither needs an
    // initialiser. These are common enough that missing them left the noise in most scripts.
    if (new RegExp(`^\\s*For\\s+Each\\s+${name}\\s+In\\s`, "i").test(first.line)
     || new RegExp(`^\\s*For\\s+${name}\\s*=`, "i").test(first.line)) { safe.add(name); continue; }

    // An assignment inside an On Error region can FAIL, leaving the variable unset - and the next
    // line is usually the Err check that reads it, which then throws under Set-StrictMode before
    // the check can run. VBScript would have had Empty there and taken the fallback path. So a
    // variable whose first assignment sits in an On Error region keeps its initialiser.
    if (oernLines && oernLines.has(first.n)) continue;

    const assign = first.line.match(new RegExp(`^\\s*${name}\\s*=\\s*(.+)$`, "i"));
    if (assign && !new RegExp(`(?<![\\w.])${name}\\b`, "i").test(assign[1])) safe.add(name);
  }
  return { skip: (n) => safe.has(n) };
}

export function convertVbs(src, opts = {}) {
  const logical = expandWith(canonicaliseIdentifiers(joinContinuations(src)));
  const { vars, funcs, loopVars, classes, unknownCalls } = discoverNames(logical);
  const usesGetRef = /\bGetRef\s*\(/i.test(src);
  /** Names the script uses AS ARRAYS somewhere, lower-cased - never a procedure reference. */
  const arrayish = new Set([
    ...[...src.matchAll(/\b(?:UBound|LBound|IsArray|Join|Filter|Erase)\s*\(\s*([A-Za-z_]\w*)/gi)].map((m) => m[1]),
    ...[...src.matchAll(/\bFor\s+Each\s+\w+\s+In\s+([A-Za-z_]\w*)/gi)].map((m) => m[1]),
    ...[...src.matchAll(/\bReDim\s+(?:Preserve\s+)?([A-Za-z_]\w*)/gi)].map((m) => m[1]),
  ].map((n) => n.toLowerCase()));
  USER_FUNCS = new Set([...funcs].map((f) => f.toLowerCase()));
  /** What parenthesizeCalls treats as callable: this file's procedures plus calls into includes. */
  const callables = new Set([...funcs, ...unknownCalls]);
  const scopes = findProcedureScopes(logical);
  const defaultMethods = findDefaultMethods(logical);
  // EVERY On Error region, not only the ones that branch on Err. A failed assignment inside any
  // of them leaves the variable unset, and the read on the next line throws under strict mode -
  // `Dim testResult: testResult = sh.RegRead(key)` under a bare On Error Resume Next did exactly
  // that when the key was missing.
  const needsInit = findVarsReadBeforeAssign(logical, findAllOernLines(logical));
  const errInfo = analyseErrorHandling(logical);
  const oern = findCheckedOernRegions(logical);
  // Set per line so the rules, which do not receive the line number, can ask which region they are
  // in. Only the OERN rule needs it, and only to tell "do not stop" from "branches on Err".
  let curLine = -1;
  const byRefMut = findByRefMutations(logical);
  const returningFuncs = findReturningFuncs(logical);
  // Caller-supplied context wins over the guess: the user knows how they deploy it and we do not.
  const ctxKind = opts.context && SCRIPT_CONTEXTS[opts.context] && opts.context !== "unknown"
    ? opts.context
    : detectContext(src);
  const out = [];

  /**
   * THE BLOCK STACK. Without it, every `End If` emits a `}` whether or not an `if` was ever opened,
   * and every unterminated construct leaves the file unbalanced - which is what produced 507
   * MissingEndCurlyBrace errors and put the parse-clean rate at ~36%. Emitting a brace is now a
   * consequence of what is actually open, not of which keyword the line happened to contain.
   */
  const blocks = [];
  const classNames = [];
  /** One entry per open Function/Sub: its name, and whether the body assigned to that name. */
  const funcStack = [];
  /** Class property names, so references inside methods can be rewritten to $this.name - PowerShell
   *  rejects a bare member reference inside a class with "MissingThis". */
  const classMembers = new Set([...findClassMembers(logical)].filter((n) => !loopVars.has(n)));
  /** Method names declared in any class, so sibling calls become `$this.Name(...)`. */
  const classMethods = findClassMethods(logical);
  const classGetters = findClassGetters(logical);
  /** Inside a class a method name is a call, never the same-named script variable. */
  const varsInClass = new Set([...vars].filter((v) => !classMethods.has(v)));

  /** Closes the innermost block of `kind`. A close keyword with nothing open is a no-op, NOT a
   *  stray `}` - one unbalanced brace desynchronises every block after it in the file. */
  const ctxClose = (kind) => {
    const i = blocks.lastIndexOf(kind);
    if (i === -1) return "";
    const extra = blocks.length - 1 - i;       // inner blocks VBScript left implicit
    blocks.length = i;
    return "}".repeat(extra + 1).split("").join("\n");
  };

  const mkCtx = (strings) => {
    const thisify = (c) => {
      // Never rewrite a property DECLARATION - `[object] $x = $null` is the declaration form, and
      // `[object] $this.x = $null` is a syntax error. Applying thisify centrally without this
      // exclusion broke more files than it fixed.
      // The `m` flag matters: `Dim a, b, c` inside a class emits THREE declarations joined by
      // newlines, and without it only the first line was recognised as a declaration - the rest
      // were rewritten to `[object] $this.b = $null`, which is a syntax error.
      if (/^\s*\[object\]\s*\$/m.test(c)) return c;
      if (blocks.lastIndexOf("class") === -1) return c;
      // `Me` is the instance.
      c = c.replace(/(?<![\w$.])Me\b/gi, "$this");
      // Only for names declared as members - a local variable in a method must NOT become $this.x,
      // and neither must a PARAMETER that shadows a member: `Function Calc(deadline, mins)` inside
      // a class with a `deadline` member emitted `Calc($this.deadline, $mins) {`, which does not
      // parse, and every use in the body then read the member instead of the argument.
      const f = funcStack[funcStack.length - 1];
      const locals = f ? scopes.locals.get(f.name) || EMPTY_SET : EMPTY_SET;
      if (classMembers.size) {
        const names = [...classMembers].filter((n) => !locals.has(n)).sort((a, b) => b.length - a.length).join("|");
        if (names) c = c.replace(new RegExp(`\\$(${names})\\b`, "g"), "$this.$1");
      }
      // A bare Property Get read is a method CALL. Inside the getter itself the same name is its
      // return variable and stays as it is; so does anything being assigned or already called.
      // Not on a DECLARATION line (`[object] Init($base, $filter)` has a parameter called filter
      // that shadows the getter), and not for a name that is a parameter or local of this method.
      if (classGetters.size && !/^\s*(?:\[\w+\]\s*)?[A-Za-z_]\w*\s*\([^)]*\)\s*\{\s*$/.test(c)) {
        const g = [...classGetters].filter((n) => !locals.has(n) && (!f || n !== f.name)).sort((a, b) => b.length - a.length).join("|");
        // `return $HasV` is emitted by closeFunc AFTER the getter has been popped, so the
        // own-name exclusion above cannot see it; rewriting it made the getter call itself forever.
        // `DSN()` - a getter called WITH parentheses - is the same call.
        if (g) c = c.replace(new RegExp(`(?<!\\breturn\\s+)\\$(${g})\\b(?:\\s*\\(\\s*\\))?(?!\\s*[(=])`, "g"), "$this.$1()");
      }
      return c;
    };

    /**
     * `New X` inside an EXPRESSION (the `Set x = New X` statement has its own rule). Only for
     * classes this script declares; RegExp and friends stay with their own handling. The
     * `(New X)(args)` idiom calls the class's Default method on the fresh instance - some
     * libraries construct every helper this way - so when the class declares one, the call is
     * routed to it by name.
     */
    const newExpr = (c) => {
      if (!classes.size) return c;
      const names = [...classes].sort((a, b) => b.length - a.length).join("|");
      // `new` is a keyword, so its case is free; the class name is canonical by now.
      let out = c.replace(new RegExp(`(?<![\\w$.])[Nn][Ee][Ww]\\s+(${names})\\b`, "g"), (m, n) => `[${n}]::new()`);
      out = out.replace(new RegExp(`\\(\\s*\\[(${names})\\]::new\\(\\)\\s*\\)\\s*\\(`, "g"),
        (m, n) => defaultMethods.has(n) ? `([${n}]::new()).${defaultMethods.get(n)}(` : m);
      return out;
    };

    /**
     * Rewrite a call to a SIBLING METHOD into `$this.Name(a, b)`.
     *
     * VBScript calls a method of the same class by bare name: `AddFilter "Name", strName`. In a
     * PowerShell class that is not a method call at all - it is a command lookup for a function
     * called AddFilter, which does not exist, so the first such call ends the script. They also
     * need real method syntax: parentheses and commas, not space-separated arguments.
     *
     * Only names declared as methods of the class currently open are touched, so a call to a
     * genuine script-level function from inside a method still works.
     */
    const thisifyCall = (c) => {
      if (!classMethods.size || blocks.lastIndexOf("class") === -1) return c;
      // Never touch the DECLARATION. `[void] AddFilter($a) {` starts with its return type, and
      // rewriting that to `[void] $this.AddFilter($a) {` is a syntax error.
      if (/^\s*\[\w+\]\s/.test(c)) return c;
      const names = [...classMethods].sort((a, b) => b.length - a.length).join("|");

      // `Name(a, b)` - already parenthesised by the time this runs.
      // `(?<![A-Za-z]-)`: a cmdlet is not a method. A class with a Sleep method turned the
      // emitted `Start-Sleep` into `Start-$this.Sleep()`.
      let out = c.replace(new RegExp(`(?<![\\w$.])(?<![A-Za-z]-)(${names})\\s*\\(`, "gi"),
        (mm, n) => `$this.${n}(`);
      // `Name a, b` - the paren-less statement form, which parenthesizeCalls leaves alone for a
      // name it does not know is callable.
      out = out.replace(new RegExp(`^(\\s*)(${names})\\s+(?!\\s*[-=])(\\S.*)$`, "i"),
        (mm, ind, n, rest) => `${ind}$this.${n}(${splitArgs(rest).map((a) => a.trim()).join(", ")})`);
      return out;
    };

    // THE RETURN VARIABLE. Inside `Function F`, a bare `F` that is not followed by `(` is the
    // function's own return variable being READ - `If Right(GetDir, 1) <> "\"` - and it needs a
    // sigil. `F(` is a recursive call and must stay one: sigilising it too would have indexBrackets
    // turn `Fact(n - 1)` into `$Fact[n - 1]`. Sub bodies are excluded (a Sub has no return
    // variable, so a bare `F` there is a call).
    //
    // A first attempt on 2026-09-21 was reverted at 340 -> 331 parse-clean; it was scoped by name
    // rather than by the function being OPEN, and the case-insensitive spellings had not yet been
    // unified. Both are true now, and the number is measured again below rather than argued.
    const returnVar = (c) => {
      const f = funcStack[funcStack.length - 1];
      if (!f || f.kind !== "func") return c;
      // The lookahead is `(` with NO whitespace: by this point translateOperators has already
      // space-separated folded arguments, so `GetDir & Chr(92)` reads `GetDir ([char]...)` and a
      // `\s*` here mistook the return variable for a call taking the char as its argument.
      return c.replace(new RegExp(`(?<![\\w$.])${f.name}\\b(?!\\()`, "g"), () => "$" + f.name);
    };
    // SCRIPT-LEVEL VARIABLES WRITTEN INSIDE A PROCEDURE. VBScript has one flat scope for them: a
    // Sub that does `depth = depth + 1` changes the script's `depth`. PowerShell's assignment
    // inside a function creates a LOCAL, so the same line silently updates a copy and the script's
    // value never moves - `If depth < 3 Then Dig` then recursed forever. Any name that is
    // script-level, and not shadowed by a parameter or a local Dim, is addressed as `$script:name`
    // inside every procedure - reads too, so a body is consistent with itself.
    //
    // Loop counters are left alone: `foreach ($script:x in ...)` is not a form PowerShell accepts,
    // and a counter is not read after its loop. Class members are left to thisify.
    const scopeGlobals = (c) => {
      const f = funcStack[funcStack.length - 1];
      if (!f) return c;
      const locals = scopes.locals.get(f.name) || EMPTY_SET;
      const assigned = scopes.assigned.get(f.name) || EMPTY_SET;
      const inClass = blocks.lastIndexOf("class") !== -1;
      return c.replace(/(?<![\w$.:])\$([A-Za-z_]\w*)\b/g, (m, name) => {
        // PowerShell's own variables are never script-level names of the VBScript. `$this` in
        // particular: rewriting it to `$script:this` parsed cleanly and broke every class.
        if (PS_AUTO_VARS.has(name.toLowerCase())) return m;
        if (locals.has(name) || loopVars.has(name) || funcs.has(name) || name === f.name) return m;
        if (inClass && (classMembers.has(name) || classGetters.has(name))) return m;
        if (scopes.globals.has(name)) return `$script:${name}`;
        // Inside a CLASS METHOD a name that is neither parameter, local, member nor assigned here
        // can only be script-level - a constant from an included file, or one this file assigns
        // somewhere the global scan does not see. PowerShell's class parser rejects it bare
        // ("Variable is not assigned in the method"); `$script:` resolves it at run time.
        if (inClass && !assigned.has(name)) return `$script:${name}`;
        return m;
      });
    };
    // `Name()` on a name this file does not declare (an included file's function, usually): in
    // PowerShell that is command Name with an EMPTY expression as its argument, a parse error.
    // `(Name)` is the call. Members, statics and cmdlets are excluded by the lookbehind.
    const emptyCall = (c) => c.replace(/(?<![\w$.:\]\-])([A-Za-z_]\w*)\s*\(\s*\)/g, (m, n) =>
      /^(if|while|switch|foreach|for|function|param|return|until|do|elseif|new|throw)$/i.test(n) ? m : `(${n})`);
    // Bare `Date`, `Time`, `Timer` that survived sigilize are the intrinsics: a variable of that
    // name would carry a `$` by now. `DatePart("m", Date)` reached the output with a bare word.
    const lateIntrinsics = (c) => c
      .replace(/(?<![\w$.:\-])Date\b(?!\s*[(.=])/gi, "(Get-Date).Date")
      .replace(/(?<![\w$.:\-])Time\b(?!\s*[(.=])/gi, "(Get-Date)")
      .replace(/(?<![\w$.:\-])Timer\b(?!\s*[(.=])/gi, "((Get-Date).TimeOfDay.TotalSeconds)");
    const fin = (c) =>
      getRef(restoreStringsPs(lateIntrinsics(emptyCall(scopeGlobals(thisifyCall(thisify(fillEmptyArgs(parenthesizeCalls(sigilize(returnVar(newExpr(translateBuiltins(c))),
        blocks.lastIndexOf("class") !== -1 ? varsInClass : vars), callables,
        blocks.lastIndexOf("class") !== -1 ? classMethods : null))))))), strings));
    /** `GetRef("name")` is the procedure as a value: `${function:name}`. After the strings are
     *  back, because the name is inside one. */
    function getRef(s) { return s.replace(/\bGetRef\(\s*"([A-Za-z_]\w*)"\s*\)/gi, "$${function:$1}"); }
    return {
      expr: (s) => fin(translateOperators(s, { isCondition: false })),
      cond: (s) => fin(translateOperators(s, { isCondition: true })),
      lhs: (s) => fin(s),
      stmt: (s) => fin(translateOperators(s, { isCondition: false })),
      raw: (s) => restoreStringsPs(s, strings),

      /**
       * Converts ONE statement for an inline context (the body of `If ... Then x`).
       *
       * Deliberately a small, separate path rather than a recursive trip through the rule table:
       * the table's rules push and pop the block stack, and an inline body opens no block, so
       * reusing them would leave the stack corrupted for the rest of the file.
       */
      statement: (s) => {
        // `Call Foo(a)` stripped up front. The RULES table has a Call rule and this parallel table
        // did not, so `If x Then Call fRaiseError(...)` emitted a call to a command named "Call".
        // It sat on an error path, which is why it survived three reviews.
        let t = s.trim();
        const callForm = t.match(/^Call\s+(.+)$/i);
        if (callForm) t = callForm[1].replace(/^\((.*)\)$/s, "$1").trim();
        if (!t) return "";
        let mm;
        if ((mm = t.match(/^Exit\s+(For|Do)$/i))) return "break";
        // Same decision as earlyReturn below: a guard clause `If n = 0 Then Exit Function` in a
        // function that assigns its name later must still return `$Name` - inside a class the
        // method is [object] and a bare `return` does not parse.
        if ((mm = t.match(/^Exit\s+(Function|Sub)$/i))) {
          const f = funcStack[funcStack.length - 1];
          return f && f.kind === "func" && returningFuncs.has(f.name) ? `return $${f.name}` : "return";
        }
        // CONVERGENCE NOTE. This is a SECOND statement table, parallel to RULES, and every fix
        // has to be made in both or the inline-If form silently diverges from the standalone one.
        // Both bugs below were exactly that: the same statement converted correctly on its own
        // line and wrongly inside `If ... Then`, three lines apart in the same file.
        //
        // Quit: matching only digits left `If failed Then WScript.Quit EXIT_FAILURE` calling a
        // WScript host that is not there - on the statement whose whole job is the exit code.
        if ((mm = t.match(/^WScript\.Quit\s*(?:\(\s*(.*?)\s*\)|\s+(.+?))?$/i))) {
          const a = (mm[1] ?? mm[2] ?? "").trim();
          return `exit ${a ? fin(translateOperators(a, { isCondition: false })) : 0}`;
        }
        // Raise: refused here as an "Err state read" while the standalone rule emitted `throw`.
        // It does not read Err, it throws. The refusal DROPPED a guard - malformed UTF-8 was
        // accepted instead of raising.
        if ((mm = t.match(/^Err\.Raise\s*\(?\s*(.+?)\s*\)?$/i))) {
          const a = splitArgs(mm[1]).map((x) => x.trim());
          const msg = a[2] || a[1] || a[0];
          return `throw ${msg ? fin(translateOperators(msg, { isCondition: false })) : '"error"'}`;
        }
        if ((mm = t.match(/^WScript\.Echo\s+(.+)$/i))) {
          const e = fin(translateOperators(mm[1], { isCondition: false }));
          // ALWAYS parenthesised. A cast must be followed by a parseable expression, and an
      // untranslated bare word like `Err.Description` is not one - `[string]Err.Description` is a
      // hard parse error where the old `Write-Output Err.Description` at least parsed as a command
      // with an argument. Wrapping keeps a line we could not fully convert from taking the file
      // down with it.
      return `[Console]::WriteLine([string](${e}))`;
        }
        if ((mm = t.match(/^Set\s+(.+?)\s*=\s*(.+)$/i)))
          return `${fin(mm[1])} = ${fin(translateOperators(mm[2], { isCondition: false }))}`;
        // `obj.Method arg` with no parentheses - the case that made this helper necessary.
        if ((mm = t.match(/^([A-Za-z_][\w.]*\.[A-Za-z_]\w*)\s+(?!\s*=)(.+)$/i)))
          return `${fin(mm[1])}(${splitArgs(mm[2]).map((a) => fin(translateOperators(a, { isCondition: false }))).join(", ")})`;
        // A bare Sub call: `include path`, `Raise 1, "msg"`. The leading name must NOT be sigilized -
        // running the whole line through fin() turned `include $path` into `$include $path`, which
        // PowerShell reads as two expressions and rejects. Subs from an included file are never in
        // the function registry, so the catch-all had classified them as variables.
        if ((mm = t.match(/^([A-Za-z_]\w*)\s+(?!\s*[-=])(.+)$/)) && !VB_KEYWORDS.has(mm[1].toLowerCase()))
          return `${mm[1]} ${splitArgs(mm[2]).map((a) => fin(translateOperators(a, { isCondition: false }))).join(" ")}`;

        if (/^[A-Za-z_][\w.()[\]]*\s*=\s*[^=]/.test(t)) {
          const i = t.indexOf("=");
          const lhs = t.slice(0, i).trim();
          // Return-by-function-name works inline too: `Case 0  GetErrorMessage = "Success"`.
          // Only the main assignment path knew about it, so function names stayed bare here and
          // PowerShell saw an assignment to an undeclared identifier.
          const openF = funcStack[funcStack.length - 1];
          if (funcs.has(lhs) || (openF && openF.kind === "func" && openF.name === lhs)) {
            const f = openF;
            if (f && f.name === lhs) f.assigned = true;
            return `$${lhs} = ${fin(translateOperators(t.slice(i + 1).trim(), { isCondition: false }))}`;
          }
          return `${fin(lhs)} = ${fin(translateOperators(t.slice(i + 1).trim(), { isCondition: false }))}`;
        }
        let call = fin(translateOperators(t, { isCondition: false }));
        if (/\.\w+$/.test(call)) call += "()";
        return call;
      },

      open: (kind) => { blocks.push(kind); },

      /** Remembers the function being opened, so `End Function` can emit its return. */
      openFunc: (name, kind) => { blocks.push("func"); funcStack.push({ name, kind, assigned: false }); },

      /** Called when the body assigns to the function's own name - that IS the return in VBScript. */
      noteReturnAssign: (name) => {
        const f = funcStack[funcStack.length - 1];
        if (f && f.name.toLowerCase() === name.toLowerCase()) f.assigned = true;
      },

      /**
       * Closes a function, emitting `return $Name` when the body assigned to the function name.
       *
       * Without this the value was computed into a local and thrown away, so every such function
       * returned nothing - and because PowerShell emits UNCAPTURED expressions instead, the caller
       * often got some unrelated intermediate value rather than an obvious error.
       */
      /** `Exit Function` returns the value assigned so far, not nothing. */
      earlyReturn: () => {
        const f = funcStack[funcStack.length - 1];
        // `returningFuncs`, not `f.assigned`: assigned only knows about lines ABOVE this one, and
        // a guard clause exits before the function assigns anything. In a class that method is
        // declared [object] (returningFuncs decides that too), where a bare `return` is "Invalid
        // return statement within non-void method". `$Name` is always declared for such a function.
        return f && f.kind === "func" && returningFuncs.has(f.name) ? `return $${f.name}` : "return";
      },

      closeFunc: () => {
        const f = funcStack.pop();
        const close = ctxClose("func");
        return f && f.assigned ? `return $${f.name}\n${close}` : close;
      },

      openClass: (name) => { classNames.push(name); blocks.push("class"); },

      skipInit: (n) => needsInit.skip(n),
      // The open procedure counts too: a Property Get is not in `funcs`, and its one-liner form
      // `Property Get StdOut : Set StdOut = stdout_ : End Property` reaches the Set rule.
      isFunc: (n) => funcs.has(n)
        || (funcStack.length > 0 && funcStack[funcStack.length - 1].kind === "func" && funcStack[funcStack.length - 1].name === n),
      isClassMethod: (n) => classMethods.has(n) && blocks.lastIndexOf("class") !== -1,
      isClass: (n) => classes.has(n),
      usesAdsi: () => /(WinNT|LDAP|GC|IIS):\/\//i.test(src),
      inClass: () => blocks.lastIndexOf("class") !== -1,

      /** Declares a class property and remembers it, so references inside methods become $this.x. */
      member: (name) => {
        // A Dim inside a METHOD is a LOCAL, not a property. PowerShell classes do not create
        // properties on assignment, so emitting `[object] $query = $null` (a correct local) and
        // then rewriting every use to `$this.query` threw on the first write. Only declarations at
        // the class BODY level - outside any method - are members.
        //
        // The block stack says which: anything pushed after the innermost "class" means we are
        // inside a method of it.
        const classAt = blocks.lastIndexOf("class");
        if (classAt !== -1 && blocks.length > classAt + 1) return `$${name} = $null`;

        // Declared as a property either way, but loop counters are NOT registered for $this
        // rewriting: `for ($this.i = 0; ...)` is not valid PowerShell.
        if (!loopVars.has(name)) classMembers.add(name);
        return `[object] $${name} = $null`;
      },

      onErrorResumeNext: () => {
        // DEPLOYED BY AN RMM: the exit code is how the tool decides success or failure.
        // $ErrorActionPreference = 'SilentlyContinue' suppresses the error WITHOUT touching
        // $LASTEXITCODE, so a script that fails, carries on, and exits 0 reports a clean install to
        // PDQ. That is a silent wrong answer with real consequences, so it is refused here even
        // though the same line is fine in a logon script.
        if (ctxKind === "rmm") {
          return { flag:
            "This is deployed by a tool that reads the EXIT CODE. Suppressing errors with "
          + "$ErrorActionPreference = 'SilentlyContinue' does not change $LASTEXITCODE, so a failure "
          + "here would still exit 0 and be reported as a successful deployment. Wrap the steps that "
          + "matter in try/catch and exit with a non-zero code when one of them fails." };
        }

        if (oern.starts.has(curLine)) {
          // THIS region branches on Err, so the error value is part of its control flow. Reset the
          // shim here; the post-pass wraps each statement in the region in try/catch so a failure
          // lands in $VbsErr and execution continues at the next statement, which is what OERN
          // does. Previously this was refused outright because the FILE read Err somewhere.
          return "VbsClearErr";
        }
        // NOT refused just because the FILE reads Err somewhere. This region does not, so it means
        // "do not stop" and the suppression below is faithful. The old file-level test turned four
        // easy regions into TODOs because a fifth one, in another function, was hard.
        return "$ErrorActionPreference = 'SilentlyContinue'   # was: On Error Resume Next"
             + (ctxKind === "logon"
                 ? "\n# NOTE: logon script - keep it non-blocking. No Read-Host, no dialogs, and watch the runtime."
                 : "");
      },

      dialogAdvice: (what) => {
        const swap = what === "InputBox" ? "Read-Host" : "Write-Output";
        if (ctxKind === "logon")
          return `${what} at logon HANGS the sign-in until somebody clicks it, and under a GPO `
               + `script nobody is looking. Replace it with ${swap}, or drop it entirely.`;
        if (ctxKind === "rmm" || ctxKind === "scheduled")
          return `${what} draws a dialog on a desktop that does not exist under SYSTEM, so it blocks `
               + `until the job times out and the deployment reports a failure. Replace it with `
               + `${swap} and let the log carry the message.`;
        if (ctxKind === "interactive")
          return `A human is running this, so a dialog is reasonable - but ${swap} is simpler. For a `
               + `real dialog, load System.Windows.Forms explicitly.`;
        return `PowerShell has no ${what}. Use ${swap} for a console script, or load `
             + `System.Windows.Forms if a dialog is genuinely required - it will not work under SYSTEM.`;
      },

      onErrorGoToZero: () => {
        // Nothing to emit when the region was converted to per-statement try/catch: the guarding
        // simply stops at this line. Refusing here used to leave a TODO in the middle of a region
        // the converter had otherwise handled completely.
        if (oern.endsChecked.has(curLine)) return "# (end of the On Error Resume Next region above)";
        return "$ErrorActionPreference = 'Continue'   # was: On Error GoTo 0";
      },

      /** Emits a function OR a class method, depending on whether a class block is open. */
      method: (name, params, kind) => {
        const inClass = blocks.lastIndexOf("class") !== -1;

        // Computed BEFORE the non-class early return. Placing it after meant the warning never
        // fired for a plain Sub, which is where ByRef mutation almost always lives.
        const byref = byRefMut.get(name);
        const warn = byref
          ? "# REVIEW: this mutates its parameter" + (byref.length > 1 ? "s " : " ")
            + byref.join(", ")
            + ". VBScript passes ByRef BY DEFAULT, so the caller's variable changed. PowerShell"
            + " passes by value and the change is LOST - no error, just a different answer."
            + " Use [ref] parameters and write to .Value, or return the new value.\n"
          : "";

        // IMPLICIT LOCALS - assigned in the body without a Dim - start as $null at the TOP of the
        // procedure. VBScript reads them as Empty until assigned; under strict mode PowerShell
        // throws, and inside a class PowerShell's parser refuses a variable whose only assignment
        // sits in a try or an if ("Variable is not assigned in the method"), which is where an
        // On Error region puts every assignment. `errDesc = Err.Description` inside the region,
        // read on the next line, was the reviewed case.
        const implicit = [...(scopes.assigned.get(name) || EMPTY_SET)]
          .filter((n) => !(scopes.locals.get(name) || EMPTY_SET).has(n) && !scopes.globals.has(n)
            && !classMembers.has(n) && !classMethods.has(n) && !classGetters.has(n) && !funcs.has(n)
            && n !== name && !PS_AUTO_VARS.has(n.toLowerCase()))
          .map((n) => `\n$${n} = $null`).join("");

        if (!inClass)
          return warn + `function ${name} (${paramList(params)}) {`
               + (returningFuncs.has(name) ? `\n$${name} = $null` : "") + implicit;

        const cls = classNames[classNames.length - 1] || "Item";
        if (/^Class_Initialize$/i.test(name)) return `${cls}(${paramList(params)}) {`;
        if (/^Class_Terminate$/i.test(name))
          return `# TODO: Class_Terminate has no PowerShell equivalent - dispose explicitly instead\n`
               + `hidden Dispose(${paramList(params)}) {`;
        // EXPLICIT return type. A PowerShell class method with no annotation is implicitly [void],
        // and `return $x` inside one is a parse error - so a method that returns must say [object].
        //
        // The return variable is also INITIALISED. PowerShell's class analyser rejects reading a
        // variable it cannot prove was assigned, and assignments that only happen inside switch
        // branches do not satisfy it. VBScript initialises the return value to Empty regardless,
        // so declaring it is the faithful thing as well as the parseable one.
        return (returningFuncs.has(name)
          ? `[object] ${name}(${paramList(params)}) {\n$${name} = $null`
          : `[void] ${name}(${paramList(params)}) {`) + implicit;
      },

      /** Closes the innermost block. A close keyword with nothing open is a no-op, NOT a stray `}` -
       *  one unbalanced brace desynchronises every block after it in the file. */
      close: (kind) => ctxClose(kind),

      /** A Case clause closes the previous clause before opening its own, because PowerShell's
       *  switch requires `value { ... }` per clause where VBScript just falls through. */
      /**
       * @param labels already-converted case values.
       *
       * A PowerShell switch clause takes ONE value, so VBScript's `Case 1, 2, 3` has to become a
       * condition block - `{ $_ -in 1,2,3 }`. Emitting the comma list verbatim produced
       * `1, 2, 3 {`, which is a parse error.
       */
      openCase: (labels) => {
        let prefix = "";
        if (blocks[blocks.length - 1] === "case") { blocks.pop(); prefix = "}\n"; }
        blocks.push("case");
        const label = Array.isArray(labels)
          ? (labels.length > 1 ? `{ $_ -in ${labels.join(", ")} }` : labels[0])
          : labels;
        return `${prefix}${label} {`;
      },

      top: () => blocks[blocks.length - 1],

      /** Exposed so the main loop can apply it to EVERY rule's output. Rules that build `$name`
       *  directly (CreateObject, Set assignment) never went through fin(), so their class-member
       *  references stayed bare and PowerShell rejected them with MissingThis. */
      thisify: (c) => thisify(c),

      /** Pops the clause pushed by openCase when the whole clause fitted on one line. */
      closeCaseInline: () => { if (blocks[blocks.length - 1] === "case") blocks.pop(); },

      closeSwitch: () => {
        let out = "";
        if (blocks[blocks.length - 1] === "case") { blocks.pop(); out += "}\n"; }
        const i = blocks.lastIndexOf("switch");
        if (i === -1) return out;
        blocks.length = i;
        return out + "}";
      },
    };
  };

  for (const { n, text, orig = text } of logical) {
    curLine = n;                        // read by ctx.onErrorResumeNext to identify its region
    if (!text.trim()) { out.push({ n, kind: "blank", vbs: text, orig, ps: "" }); continue; }

    // DROP an Authenticode signature block outright. Must sit above the comment match, which would
    // otherwise claim these first - they ARE comments - and emit `#' SIG '' <base64>` for every one
    // of them. A signed script can carry ~300 such lines, so the converted file was two thirds dead
    // base64 and the reader had to scroll past all of it to reach the script.
    //
    // Dropping rather than converting is the honest call: the block signed the VBScript BYTES, so
    // it cannot be valid for the PowerShell, and keeping a stale signature around invites someone
    // to believe the output is signed. An empty `ps` emits nothing at all.
    if (/^\s*''\s*SIG\s*''/.test(text)) { out.push({ n, kind: "comment", vbs: text, orig, ps: "" }); continue; }

    const cm = text.match(/^(\s*)(?:'|Rem\s)(.*)$/i);
    // `#requires` is a PowerShell DIRECTIVE, and a VBScript comment that happens to begin with
    // the word "requires" became one - "The #requires statement must be in one of the following
    // formats". A space after the hash keeps it a comment.
    const safeComment = (s) => (/^\s*requires\b/i.test(s) ? " " + s : s);
    if (cm) { out.push({ n, kind: "comment", vbs: text, orig, ps: `${cm[1]}#${safeComment(cm[2])}` }); continue; }

    let { code, strings } = protectStrings(text);

    // STRIP THE TRAILING COMMENT. VBScript allows `x = 5   ' why` on any line, and passing that
    // through emits `$x = 5   ' why`, where PowerShell reads the apostrophe as the start of a string
    // literal and the line - often the rest of the file - fails to parse.
    //
    // Only safe AFTER protectStrings: an apostrophe inside a string literal ("it's") is not a
    // comment, and every real literal is a placeholder by this point, so any ' left in `code` is a
    // genuine comment marker.
    //
    // This was the single largest source of parse failures, and it went unnoticed because full-line
    // comments WERE handled - so the defect was invisible in a casual read of the output.
    let trailing = "";
    const q = code.indexOf("'");
    if (q !== -1) {
      trailing = restoreStrings(code.slice(q + 1), strings).replace(/\s+$/, "");
      code = code.slice(0, q).replace(/\s+$/, "");
      if (!code.trim()) {                       // the whole line was a comment after all
        out.push({ n, kind: "comment", vbs: text, orig, ps: `#${safeComment(trailing)}` });
        continue;
      }
    }

    const indent = (code.match(/^\s*/) || [""])[0];
    const ctx = mkCtx(strings);
    const withComment = (s) => (trailing ? `${s}  #${trailing}` : s);

    // `obj.Method (arg)` - VBScript allows the space, PowerShell reads `.Method` then a stray
    // parenthesised expression. 36 corpus files failed on `db.OpenView (query)` alone. Done HERE,
    // on the source form, and nowhere later: a first version ran after the operator folds and
    // collapsed the emitted `$objItem.CreationTime (Get-Date)` - two arguments - into a call.
    code = code.replace(/\.([A-Za-z_]\w*)\s+\(/g, ".$1(");
    // A PARAMETER called like a procedure - `operation(array(i))` inside `map(array, operation)`
    // - is a procedure reference passed with GetRef, and must be INVOKED, not indexed: the
    // subscript form `$operation[...]` parses and returns nothing. Only when the script uses
    // GetRef at all, and only for the open procedure's own parameters.
    if (usesGetRef && funcStack.length) {
      // PARAMETERS only (a Dim'd local called with parentheses is an array), and not a parameter
      // the script treats as an array anywhere - UBound(arr), For Each x In arr.
      const params = scopes.params.get(funcStack[funcStack.length - 1].name) || EMPTY_SET;
      code = code.replace(/(?<![\w$.])([A-Za-z_]\w*)\s*\(/g, (mm, n) =>
        params.has(n) && !arrayish.has(n.toLowerCase()) ? `__Invoke(${n}, ` : mm);
      code = code.replace(/__Invoke\((\w+), \s*\)/g, "__Invoke($1)");
    }
    // `obj.Prop("a")("b") = v` - assignment through a DEFAULT member (WshEnvironment's Item, a
    // Dictionary's Item). `)(` before `=` is that default member; name it, and the line is an
    // ordinary property assignment for the rules below. Scripts set environment variables exactly
    // this way.
    code = code.replace(/^(\s*\S.*\))\s*\(([^()]*)\)(\s*=\s*[^=].*)$/, "$1.Item($2)$3");
    // `arr (i, j)` - a space before a subscript, on either side of `=`. Every rule matches `name(`.
    // Only for names known to be VARIABLES: a space before the parenthesis on a procedure call is
    // the ByVal idiom and must stay a call.
    code = code.replace(/(?<![\w$.])([A-Za-z_]\w*)\s+\((?=[^)])/g, (mm, n, off, s) => {
      if (VB_KEYWORDS.has(n.toLowerCase()) || VB_BUILTINS.has(n.toLowerCase())) return mm;
      if (vars.has(n) && !funcs.has(n)) return `${n}(`;
      // A FUNCTION with a space before its argument list is a call too, unless the parenthesised
      // group is only the first of several arguments (`Foo (x), y` - the ByVal idiom on a Sub
      // statement), which must stay as written for the statement rule.
      if (funcs.has(n) || unknownCalls.has(n)) {
        let depth = 0, k = off + mm.length - 1;
        for (; k < s.length; k++) { if (s[k] === "(") depth++; else if (s[k] === ")" && --depth === 0) break; }
        if (!/^\s*,/.test(s.slice(k + 1))) return `${n}(`;
      }
      return mm;
    });
    // NORMALISE the paren-less intrinsic call. VBScript allows both `SetLocale(1033)` and
    // `SetLocale 1033`, and every translation in this file matches on `Name(`. So the SAME
    // intrinsic converted at the top of a script and was left bare inside a function twenty lines
    // later, purely because one call site used parentheses and the other did not. That is the
    // consistency failure worth fixing generally rather than per-intrinsic: adding the parens here
    // routes both forms through one rule.
    //
    // Restricted to bare names in the intrinsic list, so a user-defined Sub call is untouched, and
    // skipped when an `=` follows, which would be an assignment rather than a call.
    {
      const m = code.match(/^(\s*)([A-Za-z_]\w*)\s+(?!\s*=)(\S.*)$/);
      if (m && VB_BUILTINS.has(m[2].toLowerCase()) && !VB_KEYWORDS.has(m[2].toLowerCase()) && !USER_FUNCS.has(m[2].toLowerCase())) {
        code = `${m[1]}${m[2]}(${m[3].trim()})`;
      }
    }

    let handled = false;
    for (const rule of RULES) {
      const m = code.match(rule.rx);
      if (!m) continue;

      if (rule.to === null) {
        out.push({ n, kind: "flagged", vbs: text, orig, ps: `# TODO: ${restoreStrings(code.trim(), strings)}`,
                   rule: rule.name, reason: rule.flag });
        handled = true;
        break;
      }
      const res = rule.to(m, ctx);
      if (res === SKIP) continue;   // regex matched but the rule does not really apply

      // A rule may decide AT RUNTIME that it cannot safely convert this particular script - the
      // On Error handling depends on whether Err is read elsewhere in the file, which no regex on
      // a single line can know.
      if (res && typeof res === "object" && res.flag) {
        out.push({ n, kind: "flagged", vbs: text, orig,
                   ps: `# TODO: ${restoreStrings(code.trim(), strings)}`,
                   rule: rule.name, reason: res.flag });
        handled = true;
        break;
      }
      if (res === null) {           // rule matched but intentionally emits nothing (plain Dim)
        out.push({ n, kind: "converted", vbs: text, orig, ps: "", rule: rule.name });
        handled = true;
        break;
      }
      // Restore string literals HERE, once, for every rule. Rules that build output from a raw
      // capture group (CreateObject, ExecQuery) were emitting the internal placeholder verbatim -
      // "New-Object -ComObject STR0" - because only ctx.expr() restored them. Doing it centrally
      // means a new rule cannot reintroduce the bug by forgetting.
      // thisify BEFORE restoreStrings, so a `$name` occurring inside a string literal is still a
      // placeholder and cannot be rewritten.
      out.push({ n, kind: "converted", vbs: text, orig,
                 ps: withComment(indent + restoreStringsPs(ctx.thisify(res), strings)), rule: rule.name });
      handled = true;
      break;
    }
    if (handled) continue;

    // Bare assignment / method call: the long tail, and the most common thing left.
    //
    // The left-hand side may contain a STRING KEY: `objEnv("SEE_MASK_NOZONECHECKS") = 1` and
    // `dict("key") = value` are ordinary VBScript. String literals are sentinel placeholders at
    // this point, and the sentinels are control characters that `\w` does not match, so the whole
    // statement fell through to "no rule matched" - dropping an indexed assignment entirely.
    // The left-hand side is a SEQUENCE of member accesses and subscripts, not a flat run of
    // allowed characters. A character class cannot express `arr(i, j)` or
    // `objExcel.Cells(r, c).Font.Bold`, because the subscript contains a comma and a space - so a
    // two-dimensional array assignment matched nothing and was dropped as "no rule matched". On a
    // wider corpus it became the single most common unconverted line.
    if (/^\s*[A-Za-z_]\w*(?:\s*\([^()]*\)|\.[A-Za-z_]\w*)*\s*=\s*[^=]/.test(code)) {
      const i = code.indexOf("=");
      const lhsRaw = code.slice(0, i).trim();
      // `stringEndsWith = Right(s,1) = value` - the SECOND `=` is a comparison, not a second
      // assignment. The right-hand side is therefore an expression that may contain comparisons,
      // and translating it in non-condition mode left a bare `=` that PowerShell rejects as an
      // invalid assignment target.
      const rhsRaw = code.slice(i + 1).trim();
      // ANY comparison operator, not just `=`. `IsLogEnabled = lvl<=glLoggingLevel` was missed
      // because the `<` before the `=` excluded it, leaving `<=` in the output - which PowerShell
      // reads as a redirection.
      // Safe to look for bare < and > here: string literals are still placeholders at this point.
      const rhsHasComparison = /<=|>=|<>|[<>]|(?<![-<>!=+*/])=(?!=)/.test(rhsRaw);
      // VBScript returns a value by assigning to the FUNCTION'S OWN NAME. Function names are kept
      // out of the variable registry so call sites are not mangled, which leaves this one idiom
      // emitting a bare identifier on the left of `=` - not parseable PowerShell. Sigil it here
      // only, so the assignment is valid and the call sites stay untouched.
      // The OPEN procedure's own name too, not only the Function/Sub registry: a Property Get is
      // not in `funcs`, so `HasV = ...` inside `Property Get HasV` took the plain path and the
      // getter never returned - "Not all code paths return value" in 16 corpus files.
      const openFunc = funcStack[funcStack.length - 1];
      const isReturnAssign = funcs.has(lhsRaw) || !!(openFunc && openFunc.kind === "func" && openFunc.name === lhsRaw);
      // Record it so `End Function` knows to emit `return $Name`. Assigning to the function's own
      // name IS the return in VBScript, and without this the value was computed and discarded.
      if (isReturnAssign) ctx.noteReturnAssign(lhsRaw);
      out.push({ n, kind: "converted", rule: isReturnAssign ? "function return value" : "assignment", vbs: text, orig,
                 ps: withComment(indent + `${isReturnAssign ? "$" + lhsRaw : ctx.lhs(lhsRaw)} = `
                   + (rhsHasComparison ? `(${ctx.cond(rhsRaw)})` : ctx.expr(rhsRaw))) });
      continue;
    }
    if (/^\s*[A-Za-z_][\w.]*\s*(\(.*\))?\s*$/.test(code)) {
      let call = ctx.stmt(code.trim());
      // `$obj.Close` is a property READ in PowerShell, not a call - it silently does nothing.
      // VBScript allows a parameterless method call without parens, so add them back.
      if (/\.\w+$/.test(call)) call += "()";
      out.push({ n, kind: "converted", rule: "call", vbs: text, orig, ps: withComment(indent + suppressOutput(call)) });
      continue;
    }

    out.push({ n, kind: "unknown", vbs: text, orig, ps: `# TODO: ${text.trim()}`,
               reason: "No rule matched this line. Convert it by hand." });
  }

  // Anything still open at end of file gets closed, with a note. A truncated or oddly-nested script
  // would otherwise emit PowerShell that cannot parse at all - and an unparseable file helps nobody,
  // whereas a closed block plus an explicit warning is at least actionable.
  if (blocks.length) {
    out.push({ n: logical.length, kind: "converted", rule: "auto-close",
      vbs: "", ps: `# TODO: ${blocks.length} block(s) were left open by the source - closed automatically, check nesting\n`
                 + "}".repeat(blocks.length).split("").join("\n") });
  }

  // ON ERROR RESUME NEXT, per REGION rather than per file.
  //
  // VBScript's OERN means two completely different things and the converter treated them as one.
  // Where the script never looks at Err, it means "do not stop" and an empty catch is a faithful
  // translation. Where the script BRANCHES on Err.Number, the error value is part of the control
  // flow and has to be captured. The old analysis asked "does this FILE read Err anywhere", so one
  // Err.Number check in one function turned every OERN in the script into a TODO - four easy
  // conversions refused because of one hard one.
  //
  // A checked region now wraps each of its plain statements in try/catch, which is what OERN
  // actually does: run it, record what happened, carry on at the next statement. Block openers,
  // closers and control flow are left alone - they cannot throw on their own and wrapping them
  // would break the structure.
  for (const l of out) {
    if (!oern.lines.has(l.n)) continue;
    if (l.kind !== "converted" || !l.ps || l.ps.includes("\n")) continue;
    const t = l.ps.trim();
    if (!t || t.startsWith("#")) continue;
    if (/[{}]/.test(t)) continue;                       // block structure, not a statement
    if (/^(return|exit|throw|break|continue|else|elseif|try|catch|finally|param)\b/i.test(t)) continue;
    if (/^VbsClearErr\b/.test(t)) continue;             // the reset itself must not be guarded
    // Assigning a LITERAL cannot throw, so guarding it is pure noise - and there are a lot of
    // them. Skipping these is behaviour-preserving by construction: there is no failure to record.
    if (/^\$[\w.]+\s*=\s*("(?:[^"`]|`.)*"|-?\d+(?:\.\d+)?|\$null|\$true|\$false)\s*$/.test(t)) continue;
    // The TRAILING COMMENT has to come out first. VBScript's `x = 5   ' why` becomes
    // `$x = 5   # why`, and wrapping that verbatim gives
    // `try { $x = 5   # why } catch { VbsSetErr $_ }` - where the closing brace AND the whole catch
    // are inside the comment. PowerShell then reports "The Try statement is missing its Catch or
    // Finally block", and one commented line took out every file that had one.
    const indent = l.ps.match(/^\s*/)[0];
    let code = t, trailing = "";
    let inStr = false;
    for (let k = 0; k < t.length; k++) {
      if (t[k] === '"' && t[k - 1] !== "`") inStr = !inStr;
      else if (t[k] === "#" && !inStr) { code = t.slice(0, k).trimEnd(); trailing = "  " + t.slice(k); break; }
    }
    if (!code) continue;                                 // the whole line was a comment
    // A cmdlet needs -ErrorAction Stop or try/catch never sees it. PowerShell cmdlet errors are
    // NON-terminating by default, so `try { Get-CimInstance ... } catch { }` catches nothing and
    // the Err shim stays zero while the call plainly failed - a silently wrong branch, which is
    // the whole class of bug this converter exists to avoid.
    if (/^[A-Z][a-z]+-[A-Z]\w+/.test(code) && !/-ErrorAction\b/i.test(code)) code += " -ErrorAction Stop";
    l.ps = `${indent}try { ${code} } catch { VbsSetErr $_ }${trailing}`;
  }

  // Scriptlet.TypeLib is the VBScript GUID idiom: create the object, read .Guid, and strip the
  // braces with Mid(g, 2, 36). Left as a COM call it does not fail loudly - the object creation
  // fails, the variable stays $null, and every generated name comes out with an EMPTY guid. That
  // works until two run at once and overwrite each other, which is a race nobody will debug.
  //
  // .Guid is emitted WITH its braces so the surrounding Mid(..., 2, 36) still lands correctly, and
  // upper-cased to match what the COM object returned in case anything downstream compares it.
  {
    const guidVars = new Set();
    for (const l of out) {
      const m = (l.vbs || "").trim()
        .match(/^Set\s+([A-Za-z_]\w*)\s*=\s*(?:WScript\.)?CreateObject\s*\(\s*["']Scriptlet\.TypeLib["']\s*\)/i);
      if (m) { guidVars.add(m[1]); l.ps = ""; }
    }
    for (const name of guidVars) {
      const rx = new RegExp(`\\$${name}\\.Guid\\b`, "gi");
      for (const l of out) {
        if (l.kind === "converted" && l.ps) {
          l.ps = l.ps.replace(rx, `("{" + [guid]::NewGuid().ToString().ToUpper() + "}")`);
        }
      }
    }
  }

  // VBScript.RegExp becomes a NATIVE PowerShell regex.
  //
  // This is the one line that defeats the point of the migration: a converted script still calling
  // `New-Object -ComObject VBScript.RegExp` depends on the component being removed from Windows,
  // so it breaks exactly when the port was supposed to have paid off.
  //
  // The object is stateful - Pattern, IgnoreCase and Global are set on separate lines before the
  // call - so this collects that state, drops the setup lines, and inlines it at the call site.
  //
  // TWO defaults are inverted between the languages, and both are silent:
  //   IgnoreCase defaults to FALSE in VBScript, and `-replace` is case-INSENSITIVE. Without an
  //   explicit IgnoreCase = True the faithful operator is `-creplace`.
  //   Global defaults to FALSE in VBScript, meaning replace the FIRST match only, while `-replace`
  //   replaces every one. Without Global = True the faithful form is a counted .NET Replace.
  {
    const rxVars = new Map();
    // `With re ... .Pattern = "x" ... End With` is how the script libraries of the 2000s configure
    // a RegExp, and it was the ONLY unconverted construct left in the legacy-web corpus. The
    // property assignments arrive with a LEADING DOT and no object name, so matching `name.Pattern`
    // alone missed every one of them.
    let withTarget = null;
    for (const l of out) {
      const v = (l.vbs || "").trim();
      let m;
      if ((m = v.match(/^With\s+([A-Za-z_]\w*)\s*$/i))) { withTarget = m[1]; continue; }
      if (/^End\s+With\s*$/i.test(v)) { withTarget = null; continue; }

      if ((m = v.match(/^Set\s+([A-Za-z_]\w*)\s*=\s*(?:New\s+RegExp\b|CreateObject\s*\(\s*["']VBScript\.RegExp["']\s*\))/i))) {
        rxVars.set(m[1], { pattern: null, ignoreCase: false, global: false });
        // kind too, not just ps: a line whose text is blanked but whose kind is still "unknown"
        // keeps counting against the conversion rate and keeps its TODO in the output.
        l.kind = "converted";
        l.rule = "RegExp (folded into the call site)";
        l.ps = "";                                   // the object itself has no equivalent
        continue;
      }
      if ((m = v.match(/^(?:([A-Za-z_]\w*)\s*)?\.?\s*(Pattern|IgnoreCase|Global)\s*=\s*(.+?)\s*$/i))
          || (m = v.match(/^\.(Pattern|IgnoreCase|Global)\s*=\s*(.+?)\s*$/i))) {
        // Either `re.Pattern = x` or, inside a With block, `.Pattern = x`.
        const owner = m.length === 4 ? (m[1] || withTarget) : withTarget;
        const key = (m.length === 4 ? m[2] : m[1]).toLowerCase();
        // Trailing comment CUT, on the string-protected form. `re.Pattern = "5900" 'vnc port`
        // carried its comment into the folded call site, inside the parentheses.
        const valRaw = (m.length === 4 ? m[3] : m[2]).trim();
        const { code: vc, strings: vs } = protectStrings(valRaw);
        const vq = vc.indexOf("'");
        const val = restoreStrings(vq === -1 ? vc : vc.slice(0, vq), vs).trim();
        if (!owner || !rxVars.has(owner)) continue;
        const st = rxVars.get(owner);
        if (key === "pattern") {
          st.pattern = val;
          // The CONVERTED right-hand side too. A pattern held in a variable - `re.Pattern = pat`
          // inside a helper - reached the call site as the bare VBScript name, and PowerShell saw
          // `-match pat` with a missing operand. A literal still takes the quoting path below.
          st.patternPs = l.kind === "converted"
            ? (((l.ps || "").match(/=\s*(.+?)\s*$/) || [])[1] || "").replace(/\s{2,}#.*$/, "") || null
            // The line was refused (a With block, usually), so its ps is a TODO. A bare name is
            // still a variable and gets its sigil; anything else is left for the reader.
            : (/^[A-Za-z_]\w*$/.test(val) ? "$" + val : null);
        }
        else if (key === "ignorecase") st.ignoreCase = /^true$/i.test(val);
        else st.global = /^true$/i.test(val);
        l.kind = "converted";
        l.rule = "RegExp (folded into the call site)";
        l.ps = "";                                   // folded into the call site below
      }
    }

    if (rxVars.size) {
      for (const l of out) {
        if (l.kind !== "converted" || !l.ps) continue;
        for (const [name, st] of rxVars) {
          // A VBScript literal becomes a SINGLE-quoted PowerShell string: a regex is full of `$`
          // and `\`, and a double-quoted string would interpolate the first and mangle the second.
          const lit = st.pattern && /^".*"$/s.test(st.pattern)
            ? `'${st.pattern.slice(1, -1).replace(/""/g, '"').replace(/'/g, "''")}'`
            : (st.patternPs || st.pattern);
          if (!lit) continue;

          const rxRep = new RegExp(`\\$${name}\\.Replace\\s*\\(\\s*([^,]+?)\\s*,\\s*(.+?)\\s*\\)`, "gi");
          l.ps = l.ps.replace(rxRep, (mm, subject, repl) =>
            (st.global
              ? `(${subject} ${st.ignoreCase ? "-replace" : "-creplace"} ${lit}, ${repl})`
              // Not Global: VBScript replaces only the FIRST match, which -replace cannot express.
              : `([regex]::new(${lit}${st.ignoreCase ? ", 'IgnoreCase'" : ""}).Replace(${subject}, ${repl}, 1))`));

          const rxTest = new RegExp(`\\$${name}\\.Test\\s*\\(\\s*(.+?)\\s*\\)`, "gi");
          l.ps = l.ps.replace(rxTest, (mm, subject) =>
            `(${subject} ${st.ignoreCase ? "-match" : "-cmatch"} ${lit})`);
        }
      }
    }
  }

  // WMI METHOD calls become Invoke-CimMethod.
  //
  // Get-CimInstance returns CimInstance objects, which carry PROPERTIES ONLY. VBScript's
  // SWbemObject exposed WMI methods directly, so `objService.StartService()` worked there and
  // fails here with "method not found". Earlier scripts never caught this because they only ever
  // read CIM properties.
  //
  // Scope is deliberately narrow: only variables that this script obtained FROM a WMI query - the
  // collection assigned from ExecQuery, and the loop variable iterating it - and only for the
  // method names WMI actually defines. A method call on any other object is left alone, because
  // rewriting a COM call into Invoke-CimMethod would be its own silent breakage.
  {
    const wmiVars = new Set();
    for (const l of out) {
      const v = (l.vbs || "");
      let m;
      if ((m = v.match(/^\s*Set\s+([A-Za-z_]\w*)\s*=\s*.*\.ExecQuery\b/i))) wmiVars.add(m[1]);
      if ((m = v.match(/^\s*For\s+Each\s+([A-Za-z_]\w*)\s+In\s+([A-Za-z_]\w*)/i))
          && wmiVars.has(m[2])) wmiVars.add(m[1]);
    }
    if (wmiVars.size) {
      const names = [...wmiVars].sort((a, b) => b.length - a.length).join("|");
      const WMI_METHODS = "StartService|StopService|PauseService|ResumeService|Delete|Create|"
                        + "Terminate|ChangeStartMode|Change|SetPowerState|Rename|Reconfigure";
      const rx = new RegExp(`\\$(${names})\\.(${WMI_METHODS})\\s*\\(\\s*([^()]*?)\\s*\\)`, "gi");
      for (const l of out) {
        if (l.kind !== "converted" || !l.ps || !/\$/.test(l.ps)) continue;
        let needsArgNote = false;
        const next = l.ps.replace(rx, (mm, obj, method, args) => {
          const a = splitArgs(args).map((x) => x.trim()).filter(Boolean);
          // WMI method arguments are POSITIONAL in VBScript and NAMED in Invoke-CimMethod, and the
          // names cannot be derived from the call site. Emit them as a hashtable skeleton the
          // reader fills in rather than inventing parameter names that would silently bind wrong.
          // Parameter names from a small table where they are unambiguous. WMI method arguments
          // are POSITIONAL in VBScript and NAMED in Invoke-CimMethod, and `Arg1` never binds.
          const known = WMI_METHOD_PARAMS[method.toLowerCase()];
          const argPart = a.length
            ? ` -Arguments @{ ${a.map((x, i) =>
                `${known && known[i] ? known[i] : `Arg${i + 1}`} = ${x}`).join("; ")} }`
            : "";
          if (a.length && (!known || a.length > known.length)) needsArgNote = true;
          // The method name is QUOTED. Emitting it bare let the class-method rewrite below mistake
          // it for a reference to a sibling method of the same name - `ChangeStartMode` is both a
          // WMI method and a method of the class wrapping it - and turn it into
          // `-MethodName $this.ChangeStartMode()`, a zero-argument recursive call. Quoting removes
          // the collision entirely rather than relying on the rewrite to be clever about it.
          return `(Invoke-CimMethod -InputObject $${obj} -MethodName '${method}'${argPart}).ReturnValue`;
        });
        // The note goes at the END OF THE LINE, never inside the expression. Emitting it inline
        // put `# TODO...` before the closing `).ReturnValue`, so the rest of the expression was
        // commented out and the file stopped parsing - the exact trap the FormatDateTime rule
        // above already records.
        l.ps = next + (needsArgNote ? "   # TODO: name these arguments per the WMI class" : "");
      }
    }
  }

  // Sibling METHOD calls inside a class become `$this.Name(a, b)`.
  //
  // Done over the finished lines because the emit sites disagree: the bare-sub-call path runs
  // `fin` on each ARGUMENT separately rather than on the whole statement, so a filter applied
  // inside `fin` never saw the method name at all. Same lesson as the output-capture pass - one
  // sweep over the output cannot miss a site the way a per-site filter can.
  //
  // VBScript calls a sibling by bare name (`AddFilter "Name", strName`). In a PowerShell class
  // that is a command lookup for a function that does not exist, so the FIRST such call ends the
  // script. They also need real method syntax: parentheses and commas, not spaces.
  if (classMethods.size) {
    const names = [...classMethods].sort((a, b) => b.length - a.length).join("|");
    // `(?<![A-Za-z]-)` on both forms: a cmdlet is not a method. A class with a `Sleep` method
    // turned every emitted `Start-Sleep` into `Start-$this.Sleep()`.
    const callRx = new RegExp(`(?<![\\w$.])(?<![A-Za-z]-)(${names})\\s*\\(`, "g");
    // NO bare-form rewriting here. Reconstructing `Name a b` from the emitted line means guessing
    // where the arguments split, and the commas are already gone by this point - `SetIt n * 2`
    // (one argument) came back as three. The paren-less form is handled at its emit site instead,
    // where the commas still exist. This pass only fixes calls that already have parentheses.
    let depth = 0;
    for (const l of out) {
      const v = (l.vbs || "").trim();
      if (/^Class\s+[A-Za-z_]\w*/i.test(v)) { depth++; continue; }
      if (/^End\s+Class\b/i.test(v)) { depth = Math.max(0, depth - 1); continue; }
      if (!depth || l.kind !== "converted" || !l.ps) continue;
      const t = l.ps;
      if (/^\s*#/.test(t) || /^\s*\[\w+\]\s/.test(t)) continue;   // comment, or a declaration
      // MASK THE STRINGS FIRST. Identifier rewriting has to be token-aware: this pass runs on the
      // finished line, where literals are already restored, and it happily rewrote inside one -
      // `"No services loaded. Call LoadServices() first."` became `Call $this.LoadServices()`,
      // which PowerShell then EXPANDS in a double-quoted string, so the user-facing message read
      // "Call ServiceManagementHelper.LoadServices() first." A method name is also an ordinary
      // word, so any prose mentioning one was at risk.
      const masked = [];
      let s = t.replace(/"(?:[^"`]|`.)*"|'[^']*'/g, (lit) => {
        masked.push(lit);
        return `LIT${masked.length - 1}`;
      });

      s = s.replace(callRx, (mm, n) => `$this.${n}(`);
      // ZERO-ARGUMENT calls, in any position: `VerifyServicesLoaded` on its own, or
      // `$x = GetBias`. VBScript calls a no-arg method by bare name with no parentheses at all, so
      // nothing upstream sees a call to rewrite - and a bare name in PowerShell is a command
      // lookup that fails. This was the remaining half of the $this rule: the argument forms were
      // handled and the no-argument ones were not, so the first line of the main method still died.
      //
      // Safe to do here precisely BECAUSE there are no arguments: nothing has to be reconstructed.
      // The lookahead skips names already followed by `(` (handled above) and the lookbehind skips
      // `$this.Name` and `.Name`.
      s = s.replace(new RegExp(`(?<![\\w$.])(?<![A-Za-z]-)(${names})\\b(?!\\s*\\()`, "g"), (mm, n) => `$this.${n}()`);
      l.ps = s.replace(/LIT(\d+)/g, (mm, i) => masked[Number(i)]);
    }
  }

  // PARAMETERLESS COM METHODS keep their parentheses. VBScript calls `dict.Keys` without them; in
  // PowerShell that is the METHOD OBJECT, so `foreach ($id in $processes.Keys)` iterated once
  // over a PSMethod, ran `taskkill /pid <method signature>`, and the script that exists to kill
  // a process killed nothing and exited 0. Keys/Items/RemoveAll only on variables created as a
  // Scripting.Dictionary (Items is a PROPERTY on other objects); ReadAll/ReadLine are methods on
  // every object that has them.
  {
    const dictNames = new Set();
    for (const l of out) {
      const mm = (l.vbs || "").match(/^\s*Set\s+([A-Za-z_]\w*)\s*=\s*(?:WScript\.)?CreateObject\s*\(\s*"Scripting\.Dictionary"/i);
      if (mm) dictNames.add(mm[1]);
    }
    // The dictionary usually arrives through a RETURN value or a parameter, not the variable it
    // was created in (`killProcesses(getProcessesWithName(...))`), so tracking the creating
    // variable alone missed the reviewed line. Keys and RemoveAll exist on nothing else this
    // converter can emit and are rewritten everywhere; Items is a PROPERTY on Outlook folders,
    // so it is rewritten everywhere only when the script does not touch Outlook or MAPI.
    const outlook = out.some((l) => /Outlook\.Application|MAPI/i.test(l.vbs || ""));
    const dictRx = dictNames.size
      ? new RegExp(`(\\$(?:this\\.|script:)?(?:${[...dictNames].join("|")}))\\.(Keys|Items|RemoveAll)\\b(?!\\s*\\()`, "g")
      : null;
    const anyRx = outlook ? /\.(Keys|RemoveAll)\b(?!\s*\()/g : /\.(Keys|Items|RemoveAll)\b(?!\s*\()/g;
    for (const l of out) {
      if (l.kind !== "converted" || !l.ps || /^\s*#/.test(l.ps)) continue;
      const masked = [];
      let s = l.ps.replace(/"(?:[^"`]|`.)*"|'[^']*'/g, (lit) => { masked.push(lit); return `LIT${masked.length - 1}`; });
      if (dictRx) s = s.replace(dictRx, "$1.$2()");
      s = s.replace(anyRx, ".$1()");
      // KEYS through VbsComKey: a Scripting.Dictionary over COM interop throws
      // CTL_E_ILLEGALFUNCTIONCALL on an unsigned key, and Get-CimInstance hands back ProcessId
      // as [uint32] where VBScript's WMI gave a Long. Add, Exists, Item, Remove and Key all get
      // the same conversion so a key stored converted is also looked up converted.
      if (dictNames.size)
        s = s.replace(new RegExp(`(\\$(?:this\\.|script:)?(?:${[...dictNames].join("|")}))\\.(Add|Exists|Item|Remove|Key)\\(\\s*([^,()]+?)\\s*(,|\\))`, "g"),
          (mm, obj, meth, key, sep) => `${obj}.${meth}((VbsComKey ${key})${sep}`);
      s = s.replace(/\.(ReadAll|ReadLine)\b(?!\s*\()/g, ".$1()");
      l.ps = s.replace(/LIT(\d+)/g, (mm, i) => masked[Number(i)]);
    }
  }

  // Discard the value of every call used as a STATEMENT, wherever it was emitted.
  //
  // Done as a pass over the finished lines rather than at each emit site, because patching one
  // site is how this half-shipped the first time: COM method calls got `$null =` and calls to the
  // script's OWN functions did not, so `ContentLog "msg"` still returned $true. Two of those
  // inside a function made it return @($true, $true, $false) - truthy - and the caller took the
  // wrong branch. At top level the same calls printed a stray True into the action log.
  //
  // One pass over the output cannot miss an emit site the way a per-site fix can.
  for (const l of out) {
    if (l.kind !== "converted" || !l.ps || l.ps.includes("\n")) continue;
    l.ps = suppressOutput(flattenConcat(l.ps));
  }

  // LAST PASS: catch VBScript intrinsics that survived into the output as bare words.
  //
  // This is the most important refusal in the file. Any intrinsic without a rule - SetLocale,
  // DateDiff, WeekdayName, WScript.Arguments.Named, Err.Number - was emitted UNCHANGED, and a bare
  // name in PowerShell is a command invocation. So the line parsed, shipped, and died at runtime
  // with "not recognized as a name of a cmdlet", one per intrinsic, scattered through the script.
  //
  // Whitelisting every intrinsic we handle would go stale the moment someone adds a rule. Instead
  // this asks the opposite question: does a VBScript builtin NAME still appear in the emitted
  // PowerShell, being called? If it does, no rule claimed it, whatever the reason. A converter that
  // says "I could not do this line" is usable. One that emits a call to a function that does not
  // exist is not, and looks finished while being broken.
  for (const l of out) {
    if (l.kind !== "converted" || !l.ps) continue;
    const left = leftoverIntrinsic(l.ps);
    if (!left) continue;
    l.kind = "flagged";
    l.rule = "untranslated intrinsic";
    l.reason = `${left} is a VBScript intrinsic with no PowerShell equivalent here, and it survived `
             + `into the output as a bare name. PowerShell would treat that as a command and fail at `
             + `runtime. Replace it by hand.`;

    // A line that OPENS OR CLOSES a block cannot simply be replaced by a comment: deleting
    // `If Err.Number <> 0 Then` takes its `{` with it and every brace after it is orphaned, so one
    // untranslated intrinsic turned into "Unexpected token '}'" and a file that would not parse at
    // all. Keep the code and put the warning above it instead. The line is still wrong at runtime,
    // but it is wrong in ONE place with a note attached, rather than destroying the whole file.
    // KEEP THE CODE, always. Two earlier versions of this tried to be clever about when a line was
    // safe to replace with a comment - first "are the braces unbalanced", then "does it contain a
    // brace" - and both still destroyed files, because `If Err.Number <> 0 Then` opens a block and
    // deleting it orphans every brace that follows. One untranslated intrinsic would take a whole
    // 1,700-line script from 19 parse errors to 53.
    //
    // The warning is the product here, not the deletion. The line was already going to fail at
    // runtime; leaving it in place with a TODO above it means the reader gets a file that parses,
    // opens in an editor, and points at the exact line. Deleting it buys nothing and costs the
    // rest of the file.
    l.ps = `# TODO (untranslated intrinsic): ${l.reason}\n${l.ps}`;
    l.rendered = true;      // emit `ps` as-is; do NOT rebuild this line from rule/reason/vbs
  }

  const code = out.filter((l) => l.kind !== "blank" && l.kind !== "comment");
  const converted = code.filter((l) => l.kind === "converted").length;
  return {
    lines: out,
    stats: {
      total: out.length,
      codeLines: code.length,
      converted,
      flagged: code.filter((l) => l.kind === "flagged").length,
      unknown: code.filter((l) => l.kind === "unknown").length,
      rate: code.length ? converted / code.length : 1,
    },
  };
}

/**
 * Decodes a .vbs file's bytes to text.
 *
 * Lives here rather than in the page so it can actually be TESTED: .vbs written by Windows tooling
 * is very often UTF-16LE, and decoding that as UTF-8 does not throw - it yields text with a NUL
 * between every character, which then converts to confident garbage. A silent wrong answer on the
 * very first step is the worst possible failure for this tool.
 *
 * The no-BOM heuristic exists because plenty of hand-saved files are UTF-16 without one; ASCII
 * text in UTF-16LE is half NUL bytes, which nothing else looks like.
 */
export function decodeVbsFile(buf) {
  const b = new Uint8Array(buf);
  if (b.length >= 2 && b[0] === 0xff && b[1] === 0xfe) return new TextDecoder("utf-16le").decode(buf);
  if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) return new TextDecoder("utf-16be").decode(buf);

  const head = b.subarray(0, Math.min(200, b.length));
  let nuls = 0;
  for (const x of head) if (x === 0) nuls++;
  if (head.length > 20 && nuls > head.length / 4) return new TextDecoder("utf-16le").decode(buf);

  return new TextDecoder("utf-8").decode(buf).replace(/^﻿/, "");
}

// ── migration audit ──────────────────────────────────────────────────────────
/**
 * Answers "how much work is this script to migrate?" WITHOUT emitting any code.
 *
 * This is the honest half of the tool. The converter can be subtly wrong; an audit cannot, because
 * it never claims anything runs. It also answers the question an LLM code-converter does not: the
 * planning question, which is the one with a 2027 deadline attached to it.
 */
const CAPABILITIES = [
  { key: "File system",      rx: /(Scripting\.FileSystemObject|\.(OpenTextFile|CreateTextFile|FileExists|FolderExists|DeleteFile|CopyFile|MoveFile|GetFolder|GetFile)\b)/i,
    note: "Maps cleanly onto Get-Item, Test-Path, Get-Content and Set-Content." },
  { key: "WMI",              rx: /(winmgmts:|\.ExecQuery\s*\()/i,
    note: "Becomes Get-CimInstance. Usually the tidiest part of a migration." },
  { key: "Registry",         rx: /\.Reg(Read|Write|Delete)\b/i,
    note: "Becomes Get-ItemProperty and Set-ItemProperty." },
  { key: "Process / shell",  rx: /\.(Run|Exec)\s*\(/i,
    note: "Becomes Start-Process, or the native command directly." },
  { key: "Active Directory", rx: /["'](WinNT|LDAP):\/\//i,
    note: "Worth rewriting with the ActiveDirectory module rather than translating ADSI literally." },
  { key: "Database (ADO)",   rx: /(ADODB\.|\.Recordset\b)/i,
    note: "Keeps working through the same COM objects, but Invoke-SqlCmd is usually better." },
  { key: "Classes",          rx: /^\s*Class\s+\w/im,
    note: "PowerShell 5+ has classes, but Property Get/Let/Set does not map 1:1. Expect hand work." },
  { key: "User prompts",     rx: /\b(MsgBox|InputBox)\b/i,
    note: "No equivalent under SYSTEM. Console scripts use Write-Output and Read-Host." },
  { key: "Error suppression",rx: /On\s+Error\s+Resume\s+Next/i,
    note: "The one genuinely risky construct. Needs deciding per call site, not translating." },
  { key: "Dynamic code",     rx: /\b(Eval|Execute|ExecuteGlobal)\s*\(/i,
    note: "Should be rewritten explicitly rather than carried over as Invoke-Expression." },
];

export function auditVbs(src, opts = {}) {
  const { lines, stats } = convertVbs(src, opts);
  const err = analyseErrorHandling(joinContinuations(src));
  const detected = detectContext(src);
  const ctxKind = opts.context && opts.context !== "unknown" ? opts.context : detected;

  const blockers = lines
    .filter((l) => l.kind === "flagged" || l.kind === "unknown")
    .map((l) => ({ line: l.n, kind: l.kind, rule: l.rule || "unrecognised",
                   reason: l.reason, text: (l.orig ?? l.vbs).trim() }));

  const byReason = new Map();
  for (const b of blockers) byReason.set(b.rule, (byReason.get(b.rule) || 0) + 1);

  const capabilities = CAPABILITIES
    .filter((c) => c.rx.test(src))
    .map((c) => ({ key: c.key, note: c.note }));

  // Effort is driven by the number of DISTINCT things needing a decision, not raw line count: ten
  // instances of one pattern is one decision applied ten times, and calling that ten units of work
  // would overstate every large script.
  const decisions = byReason.size;

  // Name the ACTUAL risk rather than a generic one. The first version said "including error
  // suppression" for any risky script, and happily said it about scripts whose only risk was
  // dynamic code - a report that misdescribes its own finding teaches the reader to distrust it.
  const risks = [];
  if (/On\s+Error\s+Resume\s+Next/i.test(src))
    risks.push("error suppression that changes which failures are swallowed");
  if (/\b(Eval|Execute|ExecuteGlobal)\s*\(/i.test(src))
    risks.push("code built and run from strings at runtime");
  const risky = risks.length > 0;

  let verdict, detail;
  if (blockers.length === 0) {
    verdict = "Mechanical";
    detail = "Nothing in this script needs a judgement call. Convert it and test it.";
  } else if (decisions <= 2 && !risky) {
    verdict = "Straightforward";
    detail = `${decisions} pattern${decisions === 1 ? "" : "s"} need a decision, repeated across `
           + `${blockers.length} line${blockers.length === 1 ? "" : "s"}.`;
  } else if (decisions <= 5) {
    verdict = "Moderate";
    detail = `${decisions} distinct pattern${decisions === 1 ? "" : "s"} need${decisions === 1 ? "s" : ""} a decision`
           + (risky ? `, including ${risks.join(" and ")}.` : ".");
  } else {
    verdict = "Rewrite";
    detail = `${decisions} distinct patterns need decisions. At this density it is usually faster to `
           + "rewrite the script against the same requirement than to translate it line by line.";
  }

  // Context changes the advice, so say what KIND of script this looks like. A logon script has
  // constraints a scheduled task does not: it must not block, must not prompt, and must not be slow.
  const context = [];
  if (SCRIPT_CONTEXTS[ctxKind]?.note)
    context.push({ key: SCRIPT_CONTEXTS[ctxKind].label, note: SCRIPT_CONTEXTS[ctxKind].note });

  // The exit code only matters where something reads it - say so, because it changes what
  // "converted correctly" means for the whole script.
  if (ctxKind === "rmm" && /WScript\.Quit/i.test(src))
    context.push({ key: "Exit codes are load-bearing here",
      note: "WScript.Quit becomes exit, and the deploying tool reads it. Check every path that can "
          + "fail actually reaches a non-zero exit - suppressed errors do not change $LASTEXITCODE." });

  if (err.mode === "checked")
    context.push({ key: "Branches on error codes",
      note: "Err.Number is read, so error values drive the logic. These need real try/catch with the "
          + "branches rewritten - suppressing the error would send the script down the wrong path." });
  else if (err.mode === "suppress")
    context.push({ key: "Suppresses errors without checking them",
      note: "On Error Resume Next is used but Err is never read, so the intent is just to keep "
          + "going. Converted to $ErrorActionPreference = 'SilentlyContinue'." });

  return {
    verdict, detail, capabilities, blockers, context,
    /** What we guessed, and what was actually applied - so the UI can show the guess and let the
     *  user correct it rather than silently acting on it. */
    detectedContext: detected,
    appliedContext: ctxKind,
    summary: {
      codeLines: stats.codeLines,
      cleanLines: stats.converted,
      cleanPct: stats.codeLines ? Math.round((stats.converted / stats.codeLines) * 100) : 100,
      needsReview: blockers.length,
      distinctDecisions: decisions,
    },
    topReasons: [...byReason.entries()].sort((a, b) => b[1] - a[1]).map(([rule, count]) => ({ rule, count })),
  };
}

/**
 * The runtime the emitted code calls into.
 *
 * Inlined rather than shipped as a module import so a converted script is a single file you can
 * paste and run. Only the helpers actually used are emitted, so a script that never concatenates
 * does not carry VbsConcat.
 */
const RUNTIME = {
  // VARIADIC, so `a & b & c & d` emits one call instead of a four-deep nest. $args rather than a
  // typed array parameter: a typed [object[]] would collapse a single array operand into its
  // elements, and VBScript concatenating an array is a different bug to let through silently.
  VbsConcat: 'function VbsConcat { ($args | ForEach-Object { [string]$_ }) -join "" }',
  // VBScript True is -1 (every bit set), which is the whole reason its bitwise And/Or read as
  // logical. PowerShell $true is 1, so normalising here is what lets ONE operator be correct for
  // both `mybyte And &H80` and `a = 1 And b = 2`.
  VbsNum: 'function VbsNum($v) { [long](VbsParseNum $v) }',
  // TYPE-PRESERVING, because VBScript's are. `Not True` is the Boolean False and prints "False";
  // returning the bitwise -bnot result printed "0" instead. Booleans in, boolean out; numbers in,
  // bit mask out. The suite caught this the moment the bitwise cases existed.
  VbsAnd: 'function VbsAnd($a, $b) {\n'
        + '  if (($a -is [bool]) -and ($b -is [bool])) { $a -and $b }\n'
        + '  else { (VbsNum $a) -band (VbsNum $b) }\n}',
  VbsOr: 'function VbsOr($a, $b) {\n'
       + '  if (($a -is [bool]) -and ($b -is [bool])) { $a -or $b }\n'
       + '  else { (VbsNum $a) -bor (VbsNum $b) }\n}',
  VbsXor: 'function VbsXor($a, $b) {\n'
        + '  if (($a -is [bool]) -and ($b -is [bool])) { $a -xor $b }\n'
        + '  else { (VbsNum $a) -bxor (VbsNum $b) }\n}',
  // -bnot on a PowerShell $true gives -2, which is TRUTHY - the opposite of what VBScript means.
  VbsNot: 'function VbsNot($a) { if ($a -is [bool]) { -not $a } else { -bnot (VbsNum $a) } }',
  // Asc is the ANSI code page, AscW is the UTF-16 code unit. They agree below 256 and diverge
  // above it, so `[int][char]` was right for AscW and quietly wrong for Asc on anything non-Latin.
  //
  // NOT [Text.Encoding]::Default: that is the ANSI page on Windows PowerShell 5.1 and UTF-8 on
  // PowerShell 7, so the helper would mean different things depending on which host ran it - a
  // converted script is expected to behave the same on both. Pinning to the culture's ANSI code
  // page is what the VBScript host actually used.
  VbsAsc: 'function VbsAsc($s) {\n'
        + '  $s = [string]$s\n'
        + '  if ($s.Length -eq 0) { return 0 }\n'
        + '  $enc = [Text.Encoding]::GetEncoding([cultureinfo]::CurrentCulture.TextInfo.ANSICodePage)\n'
        + '  [int]($enc.GetBytes($s.Substring(0, 1)))[0]\n}',
  // VBScript collections are 1-based for .Item but Named is keyed, so only Exists/Item matter.
  // PSCustomObject with ScriptMethods rather than a hashtable, so `.Exists("X")` and `.Item("X")`
  // from the original convert with no change at the call site.
  // A REAL Scripting.Dictionary, not a PSCustomObject with ScriptMethods.
  //
  // The PSCustomObject version supported .Exists() and .Item() and still failed, because most call
  // sites use the INDEXER - `$colNamedArguments["Command"]` - and indexing a PSCustomObject never
  // reaches an Item ScriptMethod. You get $null or an error depending on version, so the variable
  // silently stayed empty and the script did nothing. The COM dictionary supports Exists, Item, the
  // indexer and Count natively, which is the whole point: the call sites convert unchanged.
  //
  // CompareMode = 1 is text (case-insensitive) compare, matching WScript's named arguments, and it
  // MUST be set before the first add or the dictionary keeps the default binary comparison.
  // The leading comma stops PowerShell unrolling the dictionary into its keys on return.
  //
  // Reads $script:VbsScriptArgs, set once in the prelude. It cannot read $args directly: inside a
  // function $args is that FUNCTION'S arguments, and these calls are usually inside one.
  VbsNamedArgs: 'function VbsNamedArgs {\n'
              + '  $d = New-Object -ComObject Scripting.Dictionary\n'
              + '  $d.CompareMode = 1\n'
              + '  foreach ($a in $script:VbsScriptArgs) {\n'
              + '    if ($a -match \'^[/-]([^:=]+)[:=]?(.*)$\') { $d[$Matches[1]] = $Matches[2] }\n'
              + '  }\n'
              + '  ,$d\n}',
  VbsUnnamedArgs: 'function VbsUnnamedArgs { @($script:VbsScriptArgs | Where-Object { $_ -notmatch \'^[/-]\' }) }',
  // VBScript truncates toward zero on every interval, hence [int] on each. The calendar intervals
  // THROW rather than approximate: "months between" is not TotalDays/30, and a converter quietly
  // returning a number that is close is the failure mode this whole tool exists to avoid.
  // The Err shim. Number is what scripts branch on; 5 is VBScript's generic "invalid procedure
  // call", which is the closest single code to "something threw and we do not know its HRESULT".
  // Where the exception carries a real HRESULT it is used instead, so `If Err.Number = &H80070005`
  // style checks keep working.
  // The .NET exception TYPE maps to the VBScript number a script would have checked against:
  // 11 divide by zero, 9 subscript out of range, 13 type mismatch, 91 object not set, 53 file not
  // found, 76 path not found, 70 permission denied, 6 overflow. A COM error keeps its HRESULT.
  VbsSetErr: 'function VbsSetErr($e) {\n'
           + '  $n = 5\n'
           + '  $x = $e.Exception\n'
           + '  while ($x -and $x.InnerException -and $x -is [System.Management.Automation.RuntimeException]) { $x = $x.InnerException }\n'
           + '  if ($x -is [System.DivideByZeroException]) { $n = 11 }\n'
           + '  elseif ($x -is [System.IndexOutOfRangeException] -or $x -is [System.ArgumentOutOfRangeException]) { $n = 9 }\n'
           + '  elseif ($x -is [System.InvalidCastException] -or $x -is [System.FormatException]) { $n = 13 }\n'
           + '  elseif ($x -is [System.NullReferenceException]) { $n = 91 }\n'
           + '  elseif ($x -is [System.IO.FileNotFoundException]) { $n = 53 }\n'
           + '  elseif ($x -is [System.IO.DirectoryNotFoundException]) { $n = 76 }\n'
           + '  elseif ($x -is [System.UnauthorizedAccessException]) { $n = 70 }\n'
           + '  elseif ($x -is [System.OverflowException]) { $n = 6 }\n'
           + '  elseif ($x -is [System.Runtime.InteropServices.COMException]) { $n = $x.HResult }\n'
           + '  elseif ($x -and $x.HResult) { $n = $x.HResult }\n'
           + '  $script:VbsErr = [pscustomobject]@{\n'
           + '    Number = $n\n'
           + '    Description = [string]$e.Exception.Message\n'
           + '    Source = [string]$e.InvocationInfo.MyCommand.Name\n'
           + '  }\n}',
  VbsClearErr: 'function VbsClearErr {\n'
             + '  $script:VbsErr = [pscustomobject]@{ Number = 0; Description = ""; Source = "" }\n}',
  // VBScript weekdays are 1-based starting Sunday; DayOfWeek is 0-based starting Sunday.
  // Pinned to en-US (1033), NOT CurrentCulture. A thread culture assignment does not reliably
  // persist between statements in Windows PowerShell 5.1, which is what most deployment tools run,
  // so a script that sets the locale and then formats a date can still get localized day and month
  // names on a non-English machine. Scripts calling these are almost always building a fixed-format
  // string - RFC 822 headers, log timestamps - where a localized name is a defect. Pinning also
  // makes the SetLocale calls around them unnecessary rather than load-bearing.
  VbsWeekdayName: 'function VbsWeekdayName($n, $abbrev) {\n'
                + '  $f = [cultureinfo]::GetCultureInfo(1033).DateTimeFormat\n'
                + '  $d = [DayOfWeek]((([int]$n - 1) % 7 + 7) % 7)\n'
                + '  if ($abbrev) { $f.GetAbbreviatedDayName($d) } else { $f.GetDayName($d) }\n}',
  VbsMonthName: 'function VbsMonthName($n, $abbrev) {\n'
              + '  $f = [cultureinfo]::GetCultureInfo(1033).DateTimeFormat\n'
              + '  if ($abbrev) { $f.GetAbbreviatedMonthName([int]$n) } else { $f.GetMonthName([int]$n) }\n}',
  VbsDateDiff: 'function VbsDateDiff($interval, $d1, $d2) {\n'
             + '  $t = New-TimeSpan -Start ([datetime]$d1) -End ([datetime]$d2)\n'
             + '  switch ([string]$interval) {\n'
             + '    "s"  { [int]$t.TotalSeconds }\n'
             + '    "n"  { [int]$t.TotalMinutes }\n'
             + '    "h"  { [int]$t.TotalHours }\n'
             + '    "d"  { [int]$t.TotalDays }\n'
             + '    "y"  { [int]$t.TotalDays }\n'
             + '    "w"  { [int]$t.TotalDays }\n'
             + '    "ww" { [int][math]::Floor($t.TotalDays / 7) }\n'
             + '    default { throw "DateDiff interval \'$interval\' needs calendar arithmetic; convert this call by hand" }\n'
             + '  }\n}',
  // DateAdd's calendar intervals are exactly the ones a TimeSpan cannot do, and [datetime] can.
  VbsDateAdd: 'function VbsDateAdd($interval, $n, $d) {\n'
            + '  $dt = [datetime]$d\n'
            + '  switch ([string]$interval) {\n'
            + '    "yyyy" { $dt.AddYears([int]$n) }\n'
            + '    "q"    { $dt.AddMonths(3 * [int]$n) }\n'
            + '    "m"    { $dt.AddMonths([int]$n) }\n'
            + '    "y"    { $dt.AddDays([double]$n) }\n'
            + '    "d"    { $dt.AddDays([double]$n) }\n'
            + '    "w"    { $dt.AddDays([double]$n) }\n'
            + '    "ww"   { $dt.AddDays(7 * [double]$n) }\n'
            + '    "h"    { $dt.AddHours([double]$n) }\n'
            + '    "n"    { $dt.AddMinutes([double]$n) }\n'
            + '    "s"    { $dt.AddSeconds([double]$n) }\n'
            + '    default { throw "DateAdd interval \'$interval\' is not one VBScript defines" }\n'
            + '  }\n}',
  // The names VBScript's TypeName reports, for the types the converter can produce. Integer vs
  // Long follows the VALUE, as VBScript's literal typing does.
  VbsTypeName: 'function VbsTypeName($v) {\n'
             + '  if ($null -eq $v) { return "Empty" }\n'
             + '  if ($v -is [System.DBNull]) { return "Null" }\n'
             + '  if ($v -is [string]) { return "String" }\n'
             + '  if ($v -is [bool]) { return "Boolean" }\n'
             + '  if ($v -is [datetime]) { return "Date" }\n'
             + '  if ($v -is [byte]) { return "Byte" }\n'
             + '  if ($v -is [int16]) { return "Integer" }\n'
             + '  if ($v -is [int]) { if ($v -ge -32768 -and $v -le 32767) { return "Integer" } else { return "Long" } }\n'
             + '  if ($v -is [long]) { return "Long" }\n'
             + '  if ($v -is [single]) { return "Single" }\n'
             + '  if ($v -is [double]) { return "Double" }\n'
             + '  if ($v -is [decimal]) { return "Currency" }\n'
             + '  if ($v -is [array]) { return "Variant()" }\n'
             + '  if ($v -is [System.__ComObject]) { return "Object" }\n'
             + '  $v.GetType().Name\n}',
  // -1 / 0 / 1, like VBScript; mode 1 is vbTextCompare (case-insensitive), 0 is ordinal.
  VbsStrComp: 'function VbsStrComp($a, $b, $mode = 0) {\n'
            + '  $r = if ([int]$mode -eq 1) { [string]::Compare([string]$a, [string]$b, $true) } else { [string]::CompareOrdinal([string]$a, [string]$b) }\n'
            + '  if ($r -lt 0) { -1 } elseif ($r -gt 0) { 1 } else { 0 }\n}',
  VbsFormatNumber: 'function VbsFormatNumber($n, $digits = -1) {\n'
                 + '  if ([int]$digits -lt 0) { $digits = 2 }\n'
                 + '  ([double]$n).ToString("N$digits")\n}',
  // The button and icon bits are the VBScript ones; the .NET enums share the same values, and so
  // do the results (vbOK=1 ... vbNo=7).
  VbsMsgBox: 'function VbsMsgBox($prompt, $buttons = 0, $title = "") {\n'
           + '  Add-Type -AssemblyName System.Windows.Forms\n'
           + '  $btn = [System.Windows.Forms.MessageBoxButtons]([int]$buttons -band 7)\n'
           + '  $icon = switch ([int]$buttons -band 0x70) { 16 { "Error" } 32 { "Question" } 48 { "Warning" } 64 { "Information" } default { "None" } }\n'
           + '  [int][System.Windows.Forms.MessageBox]::Show([string]$prompt, [string]$title, $btn, [System.Windows.Forms.MessageBoxIcon]$icon)\n}',
  VbsDatePart: 'function VbsDatePart($interval, $d) {\n'
             + '  $dt = [datetime]$d\n'
             + '  switch ([string]$interval) {\n'
             + '    "yyyy" { $dt.Year }\n'
             + '    "q"    { [int][math]::Ceiling($dt.Month / 3) }\n'
             + '    "m"    { $dt.Month }\n'
             + '    "y"    { $dt.DayOfYear }\n'
             + '    "d"    { $dt.Day }\n'
             + '    "w"    { [int]$dt.DayOfWeek + 1 }\n'
             + '    "ww"   { [System.Globalization.CultureInfo]::InvariantCulture.Calendar.GetWeekOfYear($dt, "FirstDay", "Sunday") }\n'
             + '    "h"    { $dt.Hour }\n'
             + '    "n"    { $dt.Minute }\n'
             + '    "s"    { $dt.Second }\n'
             + '    default { throw "DatePart interval \'$interval\' is not one VBScript defines" }\n'
             + '  }\n}',
  // VBScript lets the month and day overflow (DateSerial(2020, 14, 40) is a date in 2021).
  VbsDateSerial: 'function VbsDateSerial($y, $m, $d) {\n'
               + '  ([datetime]::new([int]$y, 1, 1)).AddMonths([int]$m - 1).AddDays([int]$d - 1)\n}',
  VbsIsDate: 'function VbsIsDate($v) {\n'
           + '  if ($v -is [datetime]) { return $true }\n'
           + '  if ($null -eq $v -or "" -eq [string]$v) { return $false }\n'
           + '  $out = [datetime]0\n'
           + '  [datetime]::TryParse([string]$v, [ref]$out)\n}',
  // Measured on the host: the COM dictionary takes Int32, Int16, Double and String keys and
  // rejects UInt32 AND Int64 with CTL_E_ILLEGALFUNCTIONCALL - so an unsigned or 64-bit key
  // becomes Int32 when it fits and Double otherwise, which is what VBScript's Variant would hold.
  VbsComKey: 'function VbsComKey($k) {\n'
           + '  if ($k -is [uint16] -or $k -is [uint32] -or $k -is [uint64] -or $k -is [int64]) {\n'
           + '    if ($k -le 2147483647) { return [int]$k } else { return [double]$k }\n'
           + '  }\n  $k\n}',
  VbsBlankLines: 'function VbsBlankLines($n) { for ($i = 0; $i -lt [int]$n; $i++) { [Console]::Out.WriteLine() } }',
  VbsCByte: 'function VbsCByte($v) {\n'
          + '  $s = [string]$v\n'
          + '  if ($s -match \'^\\s*&[Hh]([0-9A-Fa-f]+)\\s*$\') { [byte][Convert]::ToInt32($Matches[1], 16) }\n'
          + '  else { [byte][math]::Round(($v -as [double]), [MidpointRounding]::ToEven) }\n}',
  // VBScript's rule is by TYPE, not by whether the text looks numeric: two strings concatenate
  // ("1" + "2" is "12"), a string and a number add (the string is converted, "17472" + 1 is
  // 17473), Empty counts as 0 or "". The first version parsed both and added, so two strings of
  // digits summed.
  VbsAdd: 'function VbsAdd($a, $b) {\n'
        + '  if ($a -is [string] -and $b -is [string]) { return $a + $b }\n'
        + '  if ($null -eq $a) { if ($b -is [string]) { return $b } else { $a = 0 } }\n'
        + '  if ($null -eq $b) { if ($a -is [string]) { return $a } else { $b = 0 } }\n'
        + '  (VbsParseNum $a) + (VbsParseNum $b)\n}',
  VbsIntDiv: 'function VbsIntDiv($a, $b) { [int][math]::Truncate(($a -as [double]) / ($b -as [double])) }',
  // Every numeric conversion goes through this, because `-as [double]` cannot read `&H1A2B` and
  // silently gives 0 - `CLng("&H" & Left(uid, 7))` put every endpoint in bin 0. VBScript's hex
  // literal is a 32-bit Long, so eight hex digits with the top bit set come back negative.
  VbsParseNum: 'function VbsParseNum($v) {\n'
             + '  if ($v -is [bool]) { if ($v) { return -1 } else { return 0 } }\n'
             + '  if ($null -eq $v) { return 0 }\n'
             + '  $s = [string]$v\n'
             + '  if ($s -match \'^\\s*&[Hh]([0-9A-Fa-f]+)\\s*&?\\s*$\') { $n = [Convert]::ToInt64($Matches[1], 16); if ($Matches[1].Length -le 8 -and $n -gt 0x7FFFFFFF) { $n -= 0x100000000 }; return $n }\n'
             + '  if ($s -match \'^\\s*&[Oo]([0-7]+)\\s*&?\\s*$\') { return [Convert]::ToInt64($Matches[1], 8) }\n'
             + '  [double]$s\n}',
  VbsCInt: 'function VbsCInt($v) { [int][math]::Round((VbsParseNum $v), [MidpointRounding]::ToEven) }',
  VbsCLng: 'function VbsCLng($v) { [long][math]::Round((VbsParseNum $v), [MidpointRounding]::ToEven) }',
  VbsCBool: 'function VbsCBool($v) { [bool]$v }',
  VbsCDbl: 'function VbsCDbl($v) { [double](VbsParseNum $v) }',
  VbsCStr: 'function VbsCStr($v) { if ($v -is [bool]) { if ($v) { "True" } else { "False" } } else { [string]$v } }',
  VbsMid: 'function VbsMid($s, [int]$start, $length) {\n'
        + '  $s = [string]$s\n'
        + '  if ($start -lt 1 -or $start -gt $s.Length) { return "" }\n'
        + '  if ($null -eq $length) { $s.Substring($start - 1) }\n'
        + '  else { $n = [math]::Min([int]$length, $s.Length - ($start - 1)); if ($n -le 0) { "" } else { $s.Substring($start - 1, $n) } }\n}',
  VbsLeft: 'function VbsLeft($s, [int]$n) { $s = [string]$s; if ($n -le 0) { "" } else { $s.Substring(0, [math]::Min($n, $s.Length)) } }',
  VbsRight: 'function VbsRight($s, [int]$n) { $s = [string]$s; if ($n -le 0) { "" } else { $s.Substring($s.Length - [math]::Min($n, $s.Length)) } }',
  VbsInStr: 'function VbsInStr($a, $b, $c) {\n'
          + '  if ($null -eq $c) { $start = 1; $hay = [string]$a; $needle = [string]$b }\n'
          + '  else { $start = [int]$a; $hay = [string]$b; $needle = [string]$c }\n'
          + '  if ($start -lt 1 -or $start -gt $hay.Length) { return 0 }\n'
          + '  $i = $hay.IndexOf($needle, $start - 1, [StringComparison]::Ordinal)\n'
          + '  if ($i -lt 0) { 0 } else { $i + 1 }\n}',
  VbsReplace: 'function VbsReplace($s, $f, $r) { ([string]$s).Replace([string]$f, [string]$r) }',
  VbsStrReverse: 'function VbsStrReverse($s) { $c = ([string]$s).ToCharArray(); [array]::Reverse($c); -join $c }',
  VbsSplit: 'function VbsSplit($s, $d = " ") { ([string]$s).Split([string]$d) }',
  VbsJoin: 'function VbsJoin($a, $d = " ") { ($a -join [string]$d) }',
  VbsUBound: 'function VbsUBound($a) { if ($null -eq $a) { -1 } else { @($a).Count - 1 } }',
  // VarType codes, per the VBScript table. Only the ones scripts actually branch on.
  VbsVarType: 'function VbsVarType($v) {\n'
            + '  if ($null -eq $v) { return 1 }\n'
            + '  if ($v -is [bool]) { return 11 }\n'
            + '  if ($v -is [int] -or $v -is [int16]) { return 2 }\n'
            + '  if ($v -is [long]) { return 3 }\n'
            + '  if ($v -is [double] -or $v -is [single]) { return 5 }\n'
            + '  if ($v -is [datetime]) { return 7 }\n'
            + '  if ($v -is [string]) { return 8 }\n'
            + '  if ($v -is [array]) { return 8192 }\n'
            + '  return 9\n}',
};

/** Emits only the helpers the converted body actually calls. */
function runtimePrelude(body) {
  // Closure over helper-to-helper calls, not a single pass. VbsAnd/VbsOr/VbsXor/VbsNot all call
  // VbsNum, and VbsNum never appears in the converted body - so a single pass emitted VbsAnd
  // without it and the script died on a missing function. Iterate until nothing new is pulled in.
  const used = new Set();
  let scan = body;
  for (let guard = 0; guard < 10; guard++) {
    const found = Object.keys(RUNTIME).filter(
      (fn) => !used.has(fn) && new RegExp(`\\b${fn}\\b`).test(scan));
    if (!found.length) break;
    found.forEach((f) => used.add(f));
    scan = found.map((f) => RUNTIME[f]).join("\n");
  }
  if (!used.size) return "";
  return "# --- VBScript runtime helpers (only the ones this script uses) ---\n"
       + "# These exist because PowerShell's obvious equivalent is WRONG, not merely uglier:\n"
       + "# VBScript's & always concatenates, its string comparison is case-sensitive, its string\n"
       + "# indexes are 1-based, and True is -1.\n"
    // Declaration order in RUNTIME, not discovery order: a helper that calls another must be
    // emitted after it, and VbsNum is discovered second while being needed first.
       + Object.keys(RUNTIME).filter((fn) => used.has(fn)).map((fn) => RUNTIME[fn]).join("\n")
    // The argument helpers need the SCRIPT'S $args, captured here at script scope. Reading $args
    // from inside the helper would give the helper's own (empty) arguments.
       + (used.has("VbsNamedArgs") || used.has("VbsUnnamedArgs")
          ? "\n$script:VbsScriptArgs = $args" : "")
    // Err has to EXIST before the first read, or Set-StrictMode throws on a script that checks
    // Err.Number before anything has thrown.
       // `\bVbsErr\b`, not `$VbsErr`: reads are emitted as `$script:VbsErr` since 2026-09-22 so
       // that a class method can see the shim, and the old substring test stopped matching them.
       + (used.has("VbsSetErr") || used.has("VbsClearErr") || /\bVbsErr\b/.test(body)
          ? '\n$script:VbsErr = [pscustomobject]@{ Number = 0; Description = ""; Source = "" }' : "")
       + "\n# --- end helpers ---\n\n";
}

/**
 * Move every top-level `function` block above the main body.
 *
 * VBScript hoists Function and Sub: a call can sit above the definition and the interpreter finds
 * it. PowerShell resolves function names at EXECUTION time, so the same layout throws "The term
 * 'X' is not recognized". Plenty of scripts put their helpers at the bottom, which is idiomatic
 * VBScript and fatal here.
 *
 * This parses cleanly, which is exactly why the 80%-of-files-parse number never saw it. A reviewer
 * running one real script found it in seconds.
 *
 * Definition ORDER among the functions is preserved, and so is the order of everything else. Only
 * the two groups move relative to each other.
 */
/**
 * Net brace count for a line, ignoring braces inside strings and comments.
 *
 * A naive count breaks hoisting: the converted output carries the original VBScript in `# TODO`
 * and `# REVIEW` comments, and a comment quoting a line with a brace in it silently shifted the
 * depth. The function block was then cut in the wrong place and its opening line was hoisted away
 * from its body, leaving an orphaned `}` that reads as "Missing function body".
 */
function braceDelta(line) {
  const bare = line
    .replace(/"(?:[^"`]|`.)*"/g, '""')
    .replace(/'[^']*'/g, "''")
    .replace(/#.*$/, "");
  return (bare.match(/\{/g) || []).length - (bare.match(/\}/g) || []).length;
}

function hoistFunctions(text) {
  const lines = text.split("\n");
  const defs = [];
  const body = [];
  let i = 0;

  while (i < lines.length) {
    // A top-level definition only. A nested one is already inside its parent's braces and moves
    // with it; matching those here would tear the parent apart.
    if (/^function\s+[\w-]+/i.test(lines[i])) {
      const block = [lines[i]];
      // Brace depth, seeded from the opening line so a `function f { ... }` one-liner closes here.
      let depth = braceDelta(lines[i]);
      let j = i + 1;
      while (j < lines.length && depth > 0) {
        depth += braceDelta(lines[j]);
        block.push(lines[j]);
        j++;
      }
      // Unbalanced means we cannot tell where it ends; leaving it in place beats moving half of it.
      if (depth !== 0) { body.push(...lines.slice(i)); break; }
      defs.push(...block, "");
      i = j;
      continue;
    }
    body.push(lines[i]);
    i++;
  }

  if (!defs.length) return text;
  return defs.join("\n") + "\n" + body.join("\n");
}

export function toPowerShell(src, opts = {}) {
  const { lines } = convertVbs(src, opts);
  const body = [];
  for (const l of lines) {
    if (l.kind === "blank") { body.push(""); continue; }
    // `rendered` means the line already carries its own TODO text in `ps` and must be emitted
    // verbatim. Without it this branch rebuilds every flagged line from rule/reason/vbs and throws
    // `ps` away - so the untranslated-intrinsic pass, which deliberately KEEPS the converted code
    // under its warning, had its output silently discarded and three separate fixes to it produced
    // byte-identical files.
    if ((l.kind === "flagged" || l.kind === "unknown") && !l.rendered) {
      body.push(`# TODO (${l.rule || "unrecognised"}): ${l.reason}`);
      body.push(`#   ${(l.orig ?? l.vbs).trim()}`);
      continue;
    }
    if (l.ps !== "") body.push(l.ps);
  }
  const text = shakeFunctions(hoistFunctions(body.join("\n")), opts);
  return runtimePrelude(text) + text;
}

/**
 * Drop top-level functions nothing reaches. A small wrapper script often pulls in a whole shared
 * include - hundreds of constants and a dozen procedures - for one call, and every unused procedure in the
 * output is a place for a bug to hide from the reviewer. Reachability starts from everything that
 * is not a function definition (the main body, classes) and follows names through the kept
 * functions; a name mentioned anywhere counts, including inside a string, so GetRef and Execute
 * keep what they name. The omitted names are listed once at the top. `opts.keepUnused` disables it.
 */
function shakeFunctions(text, opts = {}) {
  if (opts.keepUnused) return text;
  const lines = text.split("\n");
  const blocks = [];                                   // { name, start, end } of top-level functions
  for (let i = 0; i < lines.length;) {
    const m = lines[i].match(/^function\s+([\w-]+)/i);
    if (!m) { i++; continue; }
    let depth = braceDelta(lines[i]), j = i + 1;
    while (j < lines.length && depth > 0) { depth += braceDelta(lines[j]); j++; }
    if (depth !== 0) return text;                      // cannot tell where it ends; leave everything
    blocks.push({ name: m[1], start: i, end: j });
    i = j;
  }
  if (!blocks.length) return text;
  const inBlock = new Array(lines.length).fill(null);
  blocks.forEach((b, k) => { for (let i = b.start; i < b.end; i++) inBlock[i] = k; });
  const rest = lines.filter((_, i) => inBlock[i] === null).join("\n");
  const mentions = (name, s) => new RegExp(`(?<![\\w-])${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(?![\\w-])`, "i").test(s);
  const reachable = new Set(blocks.filter((b) => mentions(b.name, rest)).map((b) => b.name));
  for (let grew = true; grew;) {
    grew = false;
    for (const b of blocks) {
      if (reachable.has(b.name)) continue;
      for (const r of blocks) {
        if (!reachable.has(r.name) || r === b) continue;
        if (mentions(b.name, lines.slice(r.start, r.end).join("\n"))) { reachable.add(b.name); grew = true; break; }
      }
    }
  }
  const dropped = blocks.filter((b) => !reachable.has(b.name));
  if (!dropped.length) return text;
  const keep = lines.filter((_, i) => inBlock[i] === null || reachable.has(blocks[inBlock[i]].name));
  const note = `# ${dropped.length} procedure(s) from the source are not called by this script and were left out: `
             + dropped.map((b) => b.name).join(", ") + "\n";
  return note + keep.join("\n").replace(/\n{3,}/g, "\n\n");
}

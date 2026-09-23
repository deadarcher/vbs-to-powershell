# VBS to PowerShell

**Convert a VBScript file to PowerShell that behaves the same, not just one that looks the same.** Drop a `.vbs` file and get PowerShell back, plus a count of what still needs a human. Nothing is uploaded: the conversion runs entirely in your browser.

**Use it now (hosted):** https://getrff.com/vbs-to-powershell/

---

## Why a find-and-replace isn't enough

VBScript and PowerShell look close enough that a text swap compiles. Then it runs, finishes without an error, and gives you a different answer. These are the ones that get people, and the converter keeps VBScript's behaviour for each:

- **`&` always concatenates.** `"1" & 2` is the string `12`. Swap it for `+` and PowerShell hands you `3`.
- **`=` on strings is case-sensitive.** `-eq` isn't, so a comparison that used to fail starts passing.
- **`\` is integer division.** `7 \ 2` is `3`.
- **`True` is `-1`.** That shows up the moment it hits arithmetic or gets written to a registry value.
- **`Mid` and `InStr` count from 1.** `.Substring` and `.IndexOf` count from 0, and `InStr` returns 0 for no match where `.IndexOf` returns -1.
- **`ByRef` is the default.** In PowerShell it isn't, so a `Sub` that writes to its argument stops affecting the caller.

The output carries one small helper function per behaviour it has to keep, and only the ones your script actually uses. You get a single self-contained `.ps1` with no module to install.

## What it flags instead of converting

- **`On Error Resume Next`.** PowerShell has nothing that resumes at the next statement. `$ErrorActionPreference = "SilentlyContinue"` only quiets non-terminating errors. On a script PDQ or your RMM grades by exit code, a swallowed error lands in the console as a successful deployment.
- **`ByRef` parameters that get written to.** You get the parameter name and the line number, so you can decide between a `[ref]` and a return value.
- **Anything else it can't do faithfully** comes back as `# TODO` with your original line underneath. Thirty TODOs means thirty decisions, and it's better to see that number before you start than three days in.

It also asks how the script is deployed: logon script, RMM or PDQ push, scheduled task, or run by hand. The same line is harmless in one and a real problem in another, so the advice changes with it.

## How accurate is it

Measured on the hosted build, across 610 real sysadmin scripts pulled off GitHub: 98% of lines convert, and 94% of whole files come out as PowerShell that parses clean. A separate 168-case suite runs the VBScript and the generated PowerShell against the same inputs and compares the values, which is what catches output that runs fine and returns the wrong thing.

None of that tells you *your* script is OK. Read the output and test it somewhere you don't mind breaking.

The measurement corpus isn't in this repo. It's other people's scripts, used only to measure the converter, and it isn't ours to redistribute.

## Honest limits

- **Not a finished port.** It's a head start. Work through the TODOs before you run anything.
- **`Eval` and `Execute`** build code from strings at run time. There's no safe static translation, so they're flagged.
- **`MsgBox` and `InputBox` as statements** are flagged rather than converted, because a dialog in a script that runs unattended blocks until something kills it.
- **Encoded `.vbe` files** aren't decoded. It reads UTF-8, UTF-16 and BOMs, which covers most `.vbs` files you'll actually run into.

## Run it yourself

It's a static site. Nothing runs server-side, so any web server will do.

```
npm install
npm run dev        # http://localhost:4321
npm run build      # static files in dist/
```

Or with Docker:

```
docker run --rm -p 8080:80 ghcr.io/deadarcher/vbs-to-powershell:latest
```

or `docker compose up -d` from this folder. Then open http://localhost:8080.

## Use the engine directly

The converter is one dependency-free ES module, [`src/lib/vbsToPs1.mjs`](src/lib/vbsToPs1.mjs), so you can batch-convert a folder of scripts from Node:

```js
import { readFileSync, writeFileSync } from 'node:fs';
import { toPowerShell, auditVbs, decodeVbsFile } from './src/lib/vbsToPs1.mjs';

const buf = readFileSync('logon.vbs');
const src = decodeVbsFile(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.length));

const report = auditVbs(src, { context: 'logon' });   // logon | rmm | scheduled | interactive
console.log(report.verdict, '-', report.detail);
console.log(report.blockers.length, 'lines need a decision');

writeFileSync('logon.ps1', toPowerShell(src, { context: 'logon' }));
```

`decodeVbsFile` matters more than it looks: most `.vbs` files written by a Windows tool are UTF-16LE with a BOM, and reading one as UTF-8 gives you a NUL between every character that converts to nothing at all.

## Privacy

Your script is read and converted in the browser. There's no upload, no storage and no account, and this build loads nothing from a third party: no analytics, no font CDN. Open your browser's developer tools and watch the network tab if you want to check.

## Issues

If it converts something wrong, open an issue with the smallest VBScript that shows it and what you expected. A snippet that runs under `cscript` and prints a value is the most useful thing you can send, because that's exactly how the test suite checks it.

The engine file is kept byte-identical with the hosted copy at getrff.com, so fixes land in both.

## License

MIT. See [LICENSE](LICENSE).

---

Built by the [RFF](https://getrff.com) team. RFF is a Windows RMM: deploy, patch and remote-control your fleet from a browser tab, free for your first 100 endpoints.

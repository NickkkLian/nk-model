# nk-model

An agent skill for [Claude Code](https://code.claude.com) and [OpenAI Codex](https://developers.openai.com/codex). Turn a description of a small business and a few assumptions into a three-statement financial model in one .xlsx whose statements tie to each other; the builder writes nothing unless the balance sheet balances.

**What you get.** The income statement of the workbook `make_model.py` writes from its bundled example, an invented bike-repair shop, shown through `preview.py`. Every number is a formula, and the builder checked that the balance sheet balances to the penny. Recorded on 2026-09-30 with 0.1.7.

![nk-model: the income statement of a three-year model, every number a formula](https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/results/nk-model.png)

## Try it

Nothing is installed and nothing under `~/.claude` changes: clone, run the self-tests, run the example. It writes only `demo*` files inside the clone.

```bash
git clone https://github.com/NickkkLian/nk-model && cd nk-model
python3 scripts/make_model.py --selftest
python3 scripts/model_check.py --selftest
python3 scripts/numsrc.py --selftest
python3 scripts/ooxml.py --selftest
python3 scripts/preview.py --selftest
python3 scripts/xlformula.py --selftest
python3 scripts/make_model.py --init demo.json
python3 scripts/make_model.py demo.json -o demo-model.xlsx
python3 scripts/model_check.py demo-model.xlsx
python3 scripts/preview.py demo-model.xlsx -o demo-preview.html
```

Each self-test ends on its own line:

```text
make_model selftest: 21/21 passed
model_check selftest: 45/45 passed
selftest: 50/50
ooxml selftest: 13/13 passed
preview selftest: 16/16 passed
xlformula selftest: 14/14 passed
```

The example commands print this (recorded in a fresh copy with an empty home folder; the path of the clone is taken out):

```text
$ python3 scripts/make_model.py --init demo.json
wrote demo.json: an illustrative bike-repair business to edit
$ python3 scripts/make_model.py demo.json -o demo-model.xlsx
wrote demo-model.xlsx: Harbour Bike Repairs, 2027-2029 · 15,412 bytes · balanced to the penny in all 4 balance sheets
net profit 7,614, 15,452, 24,545 · cash at the end of each year 14,400, 26,518, 47,268 (£)
next: python3 scripts/model_check.py demo-model.xlsx
$ python3 scripts/model_check.py demo-model.xlsx
✔ demo-model.xlsx: 0 findings · balanced to the penny, every formula's stored value recomputed
$ python3 scripts/preview.py demo-model.xlsx -o demo-preview.html
wrote demo-preview.html: 6 sheets; open it, or add #s2, #s3 ... to the address for the others
```

Open `demo-model.xlsx` in a spreadsheet program, or `demo-preview.html` in a browser: its second sheet is the picture above.

### What to type

With the skill installed ([Install](#install)), ask in plain words. This is the request a recorded test run used; it never names the skill:

> My sister wants to open a small dog-grooming salon: about 18 dogs a week at first at £45 a groom, rent and a part-time assistant come to roughly £2,800 a month, and she's borrowing £12,000 over three years at around 9%. Can you put together a proper three-year financial model she can open in Excel — profit, cash and a balance sheet — so she can see whether the money works?

![nk-model](https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/social/nk-model.png)

Part of [nickkk-skills](https://github.com/NickkkLian/nickkk-skills) — agent skills that ship a self-test with every script; the Verify
section below says which of them were broken on purpose before release to prove they react.

![nk-model demo: a few assumptions in, one workbook out](https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/nk-model.gif)

The demo above was recorded with another invented business, a coffee cart; `--init` writes the bike-repair shop shown at the top of this page.

## What it does

- You write the business down as JSON (`make_model.py --init` gives an example to edit); the builder writes one workbook with six sheets: Assumptions, Income statement, Balance sheet, Cash flow, Checks and Charts.
- Every number on the statements is a formula, never a plug, and its value is stored beside it, so a viewer that does not recalculate still shows what the formulas give.
- The builder works the model out twice, through the workbook's formulas and in plain Python, and writes nothing unless the two agree and the balance sheet balances to the penny.
- `model_check.py` re-checks any workbook this skill wrote, including one edited since: eleven rules, from plugs and stored values to the two charts.
- `preview.py` shows the workbook as one HTML page, for looking at it without a spreadsheet program; click a number to see its cell and its formula.
- Standard library only, Python 3.9+. No Chrome and no spreadsheet program needed.

<details><summary>What each script does, in full</summary>

- Five invariants: every number on the statements is a formula, never a plug; the balance sheet balances to the penny at the start and in every year, recomputed from its lines rather than its totals; the three statements tie; the numbers follow the method the Assumptions sheet states; every formula's value is stored beside it, so a viewer that does not recalculate shows the numbers the formulas give.
- `scripts/make_model.py`: a business written down as JSON (`--init` writes an example to edit) becomes one workbook with six sheets — Assumptions, Income statement, Balance sheet, Cash flow, Checks, Charts. It refuses a business with a missing or out-of-range number (no amount above a billion), text with a control character or an emoji, or real numbers with no source; works the model out twice (through the workbook's formulas, and in plain Python without them); and writes nothing unless they agree and it balances.
- `scripts/model_check.py`: eleven rules on any workbook this skill wrote, including one edited since. It builds again the workbook this skill would write for the same years and compares every cell — the layout, word for word and format for format, with no cell filled that it leaves empty; no plugs, so every statement cell holds exactly this skill's formula and nothing is typed where it writes nothing; every stored value recomputed from its formula; the balance sheet adding up and balancing; the statements tying; the inputs within what the builder accepts and the numbers following the stated method; the Checks sheet telling the truth; the honesty sentence; no macros, embedded objects or links out of the file; the two charts drawing the right cells, with copies of the numbers that match; no emoji. Cash below zero by more than half a penny, at the start or at a year end, is reported as a warning. What it does not look at is under Limits.
- `scripts/preview.py`: the workbook as one HTML page — every sheet as stored, and the charts drawn from the numbers they carry — for looking at it without a spreadsheet program. The page carries the three faces the workbook names: the Latin subsets of Fraunces, Inter and Space Mono (SIL OFL 1.1, licences in `assets/fonts/`) are inlined, which adds 135,972 bytes to it. It still fetches nothing.
- `scripts/ooxml.py` and `scripts/xlformula.py`: the standard-library .xlsx writer and reader, and the checker's evaluator for the part of Excel's formula language the model uses (Excel's own order: `2^3^2` is 64, `-2^2` is 4; some things it could not test against Excel, such as `&` or `TRUE` typed into `MAX`, it refuses; where it is known to differ from Excel is under Limits).
- Standard library only, Python 3.9+. No Chrome, no spreadsheet program needed.
- In the HTML preview, click any number to see its cell, its formula and what the preview did not check (the shared number-sources layer, `numsrc.py`).

</details>

The full procedure, the boundaries and where the rules came from are in [SKILL.md](SKILL.md).

## How it works

1. Write the business down: `python3 scripts/make_model.py --init business.json` gives a filled-in example to edit (every field is described in `references/business-format.md`).
2. Say whether the numbers are real.
3. Build the workbook: `python3 scripts/make_model.py business.json -o model.xlsx`.
4. Check it: `python3 scripts/model_check.py model.xlsx` reads the file back and checks eleven things.
5. Read the cash line. If cash goes below zero (by more than half a penny) at the start or at a year end, both scripts say so: the plan needs money the model does not include.
6. Look at it without a spreadsheet if you cannot open one: `python3 scripts/preview.py model.xlsx -o preview.html` shows every sheet as stored, and the charts from the numbers they carry.
7. Hand it over with the three facts that matter: whether it balances (it always will, or it would not exist), when the cash is lowest, and what the model leaves out (it says on the Assumptions sheet).

## Why it is built this way

**The idea.** A business plan is three statements that have to agree. A model that does not balance has a mistake in it somewhere, and a model that balances only because someone typed a number into the balance sheet (a plug) is worse, because it looks right.

**Where it came from.** The discipline comes from a bank-statement categoriser of mine: totals that must reconcile to the penny, output that is rebuilt and compared rather than trusted, and a deliberate break that has to turn a check red before the check counts.

**Evidence.** What was broken on purpose to show that the self-tests can fail is under [Verify](#verify); what was run end to end, and in which agent, is under [Compatibility](#compatibility).

## Install

Pick one of four ways: three for Claude Code, one for OpenAI Codex. Skills load when a session starts, so open a **new** session after installing.

### 1 · Terminal, one command

```bash
git clone https://github.com/NickkkLian/nk-model ~/.claude/skills/nk-model
```

1. Run the command above (for one project only, clone into `.claude/skills/nk-model` inside that project).
2. Start a new Claude Code session.
3. Check it loaded: type `/nk-model` — it appears in the slash-command menu. Or just ask for the task; the skill triggers on its own.

### 2 · Claude Code in a terminal session (plugin)

The plugin route goes through the [nickkk-skills](https://github.com/NickkkLian/nickkk-skills) marketplace. Add it once; after that each skill is one command.

```
/plugin marketplace add NickkkLian/nickkk-skills
/plugin install nk-model@nickkk-skills
```

1. In a Claude Code session, run the first line (once per machine).
2. Run the second line.
3. Start a new session (or run `/reload-plugins`). The skill shows up as `nk-model:nk-model`.

Without opening a session, the same two steps work from a shell: `claude plugin marketplace add NickkkLian/nickkk-skills` then `claude plugin install nk-model@nickkk-skills`.

### 3 · Claude desktop app (Code tab)

**Add the marketplace first — Discover only searches marketplaces you have already added.**

<img src="https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/panel-route/panel-route.gif" alt="Adding the marketplace and installing a skill in the desktop app" width="640">

<sub>Recorded on 2026-09-16, when the marketplace listed ten skills, all at version 0.1.0; it lists more now. The repository list in this recording shows the recorder's own repositories because a GitHub account is connected; yours will show yours. Type the full name as in step 4.</sub>

1. In the chat box, type `/plugin marketplace` and press Enter (or open **Settings → Customize → Plugins**). The **Plugins** panel opens.
   <br><img src="https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/panel-route/step1-type-plugin-marketplace.png" alt="/plugin marketplace typed in the chat box" width="480">
2. Top right, open **Add ▾** and choose **Add marketplace**.
   <br><img src="https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/panel-route/step2-add-menu.png" alt="The Add menu with Add marketplace" width="480">
3. Choose **Add from a repository**.
   <br><img src="https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/panel-route/step3-add-from-repository.png" alt="Add marketplace dialog: Add from a repository" width="480">
4. In **URL**, type the full `NickkkLian/nickkk-skills`. At the bottom of the list choose the row **Use "NickkkLian/nickkk-skills"**, then press **Sync**.
   <br><img src="https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/panel-route/step4-url-then-sync.png" alt="URL filled in, Sync button" width="480">
5. You land on **Discover**, filtered to the new marketplace (**Filter · 1**). Find **Nk model** and press **Add**. Installed ones show **✓ Added**.
   <br><img src="https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/panel-route/step5-discover-add.png" alt="Discover list with Added and Add buttons" width="480">
6. Close the panel and start a new session.

To try it for one session without installing anything: `claude --plugin-dir ./nk-model` from a clone.

### 4 · OpenAI Codex CLI

```bash
git clone https://github.com/NickkkLian/nk-model.git ~/.agents/skills/nk-model
```

1. Run the command above (for one project only, clone into `.agents/skills/nk-model` inside that project).
2. Start a new Codex session.
3. Check it loaded, without spending a model call: `codex debug prompt-input | grep -o -- '- nk-model[a-z0-9:-]*' | sort -u` prints `- nk-model:nk-model:`. Codex adds the `nk-model:` prefix because this repository also carries a Claude Code plugin manifest. Ask for the task and the skill triggers on its own, or type `$` and pick it from the list.

## Compatibility

<details><summary>Which agents it was run in, and what each run did</summary>

| Agent | Tested | What was checked |
|---|---|---|
| Claude Code (CLI 2.1.173, macOS) | yes | In a fresh project with an isolated Claude config, inside a macOS sandbox that blocked reading the tester's ~/.claude folder (settings, session history, memory), Desktop, Documents and Downloads, SSH keys and git identity, a plain request that never names the skill triggered it and it ran its bundled script. The route 2 plugin commands were also run from a shell with an isolated config: marketplace add, install, list. The brief, invented for the test, was a paragraph about a sister opening a dog-grooming salon, with the numbers she knew (18 dogs a week at £45, about £2,800 a month in rent and help, £12,000 borrowed over three years at about 9%), asking for a three-year model to open in Excel; it never named the skill. In eleven turns the run took the example with --init, read the format reference, wrote the business (900 grooms in the first year growing 15% a year, £4 of consumables a groom, £8,000 of equipment, and £5,000 of her own money that the brief never mentioned; its reply listed each as an assumption to check), built the workbook, checked it (balanced to the penny, 0 findings) and wrote a preview for looking at it without Excel. Its business, built again with the first released version (0.1.0), balances, and the checker gives 0 findings. |
| OpenAI Codex CLI (0.155.0-alpha.9.2, gpt-5.6-sol, low reasoning, macOS) | yes | Copied into `~/.agents/skills` of a temporary home, in a fresh project, without the user's Codex config, with the same brief. Codex read SKILL.md and the format reference, started from the --init example, built the workbook, checked it and wrote the preview. Its version assumed none of her own money, so cash dips to -£289 in the first year, and Codex's reply said the plan needs at least another £300. The run's own record does not show the checker's output (the command's output came back empty). Its business, built again with the first released version (0.1.0), balances; the checker gives 0 findings and warns that cash is below zero (-289.40). |
| Cursor, Gemini CLI | no | Not tested. Their documentation says both read `~/.agents/skills`, the folder route 4 clones into; Gemini CLI asks before it activates a skill. |

Route 4 was checked for this repository: cloned from GitHub into a temporary home's `~/.agents/skills`, it was listed by the step 3 command. This skill's frontmatter uses only name, description, license, compatibility and metadata.

</details>

## Verify

```bash
python3 scripts/make_model.py --selftest
python3 scripts/model_check.py --selftest
python3 scripts/numsrc.py --selftest
python3 scripts/ooxml.py --selftest
python3 scripts/preview.py --selftest
python3 scripts/xlformula.py --selftest
```

<details><summary>What was broken on purpose before release, script by script</summary>

Python 3.9+, standard library only. 31 of model_check.py's checks were broken on purpose, one at a time,
in a sandbox copy, each turning the sample written for it red, none by a crash. The 31 lines of model_check.py that report
something were read from the source, not listed by hand: 2 only pass on what other lines found, and every other one is in the matrix.
make_model.py, preview.py, ooxml.py and xlformula.py have self-tests but no break matrix.

One workbook was also opened in Microsoft Excel, by a person, once: on 2026-09-23, Microsoft Excel for Mac 16.90.2
(build 16.90.24102719) opened the demo's workbook, the coffee cart in the GIF above as an earlier version built it
(SHA-256 0998f7164d3cce07c59f2a48321910922a69a296d5a1ddb07940cb477c110515), without a repair prompt, and with the file
set to recalculate on opening, the Checks sheet read "Balanced: every check is zero to the penny" and the cash line
read "No". That was confirmed by looking at the screen of one Mac; no screenshot was kept. The coffee cart this
version builds (the one in the GIF) differs from that file in one cell only, the cash line's formula (Checks!B11: "below
zero" became "below zero by more than half a penny"); both store "No". This version's file has not been opened in
Excel, and the automated checks above have never run in Excel.
numsrc.py, the number-sources layer shared with four other skills: each of its 16 lines that report a
finding was disabled in a sandbox copy, found by reading the source rather than listed by hand, and its self-test went
red each time; the unmutated copy stayed green.

</details>

## Limits

<details><summary>What it does not do, in full</summary>

- **Yearly, one to ten years, one product line.** Units, a price, a cost per unit and fixed costs, each growing at one rate. No months, no seasons, no product mix.
- **One way of doing each thing**: reducing-balance depreciation on equipment, working capital by days, one loan repaid in equal yearly amounts with interest on the opening balance, tax only on a profit with no losses carried forward, nothing paid out to the owners. Anything else means editing the formulas, and then `model_check.py` will say, correctly, that they are no longer this skill's formulas (K02); depending on the change it also says that the statements no longer tie (K05, as when money is paid out to the owners) or that the numbers no longer follow the stated method (K06, as with straight-line depreciation). Changing the inputs on the Assumptions sheet is the edit the model is for.
- **Not advice, and not a forecast of anything.** It shows how the assumptions fit together; whether the assumptions are right is the person's to judge.
- **Checked by reading the file, not by Excel.** The checks here read the workbook back with this skill's own reader, and the files were also opened by macOS Quick Look and by openpyxl; they have not been opened in Microsoft Excel by these checks. Edits are judged cell by cell, as the file stores them; a workbook saved again by a spreadsheet program has not been tested, and may be reported as edited where only its storage differs. A shared formula (how Excel stores a filled-down run) is read as the formula it stands for when it is stored the usual way; in a group where more than one cell carries formula text, this reader and openpyxl can expand a cell differently, so such a file can pass here while openpyxl reads other formulas. Fonts named in the file (Fraunces, Inter, Space Mono) fall back to the viewer's own where they are not installed.
- **What the checker does not look at.** It compares the cells and their number formats, what the two charts draw and their copies of the numbers, and the package's links, macros and embedded objects. An edit anywhere else passes, even one that changes what a reader sees: a conditional format that shows a number as fixed text, a third chart, a text box laid over a number, a chart's title or its copy of the year labels, hidden rows or columns, merged cells, a font colour, a data connection such as a web query (`xl/connections.xml`), text moved into a phonetic guide (read here as part of the label; openpyxl leaves it out), Excel's "precision as displayed" setting, or the note under a sheet's title, which is free text like the business's name. Open the file and look before it goes anywhere (`references/acceptance.md`).
- **Where its arithmetic is not Excel's.** The formula evaluator follows Excel's order of operations and refuses what it could not test (`&`, `ROUND`, `TRUE` or text typed into `SUM`, `MIN` or `MAX`, a range inside `AND` or `OR`). It does not refuse four cases where its answer differs from what Microsoft documents for Excel, or may: a sum that comes out a hair from zero (`1.333 + 1.225 - 1.333 - 1.225` is -2.2e-16 here; Excel 97 and later show 0); text in arithmetic (`"1"+"2"` is an error here, 3 in Excel); text reached through a reference inside `AND` or `OR` (an error here; Excel ignores it); and a TRUE worked out inside `SUM`, `MIN` or `MAX` (`SUM(1=1)` is 0 here; Microsoft's pages do not say). None was tried in Excel. Only the first can come from this skill's own formulas.
- **Text XML cannot hold.** The builder refuses control characters and emoji, but not U+FFFE or U+FFFF: a name or description holding one is written into a workbook that is not valid XML, which `model_check.py` and openpyxl cannot read. Nor are ⤴ and ⤵ refused, though Unicode lists them as emoji.
- **Very large sums.** Every input can be within its range and the sums still reach the trillions (large amounts growing by hundreds of percent a year), where floating point is not exact to the penny. The builder writes such a workbook, and the checker, which adds and compounds its own way, can then report a difference at the last penny (K05 or K06).
- **Rounded for display.** The statements show money in whole units, and prices, unit costs and the check lines to the penny; every total is worked out before rounding, so a total can be a few units off the sum of the lines it shows (at most half a unit for each number in the sum, the total included), and units, which grow by a percentage, are not whole numbers. Each statement's note says it is rounded, and the Assumptions sheet's method says the rest. A year-end cash of exactly nothing can be stored as -0.0000000000036, which Excel may show as "(0)", since its format picks the negative section by the sign of the stored value (not tried in Excel).
- **The checker knows this skill's layout.** It checks workbooks this skill wrote; a model from anywhere else is reported as not from this skill (K01), not checked.

</details>

## License

MIT. Read a script before letting it run in your environment.

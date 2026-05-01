# Installing the MS Project Field Mapper (Phase 0 build)

This is the early-phase install guide. Phase 0 ships only the foundation: a
`Run Diagnostics` button that writes a `.txt` with your host, version, and
macro-security state. The full mapping pipeline is added in later phases.

## Requirements

- **Microsoft Project Desktop** 2016 or later (Project 2016, 2019, 2021, 2024,
  or Microsoft 365 desktop subscription).
- **Not supported:** Project for the Web, Project Online, or any
  browser-based Project. Those hosts do not run VBA at all.
- **Reference:** `Microsoft Scripting Runtime` (added during install).

## 1. Unblock the bundle (Mark-of-the-Web)

If you downloaded the `.zip` from a browser, OneDrive, SharePoint synced
folder, or email:

1. Right-click the `.zip` -> **Properties**.
2. At the bottom of the General tab, tick **Unblock** if shown -> **OK**.
3. Extract the zip after unblocking. (Unblocking the zip is much faster than
   unblocking each `.bas` / `.cls` / `.frm` individually after extraction.)

If MOTW is not stripped, VBE may silently refuse to import the modules, or
Project may refuse to load forms.

## 2. Open the VBA Editor

1. Launch Microsoft Project. Open any project (or just a blank one).
2. Press **Alt + F11** to open the VBA Editor (VBE).
3. In the Project Explorer pane (left), locate `VBAProject (Project1.mpp)` (or
   whatever file you have open) -- this is where the modules will be imported.

## 3. Add the Scripting Runtime reference

The tool uses `Scripting.Dictionary`, which requires the Microsoft Scripting
Runtime reference.

1. In the VBE: **Tools -> References...**
2. Scroll to **Microsoft Scripting Runtime** -> tick the checkbox -> **OK**.
3. If you can't find it, you can leave this step -- our code uses
   `CreateObject("Scripting.Dictionary")` (late-bound), so the reference is
   helpful for IntelliSense but not strictly required.

## 4. Import the modules

In the VBE, for each file in `src/`:

1. **File -> Import File...**
2. Navigate to the `src/` folder.
3. Select the file -> **Open**.

Import in this order (dependencies cascade):

1. `modConstants.bas`
2. `modHash.bas`
3. `modEnvironment.bas`
4. `clsLogEntry.cls`
5. `clsLogger.cls`
6. `clsFieldMetadata.cls`
7. `clsProjectScan.cls`
8. `clsRunContext.cls`
9. `modFileScanner.bas`
10. `modMain.bas`
11. `frmMain.frm`      *(see step 5 if this fails)*
12. `frmProgress.frm`  *(see step 5 if this fails)*

> **Note:** When you import `.frm` files, VBE looks for matching `.frx`
> binary resource companions. This bundle does not ship any -- the forms
> have no embedded images or initialised list values, so a `.frx` isn't
> required. **If a `.frm` fails to import for any reason, follow step 5
> for that form.**

## 5. (Fallback) Manual form construction

If a `.frm` will not import on your machine, follow the matching guide:

**5a. `frmMain` (Phase 0)**

1. In the VBE: **Insert -> UserForm**. A blank `UserForm1` appears.
2. With the form selected, in the Properties window:
   - **Name** = `frmMain`
   - **Caption** = `MS Project Field Mapper`
   - **Width** ~ `450`, **Height** ~ `230`
   - **StartUpPosition** = `1 - CenterOwner`
3. From the Toolbox, drop the following controls onto the form. Set each
   control's **Name** to the value below (the rest is cosmetic):

   | Control type   | Name              | Caption                         |
   |----------------|-------------------|---------------------------------|
   | Label          | `lblTitle`        | (set in code)                   |
   | Label          | `lblPhase`        | (set in code)                   |
   | Label          | `lblHint`         | (set in code)                   |
   | Label          | `lblOutputCaption`| (set in code)                   |
   | TextBox        | `txtOutputFolder` |                                 |
   | CommandButton  | `btnBrowseOutput` | `Browse...`                     |
   | CommandButton  | `btnDiagnostics`  | `Run Diagnostics...`            |
   | CommandButton  | `btnClose`        | `Close`                         |

4. Double-click the form (any blank area) to open the code-behind.
5. Open `src/frmMain.code.bas` in any text editor. Copy the block between
   `BEGIN PASTE` and `END PASTE`. Paste it as the entire code-behind for
   `frmMain`, replacing whatever VBE put there.
6. Save the project (**Ctrl+S**). The form is now functionally identical to
   the .frm import path.

**5b. `frmProgress` (Phase 1)**

1. **Insert -> UserForm**. Set Properties:
   - **Name** = `frmProgress`
   - **Caption** = `Working...`
   - **Width** ~ `480`, **Height** ~ `320`
   - **StartUpPosition** = `1 - CenterOwner`
2. Add controls (Name property set exactly as listed):

   | Control type   | Name         | Notes                                 |
   |----------------|--------------|---------------------------------------|
   | Label          | `lblPhase`   | full-width across top                 |
   | Label          | `lblPercent` | top-right, ~40 wide, right-aligned    |
   | Label          | `pgrFrame`   | thin track; BorderStyle = 1 (single)  |
   | Label          | `pgrFill`    | inside `pgrFrame`; BackColor any contrast colour |
   | ListBox        | `lstLog`     | large central area; IntegralHeight = False |
   | CommandButton  | `btnCancel`  | bottom-left; Caption = `Cancel`       |
   | CommandButton  | `btnClose`   | bottom-right; Caption = `Close`; Enabled = False |

3. Double-click the form, replace the entire code-behind with the block
   between `BEGIN PASTE` and `END PASTE` in `src/frmProgress.code.bas`.

## 6. Verify the install (Phase 0 smoke test)

1. In the VBE, open the Immediate Window (**Ctrl+G**).
2. Type: `modMain.ShowMain` and press **Enter**.
3. The `frmMain` window should appear.
4. Click **Run Diagnostics...**
5. A confirmation dialog tells you the path of the written `.txt`.
6. Click **Yes** to open the folder; review the file. You should see:
   - Application name, version, major version
   - Office bitness (x86 or x64), OS bitness
   - AutomationSecurity state
   - Trust VBOM access flag
   - Trusted Locations list (or "(none configured)")
   - "Host supported: YES" at the bottom

If the diagnostics file is missing fields or shows "Host supported: NO", send
the file back to the maintainer for triage.

## 7. Verify the file scanner (Phase 1 smoke test)

1. Have a `.mpp` file ready -- ideally one with at least a couple of named
   custom fields (e.g. set Text1's alias to "Room" via *Add Column*, type
   some values).
2. **Alt+F8** -> select `Phase1_TestScan` -> **Run**.
3. Paste the path to your `.mpp` -> **OK**.
4. Leave the output folder blank (defaults to `%TEMP%`) -> **OK**.
5. Wait for the scan. On a small file this is sub-second; on a 50K-task
   schedule it can take 30-60 seconds. The current Phase 1 build does not
   show a live progress UI for the single-file scan -- that's wired up in
   Phase 2.
6. The success dialog reports the dump path. **Yes** opens the folder.
7. Open the dump file. You should see, at the top:
   - File path, fingerprint, source `Application.Version`
   - Task and Resource counts
   - A list of every populated or aliased custom field, one per line, with
     internal name, data type, scope, alias, populated ratio, and a sample
     of values.
   - Any duplicate aliases flagged.
   - The full session log appended at the bottom for triage.

For the same scan via the Immediate Window:

```
?modMain.Phase1_ScanToFile("C:\Projects\PlantA.mpp")
```

Returns the dump path.

### What "good" looks like

If your file has Text1 aliased "Room" with values "Room 101", "Room 102",
the dump line for that field should look like:

```
Text1   Text   Task   alias="Room"   pop=100% (3/3)   samples=[Room 101 | Room 102 | Room 103]
```

If aliases come back empty when you know they're set, or counts look wrong,
capture the dump file and send it back -- the structured log at the bottom
will show which API call failed.

### Induce-error checks

To confirm the cleanup path works, try:

- A path that doesn't exist -> log entry: `WARN ... File not found`, returns Nothing.
- A path to a non-`.mpp` file (e.g. a `.txt`) -> log entry: `ERROR ... FileOpenEx failed`, cleanup restores Application state.
- A `.mpp` that is open in another Project window -> log entry: `WARN ... File locked`, scan skipped.

After each induced error, re-run a clean scan to confirm Project's state is
not poisoned (no zombie windows; `Application.Calculation` is restored).

## 8. Optional: bind to a button

Inside Microsoft Project: **File -> Options -> Customize Ribbon** (or
**Customize the Quick Access Toolbar**) -> in **Choose commands from** pick
**Macros** -> select `modMain.ShowMain` -> **Add** -> **OK**.

## Known limitations (Phase 0 + 1)

- The scanner runs single-file only. Multi-file orchestration with cancel
  and continue-on-error lands in Phase 2.
- Field-value extraction is read-only -- the scanner never writes to source
  files. (Merge mode in Phase 7 will create copies, but Phase 1 only reads.)
- The auto-mapper, manual mapping UI, collision handling, extraction, merge,
  and persistence are not wired up yet.
- If your environment has macros globally disabled
  (`Disable all macros without notification`), you cannot run this tool. Talk
  to your IT team about Trusted Locations or signed templates.

## Troubleshooting

| Symptom | Fix |
|---|---|
| "Compile error: User-defined type not defined" on `clsLogger` etc. | You skipped the Scripting Runtime reference (Step 3) **and** are using the early-bound form. We use late binding, so this should not happen -- re-import the affected `.cls` from `src/`. |
| Diagnostics button does nothing | Open the Immediate Window, run `?modMain.RunDiagnostics()`. The path will be returned. If a runtime error appears there, copy the error number and message. |
| `.frm` import error: "Could not load some objects..." | Use the fallback in step 5 for that form. |
| Diagnostics or scan file is empty | You hit a write error; check the path returned in the message box. Make sure the folder is writable. |
| `Phase1_TestScan` reports "FileOpenEx failed (1101): ..." | The file is open in another Project window. Close it there, retry. |
| Scan produces zero fields for a file you know has custom fields | Check the bottom of the dump file -- the session log lists each slot resolution. If `Slot resolve failed` appears, the field name format may differ in your Project version. Capture the log and report. |
| Scan takes more than a minute on a small file | Set `ctx.Logger.MinSeverity = sevDebug` (already set in `Phase1_ScanToFile`) and look at the `ScanRows` debug entries in the log -- they record per-scope row counts and timings. |

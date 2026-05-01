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
2. `modEnvironment.bas`
3. `clsLogEntry.cls`
4. `clsLogger.cls`
5. `clsRunContext.cls`
6. `modMain.bas`
7. `frmMain.frm`  *(see step 5 if this fails)*

> **Note:** When you import `frmMain.frm`, VBE looks for a matching
> `frmMain.frx` (binary form resources). This bundle does not ship one --
> the form has no embedded images or initialised list values, so a `.frx`
> isn't required. **If the import fails for any reason, follow step 5.**

## 5. (Fallback) Manual form construction

If `frmMain.frm` will not import on your machine:

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
the file back to the maintainer for triage. Phase 0's whole point is to make
that triage easy.

## 7. Optional: bind to a button

Inside Microsoft Project: **File -> Options -> Customize Ribbon** (or
**Customize the Quick Access Toolbar**) -> in **Choose commands from** pick
**Macros** -> select `modMain.ShowMain` -> **Add** -> **OK**.

## Known limitations (Phase 0)

- Phase 0 only includes diagnostics. The mapping/extraction/merge pipeline is
  not yet wired up.
- If your environment has macros globally disabled
  (`Disable all macros without notification`), you cannot run this tool. Talk
  to your IT team about Trusted Locations or signed templates.

## Troubleshooting

| Symptom | Fix |
|---|---|
| "Compile error: User-defined type not defined" on `clsLogger` etc. | You skipped the Scripting Runtime reference (Step 3) **and** are using the early-bound form. We use late binding, so this should not happen -- re-import the affected `.cls` from `src/`. |
| Diagnostics button does nothing | Open the Immediate Window, run `?modMain.RunDiagnostics()`. The path will be returned. If a runtime error appears there, copy the error number and message. |
| `frmMain.frm` import error: "Could not load some objects..." | Use the fallback in step 5. |
| Diagnostics file is empty | You hit a write error; check the path returned in the message box. Make sure the folder is writable. |

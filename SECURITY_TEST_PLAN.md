# Security Test Plan — CreoRegenAppVB

**Version:** 1.0  
**Date:** March 24, 2026

---

## Table of Contents

1. [CRITICAL #1 — COM ProgID Whitelist](#critical-1--com-progid-whitelist)
2. [CRITICAL #2 — CSV Injection Prevention](#critical-2--csv-injection-prevention)
3. [CRITICAL #3 — Export Path Validation](#critical-3--export-path-validation)
4. [CRITICAL #4 — Path Traversal in Model Handling](#critical-4--path-traversal-in-model-handling)
5. [CRITICAL #5 — Information Disclosure in Error Messages](#critical-5--information-disclosure-in-error-messages)
6. [HIGH #1 — BOM Recursion Depth & Node Limits](#high-1--bom-recursion-depth--node-limits)
7. [HIGH #2 — UI-Layer Model Input Validation](#high-2--ui-layer-model-input-validation)
8. [HIGH #3 — ModelNameParser Extension Whitelist](#high-3--modelnameparser-extension-whitelist)
9. [HIGH #4 — Error Handling in Event Handlers](#high-4--error-handling-in-event-handlers)
10. [Functional Regression Tests](#functional-regression-tests)

---

## CRITICAL #1 — COM ProgID Whitelist

**File:** `CreoSessionService.vb`  
**Controls:** `AllowedProgIds`, `GetAllowedProgIdCandidates`, `IsSafeCreoClassName`, `CreatePfcObject`

### Test Cases

| ID | Input | Expected Result | Pass |
|----|-------|-----------------|------|
| C1-01 | `CreatePfcObject("pfcAsyncConnection")` | Returns COM object (or throws COM-not-registered if Creo not installed). No security exception. | [ ] |
| C1-02 | `CreatePfcObject("pfcModelDescriptor")` | Returns COM object or COM-not-registered. | [ ] |
| C1-03 | `CreatePfcObject("CpfcRegenInstructions")` | Returns COM object or COM-not-registered. | [ ] |
| C1-04 | `CreatePfcObject("MpfcModelItem")` | Returns COM object or COM-not-registered. | [ ] |
| C1-05 | `CreatePfcObject("pfcModelItem")` | Returns COM object or COM-not-registered. | [ ] |
| C1-06 | `CreatePfcObject("Windows.Shell.BrokerServiceClass")` | Throws `ArgumentException` — not in allowed list. | [ ] |
| C1-07 | `CreatePfcObject("Evil.Object")` | Throws `ArgumentException` — not in allowed list. | [ ] |
| C1-08 | `CreatePfcObject("Scripting.FileSystemObject")` | Throws `ArgumentException` — contains `.` (fails `IsSafeCreoClassName`). | [ ] |
| C1-09 | `CreatePfcObject("")` | Throws `ArgumentException`. | [ ] |
| C1-10 | `CreatePfcObject("   ")` | Throws `ArgumentException`. | [ ] |
| C1-11 | `CreatePfcObject("pfcAsync;Connection")` | Throws `ArgumentException` — `;` fails `IsSafeCreoClassName`. | [ ] |
| C1-12 | `CreatePfcObject("pfcAsync Connection")` | Throws `ArgumentException` — space fails `IsSafeCreoClassName`. | [ ] |

---

## CRITICAL #2 — CSV Injection Prevention

**File:** `MainForm.vb`  
**Control:** `EscapeCsvValues`

### Test Cases

| ID | Cell Value | Expected Output (inside quotes) | Pass |
|----|------------|--------------------------------|------|
| C2-01 | `=cmd\|"/c calc"!A1` | `'=cmd\|"/c calc"!A1` (leading `'` prepended) | [ ] |
| C2-02 | `+2+5+cmd\|"/c powershell"!A1` | `'+2+5+cmd\|"/c powershell"!A1` | [ ] |
| C2-03 | `-2+5+cmd\|"/c calc"!A1` | `'-2+5+cmd\|"/c calc"!A1` | [ ] |
| C2-04 | `@SUM(1+1)*cmd\|"/c calc"!A1` | `'@SUM(1+1)*cmd\|"/c calc"!A1` | [ ] |
| C2-05 | `Normal text value` | `Normal text value` (no change) | [ ] |
| C2-06 | `Value with "quotes"` | `Value with ""quotes""` (doubled) | [ ] |
| C2-07 | `Line1\r\nLine2` | `Line1 Line2` (CRLF replaced with space) | [ ] |
| C2-08 | `\tTabbed` | ` Tabbed` (tab replaced with space) | [ ] |
| C2-09 | (null) | Empty quoted string `""` | [ ] |
| C2-10 | `   =SUM(A1)` | `   '=SUM(A1)` (leading spaces preserved, `'` before `=`) | [ ] |

### Integration Verification

| ID | Action | Expected | Pass |
|----|--------|----------|------|
| C2-11 | Export BOM with a component whose DESCP contains `=cmd\|"/c calc"!A1` | CSV file opens in Excel without executing commands | [ ] |
| C2-12 | Export BOM with normal data | CSV content matches grid data exactly | [ ] |

---

## CRITICAL #3 — Export Path Validation

**File:** `MainForm.vb`  
**Controls:** `ValidateExportPath`, `InitializeExportDirectories`, `IsAllowedExportDirectory`

### Test Cases

| ID | Path Tested | Expected Result | Pass |
|----|-------------|-----------------|------|
| C3-01 | `C:\Users\<user>\Documents\test.csv` | Accepted — Documents is allowed | [ ] |
| C3-02 | `C:\Users\<user>\Desktop\export.csv` | Accepted — Desktop is allowed | [ ] |
| C3-03 | `C:\Users\<user>\Downloads\bom.csv` | Accepted — Downloads is allowed | [ ] |
| C3-04 | `%TEMP%\temp_export.csv` | Accepted — Temp is allowed | [ ] |
| C3-05 | `C:\Windows\System32\evil.csv` | Rejected — not in allowed directories | [ ] |
| C3-06 | `\\server\share\export.csv` | Rejected — UNC paths not allowed | [ ] |
| C3-07 | `C:\Users\<user>\Documents\test.txt` | Rejected — must be .csv extension | [ ] |
| C3-08 | `C:\Users\<user>\Documents\test.csv:hidden` | Rejected — alternate data stream detected | [ ] |
| C3-09 | (empty string) | Rejected — empty path | [ ] |
| C3-10 | `C:\Users\<user>\Documents\` & 300 `A` chars & `.csv` | Rejected — filename too long | [ ] |
| C3-11 | `C:\Users\<user>\Documents\file<name>.csv` | Rejected — invalid filename characters | [ ] |
| C3-12 | `C:\Users\<user>\Documents\subfolder\test.csv` | Accepted — subdirectory of allowed path (if folder exists) | [ ] |

---

## CRITICAL #4 — Path Traversal in Model Handling

**File:** `CreoSessionService.vb`  
**Control:** `ValidateModelIdentifier`, `ValidateParsedModelName`

### Test Cases

| ID | Model Input | Expected Result | Pass |
|----|-------------|-----------------|------|
| C4-01 | `my_part.prt` | Accepted | [ ] |
| C4-02 | `../../../windows/system32` | Rejected — contains `..` | [ ] |
| C4-03 | `..\..\secret.prt` | Rejected — contains `..` and `\` | [ ] |
| C4-04 | `C:\full\path\model.prt` | Rejected — rooted path | [ ] |
| C4-05 | `\\server\share\model.prt` | Rejected — UNC path | [ ] |
| C4-06 | `//server/share/model.prt` | Rejected — UNC path | [ ] |
| C4-07 | `model\sub\file.prt` | Rejected — contains `\` | [ ] |
| C4-08 | `model/sub/file.prt` | Rejected — contains `/` | [ ] |
| C4-09 | String of 300 characters | Rejected — exceeds MaxModelIdentifierLength (255) | [ ] |
| C4-10 | `model` & Chr(0) & `.prt` | Rejected — control character | [ ] |
| C4-11 | `model<pipe>.prt` | Rejected — invalid filename character | [ ] |
| C4-12 | (empty string) | Rejected — empty input | [ ] |

---

## CRITICAL #5 — Information Disclosure in Error Messages

**File:** `CreoSessionService.vb`  
**Controls:** `GetUserFacingErrorMessage`, `LogApplicationError`, `BuildDiagnosticMessage`

### Test Cases

| ID | Scenario | Expected UI Message | Expected Diagnostic Log | Pass |
|----|----------|--------------------|-----------------------|------|
| C5-01 | COM object creation fails (COMException) | "The operation failed." | Full exception type, message, stack trace in Trace output | [ ] |
| C5-02 | Model retrieval throws NullReferenceException | "The operation failed." | Full diagnostic in Trace | [ ] |
| C5-03 | ArgumentException("Model name is empty.") | "Model name is empty." (pass-through) | Full diagnostic in Trace | [ ] |
| C5-04 | InvalidOperationException with COM error code | "The operation failed." | COM error code visible in Trace only | [ ] |
| C5-05 | Nested exception (inner + outer) | "The operation failed." | Both exception levels logged with depth indices | [ ] |

### Verification Method

- Inspect UI log output: must never contain file paths, stack traces, COM error codes, or internal class names.
- Inspect `System.Diagnostics.Trace` listener output: must contain full diagnostic details for each error.

---

## HIGH #1 — BOM Recursion Depth & Node Limits

**File:** `CreoSessionService.vb`  
**Controls:** `MaxBomDepth` (32), `MaxBomNodes` (10,000), `visitedModels` HashSet

### Test Cases

| ID | Scenario | Expected Result | Pass |
|----|----------|-----------------|------|
| H1-01 | Assembly with 5 nesting levels | Full BOM returned normally | [ ] |
| H1-02 | Assembly with 33+ nesting levels | BOM truncated at depth 32; no crash or stack overflow | [ ] |
| H1-03 | Assembly with circular reference (A contains B, B contains A) | Circular branch stopped; BOM returned without infinite loop | [ ] |
| H1-04 | Assembly with 10,001+ total components | BOM capped at 10,000 nodes; remaining nodes omitted | [ ] |
| H1-05 | Assembly where same sub-assembly appears as sibling twice | Both instances included (visited tracking is path-based, not global) | [ ] |
| H1-06 | Single part (not assembly) | "BOM hierarchy requires an assembly model." error | [ ] |

### Verification Method

- Count nodes in returned `BomQueryResult.Rows`.
- For circular references: confirm the BOM completes in bounded time and does not hang.

---

## HIGH #2 — UI-Layer Model Input Validation

**File:** `MainForm.vb`  
**Control:** `ValidateModelInput`

### Test Cases

| ID | Input | Expected Result | Pass |
|----|-------|-----------------|------|
| H2-01 | `valid_model` | Accepted (trimmed) | [ ] |
| H2-02 | `  valid_model  ` | Accepted (whitespace trimmed) | [ ] |
| H2-03 | (empty) | Rejected — "Enter a model name." | [ ] |
| H2-04 | `   ` (whitespace only) | Rejected — "Enter a model name." | [ ] |
| H2-05 | 261-character string | Rejected — "Model name is too long." | [ ] |
| H2-06 | `path\to\model` | Rejected — "Enter a Creo model name, not a file path." | [ ] |
| H2-07 | `path/to/model` | Rejected — same as above | [ ] |
| H2-08 | `../traversal` | Rejected — same as above | [ ] |
| H2-09 | `model` & Chr(1) | Rejected — "Model name contains invalid characters." | [ ] |
| H2-10 | `model` & vbTab | Rejected — control character | [ ] |

### Affected Event Handlers

All of the following must call `ValidateModelInput` before proceeding:

- [ ] `RetrieveClickedAsync`
- [ ] `RegenerateClickedAsync`
- [ ] `LoadParametersClickedAsync`
- [ ] `LoadFeaturesClickedAsync`
- [ ] `LoadComponentsClickedAsync`
- [ ] `BuildBomClickedAsync`
- [ ] `SetSelectedParameterClickedAsync`

---

## HIGH #3 — ModelNameParser Extension Whitelist

**File:** `ModelNameParser.vb`  
**Controls:** `AllowedExtensions`, `MaxNameLength`

### Test Cases

| ID | Input | Expected Result | Pass |
|----|-------|-----------------|------|
| H3-01 | `model.prt` | Parsed: Name=model, Ext=PRT, Kind=Part | [ ] |
| H3-02 | `assembly.asm` | Parsed: Name=assembly, Ext=ASM, Kind=Assembly | [ ] |
| H3-03 | `drawing.drw` | Parsed: Name=drawing, Ext=DRW, Kind=Drawing | [ ] |
| H3-04 | `model.mfg` | Parsed: Name=model, Ext=MFG, Kind=Unknown | [ ] |
| H3-05 | `model.frm` | Parsed: Ext=FRM | [ ] |
| H3-06 | `model.lay` | Parsed: Ext=LAY | [ ] |
| H3-07 | `model.dgm` | Parsed: Ext=DGM | [ ] |
| H3-08 | `model.sec` | Parsed: Ext=SEC | [ ] |
| H3-09 | `model.skf` | Parsed: Ext=SKF | [ ] |
| H3-10 | `model.exe` | Rejected — "Unsupported model extension: .exe" | [ ] |
| H3-11 | `model.asp` | Rejected — "Unsupported model extension: .asp" | [ ] |
| H3-12 | `model.dll` | Rejected — "Unsupported model extension: .dll" | [ ] |
| H3-13 | `model.verylongextension` | Rejected — not in whitelist | [ ] |
| H3-14 | `model` (no extension) | Parsed: Kind=Unknown, HasExplicitExtension=False | [ ] |
| H3-15 | `model.prt.3` | Parsed: version stripped → Name=model, Ext=PRT | [ ] |
| H3-16 | 261-character input | Rejected — "Model name exceeds maximum length." | [ ] |
| H3-17 | (empty) | Rejected — "Model name is empty." | [ ] |
| H3-18 | `.prt` (no base name) | Rejected — "Model name is empty." | [ ] |

---

## HIGH #4 — Error Handling in Event Handlers

**File:** `MainForm.vb`  
**Controls:** `HandleUnexpectedError`, `GenericUiErrorMessage`

### Test Cases

| ID | Scenario | Expected UI Log | Pass |
|----|----------|----------------|------|
| H4-01 | `LoadParametersClickedAsync` throws unexpected exception | "ERROR: The operation could not be completed." | [ ] |
| H4-02 | `LoadFeaturesClickedAsync` throws unexpected exception | Same generic message | [ ] |
| H4-03 | `LoadComponentsClickedAsync` throws unexpected exception | Same generic message | [ ] |
| H4-04 | `BuildBomClickedAsync` throws unexpected exception | Same generic message | [ ] |
| H4-05 | `SetSelectedParameterClickedAsync` throws unexpected exception | Same generic message | [ ] |
| H4-06 | `RunOperationAsync` catches exception from service | Same generic message | [ ] |
| H4-07 | Service returns `OperationResult.Fail("...")` | Sanitized service message shown (not raw exception) | [ ] |

### Verification Method

- Confirm no stack traces, file paths, or internal class names appear in the UI log.
- Confirm `LogApplicationError` is called in every Catch block (verify via Trace listener or debugger).

---

## Functional Regression Tests

Verify all existing features still work correctly after security hardening.

### Core Operations

| ID | Test | Expected | Pass |
|----|------|----------|------|
| F-01 | Click Connect with Creo running | "Connected to active Creo session." | [ ] |
| F-02 | Click Connect without Creo running | Session starts or reports connection failure | [ ] |
| F-03 | Enter `model.prt`, click Retrieve | Model retrieved and window activated | [ ] |
| F-04 | Enter `model` (no extension), click Retrieve | Extension auto-resolved, model retrieved | [ ] |
| F-05 | Enter `model.prt.3` (versioned), click Retrieve | Version stripped, model retrieved | [ ] |
| F-06 | Click Regenerate on retrieved model | Model regenerated successfully | [ ] |
| F-07 | Click Load Parameters | Parameters grid populated | [ ] |
| F-08 | Edit parameter, click Set Selected | Parameter updated, grid refreshed | [ ] |
| F-09 | Click Load Features | Features grid populated | [ ] |
| F-10 | Click Load Components on assembly | Components grid populated | [ ] |

### BOM Operations

| ID | Test | Expected | Pass |
|----|------|----------|------|
| F-11 | Build BOM on assembly | BOM grid populated with hierarchy | [ ] |
| F-12 | Expand/collapse BOM nodes | Rows show/hide correctly | [ ] |
| F-13 | Expand All / Collapse All | All nodes toggled | [ ] |
| F-14 | Select BOM row | Item parameters auto-load in side panel | [ ] |
| F-15 | Click Show Params button | Side panel appears with parameters | [ ] |
| F-16 | Close item parameters panel (X) | Panel hidden | [ ] |
| F-17 | Set item parameter from BOM panel | Parameter updated on child component | [ ] |
| F-18 | Export BOM to CSV (Documents folder) | CSV file created with correct content | [ ] |
| F-19 | Export BOM to CSV (Desktop folder) | CSV file created | [ ] |
| F-20 | Open exported CSV in Excel | No formula execution, data displays correctly | [ ] |

### Edge Cases

| ID | Test | Expected | Pass |
|----|------|----------|------|
| F-21 | Rapid-click multiple operations | Busy state prevents concurrent execution | [ ] |
| F-22 | Enter very long valid model name (< 260 chars) | Accepted and processed | [ ] |
| F-23 | Mixed case extension (`Model.PrT`) | Normalized and accepted | [ ] |
| F-24 | Cancel Save dialog during CSV export | No file created, no error | [ ] |

---

## Performance Benchmarks

| ID | Metric | Threshold | Measured | Pass |
|----|--------|-----------|----------|------|
| P-01 | Model load time | < 2 seconds | | [ ] |
| P-02 | Parameter load time | < 3 seconds | | [ ] |
| P-03 | BOM build time (typical assembly) | < 5 seconds | | [ ] |
| P-04 | CSV export time | < 2 seconds | | [ ] |
| P-05 | Memory usage (steady state) | < 500 MB | | [ ] |

---

## Sign-Off

| Role | Name | Date | Signature |
|------|------|------|-----------|
| Developer | | | |
| Reviewer | | | |
| QA | | | |

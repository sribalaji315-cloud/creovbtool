# Dependency & Endpoint-Risk Audit — CreoRegenAppVB

**Date:** 2026-03-20  
**Auditor:** Automated security audit agent  
**Repository:** `sribalaji315-cloud/creovbtool`  
**Scope:** Full repository scan — source, project files, configuration, binaries, documentation

---

## Scope and Methodology

Every file in the repository was reviewed:

| File | Lines | Role |
|---|---|---|
| `CreoRegenAppVB.vbproj` | 54 | MSBuild project — declares references, build configuration |
| `Program.vb` | 13 | Entry point |
| `MainForm.vb` | 997 | WinForms UI layer |
| `CreoSessionService.vb` | 1 511 | COM-interop service, threading, BOM |
| `ModelNameParser.vb` | 103 | Pure-logic model-name parsing |
| `README.md` | 47 | Deployment instructions |
| `dist/CreoRegenAppVB-Update.zip` | — | Pre-built release binary |
| `.gitignore` | — | Repo hygiene |

No other files, folders, or configuration were found.

---

## Part A — Full Dependency Inventory

### A.1 Declared Package Dependencies

| # | Package / Assembly | Version | Source Type | Where Declared | Direct / Transitive |
|---|---|---|---|---|---|
| 1 | `System` | .NET Framework 4.8 (system-provided) | GAC / .NET BCL | `CreoRegenAppVB.vbproj` line 41 | Direct |
| 2 | `System.Core` | .NET Framework 4.8 (system-provided) | GAC / .NET BCL | `CreoRegenAppVB.vbproj` line 42 | Direct |
| 3 | `System.Drawing` | .NET Framework 4.8 (system-provided) | GAC / .NET BCL | `CreoRegenAppVB.vbproj` line 43 | Direct |
| 4 | `System.Windows.Forms` | .NET Framework 4.8 (system-provided) | GAC / .NET BCL | `CreoRegenAppVB.vbproj` line 44 | Direct |
| 5 | `System.Xml` | .NET Framework 4.8 (system-provided) | GAC / .NET BCL | `CreoRegenAppVB.vbproj` line 45 | Direct |
| 6 | `Microsoft.VisualBasic` | .NET Framework 4.8 (system-provided) | GAC / .NET BCL | `CreoRegenAppVB.vbproj` line 40 | Direct |

**No NuGet packages are used.** There is no `packages.config`, `Directory.Packages.props`, or Paket file.  
**No third-party or open-source libraries are present** as static references.

### A.2 Runtime-Provided / External COM Dependencies

The application communicates at runtime with PTC Creo's VB API COM server. These dependencies are **not bundled** — they must be pre-installed on the target machine.

| # | Component | ProgID Pattern | Source | Registration Requirement | Reference |
|---|---|---|---|---|---|
| 7 | Creo VB API — async connection | `pfcls.pfcAsyncConnection` / `pfcls.CpfcAsyncConnection` / `pfc.pfcAsyncConnection` | Runtime COM (PTC Creo install) | `vb_api_register.bat` must be run | `CreoSessionService.vb` line 184 |
| 8 | Creo VB API — model descriptor | `pfcls.pfcModelDescriptor` / `pfcls.CpfcModelDescriptor` | Runtime COM | Same | `CreoSessionService.vb` line 726 |
| 9 | Creo VB API — regen instructions | `pfcls.CpfcRegenInstructions` | Runtime COM | Same | `CreoSessionService.vb` line 262 |
| 10 | Creo VB API — model item factory | `pfcls.CMpfcModelItem` (also tried as `pfcls.MpfcModelItem`, `pfcls.pfcModelItem`) | Runtime COM | Same | `CreoSessionService.vb` lines 1438, 1444, 1447 |
| 11 | `parametric.exe` | — | PTC Creo install | Must be on `PATH` or fully qualified | `README.md` line 34; `CreoSessionService.vb` line 721 |
| 12 | `pro_comm_msg.exe` | — | PTC Creo install | Set via `PRO_COMM_MSG_EXE` environment variable | `README.md` line 20 |

> **Note:** The actual ProgID that resolves at runtime depends on which Creo version and VB API release is installed. The code tries four ProgID variants in order for every object class (`pfcls.<Class>`, `pfcls.C<Class>`, `pfc.<Class>`, `<Class>`) — see `CreoSessionService.vb` lines 881–893.

### A.3 Bundled Binaries

| File | Type | SHA-256 | Origin |
|---|---|---|---|
| `dist/CreoRegenAppVB-Update.zip` → `CreoRegenAppVB.exe` | PE32 .NET assembly (GUI), x64 | `6b93bb6711089f03a7dd60adc284e932ad204f90339df9c8bb879f04dbaad190` | Built from this repository's source code |

The binary is a compiled form of the source in this repository. It does **not** bundle Creo, .NET, or any third-party library. No additional DLLs are packed inside the ZIP.

### A.4 System-Inferred Runtime Dependencies

| Component | Version | How Inferred |
|---|---|---|
| .NET Framework | 4.8 | `<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>` — `CreoRegenAppVB.vbproj` line 11 |
| Windows (x64) | Any Windows supporting .NET 4.8 | `<PlatformTarget>x64</PlatformTarget>` — `CreoRegenAppVB.vbproj` lines 21, 31 |
| PTC Creo | Any version with VB API COM server | `README.md`; COM ProgIDs in `CreoSessionService.vb` |

### A.5 Build & Installer Tooling

| Item | Present | Notes |
|---|---|---|
| Solution file (`.sln`) | No | Single-project repository; opens directly as `.vbproj` |
| `packages.config` | No | No NuGet packages |
| `app.config` / `web.config` | No | No configuration files |
| Installer project (Inno Setup, WiX, NSIS) | No | — |
| Post-build / pre-build scripts | No | — |
| `AssemblyInfo.vb` | No | Assembly metadata is embedded in the `.vbproj` |
| NuGet package source | No | — |
| Paket | No | — |

---

## Part B — EDR / Endpoint Security Risk Analysis

This section identifies code patterns that **may attract scrutiny** from endpoint detection and response (EDR) tools such as SentinelOne, Microsoft Defender for Endpoint (MDE), or CrowdStrike Falcon. Risk levels reflect pattern-based heuristic likelihood of detection, **not** confirmed malicious intent. All findings are evidence-based.

---

### B.1 Late-Bound COM Object Creation via ProgID

**Risk level:** Medium (likely benign in context; may attract EDR scrutiny)

**Pattern:**
```vb
Return Activator.CreateInstance(Type.GetTypeFromProgID(progId, throwOnError:=True))
```

**Evidence:**

| Location | Line(s) | ProgIDs attempted |
|---|---|---|
| `CreoSessionService.vb` — `CreatePfcObject()` | 880–893 | `pfcls.<Class>`, `pfcls.C<Class>`, `pfc.<Class>`, `<Class>` |
| `CreoSessionService.vb` — `CreateParamValueFromInput()` | 1438 | `pfcls.CMpfcModelItem` (explicit) |

**Why it may attract EDR scrutiny:**  
Late-bound COM instantiation via `Type.GetTypeFromProgID` / `Activator.CreateInstance` is a technique used by malware (e.g., script-based loaders, LOLBins) to dynamically create COM objects — including `WScript.Shell`, `Shell.Application`, `MSXML2.XMLHTTP`, and others — without declaring them statically. EDR tools routinely monitor `RoGetActivationFactory`, `CoCreateInstance`, and equivalent .NET patterns.

**Mitigating factors:**  
All ProgID strings are compile-time constants referencing the PTC Creo namespace (`pfcls.*`, `pfc.*`). No ProgID is constructed from user input or environment variables. There is no use of `WScript.Shell`, `Shell.Application`, `MSXML2.*`, `Scripting.FileSystemObject`, or any general-purpose Windows COM automation classes.

---

### B.2 Dynamic Late-Bound Method and Property Invocation (`CallByName`)

**Risk level:** Low–Medium (likely benign; may resemble suspicious behavior under behavioral analysis)

**Pattern:**
```vb
CallByName(instance, methodName, CallType.Method, args...)
CallByName(instance, propertyName, CallType.Get)
CallByName(instance, propertyName, CallType.Set, value)
```

**Evidence:**

| Location | Lines | Usage |
|---|---|---|
| `CreoSessionService.vb` — `InvokeMethod()` | 918–932 | Dispatches Creo API method calls with 0–4 arguments |
| `CreoSessionService.vb` — `GetProperty()` | 898–903 | Reads COM object properties |
| `CreoSessionService.vb` — `SetProperty()` | 906–916 | Writes COM object properties (Set then Let fallback) |
| `CreoSessionService.vb` — `TryGetString()` | 943–963 | Safe property-then-method accessor |

All method and property names are string literals inside the source file; none are built from user input.

**Why it may attract EDR scrutiny:**  
Behavioral engines can flag applications that use reflection or late-bound dispatch to invoke methods whose names are not visible in static imports. `Microsoft.VisualBasic.CallByName` is commonly used in VBScript and VBA macro payloads.

**Mitigating factors:**  
The dispatch targets are all Creo VB API objects. No Windows system object, shell, or automation COM class is ever passed to `CallByName`. Method names are hardcoded (`"Connect"`, `"Start"`, `"Regenerate"`, `"ListParams"`, etc.).

---

### B.3 Macro / UI Automation Execution (`RunMacro` with Creo Mapkeys)

**Risk level:** Low–Medium (context-specific; may attract EDR scrutiny as scripted automation)

**Pattern:**
```vb
InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdRegenAuto`; ...")
```

**Evidence:**

| Location | Lines | Mapkey commands used |
|---|---|---|
| `CreoSessionService.vb` — `RunRegenerationMacros()` | 769–783 | `ProCmdNamedViewsGalSelect`, `ProCmdRegenAuto`, `ProCmdDwgRegenModel`, `ProCmdDwgRegenDraft`, `ProCmdRegenPart`, `ProCmdViewRefit`, `ProCmdViewRepaint`, `ProCmdViewRegenCur` |

This code is only invoked as a **fallback** path (lines 266–272) when the primary Creo API `Regenerate()` call fails.

**Why it may attract EDR scrutiny:**  
Macro/script execution injected into a running application process may match heuristics for process-injection or automated UI scripting patterns, even when the target is a CAD application.

**Mitigating factors:**  
All mapkey strings are compile-time constants. No user input is concatenated into the macro strings. The strings contain only Creo-specific commands with no shell, file system, or network semantics. The macro execution path targets the already-running Creo session that the application itself started or connected to.

---

### B.4 Private Reflection (`BindingFlags.NonPublic`)

**Risk level:** Low (likely benign; standard WinForms pattern)

**Pattern:**
```vb
Dim propertyInfo = control.GetType().GetProperty(
    "DoubleBuffered",
    BindingFlags.NonPublic Or BindingFlags.Instance)
If propertyInfo IsNot Nothing Then
    propertyInfo.SetValue(control, True, Nothing)
End If
```

**Evidence:** `MainForm.vb` — `EnableDoubleBuffering()`, lines 889–894.  
Called at `MainForm.vb` lines 411–415 for all five `DataGridView` controls.

**Why it may attract EDR scrutiny:**  
Use of `BindingFlags.NonPublic` to access private class members is a technique used in reflective injection and assembly loading exploits. Behavioral engines with .NET inspection may flag this call.

**Mitigating factors:**  
The reflection targets `System.Windows.Forms.Control.DoubleBuffered`, a well-known internal WinForms property used by countless enterprise applications to eliminate grid flicker. The target control is `this` (the form's own controls). No assembly loading, code generation, or external type resolution is performed.

---

### B.5 STA Background Thread Spawning

**Risk level:** Low (normal COM threading pattern; unlikely to trigger standalone)

**Pattern:**
```vb
_worker = New Thread(AddressOf ProcessQueue)
_worker.Name = "CreoApiWorker"
_worker.IsBackground = True
_worker.SetApartmentState(ApartmentState.STA)
_worker.Start()
```

**Evidence:** `CreoSessionService.vb` lines 166–171.

**Why it may attract EDR scrutiny:**  
Background thread creation with STA apartment state can appear in COM in-process server exploits and certain shellcode runners. When combined with COM object creation on the same thread, some heuristics may elevate the signal.

**Mitigating factors:**  
STA apartment state is a **requirement** for classic Win32 COM usage (the Creo VB API is a classic COM server). The thread's sole purpose (lines 874–877) is to drain a `BlockingCollection<Action>` of Creo API calls. This is a canonical pattern for any .NET app that wraps a COM API.

---

### B.6 File System Write (User-Controlled Path via `SaveFileDialog`)

**Risk level:** Low (standard WinForms file-save pattern)

**Pattern:**
```vb
Using dialog As New SaveFileDialog()
    ...
    Using writer As New StreamWriter(dialog.FileName, False, Encoding.UTF8)
        ...
    End Using
End Using
```

**Evidence:** `MainForm.vb` — `ExportBomClicked()`, lines 813–847.

**Context:** The user selects a path via the standard `SaveFileDialog` to export a BOM CSV. The file content is derived entirely from UI data already displayed in the grid; no external data is written.

**Why it may attract EDR scrutiny:**  
Any file write from a running .NET process is potentially monitored. Write operations outside expected application directories may trigger behavioral alerts.

**Mitigating factors:**  
Path comes from a system dialog (`SaveFileDialog`), which inherits the OS shell's path validation. The content written is plain CSV text derived from read-only CAD metadata. No executable, script, or binary content is written.

---

### B.7 External Process Implied — Creo (`parametric.exe`)

**Risk level:** Medium (context-dependent; may attract EDR scrutiny if Creo is not whitelisted)

**Pattern:**  
The application's `Connect` flow calls the Creo COM object's `Start` method when no active Creo session is found. Under the hood, the Creo VB API launches `parametric.exe` as a child process of the COM server.

```vb
_asyncConnection = InvokeMethod(creator, "Start", command, settings.TextDirectory)
```

**Evidence:** `CreoSessionService.vb` line 192; `README.md` lines 34–35.

The `BuildStartCommand()` function (`CreoSessionService.vb` lines 720–723) falls back to the string `"parametric"` if no override is provided.  
In `MainForm.vb` line 435, `ConnectClickedAsync` always passes `CreoStartCommand = "parametric"`.

**Why it may attract EDR scrutiny:**  
EDR products monitor process trees closely. A .NET GUI application spawning a CAD process (`parametric.exe`) through a COM intermediary is an unusual call chain. SentinelOne, CrowdStrike, and MDE all inspect parent–child process relationships. An unknown GUI EXE spawning `parametric.exe` via COM may be flagged if Creo is not in the sensor's whitelist.

**Mitigating factors:**  
The command string is a compile-time constant. No shell metacharacters, command injection, or dynamic path construction is present. The `TextDirectory` is always `"."` (current working directory), set at `MainForm.vb` line 436.

---

### B.8 Environment Variable Dependency (`PRO_COMM_MSG_EXE`, `PFCLS_START_DIR`)

**Risk level:** Low (configuration, not a code risk)

**Evidence:** `README.md` lines 19–22.

These environment variables are consumed by the Creo VB API COM server (not by this application's code directly). If `PRO_COMM_MSG_EXE` points to a non-PTC binary, the COM server would launch an untrusted executable. This is a **deployment risk**, not a code risk.

**Recommendation:** Validate that deployment scripts or IT provisioning processes set `PRO_COMM_MSG_EXE` to the PTC-signed binary at `<creo_loadpoint>\<datecode>\Common Files\x86e_win64\obj\pro_comm_msg.exe`.

---

## Part C — Normal Enterprise Desktop-App Behavior vs. Higher-Risk Patterns

### C.1 Normal / Expected Enterprise Desktop-App Behavior

| Behavior | Evidence |
|---|---|
| WinForms GUI with layout panels, tabs, data grids | `MainForm.vb` — entire file |
| Async/await on UI thread for non-blocking operations | `MainForm.vb` lines 431–451, 479–526 |
| COM interop to a vendor-supplied API | `CreoSessionService.vb` — all COM calls |
| CSV file export via standard Save dialog | `MainForm.vb` lines 807–848 |
| `IDisposable` pattern with resource cleanup | `CreoSessionService.vb` lines 1485–1509 |
| Structured error handling; no crash-on-failure | Throughout `CreoSessionService.vb` |
| No embedded credentials, secrets, or tokens | Confirmed — no such strings in source or binary |
| No network calls | Confirmed — no `WebClient`, `HttpClient`, `Socket`, DNS, or URL references |
| No registry access | Confirmed — no `RegistryKey`, `Registry`, or `HKEY_*` references |
| No P/Invoke or `DllImport` | Confirmed — none in any source file |
| No WMI usage | Confirmed — no `ManagementObject`, `ManagementScope`, WMI queries |

### C.2 Higher-Risk Patterns — Review Warranted

| # | Pattern | Location | Risk | Recommended Action |
|---|---|---|---|---|
| 1 | Late-bound COM via `Type.GetTypeFromProgID` + `Activator.CreateInstance` | `CreoSessionService.vb` 880–893, 1438 | Medium | Add to EDR application allowlist; document ProgID inventory to ops team |
| 2 | `CallByName` dynamic dispatch | `CreoSessionService.vb` 900–953 | Low–Medium | Review at EDR rule tuning time; acceptable given compile-time method names |
| 3 | `RunMacro` with Creo mapkey strings | `CreoSessionService.vb` 769–783 | Low–Medium | Acceptable as fallback path; monitor for unexpected macro string changes |
| 4 | COM-launched external process (`parametric.exe`) | `CreoSessionService.vb` 192; `README.md` 34 | Medium | Whitelist the Creo process tree in EDR policy; verify binary signing |
| 5 | Private reflection on WinForms controls | `MainForm.vb` 889–894 | Low | No action required; well-understood pattern |
| 6 | STA background thread + COM calls | `CreoSessionService.vb` 166–171 | Low | No action required; mandatory for COM threading model |
| 7 | Pre-built binary in `dist/` | `dist/CreoRegenAppVB-Update.zip` | Low–Medium | Ensure the binary is signed before distribution; hash documented above |

---

## Part D — Findings Not Present (Confirmed Absences)

The following were explicitly checked and are **not** present in this codebase:

- No `Shell()`, `Process.Start()`, `CreateProcessW`, or direct executable launch from VB/C# code.
- No `Registry`, `RegistryKey`, `HKEY_*` access.
- No `WebClient`, `HttpClient`, `WebRequest`, `Socket`, DNS resolution, or any network I/O.
- No `DllImport` or `extern` P/Invoke declarations.
- No `WScript.Shell`, `Shell.Application`, `Scripting.FileSystemObject` COM objects.
- No WMI (`System.Management`, `ManagementObject`).
- No cryptography (`System.Security.Cryptography`, `RijndaelManaged`, `AesCryptoServiceProvider`).
- No dynamic code compilation (`CSharpCodeProvider`, `VBCodeProvider`, `Roslyn`).
- No `Assembly.Load`, `Assembly.LoadFrom`, `Assembly.LoadFile` (no reflective assembly loading).
- No browser/`WebBrowser` control.
- No external URL references, API endpoints, update servers, or telemetry.
- No embedded resources (`*.resx` not present).
- No `app.config` or `web.config`.
- No `AssemblyInfo.vb` (metadata embedded in project file).
- No installer or bootstrapper scripts.
- No scheduled task, service installation, or autostart mechanism.
- No hardcoded credentials, API keys, passwords, tokens, or certificates.

---

## Part E — Recommendations Summary

| Priority | Recommendation |
|---|---|
| High | Sign the release binary (`CreoRegenAppVB.exe`) with an authenticode certificate before distribution. An unsigned PE32 .NET GUI binary may trigger MDE Smart App Control or SentinelOne reputation-based blocks. |
| High | Add `CreoRegenAppVB.exe` and the associated Creo process tree to the EDR application allowlist on all target machines. |
| Medium | Document and pin the Creo ProgID set (`pfcls.*`, `pfc.*`) in your IT knowledge base. Share with the EDR administrator so COM instantiation by this application can be distinguished from generic COM-based attack patterns. |
| Medium | Validate the `PRO_COMM_MSG_EXE` environment variable at application startup. Alert or log if it does not point to a PTC-signed binary. |
| Low | Consider adding an `app.config` with an explicit `<supportedRuntime>` element to ensure the correct .NET 4.8 runtime is selected without relying on implicit binding. |
| Low | Remove or `.gitignore` the `dist/` folder if binaries are managed through a CI/CD pipeline instead. Committing prebuilt binaries alongside source creates audit complexity. |
| Informational | The `OptionStrict Off` compiler setting (`CreoRegenAppVB.vbproj` line 17) permits implicit late-binding. This is intentional (see `README.md` line 45) but should be documented. Consider enabling `OptionStrict On` in the long term and using explicit `CObj()` casts at COM call sites to make the late-binding intent explicit without relying on the compiler option. |

---

## Appendix — Binary Artifact Details

**File:** `dist/CreoRegenAppVB-Update.zip` → `CreoRegenAppVB.exe`

| Attribute | Value |
|---|---|
| Compressed size | 31 KB |
| Uncompressed size | ~79 KB |
| File type | PE32 .NET assembly (GUI), x86/x64-neutral CLR header |
| Architecture declared in project | x64 (`<PlatformTarget>x64</PlatformTarget>`) |
| SHA-256 (unzipped EXE) | `6b93bb6711089f03a7dd60adc284e932ad204f90339df9c8bb879f04dbaad190` |
| Authenticode signature | Not verified (Linux analysis environment; recommend `sigcheck` on Windows) |
| Embedded ProgID strings confirmed | `Activator`, `CallByName`, `GetTypeFromProgID`, `BindingFlags`, `System.Reflection` |

---

*End of audit report.*

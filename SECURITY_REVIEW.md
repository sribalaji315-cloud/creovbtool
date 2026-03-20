# Security Review Package — CreoRegenAppVB

**Application:** `CreoRegenAppVB.exe`
**Technology:** VB.NET, .NET Framework 4.8, Windows Forms
**Review type:** Static code analysis (source review)
**Review date:** 2026-03-20
**Reviewer:** Automated static analysis, Copilot Coding Agent
**Repository:** `sribalaji315-cloud/creovbtool`
**Branch reviewed:** `copilot/build-dependency-inventory`

> This document was produced for internal IT/security review.  It is based entirely on static
> analysis of the source code in this repository.  No binary execution or dynamic analysis was
> performed.  All findings are classified conservatively.

---

## Section A: Executive Summary

### What is this application?

`CreoRegenAppVB.exe` is a **Windows Forms desktop utility** that connects to PTC Creo Parametric
(a professional 3-D CAD application) through Creo's published **VB API COM interface**.  The
tool allows an engineer to retrieve a Creo model, inspect and edit its parameters, regenerate
the model geometry, browse the feature tree, list assembly components, and export a Bill-of-
Materials (BOM) CSV file.  All interaction with Creo takes place through documented COM objects
that Creo itself registers during installation.

### Likely reasons the EXE was flagged

| # | Reason |
|---|--------|
| 1 | **Unsigned executable** — the binary has no Authenticode signature. Endpoint security products assign low reputation to new unsigned EXEs. |
| 2 | **Unknown/low-reputation binary** — internally built and privately distributed; no commercial signature or catalog entry. |
| 3 | **Late-bound COM instantiation** — the code calls `Activator.CreateInstance(Type.GetTypeFromProgID(…))` and `Microsoft.VisualBasic.Interaction.CallByName(…)` to drive Creo's COM server. Dynamic COM object creation is a pattern shared by many legitimate automation scripts and by some malware families, so it may attract heuristic scrutiny. |
| 4 | **Reflection use** — `System.Reflection` is imported and used to access a non-public WinForms property for UI double-buffering. Reflection-based property access is a known EDR trigger. |
| 5 | **Creo COM server may launch a subprocess** — when no active Creo session is found the app calls the Creo VB API `Start` method, which internally starts `parametric.exe` (the Creo process). Although `Process.Start` is never called directly by this code, the COM server performs the launch on the app's behalf. EDR may trace the parent–child relationship back to this EXE. |
| 6 | **`OptionStrict Off`** — disables strict type checking to allow late-bound COM calls. Obfuscated malware frequently turns off strict typing to hide runtime type coercions. |
| 7 | **CSV file write** — the BOM export writes a user-chosen CSV file via `StreamWriter`. File-write activity from a new/unsigned binary is sometimes flagged. |

### Overall classification

| Question | Finding |
|----------|---------|
| Does the code appear benign, suspicious, risky, or unknown? | **Benign** — no clear evidence of malicious intent was found in static review. |
| Are third-party dependencies present? | **No** — zero third-party or NuGet packages; only .NET Framework system assemblies. |
| Was any clearly malicious or system-harming code found? | **No** — none found. |
| Should IT review this before approving? | **Yes** — standard process for any new internal EXE, especially one that interacts with registered COM servers. |

### Top review points for IT

1. Confirm the COM ProgIDs used (`pfcls.pfcAsyncConnection`, `pfcls.CpfcRegenInstructions`,
   `pfcls.pfcModelDescriptor`, `pfcls.CMpfcModelItem`, `pfc.*` variants) are registered by a
   legitimate PTC Creo installation — not by any unknown software.
2. Verify the `PRO_COMM_MSG_EXE` environment variable points to the real Creo install path, not
   a hijacked binary.
3. The only file-write path is the user-initiated BOM CSV export (`SaveFileDialog` → `StreamWriter`).
   There is no write to `%TEMP%`, `AppData`, `ProgramData`, the Startup folder, or any system
   path.
4. There is no network access of any kind.
5. Code-signing the EXE with an internal certificate would eliminate most AV false positives
   caused by reasons 1 and 2 above.

---

## Section B: Dependency and Library List

### B.1 Assembly references declared in `CreoRegenAppVB.vbproj`

| Name | Version | Type | Classification | Where found | Direct / Transitive | Build / Install / Runtime | Purpose | Evidence |
|------|---------|------|----------------|-------------|---------------------|---------------------------|---------|----------|
| `Microsoft.VisualBasic` | .NET Framework 4.8 (GAC) | Assembly | **Microsoft / system** | `CreoRegenAppVB.vbproj` line 40 | Direct | Runtime | VB.NET runtime helpers, `CallByName`, `Interaction` | `<Reference Include="Microsoft.VisualBasic" />` |
| `System` | .NET Framework 4.8 (GAC) | Assembly | **Microsoft / system** | `CreoRegenAppVB.vbproj` line 41 | Direct | Runtime | Core .NET types: `String`, `Convert`, `Exception`, `Activator`, `Type`, `Environment`, etc. | `<Reference Include="System" />` |
| `System.Core` | .NET Framework 4.8 (GAC) | Assembly | **Microsoft / system** | `CreoRegenAppVB.vbproj` line 42 | Direct | Runtime | LINQ, `HashSet(Of T)`, lambda support | `<Reference Include="System.Core" />` |
| `System.Drawing` | .NET Framework 4.8 (GAC) | Assembly | **Microsoft / system** | `CreoRegenAppVB.vbproj` line 43 | Direct | Runtime | WinForms drawing: `Color`, `Font`, `Size` | `<Reference Include="System.Drawing" />` |
| `System.Windows.Forms` | .NET Framework 4.8 (GAC) | Assembly | **Microsoft / system** | `CreoRegenAppVB.vbproj` line 44 | Direct | Runtime | WinForms UI: `Form`, `Button`, `DataGridView`, `SaveFileDialog`, etc. | `<Reference Include="System.Windows.Forms" />` |
| `System.Xml` | .NET Framework 4.8 (GAC) | Assembly | **Microsoft / system** | `CreoRegenAppVB.vbproj` line 45 | Direct | Runtime | XML support (referenced but not directly called in current source) | `<Reference Include="System.Xml" />` |

> **No NuGet packages.** The repository contains no `packages.config`, no `<PackageReference>` elements, and no `nuget.config` file.

### B.2 COM objects used at runtime (late-bound, no TLB reference in project)

These are **not** referenced in the project file.  They are resolved at runtime by ProgID string
through the Windows COM registry.  They are provided by PTC Creo and must be pre-registered
before the app can connect to Creo.

| ProgID tried | COM class | Classification | Where used | Purpose | Evidence |
|-------------|-----------|----------------|-----------|---------|----------|
| `pfcls.pfcAsyncConnection` / `pfc.pfcAsyncConnection` / `pfcls.CpfcAsyncConnection` / `pfcAsyncConnection` | Creo VB API — async connection factory | **Third-party / PTC vendor** | `CreoSessionService.vb` line 184 | Create connection to Creo session | `CreatePfcObject("pfcAsyncConnection")` |
| `pfcls.pfcModelDescriptor` / variants | Creo VB API — model descriptor factory | **Third-party / PTC vendor** | `CreoSessionService.vb` line 726 | Create model descriptor for retrieval | `CreatePfcObject("pfcModelDescriptor")` |
| `pfcls.CpfcRegenInstructions` / variants | Creo VB API — regeneration instruction factory | **Third-party / PTC vendor** | `CreoSessionService.vb` lines 262, 382, 465 | Create `IpfcRegenInstructions` object passed to `model.Regenerate()` | `CreatePfcObject("CpfcRegenInstructions")` |
| `pfcls.CMpfcModelItem` | Creo VB API — parameter value factory | **Third-party / PTC vendor** | `CreoSessionService.vb` line 1438 | Create typed parameter value objects (`CreateStringParamValue`, etc.) | `Activator.CreateInstance(Type.GetTypeFromProgID("pfcls.CMpfcModelItem", True))` |
| `pfcls.MpfcModelItem` / `pfcls.pfcModelItem` (fallback) | Creo VB API — parameter value factory (alternate ProgIDs) | **Third-party / PTC vendor** | `CreoSessionService.vb` lines 1444–1448 | Fallback ProgIDs for the same factory | `CreatePfcObject("MpfcModelItem")` / `CreatePfcObject("pfcModelItem")` |

### B.3 Required runtimes

| Runtime | Version | Classification | Notes |
|---------|---------|----------------|-------|
| .NET Framework | 4.8 | **Microsoft / system** | Must be installed on target Windows machine. Included with Windows 10/11 via Windows Update. |
| PTC Creo Parametric with VB API support | Any version registered with the above ProgIDs | **Third-party / PTC vendor** | Must be installed and its COM server must be registered via `vb_api_register.bat` before this app can connect. |

### B.4 External executable referenced at runtime

| Item | Classification | How referenced | Purpose |
|------|----------------|---------------|---------|
| `parametric.exe` (Creo Parametric) | **Third-party / PTC vendor** | Passed as `command` argument to the Creo VB API `Connect`/`Start` COM method (`BuildStartCommand("parametric")`, `CreoSessionService.vb` line 721) | Launches Creo if no active session exists |
| `pro_comm_msg.exe` (Creo communications bridge) | **Third-party / PTC vendor** | Required by the Creo VB API COM server; set via environment variable `PRO_COMM_MSG_EXE` (documented in `README.md` line 20) | IPC bridge used internally by the Creo COM server; not called by this code |

> No `Process.Start`, `Shell()`, `cmd.exe`, `powershell.exe`, `wscript.exe`, `cscript.exe`,
> `mshta.exe`, `rundll32.exe`, `regsvr32.exe`, `schtasks.exe`, `wmic.exe`, `certutil.exe`,
> `bitsadmin.exe`, `taskkill`, `netsh`, `sc.exe`, `bcdedit`, or `vssadmin` references appear
> anywhere in the source code.

### B.5 Local DLL files

None.  No physical `.dll` files are present in the repository, the `dist/` folder, or any
subdirectory.

### B.6 Embedded binary resources

None.  The project contains no embedded resources (`EmbeddedResource` elements) and no
`resx`/`resources` files.

### B.7 Installer / bootstrapper

None.  No installer project (`.vdproj`, `.wixproj`, `.iss`, `.nsi`), MSI, or setup EXE is
present.  Distribution is via `dist/CreoRegenAppVB-Update.zip` — a plain zip file containing
the built output.

---

## Section C: Third-Party Dependency Summary

| Vendor / Package | Version | Purpose | Actively used in code? | Bundled with app? | Notes |
|-----------------|---------|---------|------------------------|-------------------|-------|
| **PTC Creo VB API COM server** (`pfcls.*` / `pfc.*` ProgIDs) | Unknown (varies with Creo installation) | Provides the COM objects through which this app reads/writes Creo model data, parameters, features, and BOM | Yes — all Creo interaction goes through these COM objects | No — must be pre-installed on the host machine via Creo's own installer | Standard, documented CAD automation API from PTC (Parametric Technology Corporation). It is a well-known B2B engineering software vendor. These ProgIDs are specific to Creo and are not used by malware. |

> There are **no NuGet packages, no npm/pip/Maven dependencies, no third-party DLLs bundled with
> the app, and no other vendor code** included in the build output.  The only third-party
> component is the PTC Creo VB API, which must exist on the machine independently.

---

## Section D: External Components and Binaries

| Item | Type | Origin | Bundled? | Notes |
|------|------|--------|---------|-------|
| `pfcls.dll` / Creo VB API COM server | DLL (registered COM server) | PTC Creo installation | No | Loaded by Windows COM infrastructure when `Activator.CreateInstance(Type.GetTypeFromProgID(…))` is called. Located in the Creo install tree. |
| `pro_comm_msg.exe` | EXE | PTC Creo installation | No | IPC bridge used internally by the Creo COM server. Path set in `PRO_COMM_MSG_EXE` environment variable. This app does not execute it directly. |
| `parametric.exe` | EXE | PTC Creo installation | No | Main Creo application. Started indirectly through the Creo VB API `Start` COM method when no active session exists. |
| `CreoRegenAppVB-Update.zip` | ZIP archive | Custom / local build output | Yes (in `dist/`) | Distribution package. Contains the built application files. This zip was not decompressed or analyzed; its contents are assumed to match the build output of this source. |

> **No WebView2, ActiveX/AxHost, browser controls, WebBrowser control, or embedded browser
> dependencies** are present in the project.

---

## Section E: Potential EDR-Triggering Behaviors

### E.1 Late-bound COM instantiation via `Activator.CreateInstance` + `Type.GetTypeFromProgID`

| Field | Detail |
|-------|--------|
| **Behavior** | The `CreatePfcObject()` method (line 880) and the `CreateParamValueFromInput()` method (line 1438) instantiate COM objects using `Activator.CreateInstance(Type.GetTypeFromProgID(progId, throwOnError:=True))`. |
| **Why it may attract EDR scrutiny** | Dynamic COM activation by ProgID string (rather than a statically typed TLB reference) is a technique used by script-based malware and living-off-the-land attacks to invoke COM objects without import-table evidence. EDR behavioural engines may flag `GetTypeFromProgID` calls for objects not commonly seen in the process context. |
| **Severity** | Low |
| **Confidence** | Medium — the ProgIDs (`pfcls.*`, `pfc.*`) are Creo-specific and not associated with any known malicious activity. |
| **Evidence** | `CreoSessionService.vb` lines 880–895 (`CreatePfcObject`), line 1438 (`Activator.CreateInstance(Type.GetTypeFromProgID("pfcls.CMpfcModelItem", True))`). |
| **Necessary for business function?** | **Yes** — this is the only way to drive Creo's VB API without hard-coding type library paths at build time, as documented in `README.md` line 45. |
| **Safer alternative** | Add a strong-named COM TLB reference (`<COMReference>`) to the `.vbproj` so the linker records the COM dependency statically. This reduces heuristic suspicion at the cost of requiring a specific Creo version at build time. |

### E.2 `CallByName` for dynamic property/method dispatch on COM objects

| Field | Detail |
|-------|--------|
| **Behavior** | `Microsoft.VisualBasic.Interaction.CallByName` is used 12 times in `CreoSessionService.vb` (lines 900, 908, 911, 921, 923, 925, 927, 929, 947, 953, 1382, 1387) to get/set COM properties and invoke COM methods by string name at runtime. |
| **Why it may attract EDR scrutiny** | Invoking arbitrary methods by runtime string is semantically similar to reflection-based code injection. AV heuristics may flag `CallByName` calls on COM objects as suspicious dynamic dispatch. |
| **Severity** | Low |
| **Confidence** | Low — `CallByName` on COM objects is the standard pattern recommended by Microsoft for VB.NET late-bound COM automation. |
| **Evidence** | `CreoSessionService.vb` — `GetProperty()` (line 900), `SetProperty()` (lines 908, 911), `InvokeMethod()` (lines 921, 923, 925, 927, 929), `TryGetString()` (lines 947, 953), `GetCollectionItem()` (lines 1382, 1387). |
| **Necessary for business function?** | **Yes** — required by the late-bound COM design. |
| **Safer alternative** | Same as E.1: add a TLB COM reference to the project to switch to early-bound COM calls, eliminating the need for `CallByName`. |

### E.3 Indirect process launch via Creo COM API

| Field | Detail |
|-------|--------|
| **Behavior** | When `EnsureConnectedAsync` cannot find an active Creo session, it calls `InvokeMethod(creator, "Start", command, settings.TextDirectory)` where `command = "parametric"`. The Creo COM server internally starts `parametric.exe`. This is an **indirect** process launch; `Process.Start` is never called by this code. |
| **Why it may attract EDR scrutiny** | EDR process tree tracking may identify this EXE as the ancestor of a Creo process. Some EDR rules flag parent–child relationships where an unusual parent spawns a CAD or administrative binary. |
| **Severity** | Low |
| **Confidence** | Low — the spawned process is the legitimate Creo application that the user intends to open. |
| **Evidence** | `CreoSessionService.vb` line 192: `_asyncConnection = InvokeMethod(creator, "Start", command, settings.TextDirectory)`. `BuildStartCommand` (line 721): returns `"parametric"` by default. |
| **Necessary for business function?** | **Yes** — opening Creo when not already running is a core feature. |
| **Safer alternative** | Require users to start Creo manually before running this tool (connect-only mode). This would eliminate the process-launch behaviour. |

### E.4 Reflection access to a non-public WinForms property

| Field | Detail |
|-------|--------|
| **Behavior** | `EnableDoubleBuffering()` (`MainForm.vb` line 889) uses `System.Reflection` to read and set the non-public `DoubleBuffered` property on `DataGridView` instances: `control.GetType().GetProperty("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance)`. |
| **Why it may attract EDR scrutiny** | Reflection-based access to non-public members is a technique used to bypass visibility rules. EDR may flag it alongside other dynamic-code patterns in the same binary. |
| **Severity** | Low |
| **Confidence** | Low — this is a well-documented, benign WinForms performance pattern that has been in widespread use since .NET 2.0. |
| **Evidence** | `MainForm.vb` lines 889–893. `Imports System.Reflection` at line 5. |
| **Necessary for business function?** | No — it is a UI rendering optimisation. It can be removed without affecting functionality. |
| **Safer alternative** | Sub-class `DataGridView` and override `DoubleBuffered` as `Public`, removing the reflection call entirely. |

### E.5 CSV file write via `StreamWriter`

| Field | Detail |
|-------|--------|
| **Behavior** | `ExportBomClicked` (`MainForm.vb` lines 807–848) writes a CSV file using `StreamWriter` to a path chosen by the user via `SaveFileDialog`. |
| **Why it may attract EDR scrutiny** | Any file-write activity from a new/unsigned binary may be logged. The write is to a user-selected location, not to any protected or system path. |
| **Severity** | Low |
| **Confidence** | Low — the write is explicitly user-initiated and user-scoped. |
| **Evidence** | `MainForm.vb` lines 813–847: `SaveFileDialog` → `StreamWriter(dialog.FileName, False, Encoding.UTF8)`. |
| **Necessary for business function?** | **Yes** — BOM CSV export is an advertised feature. |
| **Safer alternative** | No meaningful alternative; this is standard desktop-app behaviour. |

### E.6 Creo macro execution via `RunMacro`

| Field | Detail |
|-------|--------|
| **Behavior** | `RunRegenerationMacros()` (`CreoSessionService.vb` lines 769–782) calls `InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command \`…\`;")` multiple times as a fallback when direct API regeneration fails. These are Creo-internal mapkey macro strings that drive Creo's own UI commands. |
| **Why it may attract EDR scrutiny** | Passing shell-like command strings to an API method named `RunMacro` superficially resembles command injection or macro-based malware. An analyst unfamiliar with Creo's API may flag these strings. |
| **Severity** | Low |
| **Confidence** | Low — `RunMacro` is a published Creo VB API method; the macro strings are Creo-internal UI commands (e.g., `ProCmdRegenAuto`, `ProCmdViewRefit`) with no OS-level effect. |
| **Evidence** | `CreoSessionService.vb` lines 770–782. Example: `InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command \`ProCmdRegenAuto\`; …")`. |
| **Necessary for business function?** | **Yes** — this is the fallback regeneration path when the direct COM API call fails (e.g., due to version differences in the Creo VB API). |
| **Safer alternative** | None while retaining compatibility with multiple Creo versions. Document clearly in code comments that these are Creo-internal UI commands and not OS commands. |

---

## Section F: Malicious / System-Harming Code Review

A systematic check was performed against every pattern listed in the problem statement.  Results
are shown below.

### F.1 Pattern coverage table

| Pattern checked | Found? | Finding |
|----------------|-------|---------|
| `Process.Start` | **No** | Not present anywhere in source. |
| `Shell()` | **No** | Not present anywhere in source. |
| `CreateObject()` | **No** | Not used. Late-bound COM uses `Activator.CreateInstance(Type.GetTypeFromProgID(…))` instead, which is the idiomatic .NET equivalent but does not use the VBA `CreateObject` function directly. |
| `GetObject()` | **No** | Not present. |
| `My.Computer.Registry` / `Microsoft.Win32.Registry` | **No** | No registry reads or writes anywhere. |
| `DllImport` | **No** | Not present. |
| `Declare Function` / `Declare Sub` (PInvoke) | **No** | Not present. |
| `Assembly.Load` / `Assembly.LoadFrom` | **No** | Not present. |
| `AppDomain.AssemblyResolve` | **No** | Not present. |
| `WebBrowser` control | **No** | Not present. |
| `WebView2` | **No** | Not present. |
| `AxHost` / ActiveX embedding | **No** | Not present. |
| `HttpClient` | **No** | Not present. |
| `WebClient` | **No** | Not present. |
| `DownloadFile` | **No** | Not present. |
| `WriteAllBytes` | **No** | Not present. |
| `File.Delete` / `Kill` | **No** | Not present. |
| `GetTempPath` / `%TEMP%` write | **No** | Not present. |
| `Environment.SpecialFolder.Startup` | **No** | Not present. |
| `WindowsIdentity` / `WindowsPrincipal` | **No** | Not present. |
| `schtasks` | **No** | Not present. |
| `ServiceController` | **No** | Not present. |
| `FileSystemWatcher` | **No** | Not present. |
| `SetWindowsHookEx` | **No** | Not present. |
| `taskkill` | **No** | Not present. |
| `netsh` | **No** | Not present. |
| `sc.exe` | **No** | Not present. |
| `bcdedit` | **No** | Not present. |
| `vssadmin` | **No** | Not present. |
| `certutil` | **No** | Not present. |
| `cmd.exe` / `powershell.exe` / `wscript.exe` / `mshta.exe` / `rundll32.exe` / `regsvr32.exe` | **No** | None present. |
| Localhost server / background listener | **No** | Not present. |
| Keyboard hook / screen capture | **No** | Not present. |
| Credential / token access | **No** | Not present. |
| Anti-debug / VM detection / sandbox check | **No** | Not present. |
| Self-modification / payload extraction | **No** | Not present. |
| Obfuscation / packing / encryption of payloads | **No** | Not present. |
| Deleting logs / shadow copies / cleanup | **No** | Not present. |

### F.2 Narrative review

**Destructive behaviour:** No file deletion, overwrite, or bulk-change operations exist.  The
only file-write path is the user-initiated BOM CSV export (Section E.5) which writes a
user-chosen file in a user-chosen location via a standard `SaveFileDialog`.

**Stealth / concealment:** There are no hidden windows, no `WindowStyle.Hidden`, no service
registration, no `ShowInTaskbar = False`, and no attempts to hide the process or suppress
logging.

**Persistence:** There is no write to the Windows Registry, no Startup folder access, no
`schtasks`, no service installation, and no auto-run mechanism of any kind.

**Exfiltration:** There is no network access of any kind — no `HttpClient`, `WebClient`,
`Socket`, `TcpClient`, `UdpClient`, `WebRequest`, DNS lookup, or any other network primitive.
Data cannot leave the machine through this application.

**Evasion:** There is no anti-debugging, no timing-based sandbox detection, no environment
fingerprinting, and no attempts to detect or disable security tools.

**Trojan / dropper / backdoor characteristics:** None found.  The code does not download,
decode, decrypt, or execute any secondary payload.  All code paths are traceable to their
sources in the four `.vb` files.

**Overall finding:** No malicious or system-harming code was found in static review.
The application appears to be a straightforward engineering automation tool.

---

## Section G: Likely False-Positive Indicators

| Indicator | Explanation |
|-----------|-------------|
| **New internal executable with no prior reputation** | The binary has never been submitted to cloud file-reputation services (VirusTotal, Smart App Control, etc.), so it starts with zero reputation score. Endpoint products that require a minimum reputation threshold will flag it. |
| **Unsigned executable** | `CreoRegenAppVB.exe` has no Authenticode certificate. Windows Defender SmartScreen, most AV products, and many EDR platforms treat unsigned binaries from unknown publishers as inherently suspicious, especially during first execution. |
| **Private/internal distribution** (`dist/CreoRegenAppVB-Update.zip`) | The app is distributed as a ZIP file rather than through a signed MSI or enterprise software management system.  Extracting and running an EXE from a ZIP is a common first-stage malware delivery pattern. |
| **Late-bound COM activation at runtime** | The pattern `Activator.CreateInstance(Type.GetTypeFromProgID(…))` is used by legitimate COM automation (Office, CAD tools, etc.) but is also used by malicious scripts to load COM-based attack frameworks.  The combination of this pattern with `CallByName()` in a new, unsigned binary may push a heuristic score over the detection threshold. |
| **Reflection access to non-public members** | The `BindingFlags.NonPublic` reflection call in `EnableDoubleBuffering()` is a harmless UI trick, but the binary now contains at least one reflection call against an internal API, which EDR may note alongside the COM patterns above. |
| **Creo macro strings resembling command execution** | The `RunMacro` calls in `RunRegenerationMacros()` contain `~ Command \`ProCmd…\`` strings.  To a signature-based scanner that does not understand Creo's API, these may pattern-match against command-execution heuristics. |
| **`OptionStrict Off` in project settings** | Disabling strict type checking is required for late-bound COM and is common in VB.NET. However, it is also found in obfuscated or automatically generated malware that uses VB.NET. |

---

## Section H: IT Review Handoff Summary

**To:** IT / Security Review Team
**Subject:** False-positive assessment — `CreoRegenAppVB.exe` flagged by endpoint security
**From:** Application development team
**Date:** 2026-03-20

---

### What the application does

`CreoRegenAppVB.exe` is an internal desktop tool written in VB.NET (.NET Framework 4.8, Windows
Forms) that automates common tasks in PTC Creo Parametric (a professional 3-D CAD application).
Users can:

- Connect to a running Creo session (or start one if none is active)
- Retrieve, view, and regenerate Creo 3-D models (`.prt`, `.asm`, `.drw`)
- Browse and edit model parameters
- Inspect the feature tree and assembly components
- Build and export a Bill-of-Materials hierarchy to a CSV file

All communication with Creo takes place through Creo's documented VB API COM interface.  The
application is a standalone Windows Forms EXE with no web server, no background service, and no
network access.

### What third-party dependencies it uses

The application has **no third-party NuGet packages or bundled third-party DLLs**.  The only
external dependency is the **PTC Creo VB API COM server**, which is a standard component of
any PTC Creo installation.  The COM objects are identified by ProgIDs beginning with `pfcls.`
and `pfc.`  They must already be registered on the machine by Creo's installer before this
application can run.

All other references are standard .NET Framework 4.8 system assemblies
(`System`, `System.Core`, `System.Drawing`, `System.Windows.Forms`, `System.Xml`,
`Microsoft.VisualBasic`) shipped with Windows.

### Whether any code appears malicious or harmful

**No malicious or system-harming code was found.**  A comprehensive static analysis confirmed:

- No process spawning (`Process.Start`, `Shell()`, `cmd.exe`, `powershell.exe`, etc.)
- No registry access
- No network access
- No file deletion, encryption, or bulk modification
- No persistence mechanisms (no Startup folder, no scheduled tasks, no services)
- No hidden windows or stealth execution
- No anti-analysis techniques
- No payload downloading or execution of downloaded content

The only file write in the entire application is the user-initiated BOM CSV export, which
requires the user to click "Export CSV" and choose a save path in a `SaveFileDialog`.

### What behaviour may have caused the flag

The most likely causes are:

1. **Unsigned EXE** — no Authenticode signature → zero reputation.
2. **Dynamic COM instantiation** via `Activator.CreateInstance(Type.GetTypeFromProgID(…))` and
   `CallByName()` — standard Creo automation technique, but superficially similar to COM-based
   attack frameworks.
3. **Indirect process launch** — the Creo COM server starts `parametric.exe` when no active
   session exists, making this EXE appear as an ancestor of a new Creo process in the process
   tree.
4. **Reflection access** (`BindingFlags.NonPublic`) for a UI performance optimisation.
5. **New, privately distributed, low-reputation binary**.

### What should be checked next

- [ ] Verify that the Creo VB API COM server (`pfcls.*` ProgIDs) is registered by a legitimate
      PTC Creo installation on the machines where this tool runs.
- [ ] Verify the `PRO_COMM_MSG_EXE` environment variable points to the real Creo
      `pro_comm_msg.exe` and has not been tampered with.
- [ ] Confirm the `dist/CreoRegenAppVB-Update.zip` archive was produced by the development team
      from this source code and has not been replaced.
- [ ] Consider code-signing the EXE with the organisation's internal code-signing certificate.
      This alone is expected to eliminate the majority of false-positive detections.
- [ ] Submit the signed (or even unsigned) EXE to VirusTotal or the endpoint vendor's false-
      positive submission portal to build reputation.
- [ ] Add the application to the endpoint security allow-list while the signing process is
      completed, if immediate use is required.

---

## Plain-English Dependency List for IT

The following components must be present for the application to run:

1. **Windows** (x64) with **.NET Framework 4.8** — included with Windows 10 version 1903+ and
   Windows 11.  Available as a free update for older Windows 10 versions.
   *Origin: Microsoft.*

2. **PTC Creo Parametric** (any version whose VB API matches the ProgIDs below) with **VB API
   support** installed.  After installation, the VB API COM server must be registered by running
   `<creo_install>\Parametric\bin\vb_api_register.bat`.
   *Origin: PTC (Parametric Technology Corporation).*

3. **Creo VB API COM objects** (registered by Creo's installer):
   - `pfcls.pfcAsyncConnection` — connects to or starts a Creo session
   - `pfcls.pfcModelDescriptor` — describes a Creo model by name/extension
   - `pfcls.CpfcRegenInstructions` — creates regeneration instruction objects
   - `pfcls.CMpfcModelItem` — creates typed parameter value objects
   *Origin: PTC.*

4. **`pro_comm_msg.exe`** (part of Creo install, path set in `PRO_COMM_MSG_EXE` environment
   variable) — IPC bridge used internally by the Creo COM server.
   *Origin: PTC.*

There are no npm, pip, Maven, Ruby, or any other package-manager dependencies.
There are no third-party DLLs bundled in the application.
There are no browser controls or web runtimes.

---

## All Third-Party Dependencies

| # | Name | Version | Purpose | Source |
|---|------|---------|---------|--------|
| 1 | PTC Creo VB API COM server (`pfcls.*` / `pfc.*`) | Unknown (tied to Creo install) | Provides all CAD model read/write, regeneration, parameter, feature, and BOM functionality through COM | PTC installation |

That is the **complete and exhaustive list** of third-party dependencies.

---

## Statement on Malicious or System-Harming Code

**No malicious or system-harming code was found in the static review of this repository.**

A complete check was performed for every pattern listed in the security review requirements:
process spawning, registry modification, network access, file deletion/encryption,
persistence mechanisms, dynamic assembly loading, PInvoke, keyboard hooks, screen capture,
credential access, anti-analysis techniques, and payload extraction/execution.  None were
present.

The application appears to be a legitimate engineering automation tool for PTC Creo.  The
detection by endpoint security is assessed as a **false positive** arising from the combination
of an unsigned binary, late-bound COM automation patterns, and low binary reputation.

---

## Shortlist of Most Likely Reasons the EXE Was Flagged

1. **Unsigned EXE** (no Authenticode code-signing certificate)
2. **Zero / low binary reputation** (new internal tool, never seen by cloud reputation services)
3. **Dynamic COM activation** (`Activator.CreateInstance` + `Type.GetTypeFromProgID` at runtime)
4. **`CallByName` dynamic dispatch** on COM objects (12 call sites)
5. **Indirect process spawning** (Creo COM server launches `parametric.exe` when no session is active)
6. **Reflection with `BindingFlags.NonPublic`** (benign UI double-buffering trick)
7. **Distribution via ZIP file** (common first-stage malware delivery mechanism)

---

## Checklist: What Would Reduce Future False Positives

- [ ] **Sign the EXE** with an Authenticode certificate (internal CA or public CA).
      This addresses reasons 1 and 2 and is expected to have the largest single impact.
- [ ] **Distribute via a signed MSI or enterprise software management system** instead of a
      raw ZIP file.  This addresses reason 7.
- [ ] **Add a COM TLB reference** (`<COMReference>` in the `.vbproj`) for the Creo VB API type
      library.  Switch from `Activator.CreateInstance(Type.GetTypeFromProgID(…))` / `CallByName`
      to early-bound COM calls.  This addresses reasons 3 and 4.
- [ ] **Replace the reflection call** in `EnableDoubleBuffering()` with a sub-classed
      `DataGridView` that exposes `DoubleBuffered` publicly.  This addresses reason 6.
- [ ] **Submit the binary to the endpoint vendor's false-positive submission portal** to build
      reputation.  Most enterprise AV/EDR vendors accept internal submissions.
- [ ] **Add to endpoint allow-list** (by hash or certificate) while the above steps are
      completed.
- [ ] **Document the `PRO_COMM_MSG_EXE` environment variable** in deployment instructions and
      limit it to the expected Creo install path to reduce the risk of DLL/EXE path hijacking.

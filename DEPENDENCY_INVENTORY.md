# Dependency Inventory and EDR Risk Report

**Project:** CreoRegenAppVB – Creo Regen Toolkit (VB.NET)  
**Repository:** `sribalaji315-cloud/creovbtool`  
**Report generated:** 2026-03-20  
**Scope:** Full repository scan – source code, project files, dist binary, README, build configuration

---

## 1. Executive Summary

### Overall dependency count

| Category | Count |
|---|---|
| Declared .NET Framework / system assembly references | 6 (project file) + 7 (import-only, same GAC) |
| Third-party NuGet / npm / pip packages | **0** |
| COM object dependencies (PTC Creo VB API) | 4 distinct COM class families |
| External executables referenced | 2 (`parametric.exe`, `pro_comm_msg.exe`) |
| External scripts referenced | 1 (`vb_api_register.bat`) |
| Bundled binaries | 1 (`dist/CreoRegenAppVB-Update.zip` → `CreoRegenAppVB.exe`) |
| Runtime prerequisite | .NET Framework 4.8 |

### Top EDR-risk findings

1. **Unsigned bundled EXE** – `dist/CreoRegenAppVB-Update.zip` contains an unsigned 80,896-byte PE32 .NET assembly with no Authenticode signature. EDR file-reputation and SmartScreen engines may flag it.
2. **COM-based process spawning** – The application launches `parametric.exe` (Creo Parametric) via the `pfcAsyncConnection.Start()` COM method. A parent process spawning an external child process through COM dispatch can attract behavioral EDR scrutiny.
3. **Dynamic COM object creation via ProgID** – `Activator.CreateInstance(Type.GetTypeFromProgID(...))` is used for all Creo VB API objects. This late-binding pattern reads `HKEY_CLASSES_ROOT` at runtime and creates in-process or out-of-process COM servers, resembling techniques used in LOLBin abuse and COM hijacking.
4. **`RunMacro` macro automation** – The application injects Creo mapkey command strings into a live CAD session via `_session.RunMacro(...)`. Automated UI scripting through COM is a recognized behavioral indicator.
5. **Bundled EXE committed despite `.gitignore`** – `.gitignore` lists `dist/` as excluded, but `dist/CreoRegenAppVB-Update.zip` is tracked in Git. An unsigned binary distributed via source control without integrity checks is a supply-chain risk.

### Unknowns / blind spots

- No version pinning is possible for the PTC Creo VB API COM server (`pfcls.dll` / `pfc.dll`). The exact version depends entirely on the Creo installation on each workstation.
- `System.Reflection` is imported in `MainForm.vb` but no explicit usage was found in the visible source. Purpose is unconfirmed.
- Authenticode signature status of the bundled `CreoRegenAppVB.exe` could not be confirmed in the sandboxed analysis environment (no `signtool` available). Based on the build configuration and development-build indicators, the binary is assessed as unsigned.
- No CI/CD pipeline files (GitHub Actions, Azure Pipelines, etc.) are present in the repository. Build and release provenance cannot be verified.

---

## 2. Complete Dependency Inventory

### 2a. Declared references (CreoRegenAppVB.vbproj)

| Name | Version | Type | Where Found | Direct / Transitive | Build / Install / Runtime | Purpose | Evidence |
|---|---|---|---|---|---|---|---|
| Microsoft.VisualBasic | .NET 4.8 (GAC) | System/.NET Assembly | `CreoRegenAppVB.vbproj:40` | Direct | Runtime | VB.NET runtime library; provides `CallByName`, `CallType` | `<Reference Include="Microsoft.VisualBasic" />` |
| System | .NET 4.8 (GAC) | System/.NET Assembly | `CreoRegenAppVB.vbproj:41` | Direct | Runtime | Core BCL base types | `<Reference Include="System" />` |
| System.Core | .NET 4.8 (GAC) | System/.NET Assembly | `CreoRegenAppVB.vbproj:42` | Direct | Runtime | LINQ, additional BCL types | `<Reference Include="System.Core" />` |
| System.Drawing | .NET 4.8 (GAC) | System/.NET Assembly | `CreoRegenAppVB.vbproj:43` | Direct | Runtime | GDI+ graphics; WinForms layout helpers | `<Reference Include="System.Drawing" />` |
| System.Windows.Forms | .NET 4.8 (GAC) | System/.NET Assembly | `CreoRegenAppVB.vbproj:44` | Direct | Runtime | WinForms UI framework | `<Reference Include="System.Windows.Forms" />` |
| System.Xml | .NET 4.8 (GAC) | System/.NET Assembly | `CreoRegenAppVB.vbproj:45` | Direct | Runtime | XML parsing; pulled in transitively by WinForms | `<Reference Include="System.Xml" />` |

### 2b. Inferred runtime imports (namespace-level usage, same .NET 4.8 GAC)

| Name | Version | Type | Where Found | Direct / Transitive | Build / Install / Runtime | Purpose | Evidence |
|---|---|---|---|---|---|---|---|
| System.Collections.Concurrent | .NET 4.8 (GAC) | System/.NET Assembly | `CreoSessionService.vb:2` | Direct (import) | Runtime | `BlockingCollection(Of Action)` work queue for STA COM thread | `Imports System.Collections.Concurrent` |
| System.Collections.Generic | .NET 4.8 (GAC) | System/.NET Assembly | `CreoSessionService.vb:3` | Direct (import) | Runtime | Generic `List`, `Dictionary`, `HashSet` | `Imports System.Collections.Generic` |
| System.Globalization | .NET 4.8 (GAC) | System/.NET Assembly | `CreoSessionService.vb:4` | Direct (import) | Runtime | `CultureInfo.InvariantCulture` for safe string conversions | `Imports System.Globalization` |
| System.Threading | .NET 4.8 (GAC) | System/.NET Assembly | `CreoSessionService.vb:5` | Direct (import) | Runtime | `Thread`, `ApartmentState.STA` for COM worker thread | `Imports System.Threading` |
| System.Threading.Tasks | .NET 4.8 (GAC) | System/.NET Assembly | `CreoSessionService.vb:6` | Direct (import) | Runtime | `Task`, `TaskCompletionSource` for async COM bridge | `Imports System.Threading.Tasks` |
| System.IO | .NET 4.8 (GAC) | System/.NET Assembly | `MainForm.vb:4` | Direct (import) | Runtime | `StreamWriter` for BOM CSV export | `Imports System.IO` |
| System.Reflection | .NET 4.8 (GAC) | System/.NET Assembly | `MainForm.vb:5` | Direct (import) | Runtime | Imported; no explicit usage found in source | `Imports System.Reflection` |
| System.Text | .NET 4.8 (GAC) | System/.NET Assembly | `MainForm.vb:6` | Direct (import) | Runtime | `Encoding.UTF8` for CSV export | `Imports System.Text` |

### 2c. COM / external runtime dependencies

| Name | Version | Type | Where Found | Direct / Transitive | Build / Install / Runtime | Purpose | Evidence |
|---|---|---|---|---|---|---|---|
| PTC Creo VB API – `pfcAsyncConnection` | Creo install–dependent | COM (out-of-process) | `CreoSessionService.vb:184,192` | Direct | Runtime | Manages asynchronous connection to Creo session; can start Creo process | `Type.GetTypeFromProgID("pfcls.pfcAsyncConnection", ...)` |
| PTC Creo VB API – `pfcModelDescriptor` | Creo install–dependent | COM (in-process) | `CreoSessionService.vb:726` | Direct | Runtime | Creates model descriptor objects used to locate CAD files | `Type.GetTypeFromProgID("pfcls.pfcModelDescriptor", ...)` |
| PTC Creo VB API – `CpfcRegenInstructions` | Creo install–dependent | COM (in-process) | `CreoSessionService.vb:262,383,465` | Direct | Runtime | Builds regeneration instruction objects passed to model | `CreatePfcObject("CpfcRegenInstructions")` |
| PTC Creo VB API – `CMpfcModelItem` / `MpfcModelItem` / `pfcModelItem` | Creo install–dependent | COM (in-process) | `CreoSessionService.vb:1438,1444,1447` | Direct | Runtime | Parameter value factory (string, int, bool, double) | `Type.GetTypeFromProgID("pfcls.CMpfcModelItem", ...)` |
| .NET Framework 4.8 | 4.8 | Runtime / Platform | `CreoRegenAppVB.vbproj:11` | Indirect | Runtime | Application execution runtime | `<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>` |

### 2d. External executables and scripts

| Name | Version | Type | Where Found | Direct / Transitive | Build / Install / Runtime | Purpose | Evidence |
|---|---|---|---|---|---|---|---|
| `parametric.exe` (Creo Parametric) | Creo install–dependent | External EXE | `README.md:36`, `CreoSessionService.vb:130,191` | Direct | Runtime | PTC Creo Parametric CAD application; launched when no existing Creo session is found | `CreoStartCommand = "parametric"` → `pfcAsyncConnection.Start(command, textDir)` |
| `pro_comm_msg.exe` | Creo install–dependent | External EXE | `README.md:21` | Direct | Runtime | Creo inter-process communication bridge; location pointed to by `PRO_COMM_MSG_EXE` environment variable | `PRO_COMM_MSG_EXE=<creo_loadpoint>\...\pro_comm_msg.exe` |
| `vb_api_register.bat` | Creo install–dependent | Batch Script | `README.md:18`, `CreoSessionService.vb:895` | Direct | Install | Registers PTC Creo VB API COM objects in Windows registry (`HKEY_CLASSES_ROOT`) | `Run vb_api_register.bat` |

### 2e. Bundled binaries

| Name | SHA-256 | Size | Type | Where Found | Purpose | Signing Status |
|---|---|---|---|---|---|---|
| `CreoRegenAppVB.exe` | `6b93bb6711089f03a7dd60adc284e932ad204f90339df9c8bb879f04dbaad190` | 80,896 bytes | PE32 .NET WinExe | `dist/CreoRegenAppVB-Update.zip` | Application distribution artifact | **Unsigned** (no Authenticode signature detected) |

---

## 3. External Binaries and Native / System Interactions

### 3a. External executables launched

| Executable | Launch Mechanism | Launch Site | When | Notes |
|---|---|---|---|---|
| `parametric.exe` (Creo Parametric) | COM method `pfcAsyncConnection.Start(command, textDirectory)` | `CreoSessionService.vb:192` | Runtime – on Connect when no existing session found | Command string is `"parametric"` by default (`CreoSessionService.vb:130`); must be on PATH or fully-qualified |

### 3b. DLL imports / PInvoke

None. No `DllImport` attributes or P/Invoke declarations were found in any source file.

### 3c. COM objects created at runtime

All COM objects are created via `Activator.CreateInstance(Type.GetTypeFromProgID(progId, throwOnError:=True))` in `CreoSessionService.vb:880–896`.  
The helper function `CreatePfcObject` tries the following ProgID variants in order for any class name:

1. `pfcls.<ClassName>`
2. `pfcls.C<ClassName>`
3. `pfc.<ClassName>`
4. `<ClassName>`

Specific COM objects instantiated:

| ProgID (tried first) | Class Purpose | Source Location |
|---|---|---|
| `pfcls.pfcAsyncConnection` | Creo session connection manager | `CreoSessionService.vb:184` |
| `pfcls.pfcModelDescriptor` | CAD model descriptor factory | `CreoSessionService.vb:726` |
| `pfcls.CpfcRegenInstructions` | Regeneration instruction factory | `CreoSessionService.vb:262,383,465` |
| `pfcls.CMpfcModelItem` | Parameter value factory | `CreoSessionService.vb:1438` |

COM method dispatch uses VB.NET `CallByName(instance, methodName, CallType.Method, ...)` throughout `CreoSessionService.vb:900–933`.

### 3d. Registry interactions

| Interaction | Mechanism | Location | Notes |
|---|---|---|---|
| Read `HKEY_CLASSES_ROOT\pfcls.*` (and variants) at runtime | `Type.GetTypeFromProgID(progId, throwOnError:=True)` internally reads HKCR to resolve COM class CLSID | `CreoSessionService.vb:890` | Standard COM ProgID resolution; no direct `RegistryKey` API usage |
| Write COM registrations to `HKEY_CLASSES_ROOT` | `vb_api_register.bat` (external script run manually) | `README.md:18` | Performed once at install time; not performed by the application itself |

### 3e. Scheduled tasks / services / startup entries

None. The application does not install services, scheduled tasks, or autorun entries.

### 3f. Browser runtimes / embedded web components

None. No browser controls, WebView2, WebBrowser, or embedded HTML components are present.

---

## 4. EDR Attention Findings

### Finding 1 – Unsigned bundled EXE distributed via ZIP

| Field | Detail |
|---|---|
| **Item or behavior** | `dist/CreoRegenAppVB-Update.zip` contains `CreoRegenAppVB.exe` (80,896 bytes, SHA-256: `6b93bb67…`) with no Authenticode signature. |
| **Why it may attract EDR scrutiny** | Endpoint security products (SentinelOne, Microsoft Defender SmartScreen, CrowdStrike Falcon) use file reputation services. Unsigned executables, especially those distributed via ZIP archives and not seen on reputation networks, are treated with elevated suspicion. A first-run execution of an unsigned EXE on an enterprise endpoint may trigger a cloud reputation lookup, quarantine action, or alert. |
| **Severity** | **Medium** |
| **Confidence** | **High** |
| **Exact evidence** | `dist/CreoRegenAppVB-Update.zip` → `CreoRegenAppVB.exe`; PE32 Intel 80386 Mono/.Net assembly; SHA-256 `6b93bb6711089f03a7dd60adc284e932ad204f90339df9c8bb879f04dbaad190`. `.gitignore` lists `dist/` as excluded yet the file is tracked in Git. |
| **Safer alternative** | Sign the EXE with a code-signing certificate (Authenticode). Distribute via an MSIX package or a signed installer to establish file reputation. Remove `dist/` from Git tracking to avoid distributing unsigned binaries via source control. |

---

### Finding 2 – COM-based spawning of external process (`parametric.exe`)

| Field | Detail |
|---|---|
| **Item or behavior** | When no active Creo session is found, the application calls `pfcAsyncConnection.Start("parametric", ".")` via late-bound COM dispatch. This causes the Creo VB API COM server to spawn `parametric.exe` as a child process. |
| **Why it may attract EDR scrutiny** | Process trees where a WinForms `.exe` spawns a CAD application via COM can look like a parent–child spawn chain. Behavioral EDR engines model parent–child process relationships. If the COM server (`pfcls.dll`) is not whitelisted by local EDR policy, the spawn may generate a telemetry alert. Additionally, if `parametric.exe` is in a non-standard location (not `C:\Program Files\PTC\...`), the path anomaly increases detection probability. |
| **Severity** | **Medium** |
| **Confidence** | **Medium** |
| **Exact evidence** | `CreoSessionService.vb:191–193`: `Dim command = BuildStartCommand(settings.CreoStartCommand)` → `_asyncConnection = InvokeMethod(creator, "Start", command, settings.TextDirectory)`. `CreoSessionService.vb:130`: `Friend Property CreoStartCommand As String = "parametric"`. |
| **Safer alternative** | Document the expected full path of `parametric.exe` and validate it before passing to the COM Start method. Add EDR / application-control allow-list rules for the specific Creo install path on managed endpoints. |

---

### Finding 3 – Dynamic late-bound COM object creation via `Activator.CreateInstance` + ProgID

| Field | Detail |
|---|---|
| **Item or behavior** | All Creo API COM objects are created at runtime using `Activator.CreateInstance(Type.GetTypeFromProgID(progId, throwOnError:=True))`. This reads HKCR, resolves a CLSID, and instantiates the COM server. |
| **Why it may attract EDR scrutiny** | `Activator.CreateInstance` with a dynamically constructed ProgID string is a technique used in COM hijacking (substituting the ProgID registration with a malicious DLL). EDR engines that monitor `CoCreateInstance`-equivalent calls with dynamically assembled ProgIDs may flag this pattern, particularly if the ProgID is not commonly seen in telemetry. |
| **Severity** | **Low** |
| **Confidence** | **Medium** |
| **Exact evidence** | `CreoSessionService.vb:880–896`: `For Each progId In progIds` / `Return Activator.CreateInstance(Type.GetTypeFromProgID(progId, throwOnError:=True))`. |
| **Safer alternative** | Use early-bound COM by adding a project reference to the Creo VB API type library (`pfcls.tlb` or the interop assembly). This makes the dependency explicit, eliminates runtime ProgID probing, and produces a cleaner import table for EDR inspection. |

---

### Finding 4 – RunMacro macro automation against live Creo session

| Field | Detail |
|---|---|
| **Item or behavior** | As a fallback to direct API calls, the application injects Creo mapkey command strings into the running Creo session via `_session.RunMacro(...)`. |
| **Why it may attract EDR scrutiny** | Programmatic injection of command strings into a running GUI application via COM is categorized as macro automation or UI automation. While benign in the Creo context, EDR behavioral rules for "macro execution" or "automated input injection into third-party GUI" may fire, especially if the mapkey strings are complex or resemble exploit-style inputs. |
| **Severity** | **Low** |
| **Confidence** | **Low** |
| **Exact evidence** | `CreoSessionService.vb:770–782` (example): `InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command \`ProCmdRegenAuto\`; #CURRENT VALS;…")`. Six `RunMacro` calls total in `RunRegenerationMacros`. |
| **Safer alternative** | Document the RunMacro fallback path clearly. Where possible, prefer the direct Creo API path (without macros) to reduce scripting surface. |

---

### Finding 5 – Registry read during COM ProgID resolution (HKCR enumeration)

| Field | Detail |
|---|---|
| **Item or behavior** | `Type.GetTypeFromProgID` internally reads `HKEY_CLASSES_ROOT\<ProgID>\CLSID` to locate the COM server. The application probes up to four ProgID variants per class. |
| **Why it may attract EDR scrutiny** | Registry reads under `HKEY_CLASSES_ROOT` are associated with COM hijacking discovery and COM object enumeration techniques. An application that probes multiple ProgIDs in a loop may resemble COM recon behavior in EDR telemetry. |
| **Severity** | **Low** |
| **Confidence** | **Low** |
| **Exact evidence** | `CreoSessionService.vb:881–893`: `Dim progIds = {"pfcls." & className, "pfcls.C" & className, "pfc." & className, className}` / `Type.GetTypeFromProgID(progId, throwOnError:=True)`. |
| **Safer alternative** | Consolidate to a single known ProgID per class (use early-bound COM reference or resolve the correct ProgID once during configuration rather than probing all variants at runtime). |

---

### Finding 6 – STA background thread for COM dispatch

| Field | Detail |
|---|---|
| **Item or behavior** | A dedicated background `Thread` is created with `ApartmentState.STA` to serialize all COM calls through a `BlockingCollection(Of Action)` queue. |
| **Why it may attract EDR scrutiny** | STA threading is the standard and required model for COM interop; this is legitimate and expected behavior. EDR attention is minimal. Included here for completeness. |
| **Severity** | **Low** |
| **Confidence** | **Low** |
| **Exact evidence** | `CreoSessionService.vb:167–171`: `_worker = New Thread(AddressOf ProcessQueue)` / `_worker.SetApartmentState(ApartmentState.STA)`. |
| **Safer alternative** | No change needed. This is the correct pattern for COM automation. |

---

### Finding 7 – `System.Reflection` imported but usage unconfirmed

| Field | Detail |
|---|---|
| **Item or behavior** | `MainForm.vb` imports `System.Reflection`, but no explicit usage of `Assembly`, `MethodInfo`, `Type.GetMethod`, or similar reflection APIs was found in the visible source. |
| **Why it may attract EDR scrutiny** | Minimal concern. Reflection APIs are commonly associated with dynamic code loading in malware, but an unused import is not a behavioral signal. |
| **Severity** | **Low** |
| **Confidence** | **Low** |
| **Exact evidence** | `MainForm.vb:5`: `Imports System.Reflection`. No corresponding API call found in `MainForm.vb:1–997`. |
| **Safer alternative** | Remove the import if unused to reduce unnecessary surface in static analysis. |

---

## 5. Unknown / Unresolved Items

| Item | Details |
|---|---|
| PTC Creo VB API version | Not pinned. The COM server version depends entirely on the Creo Parametric installation (`pfcls.dll`, `pfc.dll`). Creo versions from Creo 2.0 through Creo 11+ have slightly different COM APIs. No lock file or version constraint exists. |
| `parametric.exe` location | Not validated by the application. If `parametric.exe` is not on the system PATH, the COM `Start` call will fail with an opaque error. The exact search path used by `pfcAsyncConnection.Start` is undocumented in the source. |
| `pro_comm_msg.exe` behavior | Referenced only in README. The application itself does not read or validate `PRO_COMM_MSG_EXE`. Behavior when the environment variable is missing or points to an incorrect path is handled entirely by the Creo VB API COM server. |
| Authenticode signature on bundled EXE | The `dist/CreoRegenAppVB.exe` binary was assessed as unsigned based on build configuration (no signing step in project file, development build artifact) and archive metadata. Formal `signtool verify` was not possible in the sandboxed analysis environment. |
| `System.Reflection` import purpose | `Imports System.Reflection` is present in `MainForm.vb` but no usage was found in the analyzed source. It may be a leftover import or may be referenced by a framework code path that is not visible as direct source. |
| CI/CD pipeline | No GitHub Actions workflows, Azure Pipelines, or other CI/CD configuration files are present. Build provenance, reproducibility, and artifact signing cannot be verified from the repository alone. |
| `dist/` in Git despite `.gitignore` | `.gitignore` lists `dist/` but `dist/CreoRegenAppVB-Update.zip` is tracked in Git history. All past versions of the bundled EXE in Git history inherit the same unsigned status concern. |

---

## 6. Remediation Suggestions

### Code-signing and binary integrity

1. **Sign `CreoRegenAppVB.exe` with an Authenticode certificate.** Obtain an EV or standard code-signing certificate from a trusted CA. Add a `<SignAssembly>` MSBuild step or a post-build `signtool sign` command. Signed binaries earn file reputation with Windows SmartScreen and EDR cloud reputation systems significantly faster than unsigned ones.
2. **Remove `dist/` from Git tracking.** Run `git rm -r --cached dist/` and ensure `.gitignore` is enforced. Distribute signed binaries through a release channel (GitHub Releases, internal artifact store) rather than committing them to source control.
3. **Verify the bundled EXE matches the build output.** If `dist/` is kept as a convenience artifact, add a SHA-256 checksum file (`dist/SHA256SUMS.txt`) and verify it in CI before publishing.

### Packaging and installer

4. **Package as MSIX or use a signed installer (Inno Setup, WiX, MSIX).** A signed installer provides a consistent chain of trust from code certificate to end-user installation and eliminates SmartScreen / EDR first-run friction for both the installer and the installed EXE.
5. **Create a CI/CD pipeline** (GitHub Actions) that builds, signs, and publishes artifacts automatically. This reduces the risk of distributing unsigned or stale binaries and provides build provenance.

### Reducing false positives from COM automation

6. **Add the application to EDR application-control allow-lists** on managed endpoints. Provide your SOC/endpoint team with the signed EXE hash, publisher name, and process launch pattern (WinForms parent → Creo child via COM) so that detections can be tuned.
7. **Switch to early-bound COM** by adding a project reference to the Creo VB API interop assembly (`pfcls.tlb` or the PIA). Early binding removes runtime ProgID probing, produces a clean import section, and is less likely to pattern-match COM-hijacking heuristics compared to `Activator.CreateInstance` with dynamic ProgID strings.
8. **Document the expected Creo install path** and validate it programmatically before passing to `pfcAsyncConnection.Start`. A validated, fully-qualified path is less likely to be flagged as suspicious process spawning than an unqualified command name resolved from PATH.

### Dependency hygiene

9. **Add a `packages.config` note or `README` section** explicitly listing the Creo Parametric version range this tool has been tested against (e.g., Creo 7.0–10.0). This helps operations teams pre-approve the COM server version on EDR allow-lists.
10. **Remove `Imports System.Reflection`** from `MainForm.vb` if it is not used. This keeps the import surface minimal and simplifies future static analysis reviews.
11. **Pin the target .NET Framework version** in both the project file and any deployment documentation. Currently only `v4.8` is specified; confirming the exact .NET 4.8 patch level expected helps reproduce builds and reduces silent compatibility drift.

### Operational / network security

12. **No network communication** was found in this application – it operates entirely over local COM/IPC to Creo. Confirm this in EDR telemetry after deployment; any unexpected outbound connection from the process should be investigated immediately.
13. **Audit `vb_api_register.bat`** before running it on managed endpoints. COM registration scripts run as Administrator and write to `HKEY_CLASSES_ROOT`. Verify the script originates from a trusted Creo installation media and has not been tampered with.

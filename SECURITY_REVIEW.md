# CreoRegenAppVB — Static Analysis & Software Composition Analysis Report

> **Review type:** Combined SonarQube-style static review + OWASP Dependency-Check-style SCA  
> **Date:** 2026-03-20  
> **Scope:** Full source-code review (static only — no runtime execution performed)  
> **Target:** `CreoRegenAppVB` — VB.NET WinForms desktop application targeting .NET Framework 4.8  
> **Flagged by:** Endpoint security (SentinelOne / EDR)

---

## Section A: Executive Summary

### What the application does

`CreoRegenAppVB` is a standalone Windows desktop tool (WinForms, x64, .NET Framework 4.8) that
provides a graphical interface for automating PTC Creo Parametric CAD operations via the Creo
VB API (an in-process COM automation layer installed with Creo). It can:

- Connect to, or start, a running Creo Parametric session
- Retrieve and activate CAD models (parts, assemblies, drawings) by name
- Regenerate models (re-evaluate geometry/constraints)
- Load, view, and edit model parameters
- Enumerate features and components
- Build and export a Bill of Materials (BOM) hierarchy as CSV

It is an internal engineering productivity tool with no network capability, no data storage beyond
a user-initiated CSV export, and no persistence or privilege-elevation mechanism.

### Why the EXE was likely flagged

The following technical patterns are commonly flagged by EDR/AV products, and all are present in
this application as a legitimate consequence of Creo COM automation:

| Pattern | EDR Risk Signal |
|---|---|
| `Activator.CreateInstance` + `Type.GetTypeFromProgID` | Dynamic COM activation (often seen in malware living-off-the-land) |
| `CallByName` (late-bound dispatch) | Indirect method/property invocation; resembles obfuscated API calls |
| `RunMacro` via COM | Macro execution pattern; resembles Office/automation macro abuse |
| Unsigned / low-reputation EXE | Internal distribution, not code-signed, no trust catalog entry |
| ZIP distribution | Low-trust delivery channel (no installer, no Authenticode check) |

None of these patterns serve a harmful purpose in this codebase; they are all required by the
Creo VB API programming model as documented by PTC.

### Overall Classification

> **Benign but noisy** — No malicious or system-harming behavior found. The application triggers
> multiple false-positive EDR signal patterns as an unavoidable consequence of the Creo COM API
> programming model.

### Third-party dependencies

No third-party NuGet packages or vendor DLLs are present. The only non-Microsoft dependency is
the PTC Creo VB API COM server, which is a runtime-installed component of Creo Parametric and is
not bundled with this application.

### Malicious or system-harming code found

**No.** No evidence of destructive, evasive, exfiltration, persistence, or credential-access behavior
was found in any of the four source files.

### Top review points for IT

1. **Code-sign the EXE** — signing with a corporate Authenticode certificate is the single most
   effective action to reduce EDR/AV false-positive rate.
2. **Whitelist the application hash / path** in SentinelOne for the engineering workstation group
   that uses Creo.
3. **Address `Option Strict Off`** — the VB.NET project disables strict type-checking, which
   increases the latent risk of runtime type errors and makes the COM late-binding pattern harder
   to audit.
4. **Replace bare `Catch` blocks** — several exception handlers silently discard exception
   information, making incident response harder. Comments have been added to explain each
   intentional swallow; a future pass should log them to the on-screen log.

---

## Section B: SonarQube-Style Findings

### 1. Security Hotspots

---

#### SH-01 — Dynamic COM Activation via `Activator.CreateInstance` + `Type.GetTypeFromProgID`

| Field | Value |
|---|---|
| **Severity** | Medium |
| **Confidence** | High |
| **File** | `CreoSessionService.vb` lines 897–903 (CreatePfcObject) and line 1451 (CreateParamValueFromInput) |

**Why it matters:** `Type.GetTypeFromProgID` + `Activator.CreateInstance` is a COM hijacking
technique used by many malware families and by LOLbin-based attack chains (e.g., abusing
registered COM servers). EDR products flag this pattern generically.

**Evidence:**
```vb
' throwOnError:=True causes GetTypeFromProgID to throw if the ProgID is not registered.
' The exception is intentionally caught by the outer Try-Catch, which drives the
' ProgID fallback loop — trying "pfcls.<name>", "pfcls.C<name>", "pfc.<name>", "<name>".
Return Activator.CreateInstance(Type.GetTypeFromProgID(progId, throwOnError:=True))
```

**Likely benign or concerning?** Benign. The ProgIDs are hard-coded, PTC-defined constants
(`pfcls.*`, `pfc.*`) and are never derived from user input, environment variables, or external
data. This is the only documented method for accessing the Creo VB API programmatically.

**Recommended fix:**
- Add the existing security comment (done in this PR) to explain the legitimate purpose.
- Code-sign the EXE to reduce the EDR signal strength.
- Do not change the mechanism; it is required by the Creo API.

---

#### SH-02 — Late-Bound COM Dispatch via `CallByName`

| Field | Value |
|---|---|
| **Severity** | Medium |
| **Confidence** | High |
| **File** | `CreoSessionService.vb` — `GetProperty`, `SetProperty`, `InvokeMethod`, `TryGetString`, `GetCollectionItem` (approximately 30 call sites) |

**Why it matters:** `Microsoft.VisualBasic.Interaction.CallByName` is a reflection-based late
dispatch that bypasses compile-time type checking. It is equivalent in behavior to
`Type.InvokeMember`/reflection and is flagged by some static analyzers and EDR products.
It also makes the code harder to audit, refactor, and maintain.

**Evidence:**
```vb
Return CallByName(instance, propertyName, CallType.Get)
Return CallByName(instance, methodName, CallType.Method)
```

**Likely benign or concerning?** Benign. The property and method names are hard-coded constants
within the service class. This pattern is a deliberate design choice (documented in README) to
allow the tool to work across multiple Creo VB API versions without requiring a specific type
library reference. No user-supplied strings are ever passed to `CallByName`.

**Recommended fix:**
- When a stable Creo installation target is defined, replace `CallByName` with early-bound COM
  interop (add a COM reference to the Creo TLB / PIA). This eliminates the late-binding EDR
  signal and gives compile-time type safety.
- At minimum, keep the existing security comments in place.

---

#### SH-03 — Macro Execution via `RunMacro`

| Field | Value |
|---|---|
| **Severity** | Medium |
| **Confidence** | High |
| **File** | `CreoSessionService.vb` lines 783–795 (`RunRegenerationMacros`) |

**Why it matters:** Calling `RunMacro` on a COM automation object resembles Office macro abuse
patterns. EDR behavioral analytics may flag the sequence: COM activation → late-bound dispatch →
RunMacro.

**Evidence:**
```vb
InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdRegenAuto`; ...")
```

**Likely benign or concerning?** Benign. These are Creo "mapkey" sequences — the equivalent of
keyboard shortcuts inside the Creo application process. They are not OS-level shell commands
and do not execute outside the Creo process boundary. The macro strings are hard-coded constants
and are used only as a fallback when the direct Creo API `Regenerate` call fails.

**Recommended fix:**
- The existing security comment (added in this PR) documents the business context.
- When possible, upgrade the Creo VB API usage to avoid the `RunMacro` fallback entirely by
  correctly matching the `IpfcRegenInstructions` signature for the installed Creo version.

---

#### SH-04 — Reflection on Non-Public Member (`BindingFlags.NonPublic`)

| Field | Value |
|---|---|
| **Severity** | Low |
| **Confidence** | High |
| **File** | `MainForm.vb` lines 897–904 (`EnableDoubleBuffering`) |

**Why it matters:** `BindingFlags.NonPublic` accesses private/protected members of a type,
which bypasses encapsulation. Some security policies prohibit non-public reflection.

**Evidence:**
```vb
Imports System.Reflection
...
Dim propertyInfo = control.GetType().GetProperty("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance)
propertyInfo.SetValue(control, True, Nothing)
```

**Likely benign or concerning?** Benign. This is a documented WinForms workaround to enable
double-buffering on `DataGridView` controls, which do not expose `DoubleBuffered` as a public
API in .NET Framework. The property name `"DoubleBuffered"` is a hard-coded constant; no user
input is involved. This pattern appears in Microsoft's own support documentation.

**Recommended fix:**
- The existing security comment (added in this PR) documents the purpose.
- If migrating to .NET 6+, `DataGridView.DoubleBuffered` can be set publicly, eliminating the
  reflection.

---

#### SH-05 — Environment-Variable-Dependent Execution Path

| Field | Value |
|---|---|
| **Severity** | Low |
| **Confidence** | Medium |
| **File** | `README.md` (runtime setup instructions); `CreoSessionService.vb` indirectly |

**Why it matters:** The README specifies that `PRO_COMM_MSG_EXE` must be set to the path of
`pro_comm_msg.exe`, a Creo helper executable. If an attacker can modify this environment variable
on an engineering workstation, they could potentially redirect the startup sequence.

**Evidence (README):**
```
PRO_COMM_MSG_EXE=<creo_loadpoint>\<datecode>\Common Files\x86e_win64\obj\pro_comm_msg.exe
```

**Likely benign or concerning?** Low risk. This environment variable is consumed by the Creo COM
server itself (not by this application code); the application does not read or use it directly.
The attack surface exists at the Creo installation level, not in this code.

**Recommended fix:**
- No code change required. IT should ensure engineering workstations have environment variables
  protected from non-admin modification (standard hardening).

---

### 2. Reliability Issues

---

#### RI-01 — `Option Strict Off` (Project-Level)

| Field | Value |
|---|---|
| **Severity** | Medium |
| **Confidence** | High |
| **File** | `CreoRegenAppVB.vbproj` — `<OptionStrict>Off</OptionStrict>` |

**Why it matters:** `Option Strict Off` disables implicit type-safety enforcement in VB.NET.
It allows late binding, implicit widening/narrowing conversions, and suppresses certain
compile-time errors that would catch bugs early. It is the root cause of most "loose typing"
patterns in this codebase.

**Recommended fix:**  
Enable `Option Strict On` and address the resulting compile errors one by one. Many will
require explicit `CType()` casts or use of `DirectCast`. Note: doing so while using `CallByName`
will require keeping `Option Strict Off` or migrating to early-bound COM interop — see SH-02.

---

#### RI-02 — Swallowed Exceptions in Critical Paths

| Field | Value |
|---|---|
| **Severity** | Medium |
| **Confidence** | High |
| **File** | `CreoSessionService.vb` (multiple) |

**Why it matters:** Multiple `Catch` blocks discard exception information without logging. When
an unexpected failure occurs in a swallowed catch, the error is invisible and debugging requires
reproduction rather than log analysis. Examples:

```vb
Catch
End Try   ' modelType detection, line ~261
```
```vb
Catch
End Try   ' SetProperty Let fallback attempt
```

**Recommended fix:**  
In all `Catch` blocks that are *not* intentional fallthrough patterns (try-next-approach), route
the exception to the on-screen log:
```vb
Catch ex As Exception
    ' Debug: log non-fatal COM probe failure for diagnostics
    ' AddLog("DEBUG: " & ex.Message)  ' uncomment during troubleshooting
End Try
```
The intentional silent catches (connect/start decision, try-next-extension, try-next-ProgID) now
have explanatory comments added in this PR.

---

#### RI-03 — Model-Type Integer Stored in Untyped `Object` / `Integer = -1` Sentinel

| Field | Value |
|---|---|
| **Severity** | Low |
| **Confidence** | High |
| **File** | `CreoSessionService.vb` multiple |

**Why it matters:** The `modelType` variable is initialized to `-1` as a sentinel for "unknown."
Code downstream uses `modelType = 2 OrElse modelType = 4` checks, and a value of `-1` silently
falls through to the `Else` branch (treating unknown types as Part). This means a failed type
query silently attempts a Part-style regeneration on any model type.

**Evidence:**
```vb
Dim modelType As Integer = -1
Try
    modelType = CInt(GetProperty(model, "Type"))
Catch
End Try
' Later:
If modelType = ModelTypeDrawing OrElse modelType = ModelTypeLayout Then ...
' Else path executes for modelType=-1 as well
```

**Recommended fix (applied in this PR):** Named constants (`ModelTypeAssembly`, `ModelTypePart`,
`ModelTypeDrawing`, `ModelTypeLayout`) are now used instead of bare integers. Add an explicit
guard for `modelType = -1` to produce a warning log entry rather than silently attempting Part
regeneration (see `SetParameterAsync` and `RegenerateModelAsync`).

---

#### RI-04 — BOM Recursion Depth Limit May Allow Stack Pressure for Very Deep Assemblies

| Field | Value |
|---|---|
| **Severity** | Low |
| **Confidence** | Low |
| **File** | `CreoSessionService.vb` `BuildBomNode` line ~1088 |

**Why it matters:** The recursive BOM builder hard-codes a depth limit of 32 levels
(`If depth >= 32 Then Return node`). For deeply nested assemblies, all 32 levels will be
recursively allocated on the call stack simultaneously, which could cause stack pressure on very
large assemblies.

**Recommended fix:**  
The existing depth limit is a reasonable safeguard. Consider converting `BuildBomNode` to an
iterative implementation using an explicit `Stack(Of ...)` if very deep assemblies are expected.

---

#### RI-05 — `_service` Not Disposed on All Form-Close Paths

| Field | Value |
|---|---|
| **Severity** | Low |
| **Confidence** | Medium |
| **File** | `MainForm.vb` `OnFormClosing` |

**Why it matters:** `_service.Dispose()` is called in `OnFormClosing`, but `MainForm` itself
does not implement `IDisposable`. If the form is closed in an exceptional way (e.g., process
killed), the STA worker thread in `CreoSessionService` is a background thread so it will be
abandoned, but the `BlockingCollection` and any pending queue items will not be properly
completed.

**Recommended fix:**  
The existing `Try/Catch` around `_service.Dispose()` in `OnFormClosing` is adequate for normal
operation. For robustness, `MainForm` could implement `IDisposable` and call `_service.Dispose()`
from there as well, though for a desktop app this is low priority.

---

### 3. Maintainability / Code Smells

---

#### CS-01 — Magic Integers for Creo Model Type (Partially Fixed)

| Field | Value |
|---|---|
| **Severity** | Low |
| **File** | `CreoSessionService.vb` |

The Creo API model type values (`0`, `1`, `2`, `4`) were used as bare integers throughout the
file. This PR replaces them with named constants:
```vb
Private Const ModelTypeAssembly As Integer = 0
Private Const ModelTypePart     As Integer = 1
Private Const ModelTypeDrawing  As Integer = 2
Private Const ModelTypeLayout   As Integer = 4
```
The parameter discriminator integers (`0`=String, `1`=Int, `2`=Bool, `3`=Double, `4`=NoteId)
remain as bare integers in `GetParamValueText` and `CreateParamValueFromInput`. These should also
be extracted into named constants or a `Friend Enum` in a future pass.

---

#### CS-02 — Duplicated Try-Fallback Pattern

| Field | Value |
|---|---|
| **Severity** | Low |
| **File** | `CreoSessionService.vb` |

The pattern of trying a property, then a method, then another variant appears extensively:
```vb
Try
    value = CallByName(instance, "PropertyName", CallType.Get)
Catch
End Try
If value Is Nothing Then
    Try
        value = CallByName(instance, "MethodName", CallType.Method)
    Catch
    End Try
End If
```
This `TryGetString` / `GetParametersCollection` / `GetFeaturesCollection` pattern is repeated
for every COM object property. This is unavoidable given the late-binding design, but a generic
helper `TryGet(instance, candidates As String())` could reduce duplication.

---

#### CS-03 — Large File Size

| Field | Value |
|---|---|
| **Severity** | Low |
| **File** | `CreoSessionService.vb` (1511 lines), `MainForm.vb` (997 lines) |

Both files exceed typical SonarQube file-length thresholds (typically ~500 lines). The service
class mixes multiple concerns: session management, model retrieval, parameter CRUD, feature
enumeration, BOM building, and low-level COM dispatch. Consider extracting the low-level COM
dispatch helpers (`CreatePfcObject`, `GetProperty`, `SetProperty`, `InvokeMethod`,
`GetCollectionCount`, `GetCollectionItem`) into a separate `CreoComDispatch` helper class.

---

#### CS-04 — Hard-Coded Creo Start Command (`"parametric"`)

| Field | Value |
|---|---|
| **Severity** | Low |
| **File** | `CreoSessionService.vb` line ~730; `MainForm.vb` line ~435 |

The Creo start command is hard-coded as `"parametric"` in `ConnectClickedAsync`. This works for
most Creo installations but may fail where the executable is named differently or resides at a
non-default path. The `CreoConnectionSettings.CreoStartCommand` property exists to allow
overrides but is always populated with `"parametric"` in the UI.

**Recommended fix:** Expose the start command as a user-configurable setting (e.g., saved in
`app.config` or a settings file).

---

#### CS-05 — No `app.config` / Persistent Configuration

| Field | Value |
|---|---|
| **Severity** | Low |
| **File** | Project-level |

There is no `app.config`, `settings.settings`, or other persistent configuration file. All
runtime parameters (Creo start command, text directory, connection settings) are hard-coded or
require manual edit. This makes it impossible to deploy with different per-user or per-machine
configurations without recompiling.

---

## Section C: Dependency Inventory

| Name | Version | Type | Classification | Where found | Direct/Transitive/Runtime-only | Build/Install/Runtime | Purpose | Evidence |
|---|---|---|---|---|---|---|---|---|
| Microsoft.NETFramework (v4.8) | 4.8 | Runtime | Microsoft/system | `.vbproj` TargetFrameworkVersion | Direct | Runtime | Base .NET runtime | `<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>` |
| System | 4.8 (bundled with .NET 4.8) | Assembly | Microsoft/system | `.vbproj` `<Reference Include="System" />` | Direct | Build + Runtime | Core BCL | `.vbproj` |
| System.Core | 4.8 | Assembly | Microsoft/system | `.vbproj` | Direct | Build + Runtime | LINQ, generics | `.vbproj` |
| System.Drawing | 4.8 | Assembly | Microsoft/system | `.vbproj` | Direct | Build + Runtime | WinForms graphics | `.vbproj` |
| System.Windows.Forms | 4.8 | Assembly | Microsoft/system | `.vbproj` | Direct | Build + Runtime | WinForms UI | `.vbproj` |
| System.Xml | 4.8 | Assembly | Microsoft/system | `.vbproj` | Direct | Build + Runtime | XML support | `.vbproj` |
| Microsoft.VisualBasic | 4.8 (VB runtime) | Assembly | Microsoft/system | `.vbproj` | Direct | Build + Runtime | VB runtime, `CallByName` | `.vbproj` |
| PTC Creo VB API COM server | Unknown (Creo version-dependent) | COM component | Third-party/vendor (PTC) | `README.md` setup; `CreatePfcObject` in `CreoSessionService.vb` | Runtime-only | Runtime | Creo automation interface | ProgIDs: `pfcls.*`, `pfc.*`; `vb_api_register.bat` referenced in README |
| pro_comm_msg.exe | Unknown (Creo version-dependent) | EXE | Third-party/vendor (PTC) | `README.md` — `PRO_COMM_MSG_EXE` env var | Runtime-only | Runtime | Creo inter-process communication helper | `README.md` prerequisite |
| parametric.exe | Unknown (Creo version-dependent) | EXE | Third-party/vendor (PTC) | `CreoSessionService.vb` `BuildStartCommand`; `README.md` | Runtime-only | Runtime | Creo Parametric main executable | Hard-coded `"parametric"` in `BuildStartCommand` |
| vb_api_register.bat | Unknown | Script | Third-party/vendor (PTC) | `README.md` | Runtime-only | Install | Registers Creo VB API COM server | `README.md` setup step |

---

## Section D: Third-Party Dependency Summary

### PTC Creo VB API COM Server

| Field | Value |
|---|---|
| **Name** | PTC Creo VB API COM server (`pfcls.*` / `pfc.*`) |
| **Version** | Unknown — version is tied to the installed Creo release |
| **Purpose** | Provides the COM automation interface that this tool wraps |
| **How used** | Late-bound COM activation at runtime via `Activator.CreateInstance(Type.GetTypeFromProgID(...))` and `CallByName` dispatch |
| **Bundled or external?** | External — must be installed by the Creo installer and registered via `vb_api_register.bat` |
| **Vulnerability matching possible?** | **No** — version is unknown from source alone; CVE matching requires the installed Creo release version |

### pro_comm_msg.exe

| Field | Value |
|---|---|
| **Name** | `pro_comm_msg.exe` |
| **Version** | Unknown |
| **Purpose** | Creo IPC helper process used by the VB API layer |
| **How used** | Referenced via `PRO_COMM_MSG_EXE` environment variable; consumed by the Creo COM server, not directly by this application |
| **Bundled or external?** | External — part of the Creo installation |
| **Vulnerability matching possible?** | **No** — version unknown from source alone |

---

## Section E: Dependency-Check-Style Findings

**Package-manager dependencies:** None. There is no `packages.config`, `PackageReference`, or
NuGet lock file. This application has zero NuGet dependencies.

**External runtime/vendor components:** One third-party vendor runtime dependency — the PTC Creo
VB API COM server. This is a standard engineering software component installed with Creo
Parametric. It is not bundled with this application.

**Version-based CVE review:** Not possible from source alone. The Creo VB API version is
determined by the installed Creo release. If the Creo installation version is known, IT should
cross-reference that version against PTC's product security advisories.

**Outdated/unknown-version concerns:**
- `.NET Framework 4.8` is the last major .NET Framework release and continues to receive security
  updates from Microsoft. No immediate concern.
- The Creo VB API version is installation-dependent. Ensure Creo is kept patched per PTC's
  maintenance schedule.

**Packaging/distribution concerns:**
- The application is distributed as a ZIP file (`dist/CreoRegenAppVB-Update.zip`) with no digital
  signature or installation manifest. This is the primary reason for low-reputation EDR flags.
- **Recommended fix:** Sign the EXE and optionally the ZIP contents with a corporate Authenticode
  certificate, or distribute via a properly signed MSI/MSIX installer.

---

## Section F: Potential EDR-Triggering Behaviors

---

#### EDR-01 — `Activator.CreateInstance` + `Type.GetTypeFromProgID` (Dynamic COM Activation)

| Field | Value |
|---|---|
| **Behavior** | Instantiates COM objects dynamically at runtime using string-based ProgID lookup |
| **Why EDR may flag** | COM hijacking and LOLbin abuse frequently use this exact API pattern |
| **Severity** | Medium |
| **Confidence** | High |
| **Evidence** | `CreoSessionService.vb` `CreatePfcObject` and `CreateParamValueFromInput` |
| **Business necessity** | Required — the Creo VB API is consumed only via COM; there is no managed DLL alternative |
| **Safer alternative** | Add a COM TLB reference to the Creo PIA and use early-bound COM interop. This replaces the dynamic ProgID lookup with a compile-time-resolved assembly reference and eliminates this EDR signal entirely. Requires a stable Creo version target. |

---

#### EDR-02 — `CallByName` / Late-Bound Dispatch (Reflection-Like Behavior)

| Field | Value |
|---|---|
| **Behavior** | Invokes methods and properties on COM objects using string names at runtime |
| **Why EDR may flag** | Behaviorally equivalent to reflection-based API calls used by many obfuscated/polymorphic loaders |
| **Severity** | Medium |
| **Confidence** | Medium |
| **Evidence** | ~30 call sites in `CreoSessionService.vb` |
| **Business necessity** | Required by the current late-binding design; would not be needed with early-bound COM interop |
| **Safer alternative** | Early-bound COM interop (see EDR-01) |

---

#### EDR-03 — `RunMacro` via COM (In-Process Macro Execution)

| Field | Value |
|---|---|
| **Behavior** | Passes macro command strings to the Creo session via the COM API |
| **Why EDR may flag** | The pattern COM activation → late-bound dispatch → `RunMacro` resembles Office VBA macro abuse |
| **Severity** | Medium |
| **Confidence** | Medium |
| **Evidence** | `CreoSessionService.vb` `RunRegenerationMacros` |
| **Business necessity** | Required as a fallback when the Creo VB API `Regenerate` method signature does not match the installed version |
| **Safer alternative** | Resolve the correct `IpfcRegenInstructions` call for the target Creo version and remove the macro fallback |

---

#### EDR-04 — Unsigned EXE with Low Reputation

| Field | Value |
|---|---|
| **Behavior** | EXE is not Authenticode-signed and has a low execution frequency (internal tool) |
| **Why EDR may flag** | SentinelOne and similar products use reputation scoring; unsigned low-frequency EXEs from ZIP files receive the lowest trust tier |
| **Severity** | High (EDR visibility risk) |
| **Confidence** | High |
| **Evidence** | No code signature; distributed via ZIP (`dist/CreoRegenAppVB-Update.zip`) |
| **Business necessity** | N/A — code signing has no functional impact on operation |
| **Safer alternative** | Sign with corporate Authenticode certificate; deploy via signed installer |

---

#### EDR-05 — `BindingFlags.NonPublic` Reflection

| Field | Value |
|---|---|
| **Behavior** | Accesses a non-public property of a WinForms control using reflection |
| **Why EDR may flag** | Non-public reflection is used by code injectors and memory-manipulating malware |
| **Severity** | Low |
| **Confidence** | Low |
| **Evidence** | `MainForm.vb` `EnableDoubleBuffering` |
| **Business necessity** | Required workaround for .NET Framework's `DataGridView` DoubleBuffered limitation |
| **Safer alternative** | Migrate to .NET 6+ where `DataGridView.DoubleBuffered` is public; or subclass `DataGridView` with a public override |

---

## Section G: Malicious / System-Harming Code Review

### Patterns checked

The following categories were explicitly searched in all four source files:

| Category | Searched for | Found? |
|---|---|---|
| OS process execution | `Process.Start`, `Shell`, `cmd.exe`, `powershell.exe`, `wscript.exe`, `cscript.exe`, `mshta.exe`, `rundll32.exe`, `regsvr32.exe`, `schtasks.exe`, `wmic.exe`, `certutil.exe`, `bitsadmin.exe`, `taskkill`, `sc.exe`, `netsh`, `bcdedit`, `vssadmin` | **Not found** |
| Registry access | `My.Computer.Registry`, `Microsoft.Win32.Registry`, startup/run key writes | **Not found** |
| Service installation | `ServiceController`, `sc.exe` | **Not found** |
| Scheduled tasks | `schtasks` | **Not found** |
| File deletion / encryption | `File.Delete`, `Kill` (as file command), `WriteAllBytes` to system paths | **Not found** |
| Network / download | `HttpClient`, `WebClient`, `DownloadFile`, `WebBrowser`, `WebView2` | **Not found** |
| Assembly injection | `Assembly.Load`, `Assembly.LoadFrom`, `AppDomain.AssemblyResolve` | **Not found** |
| Credential access | `WindowsIdentity`, `WindowsPrincipal`, credential/token APIs | **Not found** |
| Anti-analysis / evasion | VM detection, debugger detection, sleep loops, obfuscation | **Not found** |
| Startup persistence | `Environment.SpecialFolder.Startup`, run keys, startup folder | **Not found** |
| Screen/keyboard capture | Keyboard hooks, screen capture APIs | **Not found** |
| Payload dropping | `WriteAllBytes` to temp/exe paths, `GetTempPath` | **Not found** |
| Self-extraction | Embedded resource extraction to EXE/DLL | **Not found** |

### File write activity found

One legitimate file write exists: `ExportBomClicked` in `MainForm.vb` uses a `SaveFileDialog` +
`StreamWriter` to export a BOM to a user-chosen CSV file path. This is fully user-initiated,
user-targeted, and produces only structured CSV data. No system paths, temp directories, or
auto-write behavior is involved.

### What was found that requires mention

- **Dynamic COM activation** (`Activator.CreateInstance` + `Type.GetTypeFromProgID`) — present
  and fully explained as Creo API integration; see SH-01.
- **`CallByName` dispatch** — present and fully explained; see SH-02.
- **Macro execution via `RunMacro`** — present and fully explained as Creo-internal only; see SH-03.
- **Non-public reflection** — present and fully explained as WinForms workaround; see SH-04.

### Final statement

> **No clear malicious or system-harming behavior found in static review.**

All flagged patterns are directly attributable to the PTC Creo COM API programming model and
standard .NET WinForms development practices. There is no evidence of destructive, stealthy,
evasive, persistent, or exfiltration-related logic in any of the four source files.

---

## Section H: Likely False-Positive Indicators

| Indicator | Explanation |
|---|---|
| **Unsigned executable** | EXE is not Authenticode-signed. Low-trust EXEs from non-enterprise sources are flagged by default in many EDR products. |
| **Low reputation** | Internal-only distribution with a small user population means no cloud reputation data exists for this hash. |
| **ZIP distribution** | Delivered as a ZIP archive (`dist/CreoRegenAppVB-Update.zip`) with no installer, no trust chain, no package manager metadata. |
| **Dynamic COM activation** | `Activator.CreateInstance` + `Type.GetTypeFromProgID` pattern resembles COM hijacking techniques. |
| **Late-bound dispatch** | `CallByName` resembles obfuscated reflection used by loaders and injectors. |
| **Macro execution pattern** | `RunMacro` call resembles Office VBA macro abuse sequences. |
| **Non-public reflection** | `BindingFlags.NonPublic` resembles memory manipulation used by injectors. |
| **STA worker thread** | A dedicated STA thread with a `BlockingCollection(Of Action)` message loop resembles patterns used by certain threading-based evasion techniques (though here it is a legitimate COM threading requirement). |

---

## Section I: IT Handoff Summary

### What the app does

`CreoRegenAppVB` is an internal Windows desktop tool for PTC Creo CAD automation. Engineers use
it to connect to a running Creo Parametric session and perform operations (model retrieval,
regeneration, parameter editing, BOM export) that would otherwise require manual UI interaction
inside Creo. It writes one type of file: a user-initiated CSV export of BOM data saved to a
user-chosen path. It has no network capability, no registry access, no persistence, and no
privilege-elevation code.

### What it depends on

- **.NET Framework 4.8** — Microsoft system runtime (pre-installed on all modern Windows systems)
- **PTC Creo VB API COM server** — Installed by the Creo Parametric installer and registered via
  `vb_api_register.bat`. Version is tied to the installed Creo release. No vulnerability matching
  is possible from source alone.
- No NuGet or third-party packages.

### SonarQube-style issues found

- `Option Strict Off` (project-wide) — increases latent type-safety risk
- Multiple bare `Catch` blocks silently discard exception details — makes debugging harder
  (comments added in this PR to explain each intentional swallow)
- Magic model-type integers replaced with named constants in this PR
- Large file sizes and duplicated COM probe patterns (maintainability debt)

### Dependency-Check-style issues found

- Zero NuGet/package-manager dependencies — no standard CVE matching applicable
- One third-party vendor runtime (PTC Creo VB API) — version unknown from source; IT should
  verify Creo is current on its PTC maintenance schedule
- No Authenticode signature on the EXE — primary distribution trust issue

### Was any malicious code found?

**No.** All patterns that resemble malicious behavior (COM activation, late-bound dispatch, macro
execution, non-public reflection) are directly required by the Creo API programming model and
standard WinForms development practices. No network, persistence, credential-access, destructive,
or evasive behavior was found.

### Why was the EXE likely flagged?

The combination of: unsigned EXE + low reputation + ZIP delivery + dynamic COM activation +
`CallByName` dispatch + `RunMacro` call creates a cluster of behavioral signals that EDR products
associate with living-off-the-land (LOLbin) attack chains. None of these signals indicate actual
malicious behavior in this case, but together they push the application into the "suspicious"
detection tier.

### What should be fixed next

1. **Code-sign the EXE** with a corporate Authenticode certificate — highest-impact action.
2. **Whitelist the application** in SentinelOne for the engineering workstation group.
3. **Migrate to early-bound COM interop** (add Creo PIA/TLB reference) to eliminate dynamic COM
   activation and `CallByName` signals — requires deciding on a minimum supported Creo version.
4. **Enable `Option Strict On`** and address resulting compile errors.
5. **Log all swallowed exceptions** to the UI log at DEBUG level to improve observability.
6. **Replace the macro fallback** (`RunRegenerationMacros`) with a version-correct direct API call.

---

## Section J: Final Shortlists

### 1. Plain-English dependency list

- .NET Framework 4.8 (Microsoft, pre-installed on Windows)
- System / System.Core / System.Drawing / System.Windows.Forms / System.Xml (Microsoft BCL, .NET 4.8)
- Microsoft.VisualBasic runtime (Microsoft, .NET 4.8 — provides `CallByName`)
- PTC Creo VB API COM server (PTC, version = installed Creo release, registered by Creo installer)
- pro_comm_msg.exe (PTC, Creo IPC helper, version = installed Creo release)

### 2. All third-party dependencies

| Name | Vendor | Version | How obtained |
|---|---|---|---|
| Creo VB API COM server | PTC | Unknown (tied to Creo release) | Creo installer + `vb_api_register.bat` |
| pro_comm_msg.exe | PTC | Unknown (tied to Creo release) | Creo installer |

### 3. Most likely reasons the EXE was flagged

1. Unsigned executable (no Authenticode certificate)
2. Low reputation (internal tool, small user population, no cloud reputation data)
3. Dynamic COM activation (`Activator.CreateInstance` + `Type.GetTypeFromProgID`)
4. Late-bound dispatch via `CallByName` (resembles reflection-based obfuscation)
5. `RunMacro` call pattern (resembles Office macro abuse)
6. ZIP-based distribution (no installer trust chain)

### 4. Highest-priority fixes

| Priority | Fix | Impact |
|---|---|---|
| 1 | Authenticode code-sign the EXE | Eliminates low-reputation signal; major EDR trust improvement |
| 2 | Whitelist in SentinelOne for engineering workstations | Immediate operational relief |
| 3 | Migrate to early-bound COM interop (add Creo TLB/PIA reference) | Eliminates SH-01, SH-02, EDR-01, EDR-02 |
| 4 | Enable `Option Strict On` | Improves type safety, reduces latent runtime error risk |
| 5 | Log swallowed exceptions to the UI log | Improves observability for support and incident response |

### 5. One-sentence final conclusion

> **No malicious or system-harming behavior was found in this codebase; all EDR-flagged patterns
> are legitimate consequences of the PTC Creo COM API programming model, and the highest-impact
> remediation is Authenticode code-signing combined with EDR policy whitelisting for the
> engineering workstation group.**

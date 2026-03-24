# Security Compliance Report

**Repository:** `sribalaji315-cloud/creovbtool`
**Report Date:** 2026-03-24
**Reviewed Commit:** `6befa342d38b33c22ddca8a29fa4def120b1f41b` (chore: update latest workspace changes)
**Requested by:** IT Security (Sentinel One alert review)

---

## Executive Summary

The security fixes requested by IT have been implemented and are **compliant**. The primary change that triggered the Sentinel One alert in the previous build—macro string injection via `RunMacro` / `mapkey` commands—has been removed and replaced with native Creo COM API calls. All five security controls verified below are in place.

---

## Findings by Control Area

### 1. Macro / Shell Command Injection — FIXED ✅

**Previous risk:**
`CreoSessionService.vb` previously contained a `RunRegenerationMacros` method that executed Creo macro strings using `InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command ...")`. This pattern is indistinguishable from macro-based command injection and correctly triggered the Sentinel One alert.

**Fix applied:**
- `RunRegenerationMacros` has been **completely removed**.
- All three call-sites (in `RegenerateModelAsync`, `SetParameterAsync`, and `SetParametersAsync`) now call `InvokeNativeRegeneration`, which uses the typed Creo COM API:
  - `IpfcModel2D.Regenerate()` for drawings/layouts (no arguments).
  - `IpfcSolid.Regenerate(IpfcRegenInstructions)` invoked via `Type.InvokeMember` to correctly pass the single optional argument through the COM dispatch layer.
- No `RunMacro`, `mapkey`, or shell-style string commands exist anywhere in the codebase.

---

### 2. COM Class Allowlist (ProgID Injection Prevention) — FIXED ✅

**Previous risk:**
`CreatePfcObject` accepted any class name and passed it to `Activator.CreateInstance(Type.GetTypeFromProgID(...))` without restriction, allowing instantiation of arbitrary registered COM classes.

**Fix applied:**
`CreatePfcObject` now validates the requested class name against a hardcoded allowlist of known Creo API prefixes before any `Activator.CreateInstance` call:

```vb
Private Shared ReadOnly AllowedPfcPrefixes As String() = {
    "pfc", "Cpfc", "CMpfc", "CCpfc", "Mpfc"
}

Private Function CreatePfcObject(className As String) As Object
    If String.IsNullOrWhiteSpace(className) OrElse
       Not AllowedPfcPrefixes.Any(Function(p) className.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
        Throw New ArgumentException("Blocked ProgID: '" & className & "' is not a recognised Creo API class name.")
    End If
    ...
```

All call-sites use hardcoded string literals (`"pfcAsyncConnection"`, `"pfcModelDescriptor"`, `"CpfcRegenInstructions"`, etc.) — user-controlled input never flows into this function.

---

### 3. Application Manifest — Privilege Escalation Prevention ✅

**`app.manifest`** was added to the project with an explicit execution level declaration:

```xml
<requestedExecutionLevel level="asInvoker" uiAccess="false" />
```

- `asInvoker`: the application runs at the privilege level of the launching user; no elevation is requested or permitted.
- `uiAccess="false"`: the application cannot bypass UIPI (User Interface Privilege Isolation) to interact with higher-privilege windows.

This manifest is correctly wired into the build via `<ApplicationManifest>app.manifest</ApplicationManifest>` in the project file.

---

### 4. Assembly Signing — Binary Integrity ✅

The project file now includes:

```xml
<SignAssembly>true</SignAssembly>
<AssemblyOriginatorKeyFile>CreoRegenAppVB.snk</AssemblyOriginatorKeyFile>
```

The `CreoRegenAppVB.snk` strong-name key file is present. This cryptographically signs the output assembly, preventing substitution or tampering without key access.

---

### 5. COM Visibility — Unexposed Public API ✅

`My Project\AssemblyInfo.vb` includes:

```vb
<Assembly: ComVisible(False)>
```

This ensures the assembly's public types are not visible to COM clients by default, reducing the attack surface exposed to other COM-aware processes on the same machine.

---

## Additional Security Controls Verified (Pre-existing)

| Control | Status | Notes |
|---|---|---|
| Target framework | ✅ `.NET Framework 4.8` | Latest v4.x, includes all security patches |
| No shell / process execution | ✅ | No `Process.Start`, `Shell()`, `Cmd.exe` anywhere |
| No arbitrary file access | ✅ | CSV export uses `SaveFileDialog` (user chooses path) |
| CSV value escaping | ✅ | `EscapeCsvValues` quotes and double-escapes all values |
| BOM traversal depth cap | ✅ | `depth >= 32` guard prevents stack overflow |
| Input validation | ✅ | `ModelNameParser.Parse` rejects null/empty/whitespace |
| STA COM thread serialisation | ✅ | Single `BlockingCollection` queue prevents concurrent COM |
| Deterministic build | ✅ | `<Deterministic>true</Deterministic>` in project file |
| Option Explicit | ✅ | `<OptionExplicit>On</OptionExplicit>` — no implicit vars |
| Warning level | ✅ | `<WarningLevel>4</WarningLevel>` — maximum |
| IDisposable / resource cleanup | ✅ | `Dispose()` joins worker thread, closes COM connection |

---

## Why the Previous Build Triggered Sentinel One

The previous build contained the `RunRegenerationMacros` method which called:

```vb
InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdRegenAuto`; ...")
InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdDwgRegenModel`;...")
```

Sentinel One's behavioural detection classified the `RunMacro` + `mapkey(continued) ~ Command` pattern as potential macro-based command injection — a known attack technique. The fix removes this pattern entirely and replaces it with direct, typed COM API method calls that have no ambiguous macro-execution semantics.

---

## Conclusion

**All IT-advised security fixes are covered.** The codebase at commit `6befa34` is compliant with the security requirements. No further code changes are required to address the Sentinel One complaint or the IT security advisory.

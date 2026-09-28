# BOMForge

BOMForge is a standalone WinForms application for Creo Parametric, covering:

- Connect/start Creo session (VB API async COM)
- Retrieve model by name
- Regenerate model by name
- BOM hierarchy browsing, parameter editing and CSV export
- Desktop log UI

The assembly is still named `CreoRegenAppVB` for build continuity; BOMForge is the product name.

## Project Location

- `CreoRegenAppVB/CreoRegenAppVB.vbproj`

## Runtime Prerequisites

1. Install Creo with VB API support.
2. Register VB API COM server:
   - Run `<creo_loadpoint>\Parametric\bin\vb_api_register.bat`
3. Set environment variable:
   - `PRO_COMM_MSG_EXE=<creo_loadpoint>\<datecode>\Common Files\x86e_win64\obj\pro_comm_msg.exe`
4. Optional:
   - `PFCLS_START_DIR=<working_directory>`

## Build in Visual Studio

1. Open `CreoRegenAppVB.vbproj` in Visual Studio.
2. Build `x64` configuration.
3. Run the app.

## Usage

1. Click `Connect`.
   - The app tries to connect to active Creo first.
   - If no active session is found, it starts a new session using `parametric`.
2. Enter model name:
   - `name`
   - `name.prt`
   - `name.asm`
   - `name.drw`
3. Click `Retrieve` or `Regenerate`.

## Notes

- API calls are serialized through a dedicated STA worker thread.
- This first implementation uses late-bound COM (`CreateObject`/`CallByName`) so the codebase can be started without hard-coded TLB paths.
- Ribbon/menu callback integration is intentionally excluded from this first standalone pass.

## Licensing

BOMForge runs as a **30-day trial**, then requires a license key.

### Trial

- Starts on first launch and is tracked per Windows user, so reinstalling does not reset it.
- State is mirrored to `HKCU\Software\BOMForge\Settings` and `%APPDATA%\BOMForge\state.dat`; both are HMAC-protected and removing one mirror does not reset the trial.
- Moving the system clock backwards, or hand-editing either mirror, ends the trial immediately.
- The remaining days are shown in the title bar. When the trial ends, the app shows the activation dialog and exits unless a valid key is entered.

### Activating

Click **Activate** in the main window (or use the dialog shown at launch), then paste the key or load a `.lic` file. The key is stored at `%APPDATA%\BOMForge\license.lic` and re-verified on every launch. Keys are perpetual and are not tied to a machine.

### Issuing keys

Keys are signed offline with ECDSA P-256. Only the public key is embedded in the application, so keys cannot be forged from the shipped binary.

1. Run `tools\BOMForgeKeyGen.exe`.
2. On first run it offers to generate a key pair. It is saved DPAPI-encrypted to `%APPDATA%\BOMForge.KeyGen\private.blob` — **back this file up**. If it is lost, no further keys can be issued for the current build; if it is replaced, every key already issued stops working.
3. Copy the displayed public key into `LicenseFormat.PublicKeyBase64` and rebuild the application. This is only needed once.
4. Enter a licensee name and click **Generate Key**. Every issued key is appended to `%APPDATA%\BOMForge.KeyGen\issued-keys.csv`.

**Never ship `BOMForgeKeyGen.exe` or `private.blob` with the application.**

> The signing scheme prevents key forgery, but .NET IL is readable and the trial check itself is not obfuscated. This is a deterrent against casual copying, not protection against a determined attacker.

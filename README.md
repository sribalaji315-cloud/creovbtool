# CreoRegenApp VB.NET (Implementation Start)

This project is a standalone WinForms rewrite of the Java toolkit flow, covering:

- Connect/start Creo session (VB API async COM)
- Retrieve model by name
- Regenerate model by name
- Desktop log UI

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

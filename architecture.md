# Creo Regen Toolkit Architecture

## 1. Purpose

Creo Regen Toolkit is a Windows desktop utility that helps engineering users connect to Creo Parametric, retrieve models, regenerate models, inspect model metadata, edit parameters, and work with assembly BOM hierarchies from one interface.

The tool is designed to reduce repetitive Creo UI actions and provide a safer, structured workflow for common model operations.

## 2. Core Function

At a high level, the application:

1. Accepts a model name input from the user.
2. Connects to an existing or newly started Creo session.
3. Resolves and retrieves the target model.
4. Executes a chosen operation (regenerate, read parameters/features/components, build BOM, set parameters).
5. Displays results in dedicated data grids and logs status in the UI.

The UI remains responsive by executing Creo COM calls asynchronously through a dedicated STA worker thread in the session service.

## 3. High-Level Architecture

The solution follows a simple layered design:

- Presentation Layer
  - WinForms UI handles user interaction, validation, grids, and status logging.
- Application/Orchestration Layer
  - Event handlers coordinate UI state, call service methods, and map results to grids.
- Integration Layer
  - Creo session service encapsulates COM connection, late-bound API invocation, and model operations.
- Utility Layer
  - Model name parser normalizes user input and extension handling.

## 4. Main Components

### 4.1 Program Entry

- Starts the application in STA mode.
- Initializes visual styles and launches the main form.

### 4.2 Main Form (UI + Workflow Control)

Responsibilities:

- Render controls (buttons, tabs, parameter and BOM grids, log window).
- Manage busy-state and button enablement.
- Trigger asynchronous service operations.
- Display operation outcomes and errors.
- Handle BOM hierarchy interactions (expand/collapse/export).
- Handle item-level parameter editing from BOM selection.

### 4.3 Creo Session Service (Business + Integration)

Responsibilities:

- Create/connect/disconnect Creo sessions.
- Retrieve and regenerate models.
- Read and write model parameters.
- Read features and component-like feature references.
- Build hierarchical BOM and flatten it for UI consumption.
- Shield UI from COM details with typed result objects.

### 4.4 Model Name Parser

Responsibilities:

- Validate model input text.
- Strip version suffixes where present.
- Detect explicit extensions and map model type hints.
- Produce normalized file name candidates.

## 5. Concurrency and Threading Model

- The UI runs on the WinForms main thread.
- The service owns a single background STA worker thread and a blocking queue.
- All Creo COM operations are serialized through that queue.

Benefits:

- Avoids concurrent COM access hazards.
- Preserves deterministic operation ordering.
- Keeps UI responsive while operations execute.

## 6. Data and Result Model

The service returns structured result types to keep UI logic simple and explicit.

Key result/data objects:

- OperationResult
  - Generic success/failure + message.
- ParameterQueryResult, ParameterItem, ParameterUpdateItem
  - Parameter read/write payloads.
- FeatureQueryResult, FeatureItem
  - Feature listing payloads.
- ComponentQueryResult, ComponentItem
  - Component feature listing payloads.
- BomQueryResult, BomNodeItem, BomGridRowItem
  - Hierarchical BOM and flattened grid rows with parent-child relationships.

## 7. Detailed Feature Breakdown

### 7.1 Session Connection

Purpose:

- Attach to an active Creo session or start a new one if needed.

Behavior:

- Attempts Connect first.
- Falls back to Start when no active session is available.
- Stores active session handles for later operations.

User Value:

- One-click session bootstrap.
- Minimal setup friction when Creo is not already running.

### 7.2 Model Retrieval

Purpose:

- Retrieve and activate a model in Creo by name.

Behavior:

- Parses input into normalized candidates.
- Tries multiple extension strategies when extension is not explicit.
- Displays and activates model window on success.

User Value:

- Faster model access with tolerant name handling.

### 7.3 Model Regeneration

Purpose:

- Execute regeneration on the target model.

Behavior:

- Retrieves target model using same resolution path.
- Invokes model-type-aware regeneration paths.
- Activates/repaints model window after regeneration.

User Value:

- Quick regeneration from external UI without manual Creo navigation.

### 7.4 Load Parameters (Model Scope)

Purpose:

- Read model parameters into an editable grid.

Behavior:

- Reads parameter collection from model.
- Converts native parameter value types into display text.
- Displays current value, type, description, and editable new value column.

User Value:

- Centralized parameter visibility and bulk edit preparation.

### 7.5 Set Parameters (Model Scope)

Purpose:

- Apply one or more parameter updates to the active model.

Behavior:

- Collects rows containing user-provided new values.
- Resolves parameter by name and converts text input to typed value.
- Applies all updates and performs best-effort regeneration.
- Reloads parameter grid after update.

User Value:

- Multi-parameter updates in one action.

### 7.6 Load Features

Purpose:

- Show feature-level metadata for the model.

Behavior:

- Enumerates features and captures id, name, type, subtype, and status.

User Value:

- Quick model structure inspection and troubleshooting context.

### 7.7 Load Components

Purpose:

- List component-related feature entries from assemblies.

Behavior:

- Scans feature collection for model descriptors.
- Resolves component display names and file names.
- Returns list for tabular display.

User Value:

- Fast component awareness without manual tree traversal.

### 7.8 Build BOM Hierarchy

Purpose:

- Build and visualize a hierarchical BOM tree from an assembly.

Behavior:

- Validates assembly requirement.
- Recursively traverses child components.
- Captures level, feature ID, display name, file name, and key metadata such as DESCP and BS_CLASS.
- Includes placeholder Conditions and Inheritance columns for future user-defined data.
- Flattens hierarchy into parent-linked rows for grid rendering.

User Value:

- Assembly structure visibility with engineering metadata context.

### 7.9 BOM Expand/Collapse and Visibility Filtering

Purpose:

- Allow interactive navigation of large BOMs.

Behavior:

- Stores collapsed node ids.
- Re-renders visible subset based on ancestor collapse state.
- Supports expand-all and collapse-all operations.

User Value:

- Better usability for deep BOMs.

### 7.10 BOM Export to CSV

Purpose:

- Export current BOM grid rows to a CSV file.

Behavior:

- Includes Level, Feature ID, Name, File Name, Descp, BS_Class, Conditions, and Inheritance columns.
- Escapes CSV values safely.

User Value:

- Easy downstream sharing and analysis.

### 7.11 Item Parameter Panel (BOM Context)

Purpose:

- Show and edit parameters for the selected BOM item directly from hierarchy view.

Behavior:

- Right-side panel appears on BOM tab when user clicks "Show Params".
- Hidden by default when a new BOM is loaded.
- Panel header includes a close (X) button to dismiss the panel.
- Clicking "Show Params" reveals the panel and loads the selected item's parameters.
- Loads selected item parameters into a dedicated grid.
- Supports item-level Set Selected + Regenerate.

User Value:

- Contextual editing of child components without leaving BOM workflow.
- Panel does not consume screen space until explicitly requested.

### 7.12 Auto-Load Item Parameters on BOM Selection (Latest Enhancement)

Purpose:

- Remove the need to click Show Params for every BOM row selection.

Behavior:

- Selection change on BOM grid now triggers parameter load automatically.
- Suppresses selection-triggered loads during BOM re-render cycles.
- Prevents overlapping concurrent item-parameter loads.
- Skips redundant reload for same selected model unless user forces refresh.

User Value:

- Faster navigation and significantly fewer clicks during BOM review/edit sessions.

### 7.13 Unified Logging and Busy-State UX

Purpose:

- Keep users informed and prevent conflicting operations.

Behavior:

- Every operation logs started/success/error status.
- Busy state disables action buttons and relevant grids while operations execute.

User Value:

- Predictable workflow and clearer error diagnosis.

## 8. Input Resolution Strategy

The parser and service together support flexible model input:

- Accepts names with or without extension.
- Normalizes extension case and known model kinds.
- Tries common extension order when extension is omitted.
- Strips version suffixes before descriptor creation.

This increases retrieval success for real-world naming patterns.

## 9. COM Integration Strategy

Integration is intentionally late-bound to reduce deployment coupling.

Key points:

- Uses ProgID-based object creation with fallback variants.
- Uses guarded property/method invocation wrappers.
- Applies defensive fallbacks for collection access, descriptors, and metadata retrieval.
- Restricts created classes to expected Creo API prefix patterns.

This design improves compatibility across different Creo API registration environments.

## 10. Error Handling and Resilience

- Service methods return explicit fail results rather than throwing to UI for expected operational failures.
- Exceptions are converted into user-facing error messages with context.
- Many COM calls use fallback patterns to tolerate interface variance.
- Regeneration after parameter set is best-effort to preserve value updates even if regen warning/errors occur.

## 11. UI Layout Summary

- Top action area: model input + operation buttons.
- Main tab area:
  - Log
  - Parameters
  - BOM (columns: Toggle, Level, Feature ID, Name, File Name, Descp, BS_Class, Conditions, Inheritance)
- BOM side panel:
  - Close (X) button + Item parameter title + loading indicator
  - Item parameter grid
  - Set Selected + Regenerate action

## 12. Current Boundaries and Scope

In scope:

- Interactive desktop operations for model/session/metadata and parameter workflows.
- Hierarchical BOM exploration with item-level parameter editing.

Out of scope:

- Ribbon/menu callback integration.
- Multi-user or server-side orchestration.
- Advanced persistence/history of operation state.

## 13. Future Extension Opportunities

- Add dedicated Features and Components tabs if desired by UI flow.
- Add operation cancellation tokens for long-running BOM traversals.
- Add configurable column presets and saved views.
- Add richer export formats and optional BOM depth filters.
- Add diagnostics pane for COM registration and environment checks.

## 14. Feature-to-Function Call Map (Debug Guide)

This section lists the primary call chain for each feature so you can trace behavior quickly during debugging.

### 14.1 Session Connection

- UI Entry:
  - `MainForm.ConnectClickedAsync`
- UI Orchestration:
  - `MainForm.RunOperationAsync`
- Service Entry:
  - `CreoSessionService.EnsureConnectedAsync`
- Service Queue:
  - `CreoSessionService.Enqueue`
  - `CreoSessionService.ProcessQueue`
- Core Calls:
  - `CreoSessionService.CreatePfcObject("pfcAsyncConnection")`
  - `CreoSessionService.InvokeMethod(..., "Connect", ...)`
  - fallback `CreoSessionService.InvokeMethod(..., "Start", ...)`
  - `CreoSessionService.GetSessionFromConnection`

### 14.2 Model Retrieval

- UI Entry:
  - `MainForm.RetrieveClickedAsync`
- UI Orchestration:
  - `MainForm.RunOperationAsync`
- Service Entry:
  - `CreoSessionService.RetrieveModelAsync`
- Core Calls:
  - `ModelNameParser.Parse`
  - `CreoSessionService.TryRetrieveModel`
  - `CreoSessionService.BuildCandidateNames`
  - `CreoSessionService.CreateModelDescriptor`
  - `CreoSessionService.InvokeMethod(_session, "RetrieveModel", ...)`
  - `CreoSessionService.InvokeMethod(model, "Display")`
  - `CreoSessionService.InvokeMethod(_session, "GetModelWindow", ...)`
  - fallback `CreoSessionService.InvokeMethod(_session, "CreateModelWindow", ...)`

### 14.3 Model Regeneration

- UI Entry:
  - `MainForm.RegenerateClickedAsync`
- UI Orchestration:
  - `MainForm.RunOperationAsync`
- Service Entry:
  - `CreoSessionService.RegenerateModelAsync`
- Core Calls:
  - `ModelNameParser.Parse`
  - `CreoSessionService.TryRetrieveModel`
  - `CreoSessionService.InvokeNativeRegeneration`
  - `CreoSessionService.ActivateModelWindow`

### 14.4 Load Parameters (Model Scope)

- UI Entry:
  - `MainForm.LoadParametersClickedAsync`
- Service Entry:
  - `CreoSessionService.GetParametersAsync`
- Core Calls:
  - `ModelNameParser.Parse`
  - `CreoSessionService.TryRetrieveModel`
  - `CreoSessionService.GetParametersCollection`
  - `CreoSessionService.GetCollectionCount`
  - `CreoSessionService.GetCollectionItem`
  - `CreoSessionService.GetParamDiscr`
  - `CreoSessionService.GetParamValueText`
  - `CreoSessionService.TryGetParameterDescription`

### 14.5 Set Parameters (Model Scope)

- UI Entry:
  - `MainForm.SetSelectedParameterClickedAsync`
- UI Helpers:
  - `MainForm.CollectPendingParameterUpdates`
  - `MainForm.ClearPendingParameterUpdates`
  - `MainForm.ReloadParametersGridAsync`
- Service Entry:
  - `CreoSessionService.SetParametersAsync`
- Core Calls:
  - `ModelNameParser.Parse`
  - `CreoSessionService.TryRetrieveModel`
  - `CreoSessionService.FindParameterByName`
  - `CreoSessionService.GetParamDiscr`
  - `CreoSessionService.CreateParamValueFromInput`
  - `CreoSessionService.SetProperty`
  - `CreoSessionService.InvokeNativeRegeneration`

### 14.6 Load Features

- UI Entry:
  - `MainForm.LoadFeaturesClickedAsync`
- Service Entry:
  - `CreoSessionService.GetFeaturesAsync`
- Core Calls:
  - `ModelNameParser.Parse`
  - `CreoSessionService.TryRetrieveModel`
  - `CreoSessionService.GetFeaturesCollection`
  - `CreoSessionService.GetCollectionCount`
  - `CreoSessionService.GetCollectionItem`
  - `CreoSessionService.TryGetString`

### 14.7 Load Components

- UI Entry:
  - `MainForm.LoadComponentsClickedAsync`
- Service Entry:
  - `CreoSessionService.GetComponentsAsync`
- Core Calls:
  - `ModelNameParser.Parse`
  - `CreoSessionService.TryRetrieveModel`
  - `CreoSessionService.GetFeaturesCollection`
  - `CreoSessionService.GetCollectionCount`
  - `CreoSessionService.GetCollectionItem`
  - `CreoSessionService.GetFeatureModelDescriptor`
  - `CreoSessionService.ResolveDescriptorFileName`
  - `CreoSessionService.ResolveDescriptorDisplayName`
  - `CreoSessionService.GetFeatureId`

### 14.8 Build BOM Hierarchy

- UI Entry:
  - `MainForm.BuildBomClickedAsync`
- Service Entry:
  - `CreoSessionService.GetBomHierarchyAsync`
- BOM Build Calls:
  - `ModelNameParser.Parse`
  - `CreoSessionService.TryRetrieveModel`
  - `CreoSessionService.IsAssemblyModel`
  - `CreoSessionService.BuildBomNode` (recursive)
  - `CreoSessionService.FlattenBomNode`
  - `CreoSessionService.CountBomNodes`
- UI Render Calls:
  - `MainForm.RenderBomGrid`
  - `MainForm.IsRowVisible`
  - `MainForm.UpdateShowParametersButtonState`

### 14.9 BOM Expand/Collapse and Visibility Filtering

- UI Entry:
  - `MainForm.BomGridCellContentClick` (single-node toggle)
  - `MainForm.ExpandAllBomClicked`
  - `MainForm.CollapseAllBomClicked`
- Render Path:
  - `MainForm.RenderBomGrid`
  - `MainForm.IsRowVisible`

### 14.10 BOM Export to CSV

- UI Entry:
  - `MainForm.ExportBomClicked`
- Core Calls:
  - `MainForm.EscapeCsvValues`
  - `System.IO.StreamWriter.WriteLine`

### 14.11 Item Parameter Panel (BOM Context)

- UI Entry:
  - `MainForm.ShowItemParametersClickedAsync` (shows panel + loads params)
  - `MainForm.CloseItemParametersPanel` (hides panel via X button)
  - `MainForm.SetSelectedItemParameterClickedAsync`
- Load Calls:
  - `MainForm.LoadItemParametersForSelectedRowAsync(forceReload:=True)`
  - `CreoSessionService.GetParametersAsync`
  - `MainForm.BeginItemParametersLoading`
  - `MainForm.EndItemParametersLoading`
- Set Calls:
  - `MainForm.CollectPendingParameterUpdates`
  - `CreoSessionService.SetParametersAsync`
  - `MainForm.ReloadParametersGridAsync`

### 14.12 Auto-Load Item Parameters on BOM Selection

- UI Event Path:
  - `MainForm.BomGridSelectionChanged`
  - guard: `_suppressSelectionChanged`
  - guard: `_isLoadingItemParams`
  - `MainForm.LoadItemParametersForSelectedRowAsync(forceReload:=False)`
- Service Call:
  - `CreoSessionService.GetParametersAsync`
- Related UI State Calls:
  - `MainForm.UpdateShowParametersButtonState`
  - `MainForm.SetBusyState`

### 14.13 Unified Logging and Busy-State UX

- Logging Calls:
  - `MainForm.AddLog`
- Busy-State Calls:
  - `MainForm.SetBusyState`
- Common Operation Wrapper:
  - `MainForm.RunOperationAsync`

### 14.14 Shared Service Infrastructure (applies to all service features)

- Queue and thread:
  - `CreoSessionService.Enqueue`
  - `CreoSessionService.EnqueueParameters`
  - `CreoSessionService.EnqueueFeatures`
  - `CreoSessionService.EnqueueComponents`
  - `CreoSessionService.EnqueueBom`
  - `CreoSessionService.ProcessQueue`
- Late-bound COM helpers:
  - `CreoSessionService.CreatePfcObject`
  - `CreoSessionService.InvokeMethod`
  - `CreoSessionService.InvokeMethodSafe`
  - `CreoSessionService.GetProperty`
  - `CreoSessionService.SetProperty`
  - `CreoSessionService.TryGetString`

## 15. First Breakpoint Guide (Per Feature)

Use this as a fast debugging playbook. Start at the first breakpoint, then step into the second breakpoint to verify service execution.

### 15.1 Session Connection

- First breakpoint:
  - `MainForm.ConnectClickedAsync`
- Next breakpoint:
  - `CreoSessionService.EnsureConnectedAsync`

### 15.2 Model Retrieval

- First breakpoint:
  - `MainForm.RetrieveClickedAsync`
- Next breakpoint:
  - `CreoSessionService.RetrieveModelAsync`
- If model is not found:
  - `ModelNameParser.Parse`
  - `CreoSessionService.BuildCandidateNames`
  - `CreoSessionService.TryRetrieveModel`

### 15.3 Model Regeneration

- First breakpoint:
  - `MainForm.RegenerateClickedAsync`
- Next breakpoint:
  - `CreoSessionService.RegenerateModelAsync`
- Regen internals:
  - `CreoSessionService.InvokeNativeRegeneration`

### 15.4 Load Parameters (Model Scope)

- First breakpoint:
  - `MainForm.LoadParametersClickedAsync`
- Next breakpoint:
  - `CreoSessionService.GetParametersAsync`
- If parameter list is empty:
  - `CreoSessionService.GetParametersCollection`
  - `CreoSessionService.GetCollectionCount`
  - `CreoSessionService.GetCollectionItem`

### 15.5 Set Parameters (Model Scope)

- First breakpoint:
  - `MainForm.SetSelectedParameterClickedAsync`
- Next breakpoint:
  - `CreoSessionService.SetParametersAsync`
- Type conversion failures:
  - `CreoSessionService.GetParamDiscr`
  - `CreoSessionService.CreateParamValueFromInput`
  - `CreoSessionService.SetProperty`

### 15.6 Load Features

- First breakpoint:
  - `MainForm.LoadFeaturesClickedAsync`
- Next breakpoint:
  - `CreoSessionService.GetFeaturesAsync`
- Collection issues:
  - `CreoSessionService.GetFeaturesCollection`
  - `CreoSessionService.GetCollectionCount`

### 15.7 Load Components

- First breakpoint:
  - `MainForm.LoadComponentsClickedAsync`
- Next breakpoint:
  - `CreoSessionService.GetComponentsAsync`
- Descriptor resolution issues:
  - `CreoSessionService.GetFeatureModelDescriptor`
  - `CreoSessionService.ResolveDescriptorDisplayName`
  - `CreoSessionService.ResolveDescriptorFileName`

### 15.8 Build BOM Hierarchy

- First breakpoint:
  - `MainForm.BuildBomClickedAsync`
- Next breakpoint:
  - `CreoSessionService.GetBomHierarchyAsync`
- BOM shape/levels issues:
  - `CreoSessionService.BuildBomNode`
  - `CreoSessionService.FlattenBomNode`
  - `MainForm.RenderBomGrid`

### 15.9 BOM Expand/Collapse

- First breakpoint:
  - `MainForm.BomGridCellContentClick`
- Next breakpoint:
  - `MainForm.RenderBomGrid`
- Visibility issues:
  - `MainForm.IsRowVisible`

### 15.10 BOM Export CSV

- First breakpoint:
  - `MainForm.ExportBomClicked`
- Next breakpoint:
  - `MainForm.EscapeCsvValues`

### 15.11 Item Parameter Panel (Manual Refresh)

- First breakpoint:
  - `MainForm.ShowItemParametersClickedAsync`
- Next breakpoint:
  - `MainForm.LoadItemParametersForSelectedRowAsync`
- Service read call:
  - `CreoSessionService.GetParametersAsync`

### 15.12 Auto-Load Item Parameters on BOM Selection

- First breakpoint:
  - `MainForm.BomGridSelectionChanged`
- Next breakpoint:
  - `MainForm.LoadItemParametersForSelectedRowAsync`
- Guard-related debugging:
  - `_suppressSelectionChanged` transitions around `MainForm.RenderBomGrid`
  - `_isLoadingItemParams` transitions inside `MainForm.LoadItemParametersForSelectedRowAsync`

### 15.13 Set Item Parameters from BOM Panel

- First breakpoint:
  - `MainForm.SetSelectedItemParameterClickedAsync`
- Next breakpoint:
  - `CreoSessionService.SetParametersAsync`
- Post-write refresh:
  - `MainForm.ReloadParametersGridAsync`

### 15.14 Logging and Busy State

- First breakpoint:
  - `MainForm.RunOperationAsync`
- Next breakpoint:
  - `MainForm.SetBusyState`
- For operation message flow:
  - `MainForm.AddLog`

### 15.15 COM Invocation Failures (Cross-Cutting)

- First breakpoint:
  - `CreoSessionService.CreatePfcObject`
- Next breakpoint:
  - `CreoSessionService.InvokeMethod`
- If property/method mismatch is suspected:
  - `CreoSessionService.GetProperty`
  - `CreoSessionService.TryGetString`
  - `CreoSessionService.InvokeMethodSafe`

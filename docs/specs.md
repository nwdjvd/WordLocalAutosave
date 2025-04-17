# Word Local Autosave – Tech Spec

## Objective

Create a standalone Windows app that autosaves Microsoft Word documents locally when they are modified, with behavior identical to OneDrive’s autosave.

## Core Logic

- Connect to Word using `win32com.client.DispatchWithEvents`
- Hook into:
  - `Application.OnDocumentChange`
  - `Application.OnWindowSelectionChange`
- For every event, check if `doc.Saved == False`
- If so, call `doc.Save()` with a debounce (default: 10s)

## Edge Case Handling

| Case | Solution |
|------|----------|
| Multiple documents open | Loop through all in `wordApp.Documents` |
| Word crashes | Wrap COM access in `try/except`, reconnect |
| No events fire (rare) | Run a backup polling loop every 15s |
| User switches away from Word | Still track events, check active doc |
| Save fails | Log error to console / file |

## Backup Polling (Failsafe)

- Every 15s, loop through all open docs
- If any doc is dirty (`.Saved == False`) and last save was > debounce, call `.Save()`

## Optional Features (Future)

- Tray icon with toggle and status
- Log file with timestamped saves
- Settings file for debounce interval
- Versioning (autosave to `.bak` copies)
- GUI using PyQt or Tkinter

## Limitations

- Requires desktop Word (not web version)
- COM interface isn’t perfect — limited change events
- Debounce avoids excessive disk writes, but means instant save after every keystroke isn't guaranteed

## Reliability Strategy

- Dual event + polling system
- Handle all exceptions silently
- Ensure app doesn’t crash even if Word closes
s
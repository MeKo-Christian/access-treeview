# ADR 0001: Use C# (.NET Framework 4.8)

## Status: Accepted

## Context

We need to build two COM components for 64-bit MS Access:
- TreeEngine64: a COM in-proc DLL for tree data/logic
- TreeViewHost64: a visual ActiveX control (WinForms UserControl exposed via COM)

Both must support VBA `WithEvents` and be registerable via `regasm`.

Three languages were evaluated:

| Criterion | C# (.NET) | C++ (ATL) | Go |
|---|---|---|---|
| COM server (in-proc DLL) | Built-in `[ComVisible]`, `regasm` | Native ATL, full control | No native COM server support — not viable |
| ActiveX visual control | WinForms UserControl + COM exposure | MFC/ATL control hosting — complex | Not feasible |
| VBA `WithEvents` | Source interface via `[InterfaceType(InterfaceIsIDispatch)]` | Connection points via ATL — boilerplate-heavy | Not possible |
| Dev on Linux | Editing works, build/register/test needs Windows | Same | Same |
| Speed to build | Fastest — one ecosystem | Slowest — manual COM plumbing | N/A |

## Decision

Use **C# with .NET Framework 4.8** for both TreeEngine64 and TreeViewHost64.

- Go is not viable for COM in-proc servers or ActiveX controls
- C++ ATL works but doubles development time with manual COM plumbing
- C# provides COM visibility with attributes, WinForms UserControl hosting, and a single unified solution
- .NET Framework 4.8 (not .NET 8+) because Access COM interop requires in-proc DLLs registered via `regasm`, which is best supported on .NET Framework

## Consequences

- Must build and register on Windows (`regasm` + Access testing)
- Editing can happen on any OS via VS Code
- Single solution, shared types between engine and control
- Dependent on .NET Framework 4.8 runtime (pre-installed on modern Windows)

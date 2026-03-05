# Installer

## Prerequisites

- Windows
- [WiX Toolset v4+](https://wixtoolset.org/) (`dotnet tool install --global wix`)
- Both projects built in Release mode

## Build

```bash
# Build the solution first
dotnet build MeKoTreeView.sln -c Release

# Build the MSI
wix build installer/Product.wxs -o installer/MeKoTreeView.msi
```

## What It Does

### Install
1. Copies TreeEngine64.dll and TreeViewHost64.dll to Program Files
2. Runs `regasm /codebase` on TreeEngine64.dll (COM registration)
3. Runs `regasm /codebase /tlb` on TreeViewHost64.dll (COM + type library registration)

### Uninstall
1. Runs `regasm /unregister` on both DLLs
2. Removes files from Program Files

## Testing

After install, verify in Access VBA Immediate window:
```vba
? TypeName(CreateObject("MeKo.TreeEngine"))
' Should print: TreeEngine
```

## Silent Install

```cmd
msiexec /i MeKoTreeView.msi /qn
```

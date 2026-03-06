# AccessTreeView

solution := "AccessTreeView.slnx"
config := env("CONFIG", "Debug")

# List available recipes
default:
    @just --list

# Build the solution
build:
    dotnet build {{ solution }} -c {{ config }}

# Build in Release mode
release:
    dotnet build {{ solution }} -c Release

# Run all tests
test:
    dotnet test {{ solution }} -c {{ config }}

# Run tests with verbose output
test-verbose:
    dotnet test {{ solution }} -c {{ config }} --logger "console;verbosity=detailed"

# Clean build artifacts
clean:
    dotnet clean {{ solution }} -c {{ config }}

# Restore NuGet packages
restore:
    dotnet restore {{ solution }}

# Format all files (C# via dotnet format, markdown/json/yaml via prettier)
fmt:
    dotnet format {{ solution }}
    treefmt

# Check formatting without making changes
fmt-check:
    dotnet format {{ solution }} --verify-no-changes
    treefmt --fail-on-change

# Build the MSI installer (requires WiX Toolset)
installer: release
    wix build installer/Product.wxs -o out/MeKoTreeView.msi

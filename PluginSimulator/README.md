# D365 Plugin Simulator

A .NET Framework 4.7.1 tool that runs Dataverse / Dynamics 365 plugin DLLs **locally** against a real D365 record — **reads** are served from the live environment, **writes** are intercepted, captured, and reported instead of being applied.

Useful for:
- Debugging plugin logic against production data without touching production state
- Replaying historical changes from audit history to reproduce bugs
- Step-through debugging with Visual Studio attached (`--debug`)

## Quick Start

### GUI mode
```
PluginSimulator.exe
```
Fill in env URL, entity, record ID, and plugin DLL path, then click **Run Simulation**.

### CLI mode
```
PluginSimulator.exe ^
  --env https://yourorg.crm.dynamics.com ^
  --entity mash_servicetask ^
  --id 12345678-1234-1234-1234-123456789012 ^
  --assembly "C:\path\to\YourPlugin.dll" ^
  --message Update ^
  --changed "field1=value1,field2=value2"
```

Run `PluginSimulator.exe --help` for all options.

## Project Structure

```
src/PluginSimulator/
├── Program.cs                  CLI entry point + arg parsing
├── Authentication/             OAuth connection to Dataverse
├── Execution/                  Plugin loading, context building, audit replay
├── Models/                     Config, result, intercepted operation
├── Proxy/                      Fake IOrganizationService, ServiceProvider, Tracing, Logger
├── Reporting/                  Console output formatter
└── UI/                         WinForms shell + audit picker dialog
```

## Build

Requires:
- .NET Framework 4.7.1 dev pack
- MSBuild or Visual Studio 2019+

```
msbuild PluginSimulator.sln /p:Configuration=Release
```

To produce a single-file executable, use ILMerge (vendored under `tools/`).

## How It Works

```
Plugin DLL  ──► PluginLoader (reflection)
                    │
                    ▼
              IServiceProvider (ProxyServiceProvider)
                    │
        ┌───────────┼─────────────┬──────────────┐
        ▼           ▼             ▼              ▼
   IPluginExec  IOrgService   ITracing        ILogger
   Context     Factory                       (telemetry)
   (built       │
   from D365    ▼
   record)     ProxyOrganizationService
               ├── Reads  ──► real D365 (pass-through + dirty-cache merge)
               └── Writes ──► intercepted, logged, NOT applied
```

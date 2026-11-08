# AutoHotkey DLL process-exit regression

Run this test in an interactive Windows session with Explorer running and
64-bit AutoHotkey v2 and the x64 Windows SDK Debugging Tools (`cdb.exe`)
installed. It queries the desktop count without changing
desktops, windows, wallpapers, or hooks. It is not a headless CI test.

```powershell
cargo build -p dll --release --locked
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dll\tests\Test-ProcessExit.ps1
```

To compare an existing DLL with the source build:

```powershell
.\dll\tests\Test-ProcessExit.ps1 -LibraryPath C:\path\to\VirtualDesktopAccessor.dll -Iterations 3
```

Use `-AutoHotkeyPath` or `-DebuggerPath` for nonstandard installations.
`-TimeoutSeconds` defaults to 30; a timed-out debugger is terminated and
reported as a failure. Each run
creates a temporary folder containing per-process output and `results.json`.
`-OutputDirectory` selects another folder. Results are not uploaded.

The load-only case is a control. The count case initializes the cached COM
proxies, prints `VDA_EXIT_READY`, and calls `ExitApp(0)` with the DLL still
loaded. CDB breaks on first-chance access violations and process exit. The
parent requires the query output, no access violation, and a normal process
exit event. Checking only the process exit code is insufficient: on affected
systems the TLS cleanup exception can occur even though the exit code is zero.
The debugger command echo is excluded from exception-marker matching.
Missing DLLs/exports, failed desktop queries, timeouts, and native crashes all
fail the test. Windows crash dialogs are disabled only inside the probe.

This exercises the lifetime defect fixed by
[#110](https://github.com/Ciantic/VirtualDesktopAccessor/pull/110). The
`2024-12-16-windows11` release predates that fix. Unloading the DLL before exit
or testing only within a long-running host can conceal the defect, because the
cached COM objects then have different teardown ordering.

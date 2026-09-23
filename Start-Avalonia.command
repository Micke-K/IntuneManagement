#!/bin/sh
# Launcher for the Avalonia UI on macOS and Linux - the counterpart to Start.cmd.
#
# Named .command so macOS Finder treats it as double-clickable; the extension means
# nothing on Linux, where "./Start-Avalonia.command" works just as well.
#
# The only prerequisite is PowerShell 7.4+ (pwsh) on PATH; it brings its own .NET.
# See Docs/CrossPlatform.md.

set -e
DIR=$(cd "$(dirname "$0")" && pwd)

if ! command -v pwsh >/dev/null 2>&1; then
    echo "PowerShell 7.4+ (pwsh) is required but was not found on PATH." >&2
    echo "Install it from https://github.com/PowerShell/PowerShell and try again." >&2
    exit 1
fi

if [ "$(uname -s)" = "Darwin" ]; then
    # Cocoa only allows the GUI on the process main thread, which a normal pwsh
    # pipeline is not. The startup hook runs the script on that thread inside
    # this same pwsh (see UI/Avalonia/Bootstrap/MainThreadHook/StartupHook.cs).
    HOOK="$DIR/Bin/MainThreadHook/IntuneManagement.MainThreadHook.dll"
    if [ ! -f "$HOOK" ]; then
        echo "Main-thread hook is missing: $HOOK" >&2
        echo "Re-download the application, or rebuild it with UI/Avalonia/Bootstrap/Publish-MainThreadHook.ps1 (.NET SDK)." >&2
        exit 1
    fi
    export DOTNET_STARTUP_HOOKS="$HOOK${DOTNET_STARTUP_HOOKS:+:$DOTNET_STARTUP_HOOKS}"
    export IM_MAIN_THREAD_HOOK=1
fi

exec pwsh -NoProfile -File "$DIR/UI/Avalonia/Start-Avalonia.ps1" "$@"

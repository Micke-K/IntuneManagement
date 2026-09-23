# A wait that keeps the window alive.
#
# The engine runs on the UI thread in both backends. Write-Status pumps the UI
# once when it sets a message, but a plain Start-Sleep afterwards blocks that
# thread for the whole wait: no repaint, no input, and on Avalonia - which
# renders on the UI thread - not even the layout pass for the text just set.
# A Graph 429 back-off of 10 seconds, retried ten times, looked like a hang
# with the throttle message painted over the line above it.
#
# Wait-UIAware sleeps in slices, pumps the message loop between them, and can
# refresh a status detail with the seconds remaining so the wait visibly
# counts down. Headless (no UI provider) it is just a sliced sleep.

function Wait-UIAware
{
    param(
        [Parameter(Mandatory = $true)][double]$Seconds,
        # Status detail with {0} for the whole seconds remaining, refreshed on
        # every change. Omit to leave the status line alone.
        [string]$DetailFormat,
        [int]$SliceMilliseconds = 250
    )

    if($Seconds -le 0) { return }
    if($SliceMilliseconds -lt 50) { $SliceMilliseconds = 50 }

    $slices    = [int][Math]::Ceiling(($Seconds * 1000) / $SliceMilliseconds)
    $lastShown = -1

    for($i = 0; $i -lt $slices; $i++)
    {
        if($DetailFormat)
        {
            $remaining = [int][Math]::Ceiling($Seconds - (($i * $SliceMilliseconds) / 1000))
            if($remaining -ne $lastShown)
            {
                $lastShown = $remaining
                Write-Status -Detail ($DetailFormat -f $remaining) -SkipLog -Force
            }
        }

        Start-Sleep -Milliseconds $SliceMilliseconds

        if($script:UIProvider)
        {
            try { $script:UIProvider.InvokeUIMessagePump() } catch { }
        }
    }
}

# The Graph throttle wording, shared by the batch and single-request 429 paths.
function Wait-GraphThrottle
{
    param(
        [Parameter(Mandatory = $true)][double]$Seconds,
        [string]$BatchType = 'Graph',
        [int]$Queued = 1
    )

    Wait-UIAware -Seconds $Seconds -DetailFormat ("{0}: throttled by Graph - waiting {{0}}s ({1} request(s) queued)" -f $BatchType, $Queued)
}

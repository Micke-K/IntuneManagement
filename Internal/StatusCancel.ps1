# Cancel affordance for the status overlay.
#
# The app is effectively single-threaded: a long wait (device-code sign-in, a
# browser round trip) runs a loop on the UI thread and keeps the window alive by
# pumping messages (Invoke-UIPump). Because that pump already runs inside both
# login wait loops, a button on the status overlay DOES get its click dispatched -
# which is what makes a Cancel button possible at all without a threading rewrite.
#
# This file owns only the state and the contract. It deliberately holds no UI
# types (R12): the button lives in each backend, and its click handler calls
# Request-StatusCancel by NAME - which also satisfies the Avalonia rule that an
# event handler may not rely on captured locals.
#
# Two separate things are recorded:
#
#   the ACTION   what actually stops the pending operation (cancel a
#                CancellationTokenSource, stop an HttpListener). Supplied by the
#                caller that knows how, via Write-Status -OnCancel.
#   the REQUEST  a sticky flag the waiting loop polls, so it can break out
#                promptly instead of running to its timeout.
#
# Both matter. Without the action the operation keeps running invisibly to its
# timeout; without the flag the loop would not notice for up to its poll interval
# (or at all, if the underlying wait is not cancellable).

$script:StatusCancelAction = $null
$script:StatusCancelRequested = $false

# Arm the overlay's cancel button. Called from Write-Status -OnCancel; also
# callable directly by code that wants the action armed without touching the
# status text.
function Set-StatusCancelAction
{
    param([scriptblock]$Action)

    $script:StatusCancelAction = $Action
    # Arming starts a NEW cancellable operation, so a request left over from a
    # previous one must not immediately cancel it.
    $script:StatusCancelRequested = $false
}

# Disarm. Called when the status overlay is cleared, and from the finally block of
# whatever armed it, so a stale action can never fire against a finished operation.
function Clear-StatusCancelAction
{
    $script:StatusCancelAction = $null
    $script:StatusCancelRequested = $false
}

# True once the user has asked to cancel. Waiting loops poll this.
function Test-StatusCancelRequested
{
    return ($script:StatusCancelRequested -eq $true)
}

# Sleep for up to -Seconds, in slices, pumping the UI between them so the overlay's
# Cancel button stays clickable, and return as soon as cancel is requested.
# Returns $true when the wait ended because of a cancel.
#
# A plain Start-Sleep of several seconds blocks the single UI thread outright: no
# repaint, no click dispatch, so a Cancel button would be dead for the whole nap.
# Any polling loop that wants to be cancellable has to wait THIS way.
function Wait-StatusCancel
{
    param([int]$Seconds = 1, [int]$SliceMilliseconds = 100)

    if($Seconds -le 0) { return (Test-StatusCancelRequested) }

    $deadline = [DateTime]::UtcNow.AddSeconds($Seconds)
    while([DateTime]::UtcNow -lt $deadline)
    {
        if(Test-StatusCancelRequested) { return $true }
        Invoke-UIPump
        Start-Sleep -Milliseconds $SliceMilliseconds
    }
    return (Test-StatusCancelRequested)
}

# Invoked by the overlay's Cancel button. Sets the flag first so the flag is
# observable even if the action throws, then runs the action.
function Request-StatusCancel
{
    $script:StatusCancelRequested = $true

    $action = $script:StatusCancelAction
    if(-not $action) { return }

    # One shot: a second click must not run the action again (stopping an already
    # stopped listener throws).
    $script:StatusCancelAction = $null

    Write-Log "Cancel requested from the status window"
    try { & $action }
    catch { Write-LogError "Status cancel action failed" $_.Exception }
}

function Get-DocumentationOutput {
    <#
    .SYNOPSIS
        Gets the registered documentation output providers.

    .DESCRIPTION
        Returns the documentation output providers available to
        Start-GraphBulkDocumentation and the documentation user interfaces.
    #>
    [CmdletBinding()]
    param()

    [DocumentationRegistry]::Outputs |
        Sort-Object Name |
        Select-Object Name, Value
}

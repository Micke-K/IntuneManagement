#ImportOrder 31

class CompareOutputProviderBase
{
    [string] $Name
    [string] $Value
    [string] $Extension

    CompareOutputProviderBase([string]$name, [string]$value, [string]$extension)
    {
        $this.Name      = $name
        $this.Value     = $value
        $this.Extension = $extension
    }

    [string] FormatRows([object[]]$rows, [string[]]$props) { return "" }
    [string] ToString()                                    { return $this.Name }
}

class CompareCSVOutputProvider : CompareOutputProviderBase
{
    [string] $Delimiter = ";"

    CompareCSVOutputProvider() : base("CSV", "csv", "csv")
    {
        $this.Delimiter = ";"
    }

    [string] FormatRows([object[]]$rows, [string[]]$props)
    {
        $selected = @($rows | Select-Object -Property $props)
        if($this.Delimiter -and $this.Delimiter.Length -eq 1)
        {
            return ($selected | ConvertTo-Csv -NoTypeInformation -Delimiter ([char]$this.Delimiter)) -join [System.Environment]::NewLine
        }
        return ($selected | ConvertTo-Csv -NoTypeInformation) -join [System.Environment]::NewLine
    }
}

class CompareJsonOutputProvider : CompareOutputProviderBase
{
    CompareJsonOutputProvider() : base("JSON", "json", "json") {}

    [string] FormatRows([object[]]$rows, [string[]]$props)
    {
        return $rows | Select-Object -Property $props | ConvertTo-Json -Depth 20
    }
}

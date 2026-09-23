# Atlassian (Confluence Storage Format) output provider.
#
# Emits Confluence-compatible XHTML for pasting into a Confluence page editor
# (Rich Text -> Source view) or POSTing via the Confluence REST API as
# `representation=storage`. Structurally identical to the HTML provider
# (BasicInfo / FilteredSettings / ComplianceActions / ApplicabilityRules /
# Assignments / CustomTables + a per-run table of contents), only the emitted
# markup differs: no CSS embed, no <HTML>/<body> wrapper, code and long-text
# blocks use Confluence macros (<ac:structured-macro name='code'|'expand'>).
#
# Heading anchors: every heading carries id='<anchor>' plus an inline `anchor`
# macro of the same name, and the table of contents links that name. Anchors are
# positional - 'section-N' for every heading, numbered in emission order across
# the whole run (including headings kept out of the TOC),
# so a name is never reused. 'table-N' is reserved for the -ToT caption form,
# which no call site in this provider currently uses - table captions are plain
# level-6 headings here, as in the HTML provider. The id= attribute is a
# documented contract for consumers that parse the generated file before it is
# published (Confluence itself discards the attribute); it always equals the
# macro name. Changing the naming scheme is a breaking change for those
# consumers.
#
# Options (via $Options.Outputs.atlassian):
#   AtlassianDocumentName     - target file path. Supports Expand-FileName
#                               tokens (%MyDocuments%, %Organization%, %Date%,
#                               %DateTime%). Default: %MyDocuments%\%Organization%-%Date%.html
#   AtlassianDocumentFileType - 'Full' (single file) or 'Object' (one file per
#                               policy + a TOC index file). Default: 'Full'.
#   AtlassianTitleProperty    - H1 title of the index page. Default: 'Intune documentation'.
#   AtlassianOpenFile         - After writing, launch the file with the OS
#                               default handler. Set $false for CI runs.
#                               Default: $true.

function Invoke-InitializeAtlassianOutput {
    Add-DocumentationOutputProvider ([PSCustomObject]@{
        Name              = "Atlassian"
        Value             = "atlassian"
        # Path metadata (see DocumentationOutputHTML.ps1 header comment). Drives
        # the "default output folder" inference on the bulk-doc form, whose
        # Atlassian options panel mirrors the HTML one minus the CSS row.
        PrimaryPathOption = "AtlassianDocumentName"
        PathIsFolder      = $false
        PreProcess        = { Invoke-AtlassianPreProcessItems @args }
        NewObjectGroup    = { Invoke-AtlassianNewObjectGroup @args }
        NewObjectType     = { Invoke-AtlassianNewObjectType @args }
        Process           = { Invoke-AtlassianProcessItem @args }
        PostProcess       = { Invoke-AtlassianPostProcessItems @args }
        ProcessAllObjects = { Invoke-AtlassianProcessAllObjects @args }
    })
}

function Invoke-AtlassianPreProcessItems {
    $script:atlSectionAnchors      = @()
    $script:atlTotAnchors          = @()
    # Anchor numbering is deliberately NOT derived from the anchor lists above:
    # a -SkipTOC heading is emitted (and needs an anchor) without being listed,
    # so a list-derived number would be handed out twice. See Add-AtlassianHeader.
    $script:atlSectionCount        = 0
    $script:atlTotCount            = 0
    $script:atlBody                = $null
    $script:atlCurrentItemFileName = $null

    $fileName = Get-DocumentationOutputOption atlassian "AtlassianDocumentName" ""
    if (-not $fileName) { $fileName = "%MyDocuments%\%Organization%-%Date%.html" }
    $fileName = Expand-FileName $fileName

    $script:atlOutFile      = $fileName
    $script:atlDocumentPath = [IO.Path]::GetDirectoryName($fileName)
    $script:atlOutputType   = Get-DocumentationOutputOption atlassian "AtlassianDocumentFileType" "Full"

    if ($script:atlOutputType -eq "Object") {
        Write-Log "Atlassian: document one file for each object + index file"
    }
    else {
        Write-Log "Atlassian: document one single file for all objects"
        $script:atlOutputType = "Full"
        $script:atlBody       = [System.Text.StringBuilder]::new()
    }
}

function Invoke-AtlassianPostProcessItems {
    $userName = $null
    $mail     = ""
    $me = Get-CurrentUser
    if ($me) {
        if ($me.givenName -and $me.surname) {
            $userName = "$($me.givenName) $($me.surname)"
        }
        else {
            $userName = $me.displayName
        }
        if ($me.mail) { $mail = " ($($me.mail))" }
    }

    $orgName = Get-CurrentOrganizationName

    $title = Get-DocumentationOutputOption atlassian "AtlassianTitleProperty" "Intune documentation"
    if (-not $title) { $title = "Intune documentation" }

    # Escaped for the same reason heading text is: an '&' in the tenant's
    # organization name or in the configured title would make the document body
    # invalid XML, and Confluence rejects the upload of the whole page.
    $content = [System.Text.StringBuilder]::new()
    [void]$content.AppendLine("<h1>$(Get-AtlassianXmlText $title)</h1>")

    if (-not ((Get-DocumentationOption "SkipDocumentInfo" $false) -eq $true)) {
        if ($orgName)  { [void]$content.AppendLine("Organization: $(Get-AtlassianXmlText $orgName)") }
        if ($userName) { [void]$content.AppendLine("Generated by: $(Get-AtlassianXmlText "$userName$mail")") }
        [void]$content.AppendLine("Generated: $((Get-Date).ToShortDateString()) $((Get-Date).ToLongTimeString())")
    }

    if ($script:atlSectionAnchors.Count -gt 0) {
        [void]$content.AppendLine("<h2>Table of Contents</h2>")
        Add-AtlassianTableOfContents $content
    }

    $text = $content.ToString()
    if ($script:atlOutputType -eq "Full" -and $script:atlBody) {
        $text += $script:atlBody.ToString()
    }

    Save-DocumentationFile $text $script:atlOutFile -OpenFile:((Get-DocumentationOutputOption atlassian "AtlassianOpenFile" $true) -eq $true)
}

function Invoke-AtlassianNewObjectGroup {
    param($groupId)
    $script:atlObjectHeaderLevel = 2
    Add-AtlassianHeader (Get-DocObjectTypeString $groupId)
}

function Invoke-AtlassianNewObjectType {
    param($objectTypeName)
    $script:atlObjectHeaderLevel = 3
    Add-AtlassianHeader $objectTypeName
    $script:atlObjectHeaderLevel = 4
}

function Invoke-AtlassianProcessAllObjects {
    param($documentationInfo)
    # ScopeTags consolidated table is deferred (matches HTML provider stub).
}

function Invoke-AtlassianProcessItem {
    param($PolicyObject, $documentedObj)

    if (-not $documentedObj -or -not $PolicyObject) { return }

    # A documented object may ask to be titled by something other than its display
    # name (see Get-DocumentationDisplayName). Headings and captions follow it; the
    # file name below deliberately does not.
    $objName   = Get-DocumentationDisplayName $PolicyObject $documentedObj
    $script:docDisplayName = $objName
    $typeTitle = $PolicyObject.PolicyType.Title

    if ($script:atlOutputType -eq "Object") {
        # Table numbering restarts per file (each object is its own page), section
        # numbering does not - see the header comment on anchor uniqueness.
        $script:atlTotAnchors          = @()
        $script:atlTotCount            = 0
        $script:atlBody                = [System.Text.StringBuilder]::new()
        $script:atlCurrentItemFileName = Get-AtlassianObjectFileName $PolicyObject
    }

    Add-AtlassianHeader $objName

    try {
        foreach ($tableType in @("BasicInfo","FilteredSettings")) {
            if ($tableType -eq "BasicInfo") {
                $properties = @("Name","Value")
                $lngId      = "SettingDetails.basics"
            }
            else {
                $properties = if ($documentedObj.DefaultDocumentationProperties) {
                    $documentedObj.DefaultDocumentationProperties
                } else {
                    @("Name","Value")
                }
                $lngId = "TableHeaders.settings"
            }

            if (($documentedObj.$tableType | Measure-Object).Count -gt 0) {
                Add-AtlassianTableItems $PolicyObject $typeTitle $documentedObj.$tableType $properties $lngId -AddCategories -AddSubcategories
            }
        }

        if (($documentedObj.ComplianceActions | Measure-Object).Count -gt 0) {
            Add-AtlassianTableItems $PolicyObject $typeTitle $documentedObj.ComplianceActions @("Action","Schedule","MessageTemplate","EmailCC") "Category.complianceActionsLabel"
        }

        if (($documentedObj.ApplicabilityRules | Measure-Object).Count -gt 0) {
            Add-AtlassianTableItems $PolicyObject $typeTitle $documentedObj.ApplicabilityRules @("Rule","Property","Value") "SettingDetails.applicabilityRules"
        }

        Add-AtlassianObjectScripts $documentedObj

        foreach ($customTable in ($documentedObj.CustomTables | Sort-Object -Property Order)) {
            Add-AtlassianTableItems $PolicyObject $typeTitle $customTable.Values $customTable.Columns $customTable.LanguageId -AddCategories -AddSubcategories
        }

        if (($documentedObj.Assignments | Measure-Object).Count -gt 0) {
            if ($documentedObj.Assignments[0].RawIntent) {
                $properties = @("GroupMode","Group","Filter","FilterMode")
                $settingsObj = $documentedObj.Assignments | Where-Object { $null -ne $_.Settings } | Select-Object -First 1
                if ($settingsObj) {
                    foreach ($objProp in $settingsObj.Settings.Keys) {
                        if ($objProp -in $properties)               { continue }
                        if ($objProp -in @("Category","RawIntent")) { continue }
                        $properties += "Settings.$objProp"
                    }
                }
            }
            else {
                $hasFilter = $false
                foreach ($a in $documentedObj.Assignments) {
                    if ($a.PSObject.Properties.Name -contains "FilterMode") { $hasFilter = $true; break }
                }
                $properties = @("Group")
                if ($hasFilter) { $properties += @("Filter","FilterMode") }
            }

            Add-AtlassianTableItems $PolicyObject $typeTitle $documentedObj.Assignments $properties "TableHeaders.assignments" -AddCategories
        }
    }
    catch {
        Write-LogError "Failed to process object $objName" $_.Exception
    }

    if ($script:atlOutputType -eq "Object") {
        $fileName = Join-Path $script:atlDocumentPath $script:atlCurrentItemFileName
        Save-DocumentationFile $script:atlBody.ToString() $fileName
        $script:atlBody = $null
    }
}

function Get-AtlassianObjectFileName {
    param($PolicyObject)

    $objName = if ($PolicyObject.Name) { [string]$PolicyObject.Name } else { 'Unnamed policy' }
    $id      = if ($PolicyObject.Id)   { [string]$PolicyObject.Id }   else { $null }
    $typeId  = if ($PolicyObject.PolicyType -and $PolicyObject.PolicyType.Id) { [string]$PolicyObject.PolicyType.Id } else { $null }
    $suffix  = if ($typeId -and $id) { " [$typeId-$id]" }
               elseif ($id)          { " [$id]" }
               else                  { '' }
    return Remove-InvalidFileNameChars "$objName$suffix.html"
}

# Escape text for storage format. Confluence storage format is strict XML and
# declares only the five XML built-in entities, so a bare '&' or '<' arriving
# from tenant data (policy names, localized captions) makes the whole document
# body malformed and Confluence rejects the upload - not just that heading.
# Values inside table cells go through Set-AtlassianText, which does this plus
# the code/expand macro wrapping; headers and TOC labels need only the escape.
function Get-AtlassianXmlText {
    param([string]$Text)

    if (-not $Text) { return "" }
    return $Text.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
}

# The author-controlled link target for a heading.
#
# Confluence strips author-specified id= attributes when it converts storage
# format to ADF, so an id alone is not linkable - '#name' only ever resolves to
# an anchor macro or to Confluence's own heading-text-derived anchor.
#
# The macro stays inline inside the heading until the downstream rewrite observed
# in Docs/AtlassianAnchorVerification-2026-09-15.md has been attributed. Moving it
# to a preceding paragraph before that measurement was an unverified fix that could
# add a blank line at every heading without changing the publisher's output. Inline
# placement is legal ADF (`anchor` is an inline macro and headings accept inline
# content), keeps the jump target on the heading, and is the known baseline while
# the runbook and Confluence import paths are tested separately.
#
# Attributes are single-quoted like every other macro in this file. Consumers
# JSON-escape the document body before publishing it, and a double-quoted
# attribute arrives as ac:name=\"anchor\" and breaks the macro.
function Get-AtlassianAnchorMacro {
    param([string]$Name)

    if (-not $Name) { return "" }
    return "<ac:structured-macro ac:name='anchor' ac:schema-version='1'>" +
           "<ac:parameter ac:name=''>$Name</ac:parameter>" +
           "</ac:structured-macro>"
}

# Build the href for a TOC entry: a percent-encoded relative file name plus the
# anchor fragment, safe to drop into a single-quoted attribute.
#
# In 'Object' mode the file name comes from the policy name (see
# Get-AtlassianObjectFileName), and only path-invalid characters are stripped
# from it. '&', '<' and "'" therefore survive - one of them makes the whole
# document body malformed XML, or terminates the attribute - and a '#' in a
# policy name would open a second fragment and retarget the link. Percent-encode
# the file-name component (never the '#' that separates the fragment), then
# XML-escape what is left.
function Get-AtlassianHref {
    param([string]$FileName, [string]$Anchor)

    $target = ""
    if ($FileName) {
        # A bare file name, no directory separators, so encoding the whole string
        # is correct. EscapeDataString covers ' on .NET Core but not on every
        # .NET Framework version; the explicit replace is a no-op when it did.
        $target = [Uri]::EscapeDataString($FileName).Replace("'", "%27")
    }

    return Get-AtlassianXmlText "$target#$Anchor"
}

function Add-AtlassianHeader {
    param(
        [string]$HeaderText,
        [int]$Level = $script:atlObjectHeaderLevel,
        [switch]$ToT,
        [switch]$SkipTOC
    )

    if ($ToT) {
        # 'Table N. ' is a visible caption prefix - that is what the Markdown
        # provider does with it. It used to be prepended to the id instead of to
        # the text, producing id="Table 1. table-1": spaces and a period in an
        # identifier, and no visible numbering anywhere. The number matches the
        # 'table-N' anchor below.
        $HeaderText = "Table $($script:atlTotCount + 1). $HeaderText"
    }

    if ($script:atlBody) {
        # Every heading that reaches the body consumes a number, whether or not it
        # is listed in the TOC. Numbering off $atlSectionAnchors.Count instead gave
        # a -SkipTOC heading (script captions, Add-AtlassianObjectScripts) the same
        # 'section-N' as the next listed heading: two anchor macros with one name,
        # so the TOC entry for the policy jumped to the script caption above it.
        if ($ToT) {
            $script:atlTotCount++
            $sectionAnchor = "table-$($script:atlTotCount)"
        }
        else {
            $script:atlSectionCount++
            $sectionAnchor = "section-$($script:atlSectionCount)"
        }

        # id= is kept even though Confluence discards it: consumers parse the
        # generated file (before upload) and read the anchor from it, so it must
        # stay byte-identical to the anchor macro name.
        $anchorMacro = Get-AtlassianAnchorMacro $sectionAnchor
        [void]$script:atlBody.AppendLine("<h$Level id='$sectionAnchor'>$anchorMacro$(Get-AtlassianXmlText $HeaderText)</h$Level>")
        $fileName = $script:atlCurrentItemFileName
    }
    else {
        $sectionAnchor = $null
        $fileName      = $null
    }

    if ($ToT) {
        $script:atlTotAnchors += [PSCustomObject]@{
            Name = $HeaderText; Anchor = $sectionAnchor; Level = $Level; FileName = $fileName
        }
    }
    elseif (-not $SkipTOC) {
        $script:atlSectionAnchors += [PSCustomObject]@{
            Name = $HeaderText; Anchor = $sectionAnchor; Level = $Level; FileName = $fileName
        }
    }
}

# Render the per-run table of contents into $Content.
#
# Each entry links the anchor its heading actually emitted. Deriving the target
# from the heading text instead - which this did - is unreliable twice over:
# duplicate policy names all resolved to the first occurrence (Confluence
# disambiguates its own text-derived anchors with .1/.2, which a naive
# derivation cannot reproduce), and every character other than a plain space
# survived into the fragment, including '.', '(', ')', ':' and U+00A0.
#
# Confluence still generates its text-derived anchors, so externally saved
# '#Policy-Name' links keep working; only the TOC moves to the reliable form.
function Add-AtlassianTableOfContents {
    param(
        [System.Text.StringBuilder]$Content,
        [int]$MaxLevel = 4
    )

    foreach ($header in $script:atlSectionAnchors) {
        if ($MaxLevel -gt 0 -and $header.Level -gt $MaxLevel) { continue }
        # Nest visually via non-breaking-space padding - Confluence doesn't honour
        # CSS anchor-level classes on imported storage-format content. Use the
        # numeric reference &#160; (not the HTML entity &nbsp;): storage format is
        # strict XML and only declares the five XML built-in entities.
        $indent = ""
        for ($i = 2; $i -lt $header.Level; $i++) { $indent += "&#160;&#160;" }

        $label = Get-AtlassianXmlText $header.Name
        if ($header.Anchor) {
            $href = Get-AtlassianHref $header.FileName $header.Anchor
            [void]$Content.AppendLine("$indent<a href='$href'>$label</a>")
        }
        else {
            # Registered while no body was open (a group/type header in 'Object'
            # mode), so the heading exists in no file and has no anchor. A
            # '#'-only href would jump to the top of the page instead; emit the
            # label as plain text.
            [void]$Content.AppendLine("$indent$label")
        }
    }
}

function Add-AtlassianTableItems {
    param(
        $PolicyObject,
        [string]$TypeTitle,
        $Items,
        [string[]]$Properties,
        [string]$LngId,
        [switch]$AddCategories,
        [switch]$AddSubcategories,
        $CaptionOverride
    )

    if ($CaptionOverride) {
        $caption = $CaptionOverride
    }
    elseif ($LngId -and $PolicyObject) {
        $caption = "$((Get-LanguageString $LngId)) - $(Get-DocCaptionName $PolicyObject)"
    }
    elseif ($PolicyObject) {
        $caption = "$(Get-DocCaptionName $PolicyObject) ($TypeTitle)"
    }
    else {
        $caption = $TypeTitle
    }

    $table = [System.Text.StringBuilder]::new()
    [void]$table.AppendLine("<table>")
    [void]$table.AppendLine("<tr>")

    $columnCount = 0
    foreach ($prop in $Properties) {
        [void]$table.AppendLine("<th>$((Invoke-DocTranslateColumnHeader $prop.Split('.')[-1]))</th>")
        $columnCount++
    }
    [void]$table.AppendLine("</tr>")

    $curCategory    = ""
    $curSubCategory = ""

    foreach ($itemObj in $Items) {
        if ($itemObj.Category -and $curCategory -ne $itemObj.Category -and $AddCategories) {
            [void]$table.AppendLine("<tr><td colspan='$columnCount'><strong>$($itemObj.Category)</strong></td></tr>")
            $curCategory    = $itemObj.Category
            $curSubCategory = ""
        }

        if ($itemObj.SubCategory -and $curSubCategory -ne $itemObj.SubCategory -and $AddSubcategories) {
            [void]$table.AppendLine("<tr><td colspan='$columnCount'><em>$($itemObj.SubCategory)</em></td></tr>")
            $curSubCategory = $itemObj.SubCategory
        }

        try {
            [void]$table.AppendLine("<tr>")

            $curCol = 0
            foreach ($prop in $Properties) {
                $curCol++
                try {
                    $propArr  = $prop.Split('.')
                    $tmpObj   = $itemObj
                    $propName = $propArr[-1]
                    for ($x = 0; $x -lt ($propArr.Count - 1); $x++) {
                        $tmpObj = $tmpObj."$($propArr[$x])"
                    }

                    if ($propName -eq "Value" -and ($itemObj.FullValueTable | Measure-Object).Count -gt 0) {
                        [void]$table.AppendLine("<td><table><tr>")
                        foreach ($colProp in $itemObj.FullValueTable[0].PSObject.Properties) {
                            [void]$table.AppendLine("<th>$($colProp.Name)</th>")
                        }
                        [void]$table.AppendLine("</tr>")
                        foreach ($rowVal in $itemObj.FullValueTable) {
                            [void]$table.AppendLine("<tr>")
                            foreach ($colProp in $itemObj.FullValueTable[0].PSObject.Properties) {
                                [void]$table.AppendLine("<td>$((Set-AtlassianText $rowVal."$($colProp.Name)"))</td>")
                            }
                            [void]$table.AppendLine("</tr>")
                        }
                        [void]$table.AppendLine("</table></td>")
                    }
                    else {
                        $indent = ""
                        if ($curCol -eq 1 -and $itemObj.Level) {
                            try {
                                # One indent unit per nesting level (Level 1 = first
                                # indent), matching the HTML/MD/Word providers. Was
                                # off-by-one ($i started at 1), so Level-1 children
                                # rendered flush. Negative levels produce no indent.
                                $level = [int]$itemObj.Level
                                # &#160; (numeric non-breaking space) not &nbsp; —
                                # Confluence storage format is strict XML and only
                                # declares the five XML built-in entities, so &nbsp;
                                # is dropped/rejected. Numeric refs always render.
                                for ($i = 0; $i -lt $level; $i++) { $indent += "&#160;&#160;" }
                            } catch {}
                        }
                        [void]$table.AppendLine("<td>$($indent)$((Set-AtlassianText $tmpObj.$propName))</td>")
                    }
                }
                catch {
                    Write-LogError "Failed to add property value for $prop" $_.Exception
                }
            }
        }
        catch {
            Write-Log "Failed to process property" 2
        }
        finally {
            [void]$table.AppendLine("</tr>")
        }
    }

    [void]$table.AppendLine("</table>")
    [void]$script:atlBody.Append($table.ToString())
    # No -ToT: the caption is a plain level-6 heading, matching the HTML provider.
    # Passing -ToT here would add visible 'Table N. ' numbering (what the Markdown
    # provider does) and move the caption into $atlTotAnchors. That is a formatting
    # decision for the HTML and Atlassian outputs together, not a port detail.
    Add-AtlassianHeader $caption -Level 6
}

# Confluence Storage Format text emitter. Escapes HTML special chars in plain
# text; wraps XML-looking values in a `code` macro; wraps long text (>250
# chars) in an `expand` macro with a first-line summary as the caption.
function Set-AtlassianText {
    param([string]$Text, [switch]$NoCodeBlock)

    if (-not $Text) { return }

    $txtSummary = ""
    if ($Text.Length -gt 250) {
        $summaryMax = 40
        $idx = $Text.IndexOfAny(@("`r","`n"))
        if ($idx -gt 10 -and $idx -lt 50) { $summaryMax = $idx }
        $txtSummary = $Text.Substring(0, $summaryMax)
    }

    $isCode = $false
    if (-not $NoCodeBlock) {
        $trim = $Text.Trim()
        if ($trim.StartsWith("<") -and $trim.EndsWith(">")) {
            $isCode = $true
            $Text = "<ac:structured-macro ac:name='code' ac:schema-version='1'>" +
                    "<ac:parameter ac:name='language'>xml</ac:parameter>" +
                    "<ac:plain-text-body><![CDATA[$Text]]></ac:plain-text-body>" +
                    "</ac:structured-macro>"
            if ($txtSummary) {
                $txtSummary = $txtSummary.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
            }
        }
    }

    if (-not $isCode) {
        $Text = $Text.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;') #.Replace("`r`n",'<br />').Replace("`n",'<br />')
    }

    if ($txtSummary) {
        "<ac:structured-macro ac:name='expand' ac:schema-version='1'>" +
        "<ac:parameter ac:name='title'>$txtSummary...</ac:parameter>" +
        "<ac:rich-text-body>$Text</ac:rich-text-body>" +
        "</ac:structured-macro>"
    }
    else {
        $Text
    }
}

function Add-AtlassianObjectScripts {
    param($documentedObj)

    foreach ($scriptItem in $documentedObj.Scripts) {
        if (-not $scriptItem.ScriptContent -or -not $scriptItem.Caption) { continue }
        [void]$script:atlBody.AppendLine("<ac:structured-macro ac:name='code' ac:schema-version='1'>")
        [void]$script:atlBody.AppendLine("<ac:parameter ac:name='language'>powershell</ac:parameter>")
        [void]$script:atlBody.AppendLine("<ac:plain-text-body><![CDATA[")
        [void]$script:atlBody.AppendLine($scriptItem.ScriptContent)
        [void]$script:atlBody.AppendLine("]]></ac:plain-text-body>")
        [void]$script:atlBody.AppendLine("</ac:structured-macro>")
        Add-AtlassianHeader $scriptItem.Caption -Level 6 -SkipTOC
    }
}

Invoke-InitializeAtlassianOutput

# ADMX/ADML data helpers for the Intune Tools ADMX Import + Reg Values tools.
# Pure parsing/string/registry-support helpers + the ADMXReg* C# class defs - no
# XAML, no UIProvider. Moved out of the per-backend UI files where they were
# duplicated (and had drifted in code style) into one shared, module-internal
# home. Get-ADMXCategoryOMAURIPath was WPF-only but referenced by the Avalonia
# ADMX import; sharing it here also fixes that latent missing-function call.
# (architecture R11)

function Get-ADMXADMLString
{
    # Resolve an XML node's attribute (default: displayName) via the ADML string
    # table. ADMX uses "$(string.id)" placeholders pointing into the ADML
    # resources/stringTable; returns the raw attribute if no ADML is loaded or
    # the placeholder doesn't resolve.
    param($XmlNode, $Property = "displayName")

    $propValue = $XmlNode.$Property
    if(-not $script:_admxlngADML -or -not $XmlNode.$Property) { return $propValue }

    if($XmlNode.$Property.StartsWith('$('))
    {
        $tmp = $XmlNode.$Property.Substring(2, $XmlNode.$Property.Length - 3)
        $null, $strId = $tmp.Split('.')
    }
    else
    {
        $strId = $propValue
    }

    if($script:_admxStringTable.ContainsKey($strId)) {
        return $script:_admxStringTable[$strId]
    }
    return $propValue
}

function Get-ADMXADMLPresentationString
{
    # Look up an element's label within a presentation node — used when an ADMX
    # element doesn't have its own displayName attribute.
    param($PresentationInfo, $XmlNode)

    if(-not $script:_admxlngADML) { return $null }

    $presentationNode = $PresentationInfo.SelectSingleNode("./*[@refId='$($XmlNode.id)']")
    if($presentationNode) {
        return (?? $presentationNode.Label.'#text' $presentationNode.'#text')
    }
    return $XmlNode.id
}

function Get-ADMXCategoryIdPath
{
    # Walk a category's parentCategory chain to its root and return the
    # slash-joined path of internal names (used for tree-building dedup).
    param($CategoryId, $Delimiter = "/")

    $catObj = $script:_admxXML.policyDefinitions.categories.SelectSingleNode("$($script:_admxNSPrefix)category[@name='$CategoryId']", $script:_admxNS)
    $categories = @()
    while($catObj)
    {
        $categories += $catObj.name
        if($catObj.parentCategory.ref) {
            $catObj = $script:_admxXML.policyDefinitions.categories.SelectSingleNode("$($script:_admxNSPrefix)category[@name='$($catObj.parentCategory.ref)']", $script:_admxNS)
        }
        else { break }
    }
    [array]::Reverse($categories)
    return ($categories -join $Delimiter)
}

function Get-ADMXCategoryNamePath
{
    # Same chain as Get-ADMXCategoryIdPath but returns the resolved DISPLAY
    # names (via ADML). Cached per session because the same category-name
    # path is read for every policy under that category.
    param($CategoryId, $Delimiter = "/")

    if($script:_admxCategoryPaths.ContainsKey($CategoryId)) {
        return $script:_admxCategoryPaths[$CategoryId]
    }

    $catObj = $script:_admxXML.policyDefinitions.categories.SelectSingleNode("$($script:_admxNSPrefix)category[@name='$CategoryId']", $script:_admxNS)
    $categories = @()
    while($catObj)
    {
        $categories += Get-ADMXADMLString $catObj
        if($catObj.parentCategory.ref) {
            $catObj = $script:_admxXML.policyDefinitions.categories.SelectSingleNode("$($script:_admxNSPrefix)category[@name='$($catObj.parentCategory.ref)']", $script:_admxNS)
        }
        else { break }
    }
    [array]::Reverse($categories)

    $catPath = $categories -join $Delimiter
    $script:_admxCategoryPaths.Add($CategoryId, $catPath)
    return $catPath
}

function Get-ADMXCategoryOMAURIPath
{
    # OMA-URI uses ~ as the separator and keeps INTERNAL category names (not
    # localized display names) — separate function so the OMA-URI format
    # doesn't get accidentally entangled with the display-path formatting.
    param($CategoryId)

    $catObj = $script:_admxXML.policyDefinitions.categories.SelectSingleNode("$($script:_admxNSPrefix)category[@name='$CategoryId']", $script:_admxNS)
    $categories = @()
    while($catObj)
    {
        $categories += $catObj.name
        $catObj = $script:_admxXML.policyDefinitions.categories.SelectSingleNode("$($script:_admxNSPrefix)category[@name='$($catObj.parentCategory.ref)']", $script:_admxNS)
    }
    [array]::Reverse($categories)
    return ($categories -join "~")
}

function Get-ADMXPresentationNode
{
    # ADML defines a presentation for each policy (where the actual labels
    # live). Resolve from "$(presentation.id)" or a bare id.
    param($Item)

    if(-not $Item.Definition.presentation -or -not $script:_admxlngADML) { return $null }

    if($Item.Definition.presentation.StartsWith('$('))
    {
        $tmp = $Item.Definition.presentation.Substring(2, $Item.Definition.presentation.Length - 3)
        $null, $strId = $tmp.Split('.')
        return ($script:_admxlngADML.policyDefinitionResources.resources.presentationTable.presentation | Where-Object id -eq $strId)
    }
    return ($script:_admxlngADML.policyDefinitionResources.resources.presentationTable.presentation | Where-Object id -eq $Item.Definition.presentation)
}

function Set-ADMXSettingStatusText
{
    # Mirror the SettingStatus field as a human-friendly column value.
    param($SettingObj)

    switch ($SettingObj.SettingStatus)
    {
        0       { $SettingObj.SettingStatusText = "Disabled" }
        1       { $SettingObj.SettingStatusText = "Enabled"  }
        default { $SettingObj.SettingStatusText = "Not Configured" }
    }
}

function Add-ADMXRegClasses
{
    # Lazy-loaded the first time the Reg Values tool is opened. Same C# defs as
    # the original IntuneTools.psm1 — kept verbatim so DataContext bindings in
    # IntuneToolsADMXRegValues.xaml / IntuneToolsADMXAddRegPolicy.xaml don't
    # need to change.
    if(("ADMXRegPolicyElement" -as [type])) { return }

    $classDef = @"
    using System;
    using System.ComponentModel;
    using System.Collections.ObjectModel;

    public class ADMXRegPolicyElement : System.ComponentModel.INotifyPropertyChanged
    {
        public string DataType { get { return _dataType; } set { _dataType = value;  NotifyPropertyChanged("DataType"); NotifyPropertyChanged("DataTypeDisplayString");  } }
        private string _dataType = null;

        public string DataTypeDisplayString { get {
            if(DataType == "text")            return "String";
            else if(DataType == "multiText")  return "Multi-string";
            else if(DataType == "list")       return "List";
            else if(DataType == "decimal")    return "DWORD (32-bit)";
            else if(DataType == "longDecimal")return "QWORD (64-bit)";
            else                              return DataType;
        } }

        public string Key { get { return _key; } set { _key = value; NotifyPropertyChanged("Key"); } }
        private string _key;

        public string ValueName { get { return _valueName; } set { _valueName = value; NotifyPropertyChanged("ValueName"); } }
        private string _valueName;

        public string Value { get { return _value; } set { _value = value; NotifyPropertyChanged("Value"); } }
        private string _value;

        public string AttributePrefix { get { return _attributePrefix; } set { _attributePrefix = value; NotifyPropertyChanged("AttributePrefix"); } }
        private string _attributePrefix;

        public string AttributeSeparator { get { return _attributeSeparator; } set { _attributeSeparator = value; NotifyPropertyChanged("AttributeSeparator"); } }
        private string _attributeSeparator = ";";

        public bool AttributeSoft { get { return _attributeSoft; } set { _attributeSoft = value; NotifyPropertyChanged("AttributeSoft"); } }
        private bool _attributeSoft = false;

        public bool AttributeExpandable { get { return _attributeExpandable; } set { _attributeExpandable = value; NotifyPropertyChanged("AttributeExpandable"); } }
        private bool _attributeExpandable = false;

        public bool AttributeAdditive { get { return _attributeAdditive; } set { _attributeAdditive = value; NotifyPropertyChanged("AttributeAdditive"); } }
        private bool _attributeAdditive = false;

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(string propertyName = "")
        {
            if(PropertyChanged != null) { PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName)); }
        }
    }

    public class ADMXRegPolicy
    {
        public string PolicyName { get; set; }
        public string PolicyStatus { get; set; }
        public string Hive { get; set; }
        public string Key { get; set; }
        public bool StatusValueEnabled { get; set; }
        public string StatusValueName { get; set; }
        public System.Collections.ObjectModel.ObservableCollection<ADMXRegPolicyElement> PolicyElements { get; set; }

        public ADMXRegPolicy()
        {
            PolicyElements = new System.Collections.ObjectModel.ObservableCollection<ADMXRegPolicyElement>();
            Hive = "HKLM";
            PolicyStatus = "Enabled";
            StatusValueEnabled = true;
        }
    }

    public class ADMXRegProfile
    {
        public string ProfileName { get; set; }
        public string ProfileDescription { get; set; }
        public string PolicyType { get; set; }
        public System.Collections.ObjectModel.ObservableCollection<ADMXRegPolicy> ADMXPolicies { get; set; }
        public string XmlString { get; set; }

        public ADMXRegProfile()
        {
            ADMXPolicies = new System.Collections.ObjectModel.ObservableCollection<ADMXRegPolicy>();
            PolicyType = "Policy";
        }
    }
"@
    Add-NativeClass -ClassName "ADMXRegPolicyElement" -ClassDefinition $classDef
}

function Add-ADMXRegXmlAttribute
{
    param($XmlNode, [string]$Attribute, $Value)
    $attr = $XmlNode.OwnerDocument.CreateAttribute($Attribute)
    $attr.Value = $Value
    [void]$XmlNode.Attributes.Append($attr)
}

#ImportOrder 10
function Add-NativeClass
{
    param(
        [string]
        $ClassName,
        [string]
        $ClassDefinition,
        [String[]]
        $AssembliesPartialName
    )
    # `-as [type]` yields a Type object or $null - never $true. Comparing a Type to
    # $true coerces the right side to Type, which never matches, so this guard used
    # to be dead: Add-Type ran on every re-import, the CLR rejected the
    # already-loaded type, and the catch below logged "Failed to add type" every
    # time. Harmless but it made a clean re-import look broken.
    if ($ClassName -as [type]) { return }

    [Reflection.Assembly]::LoadWithPartialName("System.ComponentModel") | Out-Null
    foreach($Assembly in $AssembliesPartialName) {
        [Reflection.Assembly]::LoadWithPartialName($Assembly) | Out-Null
    }

    try {
        Write-Log "Add class $ClassName"
        Add-Type -TypeDefinition $ClassDefinition -IgnoreWarnings -ErrorAction Stop #-ReferencedAssemblies @('System.ComponentModel')
    }
    catch {
        Write-LogError "Failed to add type $($ClassName)" $_.Exception
        Write-LogDebug "Definition:`n$ClassDefinition"
    }
}

$classDef = @"
    using System;
    using System.ComponentModel;

    public class ObjectColumnInfo : System.ComponentModel.INotifyPropertyChanged
    {
        public string Property { get { return _property; } set { _property = value;  NotifyPropertyChanged("Property");  } }
        private string _property = null;

        public string Header { get { return _header; } set { _header = value;  NotifyPropertyChanged("Header");  } }
        private string _header = null;

        public ObjectColumnInfo(string Property, string Header)
        {
            _property = Property;
            _header = Header;
        }

        public override string ToString()
        {
            if(!String.IsNullOrEmpty(_header)) { return _header; }
            return _property ?? String.Empty;
        }

        public event PropertyChangedEventHandler PropertyChanged;  

        // This method is called by the Set accessor of each property.  
        // The CallerMemberName attribute that is applied to the optional propertyName  
        // parameter causes the property name of the caller to be substituted as an argument.  
        private void NotifyPropertyChanged(string propertyName = "")  
        {  
            if(PropertyChanged != null) { PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName)); }
        }        
    }

"@

Add-NativeClass -ClassName "ObjectColumnInfo" -ClassDefinition $classDef

$classDef = @"
    using System;
    using System.ComponentModel;

    public class NameValueObject : System.ComponentModel.INotifyPropertyChanged
    {
        public string Name { get { return _name; } set { _name = value;  NotifyPropertyChanged("Name");  } }
        private string _name = null;

        public string Value { get { return _value; } set { _value = value;  NotifyPropertyChanged("Value");  } }
        private string _value = null;

        public NameValueObject(string Name, string Value)
        {
            _name = Name;
            _value = Value;
        }

        public event PropertyChangedEventHandler PropertyChanged;  

        // This method is called by the Set accessor of each property.  
        // The CallerMemberName attribute that is applied to the optional propertyName  
        // parameter causes the property name of the caller to be substituted as an argument.  
        private void NotifyPropertyChanged(string propertyName = "")  
        {  
            if(PropertyChanged != null) { PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName)); }
        }        
    }

"@

Add-NativeClass -ClassName "NameValueObject" -ClassDefinition $classDef -AssembliesPartialName "System.ComponentModel" 

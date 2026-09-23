#ImportOrder 220

#########################################################################################
#
# Apple Enrollment Group
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class AppleEnrollmentGroup : IntunePolicyGroupBase
{
    AppleEnrollmentGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "AppleEnrollment"
        $this._Name = "Apple Enrollment"
        $this._Icon = "AppleEnrollmentTypes"
    }
}

#########################################################################################
#
# Apple Enrollment Types
#
#########################################################################################

# region Apple Enrollment Types
class AppleEnrollmentType : IntunePolicyTypeBase
{
    AppleEnrollmentType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        # _PolicyGroup must be AppleEnrollmentGroup (the group declared at the top of
        # this file). The original wiring pointed at AppleUpdateGroup, which is the
        # software-update group — wrong taxonomy.
        # _ObjectClass was never set, so IntunePolicyTypeBase.GetObject took the
        # "Object class is missing" branch and dropped every returned row, surfacing
        # as 'no Apple Enrollment Types objects matched' even though the API returned
        # data. Wire it up to AppleEnrollmentTypeObject (defined in this same file).
        $this._PolicyGroup = (Get-SingletonObject "AppleEnrollmentGroup")
        $this._APITitle = "Apple Enrollment Types"
        $this._PolicyName = "Apple Enrollment Type"
        $this._ID = "AppleEnrollmentTypes"
        $this._API = "deviceManagement/appleUserInitiatedEnrollmentProfiles"
        $this._Permissions = @("DeviceManagementServiceConfig.ReadWrite.All")
        $this._ObjectClass = "AppleEnrollmentTypeObject"
        $this._PropertiesToRemoveForUpdate = @('platform')

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class AppleEnrollmentTypeObject : IntunePolicyBase
{
    AppleEnrollmentTypeObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }
    AppleEnrollmentTypeObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "AppleEnrollmentType")
    }
}

#endregion

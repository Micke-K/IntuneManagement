#ImportOrder 220

#########################################################################################
#
# Multi Admin Approval (MAA)
#
# operationApprovalPolicies - the access policies that decide WHICH operations need a
# second administrator's approval, and which Entra groups may approve them.
#
# Exposed read-only on purpose. Creating or editing an access policy is itself an
# MAA-protected "Tenant Configuration" change that needs a second admin to approve, and
# a wrong policy can lock every administrator out of a workload. Export is enabled so
# the configuration can be documented, diffed and migrated by hand; import is not.
#
# The request queue (deviceManagement/operationApprovalRequests - where a 412 from a
# write lands; the approval code is the request id) is deliberately NOT a policy type:
# it is transient state keyed by GUID with nothing to export, and the only useful
# actions on it (approve / reject) would belong to a small tenant-admin tool.
#
# See Internal/MSGraphErrors.ps1 for the 400/403/412 classification that surfaces the
# approval code to the caller.
#
#########################################################################################

#region Operation Approval Policies

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class OperationApprovalPoliciesType : IntunePolicyTypeBase
{
    OperationApprovalPoliciesType() : Base() { $this.Init() }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "TenantAdminGroup")
        $this._PolicyName = "Multi Admin Approval Policies"
        $this._ID = "OperationApprovalPolicies"
        $this._HasPlatform = $false
        $this._API = "deviceManagement/operationApprovalPolicies"
        $this._Permissions = @("DeviceManagementRBAC.Read.All")
        # Read-only: see the file header. Editing an access policy is itself gated by
        # MAA and can lock admins out of a workload.
        $this._ShowButtons = @("View","Export")
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false
        $this._ObjectClass = "OperationApprovalPolicyObject"
        $this._Icon = "TenantSettings"

        if($null -ne $this._PolicyGroup) { $this._PolicyGroup.AddPolicyType($this) }
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        # approverGroupIds are Entra group ids, so they are meaningless in another
        # tenant without translation. Capture them the same way Conditional Access
        # captures its include/exclude groups, so the export carries the sidecars a
        # future migration would need.
        $ids = @()
        foreach($id in @($PolicyObject.JsonObject.approverGroupIds))
        {
            if([String]::IsNullOrWhiteSpace($id)) { continue }
            if($id -in $ids) { continue }

            $ids += $id
            Add-GraphMigrationObject $id "groups" "Group" ([IO.Path]::GetDirectoryName($PathToFile)) $PolicyObject._TokenId
        }
    }
}

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class OperationApprovalPolicyObject : IntunePolicyBase
{
    OperationApprovalPolicyObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }
    OperationApprovalPolicyObject() : Base() { $this.Init() }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "OperationApprovalPoliciesType")
    }
}

#endregion

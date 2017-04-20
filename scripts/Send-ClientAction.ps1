<#
.SYNOPSIS
    Invoke commands remotely on an SCCM Client for a system or systems.

.DESCRIPTION
    This function allows you to remotely trigger some of the more common actions that you would find on the local
    Configuration Manager console.

.PARAMETER -ComputerName
    Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address. When the remote computer
    is in a different domain than the local computer, the fully qualified domain name is required.

    This command defaults to localhost.

.PARAMETER -Action
    Specifies the action to be taken on the SCCM Client.  The available actions are as follows:
        HardwareInv - Runs a Hardware Inventory Cycle on the target machine.
        SoftwareInv - Runs a Software Inventory Cycle on the target machine.
        UpdateScan  - Runs a Software Updates Scan Cycle on the target machine.
        MachinePol  - Runs a Machine Policy Retrieval and Evaluation Cycle on the target machine.
        UserPolicy  - Runs a User Policy Retrieval and Evaluation Cycle on the target machine.
        FileCollect - Runs a File Collection Cycle on the target machine.

.INPUTS
    (original function only) You can pipe a computer name to Invoke-CMClient


.EXAMPLE
    Send-ClientAction -ComputerName server01 -Action HardwareInv

    The above command will invoke the Configuration Manager Client's Hardware Inventory Cycle on the targeted computer.  The return will look like the following:

       
        __GENUS          : 1
        __CLASS          : __PARAMETERS
        __SUPERCLASS     :
        __DYNASTY        : __PARAMETERS
        __RELPATH        : __PARAMETERS
        __PROPERTY_COUNT : 1
        __DERIVATION     : {}
        __SERVER         : server01
        __NAMESPACE      : ROOT\ccm
        __PATH           : \\server01\ROOT\ccm:__PARAMETERS
        ReturnValue      :
        PSComputerName   : server01

.NOTES

    Derived from the original script created by Will Anderson. 
    
    http://lastwordinnerd.com/category/posts/powershell-scripting/
    
    This script is provided AS IS without warranty of any kind.

#>


param (
    [parameter(Mandatory=True)] 
    [string[]] $ComputerName,
    [Parameter(Mandatory=$True)]
    [ValidateSet('HardwareInv','SoftwareInv','UpdateScan','MachinePol','UserPolicy','DiscoveryInv','FileCollect')]
    [string]$Action
)
Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [Parameter(Mandatory=$True)]
        [ValidateSet('HardwareInv','SoftwareInv','UpdateScan','MachinePol','UserPolicy','DiscoveryInv','FileCollect')]
        [string]$Action

)

Switch ($action) {
	'HardwareInv'  {$_action = "{00000000-0000-0000-0000-000000000001}"}
	'SoftwareInv'  {$_action = "{00000000-0000-0000-0000-000000000002}"}
	'UpdateScan'   {$_action = "{00000000-0000-0000-0000-000000000113}"}
	'MachinePol'   {$_action = "{00000000-0000-0000-0000-000000000021}"}
	'UserPolicy'   {$_action = "{00000000-0000-0000-0000-000000000027}"}
	'FileCollect'  {$_action = "{00000000-0000-0000-0000-000000000010}"}
}
Foreach ($Computer in $ComputerName) {
    if ($PSCmdlet.ShouldProcess("$action $computer")) {
        Invoke-WmiMethod -ComputerName $Computer -Namespace root\CCM -Class SMS_Client -Name TriggerSchedule -ArgumentList "$_action"
    }
}
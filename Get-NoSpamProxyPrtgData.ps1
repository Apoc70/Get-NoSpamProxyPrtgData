# copied from prtgshell.psm1

###############################################################################
# custom exe/Xml functions

function Set-PrtgResult {
    Param (
    [Parameter(mandatory=$True,Position=0)]
    [string]$Channel,
    
    [Parameter(mandatory=$True,Position=1)]
    $Value,
    
    [Parameter(mandatory=$True,Position=2)]
    [string]$Unit,

    [Parameter(mandatory=$False)]
    [alias('mw')]
    [string]$MaxWarn,

    [Parameter(mandatory=$False)]
    [alias('minw')]
    [string]$MinWarn,
    
    [Parameter(mandatory=$False)]
    [alias('me')]
    [string]$MaxError,
    
    [Parameter(mandatory=$False)]
    [alias('wm')]
    [string]$WarnMsg,
    
    [Parameter(mandatory=$False)]
    [alias('em')]
    [string]$ErrorMsg,
    
    [Parameter(mandatory=$False)]
    [alias('mo')]
    [string]$Mode,
    
    [Parameter(mandatory=$False)]
    [alias('sc')]
    [switch]$ShowChart,
    
    [Parameter(mandatory=$False)]
    [alias('ss')]
    [ValidateSet("One","Kilo","Mega","Giga","Tera","Byte","KiloByte","MegaByte","GigaByte","TeraByte","Bit","KiloBit","MegaBit","GigaBit","TeraBit")]
    [string]$SpeedSize,

	[Parameter(mandatory=$False)]
    [ValidateSet("One","Kilo","Mega","Giga","Tera","Byte","KiloByte","MegaByte","GigaByte","TeraByte","Bit","KiloBit","MegaBit","GigaBit","TeraBit")]
    [string]$VolumeSize,
    
    [Parameter(mandatory=$False)]
    [alias('dm')]
    [ValidateSet("Auto","All")]
    [string]$DecimalMode,
    
    [Parameter(mandatory=$False)]
    [alias('w')]
    [switch]$Warning,
    
    [Parameter(mandatory=$False)]
    [string]$ValueLookup
    )
    
    $StandardUnits = @("BytesBandwidth","BytesMemory","BytesDisk","Temperature","Percent","TimeResponse","TimeSeconds","Custom","Count","CPU","BytesFile","SpeedDisk","SpeedNet","TimeHours")
    $LimitMode = $false
    
    $Result  = "  <result>`n"
    $Result += "    <channel>$Channel</channel>`n"
    $Result += "    <value>$Value</value>`n"
    
    if ($StandardUnits -contains $Unit) {
        $Result += "    <unit>$Unit</unit>`n"
    } elseif ($Unit) {
        $Result += "    <unit>custom</unit>`n"
        $Result += "    <customunit>$Unit</customunit>`n"
    }
    
	if (!($Value -is [int])) { $Result += "    <float>1</float>`n" }
    if ($Mode)        { $Result += "    <mode>$Mode</mode>`n" }
    if ($MaxWarn)     { $Result += "    <limitmaxwarning>$MaxWarn</limitmaxwarning>`n"; $LimitMode = $true }
    if ($MaxError)    { $Result += "    <limitminwarning>$MinWarn</limitminwarning>`n"; $LimitMode = $true }
    if ($MaxError)    { $Result += "    <limitmaxerror>$MaxError</limitmaxerror>`n"; $LimitMode = $true }
    if ($WarnMsg)     { $Result += "    <limitwarningmsg>$WarnMsg</limitwarningmsg>`n"; $LimitMode = $true }
    if ($ErrorMsg)    { $Result += "    <limiterrormsg>$ErrorMsg</limiterrormsg>`n"; $LimitMode = $true }
    if ($LimitMode)   { $Result += "    <limitmode>1</limitmode>`n" }
    if ($SpeedSize)   { $Result += "    <speedsize>$SpeedSize</speedsize>`n" }
    if ($VolumeSize)  { $Result += "    <volumesize>$VolumeSize</volumesize>`n" }
    if ($DecimalMode) { $Result += "    <decimalmode>$DecimalMode</decimalmode>`n" }
    if ($Warning)     { $Result += "    <warning>1</warning>`n" }
    if ($ValueLookup) { $Result += "    <ValueLookup>$ValueLookup</ValueLookup>`n" }
    
    if (!($ShowChart)) { $Result += "    <showchart>0</showchart>`n" }
    
    $Result += "  </result>`n"
    
    return $Result
}

function Set-PrtgError {
	Param (
		[Parameter(Position=0)]
		[string]$PrtgErrorText
	)
	
	@"
<prtg>
  <error>1</error>
  <text>$PrtgErrorText</text>
</prtg>
"@

exit
}

# Load Module
If ( ! (Get-module NoSpamProxy )) {
    Import-Module NoSpamProxy
}

$minutes = 5
$deliveryPendingMaxWarn = 10

if(Get-Module NoSpamProxy) {
    $timespan = New-TimeSpan -Minutes $minutes

    # Gather NoSpamProxy Message Tracking information
    # -Status: Success | DispatcherError | TemporarilyBlocked | PermanentlyBlocked | PartialSuccess | DeliveryPending | Suppressed | DuplicateDrop | All
    # -Directions: FromLocal | FromExternal | All

    $inboundSuccess = (Get-NspMessageTrack -Status Success -Age $timespan -Directions FromExternal).Count
    $outboundSuccess = (Get-NspMessageTrack -Status Success -Age $timespan -Directions FromLocal).Count
    $inboundPermBlocked = (Get-NspMessageTrack -Status PermanentlyBlocked -Age $timespan -Directions FromExternal).Count
    $outboundPending = (Get-NspMessageTrack -Status DeliveryPending -Age $timespan -Directions FromLocal).Count

    # Generate PRTG Output
    $XmlOutput = "<prtg>`n"
    $XmlOutput += Set-PrtgResult -Channel 'In/Out Success' -Value $($inboundSuccess + $outboundSuccess) -Unit Count -ShowChart
    $XmlOutput += Set-PrtgResult -Channel 'Inbound Success' -Value $inboundSuccess -Unit Count -ShowChart
    $XmlOutput += Set-PrtgResult -Channel 'Outbound Success' -Value $outboundSuccess -Unit Count -ShowChart
    $XmlOutput += Set-PrtgResult -Channel 'Inbound PermanentlyBlocked' -Value $inboundPermBlocked -Unit Count -ShowChart
    if($deliveryPendingMaxWarn -ne 0) {
        $XmlOutput += Set-PrtgResult -Channel 'Inbound DeliveryPending' -Value $outboundPending -Unit Count -ShowChart -MaxWarn $deliveryPendingMaxWarn
    }
    else {
        $XmlOutput += Set-PrtgResult -Channel 'Inbound DeliveryPending' -Value $outboundPending -Unit Count -ShowChart 
    }
    $XmlOutput += "</prtg>"

    # Return Xml
    $XmlOutput
}
else {
    Set-PrtgError -PrtgErrorText 'NoSpamProxy Module failed'
}
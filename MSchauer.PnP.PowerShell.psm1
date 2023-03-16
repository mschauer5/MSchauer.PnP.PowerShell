
if (Get-Module -ListAvailable -Name 'PnP.PowerShell') {
}
else {
  Install-Module -Name PnP.PowerShell -Force -Verbose -Scope CurrentUser
}

function Get-MSchauer.PnP.PowerShell {
  (Get-Module -ListAvailable MSchauer.PnP.PowerShell).path
}

function Get-PnPListItem.ms {
	param(
    [Parameter(Mandatory=$true,Position=0)]
	  [string] $List,
    [Parameter(Mandatory=$true,Position=1)]
	  [string] $Field,
    [Parameter(Mandatory=$true,Position=2)]
	  $Value
	)
	$query = "<View><ViewFields><FieldRef Name='" + $field +"'/><FieldRef Name='ID'/></ViewFields><Query><Where><Eq><FieldRef Name='" + $field +"'/><Value Type='Text'>" + $value + "</Value></Eq></Where></Query></View>"
	$item = Get-PnPListItem -List $list -Query $query
	if ($item){
	  return $item.FieldValues.ID
	} else {
	 return 0
	}
  }

  function Remove-PnPListItem.ms {
    param(
      [Parameter(Mandatory=$true,Position=0)]
      [string] $List,
      [Parameter(Mandatory=$true,Position=1)]
      [string] $Searchfield,
      [Parameter(Mandatory=$true,Position=2)]
      $Value
    )

    $id = Get-PnPListItem.ms $List $Searchfield $Value
    if ($id -ne 0){
      Remove-PnPListItem -List $List -Identity $id -Force
    }
  }

  function Set-PnPListItem.ms
 {
  param(
    [Parameter(Mandatory=$true,Position=0)]
    [string] $List,
    [Parameter(Mandatory=$true,Position=1)]
    [string] $Searchfield,
    [Parameter(Mandatory=$true,Position=2)]
    $Values
  )

  $id = Get-PnPListItem.ms $list $searchfield $values[$searchfield]
  if ($id -eq 0){
    return  Add-PnPListItem -List $list -Values $values
  } else {
    return Set-PnPListItem -List $list -Identity $id -Values $values
  }
}

function New-PnPList.ms
 {
  param(
    [parameter(Mandatory=$true,Position=0)]
    [string] $Title,
    [parameter(Mandatory=$false)]
    [string] $DisplayName,
    [parameter(Mandatory=$true)]
    [string] $Template
  )

  $identity = $Title;
  if ($PSBoundParameters.ContainsKey('DisplayName')){
    $identity = $DisplayName
  }

  $listobj = Get-PnPList -Identity $identity

  if ($null -eq $listobj){
    New-PnPList -Title $Title -Template $Template
  }
  if ($PSBoundParameters.ContainsKey('DisplayName')){
    Set-PnPList -Identity $Title -Title $DisplayName
  }
}

function Add-PnPField.ms
 {
  param(
    [string] $List,
    [parameter(Mandatory=$true,Position=0)]
    [string] $DisplayName,
    [parameter(Mandatory=$true)]
    [string] $InternalName,
    [parameter(Mandatory=$true)]
    [string] $Type,
    [parameter(Mandatory=$false)]
    [string[]] $Choices,
    [parameter(Mandatory=$false)]
    [string] $Group,
    [Switch] $AddToDefaultView
  )

  $fld = $null;
  $isDateOnly = $false;

  $pnpField = $null;

  if ($Type -eq 'DateOnly'){
    $isDateOnly = $true;
    $Type = 'DateTime'
  }

  # Get existing field
  if ($List -eq ''){
    $fld = Get-PnPField | Where-Object {$_.InternalName -eq $InternalName}

  } else {
    $fld = Get-PnPField -List $List | Where-Object {$_.InternalName -eq $InternalName}
  }

  # Field doesn't exist
  if ($null -ne $fld){
    $pnpField =$fld;
    if ($fld.TypeAsString -ne $Type){
      
      $message = "$($InternalName) already exist with a different field type: $($fld.TypeAsString)!"
      Write-Error -Message $message
      return $pnpField;
    }

    if ($fld.Title -ne $DisplayName){
      if ($List -eq ''){
        if ($Type -eq 'Choice'){
          $pnpField = Set-PnPField -Identity $InternalName -Value @{'Title' = $DisplayName;} -Choices $Choices -Group $Group
        }
        else {
          $pnpField = Set-PnPField -Identity $InternalName -Value @{'Title' = $DisplayName;} -Group $Group
        }

      }
      else {
        if ($Type -eq 'Choice'){
          $pnpField =Set-PnPField -List $List -Identity $InternalName -Value @{'Title' = $DisplayName;} -Choices $Choices
        }
        else {
          $pnpField = Set-PnPField -List $List -Identity $InternalName -Value @{'Title' = $DisplayName;}
        }
      }
    }

    if ($true -eq $isDateOnly){
      [XML]$SchemaXml = $pnpField.SchemaXml
      $SchemaXml.Field.SetAttribute("Format","DateOnly")
      Set-PnPField -List $List -Identity $PNPField.Id -Values @{SchemaXml =$SchemaXml.OuterXml}
    }

  } else {
    if ($List -eq ''){
      if ($Type -eq 'Choice'){
        $pnpField =Add-PnPField -DisplayName $DisplayName -InternalName $InternalName -Type $Type -Choices $Choices
      } else {
        $pnpField =Add-PnPField -DisplayName $DisplayName -InternalName $InternalName -Type $Type
      }

      if ($true -eq $isDateOnly){
        [XML]$SchemaXml = $pnpField.SchemaXml
        $SchemaXml.Field.SetAttribute("Format","DateOnly")
        Set-PnPField -List $List -Identity $pnpField.Id -Values @{SchemaXml =$SchemaXml.OuterXml}
      }
    }
    else {
      if ($AddToDefaultView){
        if ($Type -eq 'Choice'){
          $pnpField =Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type -AddToDefaultView -Choices $Choices
        } else {
          $pnpField = Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type -AddToDefaultView
        }

      } else {
        if ($Type -eq 'Choice'){
          $pnpField =Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type -Choices $Choices
        } else {
          $pnpField = Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type
        }

      }

      $pnpField | format-list

      if ($true -eq $isDateOnly){
        [XML]$SchemaXml = $pnpField.SchemaXml
        $SchemaXml.Field.SetAttribute("Format","DateOnly")
        Set-PnPField -List $List -Identity $pnpField.Id -Values @{SchemaXml =$SchemaXml.OuterXml} -UpdateExistingLists
      }
    }


  }

  return $pnpField;
}


function Remove-PnPField.ms {
	param(
	  [string] $List,
    [parameter(Mandatory=$true)]
	  [string] $Identity
	)

	$fld = Get-PnPField -List $List | Where-Object {$_.InternalName -eq $Identity}

	if ($null -ne $fld){
	  Remove-PnPField -List $List -Identity $Identity -Force
	}
}

function Copy-PnPList.ms{
	param(
    [Parameter(Mandatory=$true,Position=0)]
	  [string] $SourceTitle,
    [Parameter(Mandatory=$false,Position=1)]
    [string] $SourceUrl,
    [Parameter(Mandatory=$true,Position=2)]
	  [string] $DestinationTitle,
    [Parameter(Mandatory=$false,Position=3)]
    [string] $DestinationUrl
	)

  if ($null -eq $DestinationUrl -or $DestinationUrl + '' -eq ''){
   
    $DestinationUrl = $SourceUrl
  }

  Connect-PnPOnline -Url $SourceUrl -Interactive
  $sourceCtx = Get-PnPContext

  Connect-PnPOnline -Url $DestinationUrl -Interactive
  $destinationCtx = Get-PnPConnection

  Set-PnPContext -Context $sourceCtx
  Copy-PnPList -Identity $SourceTitle -Title $DestinationTitle -Connection $destinationCtx
  
}

function Get-ListConfig{
  param(
    [Parameter(Mandatory=$true,Position=0)]
	  [string] $Value
	)
  $val = $Value.Split("|");
  return @{
    list=$val[0];
    url=$val[1];
  }
}

function Get-LocalSettings {
  param(
    [Parameter(Mandatory=$false,Position=0)]
	  [string] $Path
	)
  $config = $null; 

  $localConfig = ".\local.settings.json"

  if ($null -ne $Path){
    $localConfig = "$($Path)$($localConfig)"
  }

  $configFileExist = Test-Path $localConfig
  if ($true -eq $configFileExist)
  {
    $config  = Get-Content $localConfig | ConvertFrom-Json
  }

  return $config
}

function Set-LocalSettings {
  [Parameter(Mandatory=$true,Position=0)]
  [Object] $LocalSettings
  if ($null -ne $LocalSettings)
  {
    if ($null -ne $LocalSettings.VALUES.TenantId){
      $ENV:TenantId =$LocalSettings.VALUES.TenantId
    }

    if ($null -ne $LocalSettings.VALUES.ClientId){
      $ENV:ClientId =$LocalSettings.VALUES.ClientId
    }

    if ($null -ne $LocalSettings.VALUES.CertificateThumbPrint){
      $ENV:CertificateThumbPrint =$LocalSettings.VALUES.CertificateThumbPrint
    }
  }
}



function Connect-PnPOnline.ms{
  param(
    [Parameter(Mandatory=$true,Position=0)]
	  [string] $Url,
    [Parameter(Mandatory=$false,Position=1)]
	  [Object] $LocalSettings
	)

  if ($null -ne $LocalSettings){
    Set-LocalSettings $LocalSettings | Out-Null;
  }else {
    $workingDirectory = (Get-Item .).FullName;
    $localsettingsFile = "$($workingDirectory)\local.settings.json"
    $localsettingsFile
    $localsettingsFile_exists = Test-Path -Path $localsettingsFile -PathType Leaf
    if ($true -eq $localsettingsFile_exists){
      $LocalSettings = Get-Content -Raw -Path  $localsettingsFile | ConvertFrom-Json;
      Set-LocalSettings $LocalSettings | Out-Null;
    }
    
  }
 

  if ($null -ne $ENV:ClientId -and $null -ne $ENV:TenantId -and $null -ne $ENV:CertificateThumbPrint){
    Connect-PnPOnline -Url $Url -Tenant $ENV:TenantId -ClientId $ENV:ClientId -Thumbprint $ENV:CertificateThumbPrint
  } elseif ($ENV:Interactive){
    Connect-PnPOnline -Url $Url -Interactive
  } else {
    Connect-PnPOnline -Url $Url -ManagedIdentity
  }

 $ctx = @{
  PnPConnection = Get-PnPConnection;
  PnPContext = Get-PnPContext;
 };

 return $ctx;
}

Export-ModuleMember -Function Get-PnPListItem.ms
Export-ModuleMember -Function Remove-PnPListItem.ms
Export-ModuleMember -Function Set-PnPListItem.ms
Export-ModuleMember -Function New-PnPList.ms
Export-ModuleMember -Function Add-PnPField.ms
Export-ModuleMember -Function Remove-PnPField.ms
Export-ModuleMember -Function Copy-PnPList.ms
Export-ModuleMember -Function Get-MSchauer.PnP.PowerShell
Export-ModuleMember -Function Connect-PnPOnline.ms
Export-ModuleMember -Function Get-ListConfig
Export-ModuleMember -Function Set-LocalSettings
Export-ModuleMember -Function Get-LocalSettings


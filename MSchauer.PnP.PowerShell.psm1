
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

  if ($List -eq ''){
    $fld = Get-PnPField | Where-Object {$_.InternalName -eq $InternalName}

  } else {
    $fld = Get-PnPField -List $List | Where-Object {$_.InternalName -eq $InternalName}
  }

  if ($null -ne $fld){

    if ($fld.TypeAsString -ne $Type){
      $message = $InternalName + ' already exist with a different field type!'
      Write-Error -Message $message
      return;
    }

    if ($fld.Title -ne $DisplayName){
      if ($List -eq ''){
        if ($Type -eq 'Choice'){
          Set-PnPField -Identity $InternalName -Value @{'Title' = $DisplayName;} -Choices $Choices -Group $Group
        }
        else {
          Set-PnPField -Identity $InternalName -Value @{'Title' = $DisplayName;} -Group $Group
        }

      }
      else {
        if ($Type -eq 'Choice'){
        Set-PnPField -List $List -Identity $InternalName -Value @{'Title' = $DisplayName;} -Choices $Choices
        }
        else {
          Set-PnPField -List $List -Identity $InternalName -Value @{'Title' = $DisplayName;}
        }
      }
    }

  } else {
    if ($List -eq ''){
      if ($Type -eq 'Choice'){
        Add-PnPField -DisplayName $DisplayName -InternalName $InternalName -Type $Type -Choices $Choices
      } else {
        Add-PnPField -DisplayName $DisplayName -InternalName $InternalName -Type $Type
      }
    }
    else {
      if ($AddToDefaultView){
        if ($Type -eq 'Choice'){
          Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type -AddToDefaultView -Choices $Choices
        } else {
          Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type -AddToDefaultView
        }

      } else {
        if ($Type -eq 'Choice'){
          Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type -Choices $Choices
        } else {
          Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type
        }

      }
    }
  }

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

function Connect-PnPOnline.ms{
  param(
    [Parameter(Mandatory=$true,Position=0)]
	  [string] $Url
	)

  if ($null -ne $ENV:ClientId -and $null -ne $ENV:TenantId -and $null -ne $ENV:CertificateThumbPrint){
    Connect-PnPOnline -Url $Url -Tenant $ENV:TenantId -ClientId $ENV:ClientId -Thumbprint $ENV:CertificateThumbPrint
  } elseif ($ENV:Interactive){
    Connect-PnPOnline -Url $Url -Interactive
  } else {
    Connect-PnPOnline -Url $Url -ManagedIdentity
  }

 $ctx = @{};
 $ctx.PnPConnection = Get-PnPConnection;
 $ctx.PnPContext = Get-PnPContext;
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


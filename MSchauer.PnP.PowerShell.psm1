
if (Get-Module -ListAvailable -Name 'PnP.PowerShell') {
}
else {
  Install-Module -Name PnP.PowerShell -Force -Verbose -Scope CurrentUser
}

function Get-PnPListItem.ms {
	param(
    [Parameter(Mandatory=$true,Position=0)]
	  [string] $List,
    [Parameter(Mandatory=$true)]
	  [string] $Field,
    [Parameter(Mandatory=$true)]
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
      [Parameter(Mandatory=$true)]
      [string] $Searchfield,
      [Parameter(Mandatory=$true)]
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
    [Parameter(Mandatory=$true)]
    [string] $Searchfield,
    [Parameter(Mandatory=$true)]
    $Values
  )

  $id = Get-PnPListItem.ms $list $searchfield $values[$searchfield] $values
  if ($id -eq 0){
      Add-PnPListItem -List $list -Values $values
  } else {
      Set-PnPListItem -List $list -Identity $id -Values $values
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

function Add-PnPField.xx
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
    [Switch] $AddToDefaultView
  )

  Write-Host $Choices
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
        Set-PnPField -Identity $InternalName -Value @{'Title' = $DisplayName;} -Choices $Choices
      }
      else {
        Set-PnPField -List $List -Identity $InternalName -Value @{'Title' = $DisplayName;} -Choices $Choices
      }
    }

  } else {
    if ($List -eq ''){
      Add-PnPField -DisplayName $DisplayName -InternalName $InternalName -Type $Type -Choices $Choices
    }
    else {
      if ($AddToDefaultView){
        Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type -AddToDefaultView -Choices $Choices
      } else {
        Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type -Choices $Choices
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

Export-ModuleMember -Function Get-PnPListItem.ms
Export-ModuleMember -Function Remove-PnPListItem.ms
Export-ModuleMember -Function Set-PnPListItem.ms
Export-ModuleMember -Function New-PnPList.ms
Export-ModuleMember -Function Add-PnPField.ms
Export-ModuleMember -Function Remove-PnPField.ms
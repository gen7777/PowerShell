#requires -version 2
<#
   .SYNOPSIS
        Приводит значения ветки реестра
        HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\*\Count
        к удобочитаемому виду
#>
(gp HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\*\Count) | % {
  $rot13 = New-Object "Collections.Generic.Dictionary[Char, Char]"
  $table = {param([Int32[]]$arr, [Int32]$lim)
    $arr | % {
      $$ = if ($_ -le $lim) {$_ + 13} else {$_ - 13}
      $rot13.Add([Char]$_, [Char]$$)
    }
  }
  $table.Invoke(65..90, 77)
  $table.Invoke(97..122, 109)
  "Session Counter      LastLaunchTime Name`n$('-'*7) $('-'*7)      $('-'*14) $('-'*4)"
}{
  $_.PSObject.Properties | ? {$_.Name -notlike 'PS*'} | % {
    $name = -join ($_.Name.ToCharArray() | % {if($rot13.ContainsKey($_)){$rot13[$_]}else{$_}})
    if ($_.Value.Length -eq 16) {
      $time = [BitConverter]::ToInt64($_.Value[8..15], 0)
      '{0,7}{1,8}{2,20} {3}' -f [BitConverter]::ToUInt32($_.Value[0..3], 0), [BitConverter]::ToUInt32(
          $_.Value[4..7], 0
      ), $(if($time -ne 0){[DateTime]::FromFileTime($time)}else{''}), $name
    }
    else {
      '{0,7}{1,8}{2,20} {3}' -f [BitConverter]::ToUInt32($_.Value[4..7], 0), '', '', $name
    }
  }
}

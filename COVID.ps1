get-childitem "$env:Public\Desktop" | remove-item -Recurse -Force -Confirm:$false
get-childitem "C:\Users\Default\Desktop" | remove-item -Recurse -Force -Confirm:$false
get-childitem "C:\Windows\System32\GroupPolicy\DataStore\0\SysVol" | remove-item -Recurse -Force -Confirm:$false




function Gen-Shortcut($TargetFile,$ShortcutFile) {
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    $Shortcut.Save()
}

Gen-Shortcut -TargetFile "C:\Program Files\Palo Alto Networks\GlobalProtect\PanGPA.exe" -ShortcutFile "$env:Public\Desktop\GlobalProtect.lnk"

New-Item -Path "$env:Public\Desktop" -ItemType "directory" -Name "Documentation"

Gen-Shortcut -TargetFile "https://docs.google.com/document/d/1pR6VmeX_yWe_CQwicplGKJ1K0qcjLgG9XUMsH_az4PA" -ShortcutFile "$env:Public\Desktop\Documentation\Duo Authentication Guide.url"
Gen-Shortcut -TargetFile "https://docs.google.com/document/d/1CQS1Yo26Shoe41j7WWxWL7y-rq_JJ6Uinrk_UCn9OVI" -ShortcutFile "$env:Public\Desktop\Documentation\VPN Sign-in Guide.url"
Gen-Shortcut -TargetFile "https://docs.google.com/document/d/1-fYe000q0gSQJhzfJNN3IjBPOfKML7_u0P132WXp3Wc" -ShortcutFile "$env:Public\Desktop\Documentation\Remote Desktop Guide.url"
Gen-Shortcut -TargetFile "https://docs.google.com/document/d/1VirRFwOBw1yMItIslMpo9YmAde8wL2yMup5K0nksaFg" -ShortcutFile "$env:Public\Desktop\Documentation\File Share Mapping Guide.url"
Gen-Shortcut -TargetFile "https://docs.google.com/document/d/1uI_R5oVAx-JszyrVYbtAuAroWxOaK3V-ULxrQWh41gE" -ShortcutFile "$env:Public\Desktop\Documentation\VPN Ticket Request Guide.url"

New-ItemProperty -Path "HKLM:\SOFTWARE\Palo Alto Networks\GlobalProtect\PanSetup" -Name Portal -PropertyType String -Value "pcc-vpn.pima.edu" -Force | Out-Null


gpupdate /force

shutdown -r -t 30
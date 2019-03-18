## The purpose of this script is to connect to your Exchange On-Premises server via Powershell, import commands to your localhost and perform a variety
## of tasks without having to open a console or go through a GUI.

while ($true) {

## Captures Creds

$creds = Get-Credential

try {

## Tries to create session. On any error it kicks back to cred entry

$sesh = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri YOUR_UNIQUE_URI -Authentication PREFERRED_AUTH_METHOD -Credential $creds 

Import-PSSession $sesh -DisableNameChecking -AllowClobber -ErrorAction Stop

break

} catch {

write-host -BackgroundColor Red -ForegroundColor white "Uh oh...looks like the credentials you provided are not correct or there was another error. Please try again!"

}
}

Write-Host "You're in....(queue movie hacker scene)"

## The next section provides a GUI box for quick disconnect and removal of session afterwards

Add-Type -AssemblyName system.windows.forms

$discbox = [System.Windows.MessageBox]::Show('Click Ok to disconnect your session when complete','Input','Ok')

 switch ($discbox) {

  'Ok' {

   remove-pssession $sesh

   Write-Host "All done and squeaky clean!!"

   }
   }

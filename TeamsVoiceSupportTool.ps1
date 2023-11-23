#region Description
<#     
       .NOTES
       ==============================================================================
       Created on:         2023/11/23 
       Created by:         Drago Petrovic
       Organization:       MSB365.blog
       Filename:           TeamsVoiceSupportTool.ps1
       Current version:    V1.00     

       Find us on:
             * Website:         https://www.msb365.blog
             * Technet:         https://social.technet.microsoft.com/Profile/MSB365
             * LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
             * MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
       ==============================================================================

       .DESCRIPTION
       This script can be executed without prior customisation.
       This script is used to assign PhoneNumbers with PowerShell            
       

       .NOTES
       It is manditory to have the right Licenses already assigned to your Tenant.





       .EXAMPLE
       .\TeamsVoiceSupportTool.ps1 
             

       .COPYRIGHT
       Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
       to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
       and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
       WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
       ===========================================================================
       .CHANGE LOG
             V1.00, 2023/04/01 - DrPe - Initial version
             V2.00, 2023/07/11 - DrPe - Added the Option - Set Voice Routing

             
			 




--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "TeamsVoiceSupportTool V1.0"
$RKEY = "MSB365_TeamsVoiceSupportTool_V10"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2023 Drago Petrovic\par
$Scriptname \par
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################

function Show-CustomMenu
{
    param (
        [string]$menuname = 'Microsoft Teams Voice Support Tool'
    )

    Clear-Host

    ###############################################################################
write-host " __  __ _                           __ _     _______                        __      __   _           " -ForegroundColor Magenta
write-host "|  \/  (_)                         / _| |   |__   __|                       \ \    / /  (_)          " -ForegroundColor Magenta
write-host "| \  / |_  ___ _ __ ___  ___  ___ | |_| |_     | | ___  __ _ _ __ ___  ___   \ \  / /__  _  ___ ___  " -ForegroundColor Magenta
write-host "| |\/| | |/ __| '__/ _ \/ __|/ _ \|  _| __|    | |/ _ \/ _`  | '_ ` _ \/ __|  \ \/ / _ \| |/ __/ _ \ " -ForegroundColor Magenta
write-host "| |  | | | (__| | | (_) \__ \ (_) | | | |_     | |  __/ (_| | | | | | \__ \    \  / (_) | | (_|  __/ " -ForegroundColor Magenta
write-host "|_|  |_|_|\___|_|  \___/|___/\___/|_|  \__|    |_|\___|\__,_|_| |_| |_|___/     \/ \___/|_|\___\___| " -ForegroundColor Magenta
Start-Sleep -s 2
write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
Start-Sleep -s 2
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
write-host ""
    ###############################################################################

    Write-Host "================ $menuname ================" -ForegroundColor Magenta
    
    Write-Host "1: Choose '1' for Direct Routing"
    Write-Host "2: Choose '2' for Operator Connect"
    Write-Host "3: Choose '3' for setting the call routing policy."
    Write-Host "4: Choose '4' for setting bulk Teams Policies (for Users)"
    Write-Host "Q: Choose 'Q' to exit the Module." -ForegroundColor Red
    Write-host ""
}


# Menue aufrufen und Titel uebergeben
Show-CustomMenu –menuname 'Microsoft Teams Phone Number Assignment'

# Eingabe /Auswahl des Benutzers
#$auswahl = Read-Host "Please enter your choice" 
$auswahl = $(write-host "Please enter your choice:" -ForegroundColor Yellow -BackgroundColor Black; Read-Host)

# Optionen wählen
switch ($auswahl){
     '1' {# Get Location ID
			write-host "Gettering Tenant location ID..." -ForegroundColor Cyan
			start-sleep -s 2
			$Lid = Get-CsOnlineLisLocation | Sort-Object LocationID | select-object -ExpandProperty LocationID
			write-host "Tenant LocationID is: $Lid" -ForegroundColor White -BackgroundColor Black
			Start-Sleep -s 2

			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"UserPrincipalName","DisplayName","TelephoneNumber","CallingPolicy"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3

			# Configuring Phone Number for Teams users
			write-host "Setting the Phone numbers..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					#Set-CsOnlineVoiceUser -Identity $user.UserPrincipalName -TelephoneNumber $user.TelephoneNumber -LocationID $Lid -ErrorAction Stop ###
                    Set-CsPhoneNumberAssignment -Identity $user.UserPrincipalName -PhoneNumber $user.TelephoneNumber -PhoneNumberType DirectRouting 
					Write-Host "Phone numbers for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
                    Set-CsPhoneNumberAssignment -Identity $user.UserPrincipalName -EnterpriseVoiceEnabled $true
                    Write-Host "Enterprise Voice for the users $($user.DisplayName) enabled." -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set phone number for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All phone numbers set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			# Get updated overview
			Write-Host "Preparing showing all updated user list..." -ForegroundColor Cyan
			Start-Sleep -s 5
			Get-CsOnlineUser | ft UserPrincipalName, LineURI
			pause}
     '2' {# Get Location ID
			write-host "Gettering Tenant location ID..." -ForegroundColor Cyan
			start-sleep -s 2
			$Lid = Get-CsOnlineLisLocation | Sort-Object LocationID | select-object -ExpandProperty LocationID
			write-host "Tenant LocationID is: $Lid" -ForegroundColor White -BackgroundColor Black
			Start-Sleep -s 2

			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"UserPrincipalName","DisplayName","TelephoneNumber","CallingPolicy"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3

			# Configuring Phone Number for Teams users
			write-host "Setting the Phone numbers..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					#Set-CsOnlineVoiceUser -Identity $user.UserPrincipalName -TelephoneNumber $user.TelephoneNumber -LocationID $Lid -ErrorAction Stop ###
                    Set-CsPhoneNumberAssignment -Identity $user.UserPrincipalName -PhoneNumber $user.TelephoneNumber -PhoneNumberType OperatorConnect 
					Write-Host "Phone numbers for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
                    Set-CsPhoneNumberAssignment -Identity $user.UserPrincipalName -EnterpriseVoiceEnabled $true
                    Write-Host "Enterprise Voice for the users $($user.DisplayName) enabled." -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set phone number for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All phone numbers set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			# Get updated overview
			Write-Host "Preparing showing all updated user list..." -ForegroundColor Cyan
			Start-Sleep -s 5
			Get-CsOnlineUser | ft UserPrincipalName, LineURI
			pause}
     '3' {# Get Location ID
	 start-sleep -s 1

write-host "  #####                                                                                                                                     " -ForegroundColor Magenta
write-host " #     # ###### #####    #    #  ####  #  ####  ######    #####   ####  #    # ##### # #    #  ####     #####   ####  #      #  ####  #   # " -ForegroundColor Magenta
write-host " #       #        #      #    # #    # # #    # #         #    # #    # #    #   #   # ##   # #    #    #    # #    # #      # #    #  # #  " -ForegroundColor Magenta
write-host "  #####  #####    #      #    # #    # # #      #####     #    # #    # #    #   #   # # #  # #         #    # #    # #      # #        #   " -ForegroundColor Magenta
write-host "       # #        #      #    # #    # # #      #         #####  #    # #    #   #   # #  # # #  ###    #####  #    # #      # #        #   " -ForegroundColor Magenta
write-host " #     # #        #       #  #  #    # # #    # #         #   #  #    # #    #   #   # #   ## #    #    #      #    # #      # #    #   #   " -ForegroundColor Magenta
write-host "  #####  ######   #        ##    ####  #  ####  ######    #    #  ####   ####    #   # #    #  ####     #       ####  ###### #  ####    #   " -ForegroundColor Magenta                                                                                     

	 
	 
			write-host "Gettering Tenant location ID..." -ForegroundColor Cyan
			start-sleep -s 2
			$Lid = Get-CsOnlineLisLocation | Sort-Object LocationID | select-object -ExpandProperty LocationID
			write-host "Tenant LocationID is: $Lid" -ForegroundColor White -BackgroundColor Black
			Start-Sleep -s 2

			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"UserPrincipalName","DisplayName","TelephoneNumber","CallingPolicy"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3

			# Configuring Phone Number for Teams users
			write-host "Setting the policies..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					Grant-CsOnlineVoiceRoutingPolicy -Identity $user.UserPrincipalName -PolicyName $user.CallingPolicy
                    Write-Host "Voice Routing Policy: $($user.CallingPolicy) for the users $($user.DisplayName) is set!" -ForegroundColor Green
					Start-Sleep -s 1
                    
				}
				catch
				{
					Write-Host "Could not set the Policy $($user.CallingPolicy) for the user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "Policy for all users set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5}
	'4' {cls
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “  #     #                         ######                                         " -ForegroundColor Magenta
Write-Host “  #     #  ####  ###### #####     #     #  ####  #      #  ####  # ######  ####  " -ForegroundColor Magenta
Write-Host “  #     # #      #      #    #    #     # #    # #      # #    # # #      #      " -ForegroundColor Magenta
Write-Host “  #     #  ####  #####  #    #    ######  #    # #      # #      # #####   ####  " -ForegroundColor Magenta
Write-Host “  #     #      # #      #####     #       #    # #      # #      # #           # " -ForegroundColor Magenta
Write-Host “  #     # #    # #      #   #     #       #    # #      # #    # # #      #    # " -ForegroundColor Magenta
Write-Host “   #####   ####  ###### #    #    #        ####  ###### #  ####  # ######  ####  " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Start-sleep -s 1
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host "*****************************************************************************************" -ForegroundColor Yellow -BackgroundColor Black
Write-Host “NOTE! " -ForegroundColor Yellow -BackgroundColor black
Write-Host “     Some of the policy sets requires at least the Teams PowerShell module version 5.7.1" -ForegroundColor Yellow -BackgroundColor black
Write-Host "*****************************************************************************************" -ForegroundColor Yellow -BackgroundColor Black
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Write-Host “ " -ForegroundColor Magenta
Start-sleep -s 4
#########################################
# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"UserPrincipalName","TeamsPolicy"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
Write-Host “ "
Write-Host “ "
Write-Host “ "
Write-Host “ "
Write-Host “Select the Policy Type, you want to assign... " -ForegroundColor Cyan
Write-Host “ "
Write-Host “ "
Write-Host “ "
Write-Host “ "
###############################################################################

function Show-CustomMenu
{
    param (
        [string]$menuname = 'Microsoft Teams User Policy Assignment'
    )

    #Clear-Host

    ###############################################################################
    
Write-Host "================ Microsoft Teams User Policy Assignment ================" -ForegroundColor Magenta

Write-Host “ 1: Press ‘ 1’ for BULK TeamsAudioConferencingPolicy"
Write-Host “ 2: Press ‘ 2’ for BULK TeamsCallHoldPolicy"
Write-Host “ 3: Press ‘ 3’ for BULK TeamsCallParkPolicy"
Write-Host “ 4: Press ‘ 4’ for BULK TeamsChannelsPolicy"
Write-Host “ 5: Press ‘ 5’ for BULK TeamsComplianceRecordingPolicy"
Write-Host “ 6: Press ‘ 6’ for BULK TeamsCortanaPolicy"
Write-Host “ 7: Press ‘ 7’ for BULK TeamsEmergencyCallingPolicy"
Write-Host “ 8: Press ‘ 8’ for BULK TeamsEmergencyCallRoutingPolicy"
Write-Host “ 9: Press ‘ 9’ for BULK TeamsEnhancedEncryptionPolicy"
Write-Host “10: Press ‘10’ for BULK TeamsFeedbackPolicy"
Write-Host “11: Press ‘11’ for BULK TeamsFilesPolicy"
Write-Host “12: Press ‘12’ for BULK TeamsIPPhonePolicy"
Write-Host “13: Press ‘13’ for BULK enabling TeamsMediaLoggingPolicy" -ForegroundColor Cyan
Write-Host “14: Press ‘14’ for BULK TeamsMeetingBroadcastPolicy"
Write-Host “15: Press ‘15’ for BULK TeamsMeetingPolicy"
Write-Host “16: Press ‘16’ for BULK TeamsMessagingPolicy"
Write-Host “17: Press ‘17’ for BULK TeamsMobilityPolicy"
Write-Host “18: Press ‘18’ for BULK TeamsRoomVideoTeleConferencingPolicy (Device Policy)" -ForegroundColor Cyan
Write-Host “19: Press ‘19’ for BULK TeamsSurvivableBranchAppliancePolicy"
Write-Host “20: Press ‘20’ for BULK TeamsUpdateManagementPolicy"
Write-Host “21: Press ‘21’ for BULK TeamsUpgradePolicy"
Write-Host “22: Press ‘22’ for BULK TeamsVideoInteropServicePolicy"
Write-Host “23: Press ‘23’ for BULK TeamsVoiceApplicationsPolicy"
Write-Host “24: Press ‘24’ for BULK TeamsWorkLoadPolicy"
Write-Host “25: Press ‘25’ for BULK TeamsAppPermissionPolicy"
Write-Host “26: Press ‘26’ for BULK TeamsAppSetupPolicy"
Write-Host “27: Press ‘27’ for BULK TeamsCallingPolicy"
Write-Host “28: Press ‘28’ for BULK TeamsEventsPolicy"
Write-Host “29: Press ‘29’ for BULK TeamsMeetingBrandingPolicy"
Write-Host “30: Press ‘30’ for BULK TeamsMeetingTemplatePermissionPolicy"
Write-Host “31: Press ‘31’ for BULK TeamsSharedCallingRoutingPolicy"
Write-Host “32: Press ‘32’ for BULK TeamsShiftsPolicy"
Write-Host “33: Press ‘33’ for BULK TeamsVdiPolicy"
Write-Host “34: Press ‘34’ for BULK TeamsVirtualAppointmentsPolicy"
Write-Host “Q: Press ‘Q’ to quit.” -ForegroundColor Red
Write-host ""
}


# Menue aufrufen und Titel uebergeben
Show-CustomMenu –menuname $menuname

# Eingabe /Auswahl des Benutzers
#$auswahl = Read-Host "Please enter your choice" 
$auswahl = $(write-host "Please enter your choice:" -ForegroundColor Yellow -BackgroundColor Black; Read-Host)

# Optionen wählen
switch ($auswahl){
     '1' {          			
			write-host "Setting the TeamsAudioConferencingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsAudioCOnferencingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '2' {          			
			write-host "Setting the TeamsCallHoldPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsCallHoldPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '3' {          			
			write-host "Setting the TeamsCallParkPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsCallParkPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '4' {          			
			write-host "Setting the TeamsChannelsPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsChannelsPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '5' {          			
			write-host "Setting the TeamsComplianceRecordingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsComplianceRecordingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '6' {          			
			write-host "Setting the TeamsCortanaPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsCortanaPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '7' {          			
			write-host "Setting the TeamsEmergencyCallingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsEmergencyCallingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '8' {          			
			write-host "Setting the TeamsEmergencyCallRoutingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '9' {          			
			write-host "Setting the TeamsEnhancedEncryptionPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsEnhancedEncryptionPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '10' {          			
			write-host "Setting the TeamsFeedbackPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsFeedbackPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '11' {          			
			write-host "Setting the TeamsFilesPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsFilesPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '12' {          			
			write-host "Setting the TeamsIPPhonePolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsIPPhonePolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '13' {          			
			write-host "Setting the TeamsMediaLoggingPolicy to enabled..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsMediaLoggingPolicy -Identity $user.UserPrincipalName -PolicyName Enabled 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '14' {          			
			write-host "Setting the MeetingBroadcastPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsMeetingBroadcastPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '15' {          			
			write-host "Setting the TeamsMeetingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsMeetingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '16' {          			
			write-host "Setting the TeamsMessagingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsMessagingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '17' {          			
			write-host "Setting the TeamsMobilityPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsMobilityPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '18' {          			
			write-host "Setting the TeamsRoomVideoTeleConferencingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsRoomVideoTeleConferencingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '19' {          			
			write-host "Setting the TeamsSurvivableBranchAppliancePolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsSurvivableBranchAppliancePolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '20' {          			
			write-host "Setting the TeamsUpdateManagementPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsUpdateManagementPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '21' {          			
			write-host "Setting the TeamsUpgradePolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsUpgradePolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '22' {          			
			write-host "Setting the TeamsVideoInteropServicePolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsVideoInteropServicePolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '23' {          			
			write-host "Setting the TeamsVoiceApplicationsPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsVoiceApplicationsPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '24' {          			
			write-host "Setting the TeamsWorkLoadPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsWorkLoadPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '25' {          			
			write-host "Setting the TeamsAppPermissionPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsAppPermissionPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '26' {          			
			write-host "Setting the TeamsAppSetupPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsAppSetupPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '27' {          			
			write-host "Setting the TeamsCallingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsCallingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '28' {          			
			write-host "Setting the TeamsEventsPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsEventsPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '29' {          			
			write-host "Setting the TeamsMeetingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsMeetingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '30' {          			
			write-host "Setting the TeamsMeetingTemplatePermissionPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsMeetingTemplatePermissionPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '31' {          			
			write-host "Setting the TeamsSharedCallingRoutingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsSharedCallingRoutingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '32' {          			
			write-host "Setting the TeamsShiftsPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsShiftsPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '33' {          			
			write-host "Setting the TeamsVdiPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsVdiPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
            Write-Host "More information about this policy settings can be found on https://learn.microsoft.com/en-us/microsoftteams/teams-for-vdi" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 5

			pause}
     '34' {          			
			write-host "Setting the TeamsVirtualAppointmentsPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
					                    Grant-CsTeamsVirtualAppointmentsPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsPolicy 
					Write-Host "Policy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set Policy for user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All policies set!" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s 5

			pause}



####################


     'Q' { #Quit the script
           Write-Host "Thank you for using the Script: $Scriptname" -ForegroundColor Yellow -BackgroundColor Black
           Start-Sleep -s 2
           Write-Host "More Scripts from MSB365 can be found on https://github.com/MSB365/" -ForegroundColor Yellow -BackgroundColor Black
           Start-Sleep -s 5
           Write-Host "The script will be closed in 5 seconds..." -ForegroundColor Yellow -BackgroundColor Black
		Start-Sleep -s 5
     Exit
     }
     }

}
     
     'Q' { #Quit the script
           Write-Host "Thank you for using the Script: $Scriptname" -ForegroundColor Yellow -BackgroundColor Black
           Start-Sleep -s 2
           Write-Host "More Scripts from MSB365 can be found on https://github.com/MSB365/" -ForegroundColor Yellow -BackgroundColor Black
           Start-Sleep -s 5
           Write-Host "The script will be closed in 5 seconds..." -ForegroundColor Yellow -BackgroundColor Black
		Start-Sleep -s 5
     Exit
     }
     }

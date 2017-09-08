<#

.Requires -version 2 - Runs in Exchange Management Shell

.SYNOPSIS
.\DistributionGroupMemberReport.ps1 - It Can Display all the Distribution Group and its members on a List

Or It can Export to a CSV file
Or It can delete DistributionGroups


Example 1

[PS] C:\DG>.\DistributionGroupMemberReport.ps1


Distribution Group Member Report
----------------------------

1.Display in Shell

2.Export to CSV File

Choose The Task: 1

DisplayName                   Alias                         Primary SMTP address          Distriubtion Group
-----------                   -----                         --------------------          ------------------
Atlast1                       Atlast1                       Atlast1@targetexchange.in     Test1
Atlast2                       Atlast2                       Atlast2@careexchange.in       Test1
Blink                         Blink                         Blink@targetexchange.in       Test1
blink1                        blink1                        blink1@targetexchange.in      Test1
User2                         User2                         User2@careexchange.in         Test11
User3                         User3                         User3@careexchange.in         Test11
User4                         User4                         User4@careexchange.in         Test11
WithClient                    WithClient                    WithClient@careexchange.in    Test11
Blink                         Blink                         Blink@targetexchange.in       Test11
blink1                        blink1                        blink1@targetexchange.in      Test11

Example 2

[PS] C:\DG>.\DistributionGroupMemberReport.ps1


Distribution Group Member Report
----------------------------

1.Display in Shell

2.Export to CSV File

Choose The Task: 2
Enter the Path of CSV file (Eg. C:\DG.csv): C:\DGmembers.csv



.Author
Written By: Satheshwaran Manoharan <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Modified By: Drago Petrovic (https://msb365.abstergo.ch)

Change Log
V1.0, 11/10/2012 - Initial version

Change Log
V1.1, 02/07/2014 - Added "Enter the Distribution Group name with Wild Card"

Change Log
V1.2, 19/07/2014 - Added "Recipient OU,Distribution Group Primary SMTP address,Distribution Group Managers,Distribution Group OU"
V1.2.1, 19/07/2014 - Added "Option- Enter the Distribution Group name with Wild Card (Display)"
V1.2.2, 19/07/2014 - Added "Fixed "Hashtable-to-Object conversion is not supported in restricted language mode or a Data section"
V1.3,05/08/2014 - Hashtable-to-Object conversion is not supported - Fixed 
V1.4,30/08/2015 - 
Removed For loops - As its not listing distribution groups which has one member.
Added Value for Empty groups. It will list empty groups now as well.
V1.5,09/09/2015 - Progress Bars while exporting to CSV

Change Log - by MSB365
V2.0,08/09/2017 - Added new option <9> for deleting DistributionGroup with zero Members


--- keep it simple, but significant ---

Find me on:

* LinkedIn:	https://www.linkedin.com/in/drago-petrovic/
* Xing:     https://www.xing.com/profile/Drago_Petrovic
* Website:  https://msb365.abstergo.ch
* GitHub:   https://github.com/MSB365
* Technet:	https://social.technet.microsoft.com/Profile/MSB365


.COPYRIGHT
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

Write-host "

Distribution Group Member Report
----------------------------

1.Display in Exchange Management Shell

2.Export to CSV File

3.Enter the Distribution Group name with Wild Card (Export)

4.Enter the Distribution Group name with Wild Card (Display)

Dynamic Distribution Group Member Report
----------------------------

5.Display in Exchange Management Shell

6.Export to CSV File

7.Enter the Dynamic Distribution Group name with Wild Card (Export)

8.Enter the Dynamic Group name with Wild Card (Display)
----------------------------" -ForeGround "Cyan"

Write-host "
----------------------------
9.Delete all Distribution Groups with NO (zero) Members
----------------------------" -ForeGround "Red"

#----------------
# Script
#----------------

Write-Host "               "

$number = Read-Host "Choose The Task"
$output = @()
switch ($number)
{
	
	1 {
		
		$AllDG = Get-DistributionGroup -resultsize unlimited
		Foreach ($dg in $allDg)
		{
			$Members = Get-DistributionGroupMember $Dg.name -resultsize unlimited
			
			
			if ($members.count -eq 0)
			{
				$userObj = New-Object PSObject
				$userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Alias" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
				Write-Output $Userobj
			}
			else
			{
				Foreach ($Member in $members)
				{
					$userObj = New-Object PSObject
					$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $member.Name
					$userObj | Add-Member NoteProperty -Name "Alias" -Value $member.Alias
					$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $member.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
					Write-Output $Userobj
				}
				
			}
			
		}
		
		; Break
	}
	
	2 {
		
		$i = 0
		
		$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\DG.csv)"
		
		$AllDG = Get-DistributionGroup -resultsize unlimited
		
		Foreach ($dg in $allDg)
		{
			$Members = Get-DistributionGroupMember $Dg.name -resultsize unlimited
			
			if ($members.count -eq 0)
			{
				$managers = $Dg | Select @{ Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
				
				$userObj = New-Object PSObject
				
				
				$userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Alias" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "RecipientType" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
				$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
				$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
				$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
				$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
				$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
				#
				$userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $DG.WhenCreated
				$userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $DG.WhenChanged
				
				
				$output += $UserObj
				
			}
			else
			{
				Foreach ($Member in $members)
				{
					
					$managers = $Dg | Select @{ Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
					
					$userObj = New-Object PSObject
					
					$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $Member.Name
					$userObj | Add-Member NoteProperty -Name "Alias" -Value $Member.Alias
					$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Member.RecipientType
					$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Member.OrganizationalUnit
					$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Member.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
					$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
					$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
					$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
					$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
					#
					$userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $DG.WhenCreated
					$userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $DG.WhenChanged
					
					
					$output += $UserObj
					
				}
			}
			# update counters and write progress
			$i++
			Write-Progress -activity "Scanning Groups . . ." -status "Scanned: $i of $($allDg.Count)" -percentComplete (($i / $allDg.Count) * 100)
			$output | Export-csv -Path $CSVfile -NoTypeInformation
			
		}
		
		; Break
	}
	
	3 {
		
		$i = 0
		
		$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\DG.csv)"
		
		$Dgname = Read-Host "Enter the DG name or Range (Eg. DGname , DG*,*DG)"
		
		$AllDG = Get-DistributionGroup $Dgname -resultsize unlimited
		
		Foreach ($dg in $allDg)
		{
			
			$Members = Get-DistributionGroupMember $Dg.name -resultsize unlimited
			
			if ($members.count -eq 0)
			{
				$managers = $Dg | Select @{ Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
				
				$userObj = New-Object PSObject
				
				
				$userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Alias" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "RecipientType" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
				$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
				$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
				$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
				$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
				$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
				#
				$userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $DG.WhenCreated
				$userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $DG.WhenChanged
				
				
				$output += $UserObj
				
			}
			else
			{
				Foreach ($Member in $members)
				{
					
					$managers = $Dg | Select @{ Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
					
					$userObj = New-Object PSObject
					
					$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $Member.Name
					$userObj | Add-Member NoteProperty -Name "Alias" -Value $Member.Alias
					$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Member.RecipientType
					$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Member.OrganizationalUnit
					$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Member.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
					$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
					$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
					$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
					$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
					#
					$userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $DG.WhenCreated
					$userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $DG.WhenChanged
					
					
					$output += $UserObj
					
				}
			}
			# update counters and write progress
			$i++
			Write-Progress -activity "Scanning Groups . . ." -status "Scanned: $i of $($allDg.Count)" -percentComplete (($i / $allDg.Count) * 100)
			$output | Export-csv -Path $CSVfile -NoTypeInformation
			
		}
		
		; Break
	}
	
	4 {
		
		$Dgname = Read-Host "Enter the DG name or Range (Eg. DGname , DG*,*DG)"
		
		$AllDG = Get-DistributionGroup $Dgname -resultsize unlimited
		
		Foreach ($dg in $allDg)
		{
			
			$Members = Get-DistributionGroupMember $Dg.name -resultsize unlimited
			
			if ($members.count -eq 0)
			{
				$userObj = New-Object PSObject
				$userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Alias" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
				Write-Output $Userobj
			}
			else
			{
				Foreach ($Member in $members)
				{
					$userObj = New-Object PSObject
					$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $member.Name
					$userObj | Add-Member NoteProperty -Name "Alias" -Value $member.Alias
					$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $member.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
					Write-Output $Userobj
				}
				
			}
			
		}
		
		; Break
	}
	
	5 {
		
		$AllDG = Get-DynamicDistributionGroup -resultsize unlimited
		
		Foreach ($dg in $allDg)
		{
			
			$Members = Get-Recipient -RecipientPreviewFilter $dg.RecipientFilter -resultsize unlimited
			
			if ($members.count -eq 0)
			{
				$userObj = New-Object PSObject
				$userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Alias" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
				Write-Output $Userobj
			}
			else
			{
				Foreach ($Member in $members)
				{
					$userObj = New-Object PSObject
					$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $member.Name
					$userObj | Add-Member NoteProperty -Name "Alias" -Value $member.Alias
					$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $member.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
					Write-Output $Userobj
				}
				
			}
			
		}
		
		; Break
	}
	
	6 {
		$i = 0
		
		$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\DYDG.csv)"
		
		$AllDG = Get-DynamicDistributionGroup -resultsize unlimited
		
		Foreach ($dg in $allDg)
		{
			
			$Members = Get-Recipient -RecipientPreviewFilter $dg.RecipientFilter -resultsize unlimited
			
			if ($members.count -eq 0)
			{
				$managers = $Dg | Select @{ Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
				
				$userObj = New-Object PSObject
				
				$userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Alias" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "RecipientType" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
				$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
				$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
				$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
				$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.RecipientType
				$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
				#
				$userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $DG.WhenCreated
				$userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $DG.WhenChanged
				
				$output += $UserObj
				
			}
			else
			{
				Foreach ($Member in $members)
				{
					
					$managers = $Dg | Select @{ Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
					
					$userObj = New-Object PSObject
					
					$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $Member.Name
					$userObj | Add-Member NoteProperty -Name "Alias" -Value $Member.Alias
					$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Member.RecipientType
					$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Member.OrganizationalUnit
					$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Member.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
					$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
					$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
					$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.RecipientType
					$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
					#
					$userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $DG.WhenCreated
					$userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $DG.WhenChanged
					
					$output += $UserObj
					
				}
			}
			# update counters and write progress
			$i++
			Write-Progress -activity "Scanning Groups . . ." -status "Scanned: $i of $($allDg.Count)" -percentComplete (($i / $allDg.Count) * 100)
			$output | Export-csv -Path $CSVfile -NoTypeInformation
			
		}
		
		; Break
	}
	
	7 {
		$i = 0
		
		$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\DYDG.csv)"
		
		$Dgname = Read-Host "Enter the DG name or Range (Eg. DynmicDGname , Dy*,*Dy)"
		
		$AllDG = Get-DynamicDistributionGroup $Dgname -resultsize unlimited
		
		Foreach ($dg in $allDg)
		{
			
			$Members = Get-Recipient -RecipientPreviewFilter $dg.RecipientFilter -resultsize unlimited
			
			if ($members.count -eq 0)
			{
				$managers = $Dg | Select @{ Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
				
				$userObj = New-Object PSObject
				
				$userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Alias" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "RecipientType" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmptyGroup
				$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
				$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
				$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
				$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
				$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.RecipientType
				$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
				#
				$userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $DG.WhenCreated
				$userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $DG.WhenChanged
				
				$output += $UserObj
				
			}
			else
			{
				Foreach ($Member in $members)
				{
					
					$managers = $Dg | Select @{ Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
					
					$userObj = New-Object PSObject
					
					$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $Member.Name
					$userObj | Add-Member NoteProperty -Name "Alias" -Value $Member.Alias
					$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Member.RecipientType
					$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Member.OrganizationalUnit
					$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Member.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
					$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
					$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
					$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.RecipientType
					$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
					#
					$userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $DG.WhenCreated
					$userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $DG.WhenChanged
					
					$output += $UserObj
					
				}
			}
			# update counters and write progress
			$i++
			Write-Progress -activity "Scanning Groups . . ." -status "Scanned: $i of $($allDg.Count)" -percentComplete (($i / $allDg.Count) * 100)
			$output | Export-csv -Path $CSVfile -NoTypeInformation
			
		}
		
		; Break
	}
	
	8 {
		
		$Dgname = Read-Host "Enter the Dynamic DG name or Range (Eg. DynamicDGname , DG*,*DG)"
		
		$AllDG = Get-DynamicDistributionGroup $Dgname -resultsize unlimited
		
		Foreach ($dg in $allDg)
		{
			
			$Members = Get-Recipient -RecipientPreviewFilter $dg.RecipientFilter -resultsize unlimited
			
			if ($members.count -eq 0)
			{
				$userObj = New-Object PSObject
				$userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Alias" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmtpyGroup
				$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
				Write-Output $Userobj
			}
			else
			{
				Foreach ($Member in $members)
				{
					$userObj = New-Object PSObject
					$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $member.Name
					$userObj | Add-Member NoteProperty -Name "Alias" -Value $member.Alias
					$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $member.PrimarySmtpAddress
					$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
					Write-Output $Userobj
				}
				
			}
			
		}
		
		; Break
	}
	
	################################################################
	9 {
		
		
		$i = 0
		Write-Host "!!!with great power comes great responsibility!!!" -ForegroundColor magenta
		$TXTfile = Read-Host "Enter the Path of TXT file for the deleted DG's (Eg. C:\DG.txt)"
		
		$AllDG = Get-DistributionGroup -resultsize unlimited
		
		
		Foreach ($dg in $allDg)
		{
			$Members = Get-DistributionGroupMember $Dg.name -resultsize unlimited
			
			if ($members.count -eq 0)
			{
				Remove-DistributionGroup -Identity $dg.alias -confirm:$false
				$dg.alias | Out-File $TXTfile -Append
			}
			
			
			
			
			
			
			# update counters and write progress
			$i++
			Write-Progress -activity "Scanning Groups . . ." -status "Scanned: $i of $($allDg.Count)" -percentComplete (($i / $allDg.Count) * 100)
			
			
		}
		
		; Break
	}
	
	Default { Write-Host "No matches found , Enter Options 1 or 2" -ForeGround "red" }
	
	
}

#################################################################################
#	Theodor Hoff																#
#																				#
#	Copyright (c) 2020 Theodor Hoff <theodorhoff@hotmail.com>					#
#																				#
#	Permission to use, copy, modify, and/or distribute this software for any	#
#	purpose with or without fee is hereby granted, provided that the above		#
#	copyright notice and this permission notice appear in all copies.			#
#																				#
#	THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES	#
#	WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF			#
#	MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR		#
#	ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES		#
#	WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN		#
#	ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF		#
#	OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.				#
#################################################################################


   $Computer = $env:computername																							#Gets Current Computer name.
   
   $Type = " "																												#Defines if Laptop og Desktop (Currently hardcoded to " ")
   
   $Make = Get-CimInstance Win32_ComputerSystem | select -ExpandProperty Manufacturer										#Gets the manufacturer's name
   
   $Model = Get-CimInstance Win32_ComputerSystem | select -ExpandProperty Model												#Gets the Model of the computer
   
   $SN = wmic bios get serialnumber | select -first 1 -skip 2 | Out-String													#Gets the Serialnumber from BIOS
   
   $SN = $SN.TrimEnd()																										#Trims blank spaces after serialnumber
   
   $User = $env:username																									#Gets current Username
   
   $Site = " "																												#Defines current city/site (Currently hardcoded to " ")
   
   $Location = " "																											#Defines current location I.E sales, warehouse. etc. (Currently hardcoded to " ")

   $CPU = Get-CimInstance win32_processor | select -ExpandProperty Name 													#Gets information on the CPU

   $RAM = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb							#Gets information on the RAM, converts result to GB

   $Disk = (Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | Measure-Object -Property Size -Sum).sum/1gb			#Gets information on the Disk, converts result to GB
   
   $HDDType = " "																											#Defines current harddisk type (Currently hardcoded to " ")

   $ComputerOS = (Get-CimInstance Win32_OperatingSystem).Version															#Gets the current OS version

   switch -Wildcard ($ComputerOS){																							#Converts versionnumber to versionname
      "6.1.7600" {$OS = "Windows 7"; break}
      "6.1.7601" {$OS = "Windows 7 SP1"; break}
      "6.2.9200" {$OS = "Windows 8"; break}
      "6.3.9600" {$OS = "Windows 8.1"; break}
      "10.0.*" {$OS = "Windows 10"; break}
      default {$OS = "Unknown Operating System"; break}
   }


$report = New-Object psobject																								#Builds the report object
$report | Add-Member -MemberType NoteProperty -name SystemName -Value $Computer												#
$report | Add-Member -MemberType NoteProperty -name Type -Value $Type														#
$report | Add-Member -MemberType NoteProperty -name Make -Value $Make														#
$report | Add-Member -MemberType NoteProperty -name Model -Value $Model														#
$report | Add-Member -MemberType NoteProperty -name SN -Value $SN															#
$report | Add-Member -MemberType NoteProperty -name Name -Value $User														#
$report | Add-Member -MemberType NoteProperty -name Site -Value $Site														#
$report | Add-Member -MemberType NoteProperty -name Location -Value $Location												#
$report | Add-Member -MemberType NoteProperty -name CPU -Value $CPU															#
$report | Add-Member -MemberType NoteProperty -name Memory -Value $RAM														#
$report | Add-Member -MemberType NoteProperty -name HDDsize -Value $Disk													#
$report | Add-Member -MemberType NoteProperty -name HDDType -Value $HDDType													#
$report | Add-Member -MemberType NoteProperty -name OS -Value $OS															#End of report

$Report | Export-Csv -Path ".\Pcinfo.csv" -Delimiter ';' -NoTypeInformation -append											#Exports to .csv, adds new line if the file exists

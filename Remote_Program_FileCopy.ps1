<#PLS TLS1.2 Transnational DLL Deployment
Operations:
	* Check for existing dll
	* rename existing dll
	* copy new dll
Dependencies for this script:
	* PSKill
Changes:
	* Added pskill Version 1.0.2
	* Fixed Issue with loading from file. Version 1.0.3
	* Added timeout for pskill. Version 1.0.4
	* Testing pskill exit code to see if process is killed. Version 1.0.5
	* Updated how ComputerList is imported. Version 1.0.6
	* Add more debuging to non zero pskill return. Version 1.0.7
	* Updated code to copy if no processes to kill. Version 1.0.8
	* Update to log pskill computers in failed log. Version 1.0.9
	* Added Counter. Version 1.0.10
	* Added checking for if $Computers is a string. Version 1.0.11
	* Added trying IP address instead of dns. Version 1.0.12
	* Updated progress bar info. Version 1.0.13
	* Loop thru all DNS IP Addresss for host. Version 1.1.0
	* Added ablity to copy multible files and Stop a servcie. Version 1.2.0
#>
PARAM (
    [Array]$Computers = $null, 
    [string]$ComputerList = $null,    
    [string]$PSKillPath = $null,    
    [string]$PSServicePath = $null,      
    [string]$Program = $null,    
    [string]$Service = $null,    
    [Parameter(Mandatory=$true)][Array]$SourceFiles = $null,
    [Parameter(Mandatory=$true)][string]$Destination = $null
)
$ScriptVersion = "1.2.0"
#############################################################################
#region User Variables
#############################################################################
$SourceFileObjects=@{}
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + `
		   $MyInvocation.MyCommand.Name + "_" + `
		   (Split-Path -leaf -Path $Destination ) + "_" + `
		   (Get-Date -format yyyyMMdd-hhmm) + ".log")
$maximumRuntimeSeconds = 30
$count = 1

#############################################################################
#endregion User Variables
#############################################################################

#############################################################################
#region Setup Sessions
#############################################################################
#Start logging.
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	try { 
	Start-Transcript -Path $LogFile -Append
	} catch { 
		Stop-transcript
		Start-Transcript -Path $LogFile -Append
	} 
	Write-Host ("Script: " + $MyInvocation.MyCommand.Name)
	Write-Host ("Version: " + $ScriptVersion)
	Write-Host (" ")
}	
#Check which computer input to use to set $Computers
If ([string]::IsNullOrEmpty($ComputerList)) {
	If ([string]::IsNullOrEmpty($Computers)) {
		throw " -Computers or -ComputerList is required."
	}Else{
		# $Computers is already set.
		If (($Computers.GetType().BaseType).Name -eq "String") {
			$Computers = $Computers.split(" ")
		}
	}
}Else{
	[Array]$Computers += Get-Content -Path $ComputerList
}
#Getting Source File Info.
Foreach ($SourceFile in $SourceFiles) {
	#Check $SourceFile
	If (Test-Path $SourceFile -PathType Leaf) {
		$SourceFileObjects.(Split-Path -leaf -Path $SourceFile) = (Get-ChildItem $SourceFile)
	}Else{
		Write-Warning -Message  "-SourceFiles is not a valid file: $SourceFile"
	}
}
If ($SourceFileObjects.Count -le 0) {
	throw " No valid source files"
}
#Check for PSKill
If (-Not ([string]::IsNullOrEmpty($Program))) {
	If (-Not (Test-Path $PSKillPath -PathType Leaf)) {
		throw ("pskill.exe is not found at: " + $PSKillPath)
	}
}
#Check for PSService
If (-Not ([string]::IsNullOrEmpty($Service))) {
	If (-Not (Test-Path $PSServicePath -PathType Leaf)) {
		throw ("psservice.exe is not found at: " + $PSServicePath)
	}
}
#############################################################################
#endregion Setup Sessions
#############################################################################

#############################################################################
#region Main
#############################################################################
#Loop thru servers
Foreach ($Computer in $Computers) {
	Write-Progress -Activity ("Updating Computers with: " + (Split-Path -leaf -Path $SourceFile )) -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer) -percentComplete ($count / $Computers.count*100)
	Write-Host ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer)
	#Test Destination Path
	If (Test-Path $("\\" +  $Computer + "\" + $Destination.replace(":","$"))){
		Foreach ($SourceFileInfo in $SourceFileObjects.GetEnumerator()) {
			#Test for dll.
			If (Test-Path $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) {
				Write-Host ("`t Found at destination dll: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name))
				$DestinationFileInfo = (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name))
				
				#Test for Version Differences 
				If ($SourceFileInfo.value.VersionInfo.FileVersion -gt $DestinationFileInfo.VersionInfo.FileVersion) {
					#Copy newer version
					$NewName =($DestinationFileInfo.Name.replace(".dll","") + "_" + $DestinationFileInfo.VersionInfo.FileVersion + ".dll")
					$DestinationFileInfo = $null
					#Term Service
					If ($Service) {
						Write-Host ("`t`t Stopping Service: " + $Service)
						$process = Start-Process -FilePath $PSServicePath -ArgumentList $("\\" + $Computer + " stop " + $Service) -PassThru -NoNewWindow
						try 
						{
							$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
							If ($process.ExitCode -le 0) {
								Write-Host ("`t`tPSService successfully completed within timeout.")
							}else{
								Write-Warning -Message $('PSService could not kill process. Exit Code: ' + $process.ExitCode)
								
								continue
							}
						}catch{
							Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
							$process | Stop-Process -Force
							#Term Program
							If ($Program) {
								Write-Host ("`t`t killing program: " + $Program)
								$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
								try 
								{
									$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
									If ($process.ExitCode -le 0) {
										Write-Host ("`t`tPSKill successfully completed within timeout.")
									}else{
										Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
										continue
									}
								}catch{
									Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
									$process | Stop-Process -Force
									continue
								} 
							}
						} 
					}
					#Term Program
					If ($Program) {
						Write-Host ("`t`t killing program: " + $Program)
						$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
						try 
						{
							$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
							If ($process.ExitCode -le 0) {
								Write-Host ("`t`tPSKill successfully completed within timeout.")
							}else{
								Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
								continue
							}
						}catch{
							Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
							$process | Stop-Process -Force
							continue
						} 
					}
					#Backup Old DLL
					Write-Host ("`t`t Renaming destination dll: " + $NewName)
					Rename-Item -Path (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) -NewName $NewName
					#copy dll
					Write-Host ("`t`t Copying new dll to destination: " + $("\\" + $Destination.replace(":","$")))
					Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
					
				}Else{
					If ($SourceFileInfo.value.LastWriteTime.ToString("yyyyMMddHHmmssffff") -gt $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff")) {
						#File is newer
						$NewName =($DestinationFileInfo.Name.replace(".dll","") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + ".dll")
						$DestinationFileInfo = $null
						#Term Service
						If ($Service) {
							Write-Host ("`t`t Stopping Service: " + $Service)
							$process = Start-Process -FilePath $PSServicePath -ArgumentList $("\\" + $Computer + " stop " + $Service) -PassThru -NoNewWindow
							try 
							{
								$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
								If ($process.ExitCode -le 0) {
									Write-Host ("`t`tPSService successfully completed within timeout.")
								}else{
									Write-Warning -Message $('PSService could not kill process. Exit Code: ' + $process.ExitCode)
									
									continue
								}
							}catch{
								Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
								$process | Stop-Process -Force
								#Term Program
								If ($Program) {
									Write-Host ("`t`t killing program: " + $Program)
									$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
									try 
									{
										$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
										If ($process.ExitCode -le 0) {
											Write-Host ("`t`tPSKill successfully completed within timeout.")
										}else{
											Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
											continue
										}
									}catch{
										Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
										$process | Stop-Process -Force
										continue
									} 
								}
							} 
						}
						#Term Program
						If ($Program) {
							Write-Host ("`t`t killing program: " + $Program)
							$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
							try 
							{
								$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
								If ($process.ExitCode -le 0) {
									Write-Host ("`t`tPSKill successfully completed within timeout.")
								}else{
									Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
									continue
								}
							}catch{
								Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
								$process | Stop-Process -Force
								continue
							} 
						}
						#Backup Old DLL
						Write-Host ("`t`t Renaming destination dll: " + $NewName)
						Rename-Item -Path (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) -NewName $NewName
						#copy dll
						Write-Host ("`t`t Copying new dll to destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")) + $SourceFileInfo.value.name)
						Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
						
					}else{
						# Older version or same version
						Write-Host ("`t`t Same or Older version: " + $NewName)
						Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime)
						Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.FileVersion)
					}
				}
			}Else{
				#Copy DLL; DLL Missing
				#Term Service
				If ($Service) {
					Write-Host ("`t`t Stopping Service: " + $Service)
					$process = Start-Process -FilePath $PSServicePath -ArgumentList $("\\" + $Computer + " stop " + $Service) -PassThru -NoNewWindow
					try 
					{
						$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
						If ($process.ExitCode -le 0) {
							Write-Host ("`t`tPSService successfully completed within timeout.")
						}else{
							Write-Warning -Message $('PSService could not kill process. Exit Code: ' + $process.ExitCode)
							
							continue
						}
					}catch{
						Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
						$process | Stop-Process -Force
						#Term Program
						If ($Program) {
							Write-Host ("`t`t killing program: " + $Program)
							$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
							try 
							{
								$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
								If ($process.ExitCode -le 0) {
									Write-Host ("`t`tPSKill successfully completed within timeout.")
								}else{
									Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
									continue
								}
							}catch{
								Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
								$process | Stop-Process -Force
								continue
							} 
						}
					} 
				}
				#Term Program
				If ($Program) {
					Write-Host ("`t`t killing program: " + $Program)
					$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
					try 
					{
						$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
						If ($process.ExitCode -le 0) {
							Write-Host ("`t`tPSKill successfully completed within timeout.")
						}else{
							Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
							continue
						}
					}catch{
						Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
						$process | Stop-Process -Force
						continue
					} 
				}
				Write-Host ("`t copying missing dll to destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")))
				Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
			}
		}
	}Else{
	#Testing using IP
		$Cerror = $lastexitcode
		$OldComputer = $Computer
		Foreach ($Computer in ((Resolve-DnsName -Name $Computer).IPAddress)) {
			#Test Destination Path
			If ([string]::IsNullOrEmpty($OldComputer)) {
				If (Test-Path $("\\" +  $Computer + "\" + $Destination.replace(":","$"))){
					#Test for dlls.
					Foreach ($SourceFileInfo in $SourceFileObjects) {
						If (Test-Path $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) {
							Write-Host ("`t Found at destination dll: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name))
							$DestinationFileInfo = (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name))
							
							#Test for Version Differences 
							If ($SourceFileInfo.value.VersionInfo.FileVersion -gt $DestinationFileInfo.VersionInfo.FileVersion) {
								#Copy newer version
								$NewName =($DestinationFileInfo.Name.replace(".dll","") + "_" + $DestinationFileInfo.VersionInfo.FileVersion + ".dll")
								$DestinationFileInfo = $null
								#Term Service
								If ($Service) {
									Write-Host ("`t`t Stopping Service: " + $Service)
									$process = Start-Process -FilePath $PSServicePath -ArgumentList $("\\" + $Computer + " stop " + $Service) -PassThru -NoNewWindow
									try 
									{
										$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
										If ($process.ExitCode -le 0) {
											Write-Host ("`t`tPSService successfully completed within timeout.")
										}else{
											Write-Warning -Message $('PSService could not kill process. Exit Code: ' + $process.ExitCode)
											
											continue
										}
									}catch{
										Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
										$process | Stop-Process -Force
										#Term Program
										If ($Program) {
											Write-Host ("`t`t killing program: " + $Program)
											$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
											try 
											{
												$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
												If ($process.ExitCode -le 0) {
													Write-Host ("`t`tPSKill successfully completed within timeout.")
												}else{
													Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
													continue
												}
											}catch{
												Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
												$process | Stop-Process -Force
												continue
											} 
										}
									} 
								}
								#Term Program
								If ($Program) {
									Write-Host ("`t`t killing program: " + $Program)
									$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
									try 
									{
										$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
										If ($process.ExitCode -le 0) {
											Write-Host ("`t`tPSKill successfully completed within timeout.")
										}else{
											Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
											continue
										}
									}catch{
										Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
										$process | Stop-Process -Force
										continue
									} 
								}
								#Backup Old DLL
								Write-Host ("`t`t Renaming destination dll: " + $NewName)
								Rename-Item -Path (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) -NewName $NewName
								#copy dll
								Write-Host ("`t`t Copying new dll to destination: " + $("\\" + $Destination.replace(":","$")))
								Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
								
							}Else{
								If ($SourceFileInfo.value.LastWriteTime.ToString("yyyyMMddHHmmssffff") -gt $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff")) {
									#File is newer
									$NewName =($DestinationFileInfo.Name.replace(".dll","") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + ".dll")
									$DestinationFileInfo = $null
									#Term Service
									If ($Service) {
										Write-Host ("`t`t Stopping Service: " + $Service)
										$process = Start-Process -FilePath $PSServicePath -ArgumentList $("\\" + $Computer + " stop " + $Service) -PassThru -NoNewWindow
										try 
										{
											$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
											If ($process.ExitCode -le 0) {
												Write-Host ("`t`tPSService successfully completed within timeout.")
											}else{
												Write-Warning -Message $('PSService could not kill process. Exit Code: ' + $process.ExitCode)
												
												continue
											}
										}catch{
											Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
											$process | Stop-Process -Force
											#Term Program
											If ($Program) {
												Write-Host ("`t`t killing program: " + $Program)
												$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
												try 
												{
													$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
													If ($process.ExitCode -le 0) {
														Write-Host ("`t`tPSKill successfully completed within timeout.")
													}else{
														Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
														continue
													}
												}catch{
													Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
													$process | Stop-Process -Force
													continue
												} 
											}
										} 
									}
									#Term Program
									If ($Program) {
										Write-Host ("`t`t killing program: " + $Program)
										$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
										try 
										{
											$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
											If ($process.ExitCode -le 0) {
												Write-Host ("`t`tPSKill successfully completed within timeout.")
											}else{
												Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
												continue
											}
										}catch{
											Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
											$process | Stop-Process -Force
											continue
										} 
									}
									#Backup Old DLL
									Write-Host ("`t`t Renaming destination dll: " + $NewName)
									Rename-Item -Path (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) -NewName $NewName
									#copy dll
									Write-Host ("`t`t Copying new dll to destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")) + $SourceFileInfo.value.name)
									Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
									
								}else{
									# Older version or same version
									Write-Host ("`t`t Same or Older version: " + $NewName)
									Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime)
									Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.FileVersion)
								}
							}
						}Else{
							#Copy DLL; DLL Missing
							#Term Service
							If ($Service) {
								Write-Host ("`t`t Stopping Service: " + $Service)
								$process = Start-Process -FilePath $PSServicePath -ArgumentList $("\\" + $Computer + " stop " + $Service) -PassThru -NoNewWindow
								try 
								{
									$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
									If ($process.ExitCode -le 0) {
										Write-Host ("`t`tPSService successfully completed within timeout.")
									}else{
										Write-Warning -Message $('PSService could not kill process. Exit Code: ' + $process.ExitCode)
										
										continue
									}
								}catch{
									Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
									$process | Stop-Process -Force
									#Term Program
									If ($Program) {
										Write-Host ("`t`t killing program: " + $Program)
										$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
										try 
										{
											$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
											If ($process.ExitCode -le 0) {
												Write-Host ("`t`tPSKill successfully completed within timeout.")
											}else{
												Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
												continue
											}
										}catch{
											Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
											$process | Stop-Process -Force
											continue
										} 
									}
								} 
							}
							#Term Program
							If ($Program) {
								Write-Host ("`t`t killing program: " + $Program)
								$process = Start-Process -FilePath $PSKillPath -ArgumentList $("-t -nobanner \\" + $Computer + " " + $Program) -PassThru -NoNewWindow
								try 
								{
									$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
									If ($process.ExitCode -le 0) {
										Write-Host ("`t`tPSKill successfully completed within timeout.")
									}else{
										Write-Warning -Message $('PSKill could not kill process. Exit Code: ' + $process.ExitCode)
										continue
									}
								}catch{
									Write-Warning -Message 'PSKill exceeded timeout, will be killed now.' 
									$process | Stop-Process -Force
									continue
								} 
							}
							Write-Host ("`t copying missing dll to destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")))
							Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
						}
					}Else{
						#Error missing folder or computer.
						Write-Warning -Message ("Error: Folder or Computer does not exists by using IP: $Computer with error code: $lastexitcode " )
						Write-Host ("`t Path: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")))
						If (Test-Connection -ComputerName $Computer -Quiet){ 
							Write-Host ("`t`t Host is up") -ForegroundColor green
						}else{
							Write-Warning -Message ("`t`t Host is Down")
						}
						Write-Host ("`t Logging Failed Computer")
						Add-Content ($LogFile + "_error_computers.txt") ("$OldComputer")
					}
				}
			}Else{
				#Error missing folder or computer.
				Write-Warning -Message ("Error: Computer has IP Address $Computer in DNS: $OldComputer with error code: $Cerror")
				Write-Host ("`t Path: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")))
				If (Test-Connection -ComputerName $Computer -Quiet){ 
					Write-Host ("`t`t Host is up") -ForegroundColor green
				}else{
					Write-Warning -Message ("`t`t Host is Down")
				}
				Write-Host ("`t Logging Failed Computer")
				Add-Content ($LogFile + "_error_computers.txt") ("$OldComputer")
			}
		}
	}
	$count++
}
#############################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}

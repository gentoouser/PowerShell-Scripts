<#File Deployment
Operations:
	* Check for existing File
	* rename existing File
	* copy new File
Dependencies for this script:
	* PSKill
	* PSService
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
	* Fixed File Looping issue. Version 1.2.1
	* Added Service Start after copy. Version 1.2.2
	* Fixed Renameing bug Version 1.2.3
	* Fixed Issue to allow other extension besides dll. Version 1.3.0
	* Fixed PSService start issue Version 1.3.1
	* Replaced Resolve-DnsName with[Net.DNS]::GetHostEntry for older PS compatibility. Version 1.3.2
	* Added Verbose logging. Version 1.3.2
#>
PARAM (
    [Array]$Computers = $null, 
    [string]$ComputerList = $null,    
    [string]$PSKillPath = $null,    
    [string]$PSServicePath = $null,     
    [string]$Program = $null,    
    [string]$Service = $null,    
    [Parameter(Mandatory=$true)][Array]$SourceFiles = $null,
    [Parameter(Mandatory=$true)][string]$Destination = $null,
    [switch]$Verbose = $false
)
$ScriptVersion = "1.3.2"
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
	#Reset logging variables
	$NoChange = $false
	$UpdatesNeeded = $false
	$MissingFiles = $false
	$ComputerError = $false
	Write-Progress -Activity ("Updating Computers with: " + (Split-Path -leaf -Path $SourceFile )) -Status ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer) -percentComplete ($count / $Computers.count*100)
	Write-Host ("( " + $count + "\" + $Computers.count + ") Computer: " + $Computer)
	#Test Destination Path
	If (Test-Path $("\\" +  $Computer + "\" + $Destination.replace(":","$"))){
		Foreach ($SourceFileInfo in $SourceFileObjects.GetEnumerator()) {
			#Test for File.
			If (Test-Path $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) {
				Write-Host ("`t Found at destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name))
				$DestinationFileInfo = (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name))
				
				#Test for Version Differences 
				If ($SourceFileInfo.value.VersionInfo.FileVersion -gt $DestinationFileInfo.VersionInfo.FileVersion) {
					#Copy newer version
					$NewName =($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.VersionInfo.FileVersion + "." + $DestinationFileInfo.Extension)
					$DestinationFileInfo = $null
					#Term Service
					If ($Service) {
						Write-Host ("`t`t Stopping Service: " + $Service)
						$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "stop",'"' + $Service + '"') -PassThru -NoNewWindow
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
					#Backup Old 
					Write-Host ("`t`t Renaming destination: " + $NewName)
					Rename-Item -Path (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) -NewName $NewName
					#copy 
					Write-Host ("`t`t Copying new $SourceFile to destination: " + $("\\" + $Destination.replace(":","$")))
					Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
					$UpdatesNeeded = $true
					If ($Service) {
						Write-Host ("`t`t Stopping Service: " + $Service)
						$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "start",'"' + $Service + '"') -PassThru -NoNewWindow

						$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
						If ($process.ExitCode -le 0) {
							Write-Host ("`t`tPSService Start successfully completed within timeout.")
						}else{
							Write-Warning -Message $('PSService Start could not kill process. Exit Code: ' + $process.ExitCode)
							
							continue
						}
					}
				}Else{
					If ($SourceFileInfo.value.LastWriteTime.ToString("yyyyMMddHHmmssffff") -gt $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff")) {
						#File is newer
						$NewName =($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + "." + $DestinationFileInfo.Extension)
						$DestinationFileInfo = $null
						#Term Service
						If ($Service) {
							Write-Host ("`t`t Stopping Service: " + $Service)
							$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "stop",'"' + $Service + '"') -PassThru -NoNewWindow
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
						#Backup Old
						Write-Host ("`t`t Renaming destination: " + $NewName)
						Rename-Item -Path (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) -NewName $NewName
						#copy 
						Write-Host ("`t`t Copying new $SourceFile to destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")) + $SourceFileInfo.value.name)
						Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
						$UpdatesNeeded = $true
						If ($Service) {
							Write-Host ("`t`t Stopping Service: " + $Service)
							$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "start",'"' + $Service + '"') -PassThru -NoNewWindow

							$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
							If ($process.ExitCode -le 0) {
								Write-Host ("`t`tPSService Start successfully completed within timeout.")
							}else{
								Write-Warning -Message $('PSService Start could not kill process. Exit Code: ' + $process.ExitCode)
								
								continue
							}
						}
					}else{
						# Older version or same version
						Write-Host ("`t`t Same or Older version: " + $NewName)
						Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime)
						Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.FileVersion)
						$NoChange = $true
					}
				}
			}Else{
				#Copy; Missing
				#Term Service
				If ($Service) {
					Write-Host ("`t`t Stopping Service: " + $Service)
					$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "stop",'"' + $Service + '"') -PassThru -NoNewWindow
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
				Write-Host ("`t copying missing $SourceFile to destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")))
				Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
				$MissingFiles = $true
				If ($Service) {
					Write-Host ("`t`t Stopping Service: " + $Service)
					$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "start",'"' + $Service + '"') -PassThru -NoNewWindow

					$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
					If ($process.ExitCode -le 0) {
						Write-Host ("`t`tPSService Start successfully completed within timeout.")
					}else{
						Write-Warning -Message $('PSService Start could not kill process. Exit Code: ' + $process.ExitCode)
						
						continue
					}
				}
			}
		}
	}Else{
	#Testing using IP
		$Cerror = $lastexitcode
		If (-Not ([string]::IsNullOrEmpty($Computer))) {
			$OldComputer = $Computer
			Foreach ($Computer in (([Net.DNS]::GetHostEntry($OldComputer)).addresslist.ipaddresstostring))
				#Test Destination Path
				If (-Not ([string]::IsNullOrEmpty($Computer))) {
					If (Test-Path $("\\" +  $Computer + "\" + $Destination.replace(":","$"))){
						#Test for Files.
						Foreach ($SourceFileInfo in $SourceFileObjects) {
							If (Test-Path $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) {
								Write-Host ("`t Found at destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name))
								$DestinationFileInfo = (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name))
								
								#Test for Version Differences 
								If ($SourceFileInfo.value.VersionInfo.FileVersion -gt $DestinationFileInfo.VersionInfo.FileVersion) {
									#Copy newer version
									$NewName =($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.VersionInfo.FileVersion + "." + $DestinationFileInfo.Extension)
									$DestinationFileInfo = $null
									#Term Service
									If ($Service) {
										Write-Host ("`t`t Stopping Service: " + $Service)
										$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "stop",'"' + $Service + '"') -PassThru -NoNewWindow
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
									#Backup Old
									Write-Host ("`t`t Renaming destination: " + $NewName)
									Rename-Item -Path (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) -NewName $NewName
									#copy 
									Write-Host ("`t`t Copying new $SourceFile to destination: " + $("\\" + $Destination.replace(":","$")))
									Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
									$UpdatesNeeded = $true
									If ($Service) {
										Write-Host ("`t`t Stopping Service: " + $Service)
										$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "start",'"' + $Service + '"') -PassThru -NoNewWindow

										$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
										If ($process.ExitCode -le 0) {
											Write-Host ("`t`tPSService Start successfully completed within timeout.")
										}else{
											Write-Warning -Message $('PSService Start could not kill process. Exit Code: ' + $process.ExitCode)
											
											continue
										}
									}
								}Else{
									If ($SourceFileInfo.value.LastWriteTime.ToString("yyyyMMddHHmmssffff") -gt $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff")) {
										#File is newer
										$NewName =($DestinationFileInfo.Name.replace("." + $DestinationFileInfo.Extension,"") + "_" + $DestinationFileInfo.LastWriteTime.ToString("yyyyMMddHHmmssffff") + "." + $DestinationFileInfo.Extension)
										$DestinationFileInfo = $null
										#Term Service
										If ($Service) {
											Write-Host ("`t`t Stopping Service: " + $Service)
											$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "stop",'"' + $Service + '"') -PassThru -NoNewWindow
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
										#Backup Old
										Write-Host ("`t`t Renaming destination: " + $NewName)
										Rename-Item -Path (Get-ChildItem $("\\" +  $Computer + "\" + $Destination.replace(":","$") + "\" + $SourceFileInfo.value.name)) -NewName $NewName
										#copy 
										Write-Host ("`t`t Copying new $SourceFile to destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")) + $SourceFileInfo.value.name)
										Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
										$UpdatesNeeded = $true
										If ($Service) {
											Write-Host ("`t`t Stopping Service: " + $Service)
											$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "start",'"' + $Service + '"') -PassThru -NoNewWindow

											$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
											If ($process.ExitCode -le 0) {
												Write-Host ("`t`tPSService Start successfully completed within timeout.")
											}else{
												Write-Warning -Message $('PSService Start could not kill process. Exit Code: ' + $process.ExitCode)
												
												continue
											}
										}
									}else{
										# Older version or same version
										Write-Host ("`t`t Same or Older version: " + $NewName)
										Write-Host ("`t`t`t Destination Modified: " + $DestinationFileInfo.LastWriteTime)
										Write-Host ("`t`t`t Destination Version: " + $DestinationFileInfo.VersionInfo.FileVersion)
										$NoChange = $true
									}
								}
							}Else{
								#Copy; Missing
								#Term Service
								If ($Service) {
									Write-Host ("`t`t Stopping Service: " + $Service)
									$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "stop",'"' + $Service + '"') -PassThru -NoNewWindow
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
								Write-Host ("`t Copying missing $SourceFile to destination: " + $("\\" +  $Computer + "\" + $Destination.replace(":","$")))
								Copy-Item $SourceFile -Destination $("\\" +  $Computer + "\" + $Destination.replace(":","$"))
								$MissingFiles = $true
								If ($Service) {
									Write-Host ("`t`t Stopping Service: " + $Service)
									$process = Start-Process -FilePath $PSServicePath -ArgumentList @("\\" + $Computer, "start",'"' + $Service + '"') -PassThru -NoNewWindow

									$process | Wait-Process -Timeout $maximumRuntimeSeconds -ErrorAction Stop 
									If ($process.ExitCode -le 0) {
										Write-Host ("`t`tPSService Start successfully completed within timeout.")
									}else{
										Write-Warning -Message $('PSService Start could not kill process. Exit Code: ' + $process.ExitCode)
										
										continue
									}
								}
							}
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
						$ComputerError = $true
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
					$ComputerError = $true
				}
			}
		}
	}
	#Extra logging
	If ($Verbose) {
		If ($NoChange) {Add-Content ($LogFile + "_NoChanges.txt") ("$Computer")}
		If ($UpdatesNeeded) {Add-Content ($LogFile + "_UpdatesNeeded.txt") ("$Computer")}
		If ($MissingFiles) {Add-Content ($LogFile + "_MissingFiles.txt") ("$Computer")}
	}
	IF ($ComputerError) {Add-Content ($LogFile + "_ErrorComputers.txt") ("$OldComputer")}
	#Increase Progress counter
	$count++
}
#############################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}

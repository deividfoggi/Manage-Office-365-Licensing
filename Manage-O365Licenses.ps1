#    This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
#    THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,        
#    INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
#    We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute
#    the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks
#    to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on
#    Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us
#    and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or resultfrom the 
#    use or distribution of the Sample Code.
#    Please note: None of the conditions outlined in the disclaimer above will supersede the terms and conditions contained 
#    within the Premier Customer Services Description.

#Log the start of script execution
Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "############ Script Started ############"

#Module import
try{
    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module MSOnline -ErrorAction Stop
    Import-Module MSOLLicenseManagement -ErrorAction Stop
}
catch{
    Write-Log -LogLevel Error -UserOrGroup "SCRIPT" -Message $_.Exception.Message
    Exit
}

#Get current date/time to start measure the duration of execution
$startTimer = Get-Date

#Path to licensing config file
$licenseFilePath = ".\Licenses.csv"

#Imports the licensing config file
$licenseConfigFile = Import-Csv -Path $licenseFilePath -Delimiter ","

#Format of the license module log file
$licenseModuleLogFileName = "LicenseModule_$((Get-Date).ToString('ddMMyyyy')).log"

#Creates a function to log scripting actions
Function Write-Log(){
    param
        (
            [ValidateSet("Error", "Warning", "Info")]$LogLevel,
            $UserOrGroup,
            [string]$Message
        )

    #Name of the log file containing date/time
    $logFileName = "ManageO365Licenses_$((Get-Date).ToString('ddMMyyyy')).log"

    #Header of the log file in csv format
    $header = "datetime,user,action,message"

    #Date/time of the log entry
    $datetime = (Get-Date).ToString('dd/MM/yyyy hh:mm:ss')

    #Log entry variable containing each parameter passed when funcion is called
    $logEntry = "$datetime,$LogLevel,$UserOrGroup,$Message"

    #Check if log file already exists
    if(-not(Test-Path $logFileName)){
        #Try to create a file and add content
        try{
            New-Item -Path $logFileName -ErrorAction Stop
            Add-Content -Path $logFileName -Value $header -ErrorAction Stop
        }
        catch{
            #Prints the exception related to log file creation/write
            $_.Exception.Message
            Exit
        }
    }

    #Adds a log entry into log file    
    try{
        Add-Content -Path $logFileName -Value $logEntry -ErrorAction Stop
    }
    catch{
        $_.Exception.Message
        #Prints the exception related to log file creation/write
        Exit
    }
    
}

#Function to get o365/msoline services credentials from a file
Function GetCredentialOnDisk{

    try{
        #User with privileges to manage licensing
        $username = "admin@foggioncloud.onmicrosoft.com"
        #Get password from the file
        $password = Get-Content .\cred.sec -ErrorAction Stop | ConvertTo-SecureString -ErrorAction Stop

        #Creates a PSCredential object
        $credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

        #Write a log entry
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "Credentials to connect to MSOnline Services acquired. User $($username) has been used to connect"
        
        #Returns the PSCredential object
        return $credential
    }
    catch{
        #Write the exception related to either get password from file or create a PSCredential object
        Write-Log -LogLevel Error -UserOrGroup "SCRIPT" -Message $_.Exception.Message
    }
}

#Function to connect to MSOnline Services
Function ConnectMsolService{
    #Try to connect
    try{
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "Import MSOnline module"
        #Import Module MSOnline Services
        Import-Module MSOnline -ErrorAction Stop
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "Connecting to Microsoft Online Services"
        #Start the connection cmdlet
        Connect-MsolService -Credential (GetCredentialOnDisk) -ErrorAction Stop
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "Connected to Microsoft Online Services"
    }
    catch{
        Write-Log -LogLevel Error -UserOrGroup "SCRIPT" -Message $_.Exception.Message
        
        #Get current date/time to calculates the execution time span before finishing script execution
        $stopTimer = Get-Date

        #Add a log entry registering script stop and execution time span
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "############ Script Stopped. Execution Time: $((New-TimeSpan -Start $startTimer -End $stopTimer).ToString("dd\.hh\:mm\:ss")) ############"
        
        #Exits current PowerShell session due to an issue to connect to MSOnline Services
        Exit
    }
}

#Function to monitor group membership
Function GroupMonitor{
    
    param(
        $GroupName
    )
    
    #Making sure sensitive variables are clean
    $currentMembers = $null
    $compareMembers = $null

    #Get group membership and saves to a file
    Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Starting group membership acquisition"
    
    #Try to get a list of group members to save it in a reference file
    try{
        $currentMembers = Get-ADGroup -Identity $groupName -Properties Members -ErrorAction Stop | Select-Object -ExpandProperty Members | Get-ADUser -ErrorAction Stop
        Write-Log -LogLevel "Info" -UserOrGroup $groupName "Group members acquired succesfully"

        #If group count is greater than 0/not empty
        if(($currentMembers|Measure-Object).Count -gt 0){
            try{
                #Creates a variable with the name of the reference file containing date/time
                $logFileName =  "$($groupName)_Members_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
                #Export group membership to a csv file
                $currentMembers | Export-Csv "$($logFileName)" -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel "Error" -UserOrGroup $groupName "Group membership list saved in csv file $($logFileName)"
            }
            catch{
                #Writes a log entry with the error related to export membership to a csv file
                Write-Log -LogLevel "Error" -UserOrGroup $groupName $_.Exception.Message
            }
        }
        #If group count is zero
        else{
            Write-Log -LogLevel "Error" -UserOrGroup $groupName "Group $($groupName) is empty"
        }
    }
    catch{
        #Writes a log entry with the error related to get AD group membership
        Write-Log -LogLevel "Error" -UserOrGroup $groupName $_.Exception.Message
    }

    #Creates a variable containing the log file name related to group membership changes
    $logFileName = "$($groupName)_MembersAddedRemoved_ddMMyyyy_hhmmss.csv"

    #If csv membership file count in disk is greater than 1
    If((Get-Item "$($groupName)_Members_*" |Measure-Object).Count -gt 1){
        try{
            #Compares the current group membership with the previous execution membership which has been saved in a csv file
            $compareMembers = Compare-Object -DifferenceObject (Import-Csv (Get-Item "$($groupName)_Members_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1)) -ReferenceObject (Import-Csv (Get-Item "$($groupName)_Members_*" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 1 -First 1)) -PassThru -Property Name -ErrorAction Stop
            Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Group membership comparison ran succesfully. Users added/removed saved into file $($logFileName)"
        }
        catch{
            Write-Log -LogLevel "Error" -UserOrGroup $groupName -Message $_.Exception.Message
        }
    }
    #If csv membership file count in disk is equal or less than 1, starts "first execution mode"
    else{
        Write-Log -LogLevel Error -UserOrGroup $groupName -Message "Unable to compare current list and previous one. Script in 'First Execution Mode' or the file with group membership is missing. If in 'First Exeuction Mode', all users in the group will receive the license"

        #If the current members count is greater than 0
        if(($currentMembers|Measure-Object).Count -gt 0){

            #Once this is in "first execution mode", adds a column to each object in current member array with "Added Member" signal
            $currentMembers | Add-Member -Name "SideIndicator" -MemberType NoteProperty -Value "=>" -Force
            
            #Exports a list containing all members that have been either added or removed to/from the AD group
            try{
                $currentMembers | Export-Csv "$($logFIleName)" -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Group membership activity sucessfully exported to log file $($logFileName)"
            }
            catch{
                #Writes a log entry with the error related to csv export
                Write-Log -LogLevel "Error" -UserOrGroup $groupName -Message $_.Exception.Message
            }

            #Returns the current membership list as a result of the compare cmdlet because it is in "first execution mode"
            return $currentMembers
        }
        #If the current members count is equal to zero
        else
        {   
            #Writes a log entry with a static string error sentence that it is unable to list current members
            Write-Log -LogLevel "Error" -UserOrGroup $groupName -Message "Unable to get group current membership"
        }
    }

    #If the array with add/removed members is greater than 0/not empty (now it is not in "first execution mode")
    if(($compareMembers|Measure-Object).Count -gt 0){
            Write-Log -LogLevel "Info" -UserOrGroup $group -Message "Group membership change detected. Members added: $(($compareMembers|Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count) - Members removed: $(($compareMembers|Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).Count)"
            
            #Exports a list containing all members that have been either added or removed to/from the AD group
            try{
                $compareMembers | Export-Csv "$($logFIleName)" -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Group membership activity exporeted to log file $($logFIleName)"
            }
            catch{
                #Writes a log entry with the error related to csv export
                Write-Log -LogLevel "Error" -UserOrGroup $groupName -Message $_.Exception.Message
            }

            #Created an array object to receive AD user objects. This function cannot return the objects resulted from Compare-Object due to issues in reading the objects that have other objects inside
            $arrCompareMembers = @()

            #For each compared object
            foreach($userMember in $compareMembers){
                
                #Get the respective AD User object using its DN
                $user = Get-ADUser -Identity $userMember.DistinguishedName

                #If the object has the signal added "=>" adds the respective signal string to the AD User object in a new column
                if($userMember.SideIndicator -eq "=>"){
                    $user | Add-Member -Name SideIndicator -MemberType NoteProperty -Value "=>" -Force
                }

                #If the object has the signal removed "=>" the respective signal string to the AD User object in a new column
                if($userMember.SideIndicator -eq "<="){
                    $user | Add-Member -Name SideIndicator -MemberType NoteProperty -Value "<=" -Force
                }
                #Adds the AD user object to the array that's going to be returned
                $arrCompareMembers += $user
            }

            #Return the final list of members that have been either added or removed            
            return $arrCompareMembers
     }
}

#Function to monitor the license config file
Function LicensePlansMonitor{
    
    Param(
        $LicenseFile,
        $Group
    )

    #Creates an rray to return the result of the comparison
    $arrLicenseChangeResult = @()

    #For each license in license config file
    foreach($license in $LicenseFile|Where-Object{$_.Group -eq $Group}){
    
        #Making sure comparison variable is clean
        $comparePlans = $null

        #Creates a sufix name for the file that's going to be used to register current SKU configuration replacing : to - in order to avoid issues due to file path reserved character
        $plansCompareFileSufix = "$($license.Group)_$(($license.SKU).Replace(":","-"))"
        
        #Following line is under review and currently not being used
        #$currentPlans = ListPlans -Plans $license.Plans
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "Starting to list license plans in the license configuration file for SKU $($license.SKU)"
        Write-Log -LogLevel "Info" -UserOrGroup $license.SKU "Plans of the group $($Group) successfully listed: $($license.Plans)"

        #If Plans column for the current license is not null
        if(![string]::IsNullOrEmpty($license.Plans)){

            #Creates a variable with the current execution license plans comparison file
            $plansFileName = "$($plansCompareFileSufix)_Plans_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
            try{
                #Created an array to be returned by this function
                $arrPlans = @()
                #For each plan in Plans column (semi-comman separated)
                foreach($plan in ($license.Plans.Split(";"))){
                    #Creates a new PSObject with its columns in order to list each plan by line in plans comparison file
                    $objPlans = New-Object psobject
                    $objPlans | Add-Member -Name "Group" -MemberType NoteProperty -Value $license.Group
                    $objPlans | Add-Member -Name "SKU" -MemberType NoteProperty -Value $license.SKU
                    $objPlans | Add-Member -Name "Plan" -MemberType NoteProperty -Value $plan
                    #Increments the resulting array
                    $arrPlans += $objPlans
                }
                
                #Exports a list with all plans to a csv file
                $arrPlans | Export-Csv $plansFileName -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel Info -UserOrGroup $license.SKU "PLans of the SKU $($SKU) saved in csv file $($plansFileName)"
            }
            catch{
                Write-Log -LogLevel Error -UserOrGroup $license.SKU $_.Exception.Message
            }   
        }
        #If the Plans column is empty (For now, I've decided to treat empty PLans columns objects in a dedicated statement in order to have an specfic error treatment)
        else{
            #Creates a new PSObject with its columns in order to list each plan by line in plans comparison file
            $plansFileName = "$($plansCompareFileSufix)_Plans_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
            try{
                #Creates a new PSObject with its columns with the empty column
                $objPlans = New-Object psobject
                $objPlans | Add-Member -Name "Group" -MemberType NoteProperty -Value $license.Group
                $objPlans | Add-Member -Name "SKU" -MemberType NoteProperty -Value $license.SKU
                $objPlans | Add-Member -Name "Plan" -MemberType NoteProperty -Value $null
                $objPlans | Export-Csv $plansFileName -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel Error -UserOrGroup $groupName "No plans found in SKU $($license.SKU)"
            }
            catch{
                Write-Log -LogLevel Error -UserOrGroup $license.SKU $_.Exception.Message
            }
        }

        #If the license plan comparison files in disk is greater than 1
        If((Get-Item "$($plansCompareFileSufix)_Plans_*" |Measure-Object).Count -gt 1){
            #Creates a variable with the name of the plans comparison result file for this execution
            $plansCompareFileName = "$($plansCompareFileSufix)_PlansAddedRemoved_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
            #Compares current plans in the file from last execution with the list from current execution
            try{
                #Line 341 below is a previous way to compare that produced some inconsistent issues. 342 will be used till we cannot confirm it is the best way
                #$lastRunPlan = ListPlans -Plans (Import-Csv (Get-Item "$($plansCompareFileSufix)_Plans_*" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 1 -First 1)).Plans
                $comparePlans = Compare-Object -DifferenceObject (Import-Csv (Get-Item "$($plansCompareFileSufix)_Plans_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1)) -ReferenceObject (Import-Csv (Get-Item "$($plansCompareFileSufix)_Plans_*" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 1 -First 1)) -PassThru -Property Plan -ErrorAction Stop
                Write-Log -LogLevel Info -UserOrGroup $license.SKU -Message "Plan configuration change detected succesfully. List of plans exported into log file $($plansCompareFileName)"
            }
            catch{
                Write-Log -LogLevel "Error" -UserOrGroup $license.SKU -Message $_.Exception.Message
            }
        }
        #
        #If the license plan comparison files in disk is greater than 1
        else{
            #Creates a variable with the name of the plans comparison result file for this execution
            $plansCompareFileName = "$($plansCompareFileSufix)_PlansAddedRemoved_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
            Write-Log -LogLevel Error -UserOrGroup $license.SKU -Message "Unable to compares license plan. Script in 'First Execution Mode' ou files with plans configuration not found. If in 'First Exeuction Mode', all users in the group will receive the license"

            #If Plans column of the current license is not empty
            if(![string]::IsNullOrEmpty($license.Plans)){
        
                #Created an array to be returned by this function
                $arrPlans = @()
                foreach($plan in ($license.Plans.Split(";"))){
                    #Creates a new PSObject with its columns with the empty column
                    $objPlans = New-Object psobject
                    $objPlans | Add-Member -Name "Group" -MemberType NoteProperty -Value $license.Group
                    $objPlans | Add-Member -Name "SKU" -MemberType NoteProperty -Value $license.SKU
                    $objPlans | Add-Member -Name "Plan" -MemberType NoteProperty  -Value $plan
                    $objPlans | Add-Member -Name "SideIndicator" -MemberType NoteProperty -Value "=>"
                    $arrPlans += $objPlans
                }
                
                #Exports a list with the plans that have been either added or removed from the license config file
                try{
                    $arrPlans | Where-Object{$_.Plan -ne ""} | Export-Csv $plansCompareFileName -NoTypeInformation -ErrorAction Stop
                    Write-Log -LogLevel "Info" -UserOrGroup $license.SKU -Message "Group membership activity sucessfully exported to log file $($plansCompareFileName)"
                }
                catch{
                    Write-Log -LogLevel "Error" -UserOrGroup $license.SKU -Message $_.Exception.Message
                }
                #Increments the resultant array
                $arrLicenseChangeResult += $arrPlans
            }
            #If the Plans column is empty (For now, I've decided to treat empty Plans columns objects in a dedicated statement in order to have an specfic error treatment)
            else
            {
                #Creates a variable with the current execution license plans comparison file
                $plansCompareFileName = "$($plansCompareFileSufix)_Plans_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"

                try{
                    #Creates a new PSObject with its columns with the empty column
                    $objPlans = New-Object psobject
                    $objPlans | Add-Member -Name "Group" -MemberType NoteProperty -Value $license.Group
                    $objPlans | Add-Member -Name "SKU" -MemberType NoteProperty -Value $license.SKU
                    $objPlans | Add-Member -Name "Plan" -MemberType NoteProperty -Value $null
                    $objPlans | Export-Csv $plansFileName -NoTypeInformation -ErrorAction Stop
                    Write-Log -LogLevel Error -UserOrGroup $groupName "No plan found for SKU $($license.SKU)"
                }
                catch{
                    Write-Log -LogLevel Error -UserOrGroup $license.SKU $_.Exception.Message
                }
            }
        }

        if(($comparePlans|Measure-Object).Count -gt 0){
                Write-Log -LogLevel "Info" -UserOrGroup $license.SKU -Message "Changes in license configuration file detected. Plans added: $(($comparePlans|Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count) - Plans removed: $(($comparePlans|Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).Count)"
                
                #Exports a list with the plans that have been either added or removed from the license config file
                try{
                    $plansCompareFileName = "$($plansCompareFileSufix)_PlansAddedRemoved_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
                    $comparePlans | Where-Object{$_.Plan -ne ""} | Export-Csv $plansCompareFileName -NoTypeInformation -ErrorAction Stop
                    Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Plan configuration change detected succesfully. List of plans exported into log file $($plansCompareFileName)"
                }
                catch{
                    Write-Log -LogLevel "Error" -UserOrGroup $license.SKU -Message $_.Exception.Message
                }
                
                #Increments the resultant array
                $arrLicenseChangeResult += $comparePlans
        }
    }

    #Returns the list of resultant plans in the case Plans column is not empty 
    return $arrLicenseChangeResult | Where-Object{$_.Plan -ne ""}
}

#Function that creates an array with the plans separated by semi-comma 
Function ListPlans{
    param(
        $Plans
    )

    #Split each element by semi-comma
    $arr = $Plans.Split(";")
    return $arr
}

Function ManageLicense(){
    param(
        $Group,
        $SKU,
        $Plans
    )
    Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Starting group membership changes task"
    
    #Check if any change happens in the license config file
    $plansMonitorResult = LicensePlansMonitor -LicenseFile $licenseConfigFile -Group $Group

    #Format plans to add and/or remove and a list with all plans included (in order to remove users that left the plan)
    #If the eobject confirms that a plan has been included, adds '=>' in an exclusive variable
    if(($plansMonitorResult | Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count -gt 0){
        $plansToEnable = ListPlans -Plans ($plansMonitorResult | Where-Object{$_.SideIndicator -eq "=>"}).Plan
    }
    #If the eobject confirms that a plan has been excluded, adds '=>' in an exclusive variable
    if(($plansMonitorResult | Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).count -gt 0){
        $plansToDisable = ListPlans -Plans ($plansMonitorResult | Where-Object{$_.SideIndicator -eq "<="}).Plan
    }
    #Creates an exclusive variable with all plans
    if(($plansMonitorResult|Measure-Object).Count -gt 0){
        $allPlans = ListPlans -Plans ($plansMonitorResult).Plan
    }
    
    #If any change happend in plans config, run the new configuration for all members of the group. If a change happend in group's membership, disables the license options of the given plan from those whom left the group. Add and remove plans for those whom is in the group
    if(($plansMonitorResult|Measure-Object).Count -gt 0){

        #Creates a varibale with the group members ship activitiy
        $groupMembersChange = GroupMonitor -GroupName $Group

        #If the number of group changes is greater than 0, runs the changes for those users that are in the group and disable the license options of the given plan from those whom left the group
        if(($groupMembersChange|Measure-Object).Count -gt 0){

            Write-Log -LogLevel Info -UserOrGroup $Group -Message "Start license adition task for $(($groupMembersChange|Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count) new members"
            Write-Log -LogLevel Info -UserOrGroup $Group -Message "SKU: $($SKU). Plano(s): $($Plans)"

            #Get group members and save in a new file and runs new licensing for all members
            try{
                #Get a list of current group membership
                $currentMembers = Get-ADGroup -Identity $Group -Properties Members -ErrorAction Stop | Select-Object -ExpandProperty Members | Get-ADUser -ErrorAction Stop
                Write-Log -LogLevel "Info" -UserOrGroup $Group "Group members listed successfully"
                
                #If the number of members found is greater than 0
                if(($currentMembers|Measure-Object).Count -gt 0){
                    try{
                        #Creates a variable with the name of the group membership log file
                        $logFileName = "$($groupName)_Members_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
                        $currentMembers | Export-Csv "$($logFileName)" -NoTypeInformation -ErrorAction Stop
                        Write-Log -LogLevel "Error" -UserOrGroup $Group "Group membership saved into log file $($logFileName)"
                        
                        #For each user listed as an AD member
                        foreach($user in $currentMembers){

                            #Checks if the user is not licensed yet
                            If(-not (Get-MsolUser -UserPrincipalName $user.UserPrincipalName).IsLicensed -eq $true){
                                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "User not licensed. SKU is going to be added"
                                try{
                                    #If there is any plan to manage
                                    if(($plansToEnable|Measure-Object).count -gt 0){
                                        #Enables the license plan for the current user
                                        Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added to user with the following plans enabled: $($plansToEnable)"
                                    }
                                    #If there is any plan to disable
                                    if(($plansToDisable|Measure-Object).count -gt 0){
                                        #Disables the license plan for the current user
                                        Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added to user with the follwoing plans disabled: $($plansToDisable)"
                                    }
                                }
                                catch{
                                    #Writes the error related to plans changes to the log
                                    Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                }
                            }
                            #If the user is already licensed
                            else{
                                #Checks if the user has the current SKU ID already (passed as $SKU parameter into this current function)
                                If(-not((Get-MsolUser -UserPrincipalName $user.UserPrincipalName).Licenses.AccountSkuId|Where-Object{$_ -eq $SKU})){
                                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "User already has SKU $($SKU). Plans going to be edited"
                                    try{
                                        #If there is any plan to enable
                                        if(($plansToEnable|Measure-Object).count -gt 0){
                                            #Enables the license plan for the current user
                                            Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Plans enabled for the user: $($plansToEnable)"
                                        }
                                        #If there is any plan to disable
                                        if(($plansToDisable|Measure-Object).count -gt 0){
                                            #Disables the license plan for the current user
                                            Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Plans disabled for the user: $($plansToDisable)"
                                        }                            
                                    }
                                    catch{
                                        #Writes the error related to plans changes to the log
                                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                    }
                                }
                                #If the user has been already licensed and also has the current SKU (passed as $SKU parameter into this current function)
                                else{
                                    try{
                                        #If there is any plan to manage
                                        if(($plansToEnable|Measure-Object).count -gt 0){
                                            #Updates the license plan to add new plan configuration for the current user
                                            Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Plans added to the user: $($plansToEnable)"
                                        }
                                        #If there is any plan to disable
                                        if(($plansToDisable|Measure-Object).count -gt 0){
                                            #Updates the license plan to remove new plan configuration for the current user
                                            Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Plans disabled for the user: $($plansToDisable)"
                                        }

                                    }
                                    catch{
                                        #Writes the error related to plans changes to the log
                                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                    }
                                }
                            }
                        }
                    }
                    catch{
                        #Writes the error related to export current members
                        Write-Log -LogLevel "Error" -UserOrGroup $Group $_.Exception.Message
                    }
                }
                #If the number of members found is 0
                else{
                    Write-Log -LogLevel "Error" -UserOrGroup $Group "No member found in the group. No license management will run for this group"
                }
            }
            catch{
                #Writes the error related to get AD user objects
                Write-Log -LogLevel "Error" -UserOrGroup $groupName $_.Exception.Message
            }
            
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Start license removal task for $(($groupMembersChange|Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).Count) removed members"
            
            #If any user has been removed from the group
            if(($groupMembersChange | Where-Object{$_.SideIndicator -eq "<="} | Measure-Object).Count -gt 0){

                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "$(($groupMembersChange | Where-Object{$_.SideIndicator -eq "<="} | Measure-Object).Count) users have been removed from the group"

                #For each user whom left the group, remove the plans alls plans form last license config file edition
                foreach($user in $groupMembersChange | Where-Object{$_.SideIndicator -eq "<="}){
                    
                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "User has been removed from the group. The plans will be disabled for this user"

                    try{
                        #Remove the plan from the user
                        Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $allPlans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Plans disabled for the user: $($plansToDisable)"
                    }
                    catch{
                        #Writes the error related to license management
                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                    }
                }
            }
            #If no user has been removed from group
            else{
                Write-Log -LogLevel Error -UserOrGroup $Group -Message "No user has been removed. No license removal will run for this group"
            }
        }
        #If the number of changes in group membership is equal to 0 starts the license management on all users due to change in license config file
        else {
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "NO member added or removed on this group. No license management wil run for this group"
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Starting group membership listing"
            
            try{
                $currentMembers = Get-ADGroup -Identity $Group -Properties Members -ErrorAction Stop | Select-Object -ExpandProperty Members | Get-ADUser -ErrorAction Stop
                Write-Log -LogLevel "Info" -UserOrGroup $Group "Members listed successfully"
                #If current membership count is greater than 0
                if(($currentMembers|Measure-Object).Count -gt 0){
                    #Creates a variable with the name of the group membership log file
                    try{
                        $logFileName =  "$($groupName)_Members_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
                        $currentMembers | Export-Csv "$($logFIleName)" -NoTypeInformation -ErrorAction Stop
                        Write-Log -LogLevel "Error" -UserOrGroup $Group "Group membership exported into log file $($logFileName)"

                        foreach($user in $currentMembers){

                           #Checks if the user is not licensed yet
                            If(-not (Get-MsolUser -UserPrincipalName $user.UserPrincipalName).IsLicensed -eq $true){
                                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "User is not licensed. SKU is going to be added for this user"
                                try{
                                    #If there is a plan to enable                           
                                    if(($plansToEnable|Measure-Object).count -gt 0){
                                        #Enables the license plan for the current user
                                        Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added with the following plans enabled $($plansToEnable)"
                                    }
                                    #If there is a plan to disable    
                                    if(($plansToDisable|Measure-Object).count -gt 0){
                                        #Disables the license plan for the current user
                                        Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added with the following plans removed: $($plansToDisable)"
                                    }
                                }
                                catch{
                                   #Writes the error related to plans changes to the log
                                    Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                }
                            }
                           #If the user has been already licensed
                            else{
                                 #Checks if the user has the current SKU ID already (passed as $SKU parameter into this current function)
                                If(-not((Get-MsolUser -UserPrincipalName $user.UserPrincipalName).Licenses.AccountSkuId|Where-Object{$_ -eq $SKU})){
                                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "User is not licensed SKU $($SKU). SKU is going to be added for this user and plans will be edited"
                                    try{
                                        #If there is a plan to enable
                                        if(($plansToEnable|Measure-Object).count -gt 0){
                                            #Enables the license plan for the current user
                                            Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added with the following plans enabled: $($plansToEnable)"
                                        }
                                        #If there is a plan to disable
                                        if(($plansToDisable|Measure-Object).count -gt 0){
                                            #Disables the license plan for the current user
                                            Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added with the following plans disabled $($plansToDisable)"
                                        }                            }
                                    catch{
                                        #Writes the error related to plans changes to the log
                                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                    }
                                }
                                #If the user has been already licensed and also has the current SKU (passed as $SKU parameter into this current function)
                                else{
                                    try{
                                       #If there is a plan to enable
                                        if(($plansToEnable|Measure-Object).count -gt 0){
                                            #Updates the license plan to add new plan configuration for the current user
                                            Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Plans removed for the user: $($plansToDisable)"
                                        }
                                        #If there is a plan to disable
                                        if(($plansToDisable|Measure-Object).count -gt 0){
                                            #Updates the license plan to add new plan configuration for the current user
                                            Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Plans removed for the user: $($plansToDisable)"
                                        }

                                    }
                                    catch{
                                        #Writes the error related to plans changes to the log
                                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                    }
                                }
                            }
                        }
                    }
                    catch{
                        #Writes the error related to export current members
                        Write-Log -LogLevel "Error" -UserOrGroup $Group $_.Exception.Message
                    }
                }
                else{
                    #If the number of members found is 0
                    Write-Log -LogLevel "Error" -UserOrGroup $Group "No group member found"
                }
            }
            catch{
                #Writes the error related to get AD user objects
                Write-Log -LogLevel "Error" -UserOrGroup $groupName $_.Exception.Message
            }
        }
    }
    #If no changes has been detected in license configuration file, runs license management only for those users that have been either added or removed from the group    
    else{
        #Creates a varibale with membership changes
        $groupMembersChange = GroupMonitor -GroupName $Group

        #If the number of changes in group membership is greater than 0
        if(($groupMembersChange|Measure-Object).Count -gt 0){

            Write-Log -LogLevel Info -UserOrGroup $Group -Message "Starting license adition task for $(($groupMembersChange|Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count) new members"
            Write-Log -LogLevel Info -UserOrGroup $Group -Message "SKU: $($SKU). Plan(s): $($Plans)"
            
            #Creates an array with all plans            
            $Plans = ListPlans -Plans $Plans

            #If any user has been added to the group
            if(($groupMembersChange | Where-Object{$_.SideIndicator -eq "=>"} | Measure-Object).Count -gt 0){
                
                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "$(($groupMembersChange | Where-Object{$_.SideIndicator -eq "=>"} | Measure-Object).Count) users have been removed from the group"

                #For each user added to the group, adds the license
                foreach($user in $groupMembersChange | Where-Object{$_.SideIndicator -eq "=>"}){
                    
                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "User $($user) has been removed from the group. Plans are going to be removed for this user"


                   #Checks if the user is not licensed yet
                    If(-not (Get-MsolUser -UserPrincipalName $user.UserPrincipalName).IsLicensed -eq $true){
                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "User not licensed. SKU is going to be added"
                        try{
                            #If there is a plan to enable
                            if(($Plans|Measure-Object).count -gt 0){
                                #Enables the license plan for the current user
                                Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $Plans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added with the following plans enabled: $($plansToEnable)"
                            }
                        }
                        catch{
                            #Writes the error related to plans changes to the log
                            Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                        }
                    }
                    #If the user has been already licensed
                    else{
                        #Checks if the user has the current SKU ID already (passed as $SKU parameter into this current function)
                        If(-not((Get-MsolUser -UserPrincipalName $user.UserPrincipalName).Licenses.AccountSkuId|Where-Object{$_ -eq $SKU})){
                            try{
                                #If there is a plan to enable
                                if(($Plans|Measure-Object).count -gt 0){
                                    #Enables the license plan for the current user
                                    Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $Plans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added with the following plans disabled: $($plansToEnable)"
                                }                         
                            }
                            catch{
                                #Writes the error related to plans changes to the log
                                Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                            }
                        }
                        #If the user has been licensed and has the current SKU ID already (passed as $SKU parameter into this current function)
                        else{
                            try{
                                #If there is a plan to enable
                                if(($Plans|Measure-Object).count -gt 0){
                                    #Enables the license plan for the current user
                                    Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $Plans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added with the following plans enabled: $($plansToEnable)"
                                }

                            }
                            catch{
                                #Writes the error related to plans changes to the log
                                Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                            }
                        }
                    }
                }
            }
           #If no user has been added to the user
            else{
                Write-Log -LogLevel Error -UserOrGroup $Group -Message "No user added to the group to receive a license"
            }
            
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Starting license removal task for $(($groupMembersChange|Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).Count) members removed"
            
            #If any user has been removed from the group
            if(($groupMembersChange | Where-Object{$_.SideIndicator -eq "<="} | Measure-Object).Count -gt 0){

                #for each user whom left the group, remove the license
                foreach($user in $groupMembersChange | Where-Object{$_.SideIndicator -eq "<="}){

                    try{
                        #Disables the license plan for the current user
                        Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $Plans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) added with the following plans disabled: $($plansToEnable)"
                    }
                    catch{
                        #Writes the error related to plans changes to the log
                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                    }
                }
            }
            #If no user has been removed from group
            else{
                Write-Log -LogLevel Error -UserOrGroup $Group -Message "No user has been removed from the group to remove a license"
            }
        }
        #If the number of changes in group membership is 0
        else
        {
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "No group added or removed from the group"
        }
    }
}

#Try to connect to Microsoft Online Services
ConnectMsolService

#For each configuration listed in license config file
foreach($license in $licenseConfigFile){
        #Runs the function to manage licenses        
        ManageLicense -Group $license.Group -SKU $license.SKU -Plans $license.Plans
}

#Get current date/time to stop measure the duration of execution
$stopTimer = Get-Date

#Log the start of script execution
Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "############ Script Finished. Execution Time: $((New-TimeSpan -Start $startTimer -End $stopTimer).ToString("dd\.hh\:mm\:ss")) ############"
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

#Registra o inicio da execução dos script
Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "############ Script Iniciado ############"

#Importa os módulos
try{
    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module MSOnline -ErrorAction Stop
    Import-Module MSOLLicenseManagement -ErrorAction Stop
}
catch{
    Write-Log -LogLevel Error -UserOrGroup "SCRIPT" -Message $_.Exception.Message
    Exit
}

#Inicia o timer para calculo do tempo de execução
$startTimer = Get-Date

#Nome do arquivo de licenciamento
$licenseFilePath = ".\Licenses.csv"

#Importa o módulo com os parametros de licenciamento
$licenseConfigFile = Import-Csv -Path $licenseFilePath -Delimiter ","

#Configura o nome do logFile do modulo de licenciamento
$licenseModuleLogFileName = "LicenseModule_$((Get-Date).ToString('ddMMyyyy')).log"

#Função de gravação de log
Function Write-Log(){
    param
        (
            [ValidateSet("Error", "Warning", "Info")]$LogLevel,
            $UserOrGroup,
            [string]$Message
        )

    #Nome do arquivo de log contendo a data/hora
    $logFileName = "ManageO365Licenses_$((Get-Date).ToString('ddMMyyyy')).log"

    #Header do arquivo no formato csv
    $header = "datetime,user,action,message"

    #Data/hora da entrada no log
    $datetime = (Get-Date).ToString('dd/MM/yyyy hh:mm:ss')

    #Entrada do arquivo de log no formato csv
    $logEntry = "$datetime,$LogLevel,$UserOrGroup,$Message"

    #Se o arquivo não existir, cria o arquivo e adiciona a primeira linha como header
    if(-not(Test-Path $logFileName)){
        try{
            New-Item -Path $logFileName -ErrorAction Stop
            Add-Content -Path $logFileName -Value $header -ErrorAction Stop
        }
        catch{
            $_.Exception.Message
            #Finaliza a exeução do script
            Exit
        }
    }
    #Adiciona a entrada no arquivo de log
    
    try{
        Add-Content -Path $logFileName -Value $logEntry -ErrorAction Stop
    }
    catch{
        $_.Exception.Message
        #Finaliza a execução do script
        Exit
    }
    
}

#Função para coleta das credenciais em disco
Function GetCredentialOnDisk{

    try{
        #Usuário que será usado para se conectar no MSOnline Services
        $username = "admin@foggioncloud.onmicrosoft.com"
        #Obtem a senha do arquivo em disco
        $password = Get-Content .\cred.sec -ErrorAction Stop | ConvertTo-SecureString -ErrorAction Stop

        #Monta o objeto da credencial
        $credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "Credenciais de conexão obtidas. Usuário $($username) usado para a conexão"
        return $credential
    }
    catch{
        Write-Log -LogLevel Error -UserOrGroup "SCRIPT" -Message $_.Exception.Message
    }
}

#Função de conexão no Microsoft Online Services
Function ConnectMsolService{
    try{
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "Importando o módulo MSOnline"
        Import-Module MSOnline -ErrorAction Stop
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "Efetuando a conexão no Microsoft Online Services"
        Connect-MsolService -Credential (GetCredentialOnDisk) -ErrorAction Stop
    }
    catch{
        Write-Log -LogLevel Error -UserOrGroup "SCRIPT" -Message $_.Exception.Message
        
        #Finaliza o timer para calculo do tempo de execucão
        $stopTimer = Get-Date

        #Registra o termino do script com o tempo de execução
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "############ Script Finalizado. Tempo de execução: $((New-TimeSpan -Start $startTimer -End $stopTimer).ToString("dd\.hh\:mm\:ss")) ############"
        
        #Finaliza a exeução do script
        Exit
    }
}

#Função que monitora a entrada e saída de membros de um grupo
Function GroupMonitor{
    
    param(
        $GroupName
    )
    
    $currentMembers = $null
    $compareMembers = $null

    #Obtem membros do grupo e salva um novo arquivo
    Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Iniciando a listagem de membros do grupo"
    try{
        #Obtem a lista de membros atuais
        $currentMembers = Get-ADGroup -Identity $groupName -Properties Members -ErrorAction Stop | Select-Object -ExpandProperty Members | Get-ADUser -ErrorAction Stop
        Write-Log -LogLevel "Info" -UserOrGroup $groupName "Membros do grupo listados com sucesso"

        #Se a quantidade de membros for maior do que 0
        if(($currentMembers|Measure-Object).Count -gt 0){
            try{
                #Grava uma variável com nome do arquivo contendo data/hora
                $logFileName =  "$($groupName)_Membros_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
                #Exporta para um arquivo .csv
                $currentMembers | Export-Csv "$($logFileName)" -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel "Error" -UserOrGroup $groupName "Membros do grupo salvos no arquivo csv $($logFileName)"
            }
            catch{
                Write-Log -LogLevel "Error" -UserOrGroup $groupName $_.Exception.Message
            }
        }
        #Se a quantidade de membros for menor do que 1
        else{
            Write-Log -LogLevel "Error" -UserOrGroup $groupName "Nenhum membro encontrado no grupo"
        }
    }
    catch{
        Write-Log -LogLevel "Error" -UserOrGroup $groupName $_.Exception.Message
    }

    #Grava uma variavel com o nome do arquivo de log
    $logFileName = "$($groupName)_MembrosAdicionadosRemovidos_ddMMyyyy_hhmmss.csv"

    #Se a quantidade de arquivos de membros for maior do que 1
    If((Get-Item "$($groupName)_Membros_*" |Measure-Object).Count -gt 1){
        #Compara os membros atuais no novo arquivo com a lista criada na última execução
        try{
            #Efetua a comparação dos membros do grupo usando os dois arquivos mais recentes, sendo o ultima da execução atual e o penultimo da execução anterior
            $compareMembers = Compare-Object -DifferenceObject (Import-Csv (Get-Item "$($groupName)_Membros_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1)) -ReferenceObject (Import-Csv (Get-Item "$($groupName)_Membros_*" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 1 -First 1)) -PassThru -Property Name -ErrorAction Stop
            Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Comparação de usuários adicionados ou removidos efetuada com sucesso. Lista de usuários presente no arquivo $($logFileName)"
        }
        catch{
            Write-Log -LogLevel "Error" -UserOrGroup $groupName -Message $_.Exception.Message
        }
    }
    #Se a quantidade de arquivos de membros for menor ou igual à 1, entra em modo de primeira execução
    else{
        Write-Log -LogLevel Error -UserOrGroup $groupName -Message "Não é possivel efetuar a comparação. Primeira execução ou arquivos com a lista de membros faltando. Se for a primeira execução, todos os usuários do grupo receberão a licença"

        #Se a quantidade de itens resultantes da comparação de membros for maior do que 0
        if(($currentMembers|Measure-Object).Count -gt 0){

            #Adiciona a coluna que indica que o membro é novo no grupo para todos os membros, pois o script entrou em modo de primeira execução
            $currentMembers | Add-Member -Name "SideIndicator" -MemberType NoteProperty -Value "=>" -Force
            
            #Exporta uma lista contendo os membros que foram adicionados ou removidos do grupo
            try{
                $currentMembers | Export-Csv "$($logFIleName)" -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Lista de membros adicionados e removidos do grupo exportada com sucesso para o arquivo $($logFileName)"
            }
            catch{
                Write-Log -LogLevel "Error" -UserOrGroup $groupName -Message $_.Exception.Message
            }

            #Retorna a lista atual de membros como o resultado da comparação, pois o script entrou em modo de primeira execução
            return $currentMembers
        }
        #Se a quantidade de itens resultantes da comparação de membros for igual a 0
        else
        {
            Write-Log -LogLevel "Error" -UserOrGroup $groupName -Message "Não foi possivel obter os membros atuais do grupo"
        }
    }

    #Se a quantidadede itens resultantes da comparação de membros for maior do que 0 (agora fora do modo de primeira execução)
    if(($compareMembers|Measure-Object).Count -gt 0){
            Write-Log -LogLevel "Info" -UserOrGroup $group -Message "Alterações detectadas. Membros adicionados: $(($compareMembers|Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count) - Membros removidos: $(($compareMembers|Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).Count)"
            
            #Exporta uma lista contendo os membros que foram adicionados ou removidos do grupo
            try{
                $compareMembers | Export-Csv "$($logFIleName)" -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Lista de membros adicionados e removidos do grupo exportada com sucesso para o arquivo $($logFIleName)"
            }
            catch{
                Write-Log -LogLevel "Error" -UserOrGroup $groupName -Message $_.Exception.Message
            }

            #Cria uma array para conter os objetos de usuários do AD. Não é feito o retorno do objeto de comparação diretamente devido à problemas na leitura de objetos dentro de objetos
            $arrCompareMembers = @()

            #Para cada objeto dentro da matriz de comparação de membros
            foreach($userMember in $compareMembers){
                #Obtem o objeto do usuário no AD
                $user = Get-ADUser -Identity $userMember.DistinguishedName

                #Se o objeto contem o indicador de membro adicionado, adiciona uma coluna ao objeto usuário com o indicador
                if($userMember.SideIndicator -eq "=>"){
                    $user | Add-Member -Name SideIndicator -MemberType NoteProperty -Value "=>" -Force
                }
                #Se o objeto contem o indicador de membro removido, adiciona uma coluna ao objeto usuário com o indicador
                if($userMember.SideIndicator -eq "<="){
                    $user | Add-Member -Name SideIndicator -MemberType NoteProperty -Value "<=" -Force
                }
                #Adiciona o objeto do usuário do AD na array resultante
                $arrCompareMembers += $user
            }

            #Retorna a lista comparativa de membros
            return $arrCompareMembers
     }
}

#Função que monitora o arquivo de licenças
Function LicensePlansMonitor{
    
    Param(
        $LicenseFile,
        $Group
    )

    #Cria uma array para retorno do resultado da função
    $arrLicenseChangeResult = @()

    #Para cada licença no arquivo csv de configuração de licenças
    foreach($license in $LicenseFile|Where-Object{$_.Group -eq $Group}){
    
        #Limpa a variável de comparação para evitar sujeira
        $comparePlans = $null

        #Cria o sufixo do nome do arquivo, trocando : por - para evitar problemas na criação do arquivo (caractere : no path do arquivo indicaria drive inexistente)
        $plansCompareFileSufix = "$($license.Group)_$(($license.SKU).Replace(":","-"))"
        
        #Obtem planos da licença atual
        Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "Iniciando a listagem de planos da configuração de liçenca $($license.SKU)"
        #$currentPlans = ListPlans -Plans $license.Plans
        Write-Log -LogLevel "Info" -UserOrGroup $license.SKU "Planos do grupo listados: $($license.Plans)"

        #Se a coluna Plans da licença não estiver vazia
        if(![string]::IsNullOrEmpty($license.Plans)){

            #Armazena o nome do arquivo com a lista de planos
            $plansFileName = "$($plansCompareFileSufix)_Planos_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
            try{
                #Cria uma array resultante
                $arrPlans = @()
                #Para cada plano na coluna, separados por ;
                foreach($plan in ($license.Plans.Split(";"))){
                    #Cria um objeto ps custom e adiciona as colunas para listar os planos por linha no arquivo de planos
                    $objPlans = New-Object psobject
                    $objPlans | Add-Member -Name "Group" -MemberType NoteProperty -Value $license.Group
                    $objPlans | Add-Member -Name "SKU" -MemberType NoteProperty -Value $license.SKU
                    $objPlans | Add-Member -Name "Plan" -MemberType NoteProperty -Value $plan
                    #Acumula a variável resultante
                    $arrPlans += $objPlans
                }
                
                #Exporta a lista de planos da licença em um arquivo csv
                $arrPlans | Export-Csv $plansFileName -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel Info -UserOrGroup $license.SKU "Planos da SKU salvos no arquivo csv $($plansFileName)"
            }
            catch{
                Write-Log -LogLevel Error -UserOrGroup $license.SKU $_.Exception.Message
            }
        }
        #Se a coluna Plans da licença estiver vazia (gravação de um objeto com coluna vazia em separado para tratamento de erro separado)
        else{
            #Armazena o nome do arquivo com a lista de planos
            $plansFileName = "$($plansCompareFileSufix)_Planos_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
            try{
                #Cria um objeto ps custom com a coluna plan vazia
                $objPlans = New-Object psobject
                $objPlans | Add-Member -Name "Group" -MemberType NoteProperty -Value $license.Group
                $objPlans | Add-Member -Name "SKU" -MemberType NoteProperty -Value $license.SKU
                $objPlans | Add-Member -Name "Plan" -MemberType NoteProperty -Value $null
                $objPlans | Export-Csv $plansFileName -NoTypeInformation -ErrorAction Stop
                Write-Log -LogLevel Error -UserOrGroup $groupName "Nenhum plano encontrado na SKU $($license.SKU)"
            }
            catch{
                Write-Log -LogLevel Error -UserOrGroup $license.SKU $_.Exception.Message
            }
        }

        #Se a quantidade de arquivos de planos da licença for maior do que 1
        If((Get-Item "$($plansCompareFileSufix)_Planos_*" |Measure-Object).Count -gt 1){
            #Armazena o nome do arquivo com a lista de planos
            $plansCompareFileName = "$($plansCompareFileSufix)_PlanosAdicionadosRemovidos_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
            #Compara os planos atuais no novo arquivo com a lista criada na última execução
            try{
                #$lastRunPlan = ListPlans -Plans (Import-Csv (Get-Item "$($plansCompareFileSufix)_Planos_*" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 1 -First 1)).Plans
                $comparePlans = Compare-Object -DifferenceObject (Import-Csv (Get-Item "$($plansCompareFileSufix)_Planos_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1)) -ReferenceObject (Import-Csv (Get-Item "$($plansCompareFileSufix)_Planos_*" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 1 -First 1)) -PassThru -Property Plan -ErrorAction Stop
                Write-Log -LogLevel Info -UserOrGroup $license.SKU -Message "Comparação de planos adicionados ou removidos efetuada com sucesso. Lista de planos presente no arquivo $($plansCompareFileName)"
            }
            catch{
                Write-Log -LogLevel "Error" -UserOrGroup $license.SKU -Message $_.Exception.Message
            }
        }
        #Se a quantidade de arquivos de planos da licença for igual ou menor à 1
        else{
            #Armazena o nome do arquivo com a lista de planos
            $plansCompareFileName = "$($plansCompareFileSufix)_PlanosAdicionadosRemovidos_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
            Write-Log -LogLevel Error -UserOrGroup $license.SKU -Message "Não é possivel efetuar a comparação. Primeira execução ou arquivos com a lista de planos faltando. Se for a primeira execução, todos os usuários do grupo receberão a licença"

            #Se a coluna Plans da licença não estiver vazia
            if(![string]::IsNullOrEmpty($license.Plans)){
        
                #Cria uma array resultante
                $arrPlans = @()
                foreach($plan in ($license.Plans.Split(";"))){
                    #Cria um objeto ps custom com a coluna plan vazia
                    $objPlans = New-Object psobject
                    $objPlans | Add-Member -Name "Group" -MemberType NoteProperty -Value $license.Group
                    $objPlans | Add-Member -Name "SKU" -MemberType NoteProperty -Value $license.SKU
                    $objPlans | Add-Member -Name "Plan" -MemberType NoteProperty  -Value $plan
                    $objPlans | Add-Member -Name "SideIndicator" -MemberType NoteProperty -Value "=>"
                    $arrPlans += $objPlans
                }
                
                #Exporta uma lista contendo os planos que foram adicionados ou removidos do grupo
                try{
                    $arrPlans | Where-Object{$_.Plan -ne ""} | Export-Csv $plansCompareFileName -NoTypeInformation -ErrorAction Stop
                    Write-Log -LogLevel "Info" -UserOrGroup $license.SKU -Message "Lista de membros adicionados e removidos do grupo exportada com sucesso para o arquivo $($plansCompareFileName)"
                }
                catch{
                    Write-Log -LogLevel "Error" -UserOrGroup $license.SKU -Message $_.Exception.Message
                }
                #Acumula a matriz resultante
                $arrLicenseChangeResult += $arrPlans
            }
            #Se a coluna Plans da licença estiver vazia (gravação de um objeto com coluna vazia em separado para tratamento de erro separado)
            else
            {
                #Armazena o nome do arquivo com a lista de planos
                $plansCompareFileName = "$($plansCompareFileSufix)_Planos_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"

                try{
                    #Cria um objeto ps custom com a coluna plan vazia
                    $objPlans = New-Object psobject
                    $objPlans | Add-Member -Name "Group" -MemberType NoteProperty -Value $license.Group
                    $objPlans | Add-Member -Name "SKU" -MemberType NoteProperty -Value $license.SKU
                    $objPlans | Add-Member -Name "Plan" -MemberType NoteProperty -Value $null
                    $objPlans | Export-Csv $plansFileName -NoTypeInformation -ErrorAction Stop
                    Write-Log -LogLevel Error -UserOrGroup $groupName "Nenhum plano encontrado na SKU $($license.SKU)"
                }
                catch{
                    Write-Log -LogLevel Error -UserOrGroup $license.SKU $_.Exception.Message
                }
            }
        }

        if(($comparePlans|Measure-Object).Count -gt 0){
                Write-Log -LogLevel "Info" -UserOrGroup $license.SKU -Message "Alterações detectadas. Planos adicionados: $(($comparePlans|Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count) - Planos removidos: $(($comparePlans|Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).Count)"

                #Exporta uma lista contendo os planos que foram adicionados ou removidos do grupo
                try{
                    $plansCompareFileName = "$($plansCompareFileSufix)_PlanosAdicionadosRemovidos_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
                    $comparePlans | Where-Object{$_.Plan -ne ""} | Export-Csv $plansCompareFileName -NoTypeInformation -ErrorAction Stop
                    Write-Log -LogLevel "Info" -UserOrGroup $groupName -Message "Lista de planos adicionados e removidos do grupo exportada com sucesso para o arquivo $($plansCompareFileName)"
                }
                catch{
                    Write-Log -LogLevel "Error" -UserOrGroup $license.SKU -Message $_.Exception.Message
                }
            
                #Retorna a lista comparativa de planos
                $arrLicenseChangeResult += $comparePlans
        }
    }

    #Retorna a lista de planos resultantes caso a coluna plan não esteja vazia 
    return $arrLicenseChangeResult | Where-Object{$_.Plan -ne ""}
}

#Função que cria uma array com os planos separados por ; da coluna Plans do arquivo Licenses.csv
Function ListPlans{
    param(
        $Plans
    )

    #Separa cada element por ; e cria uma array
    $arr = $Plans.Split(";")
    return $arr
}

Function ManageLicense(){
    param(
        $Group,
        $SKU,
        $Plans
    )
    Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Iniciando a tarefa de verificação de alterações de membros no grupo"
    
    #Verifica se ocorreu alguma mudança no arquivo de licenças
    $plansMonitorResult = LicensePlansMonitor -LicenseFile $licenseConfigFile -Group $Group

    #Formata os planos para adicionar, remover e uma lista com todos os planos (para remover dos usuários que sairam do grupo)
    #Se o objeto indicar que o plano foi adicionado => grava numa variavel separada
    if(($plansMonitorResult | Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count -gt 0){
        $plansToEnable = ListPlans -Plans ($plansMonitorResult | Where-Object{$_.SideIndicator -eq "=>"}).Plan
    }
    #Se o objeto indicar que o plano foi removido <= grava numa variavel separada
    if(($plansMonitorResult | Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).count -gt 0){
        $plansToDisable = ListPlans -Plans ($plansMonitorResult | Where-Object{$_.SideIndicator -eq "<="}).Plan
    }
    #Grava uma variavel com todos os planos
    if(($plansMonitorResult|Measure-Object).Count -gt 0){
        $allPlans = ListPlans -Plans ($plansMonitorResult).Plan
    }
                
    #Se ocorreu alguma mudança na configuração de licenças, roda a nova configuração para todos do grupo. Se houve mudança no membership do grupo, remove todos os planos da configuração de quem saiu. Adiciona e remove os planos para quem esta no grupo
    if(($plansMonitorResult|Measure-Object).Count -gt 0){

        #Lista a atividade de entrada e saida do grupo
        $groupMembersChange = GroupMonitor -GroupName $Group

        #Se a quantidade de mudanças no grupo for maior do que 0 roda as mudanças para os usuários que estão no grupo e remove as licenças adicionadas e removidas da configuração de quem saiu do grupo
        if(($groupMembersChange|Measure-Object).Count -gt 0){

            Write-Log -LogLevel Info -UserOrGroup $Group -Message "Iniciando tarefas de adição de licenças para $(($groupMembersChange|Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count) novos membros"
            Write-Log -LogLevel Info -UserOrGroup $Group -Message "SKU: $($SKU). Plano(s): $($Plans)"

            #Obtem membros do grupo e salva um novo arquivo e roda o novo licenciamento para todos os mem
            try{
                #Lista os membros atuais do grupo
                $currentMembers = Get-ADGroup -Identity $Group -Properties Members -ErrorAction Stop | Select-Object -ExpandProperty Members | Get-ADUser -ErrorAction Stop
                Write-Log -LogLevel "Info" -UserOrGroup $Group "Membros do grupo listados com sucesso"
                #Se a quantidade de membros encontrados for maior do que 0
                if(($currentMembers|Measure-Object).Count -gt 0){
                    try{
                        #Cria uma variável com o nome do arquivo de log
                        $logFileName = "$($groupName)_Membros_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
                        $currentMembers | Export-Csv "$($logFileName)" -NoTypeInformation -ErrorAction Stop
                        Write-Log -LogLevel "Error" -UserOrGroup $Group "Membros do grupo salvos no arquivo csv $($logFileName)"
                        
                        #Para cada usuário listado como membro do AD
                        foreach($user in $currentMembers){

                            #Verifica se o usuário não esta licenciado
                            If(-not (Get-MsolUser -UserPrincipalName $user.UserPrincipalName).IsLicensed -eq $true){
                                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Usuário não licenciado. SKU será adicionada"
                                try{
                                    #Se existe algum plano para habilitar
                                    if(($plansToEnable|Measure-Object).count -gt 0){
                                        #Adiciona a licença ao usuário
                                        Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos $($plansToEnable) habilitados"
                                    }
                                    #Se existe algum plano para desabilitar
                                    if(($plansToDisable|Measure-Object).count -gt 0){
                                        #Adiciona a licença ao usuário
                                        Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos $($plansToDisable) removidos"
                                    }
                                }
                                catch{
                                    #Grava o erro no log
                                    Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                }
                            }
                            #Se o usuário já foi licenciado
                            else{
                                #Verifica se o usuário não contém o SKU ID atual
                                If(-not((Get-MsolUser -UserPrincipalName $user.UserPrincipalName).Licenses.AccountSkuId|Where-Object{$_ -eq $SKU})){
                                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Usuário já possui SKU $($SKU). Planos serão editados"
                                    try{
                                        #Se existe algum plano para habilitar
                                        if(($plansToEnable|Measure-Object).count -gt 0){
                                            #Adiciona a licença ao usuário
                                            Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Planos habilitados para o usuário: $($plansToEnable)"
                                        }
                                        #Se existe algum plano para desabilitar
                                        if(($plansToDisable|Measure-Object).count -gt 0){
                                            #Adiciona a licença ao usuário
                                            Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Planos removidos para o usuário: $($plansToDisable)"
                                        }                            
                                    }
                                    catch{
                                        #Grava o erro no log
                                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                    }
                                }
                                #Se o usuário foi licenciado e já tem o SKU ID, atualiza os planos
                                else{
                                    try{
                                        if(($plansToEnable|Measure-Object).count -gt 0){
                                            Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Planos adicionados para o usuário: $($plansToEnable)"
                                        }
                                        if(($plansToDisable|Measure-Object).count -gt 0){
                                            Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Planos removidos para o usuário: $($plansToDisable)"
                                        }

                                    }
                                    catch{
                                        #Grava o erro no log
                                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                    }
                                }
                            }
                        }
                    }
                    catch{
                        Write-Log -LogLevel "Error" -UserOrGroup $Group $_.Exception.Message
                    }
                }
                else{
                    Write-Log -LogLevel "Error" -UserOrGroup $Group "Nenhum membro encontrado no grupo"
                }
            }
            catch{
                Write-Log -LogLevel "Error" -UserOrGroup $groupName $_.Exception.Message
            }
            
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Iniciando tarefas de remoção de licenças para $(($groupMembersChange|Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).Count) membros removidos"
            
            #Se algum usuário foi removido do grupo
            if(($groupMembersChange | Where-Object{$_.SideIndicator -eq "<="} | Measure-Object).Count -gt 0){

                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "$(($groupMembersChange | Where-Object{$_.SideIndicator -eq "<="} | Measure-Object).Count) usuários foram removidos do grupo"

                #Para cada usuário que saiu do grupo, remove os planos adicionados e removidos após a edição da configuração de licenças
                foreach($user in $groupMembersChange | Where-Object{$_.SideIndicator -eq "<="}){
                    
                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Ö usuário $($user) foi removido do grupo. Os planos do grupo serão removidos para o usuário"

                    try{
                        #Remove a plano do usuário
                        Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $allPlans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Planos removidos para o usuário: $($plansToDisable)"
                    }
                    catch{
                        #Grava o erro no log
                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                    }
                }
            }
            #Se nenhum usuário removido do grupo for encontrado
            else{
                Write-Log -LogLevel Error -UserOrGroup $Group -Message "Nenhum usuário removido do grupo para remover licença"
            }
        }
        #Se a quantidade de mudanças do grupo for igual a 0 inicia a edição de licenças em todos os usuários do grupo devido à mudança no arquivo de configuração
        else {
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Nenhum membro foi adicionado ou removido do grupo"
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Iniciando a listagem de membros do grupo"
            
            #Obtem membros do grupo e salva um novo arquivo
            try{
                $currentMembers = Get-ADGroup -Identity $Group -Properties Members -ErrorAction Stop | Select-Object -ExpandProperty Members | Get-ADUser -ErrorAction Stop
                Write-Log -LogLevel "Info" -UserOrGroup $Group "Membros do grupo listados com sucesso"
                if(($currentMembers|Measure-Object).Count -gt 0){
                    try{
                        $logFileName =  "$($groupName)_Membros_$((Get-Date).ToString('ddMMyyyy_hhmmss')).csv"
                        $currentMembers | Export-Csv "$($logFIleName)" -NoTypeInformation -ErrorAction Stop
                        Write-Log -LogLevel "Error" -UserOrGroup $Group "Membros do grupo salvos no arquivo csv $($logFileName)"

                        foreach($user in $currentMembers){

                            #Verifica se o usuário não esta licenciado
                            If(-not (Get-MsolUser -UserPrincipalName $user.UserPrincipalName).IsLicensed -eq $true){
                                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Usuário não licenciado. SKU será adicionada"
                                try{
                                    #Se existe algum plano para habilitar
                                    if(($plansToEnable|Measure-Object).count -gt 0){
                                        #Adiciona a licença ao usuário
                                        Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos $($plansToEnable) habilitados"
                                    }
                                    #Se existe algum plano para desabilitar
                                    if(($plansToDisable|Measure-Object).count -gt 0){
                                        #Adiciona a licença ao usuário
                                        Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos $($plansToDisable) removidos"
                                    }
                                }
                                catch{
                                    #Grava o erro no log
                                    Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                }
                            }
                            #Se o usuário já foi licenciado
                            else{
                                #Verifica se o usuário não contém o SKU ID atual
                                If(-not((Get-MsolUser -UserPrincipalName $user.UserPrincipalName).Licenses.AccountSkuId|Where-Object{$_ -eq $SKU})){
                                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Usuário não possui SKU $($SKU). SKU será adicionadao e os planos serão editados"
                                    try{
                                        #Se existe algum plano para habilitar
                                        if(($plansToEnable|Measure-Object).count -gt 0){
                                            #Adiciona a licença ao usuário
                                            Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos $($plansToEnable) habilitados"
                                        }
                                        #Se existe algum plano para desabilitar
                                        if(($plansToDisable|Measure-Object).count -gt 0){
                                            #Adiciona a licença ao usuário
                                            Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos $($plansToDisable) removidos"
                                        }                            }
                                    catch{
                                        #Grava o erro no log
                                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                    }
                                }
                                #Se o usuário foi licenciado e já tem o SKU ID, atualiza os planos
                                else{
                                    try{
                                        #Se a quantidade de planos para habilitar for maior do que 0
                                        if(($plansToEnable|Measure-Object).count -gt 0){
                                            Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $plansToEnable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Planos removidos para o usuário: $($plansToDisable)"
                                        }
                                        #Se a quantidade de planos para desabilitar for maior do que 0
                                        if(($plansToDisable|Measure-Object).count -gt 0){
                                            Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $plansToDisable -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                            Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Planos removidos para o usuário: $($plansToDisable)"
                                        }

                                    }
                                    catch{
                                        #Grava o erro no log
                                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                                    }
                                }
                            }
                        }
                    }
                    catch{
                        Write-Log -LogLevel "Error" -UserOrGroup $Group $_.Exception.Message
                    }
                }
                else{
                    Write-Log -LogLevel "Error" -UserOrGroup $Group "Nenhum membro encontrado no grupo"
                }
            }
            catch{
                Write-Log -LogLevel "Error" -UserOrGroup $groupName $_.Exception.Message
            }
        }
    }
    #Se não ocorreu nenhuma mudança no arquivo de configuração de licenças, executa apenas para usuários que foram adicionados ou removidos dos grupos
    else{
        #Lista a atividade de entrada e saida do grupo
        $groupMembersChange = GroupMonitor -GroupName $Group

        #Se a quantidade de mudanças no grupo for maior do que 0
        if(($groupMembersChange|Measure-Object).Count -gt 0){

            Write-Log -LogLevel Info -UserOrGroup $Group -Message "Iniciando tarefas de adição de licenças para $(($groupMembersChange|Where-Object{$_.SideIndicator -eq "=>"}|Measure-Object).Count) novos membros"
            Write-Log -LogLevel Info -UserOrGroup $Group -Message "SKU: $($SKU). Plano(s): $($Plans)"
            
            #Cria a array de planos
            $Plans = ListPlans -Plans $Plans

            #Se algum usuário foi adicionado no grupo
            if(($groupMembersChange | Where-Object{$_.SideIndicator -eq "=>"} | Measure-Object).Count -gt 0){
                
                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "$(($groupMembersChange | Where-Object{$_.SideIndicator -eq "=>"} | Measure-Object).Count) usuários foram removidos do grupo"

                #Para cada usuário que entrou no grupo, adiciona a licença correspondente
                foreach($user in $groupMembersChange | Where-Object{$_.SideIndicator -eq "=>"}){
                    
                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "O usuário $($user) foi removido do grupo. Os planos do grupo serão removidos para o usuário"

                    #Verifica se o usuário não esta licenciado
                    If(-not (Get-MsolUser -UserPrincipalName $user.UserPrincipalName).IsLicensed -eq $true){
                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "Usuário não licenciado. SKU será adicionada"
                        try{
                            #Se existe algum plano para habilitar
                            if(($Plans|Measure-Object).count -gt 0){
                                #Adiciona a licença ao usuário
                                Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $Plans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos habilitados: $($plansToEnable)"
                            }
                        }
                        catch{
                            #Grava o erro no log
                            Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                        }
                    }
                    #Se o usuário já foi licenciado
                    else{
                        #Verifica se o usuário não contém o SKU ID atual
                        If(-not((Get-MsolUser -UserPrincipalName $user.UserPrincipalName).Licenses.AccountSkuId|Where-Object{$_ -eq $SKU})){
                            try{
                                #Se existe algum plano para habilitar
                                if(($Plans|Measure-Object).count -gt 0){
                                    #Adiciona a licença ao usuário
                                    Add-MSOLUserLicense -Location BR -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $Plans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos habilitados: $($plansToEnable)"
                                }                         
                            }
                            catch{
                                #Grava o erro no log
                                Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                            }
                        }
                        #Se o usuário foi licenciado e já tem o SKU ID, atualiza os planos
                        else{
                            try{
                                if(($Plans|Measure-Object).count -gt 0){
                                    Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToEnable $Plans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                                    Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos habilitados: $($plansToEnable)"
                                }

                            }
                            catch{
                                #Grava o erro no log
                                Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                            }
                        }
                    }
                }
            }
            #Se nenhum usuário adicionado ao grupo for encontrado
            else{
                Write-Log -LogLevel Error -UserOrGroup $Group -Message "Nenhum usuário adicionado ao grupo para receber licença"
            }
            
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Iniciando tarefas de remoção de licenças para $(($groupMembersChange|Where-Object{$_.SideIndicator -eq "<="}|Measure-Object).Count) membros removidos"
            
            #Se algum usuário foi removido do grupo
            if(($groupMembersChange | Where-Object{$_.SideIndicator -eq "<="} | Measure-Object).Count -gt 0){

                #Para cada usuário que saiu do grupo, remove a licença correspondente
                foreach($user in $groupMembersChange | Where-Object{$_.SideIndicator -eq "<="}){

                    try{
                        #Remove a plano do usuário
                        Update-MSOLUserLicensePlan -Users $user.UserPrincipalName -SKU $SKU -PlansToDisable $Plans -LogFile .\$licenseModuleLogFileName -ErrorAction Stop
                        Write-Log -LogLevel Info -UserOrGroup $user.UserPrincipalName -Message "SKU $($SKU) adicionada ao usuário com os planos removidos: $($plansToEnable)"
                    }
                    catch{
                        #Grava o erro no log
                        Write-Log -LogLevel "Error" -UserOrGroup $user.UserPrincipalName -Message $_.Exception.Message
                    }
                }
            }
            #Se nenhum usuário removido do grupo for encontrado
            else{
                Write-Log -LogLevel Error -UserOrGroup $Group -Message "Nenhum usuário removido do grupo para remover licença"
            }
        }
        #Se a quantidade de mudanças no grupo for igual a 0
        else
        {
            Write-Log -LogLevel "Info" -UserOrGroup $Group -Message "Nenhum membro foi adicionado ou removido do grupo"
        }
    }
}

#Tenta efetuar a conexão no Microsoft Online Services
ConnectMsolService

#Para cada uma das configurações de licenciamento
foreach($license in $licenseConfigFile){
        #Executa a função de gerenciamento de licenças
        ManageLicense -Group $license.Group -SKU $license.SKU -Plans $license.Plans
}

#Finaliza o timer para calculo do tempo de execucão
$stopTimer = Get-Date

#Registra o termino do script com o tempo de execução
Write-Log -LogLevel Info -UserOrGroup "SCRIPT" -Message "############ Script Finalizado. Tempo de execução: $((New-TimeSpan -Start $startTimer -End $stopTimer).ToString("dd\.hh\:mm\:ss")) ############"
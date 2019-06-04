Import-Module "$PSScriptRoot\System\ImportExcel"
#Прикручиваем файл с функциями
. $PSScriptRoot\PSJiraRFD_Funcs.ps1

#Set-JiraConfigServer -server "http://servicedesk:8080" # Настройка на наш сервер Jira
$DPC = @{"Login"=$env:USERNAME; "Pass"=""; "Name"=(Get-ADUser -Identity $env:USERNAME).Name}
$DPC.Pass =  ConvertTo-SecureString -AsPlainText 'Password' -Force #Пароль для ленивых
$Global:DPCCred =  New-Object System.Management.Automation.PSCredential -ArgumentList $DPC.Login, $DPC.Pass # Переменая учётных данных
$netCred = New-Object System.Management.Automation.PSCredential -ArgumentList "DPC\$($DPC.Login)", $DPC.Pass # Переменая учётных данных

cls
get-date

$MyCred = @{'user'='Login';'pass'='Password'}
$Tech = @{'ID'='';'NAME'=''}

#Начало искомого временного периода
$FromDay = [datetime]::ParseExact('2019-04-01','yyyy-MM-dd',$null)
#Конец искомого временного периода
$ToDay = [datetime]::ParseExact('2019-04-09','yyyy-MM-dd',$null)
$OutFileName = "Отчёт по дежурству $($tech.NAME) за $($FromDay|get-date -f 'yyyy-MM-dd') - $($ToDay|get-date -f 'yyyy-MM-dd').xlsx"

Write-Host "Получение активности и комментариев по заявкам" -ForegroundColor Green
#Write-Host

$TestDay = $FromDay
while($TestDay -le $ToDay)
{
    Write-Host
    Write-Host "Получение списка заявок за $($TestDay|get-date -f 'dd.MM.yyyy')" -ForegroundColor Green

    $sm = 1
    while ($sm -le 3)
    {
        Write-host "Смена $sm" -f DarkCyan
        $Issue = APIGet-Issues -Ucredennials $MyCred  -techexp $tech -tFrom $TestDay -Smena $sm
        $res = ''
        $res = $Issue|foreach {
            $rez = @()
            $IssueOut = ''|select key, summary, status, timetotake, timetoresolve, whobreachtotake, whobreachetoresolve, author,jtext
            $IssueOut.key = $_.key
            $IssueOut.summary = $_.fields.summary
            $IssueOut.status = $_.fields.status.name
            $IssueOut.whobreachtotake = $_.fields.Customfield_16022
            $IssueOut.whobreachetoresolve = $_.fields.Customfield_16023
            $rez += $IssueOut
            $rez += GetHaC -Issue $_ -Ucredennials $MyCred -Duser $SLAVE -techexp $Tech -ComDate $TestDay -Smena $sm
            if ($rez.Count -gt 1){$rez}
        }
        if ($res -ne $null){
            $res|select  @{Name='Заявка'; Expression={$_.'key'}},@{Name='Описание'; Expression={$_.'summary'}},@{Name='Статус заявки/Изменение статуса'; Expression={$_.'status'}},`
            @{Name='Время изменения'; Expression={$_.'timetotake'}},@{Name='Просрочка взятия'; Expression={$_.'whobreachtotake'}},@{Name='Просрочка выполнения'; Expression={$_.'whobreachetoresolve'}},`
            @{Name='Сотрудник'; Expression={$_.'author'}},@{Name='Действие'; Expression={$_.'jtext'}} -ExcludeProperty timetoresolve|Export-Excel -Path "$PSScriptRoot\$OutFileName" `
            -WorkSheetname "$($TestDay.tostring("dd.MM.yy"))_см$sm" -AutoSize -KillExcel -BoldTopRow
        }
        $sm +=1
    }
    $TestDay=$TestDay.AddDays(1)
}

get-date
Write-Host "Окрашивание файла" -ForegroundColor Green
Colorite-Excell -Path "$PSScriptRoot\$OutFileName"
Write-Host

get-date
Write-Host "Horray!!!" -f Green
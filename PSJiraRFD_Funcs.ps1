####################################################################
#Функция преобразования изменения статусов на понятные для чтения
function HashToText($hashTable) 
{
    switch ($hashTable.field){
        'assignee'                          {$res = "Назначено "+$hashTable.fromString+" на "+$hashTable.toString}
        'Просрочка SLA "Принятие в работу"' {$res = $hashTable.field+$hashTable.toString.Replace(' ,','')}
        'status'                            {$res = "Изменен статус с "+$hashTable.fromString+" на "+$hashTable.toString}
        'resolution'                        {$res = "Решение по заявке:"+$hashTable.toString}
        'assignee'                          {$res = "Назначено с "+$hashTable.fromString+" на "+$hashTable.toString}
        'description'                       {$res = "Описание с "+$hashTable.fromString+" на "+$hashTable.toString}
        'Attachment'                        {$res = "Добавлено вложение "+$hashTable.fromString+" на "+$hashTable.toString}
        'Comment'                           {$res = "Изменен комментарий "+$hashTable.fromString+" на "+$hashTable.toString}
        Default {$res = $hashTable.field+" c "+$hashTable.fromString+' на '+$hashTable.toString}
    }
    return $res
}
####################################################################

####################################################################
#Функция получения списка заявок, которые были назначены на сотрудника за 2 смену указанного дня.
function APIGet-Issues($Ucredennials,$Duser,$techexp,$tFrom,$Smena)
{
    switch ($Smena){
        1 {$url = "http://servicedesk:8080/rest/api/2/search?jql=assignee changed FROM $($techexp.ID) DURING ('$($tFrom|get-date -format "yyyy-MM-dd 01:00")', '$($tFrom|get-date -format "yyyy-MM-dd 09:30")')&expand=changelog&maxResults=900"}
        2 {$url = "http://servicedesk:8080/rest/api/2/search?jql=assignee changed FROM $($techexp.ID) DURING ('$($tFrom|get-date -format "yyyy-MM-dd 09:00")', '$($tFrom|get-date -format "yyyy-MM-dd 17:30")')&expand=changelog&maxResults=900"}
        3 {$url = "http://servicedesk:8080/rest/api/2/search?jql=assignee changed FROM $($techexp.ID) DURING ('$($tFrom|get-date -format "yyyy-MM-dd 17:00")', '$(($tFrom.AddDays(1))|get-date -format "yyyy-MM-dd 01:30")')&expand=changelog&maxResults=900"}
    }
    $Credent = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($($Ucredennials.user+":"+$Ucredennials.pass)))
    $headers = @{"Authorization" = $("Basic " + $Credent)}
    try {
        $json = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -ContentType "application/json; charset=utf-8"
        $isRes = (ConvertFrom-Json $json.Content).issues
    } catch {
        write-host "error on Get-Issue" -f DarkYellow
        Write-Host $isRes.key -f Red
        write-host $_
    }
    return $isRes
}
####################################################################


####################################################################
#Функция получения комментариев из заявки
function APIGet-Comment($Ucredennials,$Issue)
{
    $url = "$($Issue.self)/comment"
    $Credent = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($($Ucredennials.user+":"+$Ucredennials.pass)))
    $headers = @{"Authorization" = $("Basic " + $Credent)}
    try {
        $json = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -ContentType "application/json; charset=utf-8"
        $response = ConvertFrom-Json $json.Content
        $conv = $response.comments
    } catch {
        write-host "error on Get-Comment" -f Cyan
        Write-Host $Issue.key
        write-host $_
    }
    return $conv
}
####################################################################


####################################################################
#Функция получения активности сотрудника в указанной заявке
function GetHaC($Issue,$Ucredennials,$techexp,$ComDate,$Smena) #,$Duser
{
    If ($Issue -ne $null)
    {
        ##Комментарии к заявке
        $conv = @(
            APIGet-Comment -Ucredennials $Ucredennials -Issue $Issue|foreach{
                $result = ''|select id,status,author,date,timetotake,jtext,whobreachtotake, whobreachetoresolve
                $result.author = $_.author.displayname
                $result.timetotake = [datetime]::ParseExact($_.created.Substring(0,19),'yyyy-MM-ddTHH:mm:ss',$null)
                $result.date = $result.timetotake
                $result.jtext = "Комментарий:"+$_.body
                $result 
            }
        )
        ########
        ##История изменения заявки
        $conv += @(
            $Issue.changelog.histories| foreach {
                foreach ($item in $_.items) {
                    $result = ''|select id,status,author,date,timetotake,jtext,whobreachtotake, whobreachetoresolve
                    $result.id = $_.id
                    $result.author = $_.author.displayname
                    $result.timetotake = [datetime]::ParseExact($_.created.Substring(0,19),'yyyy-MM-ddTHH:mm:ss',$null)
                    $result.date = $result.timetotake
                    $result.jtext = HashToText -hashtable $item
                    $A = ($Issue.changelog.histories|where {($_.id -le $result.id) -and ($_.items.field -imatch "status")}).items
                    $B = if ($A -ne $null){(($A|where {$_.field -eq "status"})[-1]).toString} else {""}
                    $result.status = $B
                    if ($result.jtext -imatch 'Просрочка SLA "Принятие в работу"')
                    {
                        $result.whobreachtotake = (($result.jtext).replace('Просрочка SLA "Принятие в работу" c  на  ','')).replace(' ,','')
                    }
                    if ($result.jtext -imatch 'Просрочка SLA "Решение заявки"')
                    {
                        $result.whobreachetoresolve = (($result.jtext).replace('Просрочка SLA "Решение заявки" c  на  ','')).replace(' ,','')
                    }
                    $result
                }
            }
        )
        switch ($Smena){
            1 {return $($conv|where {($_.timetotake -ge $ComDate.AddHours(1.5)) -and ($_.timetotake -lt $ComDate.AddHours(9.5))}|sort -Property timetotake ) }
            2 {return $($conv|where {($_.timetotake -ge $ComDate.AddHours(9.5)) -and ($_.timetotake -lt $ComDate.AddHours(17.5))}|sort -Property timetotake ) }
            3 {return $($conv|where {($_.timetotake -ge $ComDate.AddHours(17.5)) -and ($_.timetotake -lt $ComDate.AddDays(1).AddHours(1.5))}|sort -Property timetotake ) }
        }
    }
}
####################################################################


####################################################################
#Функция разукрашивания файла Excell
Function Colorite-Excell ($Path)
{
    try
        {
        $excell = New-Object -comobject Excel.Application
        $excell.visible = $False
        $workbook = $excell.workbooks.open($Path)
        foreach ($sheet in $workbook.Sheets)
        {
            Foreach ($row in $sheet.UsedRange.Rows)
            {
                #Окрашивание заголовка
                if ($row.row -eq 1)
                {
                    $row.font.bold=$true
                    $row.interior.colorindex=36
                }
                #Окрашивание строк с номером заявки
                elseif (($sheet.Cells.Item($row.row,1).text -ne '') -and (($sheet.Cells.Item($row.row,5).text -eq '') -and ($sheet.Cells.Item($row.row,6).text -eq '')))
                {
                    $row.interior.colorindex=42
                }
                #Окрашивание строк с комментариями и изменениями
                elseif (($sheet.Cells.Item($row.row,1).text -ne '') -and (($sheet.Cells.Item($row.row,5).text -ne '') -or ($sheet.Cells.Item($row.row,6).text -ne '')))
                {
                    $row.interior.colorindex=41
                }
                #Окрашивание строк с номером заявки при наличии превышения времени взятия и/или времени решения
                elseif (($sheet.Cells.Item($row.row,1).text -eq '') -and (($sheet.Cells.Item($row.row,5).text -ne '') -or ($sheet.Cells.Item($row.row,6).text -ne ''))) 
                {
                    $row.interior.colorindex=3
                }
                #Окрашивание строк изменений в заявке при наличии превышения времени взятия и/или времени решения в текущую смену
                else
                {
                    $row.interior.colorindex=35
                }
                $workbook.Save()
            }
        }
        $workbook.Save()
        $workbook.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)|Out-Null
        Remove-Variable workbook -ErrorAction SilentlyContinue|Out-Null
        $excell.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excell)|Out-Null
        Remove-Variable excell -ErrorAction SilentlyContinue|Out-Null
    }
    catch
    {
        Write-Host $_.Error
    }
}
####################################################################

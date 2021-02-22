cls

$Servers = Get-TransportService

$Start = "01/25/2021 00:00"

$Start = [datetime]$Start

$Days = 14

$End = $Start.AddDays(14)

#$End = "02/08/2021 23:59"

#$End = [datetime]$End

#$Days = ($End - $Start).Days

Write-Output "Объём трафика с $Start по $End ($Days дней)`n"

foreach ($Server in $Servers) {

    $ReceiveMessages = Get-Messagetrackinglog -Server $Server.Name -EventID "RECEIVE" -Source "STOREDRIVER" -Start $Start -End $End -ResultSize Unlimited

    foreach ($ReceiveMessage in $ReceiveMessages) {

        $ReceiveMessagesSize = $ReceiveMessage.TotalBytes / 1mb

        $ReceiveMessagesTotalSize += $ReceiveMessagesSize

    }

    $ReceiveMessagesTotalSize = $ReceiveMessagesTotalSize / $Days

    $ReceiveMessagesTotalSizeAll += $ReceiveMessagesTotalSize


    $SendMessages = Get-Messagetrackinglog -Server $Server -EventID "DELIVER" -Start $Start -End $End -ResultSize Unlimited

    foreach ($SendMessage in $SendMessages) {

        $SendMessagesSize = $SendMessage.TotalBytes / 1mb

        $SendMessagesTotalSize += $SendMessagesSize

    }

    $SendMessagesTotalSize = $SendMessagesTotalSize / $Days

    $SendMessagesTotalSizeAll += $SendMessagesTotalSize

    $InsideReceiveMessages = Get-Messagetrackinglog -Server $Server -EventID "RECEIVE" -Source "STOREDRIVER" -Start $Start -End $End -ResultSize Unlimited | where {$_.Recipients -match "ueip.ru" -and $_.Sender -match "ueip.ru"}

    foreach ($InsideReceiveMessage in $InsideReceiveMessages) {

        $InsideReceiveMessagesSize = $InsideReceiveMessage.TotalBytes / 1mb

        $InsideReceiveMessagesTotalSize += $InsideReceiveMessagesSize

    }

    $InsideReceiveMessagesTotalSize = $InsideReceiveMessagesTotalSize / $Days

    $InsideReceiveMessagesTotalSizeAll += $InsideReceiveMessagesTotalSize

    $InsideSendMessages = Get-Messagetrackinglog -Server $Server -EventID "DELIVER" -Start $Start -End $End -ResultSize Unlimited | where {$_.Recipients -match "ueip.ru" -and $_.Sender -match "ueip.ru"}

    foreach ($InsideSendMessage in $InsideSendMessages) {

        $InsideSendMessagesSize = $InsideSendMessage.TotalBytes / 1mb

        $InsideSendMessagesTotalSize += $InsideSendMessagesSize

    }

    $InsideSendMessagesTotalSize = $InsideSendMessagesTotalSize / $Days

    $InsideSendMessagesTotalSizeAll += $InsideSendMessagesTotalSize

    Write-Output "Входящий трафик сервера $($Server.Name): $([Math]::Round($ReceiveMessagesTotalSize, 2)) МБ/день"
    
    Write-Output "Исходящий трафик сервера $($Server.Name): $([Math]::Round($SendMessagesTotalSize, 2)) МБ/день"

    Write-Output "Общий трафик сервера $($Server.Name): $([Math]::Round($ReceiveMessagesTotalSize + $SendMessagesTotalSize, 2)) МБ/день"

    Write-Output "`n"

    Write-Output "Входящий внутренний трафик сервера $($Server.Name): $([Math]::Round($InsideReceiveMessagesTotalSize, 2)) МБ/день"
    
    Write-Output "Исходящий внутренний трафик сервера $($Server.Name): $([Math]::Round($InsideSendMessagesTotalSize, 2)) МБ/день"

    Write-Output "Общий внутренний трафик сервера $($Server.Name): $([Math]::Round($InsideReceiveMessagesTotalSize + $InsideSendMessagesTotalSize, 2)) МБ/день"

    Write-Output "`n"

}

Write-Output "Входящий трафик всех серверов: $([Math]::Round($ReceiveMessagesTotalSizeAll, 2)) МБ/день"

Write-Output "Исходящий трафик всех серверов: $([Math]::Round($SendMessagesTotalSizeAll, 2)) МБ/день"

Write-Output "Общий трафик всех серверов: $([Math]::Round($ReceiveMessagesTotalSizeAll + $SendMessagesTotalSizeAll, 2)) МБ/день"

Write-Output "`n"

Write-Output "Входящий внутренний трафик всех серверов: $([Math]::Round($InsideReceiveMessagesTotalSizeAll, 2)) МБ/день"

Write-Output "Исходящий внутренний трафик всех серверов: $([Math]::Round($InsideSendMessagesTotalSizeAll, 2)) МБ/день"

Write-Output "Общий внутренний трафик всех серверов: $([Math]::Round($InsideReceiveMessagesTotalSizeAll + $InsideSendMessagesTotalSizeAll, 2)) МБ/день"
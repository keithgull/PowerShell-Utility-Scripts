# Windowsのlog-on(7001), log-off(7002) イベントの時刻を日ごとに出力するスクリプト
# 勤務報告用に使えます。
# parameter yyyymm 対象の年月
# parameter path   出力ファイル場所＋ファイル名
#                  デフォルトはスクリプトと同じ場所、デフォルトファイル名はextract-logonlogoff_${yyyymm}.csv
param (
    [string]${yyyymm},
    [string]${path} = "${PSScriptRoot}\extract-logonlogoff_${yyyymm}.csv"
)

if (-not ${yyyymm}) {
    Write-Host "パラメータに対象年月が必要です。yyyymm形式で年月を指定して実行してください"
    exit
}

${year} = [int]${yyyymm}.Substring(0, 4)
${month} = [int]${yyyymm}.Substring(4, 2)
${startDate} = Get-Date -Year ${year} -Month ${month} -Day 1 -Hour 0 -Minute 0 -Second 0
${endDate}   = ${startDate}.AddMonths(1).AddSeconds(-1)

${events} = Get-WinEvent -FilterHashtable @{
    LogName   = 'System'
    Id        = 7001, 7002
    StartTime = ${startDate}
    EndTime   = ${endDate}
}

${logonEvents}  = @{}
${logoffEvents} = @{}

foreach (${event} in ${events}) {
    ${date} = ${event}.TimeCreated.ToString("MM/dd (ddd)")
    if (${event}.Id -eq 7001) {
        if (-not ${logonEvents}[${date}] -or ${event}.TimeCreated -lt ${logonEvents}[${date}]) {
            ${logonEvents}[${date}] = ${event}.TimeCreated
        }
    } elseif (${event}.Id -eq 7002) {
        if (-not ${logoffEvents}[${date}] -or ${event}.TimeCreated -gt ${logoffEvents}[${date}]) {
            ${logoffEvents}[${date}] = ${event}.TimeCreated
        }
    }
}

${output} = "date`tlogon`tlogoff`n"
foreach (${day} in 1..${endDate}.Day) {
    ${currentDate} = Get-Date -Year ${year} -Month ${month} -Day ${day}
    ${dateString}  = ${currentDate}.ToString("MM/dd (ddd)")
    ${logon}       = ${logonEvents}[${dateString}]  ? ${logonEvents}[${dateString}].ToString("HH:mm") : "`t"
    ${logoff}      = ${logoffEvents}[${dateString}] ? ${logoffEvents}[${dateString}].ToString("HH:mm") : "`t"
    ${output} += "${dateString}`t${logon}`t${logoff}`n"
}

${output} | Out-File -FilePath ${path} -Encoding UTF8 -Force
Write-Host "ログオン・ログオフイベントを抽出しました。結果は ${path} に保存されました。"

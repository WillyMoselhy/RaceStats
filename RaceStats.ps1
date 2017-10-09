####################   RaceStats   ####################
#                                                     #
#A weekend project to help produce statistics for     #
#running races.                                       #
#                                                     #
#Usage: Create a tab separated CSV file as an input.  #
#And run the script. Can also resume from a previous  #
#session                                              #
#######################################################


$ErrorActionPreference = "stop"

#Change this value to meet the number of laps required.
$NumberOfLaps = 15


#Path to CSV file containing list of players
$RaceStats = Import-Csv -Path "D:\temp\RaceStats.csv" -Delimiter "	"

#region: Functions
function Out-ExcelCsv
{

  param($Path = "$env:temp\$(Get-Date -Format yyyyMMddHHmmss).csv")

  $input | Export-CSV -Path $Path -UseCulture -Encoding UTF8 -NoTypeInformation

  Invoke-Item -Path $Path

}
#endregion


$NumberOfPlayers = $RaceStats.Count

#Add properties to each player
foreach ($Player in $RaceStats){
    $Player | Add-Member -MemberType NoteProperty -Name "LapCount" -Value 0
    $Player | Add-Member -MemberType NoteProperty -Name "TotalTime" -Value $null
    $Player | Add-Member -MemberType NoteProperty -Name "Finished" -Value $null
    $Player | Add-Member -MemberType NoteProperty -Name "ETA" -Value $null
    for ($i=1;$i -le $NumberOfLaps;$i++){
        $Player | Add-Member -MemberType NoteProperty -Name "Lap$($i)Date" -Value $null
        $Player | Add-Member -MemberType NoteProperty -Name "Lap$($i)Total" -Value $null
        $Player | Add-Member -MemberType NoteProperty -Name "Lap$($i)Time" -Value $null
        $Player | Add-Member -MemberType NoteProperty -Name "Lap$($i)s" -Value $null

    }
}

#Save race starting time for lap time calculations
$RaceStartTime = Get-Date

$PlayersFinished = 0 #Used to exit the loop when all players finish

While ($PlayersFinished -ne $NumberOfPlayers)
{
    try{
        [int]$PlayerNumber = Read-Host -Prompt "Type Player Number to register lap (use '-#' to undo)"
    }
    catch{
        Write-Host "Invalid input!." -ForegroundColor Red
        continue
    }

    
    #region: Decide to add lap/ remove lap/ throw error
    if ($PlayerNumber -gt 0 -and $PlayerNumber -le $NumberOfPlayers) {
        $Action = "AddLap"
        $PlayerID = $PlayerNumber -1
    }
    elseif ($PlayerNumber -lt 0 -and ($PlayerNumber * -1) -le $NumberOfPlayers) {
        $Action = "RemoveLap"
        $PlayerID = ($PlayerNumber * -1) -1
    }
    else {
        Write-Host "Invalid player number." -ForegroundColor Red
        continue
    }
    #endregion

    #region: Add lap if player is not finished.
    if ($Action -eq "AddLap" -and $RaceStats[$PlayerID].Finished -ne $true){
        #Increment player lap count
        $RaceStats[$PlayerID].LapCount++
        $LapCount = $RaceStats[$PlayerID].LapCount

        #Update finish boolean for player
        if ($LapCount -eq $NumberOfLaps) {
            $RaceStats[$PlayerID].Finished = $true
        }

        #Update total racetime for player
        $TotalTime = New-TimeSpan -Start $RaceStartTime -End (Get-Date)
        $RaceStats[$PlayerID].TotalTime = $TotalTime
        $RaceStats[$PlayerID]."Lap$($LapCount)Total"= $TotalTime
        $RaceStats[$PlayerID]."Lap$($LapCount)Date" = Get-Date

        #Calculate lap time for player
        if ($RaceStats[$PlayerID].LapCount -eq 1){ #For first lap
            
            $LapTime = $TotalTime
            $RaceStats[$PlayerID]."Lap$($LapCount)Time"= $TotalTime
        }
        else { #For all other laps

            $LapTime = New-TimeSpan -Start $RaceStats[$PlayerID]."Lap$($LapCount-1)Date" `
                                    -End (Get-Date)
            $RaceStats[$PlayerID]."Lap$($LapCount)Time" = $LapTime
        }

        #Update Lap Print
        $LapPrint = "{0:mm}:{0:ss}" -f $LapTime
        $RaceStats[$PlayerID]."Lap$($LapCount)s" = $LapPrint



        
    }
    #endregion

    #region:Remove lap and reset finished state
    if ($Action -eq "RemoveLap"){

        $LapToRemove = $RaceStats[$PlayerID].LapCount

        if ($RaceStats[$PlayerID].LapCount -eq 1) {#Remove first lap
            $RaceStats[$PlayerID].TotalTime = $null
            $RaceStats[$PlayerID].Lap1Date= $null
            $RaceStats[$PlayerID].Lap1Time= $null
            $RaceStats[$PlayerID].Lap1Total= $null
            $RaceStats[$PlayerID].Lap1s= $null

            #Decrement player lap count
            $RaceStats[$PlayerID].LapCount--
            $RaceStats[$PlayerID].Finished = $null

        }
        elseif ($RaceStats[$PlayerID].LapCount -gt 1) {#all other laps
            #Total time of previous lap
            $RaceStats[$PlayerID].TotalTime = $RaceStats[$PlayerID]."Lap$($LapToRemove-1)".LapTotal

            #reset removed lap
            $RaceStats[$PlayerID].TotalTime = $null
            $RaceStats[$PlayerID]."Lap$($LapToRemove)Date" = $null
            $RaceStats[$PlayerID]."Lap$($LapToRemove)Time"= $null
            $RaceStats[$PlayerID]."Lap$($LapToRemove)Total"= $null
            $RaceStats[$PlayerID]."Lap$($LapToRemove)s"= $null

            #Decrement player lap count
            $RaceStats[$PlayerID].LapCount--
            $RaceStats[$PlayerID].Finished = $null
        }

    }
    #endregion

    #Update ETA
    $SumLaps = 0
    for ($i=1;$i -le $LapCount; $i++){
        $SumLaps += $RaceStats[$PlayerID]."Lap$($i)Time".TotalMinutes
    }
        
    if ($LapCount -gt 0) {
        $AveragePace = $SumLaps/$LapCount
        $ETA = New-TimeSpan -Seconds ($AveragePace*60*$NumberOfLaps)
        $RaceStats[$PlayerID].ETA = $ETA
    }
    else {
        $RaceStats[$PlayerID].ETA = $null
    }


    $PlayersFinished = ($RaceStats | Measure-Object -Property Finished -Sum).Sum
    #Print race

    cls
    $RaceStats | select Player*,@{l="Total";e={"{0:hh}:{0:mm}:{0:ss}" -f $_.TotalTime}},ETA,LapCount,Lap*s,Finished |ft Player*,Total,ETA,LapCount,Lap*s,Finished
}

#Export results to excel for analysis
$RaceStats | select PlayerNumber,PlayerName,TotalTime,LapCount,Lap*time,Finished | Out-ExcelCsv -Path "D:\Temp\RaceStats$(Get-Date -Date $RaceStartTime -Format "yyyyMMdd-HHmm").csv"

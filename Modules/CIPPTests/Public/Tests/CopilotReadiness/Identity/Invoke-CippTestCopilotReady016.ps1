function Invoke-CippTestCopilotReady016 {
    <#
    .SYNOPSIS
    Microsoft 365 Copilot active user count summary by app (informational, 30-day period)
    #>
    param($Tenant)

    # Reports aggregate active user counts per Copilot app (Teams, Outlook, Word, Excel, etc.)
    # for the past 30 days. Informational — shows which apps are driving Copilot adoption
    # and where engagement is low.

    try {
        $SummaryData = Get-CIPPTestData -TenantFilter $Tenant -Type 'CopilotUserCountSummary'

        if (-not $SummaryData) {
            Add-CippTestResult -TenantFilter $Tenant -TestId 'CopilotReady016' -TestType 'Identity' -Status 'Skipped' -ResultMarkdown 'No Copilot user count summary data found in database. Data collection may not yet have run for this tenant.' -Risk 'Informational' -Name 'Copilot active user count by app' -UserImpact 'Low' -ImplementationEffort 'Low' -Category 'Copilot Readiness'
            return
        }

        $Summary = if ($SummaryData.adoptionByProduct) {
            $SummaryData.adoptionByProduct | Select-Object -First 1
        } elseif ($SummaryData -is [array]) {
            $SummaryData[0]
        } else {
            $SummaryData
        }

        $MetaFields = @('reportRefreshDate', 'reportPeriod', 'reportDate', 'id')

        $AppProperties = $Summary.PSObject.Properties.Name |
            Where-Object { $_ -notin $MetaFields -and $_ -match '(EnabledUsers|ActiveUsers)$' }

        $AppNames = $AppProperties |
            ForEach-Object { $_ -replace '(EnabledUsers|ActiveUsers)$', '' } |
            Select-Object -Unique

        $TotalAppCount = @($AppNames).Count

        if (-not $AppNames -or $TotalAppCount -eq 0) {
            $Result = "No Microsoft 365 Copilot usage was detected in the past 30 days.`n`n"
            $Result += 'This tenant either has no Copilot licenses assigned or users have not yet started using Copilot features.'
            # Add-CippTestResult -TenantFilter $Tenant -TestId 'CopilotReady016' -TestType 'Identity' -Status 'Informational' -ResultMarkdown $Result -Risk 'Informational' -Name 'Copilot active user count by app' -UserImpact 'Low' -ImplementationEffort 'Low' -Category 'Copilot Readiness'
            return
        }

        $Result = "## Copilot Active Users by App (Last 30 Days)`n`n"
        $Result += "| App | Active Users |`n"
        $Result += "|-----|-------------|`n"

        # Friendly display names for known apps
        $FriendlyNames = @{
            'microsoftTeams' = 'Microsoft Teams'
            'word'           = 'Word'
            'powerPoint'     = 'PowerPoint'
            'outlook'        = 'Outlook'
            'excel'          = 'Excel'
            'oneNote'        = 'OneNote'
            'loop'           = 'Loop'
            'anyApp'         = 'Any App'
            'copilotChat'    = 'Copilot Chat'
        }

        foreach ($App in $AppNames) {
            # Use the lookup if we have one, otherwise split camelCase as a fallback
            if ($FriendlyNames.ContainsKey($App)) {
                $AppName = $FriendlyNames[$App]
            } else {
                $AppName = (($App -creplace '([a-z])([A-Z])', '$1 $2') -replace '^.', { $_.Value.ToUpper() })
            }

            $ActiveUsers = $Summary."$($App)ActiveUsers"
            $Result += "| $AppName | $ActiveUsers |`n"
        }

        if ($SummaryData.reportRefreshDate) {
            $Result += "`n*Data as of $($SummaryData.reportRefreshDate).*"
        }

        Add-CippTestResult -TenantFilter $Tenant -TestId 'CopilotReady016' -TestType 'Identity' -Status 'Informational' -ResultMarkdown $Result -Risk 'Informational' -Name 'Copilot active user count by app' -UserImpact 'Low' -ImplementationEffort 'Low' -Category 'Copilot Readiness'

    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -API 'Tests' -tenant $Tenant -message "Failed to run test CopilotReady016: $($ErrorMessage.NormalizedError)" -sev Error -LogData $ErrorMessage
        Add-CippTestResult -TenantFilter $Tenant -TestId 'CopilotReady016' -TestType 'Identity' -Status 'Failed' -ResultMarkdown "Test failed: $($ErrorMessage.NormalizedError)" -Risk 'Informational' -Name 'Copilot active user count by app' -UserImpact 'Low' -ImplementationEffort 'Low' -Category 'Copilot Readiness'
    }
}

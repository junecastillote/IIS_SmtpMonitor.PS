Function IsFeatureInstalled {
    # This function returns a boolean value whether the Web-WMI and SMTP-Server features are installed on the local host.

    $ProgressPreference = 'SilentlyContinue'
    $feature_names = @('WEB-WMI', 'SMTP-SERVER')
    $installed_features = @(Get-WindowsFeature -Name $feature_names)

    # If the result is less than the feature names to check, it means one of the feature names is invalid or doesn't exist.
    if ($installed_features.Count -lt $feature_names.Count) {
        SayError "The following features are not installed on this host: $(($feature_names | Where-Object {$installed_features.Name -notcontains $_} ) -join ';')"
        return $false
    }

    # If one or more features are not installed.
    if ($installed_features.Installed -contains $false) {
        SayError "The following features are not installed on this host: $(($installed_features | Where-Object { !$_.Installed }).Name -join ';')"
        return $false
    }

    return $true
}
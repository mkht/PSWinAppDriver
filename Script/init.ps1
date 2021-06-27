#Require -Version 5.0

$PSModuleRoot = Split-Path $PSScriptRoot -Parent
$LibPath = Join-Path $PSModuleRoot '\Lib'
Write-Debug ('$PSModuleRoot:{0}' -f $PSModuleRoot)

$Cases = @(
    @{
        Type    = 'OpenQA.Selenium.By'
        DllName = 'WebDriver.dll'
    }
    @{
        Type    = 'OpenQA.Selenium.Support.UI.SelectElement'
        DllName = 'WebDriver.Support.dll'
    }
    @{
        Type    = 'Newtonsoft.Json.JsonConverter'
        DllName = 'Newtonsoft.Json.dll'
    }
    @{
        Type    = 'Castle.Core.Configuration.IConfiguration'
        DllName = 'Castle.Core.dll'
    }
)

$Cases.ForEach( {
        if (!($_.Type -as [type])) {
            try {
                Add-Type -Path (Join-Path $LibPath $($_.DllName)) -ErrorAction Stop
            }
            catch {
                Write-Error "Couldn't load $($_.DllName)"
            }
        }
    })


# Load Appium
if (!('OpenQA.Selenium.Appium.AppiumCapabilities' -as [type])) {
    try {
        Add-Type -Path (Join-Path $LibPath 'SeleniumExtras.PageObjects.dll') -ErrorAction Stop
        Add-Type -Path (Join-Path $LibPath 'Appium.Net.dll') -ErrorAction Stop
    }
    catch {
        Write-Error "Couldn't load Appium"
    }
}

# Import Assert function
if ($AssertPath = Resolve-Path "$PSModuleRoot\Function\Assert.psm1" -ea SilentlyContinue) {
    Import-Module $AssertPath -Force
}
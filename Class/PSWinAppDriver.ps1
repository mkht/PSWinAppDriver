#Requires -Version 5
using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Appium

#region Enum:ImageFormat
Enum ImageFormat{
    Png = 0
    Jpeg = 1
    Gif = 2
    Tiff = 3
    Bmp = 4
}
#endregion

#region Enum:SelectorType
Enum SelectorType{
    None
    Id
    Name
    Tag
    ClassName
    Link
    XPath
    Css
}
#endregion

#region Class:Selector
class Selector {
    [string]$Expression
    [SelectorType]$Type = [SelectorType]::None
    [Object]$By

    Selector() {
    }

    Selector([string]$Expression) {
        $this.Expression = $Expression
    }

    Selector([string]$Expression, [SelectorType]$Type) {
        $this.Expression = $Expression
        $this.Type = $Type
        $this.By = [Selector]::GetSeleniumBy($Expression, $Type)
    }

    [string]ToString() {
        return $this.Expression
    }

    static [Selector]Parse([string]$Expression) {
        $local:ret = switch -Regex ($Expression) {
            '^id=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Id); break }
            '^name=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Name); break }
            '^tag=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Tag); break }
            '^className=(.+)' { [Selector]::new($Matches[1], [SelectorType]::ClassName); break }
            '^link=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Link); break }
            '^xpath=(.+)' { [Selector]::new($Matches[1], [SelectorType]::XPath); break }
            '^/.+' { [Selector]::new($Matches[0], [SelectorType]::XPath); break }
            '^css=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Css); break }
            Default { [Selector]::new($Expression) }
        }
        return $ret
    }

    static Hidden [Object]GetSeleniumBy([string]$Expression, [SelectorType]$Type) {
        $local:SelectorObj =
        switch ($Type) {
            'Id' { Invoke-Expression '[ByAccessibilityId]::AccessibilityId($Expression)'; break }
            'Name' { Invoke-Expression '[By]::Name($Expression)'; break }
            'Tag' { Invoke-Expression '[By]::TagName($Expression)'; break }
            'ClassName' { Invoke-Expression '[By]::ClassName($Expression)'; break }
            'Link' { Invoke-Expression '[By]::LinkText($Expression)'; break }
            'XPath' { Invoke-Expression '[By]::XPath($Expression)'; break }
            'Css' { Invoke-Expression '[By]::CssSelector($Expression)'; break }
            Default {
                throw 'Undefined selector type'
            }
        }
        return $SelectorObj
    }
}
#endregion

#region Class:SpecialKeys
class SpecialKeys {
    [hashtable]$KeyMap

    SpecialKeys() {
        $PSModuleRoot = Split-Path $PSScriptRoot -Parent
        $this.KeyMap = (ConvertFrom-StringData (Get-Content (Join-Path $PSModuleRoot 'Static\KEYMAP.txt') -raw))
    }

    [string]ConvertSeleniumKeys([string]$key) {
        if (!$this.KeyMap) { return '' }
        if ($this.KeyMap.ContainsKey($key)) {
            $tmp = $this.KeyMap.$key
            return [string](Invoke-Expression '[keys]::($tmp)')
        }
        else {
            return ('${{{0}}}' -f $key)
        }
    }
}
#endregion

#region Class:PSWinAppDriver
class PSWinAppDriver {
    #region Public Properties
    $Driver
    $Actions
    [AppiumOptions] $ApplicationOptions
    [uri]$WinAppDriverUrl = 'http://127.0.0.1:4723/'
    #endregion

    #region Hidden properties
    Hidden [string] $InstanceId
    Hidden [SpecialKeys] $SpecialKeys
    Hidden [string] $AppDriverExe
    Hidden [System.Diagnostics.Process] $AppDriverProcess
    Hidden [string] $PSModuleRoot
    Hidden [int] $ImplicitWait = 2
    Hidden [int] $PageLoadTimeout = 30
    # Hidden [System.Timers.Timer]$Timer
    # Hidden [int]$RecordInterval = 5000
    #endregion

    #region Constructor:PSWinAppDriver
    PSWinAppDriver() {
        $this.PSModuleRoot = Split-Path $PSScriptRoot -Parent
        $this.InstanceId = [string]( -join ((1..4) | ForEach-Object { Get-Random -input ([char[]]((48..57) + (65..90) + (97..122))) })) #4-digits random id
        $this.SpecialKeys = [SpecialKeys]::New()
        $this._LoadSelenium()
        $this.AppDriverExe = $this._FindWebDriver()
        $this.ApplicationOptions = $this._NewDriverOptions()
    }
    #endregion

    #region Method:SetImplicitWait()
    [void]SetImplicitWait([int]$TimeoutInSeconds) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            [int]$local:tmp = $this.ImplicitWait
            try {
                if ($TimeoutInSeconds -lt 0) {
                    $TimeSpan = [System.Threading.Timeout]::InfiniteTimeSpan
                }
                else {
                    $TimeSpan = New-TimeSpan -Seconds $TimeoutInSeconds -ea Stop
                }
                $this.Driver.Manage().Timeouts().ImplicitWait = $TimeSpan
                $this.ImplicitWait = $TimeoutInSeconds
            }
            catch {
                $this.ImplicitWait = $tmp
            }
        }
    }
    #endregion

    #region Method:Get/SetWindowSize()
    [System.Drawing.Size]GetWindowSize() {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            return $this.Driver.Manage().Window.Size
        }
    }

    [void]SetWindowSize([System.Drawing.Size]$Size) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            $this.Driver.Manage().Window.Size = $Size
        }
    }

    [void]SetWindowSize([int]$Width, [int]$Height) {
        $this.SetWindowSize([System.Drawing.Size]::New($Width, $Height))
    }
    #endregion

    #region Method:Start()
    [void]Start([string]$ApplicationPath, [string]$Arguments, [string]$WorkingDirectory) {
        if ($this.Driver) {
            $this.Quit()
        }

        $this._StartWinAppDriver()

        $this.ApplicationOptions.AddAdditionalCapability('app', $ApplicationPath)

        if (-not [string]::IsNullOrWhiteSpace($Arguments)) {
            $this.ApplicationOptions.AddAdditionalCapability('appArguments', $Arguments)
        }
        if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
            $this.ApplicationOptions.AddAdditionalCapability('appWorkingDir', $WorkingDirectory)
        }

        $this.Driver = [Windows.WindowsDriver[Windows.WindowsElement]]::new($this.WinAppDriverUrl, $this.ApplicationOptions)

        #Set default implicit wait
        if ($this.Driver) { $this.SetImplicitWait($this.ImplicitWait) }

        #Create Action instance
        if ($this.Driver) { $this.Actions = Invoke-Expression '[Interactions.Actions]::New($this.Driver)' }
    }

    [void]Start([string]$ApplicationPath) {
        $this.Start($ApplicationPath, '', '')
    }

    [void]Start([string]$ApplicationPath, [string]$Arguments) {
        $this.Start($ApplicationPath, $Arguments, '')
    }
    #endregion

    #region Method:Quit()
    [void]Quit() {
        # Stop animation recorder if running
        # if ($this.Timer) {
        #     try {
        #         $this._DisposeRecorder()
        #     }
        #     catch { }
        # }

        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            try {
                $this.Driver.Quit()
                Write-Verbose 'Application terminated successfully.'
            }
            catch {
                throw 'Failed to terminate Application.'
            }
            finally {
                $this.Driver = $null
            }
        }

        if ($this.AppDriverProcess) {
            $this.AppDriverProcess | Stop-Process
            $this.AppDriverProcess = $null
        }
    }
    #endregion

    #region Method:Close()
    [void]Close() {
        if ($null -ne $this.Driver) {
            $this.Driver.Close()
        }
    }
    #endregion

    #region Method:Open()
    [void]Open([string]$ApplicationPath, [string]$Arguments, [string]$WorkingDirectory) {
        $this.Start($ApplicationPath, $Arguments, $WorkingDirectory)
    }

    [void]Open([string]$ApplicationPath, [string]$Arguments) {
        $this.Start($ApplicationPath, $Arguments)
    }

    [void]Open([string]$ApplicationPath) {
        $this.Start($ApplicationPath)
    }
    #endregion

    #region Method:GetApplicationInfo()
    [HashTable]GetApplicationInfo() {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            return $this.Driver.Capabilities.ToDictionary()
        }
    }
    #endregion

    #region Method:GetTitle()
    [string]GetTitle() {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            return [string]$this.Driver.Title
        }
    }
    #endregion

    [string]GetAttribute([string]$Target, [string]$Attribute) {
        return [string]($this.FindElement($Target).GetAttribute($Attribute))
    }

    [string]GetText([string]$Target) {
        return [string]($this.FindElement($Target).Text)
    }

    [string]GetSelectedLabel([string]$Target) {
        return [string]($this._GetSelectElement($Target).SelectedOption.Text)
    }

    [bool]IsVisible([string]$Target) {
        return [bool]($this.FindElement($Target).Displayed)
    }

    #region Method:FindElement()
    [Object]FindElement([Selector]$Selector) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            if ($Selector.By) {
                # $this.WaitForPageToLoad($this.PageLoadTimeout)
                return $this.Driver.FindElement($Selector.By)
            }
            return $null
        }
    }

    [Object]FindElement([string]$SelectorExpression) {
        return $this.FindElement([Selector]::Parse($SelectorExpression))
    }

    [Object]FindElement([string]$SelectorExpression, [SelectorType]$Type) {
        return $this.FindElement([Selector]::New($SelectorExpression, $Type))
    }
    #endregion

    #region Method:FindElements()
    [Object[]]FindElements([Selector]$Selector) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            if ($Selector.By) {
                # $this.WaitForPageToLoad($this.PageLoadTimeout)
                return $this.Driver.FindElements($Selector.By)
            }
            return $null
        }
    }

    [Object[]]FindElements([string]$SelectorExpression) {
        return $this.FindElements([Selector]::Parse($SelectorExpression))
    }

    [Object[]]FindElements([string]$SelectorExpression, [SelectorType]$Type) {
        return $this.FindElements([Selector]::New($SelectorExpression, $Type))
    }
    #endregion

    #region Method:IsElementPresent()
    [bool]IsElementPresent([string]$SelectorExpression) {
        [int]$tmpWait = $this.ImplicitWait
        try {
            # Set implicit wait to 0 sec temporally.
            if ($this.Driver) { $this.SetImplicitWait(0) }
            return [bool]($this.FindElement([Selector]::Parse($SelectorExpression)))
        }
        catch {
            return $false
        }
        finally {
            # Reset implicit wait
            if ($this.Driver) { $this.SetImplicitWait($tmpWait) }
        }
    }
    #endregion

    #region Method:Sendkeys() & ClearAndType()
    Hidden [void]_InnerSendKeys([string]$Target, [string]$Value, [bool]$BeforeClear) {
        if ($element = $this.FindElement($Target)) {
            if ($BeforeClear) {
                $element.Clear()
            }
            $private:Ret = $Value
            $private:regex = [regex]'\$\{KEY_.+?\}'
            $regex.Matches($Value) | ForEach-Object {
                $Spec = $this.SpecialKeys.ConvertSeleniumKeys(($_.Value).SubString(2, ($_.Value.length - 3)))
                $Ret = $Ret.Replace($_.Value, $Spec)
            }
            $element.SendKeys($Ret)
        }
    }

    [void]SendKeys([string]$Target, [string]$Value) {
        $this._InnerSendKeys($Target, $Value, $false)
    }

    [void]ClearAndType([string]$Target, [string]$Value) {
        $this._InnerSendKeys($Target, $Value, $true)
    }
    #endregion

    #region Method:Click()
    [void]Click([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $element.Click()
        }
    }
    #endregion

    #region Method:DoubleClick()
    [void]DoubleClick([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $this.Actions.DoubleClick($element).build().perform()
        }
    }
    #endregion

    #region Method:RightClick()
    [void]RightClick([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $this.Actions.ContextClick($element).build().perform()
        }
    }
    #endregion

    #region Method:Select()
    [void]Select([string]$Target, [string]$Value) {
        if ($SelectElement = $this._GetSelectElement($Target)) {
            #TODO: Implement SelectByIndex
            #TODO: Implement SelectByValue
            $SelectElement.SelectByText($Value)
        }
    }
    #endregion

    #region Switch window
    [void]SelectWindow([string]$Title) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            $IsWindowFound = $false
            $Pattern = $this._PerseSeleniumPattern($Title)
            $CurrentWindow = $this.Driver.CurrentWindowHandle
            $AllWindow = $this.Driver.WindowHandles
            #Enumerate all windows
            :SWLOOP foreach ($window in $AllWindow) {
                if ($window -eq $CurrentWindow) {
                    $title = $this.Driver.Title
                }
                else {
                    $title = $this.Driver.SwitchTo().Window($window).Title
                }

                switch ($Pattern.Matcher) {
                    'Like' {
                        if ($title -like $Pattern.Pattern) {
                            $IsWindowFound = $true
                            break SWLOOP
                        }
                    }
                    'RegExp' {
                        if ($title -match $Pattern.Pattern) {
                            $IsWindowFound = $true
                            break SWLOOP
                        }
                    }
                    'Equal' {
                        if ($title -eq $Pattern.Pattern) {
                            $IsWindowFound = $true
                            break SWLOOP
                        }
                    }
                }
            }

            if (!$IsWindowFound) {
                if ($this.Driver.CurrentWindowHandle -ne $CurrentWindow) {
                    #Return current window
                    $this.Driver.SwitchTo().Window($CurrentWindow)
                }
                #throw NoSuchWindowException
                $this.Driver.SwitchTo().Window([System.Guid]::NewGuid().ToString)
            }
        }
    }
    #endregion

    #region Method:SelectFrame()
    # [void]SelectFrame([string]$Name) {
    #     if (!$this.Driver) {
    #         $this._WarnBrowserNotStarted()
    #     }
    #     else {
    #         $this.Driver.SwitchTo().DefaultContent()
    #         $this.Driver.SwitchTo().Frame($Name)
    #     }
    # }
    #endregion

    #region WaitFor*

    #region Method:WaitForElementPresent()
    [bool]WaitForElementPresent([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertElementPresent($Target) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForElementNotPresent([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertElementNotPresent($Target) }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForValue()
    [bool]WaitForValue([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertValue($Target, $Value) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotValue([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertNotValue($Target, $Value) }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForText()
    [bool]WaitForText([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertText($Target, $Value) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotText([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertNotText($Target, $Value) }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForVisible()
    [bool]WaitForVisible([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertVisible($Target) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotVisible([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertNotVisible($Target) }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForTitle()
    [bool]WaitForTitle([string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertTitle($Value) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotTitle([string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertNotTitle($Value) }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion
    #endregion WatFor*

    #region Method:Pause()
    [void]Pause([int]$WaitTimeInMilliSeconds) {
        [System.Threading.Thread]::Sleep($WaitTimeInMilliSeconds)
    }
    #endregion


    #region Method:SaveScreenShot()
    [void]SaveScreenShot([string]$FileName, [ImageFormat]$ImageFormat) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            # Path normalization
            if ($global:PSVersionTable.PSVersion -ge 6.1) {
                $NormalizedFilePath = [System.IO.Path]::GetFullPath($FileName, $PWD)  #This method can be used only on .NET Core 2.1+
            }
            else {
                if (([uri]$FileName).IsAbsoluteUri) {
                    $NormalizedFilePath = [System.IO.Path]::GetFullPath($FileName)
                }
                else {
                    $NormalizedFilePath = [System.IO.Path]::GetFullPath((Join-path $PWD $FileName))
                }
            }
            
            $SaveFolder = Split-Path $NormalizedFilePath -Parent
            if ($null -eq $SaveFolder) {
                throw [System.ArgumentNullException]::new()
            }
            elseif (! (Test-Path -LiteralPath $SaveFolder -PathType Container)) {
                New-Item $SaveFolder -ItemType Directory
            }
            #TODO:To alternate [System.Drawing.Image] class
            Invoke-Expression '$ScreenShot = [Screenshot]$this.Driver.GetScreenShot()'
            Invoke-Expression '$ScreenShot.SaveAsFile($NormalizedFilePath, [ScreenshotImageFormat]$ImageFormat)'
        }
    }

    [void]SaveScreenShot([string]$FileName) {
        [ImageFormat]$Format =
        switch ([System.IO.Path]::GetExtension($FileName)) {
            '.jpg' { [ImageFormat]::Jpeg; break }
            '.jpeg' { [ImageFormat]::Jpeg; break }
            '.bmp' { [ImageFormat]::Bmp; break }
            '.gif' { [ImageFormat]::Gif; break }
            '.tif' { [ImageFormat]::Tiff; break }
            '.tiff' { [ImageFormat]::Tiff; break }
            Default { [ImageFormat]::Png }
        }

        $this.SaveScreenShot($FileName, $Format)
    }
    #endregion

    #region Hidden Method
    Hidden [void]_LoadSelenium() {
        $LibPath = Join-Path $this.PSModuleRoot '\lib'

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
                        throw "Couldn't load $($_.DllName)"
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
                throw "Couldn't load Appium"
            }
        }
    }

    Hidden [string]_FindWebDriver() {
        $ret = $null
        $local:exe = 'WinAppDriver.exe'
        $local:SearchDirectories = @(
            (Join-Path ${env:ProgramFiles(x86)} 'Windows Application Driver')
            (Join-Path $env:ProgramFiles 'Windows Application Driver')
        )

        if ($cmd = Get-Command $local:exe -ErrorAction SilentlyContinue) {
            $ret = $cmd.Source
        }
        else {
            $local:SearchDirectories | % {
                if ($cmd = Resolve-Path (Join-Path $_ $local:exe) -ErrorAction SilentlyContinue) {
                    $ret = $cmd.Path
                }
            }
        }

        if ($ret) {
            return $ret
        }
        else {
            throw "Couldn't find $exe"
        }
    }

    Hidden [void]_StartWinAppDriver() {
        $status = $null
        $StatusUri = [System.UriBuilder]::new($this.WinAppDriverUrl)
        $StatusUri.Path = 'status'
        try { $status = Invoke-WebRequest -Uri $StatusUri.Uri -UseBasicParsing -ErrorAction Stop }catch {}
        if ($status.StatusCode -eq 200) {
            Write-Verbose 'WinAppDriver already started.'
        }
        else {
            $this.AppDriverProcess = Start-Process -FilePath $this.AppDriverExe -PassThru
        }
    }

    Hidden [AppiumOptions]_NewDriverOptions() {
        $Options = [AppiumOptions]::new()
        return $Options
    }

    Hidden [void]_WarnBrowserNotStarted([string]$Message) {
        Write-Warning $Message
    }

    Hidden [void]_WarnBrowserNotStarted() {
        $Message = 'Application is not started.'
        $this._WarnBrowserNotStarted($Message)
    }


    Hidden [bool]_WaitForBase([ScriptBlock]$Expression, [int]$Timeout) {
        if (($Timeout -lt 0) -or ($Timeout -gt 3600)) {
            throw [System.ArgumentOutOfRangeException]::New()
        }

        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $false
        }

        [int]$tmpWait = $this.ImplicitWait
        # Set implicit wait to 0 sec temporally.
        if ($this.Driver) { $this.SetImplicitWait(0) }

        $sec = 0;
        [bool]$ret = $false
        do {
            try {
                $Expression.Invoke()
                $ret = $true
                break
            }
            catch {
                $ret = $false
            }
            if ($sec -ge $Timeout) {
                $ret = $false
                throw 'Timeout'
                break
            }
            [System.Threading.Thread]::Sleep(1000)
            $sec++
        } while ($true)

        if ($this.Driver) { $this.SetImplicitWait($tmpWait) }
        return $ret
    }

    Hidden [HashTable]_PerseSeleniumPattern([string]$Pattern) {
        $local:ret = [HashTable]@{
            Matcher = ''
            Pattern = ''
        }

        switch -Regex ($Pattern) {
            '^regexp:(.+)' {
                $ret.Matcher = 'RegExp'
                $ret.Pattern = $Matches[1]
                break
            }
            '^glob:(.+)' {
                $ret.Matcher = 'Like'
                $ret.Pattern = $Matches[1]
                break
            }
            '^exact:(.+)' {
                $ret.Matcher = 'Equal'
                $ret.Pattern = $Matches[1]
                break
            }
            Default {
                $ret.Matcher = 'Like'
                $ret.Pattern = $Pattern
            }
        }
        return $ret
    }

    Hidden [Object]_GetSelectElement([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $SelectElement = $null
            Invoke-Expression '$SelectElement = New-Object "Support.UI.SelectElement" $element' -ea Stop
            return $SelectElement
        }
        else {
            return $null
        }
    }
    #endregion Hidden Method

    #region Assertion
    [void]AssertElementPresent([string]$Selector) {
        $this.IsElementPresent($Selector) | Assert -Expected $true
    }

    [void]AssertElementNotPresent([string]$Selector) {
        $this.IsElementPresent($Selector) | Assert -Expected $false
    }

    [void]AssertAlertPresent() {
        $this.IsAlertPresent() | Assert -Expected $true
    }

    [void]AssertAlertNotPresent() {
        $this.IsAlertPresent() | Assert -Expected $false
    }

    [void]AssertTitle([string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetTitle() | Assert -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertNotTitle([string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetTitle() | Assert -Not -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertLocation([string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetLocation() | Assert -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertNotLocation([string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetLocation() | Assert -Not -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertAttribute([string]$Target, [string]$Attribute, [string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetAttribute($Target, $Attribute) | Assert -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertNotAttribute([string]$Target, [string]$Attribute, [string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetAttribute($Target, $Attribute) | Assert -Not -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertValue([string]$Target, [string]$Value) {
        $this.AssertAttribute($Target, 'value', $Value)
    }

    [void]AssertNotValue([string]$Target, [string]$Value) {
        $this.AssertNotAttribute($Target, 'value', $Value)
    }

    [void]AssertText([string]$Target, [string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetText($Target) | Assert -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertNotText([string]$Target, [string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetText($Target) | Assert -Not -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertVisible([string]$Target) {
        $this.IsVisible($Target) | Assert -Expected $true
    }

    [void]AssertNotVisible([string]$Target) {
        $this.IsVisible($Target) | Assert -Expected $false
    }
    #endregion
}
#endregion

function New-PSWinAppDriver {
    [CmdletBinding()]
    [OutputType([PSWinAppDriver])]
    param()

    New-Object PSWinAppDriver
}

function New-Selector {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([Selector])]
    param(
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(Mandatory, ParameterSetName = 'WithType')]
        [string]
        $Expression,

        [Parameter(ParameterSetName = 'WithType')]
        [SelectorType]
        $Type = [SelectorType]::None
    )

    if ([string]::IsNullOrEmpty($Expression)) {
        New-Object Selector
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'Default') {
        New-Object Selector($Expression)
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'WithType') {
        New-Object Selector($Expression, $Type)
    }
}

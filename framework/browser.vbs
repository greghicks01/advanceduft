Function launchBrowser(URL)
    Dim IE
    Dim BrowserObj
    Dim scrWidth,scrHeight,strComputer,objWMIService,colCSItems,objCSItem
    
    launchBrowser = False ' assume Fail
    
    If Trim(URL)="" Then
        URL = "about:Blank"
    End If
    
    'this section gets and sets the IE object values to size of screen, then navigates to URL.
    strComputer = "." 'Sets the local pc as source
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 'gets the winmgmts scripting components
    Set colCSItems = objWMIService.ExecQuery("SELECT * FROM Win32_DisplayConfiguration") 'query the display config
    For Each objCSItem In colCSItems 'loop through config setting and set height and width values to variables
        scrWidth = objCSItem.PelsWidth
        scrHeight = objCSItem.PelsHeight
    Next
    Set IE = CreateObject("InternetExplorer.Application")
    ' reset to relative max values
    ' Added 06/08/2007
    If CInt(scrWidth)>(1024) Then
        scrWidth = (1024)
    End If
    scrHeight = scrHeight-44 ' allows for taskbar display under browser edge
    If CInt(scrHeight)>(768) Then
        scrHeight = (768)
    End If
    With IE
        .Visible = True
        .Top = 0
        .Left = 0
        .Height = scrHeight ' adjust for height of Windows Taskbar
        .Width = scrWidth ' browser is resized, but NOT Maximized!
        .Navigate URL
    End With
    ' now let's maximize the browser Windows, if possible (important for Java toolbar functions)
    'maximizeBrowserWindows ' -- still won't work with Web Add-In!
    launchBrowser = True
End Function

SytemUtil.CloseProcessbyWndTitle("")

' https://softwaresupport.softwaregrp.com/doc/KM00765806?fileName=hp_man_UFT12.00_ConnectionAgentUserGuide_pdf.pdf

'https://admhelp.microfocus.com/uft/jp/15.0/UFT_Help/Content/Addins_Guide/Wrk_Multi_Browsers_How2.htm
' Prerequisite- turn off auto updates for the browsers
' Configure the Record and Run settings to launch a browser
' Use the BROWSER_ENV environment variable to launch a browser
' Launch a browser with a test parameter
' Launch a browser using a data table parameter
' Launch a Browser using a WebUtil.LaunchBrowser step
' Dynamically load an object repository during the test run
' Add steps for browser specific behavior

BROWSER_ENV variable
IE	Opens the installed version of Internet Explorer.
IE64	Opens the installed 64-bit version of Internet Explorer.
CHROME	Opens the installed version of Google Chrome.
CHROME_HEADLESS	
Opens the locally installed version of Headless Chrome.
Supported for Chrome versions 60 and higher.
FIREFOX	Opens the installed version of Firefox.
FIREFOX64 Opens the latest version of 64-bit Mozilla Firefox that is both installed on the computer and supported by UFT.
SAFARI	Opens Safari on the remote Mac computer connected to UFT.
EDGE	Opens the installed version of Microsoft Edge with the Edge Agent for Functional Testing already enabled.
CHROME_EMULATOR	Opens Chrome in emulated mode with the specified device.
PHANTOMJS	Opens the locally installed version of the PhantomJS web toolkit.

For a Global Data Table parameter	In the Global tab of the Data pane, set the value of the parameter.	
Value to use:

IE. Opens Internet Explorer.

IE64. Opens a 64-bit version of Internet Explorer.

CHROME. Opens Google Chrome.

CHROME_HEADLESS. Opens Headless Chrome.
FIREFOX. Opens Mozilla Firefox.

FIREFOX64. Opens the latest version of 64-bit Mozilla Firefox that is both installed on the computer and supported by UFT.

SAFARI. Opens Safari on the remote Mac computer connected to UFT (defined in the Web tab of the Record and Run Settings dialog box or in the REMOTE_HOST environment variable).

EDGE. Opens the installed version of Microsoft Edge with the Edge Agent for Functional Testing already enabled.

CHROME_EMULATOR. Opens Chrome in emulated mode with the specified device.

PHANTOMJS: Opens the locally installed version of the PhantomJS web toolkit.
For details on adding mobile-relevant parameters, see ?Define Mobile Record and Run Settings?.

Launch a browser using a data table parameter
(Optional) Create a reusable action to use in all your tests for launching the browsers.

In the Data pane, open the Global tab.

In the Global tab, click the first row of the column where you want to store the parameter.

Enter the name for the parameter and click OK.

For example, you could name this parameter BrowserName (to identify it as the name of the browser to open).

In the data table, enter the .exe names for the browsers you want to open.

For example, if you need to run the test on Internet Explorer, Firefox, and Chrome, you would enter iexplore.exe, firefox.exe, and chrome.exe in the first three rows of the column, respectively:

Launch a Browser using a WebUtil.LaunchBrowser step
If you are using Business Process Testing to test your web applications, add use a WebUtil.LaunchBrowser step to launch the appropriate browsers as needed within each component.

Provide an argument for the Browser parameter, which is the same as the environment variables for Web- based environments:
CHROME
CHROME_EMULATOR
CHROME_HEADLESS
EDGE
FIREFOX
FIREFOX64
IE
IE64
PHANTOMJS
SAFARI

Dynamically load an object repository during the test run
If your test requires you to have different object repositories for each browser type, load the relevant object repositories as part of the test run without having to manually configure anything before the test run:

In the Data pane, open the Global tab.

In the Global tab, click the first row of the column where you want to store the parameter.

Enter the name for the parameter and click OK.

For example, you could name this parameter Browser (to identify it as the name of the browser on which to run the test).

In the data table, enter the names for the browsers on which you want to run the test.

Add a test step with the following format:

If DataTable("<data table parameter>") = <Browser 1> Then
	RepositoriesCollection.Add "<location to object repository>"
ElseIf DataTable("<data table parameter>") = <Browser 2> Then
	RepositoriesCollection.Add "<location to object repository>"
End If

Add steps for browser specific behavior
If you need to add steps to perform browser specific behavior in the course of the test, use test parameters to create steps for this behavior.

In the canvas, select an action.

In the Properties pane, select the Parameters tab.

In the Parameters tab, click the Add button.

In the parameters grid, provide a name for the parameter. For example, you could name the parameter ActiveBrowser to show that the value of the parameter represents the browser currently in use.

Add steps to the test. You can use the value of the parameter by using the Parameter object:

Select Case Parameter("<parameter name>")
	Case "<Browser 1>"
	'Do something specific for browser 1
	Case "<Browser 2>"
	'Do something specific for browser 2
End Select

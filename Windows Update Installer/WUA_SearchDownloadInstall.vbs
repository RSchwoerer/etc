
' http://msdn.microsoft.com/en-us/library/aa387102%28VS.85%29.aspx

'--- force this to be run with CScript (command line) -------------------------------------------------------------
Sub forceCScriptExecution
	'http://stackoverflow.com/a/5219524/504398
    Dim Arg, Str
    If Not LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
            If InStr( Arg, " " ) Then Arg = """" & Arg & """"
            Str = Str & " " & Arg
        Next
        CreateObject( "WScript.Shell" ).Run _
            "cscript //nologo """ & _
            WScript.ScriptFullName & _
            """ " & Str
        WScript.Quit
    End If
End Sub
forceCScriptExecution


WScript.Echo "Starting..."
Set updateSession = CreateObject("Microsoft.Update.Session")
updateSession.ClientApplicationID = "MSDN Sample Script"


'--- search for updates -------------------------------------------------------------------------------------------
WScript.Echo "Searching for updates..." & vbCRLF
Set updateSearcher = updateSession.CreateUpdateSearcher()
Set searchResult = updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

WScript.Echo "Found " & searchResult.Updates.Count & " updates."
WScript.Echo "List of applicable items on the machine:"

If searchResult.Updates.Count = 0 Then
    WScript.Echo "There are no applicable updates."
	WScript.Slee 5000
    WScript.Quit	'QUIT!
End If

For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    WScript.Echo "    " & I + 1 & "> " & update.Title
Next

'--- create collection of updates ---------------------------------------------------------------------------------
WScript.Echo vbCRLF & "Creating collection of updates to download:"
Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For I = 0 to searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    addThisUpdate = false
    If update.InstallationBehavior.CanRequestUserInput = true Then
        WScript.Echo "    " & I + 1 & "> skipping: " & update.Title & _
        " because it requires user input"
    Else
        If update.EulaAccepted = false Then
            REM WScript.Echo I + 1 & "> note: " & update.Title & _
            REM " has a license agreement that must be accepted:"
            REM WScript.Echo update.EulaText
            REM WScript.Echo "Do you accept this license agreement? (Y/N)"
            REM strInput = WScript.StdIn.Readline
            REM WScript.Echo 
            REM If (strInput = "Y" or strInput = "y") Then
                update.AcceptEula()
                addThisUpdate = true
            REM Else
                REM WScript.Echo I + 1 & "> skipping: " & update.Title & _
                REM " because the license agreement was declined"
            REM End If
        Else
            addThisUpdate = true
        End If
    End If
    If addThisUpdate = true Then
        REM WScript.Echo I + 1 & "> adding: " & update.Title 
        updatesToDownload.Add(update)
    End If
Next

If updatesToDownload.Count = 0 Then
    WScript.Echo "All applicable updates were skipped."
    WScript.Quit
End If
    
'--- download updates ---------------------------------------------------------------------------------------------
WScript.Echo vbCRLF & "Downloading updates..."
Set downloader = updateSession.CreateUpdateDownloader() 
downloader.Updates = updatesToDownload
downloader.Download()

Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")
rebootMayBeRequired = false
WScript.Echo vbCRLF & "Successfully downloaded updates:"

For I = 0 To searchResult.Updates.Count-1
    set update = searchResult.Updates.Item(I)
    If update.IsDownloaded = true Then
        REM WScript.Echo I + 1 & "> " & update.Title 
        updatesToInstall.Add(update)	
        If update.InstallationBehavior.RebootBehavior > 0 Then
            rebootMayBeRequired = true
        End If
    End If
Next

If updatesToInstall.Count = 0 Then
    WScript.Echo "No updates were successfully downloaded."
    WScript.Quit
End If

If rebootMayBeRequired = true Then
    WScript.Echo vbCRLF & "These updates may require a reboot."
End If


'--- install updates ----------------------------------------------------------------------------------------------
' don't prompt, just install.
REM WScript.Echo  vbCRLF & "Would you like to install updates now? (Y/N)"
REM strInput = WScript.StdIn.Readline
REM WScript.Echo 

REM If (strInput = "Y" or strInput = "y") Then
    WScript.Echo "Installing updates..."
    Set installer = updateSession.CreateUpdateInstaller()
    installer.Updates = updatesToInstall
    Set installationResult = installer.Install()
	
    'Output results of install
    WScript.Echo "Installation Result: " & _
    installationResult.ResultCode 
    WScript.Echo "Reboot Required: " & _ 
    installationResult.RebootRequired & vbCRLF 
    WScript.Echo "Listing of updates installed " & _
    "and individual installation results:" 
	
    For I = 0 to updatesToInstall.Count - 1
        REM WScript.Echo I + 1 & "> " & updatesToInstall.Item(i).Title & ": " & installationResult.GetUpdateResult(i).ResultCode 		
    Next
REM End If

WScript.Echo "Done."
WScript.Quit

:ERROR
WScript.Echo err.description


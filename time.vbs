Function IEButtons( )
    ' This function uses Internet Explorer to create a dialog.
    Dim objIE, sTitle, iErrorNum

    ' Create an IE object
    Set objIE = CreateObject( "InternetExplorer.Application" )
    ' specify some of the IE window's settings
    objIE.Navigate "about:blank"
    sTitle="Make your choice " & String( 80, "." ) 'Note: the String( 80,".") is to push "Internet Explorer" string off the window
    objIE.Document.title = sTitle
    objIE.MenuBar        = False
    objIE.ToolBar        = False
    objIE.AddressBar     = false
    objIE.Resizable      = False
    objIE.StatusBar      = False
    objIE.Width          = 350
    objIE.Height         = 500
    ' Center the dialog window on the screen
    With objIE.Document.parentWindow.screen
        objIE.Left = (.availWidth  - objIE.Width ) \ 2
        objIE.Top  = (.availHeight - objIE.Height) \ 2
    End With
    ' Wait till IE is ready
    Do While objIE.Busy
        WScript.Sleep 200
    Loop
    

    ' Insert the HTML code to prompt for user input
    objIE.Document.body.innerHTML = "<div align=""center""><h1>Unfortunately, the clock is ticking</h1><br /><p>the hours are going by. The past increases, the future recedes. Possibilities decreasing, regrets mounting<br />Do you understand?</p>" & vbcrlf _
                                  & "<p><input type=""hidden"" id=""OK"" name=""OK"" value=""0"">" _
                                  & "<input type=""submit"" value=""  I Understand   "" onClick=""window.close()"" style="" border-style: double;display:inline-block;"">" _
                                  & "<input type=""submit"" value=""  remain ignorant "" onClick=""window.close()"" style="" border-style:solid;display:inline-block;""></p></div>"

    ' Hide the scrollbars
    objIE.Document.body.style.overflow = "auto"
    ' Make the window visible
    objIE.Visible = True


    
End Function
IEButtons( )

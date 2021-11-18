Sub InternetExplorerObject()

Dim IEObject As InternetExplorer

'Create a new instance of the Internet Explorer Object
Set IEObject = New InternetExplorer

'Switch to see IE window appear during scraping
IEObject.Visible = True
    
Dim c As Range
Dim fee As Object

rw = 2

For Each c In ActiveSheet.Range("A2", ActiveSheet.Cells(Rows.Count, "A").End(xlUp))
    
    Application.StatusBar = "Getting... " & c
    
    'Navigate to a URL we specify.
    IEObject.Navigate Url:="https://www.fundinfo.com/en/LU-prof/LandingPage?query=" & c & "#tab=1" ', Flags:=navOpenInNewWindow
    
    'This Loop will keep us waiting as long as the IEObject is in a Busy state or
    'the ReadyState does not communicate complete.
    Do While IEObject.Busy = True Or IEObject.ReadyState <> READYSTATE_COMPLETE
       
       'Wait one second, and then try again
        Application.Wait Now + TimeValue("00:00:01")
       
    Loop
    
    'Print the URL we are currently at and row number.
    Debug.Print IEObject.LocationURL
    Debug.Print rw
    
       
    'Get the HTML document for the page
    Dim IEDocument As HTMLDocument
    Set IEDocument = IEObject.Document
    
    'Grab a elements collection
    Dim IEElements As IHTMLElementCollection
    Set IEElements = IEDocument.getElementsByClassName("member-only OFST452200")
        
    'skip if ISIN doesn't exist (i.e. tag doesn't exist in page)
    On Error Resume Next
        Debug.Print IEElements.Item.innerText
        ActiveSheet.Cells(rw, 2).Value = IEElements.Item.innerText
    
IEObject.Quit
Set IEObject = Nothing
    
Application.StatusBar = ""

rw = rw + 1

Next
    
Application.StatusBar = ""
    
End Sub


Public Function FundInfoMF(ISIN)


Dim IEObject As InternetExplorer

'Create a new instance of the Internet Explorer Object
Set IEObject = New InternetExplorer

'Switch to see IE window appear during scraping
IEObject.Visible = False
    
Dim fee As Object

Application.StatusBar = "Getting Mgmt Fee for " & ISIN

'Navigate to a URL we specify.
IEObject.Navigate Url:="https://www.fundinfo.com/en/LU-prof/LandingPage?query=" & ISIN & "#tab=1" ', Flags:=navOpenInNewWindow

'This Loop will keep us waiting as long as the IEObject is in a Busy state or
'the ReadyState does not communicate complete.
Do While IEObject.Busy = True Or IEObject.ReadyState <> READYSTATE_COMPLETE
   
   'Wait one second, and then try again
    Application.Wait Now + TimeValue("00:00:01")
   
Loop

'Print the URL we are currently at and row number.
'Debug.Print IEObject.LocationURL

'Get the HTML document for the page
Dim IEDocument As HTMLDocument
Set IEDocument = IEObject.Document

'Grab a elements collection
Dim IEElements As IHTMLElementCollection
Set IEElements = IEDocument.getElementsByClassName("member-only OFST452000")
    
FundInfoMF = IEElements.Item.innerText
Debug.Print FundInfoMF



IEObject.Quit

Set IEObject = Nothing

Application.StatusBar = ""

End Function

Public Function FundInfoOGC(ISIN)


Dim IEObject As InternetExplorer

'Create a new instance of the Internet Explorer Object
Set IEObject = New InternetExplorer

'Switch to see IE window appear during scraping
IEObject.Visible = False
    
Dim fee As Object

Application.StatusBar = "Getting Mgmt Fee for " & ISIN

'Navigate to a URL we specify.
IEObject.Navigate Url:="https://www.fundinfo.com/en/LU-prof/LandingPage?query=" & ISIN & "#tab=1" ', Flags:=navOpenInNewWindow

'This Loop will keep us waiting as long as the IEObject is in a Busy state or
'the ReadyState does not communicate complete.
Do While IEObject.Busy = True Or IEObject.ReadyState <> READYSTATE_COMPLETE
   
   'Wait one second, and then try again
    Application.Wait Now + TimeValue("00:00:01")
   
Loop

'Print the URL we are currently at and row number.
'Debug.Print IEObject.LocationURL

'Get the HTML document for the page
Dim IEDocument As HTMLDocument
Set IEDocument = IEObject.Document

'Grab a elements collection
Dim IEElements As IHTMLElementCollection
Set IEElements = IEDocument.getElementsByClassName("member-only OFST452200")
    
FundInfoOGC = IEElements.Item.innerText
Debug.Print FundInfoOGC



IEObject.Quit

Set IEObject = Nothing

Application.StatusBar = ""

End Function

Public Function FundInfoTER(ISIN)

Dim IEObject As InternetExplorer

'Create a new instance of the Internet Explorer Object
Set IEObject = New InternetExplorer

'Switch to see IE window appear during scraping
IEObject.Visible = False
    
Dim fee As Object

Application.StatusBar = "Getting Mgmt Fee for " & ISIN

'Navigate to a URL we specify.
IEObject.Navigate Url:="https://www.fundinfo.com/en/LU-prof/LandingPage?query=" & ISIN & "#tab=1" ', Flags:=navOpenInNewWindow

'This Loop will keep us waiting as long as the IEObject is in a Busy state or
'the ReadyState does not communicate complete.
Do While IEObject.Busy = True Or IEObject.ReadyState <> READYSTATE_COMPLETE
   
   'Wait one second, and then try again
    Application.Wait Now + TimeValue("00:00:01")
   
Loop

'Print the URL we are currently at and row number.
'Debug.Print IEObject.LocationURL

'Get the HTML document for the page
Dim IEDocument As HTMLDocument
Set IEDocument = IEObject.Document

'Grab a elements collection
Dim IEElements As IHTMLElementCollection
Set IEElements = IEDocument.getElementsByClassName("member-only OFST452100")
    
FundInfoTER = IEElements.Item.innerText
Debug.Print FundInfoTER



IEObject.Quit

Set IEObject = Nothing
Set IEElements = Nothing


Application.StatusBar = ""

End Function

Sub clearstatus()

Application.StatusBar = ""

End Sub


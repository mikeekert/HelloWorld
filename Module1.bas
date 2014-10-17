Attribute VB_Name = "Module1"
Const MAX_PATH = 260

Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

'test comment
'test 2

Private Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type

 Private Declare Function InternetOpen _
   Lib "wininet.dll" _
     Alias "InternetOpenA" _
       (ByVal sAgent As String, _
        ByVal lAccessType As Long, _
        ByVal sProxyName As String, _
        ByVal sProxyBypass As String, _
        ByVal lFlags As Long) As Long

 Private Declare Function InternetConnect _
   Lib "wininet.dll" _
     Alias "InternetConnectA" _
       (ByVal hInternetSession As Long, _
        ByVal sServerName As String, _
        ByVal nServerPort As Integer, _
        ByVal sUsername As String, _
        ByVal sPassword As String, _
        ByVal lService As Long, _
        ByVal lFlags As Long, _
        ByVal lContext As Long) As Long

 Private Declare Function FtpGetFile _
   Lib "wininet.dll" _
     Alias "FtpGetFileA" _
       (ByVal hFtpSession As Long, _
        ByVal lpszRemoteFile As String, _
        ByVal lpszNewFile As String, _
        ByVal fFailIfExists As Boolean, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Boolean

 Private Declare Function FtpPutFile _
   Lib "wininet.dll" _
     Alias "FtpPutFileA" _
       (ByVal hFtpSession As Long, _
        ByVal lpszLocalFile As String, _
        ByVal lpszRemoteFile As String, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Boolean
        
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias _
 "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) _
 As Boolean
 
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
    (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
    lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, _
    ByVal dwContent As Long) As Long


 Private Declare Function InternetCloseHandle _
   Lib "wininet.dll" _
     (ByVal hInet As Long) As Integer
     Option Explicit

Sub AddMenuItems()

   Dim menuBar As CommandBar
   Dim newMenu As CommandBarControl
   Dim menuItem As CommandBarControl
   Dim subMenuItem As CommandBarControl
   
   Set menuBar = CommandBars.Add(menuBar:=True, Position:=msoBarTop, Name:="Sub Menu Bar", Temporary:=True)
   menuBar.Visible = True
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&#"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "24 Hour Fitness"
        .OnAction = "Fitness"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "3M"
        .OnAction = "MMM"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&A"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Active Network"
        .OnAction = "ActiveNetwork"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Allied Barton"
        .OnAction = "AlliedBarton"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "American Family AFI"
        .OnAction = "AmericanFamilyAFI"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "American Family AGT"
        .OnAction = "AmericanFamilyAGT"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Amtrak"
        .OnAction = "Amtrak"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Avanade"
        .OnAction = "Avanade"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&B"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Bell"
        .OnAction = "Bell"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Bloomberg"
        .OnAction = "Bloomberg"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "BNY Mellon"
        .OnAction = "BNYMellon"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Bombardier"
        .OnAction = "Bombardier"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Bombardier Transportation"
        .OnAction = "BombardierTransportation"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "BonSecours"
        .OnAction = "BonSecours"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Booz"
        .OnAction = "Booz"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Bristol Myers"
        .OnAction = "BristolMyers"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Burberry"
        .OnAction = "Burberry"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&C"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "CACI"
        .OnAction = "CACI"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Carolina"
        .OnAction = "Carolina"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Celgene"
        .OnAction = "Celgene"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Christiana Care"
        .OnAction = "ChristianaCare"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Cintas"
        .OnAction = "Cintas"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Citrix"
        .OnAction = "Citrix"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Cleveland Clinic"
        .OnAction = "ClevelandClinic"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Comerica"
        .OnAction = "Comerica"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "ConAgra"
        .OnAction = "ConAgra"
        End With
                   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "CUNA"
        .OnAction = "CUNA"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&D"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Deluxe"
        .OnAction = "Deluxe"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "DSW"
        .OnAction = "DSW"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&E"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Ecolab"
        .OnAction = "Ecolab"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Ericcson"
        .OnAction = "Ericcson"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Expedia"
        .OnAction = "Expedia"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&F"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "First Citizens Bank"
        .OnAction = "FirstCitizensBank"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "FutureStep"
        .OnAction = "FutureStep"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&G"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Garda Aviation"
        .OnAction = "GardaAviation"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Garda Cash Services"
        .OnAction = "GardaCashServices"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Garda HR and Protective Services"
        .OnAction = "GardaHRProtectiveServices"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Gates Foundation"
        .OnAction = "GatesFoundation"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Genuine Parts"
        .OnAction = "GenuineParts"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Golder"
        .OnAction = "Golder"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Good Life"
        .OnAction = "GoodLife"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Groupon"
        .OnAction = "Groupon"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&H"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Hasbro"
        .OnAction = "Hasbro"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "HD Supply"
        .OnAction = "HDSupply"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "HealthONE"
        .OnAction = "HealthOne"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&I"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Integrity Staffing"
        .OnAction = "IntegrityStaffing"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&J"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "JCPenney"
        .OnAction = "JCPenney"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Johnson and Johnson"
        .OnAction = "JohnsonandJohnson"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&K"
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "KPMG Grad"
        .OnAction = "KPMGGrad"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "KPMG UK"
        .OnAction = "KPMGUK"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Kroger"
        .OnAction = "Kroger"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&L"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Leidos"
        .OnAction = "Leidos"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Lifetime Fitness"
        .OnAction = "LifetimeFitness"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&M"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Marathon Oil"
        .OnAction = "MarathonOil"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Mayo"
        .OnAction = "Mayo"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "McGladrey"
        .OnAction = "McGladrey"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Medtronic"
        .OnAction = "Medtronic"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Merck"
        .OnAction = "Merck"
        End With
        
         Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Micron"
        .OnAction = "Micron"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&N"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Nalco"
        .OnAction = "Nalco"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "New York Times"
        .OnAction = "NewYorkTimes"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Novo Nordisk DK"
        .OnAction = "NovoNordiskDK"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Novo Nordisk US"
        .OnAction = "NovoNordiskUS"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "North Highland"
        .OnAction = "NorthHighland"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&O"
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&P"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Panalpina"
        .OnAction = "Panalpina"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Pearson"
        .OnAction = "Pearson"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&Q"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "QHR File Checker"
        .OnAction = "QHRFileChecker"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&R"
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "RGF"
        .OnAction = "RGF"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Robert Half"
        .OnAction = "RobertHalf"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Rogers"
        .OnAction = "Rogers"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&S"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "SAIC"
        .OnAction = "SAIC"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Sandisk"
        .OnAction = "Sandisk"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Seasons"
        .OnAction = "Seasons"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Sodexo"
        .OnAction = "Sodexo"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Sprint"
        .OnAction = "Sprint"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "SRC"
        .OnAction = "SRC"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Successfactors"
        .OnAction = "Successfactors"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Sun America Financial Group"
        .OnAction = "SAFG"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&T"
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Tesoro"
        .OnAction = "Tesoro"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Texas Health"
        .OnAction = "TexasHealth"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Time Warner"
        .OnAction = "TimeWarner"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Transystems"
        .OnAction = "Transystems"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Trizetto"
        .OnAction = "Trizetto"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&U"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "UMMC"
        .OnAction = "UMMC"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Unisys"
        .OnAction = "Unisys"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "US Cellular"
        .OnAction = "USCellular"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "UTMB"
        .OnAction = "UTMB"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "UUHC"
        .OnAction = "UUHC"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&V"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Visa"
        .OnAction = "Visa"
        End With
        
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&W"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Welch Allyn"
        .OnAction = "WelchAllyn"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Wyndham"
        .OnAction = "Wyndham"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&X"
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&Y"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Yellow Pages"
        .OnAction = "YellowPages"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Yoh"
        .OnAction = "Yoh"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Yum"
        .OnAction = "Yum"
        End With
   
   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&Z"

   Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
   newMenu.Caption = "&Job Patch Macros"
   
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Allied Barton"
        .OnAction = "AlliedBartonJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "American Family AFI"
        .OnAction = "AmericanFamilyAFIJobPatch"
        End With
                
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "American Family AGT"
        .OnAction = "AmericanFamilyAGTJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Bell"
        .OnAction = "BellJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Bombardier"
        .OnAction = "BombardierJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Booz Allen"
        .OnAction = "BoozAllenJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "CACI"
        .OnAction = "CACIJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Carolina"
        .OnAction = "CarolinaJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Celgene"
        .OnAction = "CelgeneJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
        
        With menuItem
        .Caption = "Comerica"
        .OnAction = "ComericaJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
        
        With menuItem
        .Caption = "ConAgra"
        .OnAction = "ConAgraJobPatch"
        End With
        
         Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
         
         With menuItem
        .Caption = "Crossmark"
        .OnAction = "CrossmarkJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
         
         With menuItem
        .Caption = "DSW"
        .OnAction = "DSWJobPatch"
        End With
        
         Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Expedia"
        .OnAction = "ExpediaJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Garda"
        .OnAction = "GardaJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Gates Foundation"
        .OnAction = "GatesfoundationJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Genuine Parts"
        .OnAction = "GenuinePartsJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Golder"
        .OnAction = "GolderJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Goodlife Fitness"
        .OnAction = "GoodlifeJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Groupon"
        .OnAction = "GrouponJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Hasbro"
        .OnAction = "HasbroJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "KPMG"
        .OnAction = "KPMGJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "KPMG UK"
        .OnAction = "KPMGUKJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Leidos"
        .OnAction = "LeidosJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Lifetime Fitness"
        .OnAction = "LifetimeFitnessJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "McGladrey"
        .OnAction = "McGladreyJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Micron"
        .OnAction = "MicronJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Nalco"
        .OnAction = "NalcoJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "North Highland"
        .OnAction = "NorthHighlandJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "RGF"
        .OnAction = "RGFJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Robert Half"
        .OnAction = "RobertHalfJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Rogers"
        .OnAction = "RogersJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "SAIC"
        .OnAction = "SAICJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Sandisk"
        .OnAction = "SandiskJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Seasons"
        .OnAction = "SeasonsJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Sprint"
        .OnAction = "SprintJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "SRC"
        .OnAction = "SRCJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Sun America Financial Group"
        .OnAction = "SAFGJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
        
        With menuItem
        .Caption = "Tesoro"
        .OnAction = "TesoroJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
        
        With menuItem
        .Caption = "Time Warner"
        .OnAction = "TimeWarnerJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Trizetto"
        .OnAction = "TrizettoJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Unisys"
        .OnAction = "UnisysJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "US Cellular"
        .OnAction = "USCellularJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "UTMB"
        .OnAction = "UTMBJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "UUHC"
        .OnAction = "UUHCJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Visa"
        .OnAction = "VisaJobPatch"
        End With
        
        Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
   
        With menuItem
        .Caption = "Wyndham"
        .OnAction = "WyndhamJobPatch"
        End With

End Sub


Function CheckForMenu(argCaption) As Boolean

    Dim bar As CommandBarPopup, Result As Boolean
    
    Result = False
    
    With CommandBars("Worksheet Menu Bar")
        For Each bar In .Controls
            If bar.Caption = argCaption Then
                Result = True
            End If
        Next bar
    End With

    CheckForMenu = Result
    
End Function

Sub Auto_Open()
Call AddMenuItems
End Sub

Sub Medtronic()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:2").Delete
Range("G:I").Delete
Range("C:C").Delete
Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Original Status"
Range("E1").Value = "Date"
Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("H:H")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("B1").Value = "Status"
Range("H:H").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

Dim OriginalStatus
OriginalStatus = Range("H" & CurRow).Value

Select Case OriginalStatus
Case "Inbox"
Range("B" & CurRow).Value = "ATS Capture"
Case "Did Not Meet Basic Qualifications"
Range("B" & CurRow).Value = "Apply Completed"
Case "Ineligible For Consideration"
Range("B" & CurRow).Value = "Apply Completed"
Case "Internal RFT Notification"
Range("B" & CurRow).Value = "Apply Completed"
Case "Questionnaire - Did Not Meet Requirments"
Range("B" & CurRow).Value = "Apply Completed"
Case "Questionnaire Temp"
Range("B" & CurRow).Value = "Apply Completed"
Case "Assessment Complete"
Range("B" & CurRow).Value = "Qualified"
Case "Assessment in Progress"
Range("B" & CurRow).Value = "Qualified"
Case "Candidate Withdrew"
Range("B" & CurRow).Value = "Qualified"
Case "External Request To Apply"
Range("B" & CurRow).Value = "Qualified"
Case "Hiring Manager Screen"
Range("B" & CurRow).Value = "Qualified"
Case "Met Basic Qualifications"
Range("B" & CurRow).Value = "Qualified"
Case "Not Selected for Interview"
Range("B" & CurRow).Value = "Qualified"
Case "Recruiter Screen"
Range("B" & CurRow).Value = "Qualified"
Case "Sent to Hiring Manager"
Range("B" & CurRow).Value = "Qualified"
Case "Interviews Scheduled"
Range("B" & CurRow).Value = "Interviewed"
Case "Not Selected After Interviews"
Range("B" & CurRow).Value = "Interviewed"
Case "Reference Checks"
Range("B" & CurRow).Value = "Interviewed"
Case "Schedule Interviews"
Range("B" & CurRow).Value = "Interviewed"
Case "Offer Declined"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Extended"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Rescinded"
Range("B" & CurRow).Value = "Offer Made"
Case "Selected for Position"
Range("B" & CurRow).Value = "Offer Made"
Case "Hired"
Range("B" & CurRow).Value = "Hired"
Case "Offer Accepted"
Range("B" & CurRow).Value = "Hired"
End Select

CurRow = CurRow + 1

Loop

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit

Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub MMM()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:A").Delete
Range("B:I").Delete
Range("C:E").Delete
Range("D:I").Delete

Range("A1").Value = "Email"
Range("B1").Value = "App. Current CSW Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"

Range("B:B").Cut Destination:=Range("H:H")
Range("B1").Value = "Status"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

Dim OriginalStatus
OriginalStatus = Range("H" & CurRow).Value

Select Case OriginalStatus
Case "Checks"
Range("B" & CurRow).Value = "Apply Completed"
Case "Event Review"
Range("B" & CurRow).Value = "Apply Completed"
Case "Review"
Range("B" & CurRow).Value = "Apply Completed"
Case "Global Pipeline"
Range("B" & CurRow).Value = "Apply Completed"
Case "Hiring Manager Review"
Range("B" & CurRow).Value = "Qualified"
Case "Phone Screen"
Range("B" & CurRow).Value = "Qualified"
Case "Match"
Range("B" & CurRow).Value = "Qualified"
Case "Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "Offer"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Checks"
Range("B" & CurRow).Value = "Offer Made"
Case "Hire"
Range("B" & CurRow).Value = "Hired"
End Select

CurRow = CurRow + 1

Loop

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Application.DisplayAlerts = False
Sheets(1).Delete
Sheets(1).Name = "Sheet1"
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub TexasHealth()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:A").Delete
Range("B:B").Delete
Range("B:B").Delete
Range("C:E").Delete

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Date"
Range("E1").Value = "Original Status"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

Dim OriginalStatus
OriginalStatus = Range("H" & CurRow).Value

Select Case OriginalStatus
Case "Candidate Withdrew"
Range("B" & CurRow).Value = "Apply Completed"
Case "Did Not Possess Basic Qual/s"
Range("B" & CurRow).Value = "Apply Completed"
Case "Inbox"
Range("B" & CurRow).Value = "Apply Completed"
Case "Legal/Policy Review"
Range("B" & CurRow).Value = "Apply Completed"
Case "Not Selected For Interview"
Range("B" & CurRow).Value = "Apply Completed"
Case "Not Suitable"
Range("B" & CurRow).Value = "Apply Completed"
Case "Notify Not Selected For Interview"
Range("B" & CurRow).Value = "Apply Completed"
Case "Hiring Manager Review"
Range("B" & CurRow).Value = "Qualified"
Case "Internal Candidate Review"
Range("B" & CurRow).Value = "Qualified"
Case "Org/Cultural Fit Assessment"
Range("B" & CurRow).Value = "Qualified"
Case "Org/Cultural Fit Assessment in Progress"
Range("B" & CurRow).Value = "Qualified"
Case "Org/Culture Fit Assessment Review"
Range("B" & CurRow).Value = "Qualified"
Case "Qualify Applicants"
Range("B" & CurRow).Value = "Qualified"
Case "Schedule Interview"
Range("B" & CurRow).Value = "Qualified"
Case "Top Candidates"
Range("B" & CurRow).Value = "Qualified"
Case "Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "Not Selected after Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "Recruiter Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "Create Offer"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Accepted"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Declined"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Pending"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Rescinded"
Range("B" & CurRow).Value = "Offer Made"
Case "Hired"
Range("B" & CurRow).Value = "Hired"
Case "Internal Hire"
Range("B" & CurRow).Value = "Hired"
Case "Internal OnBoarding"
Range("B" & CurRow).Value = "Hired"
Case "OnBoarding"
Range("B" & CurRow).Value = "Hired"
Case "Developing Performer Inbox"
 Range("B" & CurRow).Value = "Apply Completed"
Case "Did Not Pass Internal Candidate Review"
 Range("B" & CurRow).Value = "Apply Completed"
Case "Did Not Pass Preferred Qualification Screen"
 Range("B" & CurRow).Value = "Apply Completed"
Case "Did Not Pass Telephone Screen"
 Range("B" & CurRow).Value = "Interviewed"
Case "External Candidate OnBoarding"
 Range("B" & CurRow).Value = "Hired"
Case "Hiring Manager Interview"
 Range("B" & CurRow).Value = "Interviewed"
Case "Internal Candidate OnBoarding"
 Range("B" & CurRow).Value = "Hired"
Case "Internal Inbox"
 Range("B" & CurRow).Value = "Apply Completed"
Case "Notify Not Selected After Interview"
 Range("B" & CurRow).Value = "Interviewed"
Case "Offer Approval Pending"
 Range("B" & CurRow).Value = "Offer Made"
Case "Offer Declined After Acceptance"
 Range("B" & CurRow).Value = "Offer Made"
Case "Offer Extended"
 Range("B" & CurRow).Value = "Offer Made"
Case "Significant Performer Inbox"
 Range("B" & CurRow).Value = "Apply Completed"
Case "Staged Candidates"
 Range("B" & CurRow).Value = "Qualified"
Case "Talentmine Routing Folder"
 Range("B" & CurRow).Value = "Apply Completed"
Case "Telephone Screen"
 Range("B" & CurRow).Value = "Interviewed"
Case "Top Performer Inbox"
 Range("B" & CurRow).Value = "Apply Completed"
End Select

CurRow = CurRow + 1

Loop

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub BristolMyers()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:5").Delete
Range("F:O").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim LastRowSheet2
LastRowSheet2 = Sheets(2).Range("A65536").End(xlUp).Row

Dim LastRowSheet3
LastRowSheet3 = Sheets(3).Range("A65536").End(xlUp).Row

Sheets(2).Range("A6:F" & LastRowSheet2).Cut Destination:=Sheets(1).Range("A" & LastRow + 1)
LastRow = Range("A65536").End(xlUp).Row

Sheets(3).Range("A6:F" & LastRowSheet3).Cut Destination:=Sheets(1).Range("A" & LastRow + 1)
LastRow = Range("A65536").End(xlUp).Row

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Status"
Range("E1").Value = "Date"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

Dim OriginalStatus
OriginalStatus = Range("H" & CurRow).Value

Select Case OriginalStatus
Case "New -  Rejected"
Range("B" & CurRow).Value = "Apply Completed"
Case "New -  To Be Evaluated"
Range("B" & CurRow).Value = "Apply Completed"
Case "New -  Withdrawn"
Range("B" & CurRow).Value = "Apply Completed"
Case "New - Rejected"
Range("B" & CurRow).Value = "Apply Completed"
Case "New - To Be Evaluated"
Range("B" & CurRow).Value = "Apply Completed"
Case "New - Withdrawn"
Range("B" & CurRow).Value = "Apply Completed"
Case "Hiring Manager Screen - HM Review Passed"
Range("B" & CurRow).Value = "Qualified"
Case "Hiring Manager Screen - To Be HM Screened"
Range("B" & CurRow).Value = "Qualified"
Case "Interview/Assessment -  Assessment Scheduled"
Range("B" & CurRow).Value = "Qualified"
Case "Interview/Assessment -  Interview Scheduled"
Range("B" & CurRow).Value = "Qualified"
Case "Interview/Assessment -  Interview/Assessment To be Scheduled"
Range("B" & CurRow).Value = "Qualified"
Case "Interview/Assessment - Assessment Scheduled"
Range("B" & CurRow).Value = "Qualified"
Case "Interview/Assessment - Interview Scheduled"
Range("B" & CurRow).Value = "Qualified"
Case "Interview/Assessment - Interview/Assessment To be Scheduled"
Range("B" & CurRow).Value = "Qualified"
Case "New -  Candidate Considered"
Range("B" & CurRow).Value = "Qualified"
Case "New - Candidate Considered"
Range("B" & CurRow).Value = "Qualified"
Case "New - Proceed to HM Screen"
Range("B" & CurRow).Value = "Qualified"
Case "Hiring Manager Screen -  To Be HM Screened"
Range("B" & CurRow).Value = "Qualified"
Case "New -  Proceed to HM Screen"
Range("B" & CurRow).Value = "Qualified"
Case "Hiring Manager Screen -  Rejected"
Range("B" & CurRow).Value = "Interviewed"
Case "Hiring Manager Screen -  Withdrawn"
Range("B" & CurRow).Value = "Interviewed"
Case "Hiring Manager Screen - Rejected"
Range("B" & CurRow).Value = "Interviewed"
Case "Hiring Manager Screen - Withdrawn"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview/Assessment -  Interview Completed"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview/Assessment -  Proceed to Next Step"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview/Assessment -  Rejected"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview/Assessment -  Withdrawn"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview/Assessment - Assessment Completed"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview/Assessment - Interview Completed"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview/Assessment - Proceed to Next Step"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview/Assessment - Rejected"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview/Assessment - Withdrawn"
Range("B" & CurRow).Value = "Interviewed"
Case "New - HR/Staffing Interview Completed"
Range("B" & CurRow).Value = "Interviewed"
Case "Pre-Offer Activities (non-US/PR) - Proceed to Offer"
Range("B" & CurRow).Value = "Interviewed"
Case "Pre-Offer Activities (non-US/PR) -  Rejected"
Range("B" & CurRow).Value = "Interviewed"
Case "New -  HR/Staffing Interview Completed"
Range("B" & CurRow).Value = "Interviewed"
Case "Pre-Offer Activities (non-US/PR) -  Withdrawn"
Range("B" & CurRow).Value = "Interviewed"
Case "Pre-Offer Activities (non-US/PR) - Rejected"
Range("B" & CurRow).Value = "Interviewed"
Case "Offer -  Canceled"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer -  Has Declined"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer -  Refused"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer -  Rejected"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer -  Reneged"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer -  Rescinded"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Accepted"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Approval in Progress"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Canceled"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Draft"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Extended"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Has Declined"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Offer to be made"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Refused"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Rejected"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Reneged"
Range("B" & CurRow).Value = "Offer Made"
Case "Hire -  Rejected"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - In Negotiation"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities - Withdrawn"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities - Background Check Completed"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities - Background Check Initiated"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities - Medical/Drug Test Completed"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities - Medical/Drug Test Initiated"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities - Rejected"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities - SEND EXTERNAL PRE-HIRE TO SAP"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities - To Be Post-Offer Checked"
Range("B" & CurRow).Value = "Offer Made"
Case "Pre-Offer Activities (non-US/PR) - To Be Pre-Offer Checked"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities -  Withdrawn"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities -  Rejected"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer - Rescinded"
Range("B" & CurRow).Value = "Offer Made"
Case "Post-Offer Activities - Proceed to Hire"
Range("B" & CurRow).Value = "Offer Made"
Case "Hire -  Cleared for Hire/Proceed to OnBoarding"
Range("B" & CurRow).Value = "Hired"
Case "Hire -  To Be Hired"
Range("B" & CurRow).Value = "Hired"
Case "Hire - Cleared for Hire/Proceed to OnBoarding"
Range("B" & CurRow).Value = "Hired"
Case "Hire - To Be Hired"
Range("B" & CurRow).Value = "Hired"
End Select

CurRow = CurRow + 1

Loop

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Deluxe()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:B").Delete
Range("F:G").Delete

Range("A1").Value = "Date"
Range("B1").Value = "JobID1"
Range("C1").Value = "Title"
Range("D1").Value = "Email"
Range("E1").Value = "Original Status"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("G1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

Dim OriginalStatus
OriginalStatus = Range("H" & CurRow).Value

Select Case OriginalStatus
Case "0-Filed"
Range("B" & CurRow).Value = "Apply Completed"
Case "Left Message"
Range("B" & CurRow).Value = "Apply Completed"
Case "No Interest - Deluxe"
Range("B" & CurRow).Value = "Apply Completed"
Case "Resume Review:  Hiring Manager"
Range("B" & CurRow).Value = "Apply Completed"
Case "Resume Review:  Talent Acquisition"
Range("B" & CurRow).Value = "Apply Completed"
Case "Resume Reviewed"
Range("B" & CurRow).Value = "Apply Completed"
Case "Skill/Behavioral Assessment"
Range("B" & CurRow).Value = "Apply Completed"
Case "Hold for Future Consideration"
Range("B" & CurRow).Value = "Qualified"
Case "No Interest - Candidate"
Range("B" & CurRow).Value = "Qualified"
Case "Submitted to Manager"
Range("B" & CurRow).Value = "Qualified"
Case "1st Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "2nd Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "Phone Interview - Manager"
Range("B" & CurRow).Value = "Interviewed"
Case "Phone Interview - Talent Acq"
Range("B" & CurRow).Value = "Interviewed"
Case "Offer"
Range("B" & CurRow).Value = "Offer Made"
Case "Hired"
Range("B" & CurRow).Value = "Hired"
Case "Hired - Contract"
Range("B" & CurRow).Value = "Hired"
End Select

CurRow = CurRow + 1

Loop

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select
Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub HealthOne()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Original Status"
Range("E1").Value = "Date"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("G1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

Dim OriginalStatus
OriginalStatus = Range("H" & CurRow).Value

Select Case OriginalStatus
Case "HR Review"
Range("B" & CurRow).Value = "Apply Completed"
Case "Hiring Manager Review"
Range("B" & CurRow).Value = "Qualified"
Case "Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "Offer"
Range("B" & CurRow).Value = "Offer Made"
Case "Pre-Employment Screen"
Range("B" & CurRow).Value = "Offer Made"
Case "Hire"
Range("B" & CurRow).Value = "Hired"
End Select

CurRow = CurRow + 1

Loop

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub KPMG()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:F").Delete
Range("C:C").Delete
Range("D:O").Delete
Range("E:G").Delete

Range("A1").Value = "Original Status"
Range("B1").Value = "Email"
Range("C1").Value = "JobID1"
Range("D1").Value = "Title"
Range("E1").Value = "Date"

Range("A:A").Cut Destination:=Range("H:H")
Range("B:B").Cut Destination:=Range("A:A")
Range("B1").Select
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("C:C")
Range("F:F").Delete

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

Dim OriginalStatus
OriginalStatus = Range("H" & CurRow).Value

Select Case OriginalStatus
Case "New"
Range("B" & CurRow).Value = "Apply Completed"
Case "Unconsidered"
Range("B" & CurRow).Value = "Apply Completed"
Case "Being Considered"
Range("B" & CurRow).Value = "Qualified"
Case "Cleared Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview In Progress"
Range("B" & CurRow).Value = "Interviewed"
Case "Rejected in Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "Offer Accepted"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Extended"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Rejected"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Rescinded"
Range("B" & CurRow).Value = "Offer Made"
Case "Hired"
Range("B" & CurRow).Value = "Hired"
End Select

CurRow = CurRow + 1

Loop

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Booz()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:D").Delete
Range("C:D").Delete
Range("D:D").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("I:I").Cut Destination:=Range("C:C")
Range("F:F").Cut Destination:=Range("G:G")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:O" & LastRow).Font.Size = 10
Range("A1:O" & LastRow).Font.Name = "Arial"
Range("A1:O1").Font.Color = vbBlack
Range("A1:O1").Font.Bold = True
Range("A1:O1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 150000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hire"
 DestArray(DestRow, 2) = "Hired"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Qualified"
 DestArray(DestRow, 2) = "Qualified"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Manpower()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:D").Delete
Range("C:C").Delete
Range("E:F").Delete

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Date"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Dim LastRow
LastRow = Range("A100000").End(xlUp).Row

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("A1").Value = "1"
Range("A2:A" & LastRow).Formula = "=IF(OR(RIGHT(B2,1)=""i"",RIGHT(B2,1)=""0"",RIGHT(B2,1)=""1"",RIGHT(B2,1)=""2"",RIGHT(B2,1)=""3"",RIGHT(B2,1)=""4"",RIGHT(B2,1)=""5"",RIGHT(B2,1)=""6"",RIGHT(B2,1)=""7"",RIGHT(B2,1)=""8"",RIGHT(B2,1)=""9""),"""",1)"
Range("A2:A" & LastRow).Select
Selection.Copy
Range("A2:A" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

On Error Resume Next     ' In case there are no blanks
Columns("A:A").SpecialCells(xlCellTypeConstants, xlTextValues).EntireRow.Delete
ActiveSheet.UsedRange 'Resets UsedRange for Excel 97

LastRow = Range("A100000").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Range("C2:C" & LastRow).Value = "Apply Completed"
Range("I2:I" & LastRow).Formula = Range("D2:D" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("D2:D" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("D2:D" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("D2:D" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A:A").Delete

Range("A1:G" & LastRow).Borders.Weight = xlThin
Range("A1:G" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Avanade()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("D:D").Delete

Range("A1").Value = "Email"
Range("B1").Value = "JobID1"
Range("C1").Value = "Original Status"
Range("D1").Value = "Date"
Range("E1").Value = "Title"

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("C:C")
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("E1").Value = "Title"
Range("F:F").Cut Destination:=Range("H:H")

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

Dim OriginalStatus
OriginalStatus = Range("H" & CurRow).Value

Select Case OriginalStatus
Case "New"
Range("B" & CurRow).Value = "Apply Completed"
Case "On Campus"
Range("B" & CurRow).Value = "Apply Completed"
Case "Sourcing"
Range("B" & CurRow).Value = "Apply Completed"
Case "HM Review"
Range("B" & CurRow).Value = "Qualified: Business Screen"
Case "HM Review 2"
Range("B" & CurRow).Value = "Qualified: Business Screen"
Case "CV Review"
Range("B" & CurRow).Value = "Qualified: Recruiter Screen"
Case "Recruiting Screen"
Range("B" & CurRow).Value = "Qualified: Recruiter Screen"
Case "Resume/CV Review"
Range("B" & CurRow).Value = "Qualified: Recruiter Screen"
Case "S/AM Review"
Range("B" & CurRow).Value = "Qualified: Recruiter Screen"
Case "SRT Review"
Range("B" & CurRow).Value = "Qualified: Recruiter Screen"
Case "Skills Assessment"
Range("B" & CurRow).Value = "Qualified: Technical Screen"
Case "Skills Assessment 2"
Range("B" & CurRow).Value = "Qualified: Technical Screen"
Case "Skills Assessment APAC"
Range("B" & CurRow).Value = "Qualified: Technical Screen"
Case "Skills Assessment 3"
Range("B" & CurRow).Value = "Qualified: Technical Screen"
Case "Final Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "Offer"
Range("B" & CurRow).Value = "Interviewed"
Case "HR Interview"
Range("B" & CurRow).Value = "Interviewed"
Case "PreEmployment Checks"
Range("B" & CurRow).Value = "Offer Made"
Case "Background Investigation"
Range("B" & CurRow).Value = "Offer Made"
Case "Hire"
Range("B" & CurRow).Value = "Hired"
End Select

If Range("D" & CurRow).Value = 8504 Or Range("D" & CurRow).Value = 8498 Or Range("D" & CurRow).Value = 8940 Or Range("D" & CurRow).Value = 8505 Or Range("D" & CurRow).Value = 8463 Or Range("D" & CurRow).Value = 8468 Or Range("D" & CurRow).Value = 8501 Or Range("D" & CurRow).Value = 8491 Or Range("H" & CurRow).Value = "Lead Review" Or Range("H" & CurRow).Value = "Referral Review" Then
Range(CurRow & ":" & CurRow).Delete
CurRow = CurRow - 1
LastRow = Range("A65536").End(xlUp).Row

Else
End If

CurRow = CurRow + 1

Loop

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Celgene()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:1").Delete
Range("F:J").Delete

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Original Status"
Range("E1").Value = "Date"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case " Applied"
DestArray(DestRow, 2) = "Apply Completed"
Case " Reject"
DestArray(DestRow, 2) = "Apply Completed"
Case " Under Consideratn/Meets BQ"
DestArray(DestRow, 2) = "Qualified"
Case " Interview"
DestArray(DestRow, 2) = "Interviewed"
Case " Interviewing"
DestArray(DestRow, 2) = "Interviewed"
Case " Offer"
DestArray(DestRow, 2) = "Offer Made"
Case " Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case " PreEmployment Checks"
DestArray(DestRow, 2) = "Offer Made"
Case " Ready to Hire"
DestArray(DestRow, 2) = "Offer Made"
Case " Linked"
DestArray(DestRow, 2) = "Offer Made"
Case " Hired"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub ClevelandClinic()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("A400000").End(xlUp).Row

Range("AK2:AK" & LastRow).Formula = "=IF(R2=""File Hired"",""Hired"",IF(ISBLANK(AF2),IF(ISBLANK(AD2),IF(R2=""Filed Not Hired Viable"",""Qualified"",""Apply Completed""),""Interviewed""),""Offer Made""))"
Range("AL2:AL" & LastRow).Formula = "=IF(AK2=""Hired"",AG2,IF(AK2=""Offer Made"",AF2,IF(AK2=""Interviewed"",AD2,Q2)))"
Range("AK2:AL" & LastRow).Copy
Range("AK2:AL" & LastRow).PasteSpecial xlPasteValues
Range("AL2:AL" & LastRow).NumberFormat = "mm/dd/yyyy"

Range("A:B").Delete
Range("B:I").Delete
Range("C:E").Delete
Range("D:W").Delete

Range("A1").Value = "Title"
Range("B1").Value = "JobID1"
Range("C1").Value = "Email"
Range("D1").Value = "Status"
Range("E1").Value = "Date"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("D:D")

Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("A1:G" & LastRow).Font.Size = 10
Range("A1:G" & LastRow).Font.Name = "Arial"
Range("A1:G1").Font.Color = vbBlack
Range("A1:G1").Font.Bold = True
Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & LastRow).Borders.Weight = xlThin
Range("A1:G" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub LifetimeFitness()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:3").Delete
Range("A:A").Delete
Range("A:C").Delete
Range("C:G").Delete
Range("D:G").Delete
Range("F:P").Delete

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Original Status"
Range("E1").Value = "Date"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "New - Advance"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Has Declined"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Rejected"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - To be reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "1st Interview - To Be Scheduled"
DestArray(DestRow, 2) = "Interviewed"
Case "Category Director/LTU - Approved"
DestArray(DestRow, 2) = "Interviewed"
Case "Category Director/LTU - Has Declined"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview Process - 1st Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview Process - 2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview Process - Awaiting Candidate Response"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview Process - Final Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview Process - Has Declined"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview Process - Rejected"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview Process - To be Scheduled"
DestArray(DestRow, 2) = "Interviewed"
Case "Prescreen - Awaiting Candidate Response"
DestArray(DestRow, 2) = "Interviewed"
Case "Prescreen - Has Declined"
DestArray(DestRow, 2) = "Interviewed"
Case "Prescreen - Left a Message / Waiting for information"
DestArray(DestRow, 2) = "Interviewed"
Case "Prescreen - Move to Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Prescreen - Phone Screen"
DestArray(DestRow, 2) = "Interviewed"
Case "Prescreen - Rejected"
DestArray(DestRow, 2) = "Interviewed"
Case "Prescreen - Resume Review"
DestArray(DestRow, 2) = "Interviewed"
Case "Prescreen - To Be Prescreened"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview - Rejected"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview - Has Declined"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview - Advance"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview - Rejected"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview - To be scheduled/send PP information to candidate"
DestArray(DestRow, 2) = "Interviewed"
Case "3rd Interview - To Be Scheduled"
DestArray(DestRow, 2) = "Interviewed"
Case "3rd Interview - Advance (indicate start date)"
DestArray(DestRow, 2) = "Interviewed"
Case "Phone screen - To be conducted"
DestArray(DestRow, 2) = "Interviewed"
Case "Phone screen - Rejected"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview - Has Declined"
DestArray(DestRow, 2) = "Interviewed"
Case "Background Check - Advance (indicate start date)"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Check - Background Consent Requested - Awaiting Candidate Response"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Check - Candidate Consent Verified by Manager - Request Screening Service"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Check - Has Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Check - Rejected"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Check - Waiting for Results"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Has Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Offer Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Offer to be made"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Rejected"
DestArray(DestRow, 2) = "Offer Made"
Case "Hire - Has Declined"
DestArray(DestRow, 2) = "Hired"
Case "Hire - Hired"
DestArray(DestRow, 2) = "Hired"
Case "Hire - Rejected"
DestArray(DestRow, 2) = "Hired"
Case "Hire - Waiting for Info"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Pepsi()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A1").Value = "Date"
Range("B1").Value = "JobID1"
Range("C1").Value = "Title"
Range("D1").Value = "Email"
Range("E1").Value = "Status"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("B:B")

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Dim LastRow
LastRow = Range("A100000").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Range("I2:I" & LastRow).Formula = "=IF(NOT(ISERR(FIND(""Internal"",H2,1))),1,0)"
Range("J2:J" & LastRow).Formula = "=IF(OR(B2=""Hired - Internal"",B2=""MovetoOnboarding"",B2=""Move to OnBoarding"",B2=""Offer Accepted""),1,0)"
Range("K2:K" & LastRow).Formula = "=I2+J2"

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

If Range("K" & CurRow).Value > 0 Then
Range(CurRow & ":" & CurRow).Delete
LastRow = Range("A100000").End(xlUp).Row
Else
CurRow = CurRow + 1
LastRow = Range("A100000").End(xlUp).Row
End If

Loop

Range("I:K").Delete

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Rackspace()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:E").Delete
Range("B:C").Delete
Range("C:E").Delete
Range("D:E").Delete
Range("F:H").Delete

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Date"
Range("E1").Value = "Original Status"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A600000").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 600000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Candidate Transferred"
DestArray(DestRow, 2) = "Apply Completed"
Case "Declined"
DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Contacted"
DestArray(DestRow, 2) = "Apply Completed"
Case "Pending"
DestArray(DestRow, 2) = "Apply Completed"
Case "Prospect"
DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted"
DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate"
DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Phone Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Removed Self from Consideration"
DestArray(DestRow, 2) = "Qualified"
Case "Submitted to Hiring Manager"
DestArray(DestRow, 2) = "Qualified"
Case "Skills Testing"
DestArray(DestRow, 2) = "Qualified"
Case "In-Person Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview "
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "3rd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Manager Phone Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Reference Check"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Yoh()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:A").Delete
Range("D:D").Delete
Range("F:H").Delete

Range("A1").Value = "Email"
Range("B1").Value = "JobID1"
Range("C1").Value = "Title"
Range("D1").Value = "Original Status"
Range("E1").Value = "Date"

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Range("B2:B" & LastRow).Formula = "=IF(ISERROR(YEAR(C2)),DATE(RIGHT(C2,4),MID(C2,4,2),LEFT(C2,2)),DATE(YEAR(C2),DAY(C2),MONTH(C2)))"
Range("B2:B" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("B2:B" & LastRow).ClearContents
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "AM Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Submitted"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Unsuccessful"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Web Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Presented"
 DestArray(DestRow, 2) = "Qualified"
Case "Rejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Identified"
 DestArray(DestRow, 2) = "Qualified"
Case "Potential"
 DestArray(DestRow, 2) = "Qualified"
Case "Client Interviewing"
 DestArray(DestRow, 2) = "Interviewed"
Case "Client Rejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferMade"
 DestArray(DestRow, 2) = "Offer Made"
Case "Made Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Placed"
 DestArray(DestRow, 2) = "Hired"
Case "Re-assigned"
 DestArray(DestRow, 2) = "Hired"
Case "Replaced"
 DestArray(DestRow, 2) = "Hired"
Case "To Be Replaced"
 DestArray(DestRow, 2) = "Hired"
Case "Assigned"
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub WelchAllyn()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:1").Delete
Range("D:E").Delete

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Original Status"
Range("E1").Value = "Date"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "ATS Capture"
 DestArray(DestRow, 2) = "ATS Capture"
Case "Apply Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Inbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Suitable"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Screening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Not Selected For Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Qualified"
 DestArray(DestRow, 2) = "Qualified"
Case "Schedule Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Not Selected After Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Approve Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre Employment Process"
 DestArray(DestRow, 2) = "Offer Made"
Case "Create Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer approved"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Made"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Hired"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Nalco()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:3").Delete
Range("A:A").Delete
Range("B:E").Delete
Range("C:F").Delete
Range("F:I").Delete

Range("A1").Value = "Email"
Range("B1").Value = "JobID1"
Range("C1").Value = "Title"
Range("D1").Value = "Original Status"
Range("E1").Value = "Date"

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("C65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "New"
DestArray(DestRow, 2) = "Apply Completed"
Case "Screen & Intv"
DestArray(DestRow, 2) = "Interviewed"
Case "Checks & Tests"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer"
DestArray(DestRow, 2) = "Offer Made"
Case "Hire"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Sodexo()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:V1").Value = "A"
Range("W1").Value = "Apply Completed"
Range("X1").Value = "Apply Completed"
Range("Y1").Value = "Apply Completed"
Range("Z1").Value = "Apply Completed"
Range("AA1").Value = "Apply Completed"
Range("AB1").Value = "Qualified"
Range("AC1").Value = "Qualified"
Range("AD1").Value = "Apply Completed"
Range("AE1").Value = "Qualified"
Range("AF1").Value = "Qualified"
Range("AG1").Value = "Interviewed"
Range("AH1").Value = "Interviewed"
Range("AI1").Value = "Interviewed"
Range("AJ1").Value = "Interviewed"
Range("AK1").Value = "Offer Made"
Range("AL1").Value = "Offer Made"
Range("AM1").Value = "Offer Made"
Range("AN1").Value = "Hired"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 23

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets("Sheet1").Range("A1:AP" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 42)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)
    DestArray(1, 13) = SourceArray(2, 13)
    DestArray(1, 14) = SourceArray(2, 14)
    DestArray(1, 15) = SourceArray(2, 15)
    DestArray(1, 16) = SourceArray(2, 16)
    DestArray(1, 17) = SourceArray(2, 17)
    DestArray(1, 18) = SourceArray(2, 18)
    DestArray(1, 19) = SourceArray(2, 19)
    DestArray(1, 20) = SourceArray(2, 20)
    DestArray(1, 21) = SourceArray(2, 21)
    DestArray(1, 22) = SourceArray(2, 22)
    DestArray(1, 23) = SourceArray(2, 23)
    DestArray(1, 24) = SourceArray(2, 24)
    DestArray(1, 25) = SourceArray(2, 25)
    DestArray(1, 26) = SourceArray(2, 26)
    DestArray(1, 27) = SourceArray(2, 27)
    DestArray(1, 28) = SourceArray(2, 28)
    DestArray(1, 29) = SourceArray(2, 29)
    DestArray(1, 30) = SourceArray(2, 30)
    DestArray(1, 31) = SourceArray(2, 31)
    DestArray(1, 32) = SourceArray(2, 32)
    DestArray(1, 33) = SourceArray(2, 33)
    DestArray(1, 34) = SourceArray(2, 34)
    DestArray(1, 35) = SourceArray(2, 35)
    DestArray(1, 36) = SourceArray(2, 36)
    DestArray(1, 37) = SourceArray(2, 37)
    DestArray(1, 38) = SourceArray(2, 38)
    DestArray(1, 39) = SourceArray(2, 39)
    DestArray(1, 40) = SourceArray(2, 40)
    

For CurRow = 3 To LastRow
      
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, 20)
                DestArray(DestRow, 21) = SourceArray(CurRow, 21)
                DestArray(DestRow, 22) = SourceArray(CurRow, 22)
                DestArray(DestRow, 23) = SourceArray(CurRow, 23)
                DestArray(DestRow, 24) = SourceArray(CurRow, 24)
                DestArray(DestRow, 25) = SourceArray(CurRow, 25)
                DestArray(DestRow, 26) = SourceArray(CurRow, 26)
                DestArray(DestRow, 27) = SourceArray(CurRow, 27)
                DestArray(DestRow, 28) = SourceArray(CurRow, 28)
                DestArray(DestRow, 29) = SourceArray(CurRow, 29)
                DestArray(DestRow, 30) = SourceArray(CurRow, 30)
                DestArray(DestRow, 31) = SourceArray(CurRow, 31)
                DestArray(DestRow, 32) = SourceArray(CurRow, 32)
                DestArray(DestRow, 33) = SourceArray(CurRow, 33)
                DestArray(DestRow, 34) = SourceArray(CurRow, 34)
                DestArray(DestRow, 35) = SourceArray(CurRow, 35)
                DestArray(DestRow, 36) = SourceArray(CurRow, 36)
                DestArray(DestRow, 37) = SourceArray(CurRow, 37)
                DestArray(DestRow, 38) = SourceArray(CurRow, 38)
                DestArray(DestRow, 39) = SourceArray(CurRow, 37)
                DestArray(DestRow, 40) = SourceArray(CurRow, 38)
                DestArray(DestRow, 41) = SourceArray(CurRow, 23)
                DestArray(DestRow, 42) = SourceArray(1, 23)
                
        For CurCol = 23 To 40
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(LastRow + i, 1) = SourceArray(CurRow, 1)
                DestArray(LastRow + i, 2) = SourceArray(CurRow, 2)
                DestArray(LastRow + i, 3) = SourceArray(CurRow, 3)
                DestArray(LastRow + i, 4) = SourceArray(CurRow, 4)
                DestArray(LastRow + i, 5) = SourceArray(CurRow, 5)
                DestArray(LastRow + i, 6) = SourceArray(CurRow, 6)
                DestArray(LastRow + i, 7) = SourceArray(CurRow, 7)
                DestArray(LastRow + i, 8) = SourceArray(CurRow, 8)
                DestArray(LastRow + i, 9) = SourceArray(CurRow, 9)
                DestArray(LastRow + i, 10) = SourceArray(CurRow, 10)
                DestArray(LastRow + i, 11) = SourceArray(CurRow, 11)
                DestArray(LastRow + i, 12) = SourceArray(CurRow, 12)
                DestArray(LastRow + i, 13) = SourceArray(CurRow, 13)
                DestArray(LastRow + i, 14) = SourceArray(CurRow, 14)
                DestArray(LastRow + i, 15) = SourceArray(CurRow, 15)
                DestArray(LastRow + i, 16) = SourceArray(CurRow, 16)
                DestArray(LastRow + i, 17) = SourceArray(CurRow, 17)
                DestArray(LastRow + i, 18) = SourceArray(CurRow, 18)
                DestArray(LastRow + i, 19) = SourceArray(CurRow, 19)
                DestArray(LastRow + i, 20) = SourceArray(CurRow, 20)
                DestArray(LastRow + i, 21) = SourceArray(CurRow, 21)
                DestArray(LastRow + i, 22) = SourceArray(CurRow, 22)
                DestArray(LastRow + i, 23) = SourceArray(CurRow, 23)
                DestArray(LastRow + i, 24) = SourceArray(CurRow, 24)
                DestArray(LastRow + i, 25) = SourceArray(CurRow, 25)
                DestArray(LastRow + i, 26) = SourceArray(CurRow, 26)
                DestArray(LastRow + i, 27) = SourceArray(CurRow, 27)
                DestArray(LastRow + i, 28) = SourceArray(CurRow, 28)
                DestArray(LastRow + i, 29) = SourceArray(CurRow, 29)
                DestArray(LastRow + i, 30) = SourceArray(CurRow, 30)
                DestArray(LastRow + i, 31) = SourceArray(CurRow, 31)
                DestArray(LastRow + i, 32) = SourceArray(CurRow, 32)
                DestArray(LastRow + i, 33) = SourceArray(CurRow, 33)
                DestArray(LastRow + i, 34) = SourceArray(CurRow, 34)
                DestArray(LastRow + i, 35) = SourceArray(CurRow, 35)
                DestArray(LastRow + i, 36) = SourceArray(CurRow, 36)
                DestArray(LastRow + i, 37) = SourceArray(CurRow, 37)
                DestArray(LastRow + i, 38) = SourceArray(CurRow, 38)
                DestArray(LastRow + i, 39) = SourceArray(CurRow, 39)
                DestArray(LastRow + i, 40) = SourceArray(CurRow, 40)
                DestArray(LastRow + i, 41) = SourceArray(CurRow, CurCol)
                DestArray(LastRow + i, 42) = SourceArray(1, CurCol)
                
                i = i + 1
                
            Else
            End If
        Next CurCol
        
        DestRow = DestRow + 1
        
Next CurRow

Sheets("Sheet1").Range("1:1").Delete

Sheets("Sheet1").Range("A1:AP" & LastRow + i - 1).Value = DestArray

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("AQ:AQ").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("AQ:AQ").Cut Destination:=Range("C:C")
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("E:E")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("I:I").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("AS2:AS" & LastRow + i).Formula = Range("C2:C" & LastRow + i).Value2
Range("AS2:AS" & LastRow + i).Select
Selection.Copy
Range("C2:C" & LastRow + i).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("AS2:AS" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("AS2:AS" & LastRow + i).Select
Selection.Copy
Range("C2:C" & LastRow + i).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("AS2:AS" & LastRow + i).Delete
Range("C2:C" & LastRow + i).NumberFormat = "mm-dd-yyyy"

Sheets("Sheet1").Range("A1:AR" & LastRow + i).Font.Size = 10
Sheets("Sheet1").Range("A1:AR" & LastRow + i).Font.Name = "Arial"
Sheets("Sheet1").Range("A1:AR1").Font.Color = vbBlack
Sheets("Sheet1").Range("A1:AR1").Font.Bold = True
Sheets("Sheet1").Range("A1:AR1").Interior.Color = vbYellow

Range("A1:AR" & LastRow + i).Borders.Weight = xlThin
Range("A1:AR" & LastRow + i).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub ChristianaCare()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:B").Delete
Range("C:D").Delete
Range("E:P").Delete
Range("F:H").Delete

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Original Status"
Range("D1").Value = "Date"
Range("E1").Value = "Email"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "0-Filed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "0-Status"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Did not meet min qual"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Did not meet min w/o notice"
 DestArray(DestRow, 2) = "Apply Completed"
Case "KPI Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "KPI Failed - External Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "KPI Failed - Internal Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "KPI Incomplete"
 DestArray(DestRow, 2) = "Apply Completed"
Case "KPI Pending"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Considered"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Assesment Incomplete"
 DestArray(DestRow, 2) = "Qualified"
Case "Behavioral Incomplete"
 DestArray(DestRow, 2) = "Qualified"
Case "Screened"
 DestArray(DestRow, 2) = "Qualified"
Case "Reviewed"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest Candidate"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest CCHS"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest CCHS w/o notice"
 DestArray(DestRow, 2) = "Qualified"
Case "No Skills Required"
 DestArray(DestRow, 2) = "Qualified"
Case "Assesment Failed-External"
 DestArray(DestRow, 2) = "Qualified"
Case "Behavioral Failed-External Candidate"
 DestArray(DestRow, 2) = "Qualified"
Case "Launch Assessment"
 DestArray(DestRow, 2) = "Qualified"
Case "Launch Behavioral "
 DestArray(DestRow, 2) = "Qualified"
Case "Skills 1 Pending"
 DestArray(DestRow, 2) = "Qualified"
Case "Skills 2 Pending"
 DestArray(DestRow, 2) = "Qualified"
Case "Skills 3 Pending"
 DestArray(DestRow, 2) = "Qualified"
Case "Skills 4 Pending"
 DestArray(DestRow, 2) = "Qualified"
Case "Skills Completed Review Results"
 DestArray(DestRow, 2) = "Qualified"
Case "Skills Incomplete"
 DestArray(DestRow, 2) = "Qualified"
Case "Reviewed*"
 DestArray(DestRow, 2) = "Qualified"
Case "Behavioral Pending"
 DestArray(DestRow, 2) = "Qualified"
Case "Behavioral Completed-No Skills Required"
 DestArray(DestRow, 2) = "Qualified"
Case "No Behavioral No Skills"
 DestArray(DestRow, 2) = "Qualified"
Case "No Behavioral Required"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone Screened"
 DestArray(DestRow, 2) = "Qualified"
Case "Reviewed (old)"
 DestArray(DestRow, 2) = "Qualified"
Case "Reviewed Resume"
 DestArray(DestRow, 2) = "Qualified"
Case "Final Candidate Status"
 DestArray(DestRow, 2) = "Interviewed"
Case "HR Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Mgr. Notification by HR"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Made"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Promoted "
 DestArray(DestRow, 2) = "Hired"
Case "Reviewed Refs & Approved"
 DestArray(DestRow, 2) = "Hired"
Case "Reviewed Refs not Accept."
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Hasbro()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim DrObj
Dim Pict
Set DrObj = ActiveSheet.DrawingObjects
For Each Pict In DrObj
If Left(Pict.Name, 7) = "Picture" Then
Pict.Select
Pict.Delete
End If
Next

Range("1:2").Delete
Range("F:O").Delete

Sheets(1).Cells.UnMerge

Range("A1").Value = "Email"
Range("B1").Value = "JobID1"
Range("C1").Value = "Original Status"
Range("D1").Value = "Date"
Range("E1").Value = "Title"

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("C:C")
Range("E:E").Cut Destination:=Range("H:H")
Range("G:G").Cut Destination:=Range("E:E")

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NEW"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Minimally Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Send to Manager"
 DestArray(DestRow, 2) = "Qualified"
Case "HR Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager Not Interested"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager Not Interested (RMO)"
 DestArray(DestRow, 2) = "Qualified"
Case "Resume Review Complete"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre Employment"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Manager Interviews"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Manager Interviews Complete "
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviewing"
 DestArray(DestRow, 2) = "Interviewed"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Interviewed"
Case "Recruiter Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Recruiter Interview (Phone Screen)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Recruiter Interview Complete"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"

End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("E" & LastRow + 2).Delete

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub GoodLife()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:1").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("A:A")
Range("H:H").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("L:L").Cut Destination:=Range("C:C")
Range("L:L").Delete
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("I:I").Cut Destination:=Range("D:D")
Range("I:I").Delete
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("G1").Select
ActiveCell.EntireColumn.Insert
Range("K:K").Cut Destination:=Range("H:H")
Range("K:K").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:P" & LastRow).Font.Size = 10
Range("A1:P" & LastRow).Font.Name = "Arial"
Range("A1:P1").Font.Color = vbBlack
Range("A1:P1").Font.Bold = True
Range("A1:P1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:P" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 16)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    DestArray(1, 9) = SourceArray(1, 9)
    DestArray(1, 10) = SourceArray(1, 10)
    DestArray(1, 11) = SourceArray(1, 11)
    DestArray(1, 12) = SourceArray(1, 12)
    DestArray(1, 13) = SourceArray(1, 13)
    DestArray(1, 14) = SourceArray(1, 14)
    DestArray(1, 15) = SourceArray(1, 15)
    DestArray(1, 16) = SourceArray(1, 16)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 14)

Select Case OriginalStatus
Case "Inbox"
DestArray(DestRow, 2) = "Apply Completed"
Case "Applicants Without Screening Questions Completed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Did Not Pass Resume Review"
DestArray(DestRow, 2) = "Apply Completed"
Case "Not Suitable"
DestArray(DestRow, 2) = "Apply Completed"
Case "Phone Screen Completed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Left Message"
DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected For Interview"
DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected - Future Interest"
DestArray(DestRow, 2) = "Apply Completed"
Case "Internal Applicants No Transfer Form"
DestArray(DestRow, 2) = "Apply Completed"
Case "Transferred to Other Requisition"
DestArray(DestRow, 2) = "Apply Completed"
Case "Hiring Manager Review"
DestArray(DestRow, 2) = "Qualified"
Case "Set Up Phone Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Review"
DestArray(DestRow, 2) = "Qualified"
Case "Withdrew"
DestArray(DestRow, 2) = "Qualified"
Case "Not Selected for 2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed Not Selected"
DestArray(DestRow, 2) = "Interviewed"
Case "Not Selected after 2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview Booked"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview Booked"
DestArray(DestRow, 2) = "Interviewed"
Case "Schedule 1st Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview Completed"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview Completed"
DestArray(DestRow, 2) = "Interviewed"
Case "Schedule 2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "HM Schedule 2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
DestArray(DestRow, 2) = "Hired"
Case "Started"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:P" & DestRow).Value = DestArray

Range("Q2:Q" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("Q2:Q" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("Q2:Q" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("Q2:Q" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("Q2:Q" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:P" & LastRow).Borders.Weight = xlThin
Range("A1:P" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("Q2:Q" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("Q1").Formula = "=SUM(Q2:Q" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("Q1").Value

Range("Q:Q").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:P" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub


Sub NorthHighland()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("C:D").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("Y2:Y" & LastRow).Formula = "=IF(D2="""",C2,D2)"

Range("Y2:Y" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("Y:Y").Delete

Range("Y2:Y" & LastRow).Formula = "=IF(X2<>"""",IF(E2<>"""",E2,X2),X2)"
Range("Y2:Y" & LastRow).Select
Selection.Copy
Range("X2:X" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("Y:Y").Delete

Range("D:F").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("A2:A" & LastRow).Formula = "=RIGHT(B2,FIND(""-"",B2,1)-1)"
Range("B:B").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:C1").Value = "A"
Range("D1").Value = "Apply Completed"
Range("E1").Value = "Apply Completed"
Range("F1").Value = "Apply Completed"
Range("G1").Value = "Apply Completed"
Range("H1").Value = "Apply Completed"
Range("I1").Value = "Apply Completed"
Range("J1").Value = "Apply Completed"
Range("K1").Value = "Qualified"
Range("L1").Value = "Qualified"
Range("M1").Value = "Qualified"
Range("N1").Value = "Interviewed"
Range("O1").Value = "Interviewed"
Range("P1").Value = "Qualified"
Range("Q1").Value = "Interviewed"
Range("R1").Value = "Offer Made"
Range("S1").Value = "Offer Made"
Range("T1").Value = "Offer Made"
Range("U1").Value = "Hired"

LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 4

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = ActiveSheet.Range("A1:V" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 24)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)
    DestArray(1, 13) = SourceArray(2, 13)
    DestArray(1, 14) = SourceArray(2, 14)
    DestArray(1, 15) = SourceArray(2, 15)
    DestArray(1, 16) = SourceArray(2, 16)
    DestArray(1, 17) = SourceArray(2, 17)
    DestArray(1, 18) = SourceArray(2, 18)
    DestArray(1, 19) = SourceArray(2, 19)
    DestArray(1, 20) = SourceArray(2, 20)
    DestArray(1, 21) = SourceArray(2, 21)
    DestArray(1, 22) = SourceArray(2, 22)

For CurRow = 3 To LastRow
                   
        For CurCol = 4 To 22
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, 20)
                DestArray(DestRow, 21) = SourceArray(CurRow, 21)
                DestArray(DestRow, 22) = SourceArray(CurRow, 22)
                DestArray(DestRow, 23) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 24) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            End If
        Next CurCol
               
Next CurRow

ActiveSheet.Range("1:1").Delete

ActiveSheet.Range("A1:X" & DestRow).Value = DestArray

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:Y").Delete
Range("G:G").Cut Destination:=Range("B:B")
Range("F:F").Cut Destination:=Range("C:C")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

LastRow = Range("A65536").End(xlUp).Row

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Range("H:H").Delete

ActiveSheet.Range("A1:G" & DestRow).Font.Size = 10
ActiveSheet.Range("A1:G" & DestRow).Font.Name = "Arial"
ActiveSheet.Range("A1:G1").Font.Color = vbBlack
ActiveSheet.Range("A1:G1").Font.Bold = True
ActiveSheet.Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub Expedia()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("E1").Value = "Email"
Range("U1").Value = "Status"
Range("V1").Value = "Date"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("A:A")
Range("F:F").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("V:V").Cut Destination:=Range("B:B")
Range("V:V").Delete
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("W:W").Cut Destination:=Range("C:C")
Range("W:W").Delete
Range("F:F").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Select
ActiveCell.EntireColumn.Insert
Range("B:B").Cut Destination:=Range("H:H")

Range("B1").Value = "Status"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A400000").End(xlUp).Row

Range("A1:AE" & LastRow).Font.Size = 10
Range("A1:AE" & LastRow).Font.Name = "Arial"
Range("A1:AE1").Font.Color = vbBlack
Range("A1:AE1").Font.Bold = True
Range("A1:AE1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 400000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Not Suitable"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Inbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "ERP Not Suitable"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected For Interview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hiring Manager Review-Not Suitable"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected after Interview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Request Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Did Not Pass Screening Questions"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Assessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected for Tier 3 Interview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not selected after on-campus interview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Did Not Pass Background Check"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Screening"
 DestArray(DestRow, 2) = "Qualified"
Case "Schedule Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Schedule HM Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Request Tactical Testing"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview Survey Sent"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Tier 2 Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Tier 3 Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview Complete"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview Rescheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "Selected for final interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Psft ID Request"
 DestArray(DestRow, 2) = "Offer Made"
Case "Initiate Background/Reference Check"
 DestArray(DestRow, 2) = "Offer Made"
Case "Create Offer Letter/Psft ID Req"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Approve Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Letter Sent"
 DestArray(DestRow, 2) = "Offer Made"
Case "Start Date Established/Confirmed Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "New Hire Welcome"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("AF2:AF" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("AF2:AF" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("AF2:AF" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("AF2:AF" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("AF2:AF" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:AE" & LastRow).Borders.Weight = xlThin
Range("A1:AE" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub


'TESTESTESTESTS

Sub KPMGUK()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:B").Delete
Range("D:E").Delete
Range("F:F").Delete

Range("A1").Value = "Email"
Range("B1").Value = "JobID1"
Range("C1").Value = "Title"
Range("D1").Value = "Date"
Range("E1").Value = "Original Status"

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Do While CurRow <= LastRow
Application.StatusBar = "Processing record " & CurRow & " of " & LastRow

Dim OriginalStatus
OriginalStatus = Range("H" & CurRow).Value

Select Case OriginalStatus
Case "Assessment - (UK)"
Range("B" & CurRow).Value = "Interviewed"
Case "Assessment Complete – Numerical and Verbal - (UK)"
Range("B" & CurRow).Value = "Interviewed"
Case "Candidate Declines Offer - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Candidate Offer sent - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Candidate withdrawn - (UK)"
Range("B" & CurRow).Value = "Apply Completed"
Case "Contract Approved - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Create Contract - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Export to SAP - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Hired - (UK)"
Range("B" & CurRow).Value = "Hired"
Case "Hiring Manager Review Complete - (UK)"
Range("B" & CurRow).Value = "Qualified"
Case "Hiring Manager Review Pending - (UK)"
Range("B" & CurRow).Value = "Qualified"
Case "Internal Hire - (UK)"
Range("B" & CurRow).Value = "Hired"
Case "Internal Offer Accepted - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Internal Offer Declined - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Interview Arranging - (UK)"
Range("B" & CurRow).Value = "Qualified"
Case "Interview Completed - (UK)"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview Feedback Received - (UK)"
Range("B" & CurRow).Value = "Interviewed"
Case "Interview Scheduled - (UK)"
Range("B" & CurRow).Value = "Interviewed"
Case "Invite to Interview - (UK)"
Range("B" & CurRow).Value = "Qualified"
Case "New Applicant - (UK)"
Range("B" & CurRow).Value = "Apply Completed"
Case "Offer accepted - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Offer Approval Pending"
Range("B" & CurRow).Value = "Interviewed"
Case "Offer Approved - (UK)"
Range("B" & CurRow).Value = "Interviewed"
Case "Offer Declined on Portal - (UK) "
Range("B" & CurRow).Value = "Offer Made"
Case "Onboarder Allocated - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Onboarding Checking - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Onboarding Complete - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Onboarding in Progress - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Online Withdrawal"
Range("B" & CurRow).Value = "Apply Completed"
Case "Pending Assessment - Numerical and Verbal - (UK)"
Range("B" & CurRow).Value = "Interviewed"
Case "Phone Screen - (UK)"
Range("B" & CurRow).Value = "Apply Completed"
Case "Pre-Offer discussions pending - (UK)"
Range("B" & CurRow).Value = "Interviewed"
Case "Prepare SAP export - (UK)"
Range("B" & CurRow).Value = "Offer Made"
Case "Rejected - site questions - (UK)"
Range("B" & CurRow).Value = "Apply Completed"
Case "Rejected Candidate - (DE)"
Range("B" & CurRow).Value = "Apply Completed"
Case "Rejected Candidate - (UK)"
Range("B" & CurRow).Value = "Apply Completed"
Case "Talent Pool in progress - (UK)"
Range("B" & CurRow).Value = "Qualified"
End Select

CurRow = CurRow + 1

Loop

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub FutureStep()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("W2:W" & LastRow).Formula = "=IF(isblank(S2),A2,S2)"
Range("W2:W" & LastRow).Select
Selection.Copy
Range("A2:A" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("W:W").Delete

Range("W:AR").Value = Range("A:V").Value
Range("A:V").Delete

Range("B:F").Delete
Range("C:L").Delete
Range("D:E").Delete

Range("A1").Value = "JobID1"
Range("B1").Value = "Title"
Range("C1").Value = "Email"
Range("D1").Value = "Original Status"
Range("E1").Value = "Date"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Value = "Status"
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "ATS Applicant Capture"
DestArray(DestRow, 2) = "ATS Capture"
Case "ATS Applies"
DestArray(DestRow, 2) = "Apply Completed"
Case "ATS Qualified: Recruiter Screen"
DestArray(DestRow, 2) = "Qualified: Recruiter Screen"
Case "ATS Qualified: Business Screen"
DestArray(DestRow, 2) = "Qualified: Business Screen"
Case "ATS Interviewed"
DestArray(DestRow, 2) = "Interviewed"
Case "ATS Offer Made"
DestArray(DestRow, 2) = "Offer Made"
Case "ATS Hired"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub SRC()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("F:Q").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Original Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"

Range("B:B").Cut Destination:=Range("H:H")

Range("B1").Value = "Status"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Applied"
DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - No Interest"
DestArray(DestRow, 2) = "Apply Completed"
Case "1 - Schedule Interview"
DestArray(DestRow, 2) = "Qualified"
Case "2 - Potential Candidate"
DestArray(DestRow, 2) = "Qualified"
Case "3 - Not Selected For Interview"
DestArray(DestRow, 2) = "Qualified"
Case "Declined"
DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager Review"
DestArray(DestRow, 2) = "Qualified"
Case "Manager Phone Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Phone Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Phone Screen2"
DestArray(DestRow, 2) = "Qualified"
Case "Declined after Phone Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Approve Offer"
DestArray(DestRow, 2) = "Interviewed"
Case "Create Offer"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Not Selected after Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Create Temp Offer"
DestArray(DestRow, 2) = "Interviewed"
Case "Approve Temp Offer"
DestArray(DestRow, 2) = "Interviewed"
Case "BI/DT Completed"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Temp Offer Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Start Date Confirmed"
DestArray(DestRow, 2) = "Hired"
Case "Hired"
DestArray(DestRow, 2) = "Hired"
Case "Temp Assignment Started"
DestArray(DestRow, 2) = "Hired"
Case "Temp Assignment Ended"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub NovoNordiskUS()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("B:B").Delete
Range("D:D").Delete
Range("E:F").Delete
Range("G:H").Delete
Range("H:H").Delete
Range("I:J").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:C1").Value = "A"
Range("D1").Value = "Apply Completed"
Range("E1").Value = "Qualified"
Range("F1").Value = "Interviewed"
Range("G1").Value = "Offer Made"
Range("H1").Value = "Hired"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 5

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = ActiveSheet.Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 10)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)


For CurRow = 3 To LastRow
      
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 4)
                DestArray(DestRow, 10) = "Apply Completed"
                
        For CurCol = 5 To 8
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(LastRow + i, 1) = SourceArray(CurRow, 1)
                DestArray(LastRow + i, 2) = SourceArray(CurRow, 2)
                DestArray(LastRow + i, 3) = SourceArray(CurRow, 3)
                DestArray(LastRow + i, 4) = SourceArray(CurRow, 4)
                DestArray(LastRow + i, 5) = SourceArray(CurRow, 5)
                DestArray(LastRow + i, 6) = SourceArray(CurRow, 6)
                DestArray(LastRow + i, 7) = SourceArray(CurRow, 7)
                DestArray(LastRow + i, 8) = SourceArray(CurRow, 8)
                DestArray(LastRow + i, 9) = SourceArray(CurRow, CurCol)
                DestArray(LastRow + i, 10) = SourceArray(1, CurCol)
                
                i = i + 1
                
            Else
            End If
        Next CurCol
        
        DestRow = DestRow + 1
        
Next CurRow

ActiveSheet.Range("1:1").Delete

ActiveSheet.Range("A1:J" & LastRow + i - 1).Value = DestArray

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("K:K").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("K:K").Cut Destination:=Range("C:C")
Range("F:L").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("H2:H" & LastRow + i).Formula = Range("C2:C" & LastRow + i).Value2
Range("H2:H" & LastRow + i).Select
Selection.Copy
Range("C2:C" & LastRow + i).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & LastRow + i).Select
Selection.Copy
Range("C2:C" & LastRow + i).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & LastRow + i).Delete
Range("C2:C" & LastRow + i).NumberFormat = "mm-dd-yyyy"

ActiveSheet.Range("A1:G" & LastRow + i).Font.Size = 10
ActiveSheet.Range("A1:G" & LastRow + i).Font.Name = "Arial"
ActiveSheet.Range("A1:G1").Font.Color = vbBlack
ActiveSheet.Range("A1:G1").Font.Bold = True
ActiveSheet.Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & LastRow + i).Borders.Weight = xlThin
Range("A1:G" & LastRow + i).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub Ameriprise()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:1").Delete
Range("A:A").Delete
Range("B:C").Delete
Range("C:F").Delete
Range("E:F").Delete
Range("F:K").Delete

Range("A1").Value = "JobID1"
Range("B1").Value = "Email"
Range("C1").Value = "Status"
Range("D1").Value = "Date"
Range("E1").Value = "Title"

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("C:C").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("H:H").Cut Destination:=Range("E:E")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("B1").Value = "Status"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Linked by HR"
DestArray(DestRow, 2) = "Apply Completed"
Case "Applied"
DestArray(DestRow, 2) = "Apply Completed"
Case "Linked Questionnaire"
DestArray(DestRow, 2) = "Apply Completed"
Case "BVQ Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "BVQ Sent"
DestArray(DestRow, 2) = "Apply Completed"
Case "Duplicate Applicant"
DestArray(DestRow, 2) = "Apply Completed"
Case "Not Eligible"
DestArray(DestRow, 2) = "Apply Completed"
Case "Reject will be not selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "Hold"
DestArray(DestRow, 2) = "Apply Completed"
Case "Applicant - Draft"
DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrawn"
DestArray(DestRow, 2) = "Apply Completed"
Case "Declined- Email"
DestArray(DestRow, 2) = "Apply Completed"
Case "HR Contacting"
DestArray(DestRow, 2) = "Qualified"
Case "Assessment 2 b Passed"
DestArray(DestRow, 2) = "Qualified"
Case "Pending"
DestArray(DestRow, 2) = "Qualified"
Case "Assessment 2 a Scheduled"
DestArray(DestRow, 2) = "Qualified"
Case "Assessment 1 b Passed"
DestArray(DestRow, 2) = "Qualified"
Case "Assessment 1 a Scheduled"
DestArray(DestRow, 2) = "Qualified"
Case "Refer to Hiring Leader"
DestArray(DestRow, 2) = "Qualified"
Case "Screen Qualified"
DestArray(DestRow, 2) = "Qualified"
Case "SAI Securities of America"
DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Interview"
DestArray(DestRow, 2) = "Qualified"
Case "Route"
DestArray(DestRow, 2) = "Qualified"
Case "Candidate not interested"
DestArray(DestRow, 2) = "Qualified"
Case "P2P Disqualified"
DestArray(DestRow, 2) = "Interviewed"
Case "P2P OK"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - 2"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - 3"
DestArray(DestRow, 2) = "Interviewed"
Case "Declined- Verbal"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Denied"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Pending"
DestArray(DestRow, 2) = "Offer Made"
Case "Intiate Hire going to Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Ready to Hire"
DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub TimeWarner()

Application.ScreenUpdating = False
Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("A:D").Delete
Range("G:G").Delete
Range("P:AP").Delete
Range("D:D").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:C1").Value = "A"
Range("D1").Value = "Apply Completed"
Range("E1").Value = "Apply Completed"
Range("F1").Value = "Interviewed"
Range("G1").Value = "Qualified"
Range("H1").Value = "Interviewed"
Range("I1").Value = "Interviewed"
Range("J1").Value = "Interviewed"
Range("K1").Value = "Offer Made"
Range("L1").Value = "Offer Made"
Range("M1").Value = "Hired"


Dim LastRow
LastRow = Range("A500000").End(xlUp).Row

Range("A:M").Select
Selection.Replace What:="=", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 4

Dim DestRow
DestRow = 2

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:M" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 500000, 1 To 15)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)
    DestArray(1, 13) = SourceArray(2, 13)

For CurRow = 3 To LastRow
                   
        For CurCol = 4 To 13
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 15) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            End If
        Next CurCol
               
Next CurRow

Sheets(1).Range("A1:O" & DestRow).Value = DestArray

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("Q:Q").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("Q:Q").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("H:Q").Delete

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & DestRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Sheets(1).Range("A1:G" & DestRow).Font.Size = 10
Sheets(1).Range("A1:G" & DestRow).Font.Name = "Arial"
Sheets(1).Range("A1:G1").Font.Color = vbBlack
Sheets(1).Range("A1:G1").Font.Bold = True
Sheets(1).Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub


Sub CapitalOne()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:C").Delete
Range("B:B").Delete
Range("F:G").Delete

Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Campus New - Auto Decline - Did not meet BQs"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - Declined - Auto Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - Declined - Manual Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - Move Forward (Met BQ)"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - Unreviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - Candidate Decline"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - Candidate Withdrawn"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - On Hold"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - Declined - Auto Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - Declined - Manual Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Declined - Auto Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Declined - Manual Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Unreviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Auto Decline - Did not meet BQs"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Candidate Withdrawn"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Decline - Auto Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Decline - Manual Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Declined-Failed PreVisor"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Declined-Failed Testing"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Declined-Not eligible for rehire"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Hold - Other Req"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Screening"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Testing"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Unreviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Candidate Decline"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Auto Decline - Did not meet BQs"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Candidate Decline"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Candidate Withdrawn"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Decline - Auto Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Declined - Manual Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Move Forward"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Move Forward (Met BQ)"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Not Eligible to Apply"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - On Hold"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Unreviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Auto Decline - Did not meet BQs"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Candidate Decline"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Candidate Withdrawn"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Declined - Auto Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Declined - Manual Notify"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Move Forward"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Move Forward (Met BQ)"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Not Eligible to Apply"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - On Hold"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Unreviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Acclaim Process - Has Declined"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus New - Has Declined"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Has Declined"
DestArray(DestRow, 2) = "Apply Completed"
Case "Campus Screen - Candidate Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Declined - Auto Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Declined - Manual Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Initiate Testing"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Testing Failed"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Testing Passed"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Candidate Decline"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Evaluate for Testing"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Testing Not Required"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Testing Required"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Transfer to Subtrack Req"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Declined - Auto Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Screen - Declined - Manual Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Move Forward"
DestArray(DestRow, 2) = "Qualified"
Case "Acclaim Process - Phone Interview"
DestArray(DestRow, 2) = "Qualified"
Case "Acclaim Process - Complete - Manual"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Candidate Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Candidate Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Declined - Auto Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Declined - Manual Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - In Process - Hiring Manager"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - In Process - Recruiter"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - In Process on Another Req"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Left Message/Email"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Move Forward"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - On Hold"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Sharing Candidate Profile"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Candidate Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Declined - Auto Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Declined - Manual Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Candidate Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Candidate Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Declined - Auto Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Declined - Manual Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - In Process - Hiring Manager"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - In Process - Recruiter"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - In Process on Another Req"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Left Message/Email"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Move Forward"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - On Hold"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Sharing Candidate Profile"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Candidate Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Declined - Auto Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Declined - Manual Notify"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Failed"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Initiate"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Not Required"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - On Hold"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Passed"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Required"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Evaluate"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Move Forward"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Has Declined"
DestArray(DestRow, 2) = "Qualified"
Case "Testing - Has Declined"
DestArray(DestRow, 2) = "Qualified"
Case "Campus Final Round Interview - Candidate Withdrawn"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Declined - Auto Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Declined - Manual Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Candidate Decline"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Confirmed for PD"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Eligible - Proceed to Offer"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Invited to PD"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Schedule for Power Day"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Screen - Invite/Preselect for 1st Round"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Screen - Move Forward"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Screen - On Hold"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Screen - Round 1 - Fail"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Screen - Round 1 - Pass"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Screen - Round 1 - Retrack"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Declined - Auto Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Decline - Manual Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Declined - Auto Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Declined - Manual Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Testing Passed"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - 1st Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Candidate Decline"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Candidate Withdrawn"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Evaluate if Testing is Needed"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Initiate Testing"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Interview/Testing"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - On Hold"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Proceed to Offer"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Testing Failed"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Testing Not Required"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - Testing Required"
DestArray(DestRow, 2) = "Interviewed"
Case "Acclaim Process - Complete"
DestArray(DestRow, 2) = "Interviewed"
Case "Acclaim Process - In Person Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Post Acclaim Process - Candidate Withdrawn"
DestArray(DestRow, 2) = "Interviewed"
Case "Post Acclaim Process - Declined - Auto Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Post Acclaim Process - Declined - Manual Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Post Acclaim Process - Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Post Acclaim Process - Candidate Decline"
DestArray(DestRow, 2) = "Interviewed"
Case "Post Acclaim Process - Move to Offer"
DestArray(DestRow, 2) = "Interviewed"
Case "Post Acclaim Process - Recommend/Waiting Pool"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Candidate Withdrawn"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Declined - Auto Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Declined - Manual Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Candidate Withdrawn"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Declined - Auto Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Declined - Manual Notify"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - On Hold"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Phone Screen"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Powerday"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Schedule Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - 1st Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - 2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - 3rd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Proceed to Offer"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - 3rd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Campus Final Round Interview - Hold"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Has Declined"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview/Testing - 2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Post Acclaim Process - Has Declined"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer - Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Has Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Declined - Manual Notify"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Pass"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Post Offer Application"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Has Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Declined - Manual Notify"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Has Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Refused"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Canceled"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Draft"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - In Negotiation"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Offer to be made"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Rejected"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Reneged"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Candidate Withdrawn"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Eligible - Externals/Rehires Only"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Fail"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Missing Information"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Provisional - Results Pending"
DestArray(DestRow, 2) = "Offer Made"
Case "PEC - Request PEC"
DestArray(DestRow, 2) = "Offer Made"
Case "Hire Set Up - Start Confirmed"
DestArray(DestRow, 2) = "Hired"
Case "Hire Set Up - Start Confirmed"
DestArray(DestRow, 2) = "Hired"
Case "Hire Set Up - Did Not Start"
DestArray(DestRow, 2) = "Hired"
Case "Hire Set Up - Start Confirmed"
DestArray(DestRow, 2) = "Hired"
Case "Hire Set Up - Start Confirmed"
DestArray(DestRow, 2) = "Hired"
Case "Hire Set Up - Create/Update Employee Record"
DestArray(DestRow, 2) = "Hired"
Case "Hire Set Up - Start Confirmed"
DestArray(DestRow, 2) = "Hired"
Case "Hire Set Up - Candidate Decline"
DestArray(DestRow, 2) = "Hired"
Case "Hire Set Up - Exception Results Clear"
DestArray(DestRow, 2) = "Hired"

End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub CACI()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:1").Delete
Range("A:A").Delete
Range("C:C").Delete
Range("D:D").Delete

Dim LastRow
LastRow = Range("B65536").End(xlUp).Row

Range("AA2:AA" & LastRow).Formula = "=IF(OR(J2="""",Z2=""Canceled""),""delete"","""")"

Dim CurRow1
CurRow1 = 2

Do While CurRow1 < LastRow
If Range("AA" & CurRow1).Value = "delete" Then
Range(CurRow1 & ":" & CurRow1).Delete
LastRow = Range("B65536").End(xlUp).Row
Else
CurRow1 = CurRow1 + 1
End If
Loop

Range("F:AA").Delete

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("C:C")
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("E:E")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

LastRow = Range("D65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow
Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "New Candidate"
DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified"
DestArray(DestRow, 2) = "Apply Completed"
Case "Screening"
DestArray(DestRow, 2) = "Apply Completed"
Case "Has Declined"
DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate selected to interview"
DestArray(DestRow, 2) = "Qualified"
Case "Manager Reviewing Candidate"
DestArray(DestRow, 2) = "Qualified"
Case "Interview Completed"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview Confirmation Sent"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer to be made"
DestArray(DestRow, 2) = "Interviewed"
Case "Approval in Progress"
DestArray(DestRow, 2) = "Interviewed"
Case "Approval Rejected"
DestArray(DestRow, 2) = "Interviewed"
Case "Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Approved"
DestArray(DestRow, 2) = "Offer Made"
Case "Canceled"
DestArray(DestRow, 2) = "Offer Made"
Case "Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "In Negotiation"
DestArray(DestRow, 2) = "Offer Made"
Case "Refused"
DestArray(DestRow, 2) = "Offer Made"
Case "Reneged"
DestArray(DestRow, 2) = "Offer Made"
Case "Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Draft"
DestArray(DestRow, 2) = "Offer Made"
Case "Complete Hire Process"
DestArray(DestRow, 2) = "Hired"
Case "Filled with 1099/Consultant"
DestArray(DestRow, 2) = "Hired"
Case "Filled with CACI Employee"
DestArray(DestRow, 2) = "Hired"
Case "Filled with External Candidate"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub IntelV1()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("A:E").Delete
Range("A1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert

Range("T:AB").Cut Destination:=Range("A:I")
Range("T:AB").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:I1").Value = "A"
Range("J1").Value = "Apply Completed"
Range("K1").Value = "Qualified"
Range("L1").Value = "Interviewed"
Range("M1").Value = "Interviewed"
Range("N1").Value = "Offer Made"
Range("O1").Value = "Offer Made"
Range("P1").Value = "Offer Made"
Range("Q1").Value = "Offer Made"
Range("R1").Value = "Hired"
Range("S1").Value = "Hired"

Dim LastRow
LastRow = Range("A400000").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 11

Dim DestRow
DestRow = 2

Dim i As Long
i = 0

Dim SourceArray As Variant
SourceArray = ActiveSheet.Range("A1:S" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 400000, 1 To 21)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)
    DestArray(1, 13) = SourceArray(2, 13)
    DestArray(1, 14) = SourceArray(2, 14)
    DestArray(1, 15) = SourceArray(2, 15)
    DestArray(1, 16) = SourceArray(2, 16)
    DestArray(1, 17) = SourceArray(2, 17)
    DestArray(1, 18) = SourceArray(2, 18)
    DestArray(1, 19) = SourceArray(2, 19)

For CurRow = 3 To LastRow
      
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, 10)
                DestArray(DestRow, 21) = "Apply Completed"
                
        For CurCol = 11 To 19
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(LastRow + i, 1) = SourceArray(CurRow, 1)
                DestArray(LastRow + i, 2) = SourceArray(CurRow, 2)
                DestArray(LastRow + i, 3) = SourceArray(CurRow, 3)
                DestArray(LastRow + i, 4) = SourceArray(CurRow, 4)
                DestArray(LastRow + i, 5) = SourceArray(CurRow, 5)
                DestArray(LastRow + i, 6) = SourceArray(CurRow, 6)
                DestArray(LastRow + i, 7) = SourceArray(CurRow, 7)
                DestArray(LastRow + i, 8) = SourceArray(CurRow, 8)
                DestArray(LastRow + i, 9) = SourceArray(CurRow, 9)
                DestArray(LastRow + i, 10) = SourceArray(CurRow, 10)
                DestArray(LastRow + i, 11) = SourceArray(CurRow, 11)
                DestArray(LastRow + i, 12) = SourceArray(CurRow, 12)
                DestArray(LastRow + i, 13) = SourceArray(CurRow, 13)
                DestArray(LastRow + i, 14) = SourceArray(CurRow, 14)
                DestArray(LastRow + i, 15) = SourceArray(CurRow, 15)
                DestArray(LastRow + i, 16) = SourceArray(CurRow, 16)
                DestArray(LastRow + i, 17) = SourceArray(CurRow, 17)
                DestArray(LastRow + i, 18) = SourceArray(CurRow, 18)
                DestArray(LastRow + i, 19) = SourceArray(CurRow, 19)
                DestArray(LastRow + i, 20) = SourceArray(CurRow, CurCol)
                DestArray(LastRow + i, 21) = SourceArray(1, CurCol)
                
                i = i + 1
                
            Else
            End If
        Next CurCol
        
        DestRow = DestRow + 1
        
Next CurRow

ActiveSheet.Range("1:1").Delete

ActiveSheet.Range("A1:U" & LastRow + i - 1).Value = DestArray

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("J:J").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("W:W").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("W:W").Cut Destination:=Range("C:C")
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("J:K").Cut Destination:=Range("D:E")
Range("J:K").Delete
Range("L:L").Delete
Range("F:K").Cut Destination:=Range("V:AA")
Range("H:K").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

ActiveSheet.Range("A1:W" & LastRow + i).Font.Size = 10
ActiveSheet.Range("A1:W" & LastRow + i).Font.Name = "Arial"
ActiveSheet.Range("A1:W1").Font.Color = vbBlack
ActiveSheet.Range("A1:W1").Font.Bold = True
ActiveSheet.Range("A1:W1").Interior.Color = vbYellow

Range("A1:W" & LastRow + i).Borders.Weight = xlThin
Range("A1:W" & LastRow + i).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub BonSecours()
Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("A:D").Delete
Range("C:D").Delete
Range("D:K").Delete

Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("B:B").Cut Destination:=Range("D:D")
Range("N:N").Cut Destination:=Range("B:B")
Range("J:J").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:C1").Value = "A"
Range("D1").Value = "Apply Completed"
Range("E1").Value = "Apply Completed"
Range("F1").Value = "Qualified"
Range("G1").Value = "Interviewed"
Range("H1").Value = "Interviewed"
Range("I1").Value = "Offer Made"
Range("J1").Value = "Hired"
Range("K1").Value = "Apply Completed"
Range("L1").Value = "Qualified"

Dim LastRow
LastRow = Range("A100000").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 4

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = ActiveSheet.Range("A1:L" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 400000, 1 To 14)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)

For CurRow = 3 To LastRow
                   
        For CurCol = 4 To 12
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 14) = SourceArray(1, CurCol)
                
                DestRow = DestRow + 1
            Else
            End If

        Next CurCol
        
Next CurRow

ActiveSheet.Range("1:1").Delete

ActiveSheet.Range("A1:N" & DestRow - 1).Value = DestArray

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("O:O").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("O:O").Cut Destination:=Range("C:C")
Range("F:N").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

LastRow = Range("A100000").End(xlUp).Row

Range("H2:H" & LastRow + i).Formula = Range("C2:C" & LastRow + i).Value2
Range("H2:H" & LastRow + i).Select
Selection.Copy
Range("C2:C" & LastRow + i).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & LastRow + i).Select
Selection.Copy
Range("C2:C" & LastRow + i).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & LastRow + i).Delete
Range("C2:C" & LastRow + i).NumberFormat = "mm-dd-yyyy"

ActiveSheet.Range("A1:G" & LastRow + i).Font.Size = 10
ActiveSheet.Range("A1:G" & LastRow + i).Font.Name = "Arial"
ActiveSheet.Range("A1:G1").Font.Color = vbBlack
ActiveSheet.Range("A1:G1").Font.Bold = True
ActiveSheet.Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & LastRow + i).Borders.Weight = xlThin
Range("A1:G" & LastRow + i).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub CUNA()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Range("C:C").Delete
Range("F:F,H:H,J:J,L:L,N:N,P:P,R:R,T:T,V:V,X:X,Z:Z,AB:AB,AD:AD,AF:AF,AH:AH,AJ:AL").Delete
Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("A2:A" & LastRow).Formula = "=IF(B2="""",C2,B2)"
Range("A2:A" & LastRow).Select
Selection.Copy
Range("A2:A" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("B:C").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:D1").Value = "A"
Range("D1").Value = "Apply Completed"
Range("E1").Value = "Apply Completed"
Range("F1").Value = "Qualified"
Range("G1").Value = "Interviewed"
Range("H1").Value = "Interviewed"
Range("I1").Value = "Interviewed"
Range("J1").Value = "Offer Made"
Range("K1").Value = "Interviewed"
Range("L1").Value = "Offer Made"
Range("M1").Value = "Offer Made"
Range("N1").Value = "Hired"
Range("O1").Value = "Hired"
Range("P1").Value = "Hired"
Range("Q1").Value = "Hired"
Range("R1").Value = "Offer Made"
Range("S1").Value = "Hired"

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 4

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:U" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 21)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)
    DestArray(1, 13) = SourceArray(2, 13)
    DestArray(1, 14) = SourceArray(2, 14)
    DestArray(1, 15) = SourceArray(2, 15)
    DestArray(1, 16) = SourceArray(2, 16)
    DestArray(1, 17) = SourceArray(2, 17)
    DestArray(1, 18) = SourceArray(2, 18)
    DestArray(1, 19) = SourceArray(2, 19)
    DestArray(1, 20) = SourceArray(2, 20)
    DestArray(1, 21) = SourceArray(2, 21)
    
For CurRow = 3 To LastRow
                   
        For CurCol = 4 To 21
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 21) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            End If
        Next CurCol
               
Next CurRow


Sheets(1).Range("1:1").Delete

Sheets(1).Range("A1:U" & DestRow).Value = DestArray

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("V:V").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("V:V").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("H:W").Delete

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & DestRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Sheets(1).Range("A1:G" & DestRow).Font.Size = 10
Sheets(1).Range("A1:G" & DestRow).Font.Name = "Arial"
Sheets(1).Range("A1:G1").Font.Color = vbBlack
Sheets(1).Range("A1:G1").Font.Bold = True
Sheets(1).Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

MsgBox "Loop in internalid from job dimension."

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub Attero()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A:A").Delete
Range("B:E").Delete
Range("J2:J" & LastRow).Formula = "=IF(ISBLANK(H2),""Apply Completed"",""Hired"")"
Range("K2:K" & LastRow).Formula = "=IF(J2=""Hired"",H2,MAX(E2,B2))"

Range("J2:K" & LastRow).Select
Selection.Copy
Range("J2:K" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("B:B").Delete
Range("D:H").Delete

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("A1:G" & LastRow).Font.Size = 10
Range("A1:G" & LastRow).Font.Name = "Arial"
Range("A1:G1").Font.Color = vbBlack
Range("A1:G1").Font.Bold = True
Range("A1:G1").Interior.Color = vbYellow

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:G" & LastRow).Borders.Weight = xlThin
Range("A1:G" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If

MsgBox "Split file by JobID1 by numeric and alpha numeric. If on the same file, alpha numeric will not load."
    
End Sub
Sub IntegrityStaffing()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("C:C").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Delete
Range("D:D").Cut Destination:=Range("J:J")
Range("D:E").Delete

Range("G2:G" & LastRow).Formula = "=LEFT(E2,3)"

Dim CurRow1
CurRow1 = 2

Do While CurRow1 < LastRow
If Range("G" & CurRow1).Value = "ISS" Then
Range(CurRow1 & ":" & CurRow1).Delete
LastRow = Range("A200000").End(xlUp).Row
Else
CurRow1 = CurRow1 + 1
End If
Loop

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Apply"
DestArray(DestRow, 2) = "Apply Completed"
Case "Submitted"
DestArray(DestRow, 2) = "Apply Completed"
Case "Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Hired"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("E:G").Delete
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("E1").Select
ActiveCell.EntireColumn.Insert

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Mayo()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("I:I").Delete
Range("J:J").Cut Destination:=Range("G:G")

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
'test
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "0-Filed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Longer Being Considered"
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Longer Being Considered - Physician"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Send to Search Committee"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Submitted"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Submitted- Physician"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Submitted - Research Temporary Professional"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew through Recruiter"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Application in Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Pending - No Longer Being Considered"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Review File/Credentials"
 DestArray(DestRow, 2) = "Qualified"
Case "Assessment"
 DestArray(DestRow, 2) = "Qualified"
Case "Forward to Hiring Manager"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Additional Interviews"
 DestArray(DestRow, 2) = "Interviewed"
Case "First Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "First Interview - Physician"
 DestArray(DestRow, 2) = "Interviewed"
Case "GreenJobs Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Salary/Benefit Meeting"
 DestArray(DestRow, 2) = "Interviewed"
Case "Accepted Another Position- Mayo"
 DestArray(DestRow, 2) = "Offer Made"
Case "Candidate back out of Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Rescind Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hire New Employee"
 DestArray(DestRow, 2) = "Hired"
Case "Appt Letter/Move Auth/Start Date"
 DestArray(DestRow, 2) = "Hired"
Case "Hire Demotion"
 DestArray(DestRow, 2) = "Hired"
Case "Hire Promotion"
 DestArray(DestRow, 2) = "Hired"
Case "Hire Transfer"
 DestArray(DestRow, 2) = "Hired"
Case "Hire Non-Lawson"
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I:J").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub CrossMark()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("B:B").Delete
Range("C:E").Delete
Range("D:D").Delete
Range("F:I").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("I:I").Delete

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim CurCol
CurCol = 1

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:I" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 9)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus

Case "0-Filed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew"
DestArray(DestRow, 2) = "Apply Completed"
Case "Not Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Review Resume"
DestArray(DestRow, 2) = "Apply Completed"
Case "Contact Candidate"
DestArray(DestRow, 2) = "Apply Completed"
Case "Did Not Meet Minimum Qualifications"
DestArray(DestRow, 2) = "Apply Completed"
Case "No Interest-Reviewed Resume"
DestArray(DestRow, 2) = "Apply Completed"
Case "No Interest-Request Mgr. Interview"
DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Gateway"
DestArray(DestRow, 2) = "Apply Completed"
Case "HR Disqualified"
DestArray(DestRow, 2) = "Offer Made"
Case "No Interest-Phone Screen"
DestArray(DestRow, 2) = "Apply Completed"
Case "No interest - Manager"
DestArray(DestRow, 2) = "Interviewed"
Case "Met Minimum Qualifications"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Sent To Mgr - Sams"
DestArray(DestRow, 2) = "Qualified"
Case "Request Manager Inteview – Convenience Solutions"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview – Kimberly Clark"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview – Novartis"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview – J&J"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview - Events"
DestArray(DestRow, 2) = "Interviewed"
Case "Phone Screen/Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Additional Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Verbal Offer Accepted (USA Only)"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada Part-Time Hourly)"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Conditional Offer"
DestArray(DestRow, 2) = "Offer Made"
Case "No Interest-Addtl Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "No Interest-Request Candidate Audition"
DestArray(DestRow, 2) = "Interviewed"
Case "Proceed with candidate - Manager"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Candidate Audition - Dannon"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Candidate Audition - Sam's"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview – Auto Advisor"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview-Canada CRT"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview-DCA Coder"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview-DCA Collector"
DestArray(DestRow, 2) = "Interviewed"
Case "Verbal Offer Accepted (Canada – Internal Transfer – FT Hourly)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada – Internal Transfer – FT Salaried)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada – Internal Transfer – PT Hourly)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada Full Time Hourly)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada Full Time Salaried)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada Only)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada-Novartis-FT Hourly - English)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada-Novartis-FT Hourly – French)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada-Novartis-FT Salary - English)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada-Novartis-FT Salary - French)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada-Novartis-PT Hourly - English)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Canada-Novartis-PT Hourly - French)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Quebec Full Time Hourly-English)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Quebec Full Time Hourly-French)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Quebec Full Time Salaried-English)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Quebec Full Time Salaried-French)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Quebec Part Time Hourly-English)"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted (Quebec Part Time Hourly-French)"
DestArray(DestRow, 2) = "Offer Made"
Case "Request Manager Interview – Walmart Events"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview - Nestle"
DestArray(DestRow, 2) = "Interviewed"
Case "Candidate Considering"
DestArray(DestRow, 2) = "Apply Completed"
Case "No Interest-Contacted Candidate"
DestArray(DestRow, 2) = "Apply Completed"
Case "Request Manager Interview - L'Oreal"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview – NCiM Events"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview - US CRT"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview - Wet Sampling Events"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview-Samsung"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview-Waste Mgmt Events"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview–Walmart Retail Rep"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview – Mead Johnson"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview - Kraft 727"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview–Glidden"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview - J&J Pilot"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview-Bil-Jac Events"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview-BJ's Event"
DestArray(DestRow, 2) = "Interviewed"
Case "Request Manager Interview-Kraft 726"
DestArray(DestRow, 2) = "Interviewed"
End Select
                               
                DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:I" & DestRow).Value = DestArray

CurRow = 2

Do While CurRow <= LastRow
If Range("I" & CurRow).Value > 0 Then
Range("B" & CurRow).Value = "Hired"
Range("C" & CurRow).Value = Range("I" & CurRow).Value
Else
End If
CurRow = CurRow + 1
Loop

Range("I:I").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("I2:I" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & DestRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I1:I" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Sheets(1).Range("A1:H" & DestRow).Font.Size = 10
Sheets(1).Range("A1:H" & DestRow).Font.Name = "Arial"
Sheets(1).Range("A1:H1").Font.Color = vbBlack
Sheets(1).Range("A1:H1").Font.Bold = True
Sheets(1).Range("A1:H1").Interior.Color = vbYellow

Range("A1:H" & DestRow).Borders.Weight = xlThin
Range("A1:H" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub Unisys()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:E").Delete
Range("F:N").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("C:C").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("H:H").Cut Destination:=Range("E:E")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "005 Draft"
DestArray(DestRow, 2) = "ATS Captured"
Case "115 Reject Online Screening"
DestArray(DestRow, 2) = "ATS Captured"
Case "010 Review"
DestArray(DestRow, 2) = "Apply Completed"
Case "015 Linked"
DestArray(DestRow, 2) = "Apply Completed"
Case "020-Applied"
DestArray(DestRow, 2) = "Apply Completed"
Case "030-Screen"
DestArray(DestRow, 2) = "Apply Completed"
Case "050-Route"
DestArray(DestRow, 2) = "Qualified"
Case "060-Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "069 Preliminary Offer Decided"
DestArray(DestRow, 2) = "Offer Made"
Case "070-Offer"
DestArray(DestRow, 2) = "Offer Made"
Case "071 Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "077-Offer Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "080-Ready to Hire"
DestArray(DestRow, 2) = "Hired"
Case "090-Hired"
DestArray(DestRow, 2) = "Hired"
Case "100-Hold"
DestArray(DestRow, 2) = "Apply Completed"
Case "110-Reject"
DestArray(DestRow, 2) = "Apply Completed"
Case "112 Failed Prescreening"
DestArray(DestRow, 2) = "Apply Completed"
Case "120-Withdrawn"
DestArray(DestRow, 2) = "Apply Completed"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I:J").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub NovoNordiskDK()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("C:F").Delete
Range("F:G").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Application withdrawn"
DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Applied"
DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Not Interested"
DestArray(DestRow, 2) = "Apply Completed"
Case "Headcount Removed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Not Considered"
DestArray(DestRow, 2) = "Apply Completed"
Case "Reject After On Hold (automatic email)"
DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Not Qualified"
DestArray(DestRow, 2) = "Qualified"
Case "Reject Before Interview (automatic email)"
DestArray(DestRow, 2) = "Qualified"
Case "Reject Before Interview (No email)"
DestArray(DestRow, 2) = "Qualified"
Case "Manager Resume Review (CN)"
DestArray(DestRow, 2) = "Qualified"
Case "More Qualified Candidate"
DestArray(DestRow, 2) = "Qualified"
Case "No Interest: HM Review"
DestArray(DestRow, 2) = "Qualified"
Case "No Interest: LOB Phone Screen"
DestArray(DestRow, 2) = "Qualified"
Case "No Interest: Manager Screen Review"
DestArray(DestRow, 2) = "Qualified"
Case "No Interest: Recruiter Phone Screen"
DestArray(DestRow, 2) = "Qualified"
Case "No Interest: Resume Review"
DestArray(DestRow, 2) = "Qualified"
Case "No Interest: Salary Requirements"
DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Phone Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Resume Review"
DestArray(DestRow, 2) = "Qualified"
Case "Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Reject After Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview (CN)"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "No Interest: 1st Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "No Interest: 2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("I:J").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Sprint()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("E:G").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:D1").Value = "A"
Range("E1").Value = "Apply Completed"
Range("F1").Value = "Qualified"
Range("G1").Value = "Interviewed"
Range("H1").Value = "Offer Made"
Range("I1").Value = "Hired"

Dim LastRow
LastRow = Range("A300000").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 5

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:I" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 400000, 1 To 11)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    
For CurRow = 3 To LastRow
                   
        For CurCol = 5 To 9
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 11) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            End If
        Next CurCol
               
Next CurRow


Sheets(1).Range("1:1").Delete

Sheets(1).Range("A1:K" & DestRow).Value = DestArray

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("L:L").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("L:L").Cut Destination:=Range("C:C")
Range("G:K").Delete
Range("D:D").Cut Destination:=Range("H:H")
Range("D:D").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & DestRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Sheets(1).Range("A1:G" & DestRow).Font.Size = 10
Sheets(1).Range("A1:G" & DestRow).Font.Name = "Arial"
Sheets(1).Range("A1:G1").Font.Color = vbBlack
Sheets(1).Range("A1:G1").Font.Bold = True
Sheets(1).Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub SAIC()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("AE2:AE" & LastRow).Formula = "=if(P2 = ""Internal Transfer"",""delete"","""")"

Range("AE2:AE" & LastRow).Select
Selection.Copy
Range("AE2:AE" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Dim CurRow1
CurRow1 = 2

Do While CurRow1 < LastRow
If Range("AE" & CurRow1).Value = "delete" Then
Range(CurRow1 & ":" & CurRow1).Delete
LastRow = Range("A65536").End(xlUp).Row
Else
CurRow1 = CurRow1 + 1
End If
Loop

LastRow = Range("A65536").End(xlUp).Row

Range("AG2:AG" & LastRow).Delete

Range("A:D").Delete
Range("B:D").Delete
Range("C:D").Delete
Range("D:E").Delete
Range("F:H").Delete
Range("G:J").Delete
Range("H:L").Delete

LastRow = Range("A65536").End(xlUp).Row

Range("H2:H" & LastRow).Formula = "=""T""&A2"

Range("H2:H" & LastRow).Select
Selection.Copy
Range("A2:A" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H:H").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("C:C")
Range("J:J").Cut Destination:=Range("E:E")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "New"
DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Review"
DestArray(DestRow, 2) = "Qualified"
Case "HM Review"
DestArray(DestRow, 2) = "Qualified"
Case "Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
DestArray(DestRow, 2) = "Offer Made"
Case "Hire"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub USCellular()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("H:K").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("H2:H" & LastRow).Formula = "=D2&E2&F2"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("D:F").Delete

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("C:C")
Range("G1").Select
ActiveCell.EntireColumn.Insert

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Care Center CSWBackground Check_2Candidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSWBackground Check_2In Progress 2 "
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSWBackground Check_2Move to Offer 2"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSWBackground Check_2Rejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSWBackground Check_2Requesting Add'l Information 2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSWBackground Check_2To be initiated 2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSWBackground Check_2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSWHireCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSWHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "Care Center CSWHireRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSWHireTo Be Hired"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSWHire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSWInterview_3Candidate Withdrew"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center CSWInterview_3Face to Face Interview 2"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center CSWInterview_3Interview to be scheduled 2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSWInterview_3Move to Background Check 2"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center CSWInterview_3Rejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center CSWInterview_3"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSWNew-CC_2App_Complete"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWNew-CC_2Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWNew-CC_2Move to PS CC 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWNew-CC_2Recruiter Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWNew-CC_2Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWNew-CC_2To be evaluated 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWOfferCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSWOfferOffer to be made"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSWOfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSWOffer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSWPhone Screen CC 2Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPhone Screen CC 2Left Message2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPhone Screen CC 2Move to Pre-Employment Testing 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPhone Screen CC 2Phone Screen Scheduled 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPhone Screen CC 2Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPhone Screen CC 2To be Phone Screened CC 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPhone Screen CC 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_2Assessment Completed 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_2Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_2DeGarmo Fit Assessment 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_2Move to Pre-Employment Testing 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_2Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_3Assessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_3Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_3DeGarmo Skill Assessment_2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_3Move to Interview 2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSWPre-Employment Testing_3Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSWPre-Employment Testing_3"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Background Check_2Candidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSW 11282012Background Check_2In Progress 2 "
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSW 11282012Background Check_2Move to Offer 2"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSW 11282012Background Check_2Rejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSW 11282012Background Check_2Requesting Add'l Information 2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSW 11282012Background Check_2To be initiated 2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSW 11282012Background Check_2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSW 11282012HireCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSW 11282012HireHired"
 DestArray(DestRow, 2) = "Hired"
Case "Care Center CSW 11282012HireRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSW 11282012HireTo Be Hired"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSW 11282012Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSW 11282012Interview_3Candidate Withdrew"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center CSW 11282012Interview_3Face to Face Interview 2"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center CSW 11282012Interview_3Interview to be scheduled 2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSW 11282012Interview_3Move to Background Check 2"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center CSW 11282012Interview_3Rejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center CSW 11282012Interview_3"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSW 11282012New-CC_2App_Complete"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012New-CC_2Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012New-CC_2Move to PS CC 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012New-CC_2Recruiter Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012New-CC_2Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012New-CC_2To be evaluated 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012OfferCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSW 11282012OfferOffer to be made"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSW 11282012OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSW 11282012Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center CSW 11282012Phone Screen CC 2Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Phone Screen CC 2Left Message2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Phone Screen CC 2Move to Pre-Employment Testing 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Phone Screen CC 2Phone Screen Scheduled 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Phone Screen CC 2Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Phone Screen CC 2To be Phone Screened CC 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Phone Screen CC 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_2Assessment Completed 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_2Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_2DeGarmo Fit Assessment 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_2Move to Pre-Employment Testing 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_2Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_3Assessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_3Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_3DeGarmo Skill Assessment_2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_3Move to Interview 2"
 DestArray(DestRow, 2) = "Qualified"
Case "Care Center CSW 11282012Pre-Employment Testing_3Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center CSW 11282012Pre-Employment Testing_3"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Care Center/PT/Retail/Do Not Use CSWOfferAccepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center/PT/Retail/Do Not Use CSWOfferApproval in Progress"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center/PT/Retail/Do Not Use CSWOfferApproval Rejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center/PT/Retail/Do Not Use CSWOfferApproved"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center/PT/Retail/Do Not Use CSWOfferCanceled"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center/PT/Retail/Do Not Use CSWOfferDraft"
 DestArray(DestRow, 2) = "Interviewed"
Case "Care Center/PT/Retail/Do Not Use CSWOfferExtended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center/PT/Retail/Do Not Use CSWOfferIn Negotiation"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center/PT/Retail/Do Not Use CSWOfferRefused"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center/PT/Retail/Do Not Use CSWOfferReneged"
 DestArray(DestRow, 2) = "Offer Made"
Case "Care Center/PT/Retail/Do Not Use CSWOfferRescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Do Not Use - Care Center CSWBackground CheckRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use - Care Center CSWNewCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use - Care Center CSWNewRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Background CheckCandidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW Background CheckIn Progress"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW Background CheckMove to Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Do Not Use Retail CSW Background CheckRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW Background CheckRequesting Add'l Information"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW Background CheckTo be initiated"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW Background Check"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW HireHired"
 DestArray(DestRow, 2) = "Hired"
Case "Do Not Use Retail CSW HireTo Be Hired"
 DestArray(DestRow, 2) = "Offer Made"
Case "Do Not Use Retail CSW Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Do Not Use Retail CSW InterviewB2B Simulation"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW InterviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW InterviewDDI Assessment"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW InterviewFace to Face Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Do Not Use Retail CSW InterviewInterview to be scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW InterviewMove to Background Check"
 DestArray(DestRow, 2) = "Interviewed"
Case "Do Not Use Retail CSW InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Do Not Use Retail CSW Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW NewCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW NewDeGarmo MVP Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW NewMove to Phone Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW NewRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW NewTo be evaluated"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW OfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "Do Not Use Retail CSW Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Do Not Use Retail CSW Phone ScreenCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Phone ScreenLeft Message"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Phone ScreenMove to Pre-Employment Testing"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Phone ScreenPhone Screen Scheduled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Phone ScreenRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Phone ScreenTo Be Phone Screened"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Phone Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Pre-Employment TestingCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Pre-Employment TestingDeGarmo Fit Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Pre-Employment TestingDeGarmo Skill Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Pre-Employment TestingMove to Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Do Not Use Retail CSW Pre-Employment TestingRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Pre-Employment TestingTesting to be initiated"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use Retail CSW Pre-Employment Testing"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWBackground CheckMove to Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "New -  Retail CSWBackground CheckRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "New -  Retail CSWBackground Check"
 DestArray(DestRow, 2) = "Qualified"
Case "New -  Retail CSWHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "New -  Retail CSWHireTo Be Hired"
 DestArray(DestRow, 2) = "Offer Made"
Case "New -  Retail CSWHire"
 DestArray(DestRow, 2) = "Offer Made"
Case "New -  Retail CSWInterviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Interviewed"
Case "New -  Retail CSWInterviewFace to Face Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "New -  Retail CSWInterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "New -  Retail CSWInterview"
 DestArray(DestRow, 2) = "Qualified"
Case "New -  Retail CSWNew-RTAssessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWNew-RTCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWNew-RTDeGarmo Fit Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWNew-RTDeGarmo MVP Assessment-Retail"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWNew-RTDeGarmo Skill Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWNew-RTMove to Phone Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWNew-RTRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWOfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "New -  Retail CSWOffer"
 DestArray(DestRow, 2) = "Offer Made"
Case "New -  Retail CSWPhone Screen-RTCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWPhone Screen-RTMove to Interview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWPhone Screen-RTRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWPhone Screen-RTTo Be Phone Screened"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New -  Retail CSWPhone Screen-RT"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New - Care Center CSWNew-CCCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New - Care Center CSWOfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "New - Care Center CSWPhone Screen CCRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWBackground / Drug ScreenCandidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWBackground / Drug ScreenIn Progress"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWBackground / Drug ScreenMove to Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Professional / Technical CSWBackground / Drug ScreenRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWBackground / Drug ScreenRequesting Add'l Information"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWBackground / Drug ScreenTo be initiated"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWBackground / Drug Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "Professional / Technical CSWHireTo Be Hired"
 DestArray(DestRow, 2) = "Offer Made"
Case "Professional / Technical CSWHire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Professional / Technical CSWHM ReviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWHM ReviewMove to Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWHM ReviewPending HM Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWHM ReviewRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWHM ReviewShared with HM"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWHM Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWInterviewAdd'l Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Professional / Technical CSWInterviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Interviewed"
Case "Professional / Technical CSWInterviewHM Phone Screen"
 DestArray(DestRow, 2) = "Interviewed"
Case "Professional / Technical CSWInterviewInterview to be scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWInterviewMove to Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Professional / Technical CSWInterviewOnsite Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Professional / Technical CSWInterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Professional / Technical CSWInterview"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWNewApp_Complete"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWNewCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWNewDeGarmo MVP Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWNewMove to Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWNewRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWNewTo be evaluated"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWOfferCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Professional / Technical CSWOfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "Professional / Technical CSWOfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Professional / Technical CSWOffer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Professional / Technical CSWReviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWReviewLeft Message"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWReviewPhone Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWReviewRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWReviewReview in Progress"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Professional / Technical CSWReviewSend to HM for Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Professional / Technical CSWReview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWBackground CheckCandidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSWBackground CheckIn Progress"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSWBackground CheckMove to Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSWBackground CheckRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSWBackground CheckRequesting Add'l Information"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSWBackground CheckTo be initiated"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSWBackground Check"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSWHireCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSWHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "Retail CSWHireRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSWHireTo Be Hired"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSWHire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSWInterviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSWInterviewFace to Face Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSWInterviewInterview to be scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSWInterviewMove to Background Check"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSWInterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSWInterview"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSWNew-RTApp_Complete_0"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWNew-RTCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWNew-RTMove to Phone Screen_0"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWNew-RTRecruiter Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWNew-RTRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWNew-RTTo be evaluated_0"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWOfferCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSWOfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSWOfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSWOffer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSWPhone Screen-RTCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPhone Screen-RTLeft Message"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPhone Screen-RTMove to Pre-Employment Testing"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPhone Screen-RTPhone Screen Scheduled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPhone Screen-RTRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPhone Screen-RTTo Be Phone Screened"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPhone Screen-RT"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPre-Employment TestingAssessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPre-Employment TestingCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPre-Employment TestingDeGarmo Fit Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPre-Employment TestingDeGarmo Skill Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPre-Employment TestingMove to Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSWPre-Employment TestingRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSWPre-Employment Testing"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Background CheckCandidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012Background CheckIn Progress"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012Background CheckMove to Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSW 11282012Background CheckRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012Background CheckRequesting Add'l Information"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012Background CheckTo be initiated"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012Background Check"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012HireCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSW 11282012HireHired"
 DestArray(DestRow, 2) = "Hired"
Case "Retail CSW 11282012HireRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSW 11282012HireTo Be Hired"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSW 11282012Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSW 11282012InterviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSW 11282012InterviewFace to Face Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSW 11282012InterviewInterview to be scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012InterviewMove to Background Check"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSW 11282012InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSW 11282012Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012New-RTApp_Complete_0"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012New-RTCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012New-RTMove to Phone Screen_0"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012New-RTRecruiter Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012New-RTRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012New-RTTo be evaluated_0"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012OfferCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSW 11282012OfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "Retail CSW 11282012OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSW 11282012Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Retail CSW 11282012Phone Screen-RTCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Phone Screen-RTLeft Message"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Phone Screen-RTMove to Pre-Employment Testing"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Phone Screen-RTPhone Screen Scheduled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Phone Screen-RTRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Phone Screen-RTTo Be Phone Screened"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Phone Screen-RT"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre - Employment TestingAssessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre - Employment TestingDeGarmo Skill Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre - Employment TestingMove to Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012Pre - Employment TestingRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre-Employment TestingAssessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre-Employment TestingCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre-Employment TestingDeGarmo Fit Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre-Employment TestingDeGarmo Skill Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre-Employment TestingMove to Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Retail CSW 11282012Pre-Employment TestingRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre-Employment Testing"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEBackground Check_2Candidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Care Center CSW-DO NOT USEBackground Check_2In Progress 2 "
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Care Center CSW-DO NOT USEBackground Check_2Move to Offer 2"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Care Center CSW-DO NOT USEBackground Check_2Rejected"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Care Center CSW-DO NOT USEBackground Check_2Requesting Add'l Information 2"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Care Center CSW-DO NOT USEBackground Check_2To be initiated 2"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Care Center CSW-DO NOT USEBackground Check_2"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Care Center CSW-DO NOT USEHireCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Care Center CSW-DO NOT USEHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "zz-Care Center CSW-DO NOT USEHireRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Care Center CSW-DO NOT USEHireTo Be Hired"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Care Center CSW-DO NOT USEHire"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Care Center CSW-DO NOT USEInterview_3Candidate Withdrew"
 DestArray(DestRow, 2) = "Interviewed"
Case "zz-Care Center CSW-DO NOT USEInterview_3Face to Face Interview 2"
 DestArray(DestRow, 2) = "Interviewed"
Case "zz-Care Center CSW-DO NOT USEInterview_3Interview to be scheduled 2"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Care Center CSW-DO NOT USEInterview_3Move to Background Check 2"
 DestArray(DestRow, 2) = "Interviewed"
Case "zz-Care Center CSW-DO NOT USEInterview_3Rejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "zz-Care Center CSW-DO NOT USEInterview_3"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Care Center CSW-DO NOT USENew-CC_2App_Complete"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USENew-CC_2Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USENew-CC_2Move to PS CC 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USENew-CC_2Recruiter Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USENew-CC_2Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USENew-CC_2To be evaluated 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEOfferCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Care Center CSW-DO NOT USEOfferOffer to be made"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Care Center CSW-DO NOT USEOfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Care Center CSW-DO NOT USEOffer"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Care Center CSW-DO NOT USEPhone Screen CC 2Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPhone Screen CC 2Left Message2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPhone Screen CC 2Move to Pre-Employment Testing 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPhone Screen CC 2Phone Screen Scheduled 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPhone Screen CC 2Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPhone Screen CC 2To be Phone Screened CC 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPhone Screen CC 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_2Assessment Completed 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_2Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_2DeGarmo Fit Assessment 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_2Move to Pre-Employment Testing 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_2Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_3Assessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_3Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_3DeGarmo Skill Assessment_2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_3Move to Interview 2"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_3Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Care Center CSW-DO NOT USEPre-Employment Testing_3"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEBackground CheckCandidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Retail CSW-DO NOT USEBackground CheckIn Progress"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Retail CSW-DO NOT USEBackground CheckMove to Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Retail CSW-DO NOT USEBackground CheckRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Retail CSW-DO NOT USEBackground CheckRequesting Add'l Information"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Retail CSW-DO NOT USEBackground CheckTo be initiated"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Retail CSW-DO NOT USEBackground Check"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Retail CSW-DO NOT USEHireCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Retail CSW-DO NOT USEHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "zz-Retail CSW-DO NOT USEHireRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Retail CSW-DO NOT USEHireTo Be Hired"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Retail CSW-DO NOT USEHire"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Retail CSW-DO NOT USEInterviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Interviewed"
Case "zz-Retail CSW-DO NOT USEInterviewFace to Face Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "zz-Retail CSW-DO NOT USEInterviewInterview to be scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Retail CSW-DO NOT USEInterviewMove to Background Check"
 DestArray(DestRow, 2) = "Interviewed"
Case "zz-Retail CSW-DO NOT USEInterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "zz-Retail CSW-DO NOT USEInterview"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Retail CSW-DO NOT USENew-RTApp_Complete_0"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USENew-RTCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USENew-RTMove to Phone Screen_0"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USENew-RTRecruiter Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USENew-RTRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USENew-RTTo be evaluated_0"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEOfferCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Retail CSW-DO NOT USEOfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "zz-Retail CSW-DO NOT USEOfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Retail CSW-DO NOT USEOffer"
 DestArray(DestRow, 2) = "Offer Made"
Case "zz-Retail CSW-DO NOT USEPhone Screen-RTCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPhone Screen-RTLeft Message"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPhone Screen-RTMove to Pre-Employment Testing"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPhone Screen-RTPhone Screen Scheduled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPhone Screen-RTRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPhone Screen-RTTo Be Phone Screened"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPhone Screen-RT"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPre-Employment TestingAssessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPre-Employment TestingCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPre-Employment TestingDeGarmo Fit Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPre-Employment TestingDeGarmo Skill Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPre-Employment TestingMove to Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "zz-Retail CSW-DO NOT USEPre-Employment TestingRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "zz-Retail CSW-DO NOT USEPre-Employment Testing"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use - Care Center CSWPre-Employment TestingCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use - Care Center CSWPre-Employment TestingRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use - Care Center CSWInterviewRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use - Care Center CSWInterviewRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Retail CSW 11282012Pre - Employment TestingCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Use - Care Center CSWOfferCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Pearson()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:B").Delete
Range("D:D").Delete
Range("F:I").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("D:D")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Candidate Withdrew"
DestArray(DestRow, 2) = "Apply Completed"
Case "External Portal"
DestArray(DestRow, 2) = "Apply Completed"
Case "Internal Portal"
DestArray(DestRow, 2) = "Apply Completed"
Case "No show/backed out"
DestArray(DestRow, 2) = "Apply Completed"
Case "Position Closed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Position Closed; Not Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Position Closed; Req Cancelled"
DestArray(DestRow, 2) = "Apply Completed"
Case "Prescreen Fail"
DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Nominated"
DestArray(DestRow, 2) = "Apply Completed"
Case "Submitted"
DestArray(DestRow, 2) = "Apply Completed"
Case "Vendor Portal"
DestArray(DestRow, 2) = "Apply Completed"
Case "Phone Screen; Not Selected; Did Not Meet Minimum"
DestArray(DestRow, 2) = "Apply Completed"
Case "Reviewed; Not Selected; Did Not Meet Minimum"
DestArray(DestRow, 2) = "Apply Completed"
Case "Referral Portal"
DestArray(DestRow, 2) = "Apply Completed"
Case "Application Request Sent"
DestArray(DestRow, 2) = "Qualified"
Case "Contact"
DestArray(DestRow, 2) = "Qualified"
Case "Contact; Not Selected"
DestArray(DestRow, 2) = "Qualified"
Case "Contact; Not Selected; Unable to Contact"
DestArray(DestRow, 2) = "Qualified"
Case "Assessment"
DestArray(DestRow, 2) = "Qualified"
Case "Interview Requested"
DestArray(DestRow, 2) = "Qualified"
Case "Interview Scheduled"
DestArray(DestRow, 2) = "Qualified"
Case "Phone Screen"
DestArray(DestRow, 2) = "Qualified"
Case "Phone Screen - Not Selected"
DestArray(DestRow, 2) = "Qualified"
Case "Phone Screen; Not Selected; Less Qualified"
DestArray(DestRow, 2) = "Qualified"
Case "Reviewed"
DestArray(DestRow, 2) = "Qualified"
Case "Reviewed; Not Selected"
DestArray(DestRow, 2) = "Qualified"
Case "Reviewed; Not Selected; Less Qualified"
DestArray(DestRow, 2) = "Qualified"
Case "Background/Reference Check Initiated"
DestArray(DestRow, 2) = "Interviewed"
Case "Interview Completed"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed; Not Selected"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed; Not Selected; Did Not Meet Minimum"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed; Not Selected; Less Qualified"
DestArray(DestRow, 2) = "Interviewed"
Case "Launch Offer Approval"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer Requested"
DestArray(DestRow, 2) = "Interviewed"
Case "Prepare Offer Details"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined/Rejected"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub JCPenney()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:A").Delete
Range("G:N").Delete

Dim LastRow
LastRow = Range("A300000").End(xlUp).Row

Range("H2:H" & LastRow).Formula = "=IF(OR(D2=""APP"",D2=""DRF""),""Applied"",E2)"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("D:E").Delete

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("C:C")

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Applied"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hold"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Linked"
 DestArray(DestRow, 2) = "Qualified"
Case "Route"
 DestArray(DestRow, 2) = "Qualified"
Case "BgrdCheck"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Off Accept"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "PreOffRej"
 DestArray(DestRow, 2) = "Offer Made"
Case "Ready Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub CroweHarwoth()
Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("A:B").Delete

Dim LastRow
LastRow = Range("A400000").End(xlUp).Row

Dim CurRow1
CurRow1 = 2

Do While CurRow1 < LastRow
If Left(Range("A" & CurRow1).Value, 4) = "Bin:" Or Left(Range("A" & CurRow1).Value, 22) = "Applicant Flow Status:" Then
Range(CurRow1 & ":" & CurRow1).Delete
Else
CurRow1 = CurRow1 + 1
End If
Loop

LastRow = Range("A400000").End(xlUp).Row

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("Y:Y").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("AA:AA").Cut Destination:=Range("C:C")
Range("R1").Select
ActiveCell.EntireColumn.Insert
Range("T:T").Cut Destination:=Range("R:R")
Range("T:T").Delete
Range("T1").Select
ActiveCell.EntireColumn.Insert
Range("V:V").Cut Destination:=Range("T:T")
Range("V:V").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:N1").Value = "A"
Range("O1").Value = "Qualified"
Range("P1").Value = "Hired"
Range("Q1").Value = "Qualified"
Range("R1").Value = "Interviewed"
Range("S1").Value = "Interviewed"
Range("T1").Value = "Offer Made"
Range("U1").Value = "Offer Made"
Range("V1").Value = "Qualified"
Range("W1").Value = "Qualified"
Range("X1").Value = "Qualified"
Range("Y1").Value = "Apply Completed"

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 15

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:Y" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 400000, 1 To 27)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)
    DestArray(1, 13) = SourceArray(2, 13)
    DestArray(1, 14) = SourceArray(2, 14)
    DestArray(1, 15) = SourceArray(2, 15)
    DestArray(1, 16) = SourceArray(2, 16)
    DestArray(1, 17) = SourceArray(2, 17)
    DestArray(1, 18) = SourceArray(2, 18)
    DestArray(1, 19) = SourceArray(2, 19)
    DestArray(1, 20) = SourceArray(2, 20)
    DestArray(1, 21) = SourceArray(2, 21)
    DestArray(1, 22) = SourceArray(2, 22)
    DestArray(1, 23) = SourceArray(2, 23)
    DestArray(1, 24) = SourceArray(2, 24)
    DestArray(1, 25) = SourceArray(2, 25)

For CurRow = 3 To LastRow
                   
        For CurCol = 15 To 25
            If SourceArray(CurRow, CurCol) <> "" And SourceArray(CurRow, CurCol - 11) = 1 Then
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, 20)
                DestArray(DestRow, 21) = SourceArray(CurRow, 21)
                DestArray(DestRow, 22) = SourceArray(CurRow, 22)
                DestArray(DestRow, 23) = SourceArray(CurRow, 23)
                DestArray(DestRow, 24) = SourceArray(CurRow, 24)
                DestArray(DestRow, 25) = SourceArray(CurRow, 25)
                DestArray(DestRow, 26) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 27) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            
            End If
        Next CurCol
               
Next CurRow


Sheets(1).Range("1:1").Delete

Sheets(1).Range("A1:AA" & DestRow).Value = DestArray

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("AB:AB").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("AB:AB").Cut Destination:=Range("C:C")
Range("F:AA").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & DestRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Sheets(1).Range("A1:G" & DestRow).Font.Size = 10
Sheets(1).Range("A1:G" & DestRow).Font.Name = "Arial"
Sheets(1).Range("A1:G1").Font.Color = vbBlack
Sheets(1).Range("A1:G1").Font.Bold = True
Sheets(1).Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

MsgBox "Remove Intern and University records"

Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub SAFG()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:3").Delete
Range("D:E").Delete
Range("E:O").Delete
Range("G:G").Delete
Range("H:L").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A2:G" & LastRow).Sort Key1:=Range("G2:G" & LastRow), order1:=xlAscending, Key2:=Range("A2:A" & LastRow), order2:=xlAscending, Key3:=Range("B2:B" & LastRow), order3:=xlDescending
Range("H2:H" & LastRow).Formula = "=IF(AND(G1=G2,A1=A2),""Dupe"","""")"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Dim CurRow1
CurRow1 = 2

Do While CurRow1 < LastRow
If Range("H" & CurRow1).Value = "Dupe" Then
Range(CurRow1 & ":" & CurRow1).Delete
Else
CurRow1 = CurRow1 + 1
End If
Loop

LastRow = Range("A65536").End(xlUp).Row

Range("A:B").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Not Suitable"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Inbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected After Prescreen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Prescreen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background Adverse"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected after HM Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Suitable after Recruiter Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Suitable after Resume/CV Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Resume/CV Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected For Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Schedule Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "DM Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre-qualification Kit"
 DestArray(DestRow, 2) = "Qualified"
Case "Licensing & Testing"
 DestArray(DestRow, 2) = "Qualified"
Case "Selection Kit & FINRA"
 DestArray(DestRow, 2) = "Qualified"
Case "Schedule 1st Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Not Selected after Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Not Selected after Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "DM Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "RVP Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Schedule 2nd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Schedule 3rd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Background Check"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Create Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "SRS - Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "SRS - Approve Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "SRS - Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "SRS - Offer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "SRS - Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Create & Approve Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Initiate Written Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Make Verbal Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Background Check"
 DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Written Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "SRS - Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hired-External"
 DestArray(DestRow, 2) = "Hired"
Case "Hired-Internal"
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A2:H" & LastRow).Font.Bold = False
Range("A2:H" & LastRow).Interior.Color = vbWhite

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Comerica()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:1").Delete
Range("E:E").Delete

Dim LastRow
LastRow = Range("C65536").End(xlUp).Row

Dim oneCell As Range, oneCell1 As Range
With Sheets(1)
    For Each oneCell In .UsedRange
        If oneCell.MergeCells Then
            Set oneCell1 = oneCell.MergeArea
            oneCell.UnMerge
            oneCell1.Value = oneCell.Value
        End If
    Next oneCell
End With

Range("M2:M" & LastRow).Formula = "=IF(H2="""",G2,H2)"
Range("M2:M" & LastRow).Select
Selection.Copy
Range("G2:G" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("M:M").Delete

Range("A:A").Delete
Range("D:D").Delete
Range("F:J").Delete

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "New - Disposition"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Meets Basic Qualification"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - To be Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Under Consideration"
DestArray(DestRow, 2) = "Apply Completed"
Case "New - Withdrawn"
DestArray(DestRow, 2) = "Apply Completed"
Case "HM Screen - Disposition"
DestArray(DestRow, 2) = "Qualified"
Case "HM Screen - Move Forward"
DestArray(DestRow, 2) = "Qualified"
Case "HM Screen - Phone Screened"
DestArray(DestRow, 2) = "Qualified"
Case "HM Screen - To be Reviewed"
DestArray(DestRow, 2) = "Qualified"
Case "HM Screen - Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "HR Screen - Disposition"
DestArray(DestRow, 2) = "Qualified"
Case "HR Screen - Move Forward"
DestArray(DestRow, 2) = "Qualified"
Case "HR Screen - Phone Screened"
DestArray(DestRow, 2) = "Qualified"
Case "HR Screen - To be Reviewed"
DestArray(DestRow, 2) = "Qualified"
Case "HR Screen - TSC Phone Screen Invite"
DestArray(DestRow, 2) = "Qualified"
Case "HR Screen - Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Disposition"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Phone Screened"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - To be Reviewed"
DestArray(DestRow, 2) = "Qualified"
Case "Screen - Move Forward"
DestArray(DestRow, 2) = "Qualified"
Case "1st Interview - Assessment"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview - Disposition"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview - Interview Completed - Under Consideration"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview - Move Forward"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview - Scheduled"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview - To Be Scheduled"
DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview - Withdrawn"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview - Disposition"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview - Scheduled"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview - To Be Scheduled"
DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview - Withdrawn"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer - Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Approval in Progress"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Approved"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Canceled"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Disposition"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Draft"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - In Negotiation"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Offer to be made"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Refused"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Withdrawn"
DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Verification - Background Check Engaged"
DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Verification - Disposition"
DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Verification - Pre-Hire Action(s) Passed"
DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Verification - Prehire Verification - ALL EXTERNAL HIRES "
DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Verification - To be Initiated"
DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Verification - Withdrawn"
DestArray(DestRow, 2) = "Offer Made"
Case "Hire - External Hire - Confirm Employment - Move to External New Hire Processing"
DestArray(DestRow, 2) = "Hired"
Case "Hire - Internal Hire - Confirm Employment - Move to Internal Transfer Processing"
DestArray(DestRow, 2) = "Hired"
Case "Hire - Rehire - Confirm Employment - Move to Rehire Transfer Processing"
DestArray(DestRow, 2) = "Hired"
Case "Hire - Withdrawn"
DestArray(DestRow, 2) = "Hired"
Case "Screen - Withdrawn"
DestArray(DestRow, 2) = "Qualified"
Case "2nd Interview - Move Forward"
DestArray(DestRow, 2) = "Interviewed"
Case "Pre-Hire Verification - Prehire Verification INTERNAL HIRES "
DestArray(DestRow, 2) = "Offer Made"
Case "Hire - Hire Process Completed"
DestArray(DestRow, 2) = "Hired"
Case "Offer - Reneged"
DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Verification - Prehire Verification - LOCAL - EXTERNAL HIRES "
DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Verification - Prehire Verification - REMOTE- EXTERNAL HIRES"
DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Verification - Prehire Verification - SE MI - EXTERNAL HIRES "
DestArray(DestRow, 2) = "Offer Made"
Case "Hire - Reinstate - Confirm Employment - Move to Reinstate Hire Processing"
DestArray(DestRow, 2) = "Hired"
Case "Pre-Hire Verification - Prehire Verification - EXTERNAL HIRES "
DestArray(DestRow, 2) = "Offer Made"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Visa()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

ActiveSheet.Copy

Range("1:3").Delete

Dim LastRow
LastRow = Range("C65536").End(xlUp).Row

Range("D2:D" & LastRow).Hyperlinks.Delete

Range("A:A").Delete
Range("E:G").Delete
Range("F:U").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H" & LastRow).Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Background Check"
DestArray(DestRow, 2) = "Apply Completed"
Case "New"
DestArray(DestRow, 2) = "Apply Completed"
Case "HM Review"
DestArray(DestRow, 2) = "Qualified"
Case "Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
DestArray(DestRow, 2) = "Offer Made"
Case "Hire"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub AlliedBarton()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:C").Delete
Range("D:D").Delete
Range("E:E").Delete
Range("F:J").Delete

Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A300000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Application Received"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Company Standards"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Job Requirements"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected - Personal Letter"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Considered/Reviewed - Position Filled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Phone Screened - Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Position Filled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Pre-Screen/Phone Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reassigned/Transfer"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recommend for Another Job"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Scheduled for SOBC"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Scheduled for MSO1"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Send Application"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Unable To Contact"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Under Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Waiting for Guard Card / Clearance"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NY-NJ Region Pilot"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Ineligible - Job Related Conviction"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Conduct Background Check"
 DestArray(DestRow, 2) = "Qualified"
Case "Passed SOBC"
 DestArray(DestRow, 2) = "Qualified"
Case "Passed MSO1"
 DestArray(DestRow, 2) = "Qualified"
Case "Sent to Hiring Manager"
 DestArray(DestRow, 2) = "Qualified"
Case "Create Criminal Scorecard"
 DestArray(DestRow, 2) = "Qualified"
Case "Send Criminal History Form"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed and Not Selected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Job Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Job Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Job Offer Rescinded - falsified application"
 DestArray(DestRow, 2) = "Offer Made"
Case "Job Offer Rescinded - failed background"
 DestArray(DestRow, 2) = "Offer Made"
Case "Job Offer Rescinded - failed crim check"
 DestArray(DestRow, 2) = "Offer Made"
Case "Job Offer Rescinded - failed drug test"
 DestArray(DestRow, 2) = "Offer Made"
Case "Job Offer Rescinded - failed reference/other"
 DestArray(DestRow, 2) = "Offer Made"
Case "Job Offer Rescinded - failed SOBC"
 DestArray(DestRow, 2) = "Offer Made"
Case "Job Offer Rescinded - failed MSO1"
 DestArray(DestRow, 2) = "Offer Made"
Case "Job Offer Rescinded - Reid Test Not Recommend"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hired For Another Job"
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow
Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub PepsiBottling()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:B").Delete
Range("B:C").Delete
Range("C:C").Delete
Range("F:M").Delete

Range("C:C").Cut Destination:=Range("H:H")
Range("D:D").Cut Destination:=Range("C:C")
Range("B:B").Cut Destination:=Range("D:D")

Dim LastRow
LastRow = Range("A500000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 500000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "AB - Complete"
DestArray(DestRow, 2) = "Apply Completed"
Case "AB - In Process"
DestArray(DestRow, 2) = "Apply Completed"
Case "AB - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Apply Completed"
Case "AB - Inactive - Deselect"
DestArray(DestRow, 2) = "Apply Completed"
Case "AB - Inactive - Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "AB - Incomplete"
DestArray(DestRow, 2) = "Apply Completed"
Case "AB - Incomplete (Client)"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Applicant Pool"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Completed - WOTC"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Hold - Normal"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Hold - OFCCP"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Hold - Previous Employee"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Inactive - Deselect"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Inactive - Ineligible Re-Hire"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Inactive - Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Inactive - Sourcing Req Closed"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Inactive - Sourcing Req Closed - Candidate Emailed"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Incomplete - Normal"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Incomplete - WOTC"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Not Qualified - Basic Quals"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Not Qualified - Normal"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "AP - Pending"
DestArray(DestRow, 2) = "Apply Completed"
Case "BG - Completed"
DestArray(DestRow, 2) = "Apply Completed"
Case "BG - In Process"
DestArray(DestRow, 2) = "Apply Completed"
Case "BG - Results Available"
DestArray(DestRow, 2) = "Apply Completed"
Case "CO - Applicant Pool"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Inactive - Deselect"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Inactive - Not Selected"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - Background"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - DOT Block"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - Driving"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - Drug Screen"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - Ergo"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - Misrepresentation"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - Multiple"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - No Show"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - Other"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Not Qualified - Refusal to Retest"
DestArray(DestRow, 2) = "Offer Made"
Case "CO - Waiting Results"
DestArray(DestRow, 2) = "Offer Made"
Case "HI - Hired"
DestArray(DestRow, 2) = "Hired"
Case "HI - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Offer Made"
Case "HI - Inactive - Deselect"
DestArray(DestRow, 2) = "Offer Made"
Case "HI - Inactive - Location Deselect"
DestArray(DestRow, 2) = "Offer Made"
Case "HI - Inactive - No Show"
DestArray(DestRow, 2) = "Offer Made"
Case "HI - Inactive - Terminated"
DestArray(DestRow, 2) = "Hired"
Case "HI - Need Start Date"
DestArray(DestRow, 2) = "Offer Made"
Case "HM - Applicant Pool - A"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Applicant Pool - Normal"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Do Not Interview"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Eligible to Schedule"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Inactive - Candidate Not Interested"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Inactive - Deselect"
DestArray(DestRow, 2) = "Interviewed"
Case "HM - Inactive - No Show"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Inactive - Not Selected"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Inactive - Not Selected - Req Filled/Closed/Cancelled"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Interview Scheduled"
DestArray(DestRow, 2) = "Interviewed"
Case "HM - Interview Scheduled - Normal"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Interview Scheduled - Req Filled/Closed/Cancelled"
DestArray(DestRow, 2) = "Interviewed"
Case "HM - Not Qualified"
DestArray(DestRow, 2) = "Interviewed"
Case "HM - Slate"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Slate Not Qualified"
DestArray(DestRow, 2) = "Apply Completed"
Case "OF - Applicant Pool"
DestArray(DestRow, 2) = "Interviewed"
Case "OF - Extended"
DestArray(DestRow, 2) = "Interviewed"
Case "OF - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Interviewed"
Case "OF - Inactive - Benefits"
DestArray(DestRow, 2) = "Offer Made"
Case "OF - Inactive - Found Another Job"
DestArray(DestRow, 2) = "Apply Completed"
Case "OF - Inactive - No Longer Interested in Job"
DestArray(DestRow, 2) = "Offer Made"
Case "OF - Inactive - No Response"
DestArray(DestRow, 2) = "Offer Made"
Case "OF - Inactive - Not Selected"
DestArray(DestRow, 2) = "Interviewed"
Case "OF - Inactive - Other"
DestArray(DestRow, 2) = "Offer Made"
Case "OF - Inactive - Pay"
DestArray(DestRow, 2) = "Offer Made"
Case "OF - Inactive - Req Cancelled"
DestArray(DestRow, 2) = "Offer Made"
Case "OF - Inactive - Req Filled"
DestArray(DestRow, 2) = "Offer Made"
Case "OF - Not Selected"
DestArray(DestRow, 2) = "Interviewed"
Case "OF - Pending"
DestArray(DestRow, 2) = "Interviewed"
Case "OF - Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "PI - Applicant Pool"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Eligible to Schedule"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - Deselect"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - Found other employment"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - No longer interested"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - No Show"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - Other"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - Pay"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - Sourcing Req Closed"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Not Qualified"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Scheduled"
DestArray(DestRow, 2) = "Apply Completed"
Case "PI - Inactive - Hold"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Applicant Pool"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Eligible to Schedule"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - Active on other Req - Scheduled"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - Deselect"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - Found Other Employment"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - No Longer Interested"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - No Show"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - Other"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - Sourcing Req Closed"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Inactive - Unable to Contact"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Not Qualified"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Scheduled"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Scheduled - Normal"
DestArray(DestRow, 2) = "Apply Completed"
Case "RS - Applicant Pool"
DestArray(DestRow, 2) = "Apply Completed"
Case "RS - Inactive - Accepted Another Position"
DestArray(DestRow, 2) = "Apply Completed"
Case "RS - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Apply Completed"
Case "RS - Inactive - Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "RS - Inactive - Other"
DestArray(DestRow, 2) = "Apply Completed"
Case "RS - Not Qualified"
DestArray(DestRow, 2) = "Apply Completed"
Case "RS - Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "TS - In Process"
DestArray(DestRow, 2) = "Apply Completed"
Case "TS - Inactive"
DestArray(DestRow, 2) = "Qualified"
Case "TS - Incomplete"
DestArray(DestRow, 2) = "Apply Completed"
Case "TS - Not Qualified"
DestArray(DestRow, 2) = "Qualified"
Case "TS - Qualified"
DestArray(DestRow, 2) = "Interviewed"
Case "HM - Applicant Pool - EEO"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Inactive - Active on other req - Scheduled"
DestArray(DestRow, 2) = "Qualified"
Case "HM - Proceed to Interview"
DestArray(DestRow, 2) = "Qualified"
Case "BG - Inactive - Active on other Req"
DestArray(DestRow, 2) = "Apply Completed"
Case "BG - Inactive - Not Selected"
DestArray(DestRow, 2) = "Apply Completed"
Case "CO - Not Qualified - Ergo Physical Fail - No Retest"
DestArray(DestRow, 2) = "Offer Made"
Case "PS - Applicant Pool - Tier 1"
DestArray(DestRow, 2) = "Apply Completed"
Case "PS - Applicant Pool - Tier 2"
DestArray(DestRow, 2) = "Apply Completed"
Case "HM - Schedule is Full"
DestArray(DestRow, 2) = "Qualified"
Case "AP - Inactive - Candidate Opt Out"
DestArray(DestRow, 2) = "Apply Completed"
Case "HM - Inactive - Not Qualified-Inappropriate Behavior"
DestArray(DestRow, 2) = "Apply Completed"
Case "CO - Not Qualified - DOT Physical"
DestArray(DestRow, 2) = "Offer Made"
Case "BG - Applicant Pool"
DestArray(DestRow, 2) = "Offer Made"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Ecolab()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("A300000").End(xlUp).Row

Range("D:E").Delete
Range("G:L").Delete

Range("D:D").Cut Destination:=Range("H:H")
Range("A:A").Cut Destination:=Range("D:D")
Range("B:B").Cut Destination:=Range("A:A")
Range("C:C").Cut Destination:=Range("G:G")
Range("E:E").Cut Destination:=Range("C:C")
Range("F:F").Cut Destination:=Range("E:E")

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 300000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "0-Filed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "0-Filed - China"
 DestArray(DestRow, 2) = "Apply Completed"
Case "0-Filed - EMEA"
 DestArray(DestRow, 2) = "Apply Completed"
Case "0-Filed - Mexico"
 DestArray(DestRow, 2) = "Apply Completed"
Case "0-Filed - US/CA"
 DestArray(DestRow, 2) = "Apply Completed"
Case "0-Filed (Brazil)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "0-Filed (Pacific)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "0-Filed Contractor - EcoSure"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Assessment/Evaluation (Brazil)"
 DestArray(DestRow, 2) = "Qualified"
Case "Basic Qualifications Met - No Contact"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Basic Qualifications Not Met - No Contact"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Campus 1st Round Interview (On-Campus)"
 DestArray(DestRow, 2) = "Qualified"
Case "Campus 2nd Round Interview"
 DestArray(DestRow, 2) = "Interview"
Case "Campus Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Candidate Package Sent"
 DestArray(DestRow, 2) = "Qualified"
Case "Contacted"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Contacted (Inactive)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Contacted (Optional) - EcoSure"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Ecolab Not Interested - Mexico"
 DestArray(DestRow, 2) = "Qualified"
Case "EMEA - 2nd Interview"
 DestArray(DestRow, 2) = "Interview"
Case "EMEA - Assessment"
 DestArray(DestRow, 2) = "Qualified"
Case "EMEA - Candidate Not Interested"
 DestArray(DestRow, 2) = "Interview"
Case "EMEA - Forwarding to Hiring Manager"
 DestArray(DestRow, 2) = "Interview"
Case "EMEA - Hired"
 DestArray(DestRow, 2) = "Hired"
Case "EMEA - Hiring Manager / HR Interview"
 DestArray(DestRow, 2) = "Interview"
Case "EMEA - HR Telephone Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "EMEA - Not Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA - Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "EMEA - Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA - Qualified (Inactive)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA - Refusal Sent"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA - Verbal Offer to Candidate"
 DestArray(DestRow, 2) = "Offer Made"
Case "EMEA - Written Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Employment Offer (Brazil)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Field Cand Pack Sent"
 DestArray(DestRow, 2) = "Interview"
Case "Field Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Field Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Forward to Manager (Brazil)"
 DestArray(DestRow, 2) = "Interview"
Case "G - Assessments/Testing/Reference Checking"
 DestArray(DestRow, 2) = "Interviewed"
Case "G - Basic (Minimum) Qualifications Met - No Contact"
 DestArray(DestRow, 2) = "Apply Completed"
Case "G - Basic/Minimum Qualifications Not Met - No Contact"
 DestArray(DestRow, 2) = "Apply Completed"
Case "G - Candidate Not Interested"
 DestArray(DestRow, 2) = "Apply Completed"
Case "G - Contacted"
 DestArray(DestRow, 2) = "Apply Completed"
Case "G - Ecolab Not Interested"
 DestArray(DestRow, 2) = "Qualified"
Case "G - Hired"
 DestArray(DestRow, 2) = "Hired"
Case "G - Hired to Another Req"
 DestArray(DestRow, 2) = "Apply Completed"
Case "G - Interview(s)/Field Travel"
 DestArray(DestRow, 2) = "Interviewed"
Case "G - New"
 DestArray(DestRow, 2) = "Apply Completed"
Case "G - Not Reviewed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "G - Offer Accepted/Written Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "G - Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "G - Offer Extended/Verbal Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "G - Offer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "G - Pre-Employment Screening"
 DestArray(DestRow, 2) = "Offer Made"
Case "G - Pre-Offer Screening"
 DestArray(DestRow, 2) = "Interviewed"
Case "G - Qualified - Forward (Send) to Hiring Manager"
 DestArray(DestRow, 2) = "Qualified"
Case "G - Recruiter Preliminary Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "G - Reviewed Basic Qualifications"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hired - China"
 DestArray(DestRow, 2) = "Hired"
Case "Hired - EcoSure"
 DestArray(DestRow, 2) = "Hired"
Case "Hired (Brazil)"
 DestArray(DestRow, 2) = "Hired"
Case "Hired (India)"
 DestArray(DestRow, 2) = "Hired"
Case "Hired (Pacific)"
 DestArray(DestRow, 2) = "Hired"
Case "Interview"
 DestArray(DestRow, 2) = "Interview"
Case "Interview - (Pacific)"
 DestArray(DestRow, 2) = "Interview"
Case "Manager Interview - China"
 DestArray(DestRow, 2) = "Interview"
Case "Manager Interview (Brazil)"
 DestArray(DestRow, 2) = "Interview"
Case "No Interest - Candidate (Brazil)"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest - Ecolab (Brazil)"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest Candidate"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest Candidate - EcoSure"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest Candidate (Pacific)"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest Ecolab"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest Ecolab - EcoSure"
 DestArray(DestRow, 2) = "Qualified"
Case "No Interest Ecolab (Pacific)"
 DestArray(DestRow, 2) = "Qualified"
Case "Not Interested - China"
 DestArray(DestRow, 2) = "Qualified"
Case "Not Qualified (Brazil)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer Approval Process - China"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - China"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined (Brazil)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined (Pacific)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended (Pacific)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded (Brazil)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Other Position - China"
 DestArray(DestRow, 2) = "Interview"
Case "Passed to Personnel Department - Documents and Exams (Brazil)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone Screen - EcoSure"
 DestArray(DestRow, 2) = "Qualified"
Case "Preliminary Interview (Brazil)"
 DestArray(DestRow, 2) = "Qualified"
Case "Qualified"
 DestArray(DestRow, 2) = "Interview"
Case "Qualified - No Interest (Brazil)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Qualified (Brazil)"
 DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Interview - China"
 DestArray(DestRow, 2) = "Interview"
Case "Reference Checks (Pacific)"
 DestArray(DestRow, 2) = "Interview"
Case "Reviewed Basic Qualifications"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reviewed Basic Qualifications (Inactive)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Screen (Pacific)"
 DestArray(DestRow, 2) = "Qualified"
Case "Tests/Field Trip (Pacific)"
 DestArray(DestRow, 2) = "Interview"
Case "Unacceptable Assessment (Brazil)"
 DestArray(DestRow, 2) = "Qualified"
Case "Verbal Offer - China"
 DestArray(DestRow, 2) = "Offer Made"
Case "Written Offer - China"
 DestArray(DestRow, 2) = "Offer Made"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Select
Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Golder()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("A:A").Delete
Range("B:C").Delete
Range("C:C").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:B1").Value = "A"
Range("C1").Value = "Apply Completed"
Range("D1").Value = "Qualified"
Range("E1").Value = "Qualified"
Range("F1").Value = "Interviewed"
Range("G1").Value = "Interviewed"
Range("H1").Value = "Offer Made"
Range("I1").Value = "Hired"
Range("J1").Value = "Hired"

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 3

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:J" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 12)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)

For CurRow = 3 To LastRow
                   
        For CurCol = 3 To 10
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 12) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            End If
        Next CurCol
               
Next CurRow


Sheets(1).Range("1:1").Delete

Sheets(1).Range("A1:L" & DestRow).Value = DestArray

LastRow = Range("A200000").End(xlUp).Row

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("M:M").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("M:M").Cut Destination:=Range("C:C")
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("A:A")
Range("B:B").Cut Destination:=Range("E:E")
Range("B:B").Delete
Range("F:Z").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("H:O").Delete

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("AQ2:AQ" & DestRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"
Range("H:H").Delete

Range("A1:G" & LastRow).Borders.Weight = xlThin
Range("A1:G" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:G" & LastRow).Font.Size = 10
Range("A1:G" & LastRow).Font.Name = "Arial"
Range("A1:G1").Font.Color = vbBlack
Range("A1:G1").Font.Bold = True
Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub UTMB()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("C65536").End(xlUp).Row

Range("B:E").Delete
Range("D:D").Delete
Range("E:F").Delete
Range("F:K").Delete

Range("B:B").Cut Destination:=Range("H:H")
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("D:D")
Range("F:F").Delete
Range("B1").Value = "Status"

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "005 Draft"
DestArray(DestRow, 2) = "ATS Capture"
Case "015 Linked"
DestArray(DestRow, 2) = "Apply Completed"
Case "020 Applied"
DestArray(DestRow, 2) = "Apply Completed"
Case "050 Route"
DestArray(DestRow, 2) = "Qualified"
Case "071 Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "080 Ready to Hire"
DestArray(DestRow, 2) = "Offer Made"
Case "090 Hired"
DestArray(DestRow, 2) = "Hired"
Case "100 Hold"
DestArray(DestRow, 2) = "ATS Capture"
Case "110 Reject"
DestArray(DestRow, 2) = "Apply Completed"
Case "120 Withdrawn"
DestArray(DestRow, 2) = "ATS Capture"
Case "077 Preliminary Offer Rejected"
DestArray(DestRow, 2) = "Offer Made"
Case "070 Offer"
DestArray(DestRow, 2) = "Offer Made"
Case "075 Preliminary Offer Notified"
DestArray(DestRow, 2) = "Offer Made"
Case "030 Screen"
DestArray(DestRow, 2) = "Apply Completed"
Case "060 Interview "
DestArray(DestRow, 2) = "Interviewed"
Case "069 Preliminary Offer Decided"
DestArray(DestRow, 2) = "Offer Made"
Case "076 Preliminary Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "077 Offer Rejected"
DestArray(DestRow, 2) = "Offer Made"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub FirstCitizensBank()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("C65536").End(xlUp).Row

Range("G:I").Delete

Range("A:A").Cut Destination:=Range("H:H")
Range("A:A").Delete
Range("B:B").Cut Destination:=Range("H:H")

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Candidate Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Inbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Reviewed - DMT"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Suitable"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Referral Inbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Screening"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Not Selected After Screening"
 DestArray(DestRow, 2) = "Qualified"
Case "Not Selected For Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Schedule Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "RPO Screening"
 DestArray(DestRow, 2) = "Qualified"
Case "Red"
 DestArray(DestRow, 2) = "Qualified"
Case "Yellow"
 DestArray(DestRow, 2) = "Qualified"
Case "Green"
 DestArray(DestRow, 2) = "Qualified"
Case "Create Offer"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Not Selected After Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment Process"
 DestArray(DestRow, 2) = "Hired"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub ServiceMaster()

Application.ScreenUpdating = False

Dim wb As Workbook
Dim ws As Worksheet
Dim ws1 As Worksheet
Dim rng As Range
Dim UsedCol As Integer

Set wb = ActiveWorkbook

For Each ws In wb.Worksheets
ws.Range("1:3").Delete
Next ws

Set ws1 = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
ws1.Name = "Sheet1"

Set ws = wb.Worksheets(1)
UsedCol = ws.Cells(1, 255).End(xlToLeft).Column

With ws1.Cells(1, 1).Resize(1, UsedCol)
.Value = ws.Cells(1, 1).Resize(1, UsedCol).Value
End With

Application.DisplayAlerts = False

For Each ws In wb.Worksheets
If ws.Index = wb.Worksheets.Count Then
Exit For
End If
Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(65536, 1).End(xlUp).Resize(, UsedCol))
ws1.Cells(65536, 1).End(xlUp).Offset(1).Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
ws.Delete
Next ws

Application.DisplayAlerts = True

ws1.Columns.AutoFit

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F:F").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("C65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Applying"
DestArray(DestRow, 2) = "Apply Completed"
Case "Available"
DestArray(DestRow, 2) = "Apply Completed"
Case "Contact Attempted"
DestArray(DestRow, 2) = "Apply Completed"
Case "Not Hired"
DestArray(DestRow, 2) = "Apply Completed"
Case "Phone Screen"
DestArray(DestRow, 2) = "Apply Completed"
Case "Screened Out"
DestArray(DestRow, 2) = "Apply Completed"
Case "Archived"
DestArray(DestRow, 2) = "Qualified"
Case "Assessment"
DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager Review"
DestArray(DestRow, 2) = "Qualified"
Case "Interview To Be Scheduled"
DestArray(DestRow, 2) = "Qualified"
Case "Qualified"
DestArray(DestRow, 2) = "Qualified"
Case "Skills Test (fee)"
DestArray(DestRow, 2) = "Qualified"
Case "Skills Test First Sitting (fee)"
DestArray(DestRow, 2) = "Qualified"
Case "Testing"
DestArray(DestRow, 2) = "Qualified"
Case "Skills Test Second Sitting"
DestArray(DestRow, 2) = "Qualified"
Case "Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Pending Hire"
DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = "=DATE(MID(C2,7,4),LEFT(C2,2),MID(C2,4,2))"
Range("I2:I" & LastRow).Select
Selection.Copy
Range("I2:I" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If

Application.ScreenUpdating = True

End Sub

Sub Groupon()
Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationManual

Range("1:4").Delete

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Dim CurRow1
CurRow1 = 1

Do While CurRow1 < LastRow
If Left(Range("A" & CurRow1).Value, 7) = "Status:" Or Left(Range("A" & CurRow1).Value, 5) = "Email" Then
Range(CurRow1 & ":" & CurRow1).Delete
Else
CurRow1 = CurRow1 + 1
End If
Loop

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1").Value = "Email"
Range("B1").Value = "Requisition ID"
Range("C1").Value = "Title"
Range("D1").Value = "Status"
Range("E1").Value = "Submitted"
Range("F1").Value = "Date Review / Reach Out"
Range("G1").Value = "Date Screened P1"
Range("H1").Value = "Date Screened P2"
Range("I1").Value = "Date Screened P3"
Range("J1").Value = "Date Contacted - Passive"
Range("K1").Value = "Date Submitted to Hiring Manager"
Range("L1").Value = "Date Phone Screen - Phase 1"
Range("M1").Value = "Date Phone Screen - Phase 2"
Range("N1").Value = "Date Deal Creation"
Range("O1").Value = "Date Schedule Design Screen"
Range("P1").Value = "Date Designer Screen"
Range("Q1").Value = "Date Schedule Design Screen 2"
Range("R1").Value = "Date Designer Screen 2"
Range("S1").Value = "Date Phone Screen"
Range("T1").Value = "Date Hiring Manager Phone Screen"
Range("U1").Value = "Date Additional Phone Screen"
Range("V1").Value = "Date Skype Interview"
Range("W1").Value = "Date Schedule Interview"
Range("X1").Value = "Date Interview"
Range("Y1").Value = "Date 2nd Interview"
Range("Z1").Value = "Date Offer Sent"
Range("AA1").Value = "Offer Accepted Date"
Range("AB1").Value = "Date Rejected"
Range("AC1").Value = "Date Candidate Withdrew"
Range("AD1").Value = "Date Considered; Not Available"
Range("AE1").Value = "Date Pass"
Range("AF1").Value = "Date Maybe Later"
Range("AG1").Value = "Date Offer Rejected"

LastRow = Range("A200000").End(xlUp).Row

Range("D:D").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:C1").Value = "A"
Range("D1").Value = "Apply Completed"
Range("E1").Value = "Qualified"
Range("F1").Value = "Qualified"
Range("G1").Value = "Qualified"
Range("H1").Value = "Qualified"
Range("I1").Value = "ATS Capture"
Range("J1").Value = "Qualified"
Range("K1").Value = "Interviewed"
Range("L1").Value = "Interviewed"
Range("M1").Value = "Interviewed"
Range("N1").Value = "Interviewed"
Range("O1").Value = "Interviewed"
Range("P1").Value = "Interviewed"
Range("Q1").Value = "Interviewed"
Range("R1").Value = "Interviewed"
Range("S1").Value = "Interviewed"
Range("T1").Value = "Interviewed"
Range("U1").Value = "Interviewed"
Range("V1").Value = "Interviewed"
Range("W1").Value = "Interviewed"
Range("X1").Value = "Interviewed"
Range("Y1").Value = "Offer Made"
Range("Z1").Value = "Hired"
Range("AA1").Value = "Apply Completed"
Range("AB1").Value = "Apply Completed"
Range("AC1").Value = "Apply Completed"
Range("AD1").Value = "Apply Completed"
Range("AE1").Value = "Apply Completed"
Range("AF1").Value = "Offer Made"

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 4

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:AF" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 400000, 1 To 35)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)
    DestArray(1, 13) = SourceArray(2, 13)
    DestArray(1, 14) = SourceArray(2, 14)
    DestArray(1, 15) = SourceArray(2, 15)
    DestArray(1, 16) = SourceArray(2, 16)
    DestArray(1, 17) = SourceArray(2, 17)
    DestArray(1, 18) = SourceArray(2, 18)
    DestArray(1, 19) = SourceArray(2, 19)
    DestArray(1, 20) = SourceArray(2, 20)
    DestArray(1, 21) = SourceArray(2, 21)
    DestArray(1, 22) = SourceArray(2, 22)
    DestArray(1, 23) = SourceArray(2, 23)
    DestArray(1, 24) = SourceArray(2, 24)
    DestArray(1, 25) = SourceArray(2, 25)
    DestArray(1, 26) = SourceArray(2, 26)
    DestArray(1, 27) = SourceArray(2, 27)
    DestArray(1, 28) = SourceArray(2, 28)
    DestArray(1, 29) = SourceArray(2, 29)
    DestArray(1, 30) = SourceArray(2, 30)
    DestArray(1, 31) = SourceArray(2, 31)
    DestArray(1, 32) = SourceArray(2, 32)

For CurRow = 3 To LastRow
                   
        For CurCol = 4 To 31
            If SourceArray(CurRow, CurCol) <> "" Then
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, 20)
                DestArray(DestRow, 21) = SourceArray(CurRow, 21)
                DestArray(DestRow, 22) = SourceArray(CurRow, 22)
                DestArray(DestRow, 23) = SourceArray(CurRow, 23)
                DestArray(DestRow, 24) = SourceArray(CurRow, 24)
                DestArray(DestRow, 25) = SourceArray(CurRow, 25)
                DestArray(DestRow, 26) = SourceArray(CurRow, 26)
                DestArray(DestRow, 27) = SourceArray(CurRow, 27)
                DestArray(DestRow, 28) = SourceArray(CurRow, 28)
                DestArray(DestRow, 29) = SourceArray(CurRow, 29)
                DestArray(DestRow, 30) = SourceArray(CurRow, 30)
                DestArray(DestRow, 31) = SourceArray(CurRow, 31)
                DestArray(DestRow, 32) = SourceArray(CurRow, 31)
                DestArray(DestRow, 33) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 34) = SourceArray(1, CurCol)
                DestArray(DestRow, 35) = SourceArray(2, CurCol)
                
                DestRow = DestRow + 1
                        
            Else
            
            End If
        Next CurCol
               
Next CurRow


Sheets(1).Range("1:1").Delete

Sheets(1).Range("A1:AI" & DestRow).Value = DestArray

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("AI:AI").Cut Destination:=Range("B:B")
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("AI:AI").Cut Destination:=Range("C:C")
Range("F:AH").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("I2:I" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & DestRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Sheets(1).Range("A1:H" & DestRow).Font.Size = 10
Sheets(1).Range("A1:H" & DestRow).Font.Name = "Arial"
Sheets(1).Range("A1:H1").Font.Color = vbBlack
Sheets(1).Range("A1:H1").Font.Bold = True
Sheets(1).Range("A1:H1").Interior.Color = vbYellow

Range("A1:H" & DestRow).Borders.Weight = xlThin
Range("A1:H" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Sheets(1).Range(DestRow & ":" & DestRow).Delete

Range("I2:I" & DestRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & DestRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & DestRow).Sort Key1:=Range("C2:C" & DestRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub McGladrey()

Application.ScreenUpdating = False

Dim wb As Workbook
Dim ws As Worksheet
Dim ws1 As Worksheet
Dim rng As Range
Dim UsedCol As Integer

Set wb = ActiveWorkbook

For Each ws In wb.Worksheets
ws.Range("S:S").Value = ws.Name
Next ws

Set ws1 = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
ws1.Name = "Sheet1"

Set ws = wb.Worksheets(1)
UsedCol = ws.Cells(1, 255).End(xlToLeft).Column

With ws1.Cells(1, 1).Resize(1, UsedCol)
.Value = ws.Cells(1, 1).Resize(1, UsedCol).Value
End With

Application.DisplayAlerts = False

For Each ws In wb.Worksheets
If ws.Index = wb.Worksheets.Count Then
Exit For
End If
Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(65536, 1).End(xlUp).Resize(, UsedCol))
ws1.Cells(65536, 1).End(xlUp).Offset(1).Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
ws.Delete
Next ws

Application.DisplayAlerts = True

ws1.Columns.AutoFit

Range("S1").Value = "Worksheet Name"
Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("A:A")
Range("E:E").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("I:I").Cut Destination:=Range("C:C")
Range("I:I").Delete
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("M:M").Cut Destination:=Range("D:D")
Range("M:M").Delete
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("Q:Q").Cut Destination:=Range("E:E")
Range("Q:Q").Delete

Dim LastRow
LastRow = Range("C65536").End(xlUp).Row

Range("U2:U" & LastRow).Formula = "=Left(T2,FIND("" "",T2,1)-1)"

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:U" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 21)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    DestArray(1, 9) = SourceArray(1, 9)
    DestArray(1, 10) = SourceArray(1, 10)
    DestArray(1, 11) = SourceArray(1, 11)
    DestArray(1, 12) = SourceArray(1, 12)
    DestArray(1, 13) = SourceArray(1, 13)
    DestArray(1, 14) = SourceArray(1, 14)
    DestArray(1, 15) = SourceArray(1, 15)
    DestArray(1, 16) = SourceArray(1, 16)
    DestArray(1, 17) = SourceArray(1, 17)
    DestArray(1, 18) = SourceArray(1, 18)
    DestArray(1, 19) = SourceArray(1, 19)
    DestArray(1, 20) = SourceArray(1, 20)
    DestArray(1, 21) = SourceArray(1, 21)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, 20)
                DestArray(DestRow, 21) = SourceArray(CurRow, 21)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 21)

Select Case OriginalStatus
Case "Submittals"
DestArray(DestRow, 2) = "Apply Completed"
Case "Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
DestArray(DestRow, 2) = "Offer Made"
Case "Onboarded"
DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:U" & DestRow).Value = DestArray

Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("V:W").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("V2:V" & LastRow).Formula = Range("C2:C" & DestRow).Value2
Range("V2:V" & LastRow).Select
Selection.Copy
Range("V2:V" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("V2:V" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("V2:V" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("V2:V" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:U" & LastRow).Borders.Weight = xlThin
Range("A1:U" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:U" & LastRow).Font.Size = 10
Range("A1:U" & LastRow).Font.Name = "Arial"
Range("A1:U1").Font.Color = vbBlack
Range("A1:U1").Font.Bold = True
Range("A1:U1").Interior.Color = vbYellow

Range("V2:V" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("V1").Formula = "=SUM(V2:V" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("V1").Value

Range("V:V").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:U" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If

Application.ScreenUpdating = True

End Sub
Sub Bombardier()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("O:O").Cut Destination:=Range("A:A")
Range("O:O").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("S:S").Cut Destination:=Range("C:C")
Range("S:S").Delete
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("M:M").Cut Destination:=Range("D:D")
Range("M:M").Delete
Range("E1").Select
ActiveCell.EntireColumn.Insert
Range("N:N").Cut Destination:=Range("E:E")
Range("N:N").Delete
Range("F1").Select
ActiveCell.EntireColumn.Insert
Range("F1").Select
ActiveCell.EntireColumn.Insert

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:W" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 23)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    DestArray(1, 9) = SourceArray(1, 9)
    DestArray(1, 10) = SourceArray(1, 10)
    DestArray(1, 11) = SourceArray(1, 11)
    DestArray(1, 12) = SourceArray(1, 12)
    DestArray(1, 13) = SourceArray(1, 13)
    DestArray(1, 14) = SourceArray(1, 14)
    DestArray(1, 15) = SourceArray(1, 15)
    DestArray(1, 16) = SourceArray(1, 16)
    DestArray(1, 17) = SourceArray(1, 17)
    DestArray(1, 18) = SourceArray(1, 18)
    DestArray(1, 19) = SourceArray(1, 19)
    DestArray(1, 20) = SourceArray(1, 20)
    DestArray(1, 21) = SourceArray(1, 21)
    DestArray(1, 22) = SourceArray(1, 22)
    DestArray(1, 23) = SourceArray(1, 23)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, 20)
                DestArray(DestRow, 21) = SourceArray(CurRow, 21)
                DestArray(DestRow, 22) = SourceArray(CurRow, 22)
                DestArray(DestRow, 23) = SourceArray(CurRow, 23)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 22)

Select Case OriginalStatus
Case "BA Has Declined"
 DestArray(DestRow, 2) = "Apply Completed"
Case "BA New"
 DestArray(DestRow, 2) = "Apply Completed"
Case "BA Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "BA Rescinded"
 DestArray(DestRow, 2) = "Apply Completed"
Case "BA To be Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "BA To Be Evaluated "
 DestArray(DestRow, 2) = "Apply Completed"
Case "BA Waiting For Candidate Response"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Canceled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Draft"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Eligible"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Get on the Plane"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Has Declined"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Left a message"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Refused"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Register"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Scheduled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Short List"
 DestArray(DestRow, 2) = "Apply Completed"
Case "To be Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "To be evaluated"
 DestArray(DestRow, 2) = "Apply Completed"
Case "To be Evaluated"
 DestArray(DestRow, 2) = "Apply Completed"
Case "To be Planned"
 DestArray(DestRow, 2) = "Apply Completed"
Case "To be Reconsidered"
 DestArray(DestRow, 2) = "Apply Completed"
Case "To be Rescheduled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "To be scheduled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "To Progress"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Under consideration"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Waiting for results"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn (NLI)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "BA Reviewed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reviewed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "1st Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st Interview Planned"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Phone Interview Planned"
 DestArray(DestRow, 2) = "Interviewed"
Case "Waiting for Final Interview Results"
 DestArray(DestRow, 2) = "Interviewed"
Case "Meets Standards"
 DestArray(DestRow, 2) = "Qualified"
Case "Qualified"
 DestArray(DestRow, 2) = "Qualified"
Case "Test 1 Passed"
 DestArray(DestRow, 2) = "Qualified"
Case "Test 1 Scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "Test 5 Passed"
 DestArray(DestRow, 2) = "Qualified"
Case "Testing"
 DestArray(DestRow, 2) = "Qualified"
Case "To be phone screened"
 DestArray(DestRow, 2) = "Qualified"
Case "Waiting for Managers Feedback"
 DestArray(DestRow, 2) = "Qualified"
Case "Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "BA Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "BA Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "BA Offer Renegotiation"
 DestArray(DestRow, 2) = "Offer Made"
Case "BA Offer to be made"
 DestArray(DestRow, 2) = "Offer Made"
Case "Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer to be made"
 DestArray(DestRow, 2) = "Offer Made"
Case "Promotion"
 DestArray(DestRow, 2) = "Offer Made"
Case "Reference & Security Planned"
 DestArray(DestRow, 2) = "Offer Made"
Case "BA Hired "
 DestArray(DestRow, 2) = "Hired"
Case "Hire - Manual"
 DestArray(DestRow, 2) = "Hired"
Case "Hire - SAP"
 DestArray(DestRow, 2) = "Hired"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hired to another Req"
 DestArray(DestRow, 2) = "Hired"
Case "To be hired"
 DestArray(DestRow, 2) = "Hired"
Case "BA Waiting for Salary Analysis"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Test 1 Failed"
 DestArray(DestRow, 2) = "Qualified"
Case "Test 2 to Be Scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "To Be Made"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Waiting for Candidate Answer"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Test 2 Passed"
 DestArray(DestRow, 2) = "Qualified"
Case "Standby"
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Answer"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Tests Planned"
 DestArray(DestRow, 2) = "Apply Completed"

End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:W" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("X2:X" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("X2:X" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("X2:X" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("X2:X" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("X2:X" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:W" & LastRow).Borders.Weight = xlThin
Range("A1:W" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:W" & LastRow).Font.Size = 10
Range("A1:W" & LastRow).Font.Name = "Arial"
Range("A1:W1").Font.Color = vbBlack
Range("A1:W1").Font.Bold = True
Range("A1:W1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("X2:X" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("X1").Formula = "=SUM(X2:X" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("X1").Value

Range("X:X").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:W" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub AmericanFamilyAFI()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:2").Delete
Range("G:O").Delete
Range("H:H").Delete

Dim LastRow
LastRow = Range("B65536").End(xlUp).Row

Range("B1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("H:H").Delete
Range("F1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
Range("J:J").Cut Destination:=Range("G:G")

Range("J2:J" & LastRow).Formula = "=H2&I2"

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:J" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 10)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    DestArray(1, 9) = SourceArray(1, 9)
    DestArray(1, 10) = SourceArray(1, 10)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 10)

Select Case OriginalStatus
Case ".Submitted Int/Ext.Screening Complete"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Int/Ext.To be Reviewed/Screened by Recruiter"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Int/ExtCandidate Submitted"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Int/ExtHas Declined"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Int/ExtRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Int/ExtRR - Maybe"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Int/ExtRR - No"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Int/ExtRR - Yes"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Send to Manager Int/Ext.Send to Manager"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Send to Manager Int/ExtHas Declined"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Send to Manager Int/ExtRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Send to Manager Int/ExtReview Resume"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Send to Manager Int/Ext.Interested"
 DestArray(DestRow, 2) = "Qualified"
Case ".Send to Manager Int/Ext.Manager Reviewed"
 DestArray(DestRow, 2) = "Qualified"
Case ".InterviewHas Declined"
 DestArray(DestRow, 2) = "Qualified"
Case ".InterviewRejected"
 DestArray(DestRow, 2) = "Qualified"
Case ".Interview.Interview Successful "
 DestArray(DestRow, 2) = "Interviewed"
Case ".Interview.Schedule HR interview"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Interview.Schedule HR/Mgr Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Interview.Schedule Mgr Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Interview.Complete Mgr Eval for Int Hire"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer  Select Activity: Capture Response"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer  Select Activity: Create Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer  Select Activity: Extend Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer  Select Activity: Send Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferCandidate Reneged"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferCompany Rescind"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferDraft"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferInternal Background Check Successful"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer to be made"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check References.External Background Check"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check References.Send External Hire to PeopleSoft"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check ReferencesComplete Ext Background References"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check ReferencesExt Proceed On/Int Bypass This Step"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check References.Background Checks Successful "
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check ReferencesHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check ReferencesLaunch On boarding (day after PSFT entry)"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check ReferencesRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Hire Internal/ExternalHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Hire Internal/ExternalRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Hire Internal/External           .To be hired"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Hire Internal/External           .Hired  External"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:S" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("T2:T" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("T2:T" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("T2:T" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("T2:T" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("T2:T" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:S" & LastRow).Borders.Weight = xlThin
Range("A1:S" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:S" & LastRow).Font.Size = 10
Range("A1:S" & LastRow).Font.Name = "Arial"
Range("A1:S1").Font.Color = vbBlack
Range("A1:S1").Font.Bold = True
Range("A1:S1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("T:T").NumberFormat = "General"

Range("T2:T" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("T1").Formula = "=SUM(T2:T" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("T1").Value

Range("T:T").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub AmericanFamilyAGT()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:2").Delete
Range("G:O").Delete
Range("H:H").Delete

Dim LastRow
LastRow = Range("B65536").End(xlUp).Row

Range("B1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("H:H").Delete
Range("F1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
Range("J:J").Cut Destination:=Range("G:G")

Range("J2:J" & LastRow).Formula = "=H2&I2"

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:J" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 10)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    DestArray(1, 9) = SourceArray(1, 9)
    DestArray(1, 10) = SourceArray(1, 10)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 10)

Select Case OriginalStatus
Case ".Submitted Agent/AITCandidate Submitted"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Agent/AITHas Declined"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Agent/AITRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Submitted Agent/AITTo be reviewed/screened by Recruiter"
 DestArray(DestRow, 2) = "Apply Completed"
Case ".Screening AGT/AITRR - Yes"
 DestArray(DestRow, 2) = "Qualified"
Case ".Screening AGT/AITRR - No"
 DestArray(DestRow, 2) = "Qualified"
Case ".Screening AGT/AITRR - Maybe"
 DestArray(DestRow, 2) = "Qualified"
Case ".Screening AGT/AIT.1st Attempt to Schedule Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case ".Screening AGT/AIT.2nd Attempt to Schedule Phone Screen (No Response)"
 DestArray(DestRow, 2) = "Qualified"
Case ".Screening AGT/AITPhone Screen Scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case ".Screening AGT/AIT.Sent Authorization Form"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AIT.Chally Complete"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AIT.Chally Initiated"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AIT.Criminal & MVR Complete"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AIT.MVR Ordered"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AIT.MVR Complete"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AIT.Running Background Checks"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AIT.Screening Complete"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AITCredit Complete Clear"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AITCredit Complete Provisional"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AITCriminal Background Complete"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AITHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case ".Screening AGT/AITRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "ASM Review AGT/AIT.Under Consideration"
 DestArray(DestRow, 2) = "Interviewed"
Case "ASM Review AGT/AIT.Working on Licenses"
 DestArray(DestRow, 2) = "Interviewed"
Case "ASM Review AGT/AITHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case "ASM Review AGT/AITRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "ASM Review AGT/AITSchedule 1st Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "ASM Review AGT/AITSchedule 2nd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "ASM Review AGT/AITSchedule 3rd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "ASM Review AGT/AITSchedule DX Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "ASM Review AGT/AITAwaiting Appointment"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Appointment Info AGT/AITNATP Nomination E-form Submitted"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Appointment Info AGT/AIT.Working on Licenses"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Appointment Info AGT/AITAwaiting Appointment"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Appointment Info AGT/AITHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Appointment Info AGT/AITRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer  Select Activity: Capture Response"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer  Select Activity: Create Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer  Select Activity: Extend Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer  Select Activity: Send Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferCandidate Reneged"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferCompany Rescind"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer to be made"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Awaiting NATPAwaiting NATP"
 DestArray(DestRow, 2) = "Offer Made"
Case "Awaiting NATPCandidate Withdrew"
 DestArray(DestRow, 2) = "Offer Made"
Case "Awaiting NATPRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check References.Background Checks Successful"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check References.External Background Check"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check ReferencesHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check ReferencesIntegration or Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check ReferencesLaunch On boarding (AIT external hire only - day after PSFT entry)"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check ReferencesRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Ext Hire Integration/Check References.Send External Hire to PeopleSoft"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Hire AGT/AITAwaiting NATP"
 DestArray(DestRow, 2) = "Offer Made"
Case ".Hire AGT/AITHired - AIT"
 DestArray(DestRow, 2) = "Hired"
Case ".Hire AGT/AITHired - Agent"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:S" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("T2:T" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("T2:T" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("T2:T" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("T2:T" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("T2:T" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:S" & LastRow).Borders.Weight = xlThin
Range("A1:S" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:S" & LastRow).Font.Size = 10
Range("A1:S" & LastRow).Font.Name = "Arial"
Range("A1:S1").Font.Color = vbBlack
Range("A1:S1").Font.Bold = True
Range("A1:S1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("T:T").NumberFormat = "General"

Range("T2:T" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("T1").Formula = "=SUM(T2:T" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("T1").Value

Range("T:T").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub GenuineParts()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("B100000").End(xlUp).Row

Range("A:A").Cut Destination:=Range("J:J")
Range("A:A").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("G:G").Delete
Range("F1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:L" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 12)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    DestArray(1, 9) = SourceArray(1, 9)
    DestArray(1, 10) = SourceArray(1, 10)
    DestArray(1, 11) = SourceArray(1, 11)
    DestArray(1, 12) = SourceArray(1, 12)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "010 Applied"
 DestArray(DestRow, 2) = "Apply Completed"
Case "100 Hold"
 DestArray(DestRow, 2) = "Apply Completed"
Case "110 Reject"
 DestArray(DestRow, 2) = "Apply Completed"
Case "112 Failed Prescreening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "115 Reject Online Screening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "120 Withdrawn"
 DestArray(DestRow, 2) = "Apply Completed"
Case "030 Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "015 Linked"
 DestArray(DestRow, 2) = "Qualified"
Case "019 Linked Questionnaire"
 DestArray(DestRow, 2) = "Qualified"
Case "020 Reviewed"
 DestArray(DestRow, 2) = "Qualified"
Case "050 Route"
 DestArray(DestRow, 2) = "Qualified"
Case "060 Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "070 Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "071 Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "080 Ready to Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "069 Preliminary Offer Decided"
 DestArray(DestRow, 2) = "Offer Made"
Case "075 Preliminary Offer Notified"
 DestArray(DestRow, 2) = "Offer Made"
Case "076 Preliminary Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "077 Preliminary Offer Rejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "090 Hired"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:L" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("M2:M" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("M2:M" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("M2:M" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("M2:M" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("M2:M" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:L" & LastRow).Borders.Weight = xlThin
Range("A1:L" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:L" & LastRow).Font.Size = 10
Range("A1:L" & LastRow).Font.Name = "Arial"
Range("A1:L1").Font.Color = vbBlack
Range("A1:L1").Font.Bold = True
Range("A1:L1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("M2:M" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("M1").Formula = "=SUM(M2:M" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("M1").Value

Range("M:M").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:L" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub UUHC()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("B100000").End(xlUp).Row

Range("C:D").Delete
Range("G:H").Delete

Range("G2:G" & LastRow).Formula = "=DATE(MID(D2,FIND(""/"",D2,3)+1,4),LEFT(D2,1),MID(D2,3,FIND(""/"",D2,3)-3))"
Range("G2:G" & LastRow).Select
Selection.Copy
Range("D2:D" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("G2:G" & LastRow).Delete
Range("D2:D" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A:A").Cut Destination:=Range("H:H")
Range("A:A").Delete
Range("B:B").Cut Destination:=Range("H:H")

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 11)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Applicants"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Applied - Failed Pre Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Applied - Further Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Applied - Need Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Applied - Not Qualified (Rejection Sent)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Applied - Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Applied - Post Slate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Applied - Pre Screen Flag"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Interviewed - Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Interviewed - Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Reviewed - Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Review -Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reviewed - Hold"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reviewed - Not Selected (Rejection Sent)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Sr. Recruiter Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background"
 DestArray(DestRow, 2) = "Qualified"
Case "Completed"
 DestArray(DestRow, 2) = "Qualified"
Case "Initiated"
 DestArray(DestRow, 2) = "Qualified"
Case "Not Qualified"
 DestArray(DestRow, 2) = "Qualified"
Case "Not Qualified (Rejection Sent)"
 DestArray(DestRow, 2) = "Qualified"
Case "Rescinded"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Interviewed"
Case "Applied - Setup Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "HireVue Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interested/Interviewing"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview Not Selected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Requested"
 DestArray(DestRow, 2) = "Interviewed"
Case "Review"
 DestArray(DestRow, 2) = "Interviewed"
Case "Setup Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Declined/Rejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Requested"
 DestArray(DestRow, 2) = "Offer Made"
Case "Salary Calc Complete"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hired - Candidate Instructions"
 DestArray(DestRow, 2) = "Hired"
Case "Hired - Manager Instructions"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:K" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("L2:L" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("L2:L" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("L2:L" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("L2:L" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("L2:L" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:K" & LastRow).Borders.Weight = xlThin
Range("A1:K" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:K" & LastRow).Font.Size = 10
Range("A1:K" & LastRow).Font.Name = "Arial"
Range("A1:K1").Font.Color = vbBlack
Range("A1:K1").Font.Bold = True
Range("A1:K1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("L2:L" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("L1").Formula = "=SUM(L2:L" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("L1").Value

Range("L:L").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:K" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Coventry()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:E").Delete
Range("C:E").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("A:A")
Range("E:E").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert

Dim LastRow
LastRow = Range("D65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "CSO Resume Submitted - Application Questions Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - Candidate Contacted"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - DDI Assessment"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - Has Declined"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - Keynomics"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - Phone Screened"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - Rejected"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - Resume Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - Screening Complete"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - To be Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Submitted - Application Questions Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Submitted - Candidate Contacted"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Submitted - Has Declined"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Submitted - Internal Eligibility Verification"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Submitted - Phone Screened"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Submitted - Rejected"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Submitted - Resume Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Submitted - Screening Complete"
DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Submitted - To be Reviewed"
DestArray(DestRow, 2) = "Apply Completed"
Case "CSO Resume Submitted - Internal Eligibility Verification"
DestArray(DestRow, 2) = "Apply Completed"
Case "Interviewing - To be Interviewed"
DestArray(DestRow, 2) = "Qualified"
Case "Interviewing - 1st Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewing - 2nd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewing - Data Integration"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewing - Has Declined"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewing - Interviewing Complete"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewing - Rejected"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer - Draft"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer - Offer to be made"
DestArray(DestRow, 2) = "Interviewed"
Case "Interviewing - 3rd Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer - Approved"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer - Canceled"
DestArray(DestRow, 2) = "Interviewed"
Case "Background Checks - Background Check"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Checks - Checks Successful"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Checks - Has Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Extended"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Has Declined"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Refused"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Rejected"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Reneged"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Rescinded"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Checks - Background Check Results Received"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Checks - Initiate Checks"
DestArray(DestRow, 2) = "Offer Made"
Case "Background Checks - OIG/GSA - Required for External Candidates"
DestArray(DestRow, 2) = "Offer Made"
Case "Hire - Confirm Employee Start"
DestArray(DestRow, 2) = "Hired"
Case "Hire - To be Hired"
DestArray(DestRow, 2) = "Hired"
Case "Hire - Hired/External"
DestArray(DestRow, 2) = "Hired"
Case "Hire - Hired/Internal"
DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub KPMGGrad()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:C").Delete
Range("B:B").Delete
Range("C:F").Delete
Range("D:D").Delete
Range("E:H").Delete

Dim LastRow
LastRow = Range("D200000").End(xlUp).Row

Range("E2:E" & LastRow).Formula = "=left(d2,3)"

Dim CurRow1
Dim Month1

CurRow1 = 2

Do While CurRow1 <= LastRow

Month1 = Range("E" & CurRow1).Value

Select Case Month1
Case "Jan"
Range("F" & CurRow1).Value = 1
Case "Feb"
Range("F" & CurRow1).Value = 2
Case "Mar"
Range("F" & CurRow1).Value = 3
Case "Apr"
Range("F" & CurRow1).Value = 4
Case "May"
Range("F" & CurRow1).Value = 5
Case "Jun"
Range("F" & CurRow1).Value = 6
Case "Jul"
Range("F" & CurRow1).Value = 7
Case "Aug"
Range("F" & CurRow1).Value = 8
Case "Sep"
Range("F" & CurRow1).Value = 9
Case "Oct"
Range("F" & CurRow1).Value = 10
Case "Nov"
Range("F" & CurRow1).Value = 11
Case "Dec"
Range("F" & CurRow1).Value = 12
End Select

CurRow1 = CurRow1 + 1
Loop

Range("G2:G" & LastRow).Formula = "=MID(D2,5,2)"
Range("H2:H" & LastRow).Formula = "=MID(D2,8,4)"
Range("I2:I" & LastRow).Formula = "=DATE(H2,F2,G2)"

Range("I2:I" & LastRow).Select
Selection.Copy
Range("I2:I" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("B:B").Select
ActiveCell.EntireColumn.Insert
Range("C:C").Select
ActiveCell.EntireColumn.Insert
Range("K:K").Cut Destination:=Range("C:C")
Range("F:K").Delete
Range("E:E").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Application withdrawn"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Assessment centre booking confirmed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Completed SJT - pending reject (contacts)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Completed SJT - recommend reject"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Final from Partner interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Final interview confirmed"
 DestArray(DestRow, 2) = "Interviewed"
Case "first interview booking confirmed"
 DestArray(DestRow, 2) = "Qualified"
Case "Hold - Invite to Assessment Centre"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hold - Invite to Final Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hold after Final interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Invite to assessment centre"
 DestArray(DestRow, 2) = "Interviewed"
Case "Invite to Final interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Invite to first interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Invite to online tests"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Invited to Situational Judgement Test"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New application (Manual screening)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer accepted"
 DestArray(DestRow, 2) = "Hired"
Case "Offer made"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer rejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "On hold after assessment centre"
 DestArray(DestRow, 2) = "Interviewed"
Case "On hold after first interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "On hold after SJT"
 DestArray(DestRow, 2) = "Apply Completed"
Case "On hold after tests"
 DestArray(DestRow, 2) = "Apply Completed"
Case "On hold Invite First Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Online tests complete"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Passed assessment centre"
 DestArray(DestRow, 2) = "Interviewed"
Case "Passed first interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Passed online tests"
 DestArray(DestRow, 2) = "Qualified"
Case "Passed SJT (check academics)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Passed SJT (contacts)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject after new application"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject after screening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject after screening (academics)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject after SJT"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject from assessment centre"
 DestArray(DestRow, 2) = "Interviewed"
Case "Reject from Final interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Reject from first interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Reject from online tests"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject From University Partner Screening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "SJT in progress"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pool"
 DestArray(DestRow, 2) = "Apply Completed "
Case "Vacation-Reject after e-tray"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Vacation-Reject after interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Verbal test complete"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Invite to numeracy test (2014)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Invite to online tests (2013)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Invite to verbal test (2014)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Online tests complete (2013)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject after numeracy test (2014)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject after verbal test (2014)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "University Partner Screening"
 DestArray(DestRow, 2) = "Apply Completed "
Case "Verbal test complete (2013)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn - SJT not completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Invite to numeracy test (2013)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Invite to verbal test (2013)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Invite to verbal test (2013 school leaver)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Numeracy Test Complete (2013)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "On hold screening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "On hold before situational Judgment test"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject after numeracy test (2013)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject after numeracy test (2013 school-leaver)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject after numeracy test (2013 school-leaver)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn – Verbal Test not completed (2014 graduate)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn – Numeracy Test not completed (2014 graduate)"
 DestArray(DestRow, 2) = "Apply Completed"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Intel()

Application.ScreenUpdating = False

Dim wb As Workbook
Dim ws As Worksheet
Dim ws1 As Worksheet
Dim rng As Range
Dim UsedCol As Integer

Set wb = ActiveWorkbook

For Each ws In wb.Worksheets
ws.Range("A:D").Delete
Next ws

Set ws1 = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
ws1.Name = "Sheet1"

Set ws = wb.Worksheets(1)
UsedCol = ws.Cells(1, 255).End(xlToLeft).Column

With ws1.Cells(1, 1).Resize(1, UsedCol)
.Value = ws.Cells(1, 1).Resize(1, UsedCol).Value
End With

Application.DisplayAlerts = False

For Each ws In wb.Worksheets
If ws.Index = wb.Worksheets.Count Then
Exit For
End If
Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(100000, 1).End(xlUp).Resize(, UsedCol))
ws1.Cells(100000, 1).End(xlUp).Offset(1).Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
ws.Delete
Next ws

Application.DisplayAlerts = True

ws1.Columns.AutoFit

Range("A:C").Delete
Range("B:G").Delete
Range("D:K").Delete
Range("F:F").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("A:A")
Range("E:E").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("D:D").Cut Destination:=Range("I:I")
Range("D:D").Delete

Dim LastRow
LastRow = Range("C100000").End(xlUp).Row

Range("I2:I" & LastRow).Formula = "=left(c2,3)"

Dim CurRow1
CurRow1 = 2

Do While CurRow1 <= LastRow

Dim Month1
Month1 = Range("I" & CurRow1)

Select Case Month1
Case "Jan"
Range("J" & CurRow1).Value = 1
Case "Feb"
Range("J" & CurRow1).Value = 2
Case "Mar"
Range("J" & CurRow1).Value = 3
Case "Apr"
Range("J" & CurRow1).Value = 4
Case "May"
Range("J" & CurRow1).Value = 5
Case "Jun"
Range("J" & CurRow1).Value = 6
Case "Jul"
Range("J" & CurRow1).Value = 7
Case "Aug"
Range("J" & CurRow1).Value = 8
Case "Sep"
Range("J" & CurRow1).Value = 9
Case "Oct"
Range("J" & CurRow1).Value = 10
Case "Nov"
Range("J" & CurRow1).Value = 11
Case "Dec"
Range("J" & CurRow1).Value = 12
End Select

CurRow1 = CurRow1 + 1

Loop

Range("K2:K" & LastRow).Formula = "=DATE(MID(C2,8,4),J2,MID(C2,5,2))"

Range("K2:K" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:K" & LastRow).Delete

CurRow1 = 2

Do While CurRow1 <= LastRow

Dim EmailString
EmailString = Range("A" & CurRow1)
EmailString = Replace(EmailString, " ", "")
Range("I" & CurRow1).Value = EmailString

CurRow1 = CurRow1 + 1

Loop

Range("I2:I" & LastRow).Select
Selection.Copy
Range("A2:A" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "-"
DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected or Withdrawn"
DestArray(DestRow, 2) = "Apply Completed"
Case "Staffing Screened"
DestArray(DestRow, 2) = "Apply Completed"
Case "Has Declined"
DestArray(DestRow, 2) = "Apply Completed"
Case "Interview Scheduled"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer Rejected By Candidate"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer On Hold"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Withdrawn By Intel"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accept Withdrawn"
DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If

Application.ScreenUpdating = True

End Sub

Sub RGF()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim LastRow
LastRow = Range("B100000").End(xlUp).Row

Range("A" & LastRow + 2 & ":A" & LastRow + 7).Delete

Range("C:K").Delete
Range("H:H").Delete
Range("I:I").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("I:I").Cut Destination:=Range("A:A")

Cells.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
Cells.Replace What:="1/0/1900", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:C1").Value = "A"
Range("D1").Value = "Apply Completed"
Range("E1").Value = "Qualified"
Range("F1").Value = "Interviewed"
Range("G1").Value = "Offer Made"
Range("H1").Value = "Hired"


LastRow = Range("B100000").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 4

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = ActiveSheet.Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 10)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)

For CurRow = 3 To LastRow
                   
        For CurCol = 4 To 8
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 10) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            End If
        Next CurCol
               
Next CurRow

ActiveSheet.Range("1:1").Delete

ActiveSheet.Range("A1:J" & DestRow).Value = DestArray

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:J").Delete
Range("G:G").Cut Destination:=Range("B:B")
Range("F:F").Cut Destination:=Range("C:C")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

LastRow = Range("B100000").End(xlUp).Row

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Range("H:H").Delete

ActiveSheet.Range("A1:G" & DestRow).Font.Size = 10
ActiveSheet.Range("A1:G" & DestRow).Font.Name = "Arial"
ActiveSheet.Range("A1:G1").Font.Color = vbBlack
ActiveSheet.Range("A1:G1").Font.Bold = True
ActiveSheet.Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & DestRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & DestRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & DestRow).Sort Key1:=Range("C2:C" & DestRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select
    
End Sub
Sub NorthHighlandDataDiscrepancyAnalysis()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("C:C").Delete
Range("W:X").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("A2:A" & LastRow).Formula = "=RIGHT(B2,FIND(""-"",B2,1)-1)"
Range("B:B").Delete

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:C1").Value = "A"
Range("D1").Value = "Apply Completed"
Range("E1").Value = "Apply Completed"
Range("F1").Value = "Apply Completed"
Range("G1").Value = "Apply Completed"
Range("H1").Value = "Apply Completed"
Range("I1").Value = "Apply Completed"
Range("J1").Value = "Apply Completed"
Range("K1").Value = "Qualified"
Range("L1").Value = "Qualified"
Range("M1").Value = "Qualified"
Range("N1").Value = "Interviewed"
Range("O1").Value = "Interviewed"
Range("P1").Value = "Qualified"
Range("Q1").Value = "Interviewed"
Range("R1").Value = "Offer Made"
Range("S1").Value = "Offer Made"
Range("T1").Value = "Offer Made"
Range("U1").Value = "Offer Made"
Range("V1").Value = "Hired"

LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 4

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = ActiveSheet.Range("A1:V" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 24)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)
    DestArray(1, 13) = SourceArray(2, 13)
    DestArray(1, 14) = SourceArray(2, 14)
    DestArray(1, 15) = SourceArray(2, 15)
    DestArray(1, 16) = SourceArray(2, 16)
    DestArray(1, 17) = SourceArray(2, 17)
    DestArray(1, 18) = SourceArray(2, 18)
    DestArray(1, 19) = SourceArray(2, 19)
    DestArray(1, 20) = SourceArray(2, 20)
    DestArray(1, 21) = SourceArray(2, 21)
    DestArray(1, 22) = SourceArray(2, 22)

For CurRow = 3 To LastRow
                   
        For CurCol = 4 To 22
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, 20)
                DestArray(DestRow, 21) = SourceArray(CurRow, 21)
                DestArray(DestRow, 22) = SourceArray(CurRow, 22)
                DestArray(DestRow, 23) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 24) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            End If
        Next CurCol
               
Next CurRow

ActiveSheet.Range("1:1").Delete

ActiveSheet.Range("A1:X" & DestRow).Value = DestArray

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:Y").Delete
Range("G:G").Cut Destination:=Range("B:B")
Range("F:F").Cut Destination:=Range("C:C")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

LastRow = Range("A65536").End(xlUp).Row

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Range("H:H").Delete

ActiveSheet.Range("A1:G" & DestRow).Font.Size = 10
ActiveSheet.Range("A1:G" & DestRow).Font.Name = "Arial"
ActiveSheet.Range("A1:G1").Font.Color = vbBlack
ActiveSheet.Range("A1:G1").Font.Bold = True
ActiveSheet.Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub GatesFoundation()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("D:I").Delete

Dim LastRow
LastRow = Range("B100000").End(xlUp).Row

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:C1").Value = "A"
Range("D1").Value = "Apply Completed"
Range("E1").Value = "Apply Completed"
Range("F1").Value = "Qualified"
Range("G1").Value = "Qualified"
Range("H1").Value = "Qualified"
Range("I1").Value = "Qualified"
Range("J1").Value = "Qualified"
Range("K1").Value = "Qualified"
Range("L1").Value = "Qualified"
Range("M1").Value = "Interviewed"
Range("N1").Value = "Interviewed"
Range("O1").Value = "Interviewed"
Range("P1").Value = "Offer Made"
Range("Q1").Value = "Offer Made"
Range("R1").Value = "Offer Made"
Range("S1").Value = "Hired"
Range("T1").Value = "Offer Made"
Range("U1").Value = "Offer Made"
Range("V1").Value = "Apply Completed"
Range("W1").Value = "Apply Completed"
Range("X1").Value = "Apply Completed"
Range("Y1").Value = "Apply Completed"
Range("Z1").Value = "Apply Completed"
Range("AA1").Value = "Apply Completed"

LastRow = Range("B100000").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 4

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = ActiveSheet.Range("A1:AA" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 29)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)
    DestArray(1, 13) = SourceArray(2, 13)
    DestArray(1, 14) = SourceArray(2, 14)
    DestArray(1, 15) = SourceArray(2, 15)
    DestArray(1, 16) = SourceArray(2, 16)
    DestArray(1, 17) = SourceArray(2, 17)
    DestArray(1, 18) = SourceArray(2, 18)
    DestArray(1, 19) = SourceArray(2, 19)
    DestArray(1, 20) = SourceArray(2, 20)
    DestArray(1, 21) = SourceArray(2, 21)
    DestArray(1, 22) = SourceArray(2, 22)
    DestArray(1, 23) = SourceArray(2, 23)
    DestArray(1, 24) = SourceArray(2, 24)
    DestArray(1, 25) = SourceArray(2, 25)
    DestArray(1, 26) = SourceArray(2, 26)
    DestArray(1, 27) = SourceArray(2, 27)

For CurRow = 3 To LastRow
                   
        For CurCol = 4 To 27
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, 11)
                DestArray(DestRow, 12) = SourceArray(CurRow, 12)
                DestArray(DestRow, 13) = SourceArray(CurRow, 13)
                DestArray(DestRow, 14) = SourceArray(CurRow, 14)
                DestArray(DestRow, 15) = SourceArray(CurRow, 15)
                DestArray(DestRow, 16) = SourceArray(CurRow, 16)
                DestArray(DestRow, 17) = SourceArray(CurRow, 17)
                DestArray(DestRow, 18) = SourceArray(CurRow, 18)
                DestArray(DestRow, 19) = SourceArray(CurRow, 19)
                DestArray(DestRow, 20) = SourceArray(CurRow, 20)
                DestArray(DestRow, 21) = SourceArray(CurRow, 21)
                DestArray(DestRow, 22) = SourceArray(CurRow, 22)
                DestArray(DestRow, 23) = SourceArray(CurRow, 23)
                DestArray(DestRow, 24) = SourceArray(CurRow, 24)
                DestArray(DestRow, 25) = SourceArray(CurRow, 25)
                DestArray(DestRow, 26) = SourceArray(CurRow, 26)
                DestArray(DestRow, 27) = SourceArray(CurRow, 27)
                DestArray(DestRow, 28) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 29) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            End If
        Next CurCol
               
Next CurRow

ActiveSheet.Range("1:1").Delete

ActiveSheet.Range("A1:AC" & DestRow).Value = DestArray

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:AC").Delete
Range("G:G").Cut Destination:=Range("B:B")
Range("F:F").Cut Destination:=Range("C:C")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

LastRow = Range("B100000").End(xlUp).Row

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Range("H:H").Delete

ActiveSheet.Range("A1:G" & DestRow).Font.Size = 10
ActiveSheet.Range("A1:G" & DestRow).Font.Name = "Arial"
ActiveSheet.Range("A1:G1").Font.Color = vbBlack
ActiveSheet.Range("A1:G1").Font.Bold = True
ActiveSheet.Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & DestRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & DestRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & DestRow).Sort Key1:=Range("C2:C" & DestRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select
    
End Sub

Sub HDSupply()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:B").Delete
Range("C:C").Delete
Range("F:F").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("I2:I" & LastRow).Formula = "=If(isblank(E2),D2,E2)"
Range("I2:I" & LastRow).Select
Selection.Copy
Range("I2:I" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("D:E").Delete

Range("H2:H" & LastRow).Formula = "=right(a2,4)"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("A:A").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("K:K").Cut Destination:=Range("D:D")
Range("F1").Select
ActiveCell.EntireColumn.Insert

Range("J2:J" & LastRow).Formula = "=h2&i2"
Range("J2:J" & LastRow).Select
Selection.Copy
Range("J2:J" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H:I").Delete

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "InBoxExternal Portal"
 DestArray(DestRow, 2) = "Apply Completed"
Case "InBoxInternal Portal"
 DestArray(DestRow, 2) = "Apply Completed"
Case "InBoxRecruiter"
 DestArray(DestRow, 2) = "Apply Completed"
Case "InBoxAgency"
 DestArray(DestRow, 2) = "Apply Completed"
Case "InBoxInitial DNQ"
 DestArray(DestRow, 2) = "Apply Completed"
Case "InBoxCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter ReviewReviewed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter ReviewReviewed; Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter ReviewRecruiter Phone Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter ReviewPhone Screen; Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter ReviewRecruiter Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Recruiter ReviewRecruiter Interviewed; Not Selected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Recruiter ReviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hiring Manager ReviewRecruiter Recommended"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager ReviewReviewed"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager ReviewHM Reviewed; Not Selected"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager ReviewHM Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager ReviewPhone Screen: Not Selected"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager ReviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Qualified"
Case "InterviewLaunch Skill Survey (pre-interview option)"
 DestArray(DestRow, 2) = "Qualified"
Case "InterviewHM 1st Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewLaunch Skill Survey (post-interview option)"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewSend HM Interview Feedback Summary iForm"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewHM 1st Interview; Not Selected"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewHM 2nd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewHM 2nd Interview; Not Selected"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewHM Additional Interviews"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewHM Additional Interviews; Not Selected"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewCandidate Withdrew"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferLaunch Offer Letter Worksheet Process"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferLaunch Offer Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferVerbal Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferVerbal Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer Extended - Send Offer Email"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer Declined/Rejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferSend SSN & DOB Collection iForm"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferBackground/Drug Check Initiated"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferBackground/Drug Check Completed"
 DestArray(DestRow, 2) = "Offer Made"
Case "HiredExternal Hire (Send to HRIS)"
 DestArray(DestRow, 2) = "Hired"
Case "HiredInternal Hire (Do NOT send to HRIS)"
 DestArray(DestRow, 2) = "Hired"
Case "HiredNo Show"
 DestArray(DestRow, 2) = "Hired"
Case "HiredExecutive Hire (Do NOT send to HRIS)"
 DestArray(DestRow, 2) = "Hired"
Case "InBoxEmployee Referral"
 DestArray(DestRow, 2) = "Apply Completed"
Case "IncompleteIncomplete"
 DestArray(DestRow, 2) = "ATS Captured"
Case "OfferBackground/Drug Check Failed"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferCreate Hourly Offer Letter"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferLaunch Hiring Manager Survey - Exempt Only"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferLaunch Offer Letter Worksheet process"
 DestArray(DestRow, 2) = "Offer Made"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Merck()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:1").Delete
Range("A:B").Delete
Range("D:D").Delete
Range("G:H").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert

Range("A1:J" & LastRow).Select
Selection.NumberFormat = "General"

Range("H2:H" & LastRow).Formula = "=I2&J2"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues



Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "New ProspectDeclined"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New ProspectMove Forward"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New ProspectRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New ProspectTo be evaluated"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New ProspectWaiting for info"
 DestArray(DestRow, 2) = "Apply Completed"
Case "HM ReviewApproved by HM"
 DestArray(DestRow, 2) = "Qualified"
Case "HM ReviewDeclined"
 DestArray(DestRow, 2) = "Qualified"
Case "HM ReviewForwarded to HM"
 DestArray(DestRow, 2) = "Qualified"
Case "HM ReviewRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "HM ReviewTo Be Forwarded to HM"
 DestArray(DestRow, 2) = "Qualified"
Case "Initial ScreenDeclined"
 DestArray(DestRow, 2) = "Qualified"
Case "Initial ScreenLeft a message"
 DestArray(DestRow, 2) = "Qualified"
Case "Initial ScreenMove Forward"
 DestArray(DestRow, 2) = "Qualified"
Case "Initial ScreenRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Initial ScreenScheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "Initial ScreenTo Be Initially Screened"
 DestArray(DestRow, 2) = "Qualified"
Case "Initial ScreenWaiting for info"
 DestArray(DestRow, 2) = "Qualified"
Case "1st In-Person InterviewDeclined"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st In-Person InterviewLeft a message"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st In-Person InterviewMove Forward"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st In-Person InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st In-Person InterviewScheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st In-Person InterviewTo be scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st In-Person InterviewWaiting for info"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd In-Person InterviewDeclined"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd In-Person InterviewLeft a message"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd In-Person InterviewMove Forward"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd In-Person InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd In-Person InterviewScheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd In-Person InterviewTo be scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd In-Person InterviewWaiting for info"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd In-Person InterviewDeclined"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd In-Person InterviewLeft a message"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd In-Person InterviewMove Forward"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd In-Person InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd In-Person InterviewScheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd In-Person InterviewTo be scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd In-Person InterviewWaiting for info"
 DestArray(DestRow, 2) = "Interviewed"
Case "EvaluationDeclined"
 DestArray(DestRow, 2) = "Interviewed"
Case "EvaluationIn process"
 DestArray(DestRow, 2) = "Interviewed"
Case "EvaluationMoved Forward"
 DestArray(DestRow, 2) = "Interviewed"
Case "EvaluationRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "EvaluationTo be evaluated"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferAccepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferDeclined"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferExtended"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferLeft a message"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferNegotiating"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer countered"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer pending"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer to be made"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferVerbal offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferWritten offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRefused"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferApproved"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment CheckPassed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment CheckPassed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment CheckPassed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment CheckTo be checked"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment CheckWaiting for results "
 DestArray(DestRow, 2) = "Offer Made"
Case "HireDeclined"
 DestArray(DestRow, 2) = "Offer Made"
Case "HireHired"
 DestArray(DestRow, 2) = "Hired"
Case "HireIn process"
 DestArray(DestRow, 2) = "Hired"
Case "HireRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferCanceled"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment CheckDeclined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment CheckRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Prescreen / TestDeclined"
 DestArray(DestRow, 2) = "Qualified"
Case "OfferApproval in Progress"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferReneged"
 DestArray(DestRow, 2) = "Offer Made"
Case "EvaluationMove Forward"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferApproval Rejected"
 DestArray(DestRow, 2) = "Offer Made"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub JohnsonandJohnson()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("C:F").Delete
Range("D:D").Delete
Range("G:J").Delete

Dim LastRow
LastRow = Range("A400000").End(xlUp).Row

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("I:I").Cut Destination:=Range("C:C")

Range("H2:H" & LastRow).Formula = "=F2&G2"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("F:G").ClearContents

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 400000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "GRU NewPending Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "GRU NewReject Letter Sent"
 DestArray(DestRow, 2) = "Apply Completed"
Case "GRU NewRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "GRU NewWithdrawn"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NewPending Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NewReject Letter Sent"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NewRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NewWithdrawn"
 DestArray(DestRow, 2) = "Apply Completed"
Case "GRU PreScreenInterim Letter Sent (Rejected)"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU PreScreenManager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU PreScreenManager Review Completed"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU PreScreenPending Review"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU PreScreenPrescreen Completed - GRU"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU PreScreenPrescreen Scheduled - GRU"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU PreScreenReject Letter Sent"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU PreScreenRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU PreScreenReviewed"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU PreScreenWithdrawn"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager ReviewInterim Letter Sent (Rejected)"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager ReviewMove to Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager ReviewReject Letter Sent"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager ReviewRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager ReviewUnder Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager ReviewWithdrawn"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenDue Diligence Completed"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenDue Diligence Initiated"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenDue Diligence Review Requested by SC"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenPrescreen Completed"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenPrescreen Request Submitted to SC"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenPrescreen Scheduled by SC"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenReady to Share with Hiring Manager"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenReject Letter Sent"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenShared With Hiring Manager"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre ScreenWithdrawn"
 DestArray(DestRow, 2) = "Qualified"
Case "UR_PreScreenPending Review"
 DestArray(DestRow, 2) = "Qualified"
Case "UR_PreScreenReject Letter Sent"
 DestArray(DestRow, 2) = "Qualified"
Case "UR_PreScreenRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "UR_PreScreenReviewed"
 DestArray(DestRow, 2) = "Qualified"
Case "UR_PreScreenWithdrawn"
 DestArray(DestRow, 2) = "Qualified"
Case "1st Interview1st Interview Request Submitted to SC"
 DestArray(DestRow, 2) = "Qualified"
Case "1st Interview1st Interview Scheduled by SC"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st Interview1st Interview Request Submitted to SC"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st Interview1st Interview Scheduled - Conference(Vendor)"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st Interview1st Interview Scheduled - Phone (GRU)"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st Interview1st Interview Scheduled - Phone (Vendor)"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st Interview1st Interview Scheduled by GRU"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st Interview1st Interview Scheduled by SC"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st InterviewCampus Interview Scheduled - (GRU)"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st InterviewCampus Interview Scheduled - (Vendor)"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st InterviewInvite for Phone Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st InterviewInvite to Campus Interview/ Alternate"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st InterviewInvite to Campus Interview/Primary"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st InterviewInvite to Video Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st InterviewInvite to Conference - Primary"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 1st InterviewInvite to Phone Update"
 DestArray(DestRow, 2) = "Qualified"
Case "UR_PreScreenApplicant Shared"
 DestArray(DestRow, 2) = "Qualified"
Case "1st Interview1st Interview Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewAssessment completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewAssessment scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewDue Diligence Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewInterim Letter Sent (Rejected)"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewReject Letter Sent"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewWithdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterview2nd Interview Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterview2nd Interview Request Submitted to SC"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterview2nd Interview Scheduled by SC"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterviewAssessment completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterviewAssessment scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterviewDue Diligence Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterviewDue Diligence Initiated"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterviewInterim Letter Sent (Rejected)"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterviewReject Letter Sent"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "2ndInterviewWithdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd Interview3rd Interview Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd Interview3rd Interview Request Submitted to SC"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd Interview3rd Interview Scheduled by SC"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewDue Diligence Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewDue Diligence Initiated"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewDue Diligence Review Requested by SC"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewReject Letter Sent"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewWithdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st Interview1st Interview Completed - Conference"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st InterviewAssessment completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st InterviewCampus Interview ? Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st InterviewInterim Letter Sent (Rejected)"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st InterviewPhone Interview ? Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st InterviewPhone Update Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st InterviewReject Letter Sent"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st InterviewVideo Interview Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st InterviewWithdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd Interview2nd Interview Completed - Conference"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd Interview2nd Interview Request Submitted to SC"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd Interview2nd Interview Scheduled - Invitational"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd InterviewPost Applicant to Invitational/Manager's Website"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd InterviewReject Letter Sent"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd InterviewWithdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd Interview3rd Interview Completed - Invitational"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd InterviewAssessment completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd InterviewPhone Interview ? Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd InterviewReject Letter Sent"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd InterviewSite Visit Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd InterviewWithdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU Pre OfferOffer Targeted"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU Pre OfferWithdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU Pre OfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "On Site InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Pre OfferReady for Written Offer"
 DestArray(DestRow, 2) = "Interviewed"
Case "Pre OfferReject Letter Sent"
 DestArray(DestRow, 2) = "Interviewed"
Case "Pre OfferRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Pre OfferVerbal Offer Accepted"
 DestArray(DestRow, 2) = "Interviewed"
Case "Pre OfferVerbal Offer Extended"
 DestArray(DestRow, 2) = "Interviewed"
Case "Pre OfferWithdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "UR Additional InterviewAdditional Interview - Completed In Person"
 DestArray(DestRow, 2) = "Interviewed"
Case "UR invitational InterviewOffer Targeted"
 DestArray(DestRow, 2) = "Interviewed"
Case "UR_Campus InterviewReject Letter Sent"
 DestArray(DestRow, 2) = "Interviewed"
Case "UR_Campus InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "UR_Campus InterviewWithdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "UR_Campus InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferCanceled"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferDraft"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewDue Diligence Initiated"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st Interview1st Interview Completed by GRU"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd Interview2nd Interview Scheduled by GRU"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd Interview3rd Interview Completed by GRU"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd Interview3rd Interview Request Submitted to SC"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU Pre OfferRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "UR_Campus InterviewPhone Interview ? Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferAccepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferExtended"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRefused"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferCriminal Background Check Completed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferCriminal Background Check Initiated"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferDrug Screen Initiated"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferDue Diligence Completed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferWritten Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "HireReject Letter Sent"
 DestArray(DestRow, 2) = "Offer Made"
Case "HireRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferDrug Screen Completed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferDue Diligence Review Requested by SC"
 DestArray(DestRow, 2) = "Offer Made"
Case "HireCleared for Hire"
 DestArray(DestRow, 2) = "Hired"
Case "HireCleared for Hire - GE"
 DestArray(DestRow, 2) = "Hired"
Case "HireCleared for Hire - GM"
 DestArray(DestRow, 2) = "Hired"
Case "HireCleared for Hire - Non Employee"
 DestArray(DestRow, 2) = "Hired"
Case "HireCleared for Hire - SC"
 DestArray(DestRow, 2) = "Hired"
Case "HireComplete for Hire"
 DestArray(DestRow, 2) = "Hired"
Case "HireComplete for Hire by SC"
 DestArray(DestRow, 2) = "Hired"
Case "HireMergers & Acquisitions - Cleared for Hire"
 DestArray(DestRow, 2) = "Hired"
Case "HireTraining Program"
 DestArray(DestRow, 2) = "Hired"
Case "GRU 2nd Interview2nd Interview Completed by GRU"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd InterviewPhone Interview ? Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd InterviewInterim Letter Sent (Rejected)"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU Pre OfferInterim Letter Sent (Rejected)"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU PreScreenPrescreen Request Submitted to SC"
 DestArray(DestRow, 2) = "Qualified"
Case "Hire - NonEmployeeCleared for Hire - GE"
 DestArray(DestRow, 2) = "Hired"
Case "Post OfferCredit Check Initiated"
 DestArray(DestRow, 2) = "Offer Made"
Case "GRU 2nd Interview2nd Interview Completed - Invitational"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd Interview2nd Interview Scheduled - Phone (GRU)"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd InterviewInvite to Conference - Primary"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd InterviewSite Visit Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd InterviewSite Visit Scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 3rd Interview3rd Interview Scheduled by GRU"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferReneged"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferDriving History Check Completed"
 DestArray(DestRow, 2) = "Offer Made"
Case "2ndInterviewDue Diligence Review Requested by SC"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 1st Interview1st Interview Scheduled - Video (GRU)"
 DestArray(DestRow, 2) = "Qualified"
Case "GRU 2nd Interview2nd Interview Scheduled - Video (GRU)"
 DestArray(DestRow, 2) = "Interviewed"
Case "GRU 2nd InterviewVideo Interview Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Post OfferEducation Verification Initiated"
 DestArray(DestRow, 2) = "Offer Made"
Case "GRU 1st InterviewPhone Update Scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "NewMeets Qualifications"
 DestArray(DestRow, 2) = "Qualified"
Case "OfferIn Negotiation"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferDriving History Check Initiated"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferEducation Verification Completed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post OfferEducation Verification Initiated"
 DestArray(DestRow, 2) = "Offer Made"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub BNYMellon()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("C:L").Delete
Range("D:H").Delete
Range("E:E").Delete
Range("G:Z").Delete

Dim LastRow
LastRow = Range("A300000").End(xlUp).Row

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("C:C")
Range("F:F").Delete

Range("H2:H" & LastRow).Formula = "=F2&G2"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("F:G").ClearContents

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 300000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Hire-External Hire- Move to External New Hire PeopleSoft Processing"
 DestArray(DestRow, 2) = "Hired"
Case "Hire-FTC Hire - Move to FTC PeopleSoft Processing"
 DestArray(DestRow, 2) = "Hired"
Case "Hire-Has Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hire-Internal- Alternative Process; Manual Closing Req"
 DestArray(DestRow, 2) = "Hired"
Case "Hire-Internal hire- Move to Internal Transfer PeopleSoft Processing"
 DestArray(DestRow, 2) = "Hired"
Case "Hire-Ready to Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hire-Rehire from Severance - Manual Closing Req"
 DestArray(DestRow, 2) = "Hired"
Case "Hire-Rehire- Move to Rehire Transfer PeopleSoft Processing"
 DestArray(DestRow, 2) = "Hired"
Case "Hire-Rejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Interviews1st Interview Round Completed - Under Consideration"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviews1st Interview Round Scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "Interviews2nd Interview Round Completed - Under Consideration"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviews2nd Interview Round Scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviews3rd Interview Round Completed - Under Consideration"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviews3rd Interview Round Scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewsHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewsRejected from Step Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "InterviewsTo be Interviewed"
 DestArray(DestRow, 2) = "Qualified"
Case "InterviewsTo be Interviewed (US) with invitation to provide add'l info"
 DestArray(DestRow, 2) = "Qualified"
Case "New-Government Questionnaire Approved"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New-Government Questionnaire Sent"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New-Has Declined"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New-Move Forward"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New-Rejected from Step New"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New-To be Reviewed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New-Under Consideration"
 DestArray(DestRow, 2) = "Apply Completed"
Case "OfferAccepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferApproval in Progress"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferApproval Rejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferApproved"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferCanceled"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferDraft"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferExtended"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferIn Negotiation"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferRefused"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferRescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Checks-Has Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Checks-Pre-Hire Completed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Checks-Rejected from Pre-Hire Checks Step"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Checks-To be Initiated"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire Checks-Securities License Check Completed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Screen-Forwarded to Hiring Manager (Not Shared)"
 DestArray(DestRow, 2) = "Qualified"
Case "Screen-Has Declined"
 DestArray(DestRow, 2) = "Qualified"
Case "Screen-Move Forward"
 DestArray(DestRow, 2) = "Qualified"
Case "Screen-Recruiter Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Screen-Rejected from Step Screen"
 DestArray(DestRow, 2) = "Qualified"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub BombardierTransportation()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:3").Delete
Range("A:I").Delete
Range("C:E").Delete
Range("D:E").Delete
Range("G:G").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F:G").Delete

Range("H2:H" & LastRow).Formula = "=F2&G2"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("F:G").ClearContents

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Medical CheckMedical Check 1 to Be Scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "NewShort List"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NewHas Declined"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Medical CheckHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case "DecisionShort List"
 DestArray(DestRow, 2) = "Interviewed"
Case "NewRejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NewUnder consideration"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NewStandby"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NewWaiting for info"
 DestArray(DestRow, 2) = "Apply Completed"
Case "DecisionTo be asserted"
 DestArray(DestRow, 2) = "Interviewed"
Case "NewTo be evaluated"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Medical CheckMedical Check 1 Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "TestTo be tested"
 DestArray(DestRow, 2) = "Interviewed"
Case "TestShort List"
 DestArray(DestRow, 2) = "Interviewed"
Case "DecisionUnder consideration"
 DestArray(DestRow, 2) = "Interviewed"
Case "DecisionRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Medical CheckMedical Check 3 Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "TestWaiting for results"
 DestArray(DestRow, 2) = "Interviewed"
Case "TestRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Medical CheckMedical Check 3 to Be Scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "Medical CheckRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "DecisionHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case "TestLeft a message"
 DestArray(DestRow, 2) = "Interviewed"
Case "TestScheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "TestUnder consideration"
 DestArray(DestRow, 2) = "Interviewed"
Case "GlobalHas Declined"
 DestArray(DestRow, 2) = "Apply Completed"
Case "TestHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case "HM ReviewRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone screenScheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone screenTo be phone screened"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone screenRejected"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone screenShort List"
 DestArray(DestRow, 2) = "Qualified"
Case "HM ReviewForwarded to HM"
 DestArray(DestRow, 2) = "Qualified"
Case "HM ReviewApproved by HM"
 DestArray(DestRow, 2) = "Qualified"
Case "HM ReviewTo Be Forwarded to HM"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone screenLeft a message"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone screenUnder consideration"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone screenHas Declined"
 DestArray(DestRow, 2) = "Qualified"
Case "HM ReviewHas Declined"
 DestArray(DestRow, 2) = "Qualified"
Case "1st InterviewUnder consideration"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd InterviewHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd InterviewTo be scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewScheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewShort List"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewTo be scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd InterviewShort List"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewShort List"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewScheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd InterviewScheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd InterviewUnder consideration"
 DestArray(DestRow, 2) = "Interviewed"
Case "1st InterviewLeft a message"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewTo be scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewLeft a message"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd InterviewLeft a message"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewUnder consideration"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd InterviewHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferOffer to be made"
 DestArray(DestRow, 2) = "Interviewed"
Case "Background CheckPassed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Verbal OfferTo Be Made"
 DestArray(DestRow, 2) = "Interviewed"
Case "Written OfferTo Be Drawn Up"
 DestArray(DestRow, 2) = "Offer Made"
Case "Verbal OfferOffer accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "OfferHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Verbal OfferHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Written OfferOffer accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Background CheckTo be backgr. checked"
 DestArray(DestRow, 2) = "Interviewed"
Case "OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Written OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Written OfferHas Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Verbal OfferRejected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Background CheckRejected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Background CheckHas Declined"
 DestArray(DestRow, 2) = "Interviewed"
Case "Background CheckWaiting for results"
 DestArray(DestRow, 2) = "Interviewed"
Case "HireTo be hired"
 DestArray(DestRow, 2) = "Hired"
Case "HireHired"
 DestArray(DestRow, 2) = "Hired"
Case "HireRejected"
 DestArray(DestRow, 2) = "Hired"
Case "HireHas Declined"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Bloomberg()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:B").Delete
Range("B:D").Delete
Range("D:H").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")

Range("H2:H" & LastRow).Formula = "=F2&G2"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("F:G").ClearContents

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "New Candidate1 New Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Review1 Recruiter"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Review2 Business"
 DestArray(DestRow, 2) = "Qualified"
Case "Review3 Eligibility"
 DestArray(DestRow, 2) = "Qualified"
Case "Review4 Hold"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview1 Scheduled/In Progess"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview3 Feedback"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview4 Hold"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer1 Pending Details"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer2 Details"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer3 Pending Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer4 Extend"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer5 Intent to accept"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted1 Pending Background"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted2 Background Complete"
 DestArray(DestRow, 2) = "Hired"
Case "Offer AcceptedOffer Accept"
 DestArray(DestRow, 2) = "Hired"
Case "Offer AcceptedPending Visa"
 DestArray(DestRow, 2) = "Hired"
Case "Offer AcceptedVisa in Progress"
 DestArray(DestRow, 2) = "Hired"
Case "Offer AcceptedVisa Complete"
 DestArray(DestRow, 2) = "Hired"
Case "OfferPending Candidate Decision"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired1 Complete"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub UMMC()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:G").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("C:C").Cut Destination:=Range("I:I")
Range("B:B").Cut Destination:=Range("H:H")
Range("C:C").Delete

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Inbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Suitable"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Sourced"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Phone Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Selected For Interview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Resume Received"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Finalists"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Peer Review/Share Day"
 DestArray(DestRow, 2) = "Qualified"
Case "Schedule Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Being Considered for Other Positions"
 DestArray(DestRow, 2) = "Qualified"
Case "Consider for Other Positions"
 DestArray(DestRow, 2) = "Qualified"
Case "Additional Interviews"
 DestArray(DestRow, 2) = "Interviewed"
Case "Comp Creates/Reviews Offer"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Not Selected after Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Request Offer"
 DestArray(DestRow, 2) = "Interviewed"
Case "UMMC Only - Approve Offer"
 DestArray(DestRow, 2) = "Interviewed"
Case "Recruiter Follow Up Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Notification"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment Background"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Bechtel()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:C").Delete
Range("D:I").Delete
Range("E:E").Delete
Range("F:G").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("B:B").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F:F").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject-Skills/Exp/Other (e)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject-Not considered (e)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject-Basic quals (e)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject-Position cancelled (e)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdraw-Unable to contact"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdraw-Other"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject-Temp Agency"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Line Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Functional Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Client Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Under Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Prehire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub ActiveNetwork()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:C").Delete
Range("E:I").Delete
Range("F:N").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A:A").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("A:A")
Range("B:B").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("D:D")
Range("E:E").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("E:E")
Range("F:F").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Job Seeker"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Archives"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Captured Offline"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Unsolicited (Do not use)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Captured Online/Job Boards"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Sent to Manager"
 DestArray(DestRow, 2) = "Qualified"
Case "Talent Pool"
 DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Manager Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "In-Person Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Final Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub GardaAviation()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("F:G").Delete

Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New Application"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Admin"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "(blank)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background Check"
 DestArray(DestRow, 2) = "Qualified"
Case "Candidate Withdrawal"
 DestArray(DestRow, 2) = "Qualified"
Case "Clearance Check"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre-Qualification Tests and ORT"
 DestArray(DestRow, 2) = "Qualified"
Case "Psychometric"
 DestArray(DestRow, 2) = "Qualified"
Case "Short List"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Medical Testing"
 DestArray(DestRow, 2) = "Interviewed"
Case "Transfer"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Orientation"
 DestArray(DestRow, 2) = "Offer Made"
Case "SOF Practical"
 DestArray(DestRow, 2) = "Offer Made"
Case "SOF Training Classroom"
 DestArray(DestRow, 2) = "Offer Made"
Case "Training Emodule"
 DestArray(DestRow, 2) = "Offer Made"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub GardaCashServices()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("F:G").Delete

Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "New Application"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrawal"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "(blank)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Transfer"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background Check"
 DestArray(DestRow, 2) = "Qualified"
Case "CFSC-CRFSC"
 DestArray(DestRow, 2) = "Qualified"
Case "PAL"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Interviewed"
Case "Tests"
 DestArray(DestRow, 2) = "Interviewed"
Case "Weapons Permit"
 DestArray(DestRow, 2) = "Interviewed"
Case "Agent Permit"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Medical Testing"
 DestArray(DestRow, 2) = "Interviewed"
Case "Scenarios Training"
 DestArray(DestRow, 2) = "Interviewed"
Case "Shooting Range Training"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Orientation"
 DestArray(DestRow, 2) = "Offer Made"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub GardaHRProtectiveServices()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("F:G").Delete

Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New Application"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrawal"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "(blank)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Transfer"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background Check"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Tests"
 DestArray(DestRow, 2) = "Qualified"
Case "Client Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Declined DPCS"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Security License"
 DestArray(DestRow, 2) = "Interviewed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Yum()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:B").Delete
Range("B:B").Delete
Range("D:H").Delete
Range("F:G").Delete

Range("B:B").Select
ActiveCell.EntireColumn.Insert
Range("C:C").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F:F").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "No Interest - RSC"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Regret - RSC"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background Check Info-No response"
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiting Review - RSC"
 DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Phone Screen - RSC"
 DestArray(DestRow, 2) = "Qualified"
Case "Reference Check - RSC"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Team Interview - RSC"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Team 2nd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Team 3rd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Collect Background Check Info"
 DestArray(DestRow, 2) = "Interviewed"
Case "Background Check - RSC"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Team 2nd Interview - RSC"
 DestArray(DestRow, 2) = "Interviewed"
Case "Extend Offer - RSC"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired - RSC"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Successfactors()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:A").Delete
Range("F:F").Delete

Range("B:B").Select
ActiveCell.EntireColumn.Insert
Range("C:C").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F:F").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Background /Reference Check"
 DestArray(DestRow, 2) = "Offer Made"
Case "Candidate Rejected SF – Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Candidate Rejected SF – Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Candidate Rejected SF – Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Finalist"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hiring Manager Review/Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview – In Process"
 DestArray(DestRow, 2) = "Interviewed"
Case "New Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer – Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer – Pending Approval"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer – Pending Candidate Accept"
 DestArray(DestRow, 2) = "Offer Made"
Case "Screen – Recruiter"
 DestArray(DestRow, 2) = "Qualified"
Case "SF Rejected Candidate – Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "SF Rejected Candidate - Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "SF Rejected Candidate – Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = ""
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub NewYorkTimes()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("C:G").Delete
Range("D:E").Delete
Range("F:H").Delete

Range("A:A").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("D:D").Delete
Range("B:B").Select
ActiveCell.EntireColumn.Insert
Range("C:C").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F:F").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case ""
 DestArray(DestRow, 2) = "ATS Captured"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "HR Phone Screening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Pending Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Phone Screening - Did not pass HR Phone Screening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Prescreen - Not Considered - DId Not Meet Basic Qualifications"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Onsite Interview - Failed to Respond/No Show"
 DestArray(DestRow, 2) = "Apply Completed"
Case "MGR Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Onsite Interview - Did not pass Interview Assessment by Manager"
 DestArray(DestRow, 2) = "Interviewed"
Case "Onsite Interview - Other more qualified candidate selected"
 DestArray(DestRow, 2) = "Interviewed"
Case "Onsite Interview - Withdrew Candidacy"
 DestArray(DestRow, 2) = "Interviewed"
Case "Onsite Interview Scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "Onsite Interview Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Phone Interview - Did not pass MGR Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer - Declined Job Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Did not pass background check"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Offer"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Leidos()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:3").Delete
Range("A:A").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("AE2:AE" & LastRow).Formula = "=if(P2 = ""Internal Transfer"",""delete"","""")"

Range("AE2:AE" & LastRow).Select
Selection.Copy
Range("AE2:AE" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Dim CurRow1
CurRow1 = 2

Do While CurRow1 < LastRow
If Range("AE" & CurRow1).Value = "delete" Then
Range(CurRow1 & ":" & CurRow1).Delete
LastRow = Range("A65536").End(xlUp).Row
Else
CurRow1 = CurRow1 + 1
End If
Loop

LastRow = Range("A65536").End(xlUp).Row

Range("AE2:AE" & LastRow).Delete

Range("A:D").Delete
Range("B:D").Delete
Range("C:D").Delete
Range("D:E").Delete
Range("E:H").Delete
Range("F:H").Delete
Range("G:K").Delete

LastRow = Range("A65536").End(xlUp).Row

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("A:A")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("E:E").Cut Destination:=Range("C:C")
Range("I:I").Cut Destination:=Range("E:E")
Range("F1").Select
ActiveCell.EntireColumn.Insert


Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "New"
DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Review"
DestArray(DestRow, 2) = "Qualified"
Case "HM Review"
DestArray(DestRow, 2) = "Qualified"
Case "Interview"
DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
DestArray(DestRow, 2) = "Offer Made"
Case "Hire"
DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Cintas()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:A").Delete
Range("B:B").Delete
Range("C:C").Delete
Range("D:D").Delete

Dim LastRow
LastRow = Range("B65536").End(xlUp).Row

Range("G:G").Select
ActiveCell.EntireColumn.Insert
Range("G1").Value = "Ace Date"
Range("G2:G" & LastRow).Formula = "=IF(C2=""Yes"",F2,"""")"

Range("G2:G" & LastRow).Select
Selection.Copy
Range("G2:G" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1:E1").Value = "A"
Range("F1").Value = "Apply Completed"
Range("G1").Value = "Qualified"
Range("H1").Value = "Interviewed"
Range("I1").Value = "Offer Made"
Range("J1").Value = "Hired"

LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 3

Dim CurCol
CurCol = 6

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = ActiveSheet.Range("A1:L" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 1500000, 1 To 12)

    DestArray(1, 1) = SourceArray(2, 1)
    DestArray(1, 2) = SourceArray(2, 2)
    DestArray(1, 3) = SourceArray(2, 3)
    DestArray(1, 4) = SourceArray(2, 4)
    DestArray(1, 5) = SourceArray(2, 5)
    DestArray(1, 6) = SourceArray(2, 6)
    DestArray(1, 7) = SourceArray(2, 7)
    DestArray(1, 8) = SourceArray(2, 8)
    DestArray(1, 9) = SourceArray(2, 9)
    DestArray(1, 10) = SourceArray(2, 10)
    DestArray(1, 11) = SourceArray(2, 11)
    DestArray(1, 12) = SourceArray(2, 12)

For CurRow = 3 To LastRow
                   
        For CurCol = 6 To 12
            If SourceArray(CurRow, CurCol) <> "" Then
                             
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                DestArray(DestRow, 9) = SourceArray(CurRow, 9)
                DestArray(DestRow, 10) = SourceArray(CurRow, 10)
                DestArray(DestRow, 11) = SourceArray(CurRow, CurCol)
                DestArray(DestRow, 12) = SourceArray(1, CurCol)
                               
                DestRow = DestRow + 1
                        
            Else
            End If
        Next CurCol
               
Next CurRow

ActiveSheet.Range("1:1").Delete

ActiveSheet.Range("A1:L" & DestRow).Value = DestArray

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("A:A")
Range("F:F").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("C1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Delete
Range("G:K").Delete
Range("H:H").Cut Destination:=Range("B:B")
Range("G:G").Cut Destination:=Range("C:C")
Range("F:F").Cut Destination:=Range("G:G")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

LastRow = Range("B200000").End(xlUp).Row

Range("H2:H" & DestRow).Formula = Range("C2:C" & DestRow).Value2
Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("H2:H" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("H2:H" & DestRow).Select
Selection.Copy
Range("C2:C" & DestRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("H2:H" & DestRow).Delete
Range("C2:C" & DestRow).NumberFormat = "mm-dd-yyyy"

Range("H:H").Delete

ActiveSheet.Range("A1:G" & DestRow).Font.Size = 10
ActiveSheet.Range("A1:G" & DestRow).Font.Name = "Arial"
ActiveSheet.Range("A1:G1").Font.Color = vbBlack
ActiveSheet.Range("A1:G1").Font.Bold = True
ActiveSheet.Range("A1:G1").Interior.Color = vbYellow

Range("A1:G" & DestRow).Borders.Weight = xlThin
Range("A1:G" & DestRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & DestRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & DestRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & DestRow).Sort Key1:=Range("C2:C" & DestRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select
    
End Sub

Sub CintasInterview()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:3").Delete
Range("A:A").Delete
Range("D:D").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("D:D").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert

Dim LastRow
LastRow = Range("D65536").End(xlUp).Row

Range("B2:B" & LastRow).Value = "Interviewed"

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Target()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:6").Delete
Range("A:B").Delete
Range("A:A").Delete
Range("C:K").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("D:D").Delete
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("G:G").Cut Destination:=Range("C:C")
Range("F1").Select
ActiveCell.EntireColumn.Insert
ActiveCell.EntireColumn.Insert

Dim LastRow
LastRow = Range("D65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Apply Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Apply Started"
 DestArray(DestRow, 2) = "Apply Started"
Case "ATS Capture"
 DestArray(DestRow, 2) = "ATS Capture"
Case "Event - Candidate Not Interested (R)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Event - Candidate Not Interested ®"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Event - Future Consideration (R)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Event - Invite (C) (R)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Event - Move to Requisition"
 DestArray(DestRow, 2) = "Qualified"
Case "Event - New (R)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Event - Not Invite (C) (R)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Event - Target Not Interested (R)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hire - Candidate Not Interested (R)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hire - Completed"
 DestArray(DestRow, 2) = "Hired"
Case "Hire - Hire Details Pending (R)"
 DestArray(DestRow, 2) = "Hired"
Case "Hire - Hired (R)"
 DestArray(DestRow, 2) = "Hired"
Case "Hire - Initiate Onboarding (C)"
 DestArray(DestRow, 2) = "Hired"
Case "Hire - No Show (R) (C)"
 DestArray(DestRow, 2) = "Hired"
Case "Hire - Target Not Interested (R)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "HM Interviewed"
 DestArray(DestRow, 2) = "HM Interviewed"
Case "Initial Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - Candidate Not Interested (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - Future Consideration (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - HM - Move Forward"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - HM - Not Interested"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - HM Review (C)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - HM/Technical Phone Interview (C)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - Initial Interview (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - Interview Scheduled (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - Request Interview (C)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - Request Interview (C) (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview  - Target Not Interested (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview - Candidate Not Interested (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview - Future Consideration (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview - HM - Move Forward"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview - HM Review (C)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview - HM/Technical Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview - Initial Interview (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview - Interview Scheduled (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview - Request Interview (C)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initial Interview - Target Not Interested (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Candidate Not Interested (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Future Consideration (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Interview (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Move To Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Interview - Move To Offer (R)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Interview - Moved to Associated Hiring Req (R)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Interview - Schedule Additional Interviews (C)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview - Target Not Interested (R)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Not Hired"
 DestArray(DestRow, 2) = "Not Hired"
Case "Not Qualified"
 DestArray(DestRow, 2) = "Not Qualified"
Case "Offer - Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Approval in Progress"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Approved"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Canceled"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Candidate Not Interested (R)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Draft"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - In Negotiation"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Offer to be made"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Offer to be made "
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Refused"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Reneged"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer - Target Not Interested (R)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Made"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post Offer Checks - Candidate Not Interested (R)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post Offer Checks - Check Not Required"
 DestArray(DestRow, 2) = "Hired"
Case "Post Offer Checks - Future Consideration (R)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post Offer Checks - In Progress"
 DestArray(DestRow, 2) = "Hired"
Case "Post Offer Checks - Post Offer Check Decision"
 DestArray(DestRow, 2) = "Offer Made"
Case "Post Offer Checks - Results Received"
 DestArray(DestRow, 2) = "Hired"
Case "Post Offer Checks - Target Not Interested (R)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Qualified"
 DestArray(DestRow, 2) = "Qualified"
Case "Qualified: Business Screen"
 DestArray(DestRow, 2) = "Qualified: Business Screen"
Case "Qualified: Recruiter Screen"
 DestArray(DestRow, 2) = "Qualified: Recruiter Screen"
Case "Qualified: Technical Screen"
 DestArray(DestRow, 2) = "Qualified: Technical Screen"
Case "Recruiter Interviewed"
 DestArray(DestRow, 2) = "Recruiter Interviewed"
Case "Recruiter Review - Candidate Not Interested (R)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Review - Future Consideration (R)"
 DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Review - Interview Scheduled (R)"
 DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Review - New (R)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Review - Request Interview (C)"
 DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Review - Target Not Interested (R)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Review - Under Consideration"
 DestArray(DestRow, 2) = "Qualified"
Case "Subscribe"
 DestArray(DestRow, 2) = "Subscribe"
Case "Unknown"
 DestArray(DestRow, 2) = "Unknown"
Case "Visit"
 DestArray(DestRow, 2) = "Visit"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub DSW()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("D:J").Delete
Range("E:F").Delete
Range("F:G").Delete
Range("G:I").Delete

Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("C:C")
Range("F:F").Delete
Range("F:F").Select
Range("F:F").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("D300000").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 300000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Archive"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NA"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject"
 DestArray(DestRow, 2) = "Apply Completed"
Case "W/draw"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "(Blanks)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Corp Rvw"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disq-Asmt"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Store Rvw"
 DestArray(DestRow, 2) = "Qualified"
Case "Disq-BG Ck"
 DestArray(DestRow, 2) = "Interviewed"
Case "BG Ck Done"
 DestArray(DestRow, 2) = "Interviewed"
Case "BG Ck Req"
 DestArray(DestRow, 2) = "Interviewed"
Case "Intrvw"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Transystems()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Manager-No Inteest"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Offer-Accepted Offer by Another Employer"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-Application not reviewed/considered"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-Does not meet basic qualifications"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-More qualified candidate selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-Other"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-Term of Interest-Salary"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-Unable to Contact Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew-No Show for Interview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew-Other"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew-Term of Interest- Salary"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew-Term of Interest-Position"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew-Term of Interest Location"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Recruiter Reviewed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-Term of Interest-Location"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-Term of Interest-Position"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-Not Eligible for Rehire"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected-Other "
 DestArray(DestRow, 2) = "Apply Completed"
Case "Contact 1st Attempt"
 DestArray(DestRow, 2) = "Qualified"
Case "Contact 2nd Attempt"
 DestArray(DestRow, 2) = "Qualified"
Case "Send Employment Application"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager Interest"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager-Hold Still Under Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Phone Interview Scheduled"
 DestArray(DestRow, 2) = "Qualified"
Case "Rtr Rev'd Resume-Interest"
 DestArray(DestRow, 2) = "Qualified"
Case "Under Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager-No Interest"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager 1st Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Manager 2nd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Accepted (External)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted (Internal)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended (Internal)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended (External)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer In Process"
 DestArray(DestRow, 2) = "Offer Made"
Case "Rejected Offer- Other"
 DestArray(DestRow, 2) = "Offer Made"
Case "Rejected Offer- Vendor Applicant"
 DestArray(DestRow, 2) = "Offer Made"
Case "Rejected Offer-Rescinded by TranSystems"
 DestArray(DestRow, 2) = "Offer Made"
Case "Rejected Offer-Accepted Counter"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hiring"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Rogers()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:H").Delete

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Candidate Withdrew - Dissatisfied with Hours (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Dissatisfied with Location (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Dissatisfied with Pay (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - No Show for Orientation (Auto Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Other (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Other Position Accepted with Different Employer (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Create Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Declined - Candidate Moved to New Req (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Candidate Verbally Dispositioned (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Does Not Meet Qualifications (Auto Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Not a Fit - Assessment (Auto Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Not Reviewed, Auto-Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Previous Employee, Not Eligible (Auto Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Req Cancelled (Auto Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Req Cancelled (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Req Filled (Auto Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Req Filled (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hired (External Only)"
 DestArray(DestRow, 2) = "Hired"
Case "Hired (Internal Only)"
 DestArray(DestRow, 2) = "Hired"
Case "Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "New Application(s)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Application Falsified (No Email)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Criminal History/Credit Verification (No Email)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Onboard"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Short-List"
 DestArray(DestRow, 2) = "Qualified"
Case "Under Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hired On Other Requisition"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Requisition Closed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer Rescinded - Reference Check(s) (No Email)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Onboard (External Only)"
 DestArray(DestRow, 2) = "Offer Made"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Micron()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:B").Delete
Range("G:O").Delete

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Dim LastRow
LastRow = Range("A300000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 300000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "(blank)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hired On Other Requisition"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected - Not willing to travel"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Did not complete the hiring process"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Duplicate candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Lacks required classes or coursework"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Lacks required experience"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Lacks required knowledge"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Lacks required major"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected More qualified candidate selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Not a US Worker"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Not willing to relocate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Other"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Unable to Contact Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Unfavorable performance during previous employment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected Unsatisfactory references"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Requisition Closed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "HM Phone"
 DestArray(DestRow, 2) = "Qualified"
Case "Keep in View"
 DestArray(DestRow, 2) = "Qualified"
Case "Moving to Specific Req"
 DestArray(DestRow, 2) = "Qualified"
Case "Rec Phone"
 DestArray(DestRow, 2) = "Qualified"
Case "Short List 2"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Rejected Failed drug test"
 DestArray(DestRow, 2) = "Interviewed"
Case "Rejected Offer was withdrawn"
 DestArray(DestRow, 2) = "Interviewed"
Case "Rejected Salary expectations are not in line with the role"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - Feels that benefits package is unsatisfactory"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - Feels that cost of living is too high"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - Feels that position is not in line with career goals"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - Feels that relocation package is insufficient"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - Feels that work environment is unsuitable"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - Feels that work hours are unsuitable"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - Not able / not willing to relocate"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - Not able / not willing to travel"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined Accepted counter offer from current employer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined Other"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined Salary below expectations"
 DestArray(DestRow, 2) = "Offer Made"
Case "PreHire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Rejected - Failed background check"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Send to SAP"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Wyndham()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("I:N").Delete

Range("B1").Select
ActiveCell.EntireColumn.Insert

Dim LastRow
LastRow = Range("A400000").End(xlUp).Row

Range("B2:B" & LastRow).Formula = "=C2&D2&E2"

Range("B2:B" & LastRow).Select
Selection.Copy
Range("B2:B" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("C:E").Delete

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 400000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Hiring Mgr ReviewNew ApplicantsNew Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hiring Mgr ReviewNew ApplicantsRecruiter Recommended"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Mgr ReviewInterviewsInterviews Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Mgr ReviewOffer ExtendedAccepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hiring Mgr ReviewOffer ExtendedRefused"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hiring Mgr ReviewOffer ExtendedExtended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hiring Mgr ReviewOffer ExtendedRescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hiring Mgr ReviewHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "Assessment QualifierNew ApplicantsNew Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Assessment QualifierInterviewsInterviews Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Assessment QualifierOffer ExtendedAccepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Assessment QualifierOffer ExtendedRefused"
 DestArray(DestRow, 2) = "Offer Made"
Case "Assessment QualifierOffer ExtendedExtended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Assessment QualifierOffer ExtendedRescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Assessment QualifierHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "Default Prof CSWNew ApplicantsNew Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default Prof CSWNew ApplicantsRecruiter Recommended"
 DestArray(DestRow, 2) = "Qualified"
Case "Default Prof CSWInterviewsInterviews Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Default Prof CSWOffer ExtendedAccepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Default Prof CSWOffer ExtendedRefused"
 DestArray(DestRow, 2) = "Offer Made"
Case "Default Prof CSWOffer ExtendedExtended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Default Prof CSWOffer ExtendedRescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Default Prof CSWHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "ExecutiveNew ApplicantsNew Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "ExecutiveOffer ExtendedAccepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "ExecutiveOffer ExtendedRefused"
 DestArray(DestRow, 2) = "Offer Made"
Case "ExecutiveOffer ExtendedExtended"
 DestArray(DestRow, 2) = "Offer Made"
Case "ExecutiveOffer ExtendedRescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "ExecutiveHireHired"
 DestArray(DestRow, 2) = "Hired"
Case "Hiring Mgr Review WbWNew Applicants WbWNew Applicant WbW"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hiring Mgr Review WbWNew Applicants WbWRecruiter Recommended WbW"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Mgr Review WbWInterviewsInterviews Completed"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Mgr Review WbWOffer ExtendedAccepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hiring Mgr Review WbWOffer ExtendedRefused"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hiring Mgr Review WbWOffer ExtendedExtended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hiring Mgr Review WbWOffer ExtendedRescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hiring Mgr Review WbWHireHired"
 DestArray(DestRow, 2) = "Hired"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Citrix()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("1:3").Delete
Range("A:A").Delete
Range("E:E").Delete
Range("G:J").Delete

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("H2:H" & LastRow).Formula = "=F2&D2"
Range("H2:H" & LastRow).Select
Selection.Copy
Range("H2:H" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

Range("D:D").Delete
Range("E:E").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("C:C").Cut Destination:=Range("A:A")
Range("E:E").Cut Destination:=Range("C:C")
Range("E:E").Delete
Range("B:B").Cut Destination:=Range("G:G")
Range("F:F").Cut Destination:=Range("H:H")

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Americas CollegeBackground - Completed - Alert"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeBackground - Completed - Clear"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeBackground Verification - In Progress"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeCandidate - Starts"
 DestArray(DestRow, 2) = "Hired"
Case "Americas CollegeConduct On-Campus Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeConduct On-Site Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeConduct Phone Interview -"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeCreate Offer Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeCreate Offer Letter"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas CollegeDid Not Pass Screening Questions"
 DestArray(DestRow, 2) = "Qualified"
Case "Americas CollegeHiring Manager Review -"
 DestArray(DestRow, 2) = "Qualified"
Case "Americas CollegeInbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas CollegeNot Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas CollegeNot Selected - College"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas CollegeNot Selected - College Future Interest"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas CollegeNot Selected After Hiring Mgr Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas CollegeOffer Declined -"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas CollegeOffer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas CollegeRecruiter Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Americas CollegeRoute Offer for Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeSchedule On-Site Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeSchedule Phone Interview -"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeSourced - External"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas CollegeStart Date Pending"
 DestArray(DestRow, 2) = "Hired"
Case "Americas WorkflowBackground - Completed - Alert"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowBackground - Completed - Clear"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowBackground Ticketing - In Progress"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowBackground Verification - In Progress"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowCandidate Review- Agency"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas WorkflowCandidate Review- External"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas WorkflowConduct 2nd Phone Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowConduct On-Site Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowConduct Phone/GTM Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowCreate Offer Approval Form"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowCreate Offer Letter"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowDid Not Pass Screening Questions"
 DestArray(DestRow, 2) = "Qualified"
Case "Americas WorkflowHiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Americas WorkflowInbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas WorkflowNot Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas WorkflowNot Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas WorkflowNot Selected After Hiring Mgr Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas WorkflowNot Selected After Interview- Future Interest"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowNot Selected After Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowNot Selected After Phone Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowOffer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowOffer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowOffer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowRecruiter Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Americas WorkflowRoute Offer for Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowSchedule 2nd Phone Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowSchedule On-Site Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowSchedule Phone/GTM Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowSourced - External"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas WorkflowStart Date Pending"
 DestArray(DestRow, 2) = "Hired"
Case "APAC CollegeCandidate - Started"
 DestArray(DestRow, 2) = "Hired"
Case "APAC CollegeOffer Accepted -"
 DestArray(DestRow, 2) = "Offer Made"
Case "APAC CollegeRoute Offer for Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowCandidate Review- External"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC WorkflowCandidate Testing"
 DestArray(DestRow, 2) = "Qualified"
Case "APAC WorkflowConduct 2nd Phone Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowConduct On-Site Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowConduct Phone/GTM Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowDid Not Pass Screening Questions"
 DestArray(DestRow, 2) = "Qualified"
Case "APAC WorkflowHiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "APAC WorkflowInbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC WorkflowNo Show"
 DestArray(DestRow, 2) = "Qualified"
Case "APAC WorkflowNot Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC WorkflowNot Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC WorkflowNot Selected After Hiring Mgr Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC WorkflowNot Selected After Interview- Future Interest"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowNot Selected After Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowNot Selected After Phone Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowNot Selected After Phone Interview."
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowOffer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "APAC WorkflowOffer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "APAC WorkflowRecruiter Phone Screen"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowRoute Offer for Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowSchedule On-Site Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowSchedule Phone/GTM Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowStart Date Pending"
 DestArray(DestRow, 2) = "Hired"
Case "EMEA WorkflowBackground Check"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowCandidate Review- External"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA WorkflowConduct On-Site Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowConduct Phone/GTM Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowCreate Offer Approval Form"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowCreate Offer Letter"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowDid Not Pass Screening Questions"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA WorkflowHiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "EMEA WorkflowInbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA WorkflowMake Verbal Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "EMEA WorkflowNot Qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA WorkflowNot Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA WorkflowNot Selected After Hiring Mgr Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA WorkflowNot Selected After Interview- Future Interest"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowNot Selected After Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowNot Selected After Phone Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowNot Selected After Phone Interview."
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowOffer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "EMEA WorkflowOffer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "EMEA WorkflowRecruiter Phone Screen"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowRoute Offer for Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowSchedule On-Site Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowSchedule Phone/GTM Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "EMEA WorkflowStart Date Pending"
 DestArray(DestRow, 2) = "Hired"
Case "Americas CollegeBackground Check Initiate"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeBackground Ticketing - In Progress"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeConduct Phone or GTM Interview -"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeOffer Accepted -"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas CollegeReviewed – Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas CollegeSchedule 2nd Phone or GTM Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas CollegeSchedule Phone or GTM Interview -"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowBackground Check Initiate"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowCandidate - Started"
 DestArray(DestRow, 2) = "Hired"
Case "Americas WorkflowCandidate Review- Internal"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas WorkflowConduct 2nd Phone or GTM Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Americas WorkflowCreate Internal Transfer Form"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowInitiate Internal Transfer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowInternal Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowMake Verbal Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Americas WorkflowReviewed – Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Americas WorkflowSchedule 2nd Phone or GTM Interview(s)"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC CollegeInbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC CollegeNot Selected - College"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC CollegeNot Selected - College Future Interest"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC CollegeOffer Declined -"
 DestArray(DestRow, 2) = "Offer Made"
Case "APAC CollegeReviewed – Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC WorkflowBackground Checks"
 DestArray(DestRow, 2) = "Interviewed"
Case "APAC WorkflowCandidate - Started"
 DestArray(DestRow, 2) = "Hired"
Case "APAC WorkflowCandidate Review- Internal"
 DestArray(DestRow, 2) = "Apply Completed"
Case "APAC WorkflowCreate Offer Letter"
 DestArray(DestRow, 2) = "Offer Made"
Case "APAC WorkflowOffer Rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "APAC WorkflowReviewed – Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "EMEA WorkflowCandidate - Started"
 DestArray(DestRow, 2) = "Hired"
Case "EMEA WorkflowReviewed – Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pipeline - AmericasConduct Exploratory Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Talent Pipeline - AmericasDid Not Pass Screening Questions"
 DestArray(DestRow, 2) = "Qualified"
Case "Talent Pipeline - AmericasInbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pipeline - AmericasNot Selected after Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Talent Pipeline - AmericasRecruiter Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Talent Pipeline - AmericasReviewed – Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pipeline - AmericasSchedule Exploratory Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Talent Pipeline - AmericasScreening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pipeline - AmericasTalent Pool"
 DestArray(DestRow, 2) = "Qualified"
Case "Talent Pipeline - APACCandidate Review- External"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pipeline - APACCandidate Selected"
 DestArray(DestRow, 2) = "Offer Made"
Case "Talent Pipeline - APACNot Selected After Phone Interview."
 DestArray(DestRow, 2) = "Interviewed"
Case "Talent Pipeline - APACRecruiter Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Talent Pipeline - APACReviewed – Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pipeline - EMEACandidate Review- External"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pipeline - EMEAInbox"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pipeline - EMEANot Selected after Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Talent Pipeline - EMEARecruiter Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Talent Pipeline - EMEAReviewed – Not Selected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Talent Pipeline - EMEASchedule Exploratory Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Talent Pipeline - EMEATalent Pool"
 DestArray(DestRow, 2) = "Qualified"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Bell()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:L").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("D:D").Cut Destination:=Range("A:A")
Range("B:B").Cut Destination:=Range("D:D")
Range("B1").Select
ActiveCell.EntireColumn.Insert
Range("H:H").Cut Destination:=Range("C:C")
Range("D1").Select
ActiveCell.EntireColumn.Insert
Range("F:F").Cut Destination:=Range("D:D")

Dim LastRow
LastRow = Range("A250000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 250000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Availability"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Candidate Accepted Another Position"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - CCS2/SL2/MPPS2/SS2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Does Not Meet Qualifications"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Filled Internally"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Job Closed/Not Filled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Language"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - More Qualified Candidate(s)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - No Show"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Other"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - PI"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Salary Requirements"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - SIM"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Student"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Left Message"
 DestArray(DestRow, 2) = "Apply Completed"
Case "On Hold"
 DestArray(DestRow, 2) = "Interviewed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Assessment Initiated"
 DestArray(DestRow, 2) = "Qualified"
Case "Assessments"
 DestArray(DestRow, 2) = "Qualified"
Case "Assessments Completed"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Pre Assessments Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Screening Shortlist"
 DestArray(DestRow, 2) = "Qualified"
Case "Disqualified - Poor Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview 1"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview 1 / Phone Screen"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview 2"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview 3"
 DestArray(DestRow, 2) = "Interviewed"
Case "Background Checks Completed"
 DestArray(DestRow, 2) = "Offer Made"
Case "Background Checks Initiated"
 DestArray(DestRow, 2) = "Offer Made"
Case "Disqualified - Background Check"
 DestArray(DestRow, 2) = "Offer Made"
Case "Disqualified - Candidate Declined Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Manager Requests Background Checks"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Approval"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired - External"
 DestArray(DestRow, 2) = "Hired"
Case "Hired - Internal"
 DestArray(DestRow, 2) = "Hired"
Case "Internal / Rehire"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - Candidate Declined Offer - Salary"
 DestArray(DestRow, 2) = "Offer Made"
Case "Disqualified - Location"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified - withdrawn (recruiter)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disqualified- Under consideration- internal role"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hired - Agency"
 DestArray(DestRow, 2) = "Hired"
Case "Hired - CAF Vet"
 DestArray(DestRow, 2) = "Hired"
Case "Pre Assessments Initiated"
 DestArray(DestRow, 2) = "Apply Completed"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Seasons()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:L").Delete

Range("A1").Select
ActiveCell.EntireColumn.Insert
Range("C:C").Cut Destination:=Range("A:A")
Range("E:E").Cut Destination:=Range("C:C")
Range("D:D").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("D:D")
Range("G:G").Cut Destination:=Range("E:E")
Range("B:B").Cut Destination:=Range("G:G")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Available"
 DestArray(DestRow, 2) = "Apply Completed"
Case "NEW"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Left Voicemail"
 DestArray(DestRow, 2) = "Qualified"
Case "Sent Email"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "HM Reviewing"
 DestArray(DestRow, 2) = "Qualified"
Case "Interviewing"
 DestArray(DestRow, 2) = "Interviewed"
Case "Decision Point"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment"
 DestArray(DestRow, 2) = "Offer Made"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Ericcson()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:G").Delete

Range("F:F").Cut Destination:=Range("G:G")
Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Assessment"
 DestArray(DestRow, 2) = "Interviewed"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background Check and reference"
 DestArray(DestRow, 2) = "Interviewed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hiring Manager Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview - additional"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Onboarding"
 DestArray(DestRow, 2) = "Hired"
Case "Phone Screening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rehire Check"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected and notified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Release Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "Testing"
 DestArray(DestRow, 2) = "Interviewed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Release Offer Approval"
 DestArray(DestRow, 2) = "Interviewed"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub ConAgra()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:K").Delete

Range("F:F").Cut Destination:=Range("G:G")
Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "2nd Onsite Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "2nd Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "3rd Onsite Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background Check"
 DestArray(DestRow, 2) = "Offer Made"
Case "College Campus Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - Education/Certification Requirements"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - Incomplete Application"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - Location Requirements"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - No Work Authorization"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - Other, please record reason in the comment box"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - Overtime Requirements"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - Salary Requirements"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - Shift or Rotation Requirements"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - Travel Requirements"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Does Not Meet Qualifications - Work Experience/History"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Meets Basic Qualifications"
 DestArray(DestRow, 2) = "Qualified"
Case "No Further Action Taken - Candidate failed pre-employment test"
 DestArray(DestRow, 2) = "Interviewed"
Case "No Further Action Taken - Candidate failed to appear for scheduled interview."
 DestArray(DestRow, 2) = "Qualified"
Case "No Further Action Taken - Candidate failed to complete application."
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Further Action Taken - Candidate former temporary employee, but poor  performance or attendance."
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Further Action Taken - Candidate meets basic qualifications, but a more qualified available."
 DestArray(DestRow, 2) = "Qualified"
Case "No Further Action Taken - Candidate not eligible for rehire."
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Further Action Taken - Candidate rejected due to poor work history."
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Further Action Taken - Other, please record reason in the comment box"
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Further Action Taken - Resume Never Viewed."
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Further Action Taken - Unable to contact candidate."
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined - Â due to inability to work full-time/part-time."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined â€“ accepted another position"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined â€“ due to benefits."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined â€“ due to job requirements (i.e. overtime, travel)."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined â€“ due to location."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined â€“ due to salary."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined â€“ due to shift or rotation."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined â€“ no reason given."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded â€“ due to applicant failed to report on start date."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded â€“ due to application or resume proven false."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer rescinded â€“ due to negative results on background report."
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded â€“ due to positive results on drug test."
 DestArray(DestRow, 2) = "Offer Made"
Case "Onsite Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Ready to Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Requisition Closed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew - accepted another position"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew - cannot meet shift or rotation"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew - due to job requirements (i.e., overtime, travel)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew - due to location"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew - due to salary"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew - no longer interested in employment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew - no reason given"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrew - unable to work full-time/part-time"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined DPCS"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Admin"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer Rescinded â€“ due to failed to pass pre-employment/post offer physical exam."
 DestArray(DestRow, 2) = "Offer Made"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Kroger()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("F:F").Cut Destination:=Range("G:G")
Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Qualified"
Case "Candidate Withdrew - Dissatisfied with Hours"
 DestArray(DestRow, 2) = "Qualified"
Case "Candidate Withdrew - Dissatisfied with Location"
 DestArray(DestRow, 2) = "Qualified"
Case "Candidate Withdrew - Dissatisfied with Pay"
 DestArray(DestRow, 2) = "Qualified"
Case "Candidate Withdrew - No Call No Show on Start Date (Auto Emails)"
 DestArray(DestRow, 2) = "Hired"
Case "Candidate Withdrew - No Show for Interview (Auto Emails)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Other"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Other Position Accepted with Different Employer"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Other Position Accepted with Kroger"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Create Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Declined - Interviewed, Not Proceeding (Auto Emails)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Declined - Interviewed, Not Proceeding (No Auto Email)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Declined - Not Reviewed, Not Proceeding (Auto Emails)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Requisition Cancelled (Auto Emails)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Reviewed, Not Proceeding (Auto Emails)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Reviewed, Not Proceeding (No Auto Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Unable to Contact Applicant (Auto Emails)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Interviewing"
 DestArray(DestRow, 2) = "Interviewed"
Case "Manager Reviewing"
 DestArray(DestRow, 2) = "Qualified"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended (Internal Candidate)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Application Falsified"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Criminal History Verification"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Did not pass Drug Test"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Did not pass Motor Vehicle Report"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Education History Verification"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Employment History Verification"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - SSN Verification"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Unable to Produce Work Authorization Documentation"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Board"
 DestArray(DestRow, 2) = "Hired"
Case "Pre-Screening"
 DestArray(DestRow, 2) = "Qualified"
Case "Requisition Closed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reviewing"
 DestArray(DestRow, 2) = "Qualified"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background Check Pending"
 DestArray(DestRow, 2) = "Offer Made"
Case "Candidate Withdrew - Dissatisfied with Benefits"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Phone Screen, Not Proceeding (Auto Emails)"
 DestArray(DestRow, 2) = "Qualified"
Case "Declined - Phone Screen, Not Proceeding (No Auto Email)"
 DestArray(DestRow, 2) = "Qualified"
Case "Hire Completed"
 DestArray(DestRow, 2) = "Hired"
Case "Interview Scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed, Back-Up Candidate"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed, Decline"
 DestArray(DestRow, 2) = "Interviewed"
Case "New Application(s)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Reviewed"
 DestArray(DestRow, 2) = "Qualified"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Trizetto()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("A:B").Delete
Range("G:G").Delete

Range("F:F").Cut Destination:=Range("G:G")
Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "TA Review"
 DestArray(DestRow, 2) = "Qualified"
Case "TA Scheduling"
 DestArray(DestRow, 2) = "Qualified"
Case "TA Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "HM Review"
 DestArray(DestRow, 2) = "Interviewed"
Case "HM Phone"
 DestArray(DestRow, 2) = "Interviewed"
Case "HM F2F"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Request"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Approved / Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Onboarding"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Reject - Another Candidate Seleted"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject - Does Not Meet Basic Qualifications"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject - Offer in Progress Other Applicant"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject - Requisition Closed without Filling"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "HM â€“ Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Approved"
 DestArray(DestRow, 2) = "Offer Made"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub
Sub Tesoro()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:K").Delete

Range("F:F").Cut Destination:=Range("G:G")
Range("B:B").Cut Destination:=Range("H:H")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Background"
 DestArray(DestRow, 2) = "Offer Made"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Extend Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Future Consideration"
 DestArray(DestRow, 2) = "Qualified"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hired On Other Requisition"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "No Further Interest After Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Requisition Closed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Screening"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Testing"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Unconsidered Email Sent"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Unconsidered No Email Sent"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Panalpina()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("D:D").Delete
Range("F:F").Delete

Range("D:D").Cut Destination:=Range("H:H")
Range("A:A").Cut Destination:=Range("D:D")
Range("C:C").Cut Destination:=Range("A:A")
Range("E:E").Cut Destination:=Range("C:C")
Range("B:B").Cut Destination:=Range("E:E")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Job Cancelled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Short List"
 DestArray(DestRow, 2) = "Qualified"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Assessment"
 DestArray(DestRow, 2) = "Interviewed"
Case "Background check"
 DestArray(DestRow, 2) = "Interviewed"
Case "Pre-Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "2nd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Onsite Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Phone Interview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Background Check"
 DestArray(DestRow, 2) = "Interviewed"
Case "Rejected & notified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Requisition Closed"
 DestArray(DestRow, 2) = "Apply Completed"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Sandisk()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:K").Delete

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Disposition"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Interview Request"
 DestArray(DestRow, 2) = "Qualified"
Case "Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Background Check / Offer Prep"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Submitted to SAP"
 DestArray(DestRow, 2) = "Hired"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub YellowPages()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 65536, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Accepted another position - external"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Accepted another position - internal"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Did not meet basic qualifications"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Failed Background check"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "No call/no show"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not most qualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not most qualified - experience"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not most qualified - skills"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Onsite Interview 1"
 DestArray(DestRow, 2) = "Interviewed"
Case "Onsite Interview 2"
 DestArray(DestRow, 2) = "Interviewed"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Qualified"
Case "Ready to Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Requisition Closed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Selected Another Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Unable to contact"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Not Interested"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Do Not Hire Record"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Manager Review"
 DestArray(DestRow, 2) = "Qualified"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Carolina()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:K").Delete

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Dim LastRow
LastRow = Range("A200000").End(xlUp).Row

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 200000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)
                
Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "(blank)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hired for Another Position"
 DestArray(DestRow, 2) = "Apply Completed"
Case "HR Phone Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Manager Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "ND (No Decision) - Not Reviewed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "ND (No Decision), Ineligible Employee"
 DestArray(DestRow, 2) = "Apply Completed"
Case "ND (No Decision), Position Cancelled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "No Show for Interview"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Not Qualified - based on not possessing basic qualifications"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Resume - Incomplete/Unprofessional"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Salary Expectation"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Short List (Data Management Technique)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn - No Response/Incomplete Contact Info"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Check References"
 DestArray(DestRow, 2) = "Interviewed"
Case "Due Diligence"
 DestArray(DestRow, 2) = "Interviewed"
Case "Not Qualified - based on due diligence"
 DestArray(DestRow, 2) = "Interviewed"
Case "Initiate Background"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Less Qualified than other applicant who was hired"
 DestArray(DestRow, 2) = "Interviewed"
Case "Provider Services Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Declined Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer/Health Assessment"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hired On Other Requisition"
 DestArray(DestRow, 2) = "Apply Completed"
End Select
                               
DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"

Range("I2:I" & LastRow).Formula = Range("C2:C" & LastRow).Value2
Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues

With Range("I2:I" & LastRow)
    .Value = Evaluate("IF(ROW(" & .Address & "),ROUNDDOWN(" & .Address & ",0))")
    .NumberFormat = "0"
End With

Range("I2:I" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("I2:I" & LastRow).Delete
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic
Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Burberry()
Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:K").Delete

Range("B:B").Cut Destination:=Range("H:H")
Range("C:C").Cut Destination:=Range("I:I")
Range("B:B").Delete

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A150000").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Range("B2:B" & LastRow).Formula = "=IF(ISERROR(YEAR(C2)),DATE(RIGHT(C2,4),MID(C2,4,2),LEFT(C2,2)),DATE(YEAR(C2),DAY(C2),MONTH(C2)))"
Range("B2:B" & LastRow).Select
Selection.Copy
Range("C2:C" & LastRow).Select
Selection.PasteSpecial Paste:=xlPasteValues
Range("B2:B" & LastRow).ClearContents
Range("C2:C" & LastRow).NumberFormat = "mm-dd-yyyy"

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 150000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Forwarded"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Invite to Apply"
 DestArray(DestRow, 2) = "Apply Completed"
Case "New"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Line Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Please Invite for Interview"
 DestArray(DestRow, 2) = "Qualified"
Case "Phone Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Assessment Centre"
 DestArray(DestRow, 2) = "Qualified"
Case "F2F Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Pre-Employment Checks"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Approval"
 DestArray(DestRow, 2) = "Interviewed"
Case "Under Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Accepted/Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Pending Reject"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject - Post Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject - Post Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Reject - Role Filled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reject - No Email"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Application Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted on Demand by Admin"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted on Demand by Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined DPCS"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Admin"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub MarathonOil()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:I").Delete

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A150000").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 150000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Lacks Necessary Skills/Experience"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hiring Manager Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "HR Rejected"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Selected"
 DestArray(DestRow, 2) = "Qualified"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Recruiter Review"
 DestArray(DestRow, 2) = "Qualified"
Case "External Candidate Interview Scheduled"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interview Requested"
 DestArray(DestRow, 2) = "Interviewed"
Case "Phone Interview Requested"
 DestArray(DestRow, 2) = "Interviewed"
Case "External Candidate Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Internal Candidate Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Employment"
 DestArray(DestRow, 2) = "Offer Made"
Case "Pre-Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Verbal Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub RobertHalf()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("G:L").Delete

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Range("A1").Select
ActiveCell.EntireRow.Insert

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A65536").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 150000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Lacks skills/qualifications"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reviewed"
 DestArray(DestRow, 2) = "Qualified"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Send for Hire"
 DestArray(DestRow, 2) = "Offer Made"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Rejected - Other"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Contact Attempted"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Insufficient Job Related Knowledge"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Location"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Salary Expectations"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Present"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Temp/Temp to Hire Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Other Internal Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Forwarded to Other Requisition"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Another Applicant Hired"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Assessment Pending"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Assessment Completed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer in Progress"
 DestArray(DestRow, 2) = "Offer Made"
Case "Contractor/Temp Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Lacks Work Experience"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Lacks Education"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Benefit Expectations"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Send 2nd Stage App"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Misrepresentation"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Accepted Another Offer"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Poor Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Position Cancelled/Put on Hold"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Interview Process"
 DestArray(DestRow, 2) = "Do Not Load"
Case "No Offer - Accepted Another Job"
 DestArray(DestRow, 2) = "Do Not Load"
Case "No Offer - Conflict of Interest"
 DestArray(DestRow, 2) = "Do Not Load"
Case "No Offer - Contractor Candidate"
 DestArray(DestRow, 2) = "Do Not Load"
Case "No Offer - Job Cancelled"
 DestArray(DestRow, 2) = "Do Not Load"
Case "No Offer - Lacks Education"
 DestArray(DestRow, 2) = "Do Not Load"
Case "No Offer - Lacks Work Experience"
 DestArray(DestRow, 2) = "Do Not Load"
Case "No Offer - Misrepresentation"
 DestArray(DestRow, 2) = "Do Not Load"
Case "Offer Process"
 DestArray(DestRow, 2) = "Do Not Load"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Do Not Load"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Do Not Load"
Case "Phone Interview"
 DestArray(DestRow, 2) = "Do Not Load"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Amtrak()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:B").Cut Destination:=Range("H:H")
Range("F:F").Cut Destination:=Range("G:G")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A100000").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "Auto Disqualified"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Dissatisfied with Hours (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Dissatisfied with Location (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Dissatisfied with Pay (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - No Show for Assessment (Send Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - No Show for First Day of Employment (No Email)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Candidate Withdrew - No Show for Interview (Send Email)"
 DestArray(DestRow, 2) = "Qualified"
Case "Candidate Withdrew - Other (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Other Position Accepted with Different Employer (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Candidate Withdrew - Other Position Accepted within our Company (No Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Create Offer"
 DestArray(DestRow, 2) = "Interviewed"
Case "Declined - Application Received After Job Closed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Data Management Technique (Num Limits)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Ineligible for Re-hire Prior Employee"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Interviewed, Not Proceeding (No Email)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Declined - Interviewed, Not Proceeding (Send Email)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Declined - Not Reviewed, Not Proceeding (Send Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Position on Hold/Eliminated/Canceled"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Reviewed, Not Proceeding (Send Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Unable to Contact Applicant (Send Email)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Declined - Unsatisfactory Performance"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Default"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Deleted On Demand By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "HiredAtSAP"
 DestArray(DestRow, 2) = "Hired"
Case "Interviewing"
 DestArray(DestRow, 2) = "Interviewed"
Case "Invite to Additional Assessments"
 DestArray(DestRow, 2) = "Qualified"
Case "Invite to IPCS"
 DestArray(DestRow, 2) = "Qualified"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Approved"
 DestArray(DestRow, 2) = "Interviewed"
Case "Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Criminal History Verification (No Email)"
 DestArray(DestRow, 2) = "Offer Made"
Case "Requisition Closed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Screening"
 DestArray(DestRow, 2) = "Qualified"
Case "SendToSAP"
 DestArray(DestRow, 2) = "Offer Made"
Case "TransferredToSAP"
 DestArray(DestRow, 2) = "Offer Made"
Case "TransferredToSAPError"
 DestArray(DestRow, 2) = "Offer Made"
Case "Withdrawn By Candidate"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer Rescinded - Other"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer Rescinded - Did not pass Physcial (No Email)"
 DestArray(DestRow, 2) = "Offer Made"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

Sub Fitness()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Range("B:C").Delete
Range("C:C").Delete
Range("D:D").Delete
Range("E:H").Delete
Range("G:K").Delete

Range("A:A").Cut Destination:=Range("G:G")
Range("B:B").Cut Destination:=Range("A:A")
Range("E:E").Cut Destination:=Range("H:H")
Range("D:D").Cut Destination:=Range("E:E")
Range("C:C").Cut Destination:=Range("D:D")
Range("F:F").Cut Destination:=Range("C:C")

Range("A1").Value = "Email"
Range("B1").Value = "Status"
Range("C1").Value = "Date"
Range("D1").Value = "JobID1"
Range("E1").Value = "Title"
Range("F1").Value = "J2wMemberID"
Range("G1").Value = "ClientApplicantID"
Range("H1").Value = "Original Status"

Dim LastRow
LastRow = Range("A100000").End(xlUp).Row

Range("A1:H" & LastRow).Font.Size = 10
Range("A1:H" & LastRow).Font.Name = "Arial"
Range("A1:H1").Font.Color = vbBlack
Range("A1:H1").Font.Bold = True
Range("A1:H1").Interior.Color = vbYellow

Dim CurRow
CurRow = 2

Dim DestRow
DestRow = 2

Dim i As Long
i = 1

Dim SourceArray As Variant
SourceArray = Sheets(1).Range("A1:H" & LastRow)

Dim DestArray As Variant
ReDim DestArray(1 To 100000, 1 To 8)

    DestArray(1, 1) = SourceArray(1, 1)
    DestArray(1, 2) = SourceArray(1, 2)
    DestArray(1, 3) = SourceArray(1, 3)
    DestArray(1, 4) = SourceArray(1, 4)
    DestArray(1, 5) = SourceArray(1, 5)
    DestArray(1, 6) = SourceArray(1, 6)
    DestArray(1, 7) = SourceArray(1, 7)
    DestArray(1, 8) = SourceArray(1, 8)
    
For CurRow = 2 To LastRow
                                              
                DestArray(DestRow, 1) = SourceArray(CurRow, 1)
                DestArray(DestRow, 2) = SourceArray(CurRow, 2)
                DestArray(DestRow, 3) = SourceArray(CurRow, 3)
                DestArray(DestRow, 4) = SourceArray(CurRow, 4)
                DestArray(DestRow, 5) = SourceArray(CurRow, 5)
                DestArray(DestRow, 6) = SourceArray(CurRow, 6)
                DestArray(DestRow, 7) = SourceArray(CurRow, 7)
                DestArray(DestRow, 8) = SourceArray(CurRow, 8)

Dim OriginalStatus
OriginalStatus = SourceArray(CurRow, 8)

Select Case OriginalStatus
Case "0-Filed"
 DestArray(DestRow, 2) = "Apply Completed"
Case "2nd Interview"
 DestArray(DestRow, 2) = "Interviewed"
Case "Candidate Review"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Corp / Field Mgt Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Corp/Field Mgt Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "Field Offer"
 DestArray(DestRow, 2) = "Offer Made"
Case "Field Offer Extended"
 DestArray(DestRow, 2) = "Offer Made"
Case "HIDE"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hired"
 DestArray(DestRow, 2) = "Hired"
Case "Hired (other req)"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Hiring Manager Review"
 DestArray(DestRow, 2) = "Qualified"
Case "Incomplete Assessment"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Interview (face to face)"
 DestArray(DestRow, 2) = "Interviewed"
Case "Interviewed Not Interested"
 DestArray(DestRow, 2) = "Interviewed"
Case "Left Message 2"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Left Message 3"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Left Message/Attempted Contact"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Offer Accepted"
 DestArray(DestRow, 2) = "Hired"
Case "Offer Declined"
 DestArray(DestRow, 2) = "Offer Made"
Case "Offer rescinded"
 DestArray(DestRow, 2) = "Offer Made"
Case "Phone Screen"
 DestArray(DestRow, 2) = "Apply Completed"
Case "Reviewed Not Interested"
 DestArray(DestRow, 2) = "Apply Completed"
End Select

DestRow = DestRow + 1
               
Next CurRow

Sheets(1).Range("A1:H" & DestRow).Value = DestArray

Range("A1:H" & LastRow).Borders.Weight = xlThin
Range("A1:H" & LastRow).Borders.ColorIndex = xlAutomatic

ActiveSheet.UsedRange.Columns.AutoFit
Sheets(1).Name = "Sheet1"

Dim ws As Worksheet
For Each ws In Sheets
Application.DisplayAlerts = False
If ws.Name <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True

Range("I2:I" & LastRow).Formula = "=IF(DATE(YEAR(C2),MONTH(C2),DAY(C2))>DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY())+90),1,0)"
Range("I1").Formula = "=SUM(I2:I" & LastRow & ")"

Dim FutureDate As Integer
FutureDate = Range("I1").Value

Range("I:I").Delete

If FutureDate > 0 Then
MsgBox "This file contains records with dates greater than 90 days from today's date."
Range("A2:H" & LastRow).Sort Key1:=Range("C2:C" & LastRow), order1:=xlDescending
Else
End If

Application.StatusBar = False
Application.ScreenUpdating = True

Range("A1").Select

Dim StatusCount
StatusCount = WorksheetFunction.CountA(Range("B2:B" & LastRow))

If StatusCount <> (LastRow - 1) Then
MsgBox "Some records contained a status not accounted for in mapping. Please manually update these records and update the mapping logic in the VBA code."
Else
End If
    
End Sub

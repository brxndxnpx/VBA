Option Explicit
Option Private Module

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     Environment functions pertaining to the user and the user's machine information.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary
'''     Combines paths by utilizing Scripting.FileSystemObject.
Public Function PathCombine(ParamArray Paths() As Variant) As String
    Dim FS     As Object: Set FS = CreateObject("Scripting.FileSystemObject")
    Dim i      As Long, output As String
    
    For i = LBound(Paths) To UBound(Paths)
        If Paths(i) <> vbNullString Then output = FS.BuildPath(output, Paths(i))
    Next
    PathCombine = output
End Function

''' Summary:
'''     The user's desktop path.
Public Function Desktop() As String: Desktop = PathCombine(IIf(OneDrive <> vbNullString, OneDrive, UserProfile), "Desktop"): End Function

''' Summary:
'''     The user's documents path.
Public Function Documents() As String: Documents = PathCombine(IIf(OneDrive <> vbNullString, OneDrive, UserProfile), "Documents"): End Function

''' Summary:
'''     The user's downloads path.
Public Function Downloads() As String: Downloads = PathCombine(IIf(OneDrive <> vbNullString, OneDrive, UserProfile), "Downloads"): End Function

''' Summary:
'''     The user's profile path.
Public Function UserProfile() As String: UserProfile = Environ$("UserProfile"): End Function

''' Summary:
'''     The user's OneDrive.
Public Function OneDrive() As String: OneDrive = Environ$("OneDrive"): End Function

''' Summary:
'''     The user's temporary files path.
Public Function Temp() As String: Temp = Environ$("TEMP"): End Function

''' Summary:
'''     The user's application data path.
Public Function AppData() As String: AppData = Environ$("APPDATA"): End Function

''' Summary:
'''     The user's home path.
Public Function HomePath() As String: HomePath = Environ$("HOMEPATH"): End Function

''' Summary:
'''     The window's root path.
Public Function SystemRoot() As String: SystemRoot = Environ$("SYSTEMROOT"): End Function

''' Summary:
'''     The user's 32x program files.
Public Function ProgramFiles32() As String: ProgramFiles32 = Environ$("PROGRAMFILES"): End Function

''' Summary:
'''     The user's 64x program files.
Public Function ProgramFiles64() As String: ProgramFiles64 = Environ$("PROGRAMFILES(X86)"): End Function

''' Summary:
'''     Excel's default library path.
Public Function ExcelLibraryPath() As String: ExcelLibraryPath = Application.LibraryPath: End Function

''' Summary:
'''     Excel's default user library path.
Public Function ExcelUserLibraryPath() As String: ExcelUserLibraryPath = Application.UserLibraryPath: End Function

''' Summary:
'''     The CPU/Processor info - whether the user uses 32/64 bit.
Public Function CPUProcessor() As String: CPUProcessor = Environ$("PROCESSOR_IDENTIFIER"): End Function

''' Summary:
'''     Excel's default startup path. This is where your PERSONAL.XLSB file is stored.
Public Function ExcelStartupPath() As String: ExcelStartupPath = Application.StartupPath: End Function

''' Summary:
'''     Microsoft office Ribbon path.
Public Function OfficeRibbonPath() As String: OfficeRibbonPath = Environ$("LocalAppData") & "\Microsoft\Office\": End Function

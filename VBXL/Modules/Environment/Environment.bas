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
Public Function DesktopPath() As String: DesktopPath = PathCombine(IIf(OneDrivePath <> vbNullString, OneDrivePath, UserProfilePath), "Desktop"): End Function

''' Summary:
'''     The user's documents path.
Public Function DocumentsPath() As String: DocumentsPath = PathCombine(IIf(OneDrivePath <> vbNullString, OneDrivePath, UserProfilePath), "Documents"): End Function

''' Summary:
'''     The user's downloads path.
Public Function DownloadsPath() As String: DownloadsPath = PathCombine(IIf(OneDrivePath <> vbNullString, OneDrivePath, UserProfilePath), "Downloads"): End Function

''' Summary:
'''     The user's profile path.
Public Function UserProfilePath() As String: UserProfilePath = Environ$("UserProfile"): End Function

''' Summary:
'''     The user's OneDrive.
Public Function OneDrivePath() As String: OneDrivePath = Environ$("OneDrive"): End Function

''' Summary:
'''     The user's temporary files path.
Public Function TempPath() As String: TempPath = Environ$("TEMP"): End Function

''' Summary:
'''     The user's application data path.
Public Function AppDataPath() As String: AppDataPath = Environ$("APPDATA"): End Function

''' Summary:
'''     The user's home path. This is the same as UserProfilePath but without the drive letter.
Public Function HomePath() As String: HomePath = Environ$("HOMEPATH"): End Function

''' Summary:
'''     The window's root path.
Public Function SystemRootPath() As String: SystemRootPath = Environ$("SYSTEMROOT"): End Function

''' Summary:
'''     The user's 32x program files.
Public Function ProgramFiles32Path() As String: ProgramFiles32Path = Environ$("PROGRAMFILES"): End Function

''' Summary:
'''     The user's 64x program files.
Public Function ProgramFiles64Path() As String: ProgramFiles64Path = Environ$("PROGRAMFILES(X86)"): End Function

''' Summary:
'''     Excel's default library path.
Public Function ExcelLibraryPath() As String: ExcelLibraryPath = Application.LibraryPath: End Function

''' Summary:
'''     Excel's default user library path.
Public Function ExcelUserLibraryPath() As String: ExcelUserLibraryPath = Application.UserLibraryPath: End Function

''' Summary:
'''     Excel's default startup path. This is where your PERSONAL.XLSB file is stored.
Public Function ExcelStartupPath() As String: ExcelStartupPath = Application.StartupPath: End Function

''' Summary:
'''     Microsoft office Ribbon path.
Public Function OfficeRibbonPath() As String: OfficeRibbonPath = Environ$("LocalAppData") & "\Microsoft\Office\": End Function

''' Summary:
'''     The CPU/Processor info - whether the user uses 32/64 bit.
Public Function CPUProcessorInfo() As String: CPUProcessorInfo = Environ$("PROCESSOR_IDENTIFIER"): End Function

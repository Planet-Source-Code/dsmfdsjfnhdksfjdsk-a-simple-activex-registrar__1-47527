VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBRegSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CREATE_SUSPENDED = &H4
Private Const INFINITE = &HFFFFFFFF
Private Const STATUS_WAIT_0 = &H0
Private Const STATUS_ABANDONED_WAIT_0 = &H80
Private Const STATUS_TIMEOUT = &H102
Private Const WAIT_FAILED = &HFFFFFFFF
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Private Const WAIT_ABANDONED = ((STATUS_ABANDONED_WAIT_0) + 0)
Private Const WAIT_TIMEOUT = STATUS_TIMEOUT
Private Const STATUS_PENDING = &H103
Private Const STILL_ACTIVE = STATUS_PENDING

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
        (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
        (ByVal hLibModule As Long) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
        (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function CreateThread Lib "kernel32" _
        (lpThreadAttributes As Any, ByVal dwStackSize As Long, _
        lpStartAddress As Long, lpParameter As Any, _
        ByVal dwCreationFlags As Long, lpThreadID As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function ResumeThread Lib "kernel32" _
        (ByVal hThread As Long) As Long

Private Declare Function GetExitCodeThread Lib "kernel32" _
        (ByVal hThread As Long, lpExitCode As Long) As Long

Public Function fVBRegServer(ByVal strFilePath As String, _
                Optional ByVal blnRegister = True) As Boolean
                
Dim lngModuleHandle As Long
Dim lngFunctionAdr  As Long
Dim lngThreadID     As Long
Dim lngThreadHandle As Long
Dim lngExitCode     As Long
Dim blnSuccess      As Boolean

    lngModuleHandle = LoadLibrary(strFilePath)
      
    If blnRegister Then
        lngFunctionAdr = GetProcAddress(lngModuleHandle, "DllRegisterServer")
    Else
        lngFunctionAdr = GetProcAddress(lngModuleHandle, "DllUnregisterServer")
    End If
    
    If lngFunctionAdr <> 0 Then
    
        lngThreadHandle = CreateThread(ByVal 0, 0, ByVal lngFunctionAdr, ByVal 0, 0, lngThreadID)
        
        If lngThreadHandle Then
        
            blnSuccess = (WaitForSingleObject(lngThreadHandle, 10000) = WAIT_OBJECT_0)
          
            If Not blnSuccess Then
            
                Call GetExitCodeThread(lngThreadHandle, lngExitCode)
                Call ExitThread(lngExitCode)
            End If
    
            Call CloseHandle(lngThreadHandle)
        End If
    End If
    
    If lngModuleHandle Then Call FreeLibrary(lngModuleHandle)
    
    fVBRegServer = blnSuccess
    
End Function

Public Function IsDLLActiveX(ByVal strDLLPath As String, Optional ByVal RaiseError As Boolean) As Boolean
Dim lngHMod         As Long
Dim lngLastDllError As Long
  
    lngHMod = LoadLibrary(strDLLPath)
    
    If lngHMod = 0 Then
        If RaiseError Then
            lngLastDllError = Err.LastDllError
            Err.Raise 10000 + lngLastDllError, "IsDLLActiveX", "LoadLibrary-Error: " & lngLastDllError
        End If
    End If
  
    IsDLLActiveX = Abs(CBool(GetProcAddress(lngHMod, "DllRegisterServer")))
    Call FreeLibrary(lngHMod)
End Function


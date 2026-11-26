Attribute VB_Name = "Common"
Option Explicit

'==============================================================================
'        modWinsock2 - Pure VB6 replacement for MSWinsockLib.Winsock
'==============================================================================
'
' O    O    O
'  \__/ \__/
'  /=||=||=\   oouuuunnnnnnnnmmmmmmmmmmmmmm\
' // ||_||                WinsockVB6        \
' \\ /\ #\     oouuuunnnnnnnnmmmmmmmmmmmmmmmm\
' /=(  \  )==>       by EBArtSoft@ + AI       \
'//  \O_\/         Copyright © 2006 - 2026     \
'\\  || ||        https://www.ebartsoft.com     \
' \==||=||==/  oouuuunnnnnnnnmmmmmmmmmmmmmmmmmmmm\
' ===========
'==== E.B ====
'
'=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!
'
'                            AI ASSISTANCE
'
'                  OpenCode - DeepSeek V4.1 Flash 1017
'
'=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!
'
' MIT License
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
'

Private mLockedObject As Collection

'================================================
' Main
'================================================
Public Sub Main()
    #If DEVMODE Then
        DevLog "Main"
    #End If
End Sub

'================================================
' New_WinsockVB6
'================================================
Public Function New_WinsockVB6() As Winsock
    #If DEVMODE Then
        DevLog "New_WinsockVB6"
    #End If
    StartVBRuntime
    Set New_WinsockVB6 = New_Winsock2()
End Function

'================================================
' DevLog
'================================================
Private Sub DevLog(ByRef v As String)
    MessageBox 0, v, vbNullString, vbExclamation
End Sub

'================================================
' DllGetClassObject_
'================================================
Private Function DllGetClassObject_(ByRef a As CLSID, ByRef b As CLSID, ByRef c As IClassFactory) As Long
    Rnd
End Function

'================================================
' AdrsOf
'================================================
Private Function AdrsOf(ByVal pAddressOf As Long) As Long
    AdrsOf = pAddressOf
End Function

'====================================================
' RtSetProcAddress
'====================================================
Private Sub RtSetProcAddress(ByVal pSrc As Long, ByVal pNew As Long)
    Dim vPro As Long
    If (pSrc = 0) Or (pNew = 0) Then Exit Sub
    VirtualProtect ByVal pSrc, 5, 64, vPro
    MemoryCopy ByVal pSrc, 233&, 1
    MemoryCopy ByVal pSrc + 1, pNew - pSrc - 5, 4
    VirtualProtect ByVal pSrc, 5, vPro, 0
End Sub

'================================================
' GetDllHandle
'================================================
Private Function GetDllHandle() As Long
    Dim e As MEMORY_BASIC_INFORMATION
    VirtualQuery ByVal AdrsOf(AddressOf DllGetClassObject_), e, Len(e)
    GetDllHandle = e.AllocationBase
End Function

'================================================
' StartVBRuntime
'================================================
Public Sub StartVBRuntime()
    Dim f                   As IClassFactory
    Dim vIClassFactory      As CLSID
    Dim vIDummy             As CLSID
    Dim vDllGetClassObject  As Long
    Dim vModuleHandle       As Long
    
    If mLockedObject Is Nothing Then
        
        #If DEVMODE Then
            DevLog "StartVBRuntime"
        #End If
        
        CoInitialize 0
        
        With vIClassFactory '{00000001-0000-0000-C000-000000000046}
            .Data1 = &H1&
            .Data4(0) = &HC0&
            .Data4(7) = &H46&
        End With
    
        vModuleHandle = GetDllHandle()
        If vModuleHandle <> 0 Then
            vDllGetClassObject = GetProcAddress(vModuleHandle, "DllGetClassObject")
            If vDllGetClassObject <> 0 Then
                RtSetProcAddress AddressOf DllGetClassObject_, vDllGetClassObject
                If DllGetClassObject_(vIDummy, vIClassFactory, f) = 0 Then
                    Set mLockedObject = New Collection
                Else
                    ' It will fail but we don't care runtime is already initialized
                    Set mLockedObject = New Collection
                    #If DEVMODE Then
                        DevLog "DllGetClassObject() Failed"
                    #End If
                End If
            Else
                #If DEVMODE Then
                    DevLog "GetProcAddress() Failed"
                #End If
            End If
        Else
            #If DEVMODE Then
                DevLog "GetDllHandle() Failed"
            #End If
        End If
    
    End If
    
End Sub

''================================================
'' StartVBRuntime
''================================================
'Public Sub StartVBRuntime()
'    Dim f                   As IClassFactory
'    Dim vIClassFactory      As CLSID
'    Dim vIDispatch          As CLSID
'    Dim vIDummy             As CLSID
'    Dim vDllGetClassObject  As Long
'    Dim vModuleHandle       As Long
'
'    If mLockedObject Is Nothing Then
'
'        ''MessageBox 0, "StartVBRuntime() [Start]", vbNullString, vbExclamation
'
'        CoInitialize 0
'
''        ' IDummy
''        With vIDummy '{EBA750F7-0000-0000-0000-00008E35DD3B}
''            .Data1 = &HEBA750F7
''            .Data4(4) = &H8E&
''            .Data4(5) = &H35&
''            .Data4(6) = &HDD&
''            .Data4(7) = &H3B&
''        End With
'
'        ' IClassFactory
'        With vIClassFactory '{00000001-0000-0000-C000-000000000046}
'            .Data1 = &H1&
'            .Data4(0) = &HC0&
'            .Data4(7) = &H46&
'        End With
'
'        ' IDIspatch
'        With vIDispatch '{00020400-0000-0000-C000-000000000046}
'            .Data1 = &H20400
'            .Data4(0) = &HC0&
'            .Data4(7) = &H46&
'        End With
'
'        vModuleHandle = GetDllHandle()
'        If vModuleHandle <> 0 Then
'            'If GetDllDummyCLSID(mhInstance, vIDummy) Then
'                vDllGetClassObject = GetProcAddress(vModuleHandle, "DllGetClassObject")
'                If vDllGetClassObject Then
'                    RtSetProcAddress AddressOf DllGetClassObject_, vDllGetClassObject
'                    ''MessageBox 0, "Calling DllGetClassObject()", vbNullString, vbInformation
'                    If DllGetClassObject_(vIDummy, vIClassFactory, f) = 0 Then
'                        If f.CreateInstance(Nothing, vIDispatch, mLockedObject) = 0 Then
'                            'MessageBox 0, "CreateInstance() Success", vbNullString, vbInformation
'                        Else
'                            'MessageBox 0, "CreateInstance() Failed", vbNullString, vbInformation
'                        End If
'                    Else
'                        ''MessageBox 0, "DllGetClassObject() Failed", vbNullString, vbCritical
'                        Set mLockedObject = New Dummy
'                    End If
'                Else
'                    ''MessageBox 0, "GetProcAddress() Failed", vbNullString, vbCritical
'                End If
'            'Else
'            '    MessageBox 0, "GetDllDummyCLSID() Failed", vbNullString, vbCritical
'            'End If
'        Else
'            ''MessageBox 0, "GetDllHandle() Failed", vbNullString, vbCritical
'        End If
'
'        ''MessageBox 0, "StartVBRuntime() [End]", vbNullString, vbExclamation
'    Else
'        ''MessageBox 0, "Instance Already Set", vbNullString, vbExclamation
'    End If
'
'End Sub

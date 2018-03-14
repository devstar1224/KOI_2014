Attribute VB_Name = "modMinSize"
Public mHeight As Long, mWidth As Long


Option Explicit


'Private Sub Form_Load()
''--폼 최소사이즈 적용, 사이즈조정 사용 안할시엔 주석처리--
'mWidth = 3735
'mHeight = 2055
'gHW = Me.hWnd
'Hook
'
'End Sub
''
''
'Private Sub Form_Unload(Cancel As Integer)
''--이부분 역시, 폼최소사이즈관련소스, 사이즈조정 안할시 주석처리--
'       Unhook
'End Sub
'
'

            Private Const GWL_WNDPROC = -4
            Private Const WM_GETMINMAXINFO = &H24

            Private Type PointAPI
                X As Long
                Y As Long
            End Type

            Private Type MINMAXINFO
                ptReserved As PointAPI
                ptMaxSize As PointAPI
                ptMaxPosition As PointAPI
                ptMinTrackSize As PointAPI
                ptMaxTrackSize As PointAPI
            End Type

            Global lpPrevWndProc As Long
            Global gHW As Long

            Private Declare Function DefWindowProc Lib "user32" Alias _
               "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                ByVal wParam As Long, ByVal lParam As Long) As Long
            Private Declare Function CallWindowProc Lib "user32" Alias _
               "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                ByVal hWnd As Long, ByVal Msg As Long, _
                ByVal wParam As Long, ByVal lParam As Long) As Long
            Private Declare Function SetWindowLong Lib "user32" Alias _
               "SetWindowLongA" (ByVal hWnd As Long, _
                ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
            Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias _
               "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, _
                ByVal cbCopy As Long)
            Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias _
               "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, _
                ByVal cbCopy As Long)

            Public Sub Hook()
                'Start subclassing.
                lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
                   AddressOf WindowProc)
            End Sub

            Public Sub Unhook()
                Dim temp As Long

                'Cease subclassing.
                temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
            End Sub

            Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
               ByVal wParam As Long, ByVal lParam As Long) As Long
                Dim MinMax As MINMAXINFO

                'Check for request for min/max window sizes.
                If uMsg = WM_GETMINMAXINFO Then
                    'Retrieve default MinMax settings
                    CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

                    'Specify new minimum size for window.
                    MinMax.ptMinTrackSize.X = mWidth / 15
                    MinMax.ptMinTrackSize.Y = mHeight / 15

                    'Specify new maximum size for window.
                    'MinMax.ptMaxTrackSize.x = 500
                    'MinMax.ptMaxTrackSize.y = 500

                    'Copy local structure back.
                    CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

                    WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
                Else
                    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, _
                       wParam, lParam)
                End If
            End Function






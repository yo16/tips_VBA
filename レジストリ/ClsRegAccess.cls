VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsRegAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 参考：http://tsware.jp/tips/tips_283.htm 他

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
          (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
          ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
          (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
          ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
          (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
          lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Const ERROR_SUCCESS = 0
Private Const REG_SZ = 1
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_WOW64_64KEY = &H100   ' xp-64で必要？
Private Const KEY_WOW64_32KEY = &H200   ' xp-64で必要？
Public HKEY_CLASSES_ROOT   As Long
Public HKEY_CURRENT_USER   As Long
Public HKEY_LOCAL_MACHINE  As Long
Public HKEY_USERS          As Long
Public HKEY_CURRENT_CONFIG As Long


' initialize
Private Sub Class_Initialize()
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Sub

' 概要 レジストリに値を設定する
' 引数  lngRootKey : レジストリルートキー
' 　    strSubKey  : レジストリサブキー
' 　    strName    : 名前
' 　    strValue   : 設定する値
' 戻り値 [ True:成功 | False:失敗 ]
Public Function SetRegValue(lngRootKey As Long, strSubKey As String, _
                    strName As String, strValue As String)
  Dim lngRet As Long
  Dim hWnd As Long

  ' ハンドルを開く
  hWnd = Application.hWnd
  lngRet = RegOpenKeyEx(lngRootKey, strSubKey, 0, KEY_SET_VALUE, hWnd)
  ' 値を設定
  lngRet = RegSetValueEx(hWnd, strName, 0, REG_SZ, ByVal strValue, LenB(strValue))
  ' ハンドルを閉じる
  RegCloseKey hWnd
  
  ' 成功/失敗の判断を戻り値に設定
  SetRegValue = (lngRet = ERROR_SUCCESS)
End Function


' 概要 レジストリの値を取得する
' 引数 lngRootKey : レジストリルートキー
'      strSubKey  : レジストリサブキー
'      strName    : 名前
' 戻り値 取得したレジストリの値
Public Function GetRegValue(lngRootKey As Long, strSubKey As String, _
                    strName As String) As String
  Dim lngRet As Long
  Dim hWnd As Long
  Dim strValue As String

  ' ハンドルを開く
  hWnd = Application.hWnd
  lngRet = RegOpenKeyEx(lngRootKey, strSubKey, 0, KEY_QUERY_VALUE, hWnd)
  ' 受け取り値用のバッファを確保
  strValue = String(255, " ")
  ' 値を取得
  lngRet = RegQueryValueEx(hWnd, strName, 0, 0, ByVal strValue, LenB(strValue))
  ' ハンドルを閉じる
  RegCloseKey hWnd
  
  ' 取得した値からNull以降の空白を取り除く
  strValue = Left(strValue, InStr(strValue, vbNullChar) - 1)
  ' 取得した値を戻り値に設定
  GetRegValue = strValue
End Function


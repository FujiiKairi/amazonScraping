Attribute VB_Name = "Module4"
'商品データをすべて削除する。
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Declare Function DeleteUrlCacheEntry Lib "wininet" _
    Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Declare Function FindWindow Lib "User32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Declare Function FindWindowEx Lib "User32.dll" _
    Alias "FindWindowExA" ( _
    ByVal hWndParent As Long, _
    ByVal hwndChildAfter As Long, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String) As Long
Private Declare PtrSafe Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long


Sub amazonScraping()
    
'---コード1｜インターネットに接続してブラウザを開く---
    Dim objIE As InternetExplorer                               '操作するIEを入れるオブジェクトを準備
    Set objIE = CreateObject("InternetExplorer.Application")    '上のオブジェクトにIEを入れる
    objIE.Visible = True                                       '見えるように

'---コード2｜商品登録のページを開く---
    objIE.navigate "https://admin.thebase.in/shop_admin/items"               'このページを開く
    Call IEWait(objIE)   'IEを待機                              IEWaitを呼ぶ
    Call WaitFor(3) '3秒停止                                    WaitForをよぶ
    On Error GoTo ErrLabel
    
'---コード3｜商品をひとつづつ削除する---
    Dim objLink As Object
    Dim numProduct As Integer
    Dim productIndex As Integer '商品のインデックス
    numProduct = 0
    'まず商品の数を数える
    For Each objLink In objIE.document.getElementsByClassName("c-iconBtn__icon i-trash")
        numProduct = numProduct + 1
    Next
    numProduct = numProduct - 1
    For pageIndex = 0 To numProduct
        For Each objLink In objIE.document.getElementsByClassName("c-iconBtn__icon i-trash")
            objLink.Click
            Sleep 5000
        Next
    Next
    For Each objLink In objIE.document.getElementsByClassName("c-iconBtn__icon i-trash")
        'Set anchor = objLink.getElementsByTagName("a")(0)
        'tmp = Split(anchor.href, "/")
        'Worksheets("Sheet1").Cells(productIndex, 1).Value = tmp(5) 'ASIN情報
        'Worksheets("Sheet1").Cells(productIndex, 2).Value = tmp(0) & "/" & tmp(1) & "/" & tmp(2) & "/" & tmp(4) & "/" & tmp(5) & "/" 'サイトＵＲＬ
        'productIndex = productIndex + 1
    Next
    MsgBox numProduct
    Call WaitFor(3)
'---コード4｜IEを閉じる---
    
    objIE.Quit
    Set objIE = Nothing
    MsgBox ("終わりました")
ErrLabel:
    'msg = msg & "エラーが発生しました"
    Resume Next
End Sub

Function OpenPage(ByVal url As String, ByRef objIE As Object)
    objIE.navigate (url)               'このページを開く
    Call WaitFor(3) '3秒停止                                    WaitForをよぶ
End Function

'---コード2-1｜IEを待機する関数---
Function IEWait(ByRef objIE As Object)                      'オブジェクトを参照渡し
    Do While objIE.Busy = True Or objIE.readyState <> 4     'busyプロパティがtrueもしくはreadystateが4（IEオブジェクトの全データ読み込み完了状態）
        DoEvents
    Loop
End Function

'---コード2-2｜指定した秒だけ停止する関数---
Function WaitFor(ByVal second As Integer)
    Dim futureTime As Date
 
    futureTime = DateAdd("s", second, Now)                  'functimeを今の時間+second(ここでは3秒)にする
 
    While Now < futureTime
        DoEvents                                            'キャンセルボタンなどのイベントが起こった時にその処理をOSにわたす
    Wend                                                    'while do loop と同じ
End Function









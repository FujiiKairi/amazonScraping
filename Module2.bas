Attribute VB_Name = "Module2"
'エクセルの情報からBASEに出品する
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
    objIE.navigate "https://admin.thebase.in/shop_admin/items/?page=1"               'このページを開く
    Call IEWait(objIE)   'IEを待機                              IEWaitを呼ぶ
    Call WaitFor(3) '3秒停止                                    WaitForをよぶ
    
    
    On Error GoTo ErrLabel


'---コード4｜商品情報入力---
    Dim objtag, objsubmit As Object
    Dim FolderName As String    '作成したいフォルダパスを格納'
    Dim buf As String, cnt As Long   'フォルダの画像を参照するために使う
    Dim cbData As New DataObject
    Dim stock As Integer
    
    For numProduct = 2 To 100 'エクセルシートに合わせて2～100
        Sleep 1000
        If InStr(Worksheets("Sheet1").Cells(numProduct, 6).Value, "在庫あり") > 0 Then
            '---商品登録画面へ行く---
            For Each objLink In objIE.document.getElementsByClassName("i-plus c-submitBtn__iconLeft")
                    objLink.Click
            Next
            Call WaitFor(3)
            '----------------------------------ここから画像の登録
            FolderName = "C:\Users\mnuus\OneDrive\デスクトップ\kasedoll\Juely\" & Worksheets("Sheet1").Cells(numProduct, 1).Value & "\"
            buf = Dir(FolderName & "*.*")
            
            Do While buf <> ""          '画像フォルダの中をすべて回す
                
                buf = FolderName + buf
                'DataObjectにメッセージを格納
                'cbData.SetText (buf)
                'DataObjectのデータをクリップボードに格納
                'cbData.PutInClipboard
                'Set cbData = Nothing
                ClipBoard_SetData (buf)
                '----------------------------------ここからwinAPIで画像参
                objIE.document.parentWindow.execScript "window.setTimeout(""document.getElementsByTagName('input')(1).click();"",10);"
                '1秒待機
                Call WaitFor(1)
                SendKeys "^v"
                SendKeys "{Enter}"
                 
                Call WaitFor(1)
                buf = Dir()
            Loop
            
            '------------------------------------ここから説明文の登録
            'objIE.document.getElementById("itemDetail_detail").Focus
            'ClipBoard_SetData (Worksheets("Sheet1").Cells(numProduct, 4).Value)
            'SendKeys "^v"
            'Sleep 500
            'objtag.Blur
            
            '------------------------------------ここから価格の登録
            'objIE.document.getElementById("itemDetail_price").Value = Worksheets("Sheet1").Cells(numProduct, 5).Value
            objIE.document.getElementById("itemDetail_price").Focus
            ClipBoard_SetData (Worksheets("Sheet1").Cells(numProduct, 5).Value)
            SendKeys "^v"
            Sleep 500
            objtag.Blur
            '------------------------------------ここから個数の登録
            stock = CInt(2 * Rnd + 1) '乱数で2か3を生成
            'objIE.document.getElementById("itemDetail_stock").Value = stock
            objIE.document.getElementById("itemDetail_stock").Focus
            ClipBoard_SetData (stock)
            SendKeys "^v"
            Sleep 500
            objtag.Blur
            
            '----------------------------------ここから商品名登録
            
            For Each objtag In objIE.document.getElementsByTagName("input")
            If objtag.ID = "itemDetail_name" Then
                objtag.Focus
                ClipBoard_SetData (Worksheets("Sheet1").Cells(numProduct, 3).Value)
                SendKeys "^v"
                Sleep 500
                objtag.Blur
                'objtag.Value = Worksheets("Sheet1").Cells(numProduct, 3).Value
                Exit For
            End If
            Next
            '------------------------------------登録ボタンを押す
            For Each objtag In objIE.document.getElementsByTagName("button")
                If InStr(objtag.outerHTML, "登録する") > 0 Then
                    objtag.Click
                    Sleep 500
                    Exit For
                End If
            Next
            '-------------------------------------商品管理画面に戻る
            Sleep 1000
            For Each objtag In objIE.document.getElementsByTagName("a")
                If InStr(objtag.outerHTML, "商品管理") > 0 Then
                    objtag.Click
                    Exit For
                End If

            Next
        End If
    Next
'---コード5｜取得したurlからより詳しい商品情報を入手---
    Call WaitFor(3)
'---コード6｜IEを閉じる---
    
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







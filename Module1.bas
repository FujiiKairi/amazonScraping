Attribute VB_Name = "Module1"
'amazonから指定したランキングのデータをエクセルに移す
Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Declare Function DeleteUrlCacheEntry Lib "wininet" _
    Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Sub amazonScraping()
    
'---コード1｜インターネットに接続してブラウザを開く---
    Dim objIE As InternetExplorer                               '操作するIEを入れるオブジェクトを準備
    Set objIE = CreateObject("InternetExplorer.Application")    '上のオブジェクトにIEを入れる
    objIE.Visible = False                                       '見えるように

'---コード2｜インターネットの特定のページを開く---
    objIE.navigate "https://www.amazon.co.jp/ranking?type=most-gifted&ref_=nav_cs_npx_giftideas_T1"               'このページを開く
    Call IEWait(objIE)   'IEを待機                              IEWaitを呼ぶ
    Call WaitFor(3) '3秒停止                                    WaitForをよぶ
    
    

'---コード3｜ジャンル選択---
    Dim objCategory As Object
    For Each objCategory In objIE.document.getElementsByTagName("a")
        If InStr(objCategory.outerHTML, "ジュエリー") > 0 Then
            objCategory.Click
            Call WaitFor(3)
            Exit For
        End If
    Next
    
    Worksheets("Sheet1").Cells(1, 1).Value = "ASIN"
    Worksheets("Sheet1").Cells(1, 2).Value = "URL"
    Worksheets("Sheet1").Cells(1, 3).Value = "商品名"
    Worksheets("Sheet1").Cells(1, 4).Value = "商品説明"
    Worksheets("Sheet1").Cells(1, 5).Value = "カート価格"
    Worksheets("Sheet1").Cells(1, 6).Value = "在庫状況"
    Worksheets("Sheet1").Cells(1, 7).Value = "画像1"
    Worksheets("Sheet1").Cells(1, 8).Value = "画像2"
    Worksheets("Sheet1").Cells(1, 9).Value = "画像3"
    Worksheets("Sheet1").Cells(1, 10).Value = "画像4"
'---コード4｜商品リンクとASINをランキング上位100件取得---
    Dim objLink As Object
    Dim productIndex As Integer '商品の個数のインデックス
    Dim pageIndex As Integer 'ページ数のインデックス
    Dim tmp As Variant
    Dim pageArray(5) As HTMLAnchorElement
    Dim anchor As HTMLAnchorElement
    Dim objtsugi As Object
    
    productIndex = 2
    
    
    For pageIndex = 1 To 5
        For Each objLink In objIE.document.getElementsByClassName("a-fixed-left-grid-col a-col-right")
            Set anchor = objLink.getElementsByTagName("a")(0)
            tmp = Split(anchor.href, "/")
            Worksheets("Sheet1").Cells(productIndex, 1).Value = tmp(5) 'ASIN情報
            Worksheets("Sheet1").Cells(productIndex, 2).Value = tmp(0) & "/" & tmp(1) & "/" & tmp(2) & "/" & tmp(4) & "/" & tmp(5) & "/" 'サイトＵＲＬ
            productIndex = productIndex + 1
        Next
    If pageIndex = 1 Then
    For Each objtsugi In objIE.document.getElementsByTagName("a")
        If InStr(objtsugi.outerHTML, "21-40") > 0 Then
            objtsugi.Click
            Call WaitFor(3)
            Exit For
        End If
    Next
    End If
    If pageIndex = 2 Then
        For Each objtsugi In objIE.document.getElementsByTagName("a")
        If InStr(objtsugi.outerHTML, "41-60") > 0 Then
            objtsugi.Click
            Call WaitFor(3)
            Exit For
        End If
    Next
    End If
    If pageIndex = 3 Then
        For Each objtsugi In objIE.document.getElementsByTagName("a")
        If InStr(objtsugi.outerHTML, "61-80") > 0 Then
            objtsugi.Click
            Call WaitFor(3)
            Exit For
        End If
    Next
    End If
    If pageIndex = 4 Then
        For Each objtsugi In objIE.document.getElementsByTagName("a")
        If InStr(objtsugi.outerHTML, "81-100") > 0 Then
            objtsugi.Click
            Call WaitFor(3)
            Exit For
        End If
    Next
    End If
    Next
    
'---コード5｜取得したurlからより詳しい商品情報を入手---
    On Error GoTo ErrLabel
    
    Dim numProduct As Integer
    Dim index2 As Integer
    Dim descriptionText As String
    Dim arr2() As String
    Dim coltd As IHTMLElementCollection
    Dim el As IHTMLElement
    Dim FolderName As String    '作成したいフォルダパスを格納'
    Dim elm As Object
    Dim numPicture As Integer
    Dim numDownPicture As Integer
    Dim numExcelPicture As Integer
    Dim i2 As Integer
    Dim imagesTextArray() As String
    Dim imagesTextArray2() As String
    Dim srcText As String
    Dim imgURL As String, fileName As String, savePath As String
    Dim cacheDel As Long, result As Long
    For numProduct = 2 To 101 'エクセルシートに合わせて2～100
        objIE.navigate (Cells(numProduct, 2).Value)
        Call WaitFor(3)
        '----------------------------------------ここから商品名の取得
        Worksheets("Sheet1").Cells(numProduct, 3).Value = objIE.document.getElementById("productTitle").innerHTML
        Worksheets("Sheet1").Cells(numProduct, 3).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 3).Value, vbLf, "")
        Worksheets("Sheet1").Cells(numProduct, 3).Value = Trim(Worksheets("Sheet1").Cells(numProduct, 3).Value)
        'Worksheets("Sheet1").Cells(numProduct, 3).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 3).Value, " ", "")
        'Worksheets("Sheet1").Cells(numProduct, 3).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 3).Value, "　", "")
        
        '----------------------------------------ここから説明文の取得
        descriptionText = ""
        For Each objLink In objIE.document.getElementsByClassName("a-unordered-list a-vertical a-spacing-none") 'ここは実質繰り返さない
            For index2 = 0 To objLink.getElementsByTagName("span").Length - 1
                descriptionText = descriptionText + objLink.getElementsByTagName("span")(index2).innerHTML
            Next
            Worksheets("Sheet1").Cells(numProduct, 4).Value = descriptionText
            arr2 = Split(descriptionText, vbCrLf)
            descriptionText = ""
            For i = LBound(arr2) To UBound(arr2)
                If Len(arr2(i)) > 1 Then
                    descriptionText = descriptionText & arr2(i)
                    End If
            Next i
        Next
        '----------------------------------------ここから在庫状況の取得
        Set coltd = objIE.document.getElementById("availability").getElementsByTagName("span")
        Worksheets("Sheet1").Cells(numProduct, 6).Value = Replace(coltd(0).innerHTML, " ", "")
        Worksheets("Sheet1").Cells(numProduct, 6).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 6).Value, vbLf, "")
        Worksheets("Sheet1").Cells(numProduct, 6).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 6).Value, " ", "")
        Worksheets("Sheet1").Cells(numProduct, 6).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 6).Value, "　", "")
        '----------------------------------------ここから価格の取得
        If InStr(Worksheets("Sheet1").Cells(numProduct, 6).Value, "現在在庫") > 0 Then
            Worksheets("Sheet1").Cells(numProduct, 5).Value = 0
        Else
            If InStr(objIE.document.getElementById("ppd").innerHTML, "中古") > 0 Then
                Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(objIE.document.getElementById("newBuyBoxPrice").innerHTML, "￥", "")
            Else
                Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(objIE.document.getElementById("price_inside_buybox").innerHTML, "￥", "")
            End If
        End If
        Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 5).Value, vbLf, "")
        Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 5).Value, " ", "")
        Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 5).Value, "　", "")
        
        
        '----------------------------------------ここから画像の取得
        FolderName = "C:\Users\mnuus\OneDrive\デスクトップ\kasedoll\Juely\" & Worksheets("Sheet1").Cells(numProduct, 1).Value
        If Dir(FolderName, vbDirectory) = "" Then   '同名のフォルダがない場合フォルダを作成'
            MkDir FolderName
        End If
        '小さい画像の数を数えてクリックする
        numPicture = 0
        For Each objLink In objIE.document.getElementsByClassName("a-spacing-small item imageThumbnail a-declarative")
            objLink.Click
            
        Next
        numPicture = objIE.document.getElementsByClassName("a-dynamic-image a-stretch-vertical").Length
        If numPicture < 4 Then
            numDownPicture = numPicture
        Else
            numDownPicture = 4
        End If
        
        
        For i2 = 0 To numDownPicture - 1
        imagesTextArray = Split(objIE.document.getElementsByClassName("a-dynamic-image a-stretch-vertical")(i2).outerHTML, " ") '画像があるクラス
            For i = 0 To UBound(imagesTextArray) - LBound(imagesTextArray) - 1
                If InStr(imagesTextArray(i), "src=") <> 0 Then
                    srcText = imagesTextArray(i)
                    Exit For
                End If
            Next
            imagesTextArray2 = Split(srcText, """")
            Worksheets("Sheet1").Cells(numProduct, 7 + i2).Value = imagesTextArray2(1)
            '---------ここから画像ファイル保存
            imgURL = imagesTextArray2(1)
            If InStr(imgURL, ".jpg") <> 0 Then
                '画像ファイル名
                  fileName = Mid(imgURL, InStrRev(imgURL, "/") + 1)
                
                  '画像保存先(+画像ファイル名）
                  savePath = FolderName & "\" & i2 & ".jpg"
                
                  'キャッシュクリア
                  cacheDel = DeleteUrlCacheEntry(imgURL)
                
                  '画像ダウンロード
                  result = URLDownloadToFile(0, imgURL, savePath, 0, 0)
            End If
        Next
        
       
    Next
'---コード6｜IEを閉じる---
    MsgBox ("終わりました")
    objIE.Quit
    Set objIE = Nothing
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





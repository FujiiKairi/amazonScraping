Attribute VB_Name = "Module1"
'amazon����w�肵�������L���O�̃f�[�^���G�N�Z���Ɉڂ�
Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Declare Function DeleteUrlCacheEntry Lib "wininet" _
    Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Sub amazonScraping()
    
'---�R�[�h1�b�C���^�[�l�b�g�ɐڑ����ău���E�U���J��---
    Dim objIE As InternetExplorer                               '���삷��IE������I�u�W�F�N�g������
    Set objIE = CreateObject("InternetExplorer.Application")    '��̃I�u�W�F�N�g��IE������
    objIE.Visible = False                                       '������悤��

'---�R�[�h2�b�C���^�[�l�b�g�̓���̃y�[�W���J��---
    objIE.navigate "https://www.amazon.co.jp/ranking?type=most-gifted&ref_=nav_cs_npx_giftideas_T1"               '���̃y�[�W���J��
    Call IEWait(objIE)   'IE��ҋ@                              IEWait���Ă�
    Call WaitFor(3) '3�b��~                                    WaitFor�����
    
    

'---�R�[�h3�b�W�������I��---
    Dim objCategory As Object
    For Each objCategory In objIE.document.getElementsByTagName("a")
        If InStr(objCategory.outerHTML, "�W���G���[") > 0 Then
            objCategory.Click
            Call WaitFor(3)
            Exit For
        End If
    Next
    
    Worksheets("Sheet1").Cells(1, 1).Value = "ASIN"
    Worksheets("Sheet1").Cells(1, 2).Value = "URL"
    Worksheets("Sheet1").Cells(1, 3).Value = "���i��"
    Worksheets("Sheet1").Cells(1, 4).Value = "���i����"
    Worksheets("Sheet1").Cells(1, 5).Value = "�J�[�g���i"
    Worksheets("Sheet1").Cells(1, 6).Value = "�݌ɏ�"
    Worksheets("Sheet1").Cells(1, 7).Value = "�摜1"
    Worksheets("Sheet1").Cells(1, 8).Value = "�摜2"
    Worksheets("Sheet1").Cells(1, 9).Value = "�摜3"
    Worksheets("Sheet1").Cells(1, 10).Value = "�摜4"
'---�R�[�h4�b���i�����N��ASIN�������L���O���100���擾---
    Dim objLink As Object
    Dim productIndex As Integer '���i�̌��̃C���f�b�N�X
    Dim pageIndex As Integer '�y�[�W���̃C���f�b�N�X
    Dim tmp As Variant
    Dim pageArray(5) As HTMLAnchorElement
    Dim anchor As HTMLAnchorElement
    Dim objtsugi As Object
    
    productIndex = 2
    
    
    For pageIndex = 1 To 5
        For Each objLink In objIE.document.getElementsByClassName("a-fixed-left-grid-col a-col-right")
            Set anchor = objLink.getElementsByTagName("a")(0)
            tmp = Split(anchor.href, "/")
            Worksheets("Sheet1").Cells(productIndex, 1).Value = tmp(5) 'ASIN���
            Worksheets("Sheet1").Cells(productIndex, 2).Value = tmp(0) & "/" & tmp(1) & "/" & tmp(2) & "/" & tmp(4) & "/" & tmp(5) & "/" '�T�C�g�t�q�k
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
    
'---�R�[�h5�b�擾����url������ڂ������i�������---
    On Error GoTo ErrLabel
    
    Dim numProduct As Integer
    Dim index2 As Integer
    Dim descriptionText As String
    Dim arr2() As String
    Dim coltd As IHTMLElementCollection
    Dim el As IHTMLElement
    Dim FolderName As String    '�쐬�������t�H���_�p�X���i�['
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
    For numProduct = 2 To 101 '�G�N�Z���V�[�g�ɍ��킹��2�`100
        objIE.navigate (Cells(numProduct, 2).Value)
        Call WaitFor(3)
        '----------------------------------------�������珤�i���̎擾
        Worksheets("Sheet1").Cells(numProduct, 3).Value = objIE.document.getElementById("productTitle").innerHTML
        Worksheets("Sheet1").Cells(numProduct, 3).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 3).Value, vbLf, "")
        Worksheets("Sheet1").Cells(numProduct, 3).Value = Trim(Worksheets("Sheet1").Cells(numProduct, 3).Value)
        'Worksheets("Sheet1").Cells(numProduct, 3).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 3).Value, " ", "")
        'Worksheets("Sheet1").Cells(numProduct, 3).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 3).Value, "�@", "")
        
        '----------------------------------------��������������̎擾
        descriptionText = ""
        For Each objLink In objIE.document.getElementsByClassName("a-unordered-list a-vertical a-spacing-none") '�����͎����J��Ԃ��Ȃ�
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
        '----------------------------------------��������݌ɏ󋵂̎擾
        Set coltd = objIE.document.getElementById("availability").getElementsByTagName("span")
        Worksheets("Sheet1").Cells(numProduct, 6).Value = Replace(coltd(0).innerHTML, " ", "")
        Worksheets("Sheet1").Cells(numProduct, 6).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 6).Value, vbLf, "")
        Worksheets("Sheet1").Cells(numProduct, 6).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 6).Value, " ", "")
        Worksheets("Sheet1").Cells(numProduct, 6).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 6).Value, "�@", "")
        '----------------------------------------�������牿�i�̎擾
        If InStr(Worksheets("Sheet1").Cells(numProduct, 6).Value, "���ݍ݌�") > 0 Then
            Worksheets("Sheet1").Cells(numProduct, 5).Value = 0
        Else
            If InStr(objIE.document.getElementById("ppd").innerHTML, "����") > 0 Then
                Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(objIE.document.getElementById("newBuyBoxPrice").innerHTML, "��", "")
            Else
                Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(objIE.document.getElementById("price_inside_buybox").innerHTML, "��", "")
            End If
        End If
        Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 5).Value, vbLf, "")
        Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 5).Value, " ", "")
        Worksheets("Sheet1").Cells(numProduct, 5).Value = Replace(Worksheets("Sheet1").Cells(numProduct, 5).Value, "�@", "")
        
        
        '----------------------------------------��������摜�̎擾
        FolderName = "C:\Users\mnuus\OneDrive\�f�X�N�g�b�v\kasedoll\Juely\" & Worksheets("Sheet1").Cells(numProduct, 1).Value
        If Dir(FolderName, vbDirectory) = "" Then   '�����̃t�H���_���Ȃ��ꍇ�t�H���_���쐬'
            MkDir FolderName
        End If
        '�������摜�̐��𐔂��ăN���b�N����
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
        imagesTextArray = Split(objIE.document.getElementsByClassName("a-dynamic-image a-stretch-vertical")(i2).outerHTML, " ") '�摜������N���X
            For i = 0 To UBound(imagesTextArray) - LBound(imagesTextArray) - 1
                If InStr(imagesTextArray(i), "src=") <> 0 Then
                    srcText = imagesTextArray(i)
                    Exit For
                End If
            Next
            imagesTextArray2 = Split(srcText, """")
            Worksheets("Sheet1").Cells(numProduct, 7 + i2).Value = imagesTextArray2(1)
            '---------��������摜�t�@�C���ۑ�
            imgURL = imagesTextArray2(1)
            If InStr(imgURL, ".jpg") <> 0 Then
                '�摜�t�@�C����
                  fileName = Mid(imgURL, InStrRev(imgURL, "/") + 1)
                
                  '�摜�ۑ���(+�摜�t�@�C�����j
                  savePath = FolderName & "\" & i2 & ".jpg"
                
                  '�L���b�V���N���A
                  cacheDel = DeleteUrlCacheEntry(imgURL)
                
                  '�摜�_�E�����[�h
                  result = URLDownloadToFile(0, imgURL, savePath, 0, 0)
            End If
        Next
        
       
    Next
'---�R�[�h6�bIE�����---
    MsgBox ("�I���܂���")
    objIE.Quit
    Set objIE = Nothing
ErrLabel:
    'msg = msg & "�G���[���������܂���"
    Resume Next
End Sub

Function OpenPage(ByVal url As String, ByRef objIE As Object)
    objIE.navigate (url)               '���̃y�[�W���J��
    Call WaitFor(3) '3�b��~                                    WaitFor�����
End Function

'---�R�[�h2-1�bIE��ҋ@����֐�---
Function IEWait(ByRef objIE As Object)                      '�I�u�W�F�N�g���Q�Ɠn��
    Do While objIE.Busy = True Or objIE.readyState <> 4     'busy�v���p�e�B��true��������readystate��4�iIE�I�u�W�F�N�g�̑S�f�[�^�ǂݍ��݊�����ԁj
        DoEvents
    Loop
End Function

'---�R�[�h2-2�b�w�肵���b������~����֐�---
Function WaitFor(ByVal second As Integer)
    Dim futureTime As Date
 
    futureTime = DateAdd("s", second, Now)                  'functime�����̎���+second(�����ł�3�b)�ɂ���
 
    While Now < futureTime
        DoEvents                                            '�L�����Z���{�^���Ȃǂ̃C�x���g���N���������ɂ��̏�����OS�ɂ킽��
    Wend                                                    'while do loop �Ɠ���
End Function





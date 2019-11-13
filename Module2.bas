Attribute VB_Name = "Module2"
'�G�N�Z���̏�񂩂�BASE�ɏo�i����
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
    
'---�R�[�h1�b�C���^�[�l�b�g�ɐڑ����ău���E�U���J��---
    Dim objIE As InternetExplorer                               '���삷��IE������I�u�W�F�N�g������
    Set objIE = CreateObject("InternetExplorer.Application")    '��̃I�u�W�F�N�g��IE������
    objIE.Visible = True                                       '������悤��

'---�R�[�h2�b���i�o�^�̃y�[�W���J��---
    objIE.navigate "https://admin.thebase.in/shop_admin/items/?page=1"               '���̃y�[�W���J��
    Call IEWait(objIE)   'IE��ҋ@                              IEWait���Ă�
    Call WaitFor(3) '3�b��~                                    WaitFor�����
    
    
    On Error GoTo ErrLabel


'---�R�[�h4�b���i������---
    Dim objtag, objsubmit As Object
    Dim FolderName As String    '�쐬�������t�H���_�p�X���i�['
    Dim buf As String, cnt As Long   '�t�H���_�̉摜���Q�Ƃ��邽�߂Ɏg��
    Dim cbData As New DataObject
    Dim stock As Integer
    
    For numProduct = 2 To 100 '�G�N�Z���V�[�g�ɍ��킹��2�`100
        Sleep 1000
        If InStr(Worksheets("Sheet1").Cells(numProduct, 6).Value, "�݌ɂ���") > 0 Then
            '---���i�o�^��ʂ֍s��---
            For Each objLink In objIE.document.getElementsByClassName("i-plus c-submitBtn__iconLeft")
                    objLink.Click
            Next
            Call WaitFor(3)
            '----------------------------------��������摜�̓o�^
            FolderName = "C:\Users\mnuus\OneDrive\�f�X�N�g�b�v\kasedoll\Juely\" & Worksheets("Sheet1").Cells(numProduct, 1).Value & "\"
            buf = Dir(FolderName & "*.*")
            
            Do While buf <> ""          '�摜�t�H���_�̒������ׂĉ�
                
                buf = FolderName + buf
                'DataObject�Ƀ��b�Z�[�W���i�[
                'cbData.SetText (buf)
                'DataObject�̃f�[�^���N���b�v�{�[�h�Ɋi�[
                'cbData.PutInClipboard
                'Set cbData = Nothing
                ClipBoard_SetData (buf)
                '----------------------------------��������winAPI�ŉ摜�Q
                objIE.document.parentWindow.execScript "window.setTimeout(""document.getElementsByTagName('input')(1).click();"",10);"
                '1�b�ҋ@
                Call WaitFor(1)
                SendKeys "^v"
                SendKeys "{Enter}"
                 
                Call WaitFor(1)
                buf = Dir()
            Loop
            
            '------------------------------------��������������̓o�^
            'objIE.document.getElementById("itemDetail_detail").Focus
            'ClipBoard_SetData (Worksheets("Sheet1").Cells(numProduct, 4).Value)
            'SendKeys "^v"
            'Sleep 500
            'objtag.Blur
            
            '------------------------------------�������牿�i�̓o�^
            'objIE.document.getElementById("itemDetail_price").Value = Worksheets("Sheet1").Cells(numProduct, 5).Value
            objIE.document.getElementById("itemDetail_price").Focus
            ClipBoard_SetData (Worksheets("Sheet1").Cells(numProduct, 5).Value)
            SendKeys "^v"
            Sleep 500
            objtag.Blur
            '------------------------------------����������̓o�^
            stock = CInt(2 * Rnd + 1) '������2��3�𐶐�
            'objIE.document.getElementById("itemDetail_stock").Value = stock
            objIE.document.getElementById("itemDetail_stock").Focus
            ClipBoard_SetData (stock)
            SendKeys "^v"
            Sleep 500
            objtag.Blur
            
            '----------------------------------�������珤�i���o�^
            
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
            '------------------------------------�o�^�{�^��������
            For Each objtag In objIE.document.getElementsByTagName("button")
                If InStr(objtag.outerHTML, "�o�^����") > 0 Then
                    objtag.Click
                    Sleep 500
                    Exit For
                End If
            Next
            '-------------------------------------���i�Ǘ���ʂɖ߂�
            Sleep 1000
            For Each objtag In objIE.document.getElementsByTagName("a")
                If InStr(objtag.outerHTML, "���i�Ǘ�") > 0 Then
                    objtag.Click
                    Exit For
                End If

            Next
        End If
    Next
'---�R�[�h5�b�擾����url������ڂ������i�������---
    Call WaitFor(3)
'---�R�[�h6�bIE�����---
    
    objIE.Quit
    Set objIE = Nothing
    MsgBox ("�I���܂���")
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







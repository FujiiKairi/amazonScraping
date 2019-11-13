Attribute VB_Name = "Module4"
'���i�f�[�^�����ׂč폜����B
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
    objIE.navigate "https://admin.thebase.in/shop_admin/items"               '���̃y�[�W���J��
    Call IEWait(objIE)   'IE��ҋ@                              IEWait���Ă�
    Call WaitFor(3) '3�b��~                                    WaitFor�����
    On Error GoTo ErrLabel
    
'---�R�[�h3�b���i���ЂƂÂ폜����---
    Dim objLink As Object
    Dim numProduct As Integer
    Dim productIndex As Integer '���i�̃C���f�b�N�X
    numProduct = 0
    '�܂����i�̐��𐔂���
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
        'Worksheets("Sheet1").Cells(productIndex, 1).Value = tmp(5) 'ASIN���
        'Worksheets("Sheet1").Cells(productIndex, 2).Value = tmp(0) & "/" & tmp(1) & "/" & tmp(2) & "/" & tmp(4) & "/" & tmp(5) & "/" '�T�C�g�t�q�k
        'productIndex = productIndex + 1
    Next
    MsgBox numProduct
    Call WaitFor(3)
'---�R�[�h4�bIE�����---
    
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









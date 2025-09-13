VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Sub PdfInit Lib "PDFLib" ()
Private Declare PtrSafe Sub OpenPDF Lib "PDFLib" (ByVal filename As LongPtr, Optional password As LongPtr = 0)
Private Declare PtrSafe Function GetPageCount Lib "PDFLib" () As Long
Private Declare PtrSafe Sub ExtractPageText Lib "PDFLib" (ByVal page As Long, ByRef buffer As LongPtr, ByRef len_ As Long)
Private Declare PtrSafe Sub ClosePDF Lib "PDFLib" ()
Private Declare PtrSafe Sub PDFFree Lib "PDFLib" ()
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#Else
 Enum LongPtr
  
  [_]
 End Enum
Private Declare Sub PdfInit Lib "PDFLib" ()
Private Declare Sub OpenPDF Lib "PDFLib" (ByVal filename As LongPtr, Optional password As LongPtr = 0)
Private Declare Function GetPageCount Lib "PDFLib" () As Long
Private Declare Sub ExtractPageText Lib "PDFLib" (ByVal page As Long, ByRef buffer As LongPtr, ByRef len_ As Long)
Private Declare Sub ClosePDF Lib "PDFLib" ()
Private Declare Sub PDFFree Lib "PDFLib" ()
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If

Private Function PtrToStr(ByVal ptr As LongPtr, len_ As Long) As Byte() '���ַ�������ָ��ת�ַ���
    Dim buffer() As Byte, n As Long
    n = len_ * 2
    ReDim buffer(0 To n - 1)
    ' �����ڴ浽��ȫ���ֽ�����
    CopyMemory buffer(0), ByVal ptr, n
    ' ���ֽ�����ת��Ϊ�ַ���
    PtrToStr = buffer()
End Function

Sub test()
Dim LibPath As String
#If Win64 Then
LibPath = "\win64"
#Else
LibPath = "\win32"
#End If
LibPath = App.Path & LibPath
ChDrive LibPath
ChDir LibPath
Dim buffer As LongPtr, slen As Long, pdfText As String
PdfInit
OpenPDF StrPtr(App.Path & "\1.pdf")
MsgBox GetPageCount() '��ȡpdf�ļ���ҳ��
ExtractPageText 1, buffer, slen '��ȡ��һҳ������,��һ����
pdfText = PtrToStr(buffer, slen)
Debug.Print pdfText
ClosePDF
PDFFree
End Sub

Private Sub Form_Load()
test
End Sub

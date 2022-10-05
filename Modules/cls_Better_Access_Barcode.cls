VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Better_Access_Barcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'############################################################################################
'# Copyright (c) 2022 Thomas Möller                                                         #
'# MIT License  => https://github.com/team-moeller/better-access-barcode/blob/main/LICENSE  #
'# Version 1.06.10  published: 05.10.2022                                                   #
'############################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum bcScriptSource
    CDN = 1
    LocalFile = 2
End Enum

Public Enum bcTextAlign
    taLeft = 1
    taCenter = 2
    taRight = 3
End Enum

Public Enum bcFormat
    CODE128A = 1
    CODE128B = 2
    CODE128C = 3
    EAN13 = 4
    UPC = 5
    EAN8 = 6
    EAN5 = 7
    EAN2 = 8
    CODE39 = 9
    ITF14 = 10
    MSI = 11
    MSI10 = 12
    MSI11 = 13
    MSI1010 = 14
    MSI1110 = 15
    pharmacode = 16
    codabar = 17
End Enum


'### Members

Private WithEvents m_Control As WebbrowserControl
Attribute m_Control.VB_VarHelpID = -1
Private m_JsBarcodeSource As bcScriptSource
Private m_JsBarcodeCDN As String
Private m_JsBarcodePathFilename As String
Private m_Data As String
Private m_Format As bcFormat
Private m_BarWidth As Long
Private m_Height As Long
Private m_Background As String
Private m_LineColor As String
Private m_Margin As Long
Private m_DisplayValue As Boolean
Private m_Font As String
Private m_IsBold As Boolean
Private m_IsItalic As Boolean
Private m_FontSize As Long
Private m_TextAlign As bcTextAlign


'### Properties

Public Property Set Control(newValue As WebbrowserControl)
    Set m_Control = newValue
End Property
Public Property Get Control() As WebbrowserControl
    Set Control = m_Control
End Property

Public Property Let JsBarcodeSource(ByVal newValue As bcScriptSource)
    m_JsBarcodeSource = newValue
End Property
Public Property Get JsBarcodeSource() As bcScriptSource
    JsBarcodeSource = m_JsBarcodeSource
End Property

Public Property Let JsBarcodeCDN(ByVal newValue As String)
    m_JsBarcodeCDN = newValue
End Property
Public Property Get JsBarcodeCDN() As String
    JsBarcodeCDN = m_JsBarcodeCDN
End Property

Public Property Let JsBarcodePathFilename(ByVal newValue As String)
    m_JsBarcodePathFilename = newValue
End Property
Public Property Get JsBarcodePathFilename() As String
    JsBarcodePathFilename = m_JsBarcodePathFilename
End Property

Public Property Let Data(newValue As String)
    m_Data = newValue
End Property
Public Property Get Data() As String
    Data = m_Data
End Property

Public Property Let Format(newValue As bcFormat)
    m_Format = newValue
End Property
Public Property Get Format() As bcFormat
    Format = m_Format
End Property

Public Property Let BarWidth(newValue As Long)
    m_BarWidth = newValue
End Property
Public Property Get BarWidth() As Long
    BarWidth = m_BarWidth
End Property

Public Property Let Height(newValue As Long)
    m_Height = newValue
End Property
Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Let Background(ByVal newValue As String)
    m_Background = newValue
End Property
Public Property Get Background() As String
    Background = m_Background
End Property

Public Property Let LineColor(ByVal newValue As String)
    m_LineColor = newValue
End Property
Public Property Get LineColor() As String
    LineColor = m_LineColor
End Property

Public Property Let Margin(newValue As Long)
    m_Margin = newValue
End Property
Public Property Get Margin() As Long
    Margin = m_Margin
End Property

Public Property Let DisplayValue(newValue As Boolean)
    m_DisplayValue = newValue
End Property
Public Property Get DisplayValue() As Boolean
    DisplayValue = m_DisplayValue
End Property

Public Property Let Font(ByVal newValue As String)
    m_Font = newValue
End Property
Public Property Get Font() As String
    Font = m_Font
End Property

Public Property Let IsBold(newValue As Boolean)
    m_IsBold = newValue
End Property
Public Property Get IsBold() As Boolean
    IsBold = m_IsBold
End Property

Public Property Let IsItalic(newValue As Boolean)
    m_IsItalic = newValue
End Property
Public Property Get IsItalic() As Boolean
    IsItalic = m_IsItalic
End Property

Public Property Let FontSize(newValue As Long)
    m_FontSize = newValue
End Property
Public Property Get FontSize() As Long
    FontSize = m_FontSize
End Property

Public Property Let TextAlign(ByVal newValue As bcTextAlign)
    m_TextAlign = newValue
End Property
Public Property Get TextAlign() As bcTextAlign
    TextAlign = m_TextAlign
End Property


'### Instanzing

Private Sub Class_Initialize()

    Me.JsBarcodeSource = bcScriptSource.CDN
    Me.JsBarcodeCDN = "https://cdn.jsdelivr.net/npm/jsbarcode@3.11.0/dist/JsBarcode.all.min.js"
    Me.JsBarcodePathFilename = CurrentProject.Path & "\JsBarcode.all.min.js"
    
    Me.DisplayValue = True
    Me.Font = "Monospace"
    Me.FontSize = 12
    Me.TextAlign = bcTextAlign.taCenter
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub ShowBarcode()

    Dim Datei As String
    
    CreateHtmlFile
       
    Datei = CurrentProject.Path & "\BetterAccessBarcode.html"
    
    Control.Object.Navigate Datei
    
End Sub


'### Private Methods

Private Sub CreateHtmlFile()

    'Variablen deklarieren
    Dim d As Integer
    Dim strPath As String
    Dim strFilename As String
    Dim strContent As String
    Const ForWriting As Long = 2
    
    strPath = CurrentProject.Path
    strFilename = "\BetterAccessBarcode.html"
    
    strContent = strContent & getHead
    strContent = strContent & getScript
    strContent = strContent & getFoot
 
    d = FreeFile()
    Open strPath & strFilename For Output As d
    Print #d, strContent;
 
    Close d
  
End Sub

Private Function getHead() As String

    'Variablen deklarieren
    Dim strText As String
    
    strText = strText & "<!DOCTYPE html>" & vbCrLf
    strText = strText & "<!-- saved from url=(0014)about:internet -->" & vbCrLf
    strText = strText & "<html lang='en'>" & vbCrLf
    strText = strText & "<head>" & vbCrLf
    strText = strText & "    <meta charset='utf-8'>" & vbCrLf
    strText = strText & "    <meta http-equiv='X-UA-Compatible' content='IE=Edge'/>" & vbCrLf
    strText = strText & "    <title>Better Access Barcode</title>" & vbCrLf
    strText = strText & "    <style>body{margin:0;padding:0;}</style>" & vbCrLf
    strText = strText & "    <script src=" & getLocation & "></script>" & vbCrLf
    strText = strText & "</head>" & vbCrLf
    strText = strText & "<body>" & vbCrLf
    
    getHead = strText
    
End Function

Private Function getScript() As String

    'Variablen deklarieren
    Dim strText As String
    
    strText = strText & "    <canvas id='barcode'></canvas>" & vbCrLf
    strText = strText & "    <script>" & vbCrLf
    strText = strText & "        JsBarcode('#barcode', '" & Me.Data & "', {" & vbCrLf
    strText = strText & "            format:'" & getBarcodeFormatString(Me.Format) & "'," & vbCrLf
    strText = strText & "            width:" & Me.BarWidth & "," & vbCrLf
    strText = strText & "            height:" & Me.Height & "," & vbCrLf
    strText = strText & "            background:'" & Me.Background & "'," & vbCrLf
    strText = strText & "            lineColor:'" & Me.LineColor & "'," & vbCrLf
    strText = strText & "            margin:" & Me.Margin & "," & vbCrLf
'    strText = strText & "            marginTop:10," & vbCrLf
'    strText = strText & "            marginLeft:10," & vbCrLf
'    strText = strText & "            marginRight:10," & vbCrLf
'    strText = strText & "            marginBottom:10," & vbCrLf
    strText = strText & "            displayValue:" & IIf(Me.DisplayValue, "true", "false") & " ," & vbCrLf
    strText = strText & "            font:'" & Me.Font & "'," & vbCrLf
    strText = strText & "            fontOptions:" & getFontOptions & "," & vbCrLf
    strText = strText & "            fontSize:" & Me.FontSize & "," & vbCrLf
    strText = strText & "            textAlign:" & getTextAlign & "," & vbCrLf
    strText = strText & "            textPosition:'bottom'," & vbCrLf
    strText = strText & "            textMargin:2" & vbCrLf
    strText = strText & "        });" & vbCrLf
    strText = strText & "    </script>" & vbCrLf
    
    getScript = strText

End Function

Private Function getFoot() As String

    'Variablen deklarieren
    Dim strText As String
    
    strText = strText & "</body>" & vbCrLf
    strText = strText & "</html>"
    
    getFoot = strText
    
End Function

Private Function getBarcodeFormatString(This_Format As bcFormat) As String

    'Variablen deklarieren
    Dim strText As String
    
    Select Case This_Format
        Case bcFormat.CODE128A
            strText = "CODE128A"
        Case bcFormat.CODE128B
            strText = "CODE128B"
        Case bcFormat.CODE128C
            strText = "CODE128C"
        Case bcFormat.EAN13
            strText = "EAN13"
        Case bcFormat.UPC
            strText = "UPC"
        Case bcFormat.EAN8
            strText = "EAN8"
        Case bcFormat.EAN5
            strText = "EAN5"
        Case bcFormat.EAN2
            strText = "EAN2"
        Case bcFormat.CODE39
            strText = "CODE39"
        Case bcFormat.ITF14
            strText = "ITF14"
        Case bcFormat.MSI
            strText = "MSI"
        Case bcFormat.MSI10
            strText = "MSI10"
        Case bcFormat.MSI11
            strText = "MSI11"
        Case bcFormat.MSI1010
            strText = "MSI1010"
        Case bcFormat.MSI1110
            strText = "MSI1110"
        Case bcFormat.pharmacode
            strText = "pharmacode"
        Case bcFormat.codabar
            strText = "codabar"
        Case Else
            strText = "CODE128"
    End Select
    
    getBarcodeFormatString = strText
    
End Function

Private Function getLocation() As String

    'Declarations
    Dim strText As String
    
    Select Case Me.JsBarcodeSource
        Case bcScriptSource.CDN
            strText = "'" & Me.JsBarcodeCDN & "'"
        Case bcScriptSource.LocalFile
            strText = "'" & Me.JsBarcodePathFilename & "'"
        Case Else
            'Do nothing
    End Select
            
    getLocation = strText

End Function

Private Function getFontOptions() As String

    'Declarations
    Dim strText As String

    If Me.IsBold = True Then
        strText = "bold"
    End If
    If Me.IsItalic = True Then
        If Len(strText) = 0 Then
            strText = "italic"
        Else
            strText = "bold italic"
        End If
    End If
    strText = "'" & strText & "'"
        
    getFontOptions = strText

End Function

Private Function getTextAlign() As String

    'Declarations
    Dim strText As String

    Select Case Me.TextAlign
        Case bcTextAlign.taLeft
            strText = "left"
        Case bcTextAlign.taCenter
            strText = "cnter"
        Case bcTextAlign.taRight
            strText = "right"
        Case Else
            strText = "center"
    End Select

    strText = "'" & strText & "'"
        
    getTextAlign = strText

End Function


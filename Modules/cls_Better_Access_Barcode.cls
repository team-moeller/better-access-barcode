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
'# Version 0.94.03  published: 28.09.2022                                                   #
'############################################################################################

Option Compare Database
Option Explicit


'### Members

Private WithEvents m_Control As WebbrowserControl
Attribute m_Control.VB_VarHelpID = -1
Private m_Data As String


'### Properties

Public Property Set Control(This_Control As WebbrowserControl)
    Set m_Control = This_Control
End Property
Public Property Get Control() As WebbrowserControl
    Set Control = m_Control
End Property

Public Property Let Data(This_Data As String)
    m_Data = This_Data
End Property
Public Property Get Data() As String
    Data = m_Data
End Property

'### Instanzing

Private Sub Class_Initialize()
    ' Do Nothing
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
    strText = strText & "    <script src='https://cdn.jsdelivr.net/npm/jsbarcode@3.11.0/dist/JsBarcode.all.min.js'></script>" & vbCrLf
    strText = strText & "</head>" & vbCrLf
    strText = strText & "<body>" & vbCrLf
    
    getHead = strText
    
End Function

Private Function getScript() As String

    'Variablen deklarieren
    Dim strText As String
    
    strText = strText & "    <canvas id='barcode'></canvas>" & vbCrLf
    strText = strText & "    <script>" & vbCrLf
    strText = strText & "        JsBarcode('#barcode', '" & Data & "', {" & vbCrLf
    strText = strText & "            background:'#ffffff'," & vbCrLf
    strText = strText & "            displayValue:true," & vbCrLf
    strText = strText & "            format:'CODE39'," & vbCrLf
    strText = strText & "            font:'fantasy'," & vbCrLf
    strText = strText & "            fontOptions:'bold italic'," & vbCrLf
    strText = strText & "            fontSize:12," & vbCrLf
    strText = strText & "            height:75," & vbCrLf
    strText = strText & "            lineColor:'#000000'," & vbCrLf
    strText = strText & "            margin: 10," & vbCrLf
    strText = strText & "            marginBottom:10," & vbCrLf
    strText = strText & "            marginLeft:10," & vbCrLf
    strText = strText & "            marginRight:10," & vbCrLf
    strText = strText & "            marginTop:10," & vbCrLf
    strText = strText & "            textAlign:'center'," & vbCrLf
    strText = strText & "            textMargin:2," & vbCrLf
    strText = strText & "            textPosition:'bottom'," & vbCrLf
    strText = strText & "            width:2" & vbCrLf
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

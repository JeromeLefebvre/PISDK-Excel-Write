Attribute VB_Name = "Module1"
'   Copyright 2016 OSIsoft, LLC.
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'       http://www.apache.org/licenses/LICENSE-2.0
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

Option Explicit

Sub WriteDataToPI_Click()

On Error GoTo ErrH_Click
    
    ' Excel�̃V�[�g�̑I��
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Set sourceBook = ActiveWorkbook
    Set sourceSheet = sourceBook.Sheets("Sheet1")
    
    ' �V�[�g����f�[�^�̎��W
    Dim serverName As String
    Dim tagName As String
    Dim timestamp As String
    Dim value As String

    serverName = sourceSheet.Range("B1")
    tagName = sourceSheet.Range("B2")
    timestamp = sourceSheet.Range("B3")
    value = sourceSheet.Range("B4")
    
    'PI Server�ɐڑ����A�^�O�̒�`�̐ݒ�
    Dim myServer As PISDK.Server
    Dim myTag As PISDK.PIPoint
    Set myServer = Servers(serverName)
    Set myTag = myServer.PIPoints(tagName)
        
    ' �l�̓o�^
    myTag.Data.UpdateValue value, timestamp

    ' �I�u�W�F�N�g�Ɛڑ��̏���
    If myServer.Connected Then
      myServer.Close
    End If
    
    Set myServer = Nothing
    Set myTag = Nothing
    
Exit_Click:
    Exit Sub
        
ErrH_Click:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_Click
End Sub

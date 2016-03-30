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
    ' Select the sheet
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Set sourceBook = ActiveWorkbook
    Set sourceSheet = sourceBook.Sheets("Sheet1")
    
    ' �V�[�g����f�[�^�̎��W
    ' Get data from the sheet
    Dim serverName As String
    Dim tagName As String
    Dim timestamp As String
    Dim value As String

    serverName = sourceSheet.Range("B1")
    tagName = sourceSheet.Range("B2")
    timestamp = sourceSheet.Range("B3")
    value = sourceSheet.Range("B4")
    
    'PI Server�ɐڑ����A�^�O�̒�`�̐ݒ�
    'Connect to the PI Server and retrive the tag
    Dim myServer As PISDK.Server
    Dim myTag As PISDK.PIPoint
    
    Set myServer = Servers(serverName)
    'Explict Login���g���ꍇ�́A���[�U�[���ƃp�X���[�h�̐ݒ�
    'If explicit login is required specify the username and password
    'myServer.Open ("uid=piLoginDemo;pwd=!")
    
    Set myTag = myServer.PIPoints(tagName)

    ' �l�̓o�^
    ' Update the tag
    myTag.Data.UpdateValue value, timestamp
    
Exit_Click:
    Exit Sub
        
ErrH_Click:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_Click
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDraftReply
   Caption         =   "Draft Reply"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   OleObjectBlob   =   "frmDraftReply.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDraftReply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmDraftReply - LLM Draft Reply Viewer v2.0
'===============================================================================
' Simple UserForm to display an LLM-drafted reply with copy/create actions.
'
' SETUP INSTRUCTIONS:
' 1. In VBA Editor: Insert -> UserForm
' 2. Set form properties: Name=frmDraftReply, Caption="Draft Reply",
'    Width=400, Height=350
' 3. Add these controls:
'    - TextBox: txtDraft (MultiLine=True, ScrollBars=2, Width=370, Height=220)
'    - CommandButton: cmdCopy     Caption="Copy to Clipboard"
'    - CommandButton: cmdCreateReply  Caption="Create Reply Email"
'    - CommandButton: cmdClose    Caption="Close"
' 4. Paste this code into the UserForm's code-behind module
'===============================================================================

Option Explicit

Private originalMail As Outlook.MailItem

' Initialize the form with the draft text and original email reference
Public Sub Initialize(ByVal draftText As String, ByVal mail As Outlook.MailItem)
    txtDraft.text = draftText
    Set originalMail = mail
End Sub

Private Sub cmdCopy_Click()
    ' Copy draft text to clipboard using MSForms.DataObject
    Dim dataObj As MSForms.DataObject
    Set dataObj = New MSForms.DataObject
    dataObj.SetText txtDraft.text
    dataObj.PutInClipboard
    Set dataObj = Nothing

    MsgBox "Draft copied to clipboard!", vbInformation, "Copy"
End Sub

Private Sub cmdCreateReply_Click()
    On Error GoTo ErrorHandler

    If originalMail Is Nothing Then
        MsgBox "Original email reference lost. Please copy the text manually.", vbExclamation
        Exit Sub
    End If

    ' Create a reply to the original email
    Dim replyMail As Outlook.MailItem
    Set replyMail = originalMail.Reply

    ' Prepend the draft text to the reply body
    replyMail.Body = txtDraft.text & vbCrLf & vbCrLf & replyMail.Body

    ' Display the reply (don't send automatically)
    replyMail.Display

    ' Close this form
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "Error creating reply: " & Err.Description, vbCritical, "Create Reply"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

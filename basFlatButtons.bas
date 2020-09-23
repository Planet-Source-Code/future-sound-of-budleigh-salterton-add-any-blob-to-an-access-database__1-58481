Attribute VB_Name = "modFlatButtons"
Option Explicit

'to use the subroutine, you have to put a picturebox named picFlat
'on each form from where you will call this sub
'pictureboxe's properties should be: Index=0, Appearance = Flat,
'BorderStyle=fixed single, Visible=False

Public Sub MakeFlatButtons(frm As Form)
   Dim ButtonCount As Integer
   Dim c As Control
   On Error Resume Next

   For Each c In frm
      If TypeOf c Is CommandButton Then

         ButtonCount = ButtonCount + 1
      
         Load frm.picFlat(ButtonCount)
         With frm.picFlat(ButtonCount)
            .Visible = True
            .Left = c.Left
            .Top = c.Top
            .Width = c.Width
            .Height = c.Height
         End With
         
         Set c.Container = frm.picFlat(ButtonCount)
         c.Left = 0
         c.Top = 0
      End If
   Next c

End Sub

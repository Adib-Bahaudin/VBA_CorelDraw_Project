VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub copi_code()

    Dim Var1 As Shape
    Dim Var2 As Shape
    Dim Var3 As Shape
    Dim Var4 As Shape
    
    ' Pastikan ada objek yang dipilih
    If ActiveShape Is Nothing Then
        MsgBox "Tidak ada objek yang dipilih.", vbExclamation
        Exit Sub
    End If
    
    ' copi objek
    Set Var1 = ActiveShape
    Set Var2 = Var1.Duplicate
    Set Var3 = Var1.Duplicate
    Set Var4 = Var1.Duplicate
    
    ActiveDocument.Unit = cdrMillimeter             ' merubah satuan menjadi mm
    ActiveDocument.ReferencePoint = cdrCenter       ' fokus objek berada di center
    ActiveDocument.MasterPage.SetSize 148#, 210#    ' Merubah ukuran lembar kerja
    
    ' menyusun objek
    Offset = Var1.SizeWidth + 2
    
    Var2.Move Offset, 0
    Var3.Move (Offset * 2), 0
    Var4.Move (Offset * 3), 0
    
    ' memilih objek
    ActiveDocument.ClearSelection   ' kosongkan selection dulu
    ActiveDocument.AddToSelection Var1, Var2, Var3, Var4
    
    ' menyatukan objek
    Dim Varall As Shape
    Set Varall = ActiveSelection.Group
    
    ' set objek ketengah halaman
    Varall.AlignAndDistribute 3, 3, 2, 0, False, 2

End Sub

Sub Cetak_Pas_Foto()

    If ActiveShape Is Nothing Then
        MsgBox "Tidak ada objek yang dipilih.", vbExclamation
        Exit Sub
    End If
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.MasterPage.SetSize 148#, 210#
    
    If ActiveShape.SizeWidth = 40 And ActiveShape.SizeHeight = 60 Then
    
        Dim En As form_4x6
        Set En = New form_4x6
        En.Show
        
        If En.IsCancelled Then
            Exit Sub
        End If
        
        Call Empat_Enam(En.Jumlah)
        
    ElseIf ActiveShape.SizeWidth = 30 And ActiveShape.SizeHeight = 40 Then
    
        Dim Em As form_3x4
        Set Em = New form_3x4
        Em.Show
        
        If Em.IsCancelled Then
            Exit Sub
        End If
        
        Call Tiga_Empat(Em.Jumlah)
    
    ElseIf ActiveShape.SizeWidth = 20 And ActiveShape.SizeHeight = 30 Then
    
        Dim Ti As form_2x3
        Set Ti = New form_2x3
        Ti.Show
        
        If Ti.IsCancelled Then
            Exit Sub
        End If
        
        Call Dua_Tiga(Ti.Jumlah)
        
    Else
        MsgBox "Ukuran Pas Foto Tidak Valid.", vbExclamation
    End If
End Sub

Private Sub Dua_Tiga(Jumlah As Integer)

    Dim Var() As Shape
    Dim i As Integer
    
    ReDim Var(1 To 6)
    
    Set Var(1) = ActiveShape
    
    For i = 2 To 6
        Set Var(i) = Var(1).Duplicate
    Next i
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    
    Dim Offset As Double
    Dim H_Offset As Double
    Offset = Var(1).SizeWidth + 2
    H_Offset = Var(1).SizeHeight + 2
    
    If Jumlah >= 6 Then
        Var(2).Move Offset, 0
        Var(3).Move (Offset * 2), 0
        Var(4).Move (Offset * 3), 0
        Var(5).Move (Offset * 4), 0
        Var(6).Move (Offset * 5), 0
    End If
    
    ActiveDocument.ClearSelection
    ActiveDocument.AddToSelection Var()
    
    Dim Varall As Shape
    Set Varall = ActiveSelection.Group
    
    Varall.AlignAndDistribute 3, 3, 2, 0, False, 2
    
    If Jumlah >= 12 Then
    
        Dim Var12 As Shape
        Set Var12 = Varall.Duplicate
        
        Var12.Move 0, H_Offset
        
        If Jumlah = 12 Then
            
            ActiveDocument.ClearSelection
            ActiveDocument.AddToSelection Varall, Var12
            
        ElseIf Jumlah >= 18 Then
            
            Dim Var18 As Shape
            Set Var18 = Varall.Duplicate
            
            Var18.Move 0, (H_Offset * 2)
            
            If Jumlah = 18 Then
            
                ActiveDocument.ClearSelection
                ActiveDocument.AddToSelection Varall, Var12, Var18
            
            ElseIf Jumlah >= 24 Then
                
                Dim Var24 As Shape
                Set Var24 = Varall.Duplicate
                
                Var24.Move 0, (H_Offset * 3)
                
                If Jumlah = 24 Then
                
                    ActiveDocument.ClearSelection
                    ActiveDocument.AddToSelection Varall, Var12, Var18, Var24
                    
                ElseIf Jumlah >= 30 Then
                
                    Dim Var30 As Shape
                    Set Var30 = Varall.Duplicate
                
                    Var30.Move 0, (H_Offset * 4)
                    
                    If Jumlah = 30 Then
                    
                        ActiveDocument.ClearSelection
                        ActiveDocument.AddToSelection Varall, Var12, Var18, Var24, Var30
                    
                    ElseIf Jumlah = 36 Then
                    
                        Dim Var36 As Shape
                        Set Var36 = Varall.Duplicate
                
                        Var36.Move 0, (H_Offset * 5)
                        
                        ActiveDocument.ClearSelection
                        ActiveDocument.AddToSelection Varall, Var12, Var18, Var24, Var30, Var36
                        
                    End If
                End If
            End If
        End If
    
    Dim Varfixall As Shape
    Set Varfixall = ActiveSelection.Group
    
    Varfixall.AlignAndDistribute 3, 3, 2, 0, False, 2
    
    End If
End Sub

Private Sub Tiga_Empat(Jumlah As Integer)

    Dim Var() As Shape
    Dim i As Integer
    
    ReDim Var(1 To 4)
    
    Set Var(1) = ActiveShape
    
    For i = 2 To 4
        Set Var(i) = Var(1).Duplicate
    Next i
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    
    Dim Offset As Double
    Dim H_Offset As Double
    Offset = Var(1).SizeWidth + 2
    H_Offset = Var(1).SizeHeight + 2
    
    If Jumlah >= 4 Then
        Var(2).Move Offset, 0
        Var(3).Move (Offset * 2), 0
        Var(4).Move (Offset * 3), 0
    End If
    
    ActiveDocument.ClearSelection
    ActiveDocument.AddToSelection Var()
    
    Dim Varall As Shape
    Set Varall = ActiveSelection.Group
    
    Varall.AlignAndDistribute 3, 3, 2, 0, False, 2
    
    If Jumlah >= 8 Then
        
        Dim Varall8 As Shape
        Set Varall8 = Varall.Duplicate
        
        Varall8.Move 0, H_Offset
        
        If Jumlah = 8 Then
        
            ActiveDocument.ClearSelection
            ActiveDocument.AddToSelection Varall, Varall8
        
        ElseIf Jumlah >= 12 Then
        
            Dim Varall12 As Shape
            Set Varall12 = Varall.Duplicate
        
            Varall12.Move 0, (H_Offset * 2)
            
            If Jumlah = 12 Then
            
                ActiveDocument.ClearSelection
                ActiveDocument.AddToSelection Varall, Varall8, Varall12
            
            ElseIf Jumlah = 16 Then
            
                Dim Varall16 As Shape
                Set Varall16 = Varall.Duplicate

                Varall16.Move 0, (H_Offset * 3)

                ActiveDocument.ClearSelection
                ActiveDocument.AddToSelection Varall, Varall8, Varall12, Varall16
                
            End If
        End If
        
        Dim Varfixall As Shape
        Set Varfixall = ActiveSelection.Group
        
        Varfixall.AlignAndDistribute 3, 3, 2, 0, False, 2
        
    End If
End Sub

Private Sub Empat_Enam(Jumlah As Integer)

    Dim Var() As Shape
    Dim i As Integer
    
    ReDim Var(1 To 3)
    
    Set Var(1) = ActiveShape
    
    For i = 2 To 3
        Set Var(i) = Var(1).Duplicate
    Next i
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    
    Dim Offset As Integer
    Dim H_Offset As Integer
    Offset = Var(1).SizeWidth + 2
    H_Offset = Var(1).SizeHeight + 2
    
    If Jumlah >= 3 Then
        Var(2).Move Offset, 0
        Var(3).Move (Offset * 2), 0
    End If
    
    ActiveDocument.ClearSelection
    ActiveDocument.AddToSelection Var()
    
    Dim Varall As Shape
    Set Varall = ActiveSelection.Group
    
    Varall.AlignAndDistribute 3, 3, 2, 0, False, 2
    
    If Jumlah >= 6 Then
    
        Dim Varall6 As Shape
        Set Varall6 = Varall.Duplicate
        
        Varall6.Move 0, H_Offset
        
        If Jumlah = 9 Then
            
            Dim Varall9 As Shape
            Set Varall9 = Varall.Duplicate
            
            Varall9.Move 0, (H_Offset * 2)
            
            ActiveDocument.ClearSelection
            ActiveDocument.AddToSelection Varall, Varall6, Varall9
            
            Dim Var9all As Shape
            Set Var9all = ActiveSelection.Group
        
            Var9all.AlignAndDistribute 3, 3, 2, 0, False, 2
            
            Exit Sub
        
        End If
        
        ActiveDocument.ClearSelection
        ActiveDocument.AddToSelection Varall, Varall6
        
        Dim Var6all As Shape
        Set Var6all = ActiveSelection.Group
        
        Var6all.AlignAndDistribute 3, 3, 2, 0, False, 2
        
    End If
End Sub

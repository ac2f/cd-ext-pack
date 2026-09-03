Attribute VB_Name = "ac2fLength"
'=====================================================================
'  ac2f pack  --  ac2fLength
'
'  Vektörlerin çevresindeki (kontur) çizgilerin toplam uzunluğunu ölçer.
'
'  Makrolar:
'    ac2fUzunlukOlc       - seçimi ölçer ve raporlar
'    ac2fUzunlukEtiketle  - seçimi ölçer ve sonucu sayfaya metin olarak koyar
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Uzunluk Ölçümü"
Private Const MAX_DETAY As Long = 15

'---------------------------------------------------------------------
' Seçili nesnelerin toplam kontur uzunluğunu ölçer ve rapor gösterir.
'---------------------------------------------------------------------
Public Sub ac2fUzunlukOlc()
Attribute ac2fUzunlukOlc.VB_Description = "ac2f pack: Secili vektorlerin toplam kontur uzunlugunu olcer"
    Dim res As ac2fResult

    res = ac2fMeasureSelection()
    If Not res.Ok Then
        ac2fWarn res.Message, CAPTION_
        Exit Sub
    End If

    If res.SubCount = 0 Then
        ac2fWarn "Seçimde ölçülebilir bir yol bulunamadı." & vbCrLf & vbCrLf & _
                 "Bitmap, boş metin veya uzunluğu sıfır olan nesneler ölçülemez.", CAPTION_
        Exit Sub
    End If

    ac2fInfo ac2fUzunlukRaporu(res), CAPTION_
End Sub

'---------------------------------------------------------------------
' Ölçüm sonucunu sayfaya artistik metin olarak yerleştirir.
'---------------------------------------------------------------------
Public Sub ac2fUzunlukEtiketle()
Attribute ac2fUzunlukEtiketle.VB_Description = "ac2f pack: Olcum sonucunu sayfaya metin olarak ekler"
    Dim res As ac2fResult
    Dim sr As ShapeRange
    Dim txt As String

    If ActiveDocument Is Nothing Then
        ac2fWarn "Önce bir belge açın.", CAPTION_
        Exit Sub
    End If

    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Ölçülecek nesne seçilmedi.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "Ölçülecek nesne seçilmedi.", CAPTION_
        Exit Sub
    End If

    res = ac2fMeasureRange(sr)
    If Not res.Ok Then
        ac2fWarn res.Message, CAPTION_
        Exit Sub
    End If
    If res.SubCount = 0 Then
        ac2fWarn "Seçimde ölçülebilir bir yol bulunamadı.", CAPTION_
        Exit Sub
    End If

    txt = "Toplam uzunluk: " & ac2fFmt(res.TotalMM) & " mm" & _
          "  (" & ac2fFmt(res.TotalMM / 1000#, 2) & " m)"

    If ac2fEtiketYerlestir(sr, txt) Then
        ac2fInfo "Etiket sayfaya eklendi." & vbCrLf & vbCrLf & txt, CAPTION_
    Else
        ac2fWarn "Etiket eklenemedi, sonuç yalnızca burada gösteriliyor." & _
                 vbCrLf & vbCrLf & txt, CAPTION_
    End If
End Sub

'=====================================================================
' İç fonksiyonlar
'=====================================================================

' Ölçüm sonucundan okunabilir bir rapor metni üretir.
Public Function ac2fUzunlukRaporu(ByRef res As ac2fResult) As String
    Dim s As String
    Dim i As Long
    Dim enUzun As Double, enKisa As Double

    enKisa = -1
    For i = 0 To res.SubCount - 1
        If res.SubLenMM(i) > enUzun Then enUzun = res.SubLenMM(i)
        If enKisa < 0 Or res.SubLenMM(i) < enKisa Then enKisa = res.SubLenMM(i)
    Next i
    If enKisa < 0 Then enKisa = 0

    s = "TOPLAM UZUNLUK" & vbCrLf
    s = s & "   " & ac2fFmtLength(res.TotalMM) & vbCrLf & vbCrLf

    s = s & "SEÇİM ÖZETİ" & vbCrLf
    s = s & "   Nesne sayısı        : " & res.ShapeCount & vbCrLf
    s = s & "   Kontur (alt yol)    : " & res.SubCount & vbCrLf
    s = s & "      kapalı           : " & (res.SubCount - res.OpenCount) & vbCrLf
    s = s & "      açık             : " & res.OpenCount & vbCrLf & vbCrLf

    s = s & "KONTUR İSTATİSTİĞİ" & vbCrLf
    s = s & "   En uzun             : " & ac2fFmt(enUzun) & " mm" & vbCrLf
    s = s & "   En kısa             : " & ac2fFmt(enKisa) & " mm" & vbCrLf
    s = s & "   Ortalama            : " & ac2fFmt(res.TotalMM / res.SubCount) & " mm" & vbCrLf

    If res.ItemCount > 0 Then
        s = s & vbCrLf & "NESNE KIRILIMI"
        If res.ItemCount > MAX_DETAY Then
            s = s & " (ilk " & MAX_DETAY & " / " & res.ItemCount & ")"
        End If
        s = s & vbCrLf

        For i = 0 To res.ItemCount - 1
            If i >= MAX_DETAY Then Exit For
            s = s & "   " & (i + 1) & ". " & res.Items(i).Name & " - " & _
                    ac2fFmt(res.Items(i).LengthMM) & " mm" & _
                    " (" & res.Items(i).SubPaths & " kontur)" & vbCrLf
        Next i
    End If

    ac2fUzunlukRaporu = s
End Function

' Seçimin altına artistik metin ekler. Başarılıysa True döner.
Private Function ac2fEtiketYerlestir(ByVal sr As ShapeRange, ByVal txt As String) As Boolean
    Dim oldUnit As cdrUnit
    Dim unitChanged As Boolean
    Dim x As Double, y As Double, w As Double, h As Double
    Dim t As Shape

    On Error GoTo Fail

    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    sr.GetBoundingBox x, y, w, h

    ActiveDocument.BeginCommandGroup ac2fTitle("etiket")
    Set t = ActiveLayer.CreateArtisticText(x, y - 8#, txt)
    On Error Resume Next
    t.Text.Story.Size = 10
    On Error GoTo Fail
    ActiveDocument.EndCommandGroup

    ac2fEtiketYerlestir = True

Cleanup:
    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Exit Function

Fail:
    ac2fEtiketYerlestir = False
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    Resume Cleanup
End Function

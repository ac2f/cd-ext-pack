Attribute VB_Name = "ac2fCore"
'=====================================================================
'  ac2f pack  --  ac2fCore
'
'  Ortak sabitler, ayar yönetimi, yardımcı fonksiyonlar ve geometri
'  ölçüm çekirdeği.
'
'  Bu modül doğrudan çalıştırılmaz; ac2fLength, ac2fLedModule ve
'  ac2fMenu modülleri tarafından kullanılır.
'=====================================================================
Option Explicit

'---------------------------------------------------------------------
' Paket kimliği
'---------------------------------------------------------------------
Public Const AC2F_NAME    As String = "ac2f pack"
Public Const AC2F_VERSION As String = "1.0.0"
Public Const AC2F_REG_APP As String = "ac2fPack"
Public Const AC2F_REG_SEC As String = "Ayarlar"

'---------------------------------------------------------------------
' Varsayılan ayarlar (ac2fAyarlar ile değiştirilebilir)
'---------------------------------------------------------------------
Public Const AC2F_DEF_SPACING    As Double = 100#  ' modül aralığı (mm)
Public Const AC2F_DEF_LEDS       As Long = 3       ' modül başına LED adedi
Public Const AC2F_DEF_MODULE_W   As Double = 0.72  ' modül gücü (W)
Public Const AC2F_DEF_PSU_W      As Double = 60#   ' güç kaynağı kapasitesi (W)
Public Const AC2F_DEF_SAFETY     As Double = 20#   ' güvenlik payı (%)
Public Const AC2F_DEF_MINPERPATH As Long = 1       ' alt yol başına en az modül
Public Const AC2F_DEF_METHOD     As Long = 1       ' 1 = çevre, 2 = orta hat tahmini
Public Const AC2F_DEF_FACTOR     As Double = 1#    ' düzeltme katsayısı

'---------------------------------------------------------------------
' Ayar anahtarları
'---------------------------------------------------------------------
Public Const AC2F_K_SPACING    As String = "ModulAraligiMM"
Public Const AC2F_K_LEDS       As String = "ModulBasinaLed"
Public Const AC2F_K_MODULE_W   As String = "ModulGucuW"
Public Const AC2F_K_PSU_W      As String = "GucKaynagiW"
Public Const AC2F_K_SAFETY     As String = "GuvenlikPayiYuzde"
Public Const AC2F_K_MINPERPATH As String = "AltYolBasinaEnAzModul"
Public Const AC2F_K_METHOD     As String = "HesapYontemi"
Public Const AC2F_K_FACTOR     As String = "DuzeltmeKatsayisi"

'---------------------------------------------------------------------
' Veri yapıları
'---------------------------------------------------------------------

' Ölçülen tek bir nesnenin özeti
Public Type ac2fItem
    Name     As String
    LengthMM As Double
    SubPaths As Long
End Type

' Bir seçimin tüm ölçüm sonucu
Public Type ac2fResult
    Ok          As Boolean
    Message     As String
    TotalMM     As Double
    ShapeCount  As Long
    SubCount    As Long
    OpenCount   As Long
    ItemCount   As Long
    Items()     As ac2fItem
    SubLenMM()  As Double
    SubClosed() As Boolean
End Type

' Ölçüm sırasında oluşturulan geçici kopya sayısı.
Private m_TempShapes As Long

'=====================================================================
' AYARLAR  (Windows kayıt defteri: VB and VBA Program Settings)
'=====================================================================

Public Function ac2fGetStr(ByVal key As String, ByVal defValue As String) As String
    Dim v As String
    On Error Resume Next
    v = GetSetting(AC2F_REG_APP, AC2F_REG_SEC, key, defValue)
    If Len(Trim$(v)) = 0 Then v = defValue
    ac2fGetStr = v
End Function

Public Sub ac2fSetStr(ByVal key As String, ByVal value As String)
    On Error Resume Next
    SaveSetting AC2F_REG_APP, AC2F_REG_SEC, key, value
End Sub

' Sayılar yerelden bağımsız olsun diye her zaman nokta ondalıklı saklanır.
Public Function ac2fGetNum(ByVal key As String, ByVal defValue As Double) As Double
    Dim s As String
    s = ac2fGetStr(key, "")
    If Len(s) = 0 Then
        ac2fGetNum = defValue
    Else
        ac2fGetNum = Val(s)
    End If
End Function

Public Sub ac2fSetNum(ByVal key As String, ByVal value As Double)
    ac2fSetStr key, ac2fNumStr(value)
End Sub

Public Function ac2fGetLng(ByVal key As String, ByVal defValue As Long) As Long
    ac2fGetLng = CLng(ac2fGetNum(key, CDbl(defValue)))
End Function

Public Sub ac2fSetLng(ByVal key As String, ByVal value As Long)
    ac2fSetStr key, ac2fNumStr(CDbl(value))
End Sub

' Tüm ayarları fabrika değerlerine döndürür.
Public Sub ac2fResetSettings()
    On Error Resume Next
    DeleteSetting AC2F_REG_APP, AC2F_REG_SEC
End Sub

'=====================================================================
' YARDIMCI FONKSİYONLAR
'=====================================================================

' Sayıyı yerelden bağımsız metne çevirir. Str$ baştaki sıfırı attığı için
' ".72" yerine "0.72" üretilir; hem kayıt hem de InputBox varsayılanı için.
Public Function ac2fNumStr(ByVal v As Double) As String
    Dim s As String
    s = Trim$(Str$(v))
    If Left$(s, 1) = "." Then
        s = "0" & s
    ElseIf Left$(s, 2) = "-." Then
        s = "-0" & Mid$(s, 2)
    End If
    ac2fNumStr = s
End Function

' Kullanıcıdan gelen metni sayıya çevirir. Hem "12,5" hem "12.5" kabul edilir.
' Binlik ayırıcı kullanılmamalıdır.
Public Function ac2fParseNum(ByVal s As String, ByVal defValue As Double) As Double
    Dim t As String, sep As String
    t = Trim$(s)
    If Len(t) = 0 Then
        ac2fParseNum = defValue
        Exit Function
    End If
    sep = Mid$(CStr(1.5), 2, 1)          ' yerel ondalık ayırıcı
    t = Replace$(t, ".", sep)
    t = Replace$(t, ",", sep)
    If Not IsNumeric(t) Then
        ac2fParseNum = defValue
        Exit Function
    End If
    ac2fParseNum = CDbl(t)
End Function

' Pozitif değerler için yukarı yuvarlama.
Public Function ac2fCeil(ByVal v As Double) As Double
    If v - Int(v) > 0.000001 Then
        ac2fCeil = Int(v) + 1#
    Else
        ac2fCeil = Int(v)
    End If
End Function

Public Function ac2fFmt(ByVal v As Double, Optional ByVal dec As Long = 2) As String
    If dec > 0 Then
        ac2fFmt = Format$(v, "#,##0." & String$(dec, "0"))
    Else
        ac2fFmt = Format$(v, "#,##0")
    End If
End Function

' Bir uzunluğu mm / cm / m olarak tek satırda gösterir.
Public Function ac2fFmtLength(ByVal mm As Double) As String
    ac2fFmtLength = ac2fFmt(mm) & " mm   |   " & _
                    ac2fFmt(mm / 10#) & " cm   |   " & _
                    ac2fFmt(mm / 1000#, 3) & " m"
End Function

Public Function ac2fTitle(ByVal caption As String) As String
    ac2fTitle = AC2F_NAME & " " & AC2F_VERSION & " - " & caption
End Function

Public Sub ac2fInfo(ByVal msg As String, ByVal caption As String)
    MsgBox msg, vbInformation, ac2fTitle(caption)
End Sub

Public Sub ac2fWarn(ByVal msg As String, ByVal caption As String)
    MsgBox msg, vbExclamation, ac2fTitle(caption)
End Sub

'=====================================================================
' ÖLÇÜM ÇEKİRDEĞİ
'=====================================================================

' Aktif seçimi ölçer. Belge ya da seçim yoksa Ok = False döner.
Public Function ac2fMeasureSelection() As ac2fResult
    Dim res As ac2fResult
    Dim sr As ShapeRange

    If ActiveDocument Is Nothing Then
        res.Message = "Önce bir belge açın."
        ac2fMeasureSelection = res
        Exit Function
    End If

    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        res.Message = "Ölçülecek nesne seçilmedi."
        ac2fMeasureSelection = res
        Exit Function
    End If
    If sr.Count = 0 Then
        res.Message = "Ölçülecek nesne seçilmedi."
        ac2fMeasureSelection = res
        Exit Function
    End If

    ac2fMeasureSelection = ac2fMeasureRange(sr)
End Function

' Verilen ShapeRange içindeki tüm yolların toplam uzunluğunu (mm) ölçer.
' Gruplar ve PowerClip içerikleri özyinelemeli olarak taranır.
Public Function ac2fMeasureRange(ByVal sr As ShapeRange) As ac2fResult
    Dim res As ac2fResult
    Dim oldUnit As cdrUnit
    Dim oldOpt As Boolean
    Dim unitChanged As Boolean
    Dim groupOpen As Boolean
    Dim i As Long

    ReDim res.Items(0 To 63)
    ReDim res.SubLenMM(0 To 255)
    ReDim res.SubClosed(0 To 255)

    m_TempShapes = 0

    On Error GoTo Fail

    oldOpt = Application.Optimization
    Application.Optimization = True

    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    ActiveDocument.BeginCommandGroup ac2fTitle("ölçüm")
    groupOpen = True

    For i = 1 To sr.Count
        ac2fCollect sr(i), res
    Next i

    ActiveDocument.EndCommandGroup
    groupOpen = False

    res.Ok = True
    res.Message = ""

Cleanup:
    On Error Resume Next
    If groupOpen Then ActiveDocument.EndCommandGroup
    ' Geçici kopyalar zaten silindi; yine de belgeyi kesin olarak
    ' ilk haline döndürmek için komut grubu geri alınır. Hiç geçici
    ' nesne üretilmediyse kullanıcının önceki işlemini geri almamak
    ' adına Undo çağrılmaz.
    If m_TempShapes > 0 Then ActiveDocument.Undo
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    m_TempShapes = 0
    ac2fMeasureRange = res
    Exit Function

Fail:
    res.Ok = False
    res.Message = "Ölçüm sırasında hata oluştu: " & Err.Description
    Resume Cleanup
End Function

'---------------------------------------------------------------------
' Tek bir şekli sonuca ekler (özyinelemeli).
'---------------------------------------------------------------------
Private Sub ac2fCollect(ByVal s As Shape, ByRef res As ac2fResult)
    Dim i As Long
    Dim cv As Curve
    Dim dup As Shape

    If s Is Nothing Then Exit Sub

    On Error GoTo Skip

    ' --- Yol taşımayan nesneler: DisplayCurve bunlarda çerçeveyi
    '     döndürebileceği için baştan elenir. ---
    Select Case s.Type
        Case cdrBitmapShape, cdrOLEObjectShape
            Exit Sub
    End Select

    ' --- Grup: içeriğini tara ---
    If s.Type = cdrGroupShape Then
        For i = 1 To s.Shapes.Count
            ac2fCollect s.Shapes(i), res
        Next i
        Exit Sub
    End If

    ' --- PowerClip içeriği varsa onu da tara ---
    ac2fCollectPowerClip s, res

    ' --- Eğri temsilini elde et ---
    Set cv = Nothing
    On Error Resume Next
    Set cv = s.DisplayCurve
    On Error GoTo Skip

    If Not cv Is Nothing Then
        ac2fAddCurve cv, ac2fShapeName(s), res
        Exit Sub
    End If

    ' DisplayCurve yoksa geçici kopya üzerinden eğriye çevir.
    Set dup = s.Duplicate(0, 0)
    m_TempShapes = m_TempShapes + 1

    On Error Resume Next
    dup.ConvertToCurves
    On Error GoTo SkipDup

    If dup.Type = cdrGroupShape Then
        ac2fCollect dup, res
        dup.Delete
        Exit Sub
    End If

    Set cv = Nothing
    On Error Resume Next
    Set cv = dup.Curve
    On Error GoTo SkipDup

    If cv Is Nothing Then GoTo SkipDup

    ac2fAddCurve cv, ac2fShapeName(s), res
    dup.Delete
    Exit Sub

SkipDup:
    On Error Resume Next
    If Not dup Is Nothing Then dup.Delete
    Exit Sub

Skip:
    ' Ölçülemeyen nesne sessizce atlanır.
End Sub

Private Sub ac2fCollectPowerClip(ByVal s As Shape, ByRef res As ac2fResult)
    Dim pc As Object
    Dim i As Long

    On Error Resume Next
    Set pc = s.PowerClip
    If pc Is Nothing Then Exit Sub

    For i = 1 To pc.Shapes.Count
        ac2fCollect pc.Shapes(i), res
    Next i
End Sub

Private Function ac2fShapeName(ByVal s As Shape) As String
    Dim n As String
    On Error Resume Next
    n = s.Name
    If Len(Trim$(n)) = 0 Then n = "Nesne " & CStr(s.StaticID)
    ac2fShapeName = n
End Function

'---------------------------------------------------------------------
' Bir eğrinin alt yollarını sonuca yazar.
'---------------------------------------------------------------------
Private Sub ac2fAddCurve(ByVal cv As Curve, ByVal shapeName As String, ByRef res As ac2fResult)
    Dim i As Long
    Dim n As Long
    Dim spLen As Double
    Dim shpLen As Double
    Dim used As Long

    On Error Resume Next
    n = cv.SubPaths.Count
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    If n = 0 Then Exit Sub

    For i = 1 To n
        spLen = 0
        On Error Resume Next
        spLen = cv.SubPaths(i).Length
        On Error GoTo 0

        If spLen > 0 Then
            shpLen = shpLen + spLen
            used = used + 1

            If res.SubCount > UBound(res.SubLenMM) Then
                ReDim Preserve res.SubLenMM(0 To (UBound(res.SubLenMM) + 1) * 2 - 1)
                ReDim Preserve res.SubClosed(0 To UBound(res.SubLenMM))
            End If

            res.SubLenMM(res.SubCount) = spLen
            res.SubClosed(res.SubCount) = ac2fIsClosed(cv, i)
            If Not res.SubClosed(res.SubCount) Then res.OpenCount = res.OpenCount + 1
            res.SubCount = res.SubCount + 1
        End If
    Next i

    If shpLen <= 0 Then Exit Sub

    If res.ItemCount > UBound(res.Items) Then
        ReDim Preserve res.Items(0 To (UBound(res.Items) + 1) * 2 - 1)
    End If

    res.Items(res.ItemCount).Name = shapeName
    res.Items(res.ItemCount).LengthMM = shpLen
    res.Items(res.ItemCount).SubPaths = used
    res.ItemCount = res.ItemCount + 1

    res.ShapeCount = res.ShapeCount + 1
    res.TotalMM = res.TotalMM + shpLen
End Sub

Private Function ac2fIsClosed(ByVal cv As Curve, ByVal idx As Long) As Boolean
    Dim b As Boolean
    On Error Resume Next
    b = cv.SubPaths(idx).Closed
    ac2fIsClosed = b
End Function

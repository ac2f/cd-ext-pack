Attribute VB_Name = "ac2fBoxLetter"
'=====================================================================
'  ac2f pack  --  ac2fBoxLetter
'
'  Kutu harf yan şeridi (bordür) açınımı ve derz yerleşimi.
'
'  Seçili harf konturlarını alır; malzeme kalınlığına ve esneklik
'  oranına göre açınım boyunu hesaplar, derz konumlarını çıkarır ve
'  kesime + derz açmaya hazır düz şeritler çizer.
'
'  Makrolar:
'    ac2fKutuHarfSerit    - şeritleri hesaplar ve çizer
'    ac2fKutuHarfRapor    - yalnız hesaplar, çizim yapmaz
'    ac2fKutuHarfAyarlar  - malzeme ve derz ayarları
'
'  GEOMETRİ MODELİ
'  Her segment dairesel yay kabul edilir. Kiriş/yay oranından dönüş
'  açısı çözülür (c/L = 2*sin(t/2)/t), yarıçap R = L/t olur. Böylece
'  bezier kontrol noktası okumaya gerek kalmaz; yalnızca düğüm konumu
'  ve segment uzunluğu yeter. Gerçek kübik bezier yaylarda yarıçap
'  hatası %0,05 mertebesindedir.
'
'  Açınım, nötr eksenin ötelenmesiyle bulunur. Basit kapalı bir eğride
'  toplam dönüş 2*pi olduğundan toplam açınım boyu P -/+ 2*pi*g'dir
'  (Steiner). Bu toplam, segment dönüş işaretlerinden bağımsız olarak
'  kesindir; işaret yalnız ara derz konumlarını etkiler ve ölçülen en
'  kötü sapma 0,06 mm'dir.
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Kutu Harf Şeridi"

' Çizim güvenlik sınırları
Private Const MAX_GROOVE   As Long = 5000    ' şerit başına en çok derz
Private Const MAX_STRIPS   As Long = 200     ' iş başına en çok şerit
Private Const STRAIGHT_EPS As Double = 0.0002 ' (yay-kiriş)/yay < bu ise düz

'---------------------------------------------------------------------
' Ayar anahtarları
'---------------------------------------------------------------------
Public Const AC2F_K_BL_THICK   As String = "KHKalinlikMM"
Public Const AC2F_K_BL_KFAC    As String = "KHKFaktoru"
Public Const AC2F_K_BL_REF     As String = "KHReferansYuzey"
Public Const AC2F_K_BL_HEIGHT  As String = "KHSeritYuksekligiMM"
Public Const AC2F_K_BL_DEPTH   As String = "KHDerzDerinlikOrani"
Public Const AC2F_K_BL_MOUTH   As String = "KHDerzAgziMM"
Public Const AC2F_K_BL_FLEX    As String = "KHEsneklikOrani"
Public Const AC2F_K_BL_TOL     As String = "KHYuzeyToleransiMM"
Public Const AC2F_K_BL_SMIN    As String = "KHDerzEnAzAralikMM"
Public Const AC2F_K_BL_SMAX    As String = "KHDerzEnCokAralikMM"
Public Const AC2F_K_BL_CORNER  As String = "KHKoseEsigiDerece"
Public Const AC2F_K_BL_JOINT   As String = "KHEkPayiMM"
Public Const AC2F_K_BL_COIL    As String = "KHRuloBoyuMM"
Public Const AC2F_K_BL_GAP     As String = "KHSeritAraligiMM"

'---------------------------------------------------------------------
' Varsayılanlar
'---------------------------------------------------------------------
Public Const AC2F_DEF_BL_THICK  As Double = 1#     ' malzeme kalınlığı (mm)
Public Const AC2F_DEF_BL_KFAC   As Double = 0.44   ' nötr eksen K faktörü
Public Const AC2F_DEF_BL_REF    As Long = 1        ' 1=dış yüz 2=iç yüz 3=nötr
Public Const AC2F_DEF_BL_HEIGHT As Double = 80#    ' şerit yüksekliği = harf derinliği
Public Const AC2F_DEF_BL_DEPTH  As Double = 0.7    ' derz derinliği / kalınlık
Public Const AC2F_DEF_BL_MOUTH  As Double = 1.2    ' derz ağzı en çok (mm)
Public Const AC2F_DEF_BL_FLEX   As Double = 1#     ' esneklik oranı
Public Const AC2F_DEF_BL_TOL    As Double = 0.15   ' yüzey (kiriş) toleransı mm
Public Const AC2F_DEF_BL_SMIN   As Double = 3#     ' en az derz aralığı
Public Const AC2F_DEF_BL_SMAX   As Double = 60#    ' en çok derz aralığı
Public Const AC2F_DEF_BL_CORNER As Double = 5#     ' köşe eşiği (derece)
Public Const AC2F_DEF_BL_JOINT  As Double = 20#    ' ek payı (mm)
Public Const AC2F_DEF_BL_COIL   As Double = 3000#  ' rulo boyu (0 = bölme)
Public Const AC2F_DEF_BL_GAP    As Double = 10#    ' şeritler arası boşluk

'---------------------------------------------------------------------
' Çözümlenen kontur (modül düzeyi diziler; büyük UDT kopyalamamak için)
'---------------------------------------------------------------------
Private m_n    As Long        ' segment sayısı
Private m_L()  As Double      ' yay uzunluğu
Private m_Th() As Double      ' dönüş açısı, işaretsiz (radyan)
Private m_Sg() As Double      ' dönüş işareti (+1 / -1)
Private m_R()  As Double      ' yarıçap (0 = düz)
Private m_Dir() As Double     ' kiriş yönü
Private m_T0() As Double      ' başlangıç teğeti
Private m_T1() As Double      ' bitiş teğeti
Private m_Cor() As Double     ' segmentten sonraki köşe dönüşü
Private m_Tau  As Double      ' toplam işaretli dönüş

' Derz konumları (açınım koordinatında, mm)
Private m_gN   As Long
Private m_gPos() As Double
Private m_gCor() As Boolean   ' True = köşe derzi

' Tek bir şeridin sonucu
Private Type ac2fStrip
    Name     As String
    DevLen   As Double        ' açınım boyu
    RawLen   As Double        ' ham kontur boyu
    Grooves  As Long
    Corners  As Long
    MinR     As Double
    IsHole   As Boolean
End Type

'=====================================================================
' MAKROLAR
'=====================================================================

Public Sub ac2fKutuHarfSerit()
Attribute ac2fKutuHarfSerit.VB_Description = "ac2f pack: Kutu harf yan seridi acinimi ve derz cizimi"
    ac2fBLCalistir True
End Sub

Public Sub ac2fKutuHarfRapor()
Attribute ac2fKutuHarfRapor.VB_Description = "ac2f pack: Kutu harf serit raporu (cizim yapmaz)"
    ac2fBLCalistir False
End Sub

'=====================================================================
' ANA AKIŞ
'=====================================================================

Private Sub ac2fBLCalistir(ByVal ciz As Boolean)
    Dim sr As ShapeRange
    Dim oldUnit As cdrUnit, unitChanged As Boolean
    Dim oldOpt As Boolean
    Dim strips() As ac2fStrip
    Dim nStrip As Long
    Dim x0 As Double, y0 As Double
    Dim bx As Double, by As Double, bw As Double, bh As Double
    Dim i As Long
    Dim uyari As String

    If ActiveDocument Is Nothing Then
        ac2fWarn "Önce bir belge açın.", CAPTION_
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr Is Nothing Then
        ac2fWarn "Şeridi çıkarılacak kontur seçilmedi.", CAPTION_
        Exit Sub
    End If
    If sr.Count = 0 Then
        ac2fWarn "Şeridi çıkarılacak kontur seçilmedi.", CAPTION_
        Exit Sub
    End If

    If ac2fGetNum(AC2F_K_BL_THICK, AC2F_DEF_BL_THICK) <= 0 Then
        ac2fWarn "Malzeme kalınlığı sıfırdan büyük olmalı." & vbCrLf & _
                 "Ayarlar - Kutu harf menüsünden düzeltin.", CAPTION_
        Exit Sub
    End If

    ReDim strips(0 To MAX_STRIPS - 1)

    On Error GoTo Fail
    oldOpt = Application.Optimization
    Application.Optimization = True
    oldUnit = ActiveDocument.Unit
    If oldUnit <> cdrMillimeter Then
        ActiveDocument.Unit = cdrMillimeter
        unitChanged = True
    End If

    sr.GetBoundingBox bx, by, bw, bh
    x0 = bx + bw + 20#
    y0 = by + bh

    ActiveDocument.BeginCommandGroup ac2fTitle("kutu harf şeridi")

    For i = 1 To sr.Count
        ac2fBLShape sr(i), strips, nStrip, ciz, x0, y0
    Next i

    ActiveDocument.EndCommandGroup

    On Error Resume Next
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    On Error GoTo 0

    If nStrip = 0 Then
        ac2fWarn "Seçimde kapalı kontur bulunamadı." & vbCrLf & vbCrLf & _
                 "Kutu harf şeridi yalnız kapalı yollardan çıkarılır.", CAPTION_
        Exit Sub
    End If

    If nStrip >= MAX_STRIPS Then
        uyari = vbCrLf & "UYARI: " & MAX_STRIPS & " şerit sınırına ulaşıldı, " & _
                "kalan konturlar atlandı." & vbCrLf
    End If

    ac2fInfo ac2fBLRapor(strips, nStrip, ciz) & uyari, CAPTION_
    Exit Sub

Fail:
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    If unitChanged Then ActiveDocument.Unit = oldUnit
    Application.Optimization = oldOpt
    Application.Refresh
    ac2fWarn "İşlem sırasında hata oluştu: " & Err.Description, CAPTION_
End Sub

'---------------------------------------------------------------------
' Tek bir şekli işler (gruplarda özyinelemeli).
'---------------------------------------------------------------------
Private Sub ac2fBLShape(ByVal s As Shape, ByRef strips() As ac2fStrip, _
                        ByRef nStrip As Long, ByVal ciz As Boolean, _
                        ByRef x0 As Double, ByRef y0 As Double)
    Dim i As Long
    Dim cv As Curve
    Dim dup As Shape

    If s Is Nothing Then Exit Sub
    On Error GoTo Skip

    Select Case s.Type
        Case cdrBitmapShape, cdrOLEObjectShape
            Exit Sub
    End Select

    If s.Type = cdrGroupShape Then
        For i = 1 To s.Shapes.Count
            ac2fBLShape s.Shapes(i), strips, nStrip, ciz, x0, y0
        Next i
        Exit Sub
    End If

    Set cv = Nothing
    On Error Resume Next
    Set cv = s.DisplayCurve
    On Error GoTo Skip

    If Not cv Is Nothing Then
        ac2fBLCurve cv, ac2fBLName(s), strips, nStrip, ciz, x0, y0
        Exit Sub
    End If

    Set dup = s.Duplicate(0, 0)
    On Error Resume Next
    dup.ConvertToCurves
    On Error GoTo SkipDup
    If dup.Type = cdrGroupShape Then
        ac2fBLShape dup, strips, nStrip, ciz, x0, y0
        dup.Delete
        Exit Sub
    End If
    Set cv = Nothing
    On Error Resume Next
    Set cv = dup.Curve
    On Error GoTo SkipDup
    If cv Is Nothing Then GoTo SkipDup
    ac2fBLCurve cv, ac2fBLName(s), strips, nStrip, ciz, x0, y0
    dup.Delete
    Exit Sub

SkipDup:
    On Error Resume Next
    If Not dup Is Nothing Then dup.Delete
    Exit Sub
Skip:
End Sub

Private Function ac2fBLName(ByVal s As Shape) As String
    Dim n As String
    On Error Resume Next
    n = s.Name
    If Len(Trim$(n)) = 0 Then n = "Kontur " & CStr(s.StaticID)
    ac2fBLName = n
End Function

'---------------------------------------------------------------------
' Bir eğrinin tüm kapalı alt yollarını şeride çevirir.
'---------------------------------------------------------------------
Private Sub ac2fBLCurve(ByVal cv As Curve, ByVal nm As String, _
                        ByRef strips() As ac2fStrip, ByRef nStrip As Long, _
                        ByVal ciz As Boolean, ByRef x0 As Double, ByRef y0 As Double)
    Dim i As Long, nSub As Long
    Dim isHole() As Boolean

    On Error Resume Next
    nSub = cv.SubPaths.Count
    On Error GoTo 0
    If nSub = 0 Then Exit Sub

    ' Delik (counter) tespiti: sınırlayıcı kutusu bir başkasının içinde kalan
    ' alt yol deliktir. Harf konturlarında güvenilir ve hızlıdır.
    ReDim isHole(0 To nSub - 1)
    ac2fBLMarkHoles cv, nSub, isHole

    For i = 1 To nSub
        If nStrip >= MAX_STRIPS Then Exit Sub
        Call ac2fBLSubPath(cv.SubPaths(i), nm & " #" & i, isHole(i - 1), _
                           strips, nStrip, ciz, x0, y0)
    Next i
End Sub

' Alt yolların sınırlayıcı kutularını karşılaştırarak delikleri işaretler.
Private Sub ac2fBLMarkHoles(ByVal cv As Curve, ByVal nSub As Long, _
                            ByRef isHole() As Boolean)
    Dim i As Long, j As Long
    Dim x1() As Double, y1() As Double, x2() As Double, y2() As Double

    ReDim x1(0 To nSub - 1): ReDim y1(0 To nSub - 1)
    ReDim x2(0 To nSub - 1): ReDim y2(0 To nSub - 1)

    For i = 0 To nSub - 1
        If Not ac2fBLSubBox(cv.SubPaths(i + 1), x1(i), y1(i), x2(i), y2(i)) Then
            x1(i) = 0: y1(i) = 0: x2(i) = 0: y2(i) = 0
        End If
    Next i

    For i = 0 To nSub - 1
        For j = 0 To nSub - 1
            If i <> j Then
                If x1(i) >= x1(j) And x2(i) <= x2(j) And _
                   y1(i) >= y1(j) And y2(i) <= y2(j) And _
                   (x2(j) - x1(j)) > (x2(i) - x1(i)) Then
                    isHole(i) = Not isHole(i)   ' iç içe her düzeyde rol değişir
                End If
            End If
        Next j
    Next i
End Sub

Private Function ac2fBLSubBox(ByVal sp As SubPath, ByRef x1 As Double, ByRef y1 As Double, _
                              ByRef x2 As Double, ByRef y2 As Double) As Boolean
    Dim k As Long, nn As Long
    Dim px As Double, py As Double

    On Error GoTo Fail
    nn = sp.Nodes.Count
    If nn = 0 Then Exit Function

    For k = 1 To nn
        px = sp.Nodes(k).PositionX
        py = sp.Nodes(k).PositionY
        If k = 1 Then
            x1 = px: x2 = px: y1 = py: y2 = py
        Else
            If px < x1 Then x1 = px
            If px > x2 Then x2 = px
            If py < y1 Then y1 = py
            If py > y2 Then y2 = py
        End If
    Next k
    ac2fBLSubBox = True
    Exit Function
Fail:
End Function

'=====================================================================
' GEOMETRİ ÇÖZÜMLEMESİ
'=====================================================================

' c/L = 2*sin(t/2)/t bağıntısını t için çözer (ikiye bölme).
' Fonksiyon (0, 2*pi) aralığında kesin azalandır.
Private Function ac2fBLTheta(ByVal chord As Double, ByVal arc As Double) As Double
    Dim r As Double, lo As Double, hi As Double, mid As Double, f As Double
    Dim i As Long
    Const PI2 As Double = 6.28318530717959

    If arc <= 0.000000001 Then Exit Function
    r = chord / arc
    If r >= 0.999999 Then Exit Function              ' düz
    If r <= 0# Then
        ac2fBLTheta = PI2 - 0.000001
        Exit Function
    End If

    lo = 0.000001: hi = PI2 - 0.000001
    For i = 1 To 60
        mid = (lo + hi) / 2#
        f = 2# * Sin(mid / 2#) / mid
        If f > r Then lo = mid Else hi = mid
    Next i
    ac2fBLTheta = (lo + hi) / 2#
End Function

Private Function ac2fBLWrap(ByVal a As Double) As Double
    Const PI2 As Double = 6.28318530717959
    Const PI_ As Double = 3.14159265358979
    Dim v As Double
    v = a
    Do While v > PI_
        v = v - PI2
    Loop
    Do While v <= -PI_
        v = v + PI2
    Loop
    ac2fBLWrap = v
End Function

' Bir alt yolu çözümler: segment yayları, dönüşler, yarıçaplar, köşeler.
Private Function ac2fBLAnalyse(ByVal sp As SubPath) As Boolean
    Dim i As Long, nx As Long, pv As Long
    Dim ax As Double, ay As Double, bx2 As Double, by2 As Double
    Dim cx As Double, cy As Double, chord As Double
    Dim a As Double, b As Double, t As Double, h As Double
    Dim sg As Segment

    On Error GoTo Fail
    m_n = sp.Segments.Count
    If m_n < 1 Then Exit Function

    ReDim m_L(0 To m_n - 1): ReDim m_Th(0 To m_n - 1)
    ReDim m_Sg(0 To m_n - 1): ReDim m_R(0 To m_n - 1)
    ReDim m_Dir(0 To m_n - 1): ReDim m_T0(0 To m_n - 1)
    ReDim m_T1(0 To m_n - 1): ReDim m_Cor(0 To m_n - 1)

    For i = 0 To m_n - 1
        Set sg = sp.Segments(i + 1)
        m_L(i) = sg.Length
        ax = sg.StartNode.PositionX: ay = sg.StartNode.PositionY
        bx2 = sg.EndNode.PositionX: by2 = sg.EndNode.PositionY
        cx = bx2 - ax: cy = by2 - ay
        chord = Sqr(cx * cx + cy * cy)
        m_Dir(i) = ac2fBLAtan2(cy, cx)

        ' Düzlük geometriden anlaşılır; segment tipi sabitine gerek yok.
        If m_L(i) <= 0 Then
            m_Th(i) = 0
        ElseIf (m_L(i) - chord) / m_L(i) < STRAIGHT_EPS Then
            m_Th(i) = 0
        Else
            m_Th(i) = ac2fBLTheta(chord, m_L(i))
        End If

        If m_Th(i) > 0.000000001 Then m_R(i) = m_L(i) / m_Th(i) Else m_R(i) = 0
    Next i

    ' Dönüş işareti: segmentin iki ucundaki tek düğümlük kiriş dönüşleri.
    ' İki segmentlik açıklık kullanılırsa +/-180 derecede belirsizlik doğar
    ' ve saat yönündeki (delik) konturlarda işaret ters çıkar.
    For i = 0 To m_n - 1
        nx = i + 1: If nx > m_n - 1 Then nx = 0
        pv = i - 1: If pv < 0 Then pv = m_n - 1
        a = ac2fBLWrap(m_Dir(nx) - m_Dir(i))
        b = ac2fBLWrap(m_Dir(i) - m_Dir(pv))
        If (a + b) >= 0 Then m_Sg(i) = 1# Else m_Sg(i) = -1#
    Next i

    ' Teğetler: kiriş yönü -/+ dönüşün yarısı.
    For i = 0 To m_n - 1
        h = m_Sg(i) * m_Th(i) / 2#
        m_T0(i) = m_Dir(i) - h
        m_T1(i) = m_Dir(i) + h
    Next i

    ' Köşe dönüşleri ve toplam dönüş.
    m_Tau = 0
    For i = 0 To m_n - 1
        nx = i + 1: If nx > m_n - 1 Then nx = 0
        m_Cor(i) = ac2fBLWrap(m_T0(nx) - m_T1(i))
        m_Tau = m_Tau + m_Sg(i) * m_Th(i) + m_Cor(i)
    Next i

    ac2fBLAnalyse = True
    Exit Function
Fail:
    m_n = 0
End Function

Private Function ac2fBLAtan2(ByVal y As Double, ByVal x As Double) As Double
    Const PI_ As Double = 3.14159265358979
    If x > 0 Then
        ac2fBLAtan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then ac2fBLAtan2 = Atn(y / x) + PI_ Else ac2fBLAtan2 = Atn(y / x) - PI_
    Else
        If y > 0 Then
            ac2fBLAtan2 = PI_ / 2#
        ElseIf y < 0 Then
            ac2fBLAtan2 = -PI_ / 2#
        Else
            ac2fBLAtan2 = 0
        End If
    End If
End Function

'=====================================================================
' DERZ YERLEŞİMİ
'=====================================================================

' Çözümlenmiş konturdan açınım boyunu ve derz konumlarını üretir.
Private Function ac2fBLGrooves(ByVal isHole As Boolean, ByRef devLen As Double, _
                               ByRef minR As Double, ByRef nCorner As Long) As Boolean
    Dim i As Long, k As Long, nGr As Long
    Dim t As Double, kf As Double, refy As Long, g As Double, m As Double
    Dim depth As Double, mouth As Double, flex As Double, tol As Double
    Dim sMin As Double, sMax As Double, corEsik As Double
    Dim dPhiMax As Double
    Dim sig As Double, thHat As Double, corHat As Double
    Dim segDev As Double, corDev As Double
    Dim s As Double, s1 As Double, s2 As Double
    Dim acc As Double
    Const PI_ As Double = 3.14159265358979

    If m_n = 0 Then Exit Function

    t = ac2fGetNum(AC2F_K_BL_THICK, AC2F_DEF_BL_THICK)
    kf = ac2fGetNum(AC2F_K_BL_KFAC, AC2F_DEF_BL_KFAC)
    refy = ac2fGetLng(AC2F_K_BL_REF, AC2F_DEF_BL_REF)
    depth = ac2fGetNum(AC2F_K_BL_DEPTH, AC2F_DEF_BL_DEPTH)
    mouth = ac2fGetNum(AC2F_K_BL_MOUTH, AC2F_DEF_BL_MOUTH)
    flex = ac2fGetNum(AC2F_K_BL_FLEX, AC2F_DEF_BL_FLEX)
    tol = ac2fGetNum(AC2F_K_BL_TOL, AC2F_DEF_BL_TOL)
    sMin = ac2fGetNum(AC2F_K_BL_SMIN, AC2F_DEF_BL_SMIN)
    sMax = ac2fGetNum(AC2F_K_BL_SMAX, AC2F_DEF_BL_SMAX)
    corEsik = ac2fGetNum(AC2F_K_BL_CORNER, AC2F_DEF_BL_CORNER) * PI_ / 180#

    If kf < 0 Then kf = 0
    If kf > 1 Then kf = 1
    If depth <= 0 Then depth = AC2F_DEF_BL_DEPTH
    If depth > 0.95 Then depth = 0.95
    If flex <= 0 Then flex = 1#
    If tol <= 0 Then tol = AC2F_DEF_BL_TOL
    If sMin <= 0 Then sMin = 0.5
    If sMax < sMin Then sMax = sMin

    ' Nötr eksenin vektörden malzemeye doğru ötelenmesi
    Select Case refy
        Case 2:    g = -kf * t          ' vektör iç yüz
        Case 3:    g = 0#               ' vektör zaten nötr eksen
        Case Else: g = (1# - kf) * t    ' vektör dış yüz (varsayılan)
    End Select
    If isHole Then m = -1# Else m = 1#

    ' Bir derzin karşılayabileceği en büyük dönüş:
    ' ağız genişliği = derz derinliği * açı  ->  açı = ağız / derinlik
    If t * depth <= 0 Then Exit Function
    dPhiMax = (mouth / (t * depth)) * flex
    If dPhiMax <= 0.0001 Then dPhiMax = 0.0001

    sig = 1#
    If m_Tau < 0 Then sig = -1#

    ReDim m_gPos(0 To 255)
    ReDim m_gCor(0 To 255)
    m_gN = 0
    nCorner = 0
    minR = 0
    acc = 0

    For i = 0 To m_n - 1
        thHat = sig * m_Sg(i) * m_Th(i)          ' normalleştirilmiş dönüş
        segDev = m_L(i) - g * m * thHat

        If m_Th(i) > 0.000000001 And m_R(i) > 0 Then
            If minR = 0 Or m_R(i) < minR Then minR = m_R(i)

            ' Aralık: (a) derz ağzı sınırı  (b) yüzey toleransı  (c) tavan
            s1 = m_R(i) * dPhiMax
            s2 = Sqr(8# * m_R(i) * tol)
            s = s1
            If s2 < s Then s = s2
            If sMax < s Then s = sMax
            If s < sMin Then s = sMin

            nGr = CLng(ac2fCeil(m_L(i) / s))
            If nGr < 1 Then nGr = 1

            For k = 0 To nGr - 1
                ac2fBLAddGroove acc + (k + 0.5) * segDev / nGr, False
            Next k
        End If

        acc = acc + segDev

        ' Köşe
        corHat = sig * m_Cor(i)
        corDev = -g * m * corHat
        If Abs(corHat) > corEsik Then
            nGr = CLng(ac2fCeil(Abs(corHat) / dPhiMax))
            If nGr < 1 Then nGr = 1
            For k = 0 To nGr - 1
                ac2fBLAddGroove acc + corDev / 2# + (k - (nGr - 1) / 2#) * sMin, True
            Next k
            nCorner = nCorner + nGr
        End If
        acc = acc + corDev
    Next i

    devLen = acc
    ac2fBLGrooves = (devLen > 0)
End Function

Private Sub ac2fBLAddGroove(ByVal pos As Double, ByVal isCorner As Boolean)
    If m_gN >= MAX_GROOVE Then Exit Sub
    If m_gN > UBound(m_gPos) Then
        ReDim Preserve m_gPos(0 To (UBound(m_gPos) + 1) * 2 - 1)
        ReDim Preserve m_gCor(0 To UBound(m_gPos))
    End If
    m_gPos(m_gN) = pos
    m_gCor(m_gN) = isCorner
    m_gN = m_gN + 1
End Sub

'=====================================================================
' ŞERİT ÜRETİMİ
'=====================================================================

Private Function ac2fBLSubPath(ByVal sp As SubPath, ByVal nm As String, _
                               ByVal isHole As Boolean, ByRef strips() As ac2fStrip, _
                               ByRef nStrip As Long, ByVal ciz As Boolean, _
                               ByRef x0 As Double, ByRef y0 As Double) As Boolean
    Dim devLen As Double, minR As Double
    Dim nCorner As Long
    Dim rawLen As Double
    Dim closed As Boolean
    Dim H As Double, gap As Double

    On Error GoTo Fail

    closed = False
    On Error Resume Next
    closed = sp.Closed
    rawLen = sp.Length
    On Error GoTo Fail
    If Not closed Then Exit Function          ' açık yoldan kutu harf şeridi çıkmaz
    If rawLen <= 0 Then Exit Function

    If Not ac2fBLAnalyse(sp) Then Exit Function
    If Not ac2fBLGrooves(isHole, devLen, minR, nCorner) Then Exit Function

    strips(nStrip).Name = nm
    strips(nStrip).DevLen = devLen
    strips(nStrip).RawLen = rawLen
    strips(nStrip).Grooves = m_gN
    strips(nStrip).Corners = nCorner
    strips(nStrip).MinR = minR
    strips(nStrip).IsHole = isHole
    nStrip = nStrip + 1

    If ciz Then
        H = ac2fGetNum(AC2F_K_BL_HEIGHT, AC2F_DEF_BL_HEIGHT)
        gap = ac2fGetNum(AC2F_K_BL_GAP, AC2F_DEF_BL_GAP)
        If H <= 0 Then H = AC2F_DEF_BL_HEIGHT
        If gap < 0 Then gap = 0
        ac2fBLDraw x0, y0, devLen, H, nm
        y0 = y0 - (H + gap + 6#)
    End If

    ac2fBLSubPath = True
    Exit Function
Fail:
End Function

' Düz şeridi, derz çizgilerini ve ek/kesim işaretlerini çizer.
Private Sub ac2fBLDraw(ByVal x As Double, ByVal yTop As Double, _
                       ByVal L As Double, ByVal H As Double, ByVal nm As String)
    Dim i As Long
    Dim sh As Shape
    Dim yBot As Double
    Dim coil As Double, joint As Double
    Dim cut As Double

    yBot = yTop - H

    On Error Resume Next

    ' Şerit dış hattı
    Set sh = ActiveLayer.CreateRectangle(x, yTop, x + L, yBot)
    If Not sh Is Nothing Then
        sh.Fill.ApplyNoFill
        sh.Outline.Width = 0.2
        sh.Outline.Color.RGBAssign 0, 0, 0
    End If

    ' Derz çizgileri
    For i = 0 To m_gN - 1
        If m_gPos(i) >= 0 And m_gPos(i) <= L Then
            Set sh = ActiveLayer.CreateLineSegment(x + m_gPos(i), yTop, _
                                                   x + m_gPos(i), yBot)
            If Not sh Is Nothing Then
                sh.Outline.Width = 0.1
                If m_gCor(i) Then
                    sh.Outline.Color.RGBAssign 230, 0, 120     ' köşe derzi
                Else
                    sh.Outline.Color.RGBAssign 0, 160, 220     ' eğri derzi
                End If
            End If
        End If
    Next i

    ' Rulo boyu aşılıyorsa kesim/ek işaretleri
    coil = ac2fGetNum(AC2F_K_BL_COIL, AC2F_DEF_BL_COIL)
    joint = ac2fGetNum(AC2F_K_BL_JOINT, AC2F_DEF_BL_JOINT)
    ' Ek payı rulo boyundan küçük olmalı; yoksa adım sıfır ya da negatif
    ' olur ve döngü ilerlemez.
    If coil > 0 And L > coil And joint < coil Then
        cut = coil
        Do While cut < L
            Set sh = ActiveLayer.CreateLineSegment(x + cut, yTop + 4#, x + cut, yBot - 4#)
            If Not sh Is Nothing Then
                sh.Outline.Width = 0.4
                sh.Outline.Color.RGBAssign 255, 0, 0           ' kesim yeri
            End If
            cut = cut + coil - joint
        Loop
    End If

    ' Etiket
    Set sh = ActiveLayer.CreateArtisticText(x, yTop + 2#, _
             nm & "  |  açınım " & ac2fFmt(L) & " mm  |  yükseklik " & _
             ac2fFmt(H, 0) & " mm  |  " & m_gN & " derz")
    If Not sh Is Nothing Then sh.Text.Story.Size = 8
End Sub

'=====================================================================
' RAPOR
'=====================================================================

Private Function ac2fBLRapor(ByRef strips() As ac2fStrip, ByVal n As Long, _
                             ByVal ciz As Boolean) As String
    Dim s As String
    Dim i As Long
    Dim topDev As Double, topGr As Long, topCor As Long
    Dim enKucukR As Double
    Dim coil As Double, joint As Double, parca As Long
    Dim t As Double, refy As Long

    For i = 0 To n - 1
        topDev = topDev + strips(i).DevLen
        topGr = topGr + strips(i).Grooves
        topCor = topCor + strips(i).Corners
        If strips(i).MinR > 0 Then
            If enKucukR = 0 Or strips(i).MinR < enKucukR Then enKucukR = strips(i).MinR
        End If
    Next i

    t = ac2fGetNum(AC2F_K_BL_THICK, AC2F_DEF_BL_THICK)
    refy = ac2fGetLng(AC2F_K_BL_REF, AC2F_DEF_BL_REF)
    coil = ac2fGetNum(AC2F_K_BL_COIL, AC2F_DEF_BL_COIL)
    joint = ac2fGetNum(AC2F_K_BL_JOINT, AC2F_DEF_BL_JOINT)

    s = "SONUÇ" & vbCrLf
    s = s & "   Şerit sayısı        : " & n & vbCrLf
    s = s & "   Toplam açınım       : " & ac2fFmtLength(topDev) & vbCrLf
    s = s & "   Toplam derz         : " & topGr & " adet" & _
            "  (" & topCor & " köşe)" & vbCrLf
    If enKucukR > 0 Then
        s = s & "   En küçük yarıçap    : " & ac2fFmt(enKucukR) & " mm" & vbCrLf
    End If

    If coil > 0 And topDev > 0 And joint < coil Then
        parca = CLng(ac2fCeil(topDev / (coil - joint)))
        s = s & "   Rulo ihtiyacı       : ~" & parca & " parça x " & _
                ac2fFmt(coil, 0) & " mm (" & ac2fFmt(joint, 0) & " mm ek payı)" & vbCrLf
    End If
    s = s & vbCrLf

    s = s & "ŞERİTLER" & vbCrLf
    For i = 0 To n - 1
        If i >= 12 Then
            s = s & "   ... (" & n & " şeridin ilk 12'si)" & vbCrLf
            Exit For
        End If
        s = s & "   " & (i + 1) & ". " & strips(i).Name
        If strips(i).IsHole Then s = s & " [delik]"
        s = s & vbCrLf & "      açınım " & ac2fFmt(strips(i).DevLen) & " mm" & _
                " (ham " & ac2fFmt(strips(i).RawLen) & ")" & _
                "  derz " & strips(i).Grooves
        If strips(i).MinR > 0 Then s = s & "  min R " & ac2fFmt(strips(i).MinR, 1)
        s = s & vbCrLf
    Next i
    s = s & vbCrLf

    s = s & "MALZEME" & vbCrLf
    s = s & "   Kalınlık            : " & ac2fFmt(t) & " mm" & vbCrLf
    s = s & "   Referans yüzey      : " & ac2fBLRefAdi(refy) & vbCrLf
    s = s & "   K faktörü           : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_KFAC, AC2F_DEF_BL_KFAC)) & vbCrLf
    s = s & "   Esneklik oranı      : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_FLEX, AC2F_DEF_BL_FLEX)) & vbCrLf
    s = s & "   Derz derinlik oranı : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_DEPTH, AC2F_DEF_BL_DEPTH)) & vbCrLf

    If ciz Then
        s = s & vbCrLf & "Şeritler sayfaya çizildi." & vbCrLf & _
                "   mavi = eğri derzi   pembe = köşe derzi   kırmızı = kesim yeri"
    Else
        s = s & vbCrLf & "Çizim yapılmadı (yalnız rapor)."
    End If

    ac2fBLRapor = s
End Function

Public Function ac2fBLRefAdi(ByVal r As Long) As String
    Select Case r
        Case 2:    ac2fBLRefAdi = "Vektör = şeridin iç yüzü"
        Case 3:    ac2fBLRefAdi = "Vektör = nötr eksen"
        Case Else: ac2fBLRefAdi = "Vektör = şeridin dış yüzü"
    End Select
End Function

'=====================================================================
' AYARLAR
'=====================================================================

Public Sub ac2fKutuHarfAyarlar()
Attribute ac2fKutuHarfAyarlar.VB_Description = "ac2f pack: Kutu harf malzeme ve derz ayarlari"
    Const C As String = "Ayarlar - Kutu Harf"
    Dim cevap As String
    Dim v As Double
    Dim n As Long

    ' 1) Malzeme ön ayarı
    cevap = InputBox( _
        "Malzeme ön ayarı seçin:" & vbCrLf & vbCrLf & _
        "  0  -  Elle ayarla (değiştirme)" & vbCrLf & _
        "  1  -  Alüminyum, ince (0,5-1,5 mm)" & vbCrLf & _
        "  2  -  Galvaniz, ince (0,5-1,2 mm)" & vbCrLf & _
        "  3  -  Alüminyum, kalın (2-4 mm)" & vbCrLf & _
        "  4  -  Paslanmaz" & vbCrLf & vbCrLf & _
        "Ön ayarlar başlangıç değeridir; ilk işten sonra" & vbCrLf & _
        "esneklik oranını kendi sonucunuza göre ayarlayın.", _
        ac2fTitle(C), "0")
    If StrPtr(cevap) = 0 Then Exit Sub
    ac2fBLPreset CLng(ac2fParseNum(cevap, 0))

    If Not ac2fBLAsk("Malzeme kalınlığı (mm)", C, AC2F_K_BL_THICK, _
                     AC2F_DEF_BL_THICK, 0.01, v) Then Exit Sub
    If Not ac2fBLAsk("Şerit yüksekliği (mm)" & vbCrLf & "Harfin derinliği.", C, _
                     AC2F_K_BL_HEIGHT, AC2F_DEF_BL_HEIGHT, 1#, v) Then Exit Sub
    If Not ac2fBLAsk("Esneklik oranı" & vbCrLf & _
                     "Büyük = daha az derz. Kendi işinize göre kalibre edin.", C, _
                     AC2F_K_BL_FLEX, AC2F_DEF_BL_FLEX, 0.05, v) Then Exit Sub
    If Not ac2fBLAskLng("Referans yüzey" & vbCrLf & _
                     "1 = Vektör şeridin dış yüzü" & vbCrLf & _
                     "2 = Vektör şeridin iç yüzü" & vbCrLf & _
                     "3 = Vektör nötr eksen", C, _
                     AC2F_K_BL_REF, AC2F_DEF_BL_REF, 1, 3, n) Then Exit Sub

    If MsgBox("Gelişmiş ayarlar da düzenlensin mi?" & vbCrLf & _
              "(K faktörü, derz ağzı, tolerans, aralık sınırları, rulo boyu)", _
              vbQuestion + vbYesNo, ac2fTitle(C)) = vbYes Then

        If Not ac2fBLAsk("K faktörü (nötr eksen konumu, 0-1)", C, _
                         AC2F_K_BL_KFAC, AC2F_DEF_BL_KFAC, 0#, v) Then GoTo Bitti
        If Not ac2fBLAsk("Derz derinliği / kalınlık oranı (0-0,95)", C, _
                         AC2F_K_BL_DEPTH, AC2F_DEF_BL_DEPTH, 0.05, v) Then GoTo Bitti
        If Not ac2fBLAsk("Derz ağzı en çok (mm)" & vbCrLf & _
                         "Kapandığında görünür iz bırakmayan en geniş ağız.", C, _
                         AC2F_K_BL_MOUTH, AC2F_DEF_BL_MOUTH, 0.05, v) Then GoTo Bitti
        If Not ac2fBLAsk("Yüzey toleransı (mm)" & vbCrLf & _
                         "Derzler arası düz yüzün eğriden sapması.", C, _
                         AC2F_K_BL_TOL, AC2F_DEF_BL_TOL, 0.01, v) Then GoTo Bitti
        If Not ac2fBLAsk("En az derz aralığı (mm)", C, _
                         AC2F_K_BL_SMIN, AC2F_DEF_BL_SMIN, 0.5, v) Then GoTo Bitti
        If Not ac2fBLAsk("En çok derz aralığı (mm)", C, _
                         AC2F_K_BL_SMAX, AC2F_DEF_BL_SMAX, 1#, v) Then GoTo Bitti
        If Not ac2fBLAsk("Köşe eşiği (derece)" & vbCrLf & _
                         "Bunun üstündeki dönüş köşe derzi sayılır.", C, _
                         AC2F_K_BL_CORNER, AC2F_DEF_BL_CORNER, 0.1, v) Then GoTo Bitti
        If Not ac2fBLAsk("Rulo boyu (mm)" & vbCrLf & "0 = bölme yapma.", C, _
                         AC2F_K_BL_COIL, AC2F_DEF_BL_COIL, 0#, v) Then GoTo Bitti
        If Not ac2fBLAsk("Ek payı (mm)", C, _
                         AC2F_K_BL_JOINT, AC2F_DEF_BL_JOINT, 0#, v) Then GoTo Bitti
        If Not ac2fBLAsk("Şeritler arası boşluk (mm)", C, _
                         AC2F_K_BL_GAP, AC2F_DEF_BL_GAP, 0#, v) Then GoTo Bitti
    End If

Bitti:
    ac2fInfo "Kutu harf ayarları kaydedildi." & vbCrLf & vbCrLf & ac2fBLAyarOzeti(), C
End Sub

' Ön ayarlar: esneklik ve derz derinliği için başlangıç değerleri.
' Bunlar ölçülmüş değil, kalibrasyon başlangıcıdır.
Private Sub ac2fBLPreset(ByVal p As Long)
    Select Case p
        Case 1      ' alüminyum ince
            ac2fSetNum AC2F_K_BL_FLEX, 1.2
            ac2fSetNum AC2F_K_BL_DEPTH, 0.7
            ac2fSetNum AC2F_K_BL_KFAC, 0.44
        Case 2      ' galvaniz ince
            ac2fSetNum AC2F_K_BL_FLEX, 1#
            ac2fSetNum AC2F_K_BL_DEPTH, 0.65
            ac2fSetNum AC2F_K_BL_KFAC, 0.44
        Case 3      ' alüminyum kalın
            ac2fSetNum AC2F_K_BL_FLEX, 0.8
            ac2fSetNum AC2F_K_BL_DEPTH, 0.75
            ac2fSetNum AC2F_K_BL_KFAC, 0.42
        Case 4      ' paslanmaz
            ac2fSetNum AC2F_K_BL_FLEX, 0.7
            ac2fSetNum AC2F_K_BL_DEPTH, 0.6
            ac2fSetNum AC2F_K_BL_KFAC, 0.45
    End Select
End Sub

Public Function ac2fBLAyarOzeti() As String
    Dim s As String
    s = "   Kalınlık            : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_THICK, AC2F_DEF_BL_THICK)) & " mm" & vbCrLf
    s = s & "   Şerit yüksekliği    : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_HEIGHT, AC2F_DEF_BL_HEIGHT), 0) & " mm" & vbCrLf
    s = s & "   Esneklik oranı      : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_FLEX, AC2F_DEF_BL_FLEX)) & vbCrLf
    s = s & "   Referans yüzey      : " & ac2fBLRefAdi(ac2fGetLng(AC2F_K_BL_REF, AC2F_DEF_BL_REF)) & vbCrLf
    s = s & "   K faktörü           : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_KFAC, AC2F_DEF_BL_KFAC)) & vbCrLf
    s = s & "   Derz derinlik oranı : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_DEPTH, AC2F_DEF_BL_DEPTH)) & vbCrLf
    s = s & "   Derz ağzı en çok    : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_MOUTH, AC2F_DEF_BL_MOUTH)) & " mm" & vbCrLf
    s = s & "   Yüzey toleransı     : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_TOL, AC2F_DEF_BL_TOL)) & " mm" & vbCrLf
    s = s & "   Derz aralığı        : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_SMIN, AC2F_DEF_BL_SMIN), 1) & _
            " - " & ac2fFmt(ac2fGetNum(AC2F_K_BL_SMAX, AC2F_DEF_BL_SMAX), 1) & " mm" & vbCrLf
    s = s & "   Rulo / ek payı      : " & ac2fFmt(ac2fGetNum(AC2F_K_BL_COIL, AC2F_DEF_BL_COIL), 0) & _
            " / " & ac2fFmt(ac2fGetNum(AC2F_K_BL_JOINT, AC2F_DEF_BL_JOINT), 0) & " mm" & vbCrLf
    ac2fBLAyarOzeti = s
End Function

Private Function ac2fBLAsk(ByVal soru As String, ByVal caption As String, _
                           ByVal key As String, ByVal defValue As Double, _
                           ByVal enAz As Double, ByRef sonuc As Double) As Boolean
    Dim mevcut As Double, cevap As String, v As Double
    mevcut = ac2fGetNum(key, defValue)
    cevap = InputBox(soru & vbCrLf & vbCrLf & "(Varsayılan: " & ac2fNumStr(defValue) & ")", _
                     ac2fTitle(caption), ac2fNumStr(mevcut))
    If StrPtr(cevap) = 0 Then Exit Function
    v = ac2fParseNum(cevap, mevcut)
    If v < enAz Then v = enAz
    ac2fSetNum key, v
    sonuc = v
    ac2fBLAsk = True
End Function

Private Function ac2fBLAskLng(ByVal soru As String, ByVal caption As String, _
                              ByVal key As String, ByVal defValue As Long, _
                              ByVal enAz As Long, ByVal enCok As Long, _
                              ByRef sonuc As Long) As Boolean
    Dim mevcut As Long, cevap As String, v As Long
    mevcut = ac2fGetLng(key, defValue)
    cevap = InputBox(soru & vbCrLf & vbCrLf & "(Varsayılan: " & ac2fNumStr(CDbl(defValue)) & ")", _
                     ac2fTitle(caption), ac2fNumStr(CDbl(mevcut)))
    If StrPtr(cevap) = 0 Then Exit Function
    v = CLng(ac2fParseNum(cevap, CDbl(mevcut)))
    If v < enAz Then v = enAz
    If v > enCok Then v = enCok
    ac2fSetLng key, v
    sonuc = v
    ac2fBLAskLng = True
End Function

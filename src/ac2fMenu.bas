Attribute VB_Name = "ac2fMenu"
'=====================================================================
'  ac2f pack  --  ac2fMenu
'
'  Paketin giriş noktası, ayarlar ve hakkında ekranı.
'
'  Makrolar:
'    ac2fPack      - ana menü
'    ac2fAyarlar   - ayar düzenleme
'    ac2fHakkinda  - sürüm bilgisi
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "Ana Menü"

'---------------------------------------------------------------------
' Paketin ana menüsü.
'---------------------------------------------------------------------
Public Sub ac2fPack()
Attribute ac2fPack.VB_Description = "ac2f pack: Ana menu"
    Dim m As String
    Dim secim As String

    m = "Yapmak istediğiniz işlemin numarasını girin:" & vbCrLf & vbCrLf & _
        "  1  -  Uzunluk ölçümü" & vbCrLf & _
        "  2  -  Uzunluk ölçümü + sayfaya etiket" & vbCrLf & _
        "  3  -  3'lü LED modül hesabı" & vbCrLf & _
        "  4  -  3'lü LED modül hesabı (aralığı sorarak)" & vbCrLf & _
        "  5  -  Kutu harf şeridi (hesapla ve çiz)" & vbCrLf & _
        "  6  -  Kutu harf raporu (çizim yok)" & vbCrLf & _
        "  7  -  Ayarlar - LED modül" & vbCrLf & _
        "  8  -  Ayarlar - Kutu harf" & vbCrLf & _
        "  9  -  Hakkında" & vbCrLf

    secim = InputBox(m, ac2fTitle(CAPTION_), "1")
    If StrPtr(secim) = 0 Then Exit Sub          ' İptal
    secim = Trim$(secim)
    If Len(secim) = 0 Then Exit Sub

    Select Case secim
        Case "1": ac2fUzunlukOlc
        Case "2": ac2fUzunlukEtiketle
        Case "3": ac2fLedModulHesapla
        Case "4": ac2fLedHizliHesap
        Case "5": ac2fKutuHarfSerit
        Case "6": ac2fKutuHarfRapor
        Case "7": ac2fAyarlar
        Case "8": ac2fKutuHarfAyarlar
        Case "9": ac2fHakkinda
        Case Else
            ac2fWarn "Geçersiz seçim: " & secim, CAPTION_
    End Select
End Sub

'---------------------------------------------------------------------
' Ayarları sırayla sorar ve kaydeder.
'---------------------------------------------------------------------
Public Sub ac2fAyarlar()
Attribute ac2fAyarlar.VB_Description = "ac2f pack: Ayarlari duzenler"
    Const C As String = "Ayarlar"
    Dim v As Double
    Dim n As Long

    If Not ac2fSorNum("Modül aralığı (mm)" & vbCrLf & _
                      "İki LED modülü arasındaki mesafe.", C, _
                      AC2F_K_SPACING, AC2F_DEF_SPACING, 0.001, v) Then Exit Sub

    If Not ac2fSorLng("Modül başına LED adedi" & vbCrLf & _
                      "3'lü modüller için 3.", C, _
                      AC2F_K_LEDS, AC2F_DEF_LEDS, 1, n) Then Exit Sub

    If Not ac2fSorNum("Modül gücü (W)" & vbCrLf & _
                      "Tek bir modülün çektiği güç.", C, _
                      AC2F_K_MODULE_W, AC2F_DEF_MODULE_W, 0#, v) Then Exit Sub

    If Not ac2fSorNum("Güç kaynağı kapasitesi (W)" & vbCrLf & _
                      "0 girilirse güç kaynağı adedi hesaplanmaz.", C, _
                      AC2F_K_PSU_W, AC2F_DEF_PSU_W, 0#, v) Then Exit Sub

    If Not ac2fSorNum("Güvenlik payı (%)" & vbCrLf & _
                      "Güç kaynağı seçiminde eklenecek pay.", C, _
                      AC2F_K_SAFETY, AC2F_DEF_SAFETY, 0#, v) Then Exit Sub

    If Not ac2fSorLng("Kontur başına en az modül" & vbCrLf & _
                      "Kısa parçaların boş kalmaması için.", C, _
                      AC2F_K_MINPERPATH, AC2F_DEF_MINPERPATH, 0, n) Then Exit Sub

    If Not ac2fSorLng("Hesap yöntemi" & vbCrLf & _
                      "1 = Çevre bazlı (tam kontur)" & vbCrLf & _
                      "2 = Orta hat tahmini (kontur/2, içi boş harfler için)", C, _
                      AC2F_K_METHOD, AC2F_DEF_METHOD, 1, n, 2) Then Exit Sub

    If Not ac2fSorNum("Düzeltme katsayısı" & vbCrLf & _
                      "Sonucu ölçeklemek için. 1 = değişiklik yok.", C, _
                      AC2F_K_FACTOR, AC2F_DEF_FACTOR, 0.01, v) Then Exit Sub

    ac2fInfo "Ayarlar kaydedildi." & vbCrLf & vbCrLf & ac2fAyarOzeti(), C
End Sub

'---------------------------------------------------------------------
' Sürüm ve içerik bilgisi.
'---------------------------------------------------------------------
Public Sub ac2fHakkinda()
Attribute ac2fHakkinda.VB_Description = "ac2f pack: Surum ve icerik bilgisi"
    Dim s As String

    s = AC2F_NAME & "  " & AC2F_VERSION & vbCrLf & vbCrLf
    s = s & "CorelDRAW için eklenti paketi." & vbCrLf & vbCrLf
    s = s & "İÇERİK" & vbCrLf
    s = s & "   • Uzunluk ölçümü" & vbCrLf
    s = s & "     Seçili vektörlerin çevresindeki çizgilerin" & vbCrLf
    s = s & "     toplam uzunluğunu ölçer." & vbCrLf & vbCrLf
    s = s & "   • 3'lü LED modül hesabı" & vbCrLf
    s = s & "     Ölçülen uzunluğa göre gereken modül, LED," & vbCrLf
    s = s & "     güç ve güç kaynağı adedini çıkarır." & vbCrLf & vbCrLf
    s = s & "   • Kutu harf şeridi" & vbCrLf
    s = s & "     Yan bordürün açınımını ve derz yerleşimini" & vbCrLf
    s = s & "     çıkarıp kesime hazır düz şerit çizer." & vbCrLf & vbCrLf
    s = s & "LED AYARLARI" & vbCrLf & ac2fAyarOzeti() & vbCrLf
    s = s & "KUTU HARF AYARLARI" & vbCrLf & ac2fBLAyarOzeti()

    ac2fInfo s, "Hakkında"
End Sub

'---------------------------------------------------------------------
' Tüm ayarları fabrika değerlerine döndürür.
'---------------------------------------------------------------------
Public Sub ac2fAyarlariSifirla()
Attribute ac2fAyarlariSifirla.VB_Description = "ac2f pack: Ayarlari varsayilana dondurur"
    If MsgBox("Tüm ac2f pack ayarları varsayılan değerlere dönecek." & vbCrLf & _
              "Devam edilsin mi?", vbQuestion + vbYesNo, _
              ac2fTitle("Ayarları Sıfırla")) <> vbYes Then Exit Sub

    ac2fResetSettings
    ac2fInfo "Ayarlar sıfırlandı." & vbCrLf & vbCrLf & ac2fAyarOzeti(), "Ayarları Sıfırla"
End Sub

'=====================================================================
' İç fonksiyonlar
'=====================================================================

Public Function ac2fAyarOzeti() As String
    Dim s As String

    s = "   Modül aralığı        : " & ac2fFmt(ac2fGetNum(AC2F_K_SPACING, AC2F_DEF_SPACING)) & " mm" & vbCrLf
    s = s & "   Modül başına LED     : " & ac2fGetLng(AC2F_K_LEDS, AC2F_DEF_LEDS) & vbCrLf
    s = s & "   Modül gücü           : " & ac2fFmt(ac2fGetNum(AC2F_K_MODULE_W, AC2F_DEF_MODULE_W)) & " W" & vbCrLf
    s = s & "   Güç kaynağı          : " & ac2fFmt(ac2fGetNum(AC2F_K_PSU_W, AC2F_DEF_PSU_W), 0) & " W" & vbCrLf
    s = s & "   Güvenlik payı        : %" & ac2fFmt(ac2fGetNum(AC2F_K_SAFETY, AC2F_DEF_SAFETY), 0) & vbCrLf
    s = s & "   Kontur başına en az  : " & ac2fGetLng(AC2F_K_MINPERPATH, AC2F_DEF_MINPERPATH) & " modül" & vbCrLf
    s = s & "   Yöntem               : " & ac2fYontemAdi(ac2fGetLng(AC2F_K_METHOD, AC2F_DEF_METHOD)) & vbCrLf
    s = s & "   Düzeltme katsayısı   : " & ac2fFmt(ac2fGetNum(AC2F_K_FACTOR, AC2F_DEF_FACTOR)) & vbCrLf

    ac2fAyarOzeti = s
End Function

' Ondalıklı ayar sorar. İptal edilirse False döner.
Private Function ac2fSorNum(ByVal soru As String, ByVal caption As String, _
                            ByVal key As String, ByVal defValue As Double, _
                            ByVal enAz As Double, ByRef sonuc As Double) As Boolean
    Dim mevcut As Double
    Dim cevap As String
    Dim v As Double

    mevcut = ac2fGetNum(key, defValue)
    cevap = InputBox(soru & vbCrLf & vbCrLf & _
                     "(Varsayılan: " & ac2fNumStr(defValue) & ")", _
                     ac2fTitle(caption), ac2fNumStr(mevcut))
    If StrPtr(cevap) = 0 Then Exit Function

    v = ac2fParseNum(cevap, mevcut)
    If v < enAz Then v = enAz

    ac2fSetNum key, v
    sonuc = v
    ac2fSorNum = True
End Function

' Tam sayılı ayar sorar. İptal edilirse False döner.
Private Function ac2fSorLng(ByVal soru As String, ByVal caption As String, _
                            ByVal key As String, ByVal defValue As Long, _
                            ByVal enAz As Long, ByRef sonuc As Long, _
                            Optional ByVal enCok As Long = 0) As Boolean
    Dim mevcut As Long
    Dim cevap As String
    Dim v As Long

    mevcut = ac2fGetLng(key, defValue)
    cevap = InputBox(soru & vbCrLf & vbCrLf & _
                     "(Varsayılan: " & ac2fNumStr(CDbl(defValue)) & ")", _
                     ac2fTitle(caption), ac2fNumStr(CDbl(mevcut)))
    If StrPtr(cevap) = 0 Then Exit Function

    v = CLng(ac2fParseNum(cevap, CDbl(mevcut)))
    If v < enAz Then v = enAz
    If enCok > 0 And v > enCok Then v = enCok

    ac2fSetLng key, v
    sonuc = v
    ac2fSorLng = True
End Function

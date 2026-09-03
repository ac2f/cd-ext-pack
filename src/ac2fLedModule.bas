Attribute VB_Name = "ac2fLedModule"
'=====================================================================
'  ac2f pack  --  ac2fLedModule
'
'  3'lü modül LED yerleşimi için adet hesabı.
'
'  Ölçülen kontur uzunluğunu modül aralığına bölerek gereken modül
'  sayısını, LED adedini, toplam gücü ve güç kaynağı ihtiyacını çıkarır.
'
'  Makrolar:
'    ac2fLedModulHesapla  - kayıtlı ayarlarla hesaplar
'    ac2fLedHizliHesap    - modül aralığını sorup tek seferlik hesaplar
'=====================================================================
Option Explicit

Private Const CAPTION_ As String = "3'lü LED Modül Hesabı"

' Hesap yöntemleri
Public Const AC2F_METHOD_CEVRE   As Long = 1   ' her konturun tam uzunluğu
Public Const AC2F_METHOD_ORTAHAT As Long = 2   ' kontur uzunluğunun yarısı

' Hesap sonucu
Public Type ac2fLedResult
    Ok           As Boolean
    Message      As String
    LengthMM     As Double   ' hesaba giren efektif uzunluk
    RawLengthMM  As Double   ' ölçülen ham kontur uzunluğu
    PathCount    As Long
    ModuleCount  As Long
    LedCount     As Long
    TotalWatt    As Double
    NeededWatt   As Double
    PsuCount     As Long
    SpacingMM    As Double
    Method       As Long
    Factor       As Double
End Type

'---------------------------------------------------------------------
' Kayıtlı ayarlarla hesaplar.
'---------------------------------------------------------------------
Public Sub ac2fLedModulHesapla()
Attribute ac2fLedModulHesapla.VB_Description = "ac2f pack: 3'lu LED modul adedini hesaplar"
    ac2fLedHesapCalistir ac2fGetNum(AC2F_K_SPACING, AC2F_DEF_SPACING)
End Sub

'---------------------------------------------------------------------
' Modül aralığını sorup tek seferlik hesaplar (ayarları değiştirmez).
'---------------------------------------------------------------------
Public Sub ac2fLedHizliHesap()
Attribute ac2fLedHizliHesap.VB_Description = "ac2f pack: Modul araligini sorarak LED modul adedini hesaplar"
    Dim varsayilan As Double
    Dim cevap As String
    Dim aralik As Double

    varsayilan = ac2fGetNum(AC2F_K_SPACING, AC2F_DEF_SPACING)

    cevap = InputBox("Modül aralığı (mm):", ac2fTitle(CAPTION_), _
                     ac2fNumStr(varsayilan))
    If StrPtr(cevap) = 0 Then Exit Sub      ' İptal

    aralik = ac2fParseNum(cevap, varsayilan)
    If aralik <= 0 Then
        ac2fWarn "Modül aralığı sıfırdan büyük olmalı.", CAPTION_
        Exit Sub
    End If

    ac2fLedHesapCalistir aralik
End Sub

'=====================================================================
' İç fonksiyonlar
'=====================================================================

Private Sub ac2fLedHesapCalistir(ByVal aralikMM As Double)
    Dim res As ac2fResult
    Dim led As ac2fLedResult

    If aralikMM <= 0 Then
        ac2fWarn "Modül aralığı sıfırdan büyük olmalı." & vbCrLf & _
                 "Ayarlar menüsünden düzeltebilirsiniz.", CAPTION_
        Exit Sub
    End If

    res = ac2fMeasureSelection()
    If Not res.Ok Then
        ac2fWarn res.Message, CAPTION_
        Exit Sub
    End If
    If res.SubCount = 0 Then
        ac2fWarn "Seçimde ölçülebilir bir yol bulunamadı.", CAPTION_
        Exit Sub
    End If

    led = ac2fLedHesapla(res, aralikMM)
    If Not led.Ok Then
        ac2fWarn led.Message, CAPTION_
        Exit Sub
    End If

    ac2fInfo ac2fLedRaporu(led), CAPTION_
End Sub

'---------------------------------------------------------------------
' Ölçüm sonucundan modül/LED/güç hesabını yapar.
'---------------------------------------------------------------------
Public Function ac2fLedHesapla(ByRef res As ac2fResult, _
                               ByVal aralikMM As Double) As ac2fLedResult
    Dim out As ac2fLedResult
    Dim i As Long
    Dim uzunluk As Double
    Dim adet As Double
    Dim toplamAdet As Double
    Dim ledBasina As Long
    Dim modulW As Double
    Dim psuW As Double
    Dim pay As Double
    Dim enAz As Long
    Dim yontem As Long
    Dim katsayi As Double

    If aralikMM <= 0 Then
        out.Message = "Modül aralığı sıfırdan büyük olmalı."
        ac2fLedHesapla = out
        Exit Function
    End If

    ledBasina = ac2fGetLng(AC2F_K_LEDS, AC2F_DEF_LEDS)
    modulW = ac2fGetNum(AC2F_K_MODULE_W, AC2F_DEF_MODULE_W)
    psuW = ac2fGetNum(AC2F_K_PSU_W, AC2F_DEF_PSU_W)
    pay = ac2fGetNum(AC2F_K_SAFETY, AC2F_DEF_SAFETY)
    enAz = ac2fGetLng(AC2F_K_MINPERPATH, AC2F_DEF_MINPERPATH)
    yontem = ac2fGetLng(AC2F_K_METHOD, AC2F_DEF_METHOD)
    katsayi = ac2fGetNum(AC2F_K_FACTOR, AC2F_DEF_FACTOR)

    If ledBasina < 1 Then ledBasina = AC2F_DEF_LEDS
    If enAz < 0 Then enAz = 0
    If katsayi <= 0 Then katsayi = 1#
    If yontem <> AC2F_METHOD_ORTAHAT Then yontem = AC2F_METHOD_CEVRE

    For i = 0 To res.SubCount - 1
        uzunluk = res.SubLenMM(i)
        If yontem = AC2F_METHOD_ORTAHAT Then uzunluk = uzunluk / 2#

        If res.SubClosed(i) Then
            ' Kapalı halkada modüller çevre boyunca eşit dağıtılır.
            adet = ac2fCeil(uzunluk / aralikMM)
        Else
            ' Açık yolda iki uç da modül alır.
            adet = Int(uzunluk / aralikMM) + 1#
        End If

        If adet < enAz Then adet = enAz

        toplamAdet = toplamAdet + adet
        out.LengthMM = out.LengthMM + uzunluk
    Next i

    toplamAdet = ac2fCeil(toplamAdet * katsayi)

    out.Ok = True
    out.RawLengthMM = res.TotalMM
    out.PathCount = res.SubCount
    out.SpacingMM = aralikMM
    out.Method = yontem
    out.Factor = katsayi
    out.ModuleCount = CLng(toplamAdet)
    out.LedCount = out.ModuleCount * ledBasina
    out.TotalWatt = out.ModuleCount * modulW
    out.NeededWatt = out.TotalWatt * (1# + pay / 100#)

    If psuW > 0 Then
        out.PsuCount = CLng(ac2fCeil(out.NeededWatt / psuW))
    Else
        out.PsuCount = 0
    End If

    ac2fLedHesapla = out
End Function

'---------------------------------------------------------------------
' Hesap sonucunu okunabilir rapora çevirir.
'---------------------------------------------------------------------
Public Function ac2fLedRaporu(ByRef led As ac2fLedResult) As String
    Dim s As String
    Dim ledBasina As Long
    Dim modulW As Double
    Dim psuW As Double
    Dim pay As Double

    ledBasina = ac2fGetLng(AC2F_K_LEDS, AC2F_DEF_LEDS)
    modulW = ac2fGetNum(AC2F_K_MODULE_W, AC2F_DEF_MODULE_W)
    psuW = ac2fGetNum(AC2F_K_PSU_W, AC2F_DEF_PSU_W)
    pay = ac2fGetNum(AC2F_K_SAFETY, AC2F_DEF_SAFETY)

    s = "SONUÇ" & vbCrLf
    s = s & "   Modül adedi         : " & led.ModuleCount & " adet" & vbCrLf
    s = s & "   LED adedi           : " & led.LedCount & " adet" & _
            "  (" & ledBasina & " LED/modül)" & vbCrLf & vbCrLf

    s = s & "GÜÇ" & vbCrLf
    s = s & "   Toplam güç          : " & ac2fFmt(led.TotalWatt) & " W" & vbCrLf
    s = s & "   %" & ac2fFmt(pay, 0) & " pay ile        : " & _
            ac2fFmt(led.NeededWatt) & " W" & vbCrLf
    If led.PsuCount > 0 Then
        s = s & "   Güç kaynağı         : " & led.PsuCount & " adet x " & _
                ac2fFmt(psuW, 0) & " W" & vbCrLf
    End If
    s = s & vbCrLf

    s = s & "ÖLÇÜM" & vbCrLf
    s = s & "   Ham kontur uzunluğu : " & ac2fFmt(led.RawLengthMM) & " mm" & vbCrLf
    s = s & "   Hesaba giren uzunluk: " & ac2fFmt(led.LengthMM) & " mm" & vbCrLf
    s = s & "   Kontur sayısı       : " & led.PathCount & vbCrLf & vbCrLf

    s = s & "KULLANILAN AYARLAR" & vbCrLf
    s = s & "   Modül aralığı       : " & ac2fFmt(led.SpacingMM) & " mm" & vbCrLf
    s = s & "   Yöntem              : " & ac2fYontemAdi(led.Method) & vbCrLf
    s = s & "   Düzeltme katsayısı  : " & ac2fFmt(led.Factor) & vbCrLf
    s = s & "   Modül gücü          : " & ac2fFmt(modulW) & " W" & vbCrLf

    ac2fLedRaporu = s
End Function

Public Function ac2fYontemAdi(ByVal yontem As Long) As String
    If yontem = AC2F_METHOD_ORTAHAT Then
        ac2fYontemAdi = "Orta hat tahmini (kontur/2)"
    Else
        ac2fYontemAdi = "Çevre bazlı (tam kontur)"
    End If
End Function

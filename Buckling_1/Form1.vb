﻿Imports System.IO
Imports System.Text
Imports System.Math

Public Class Form1
    Public fy As Double         'Yield stress
    Dim Ym As Double            'Safety factor
    Dim E_mod As Double         'Elasticity [N/mm2]

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles NumericUpDown47.ValueChanged, NumericUpDown46.ValueChanged, NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown43.ValueChanged, NumericUpDown42.ValueChanged, NumericUpDown39.ValueChanged, NumericUpDown36.ValueChanged, Button6.Click
        DNV_chapter5_0()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter, NumericUpDown8.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown11.ValueChanged, Button2.Click
        DNV_chapter6_2()    'Unstiffened plate longitudinal uniform loading
        DNV_chapter6_3()    'Unstiffened plate transverse uniform loading
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles TabPage1.Enter, NumericUpDown9.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, Button1.Click
        'Determine the weight and the stress due to weight
        Dim box_l, box_w, box_h, box_thick As Double
        Dim w_roof, w_longpanel, w_shortpanel As Double
        Dim wght_fact, wght_explo, wght_total As Double
        Dim vacuum_p, vacuum_wght As Double
        Dim total_wght_on_walls As Double
        Dim comp_stress, total_wall_area As Double

        'Dimensions
        box_l = NumericUpDown4.Value        'Length
        box_w = NumericUpDown5.Value        'Width
        box_h = NumericUpDown1.Value        'Height
        box_thick = NumericUpDown36.Value   'Wall thickness
        wght_fact = NumericUpDown9.Value    'Stiff+girder+expl_doors+insu
        wght_explo = NumericUpDown2.Value   'Explosion doors on roof

        'Weight
        w_roof = box_l * box_w * box_thick / 1000 ^ 3 * 8000 * wght_fact
        w_roof += wght_explo
        w_longpanel = box_l * box_h * box_thick / 1000 ^ 3 * 8000 * wght_fact
        w_shortpanel = box_w * box_w * box_thick / 1000 ^ 3 * 8000 * wght_fact
        wght_total = w_roof + 2 * w_longpanel + 2 * w_shortpanel

        'Vacuum
        vacuum_p = NumericUpDown45.Value                            '[Pa]
        vacuum_wght = vacuum_p * box_l * box_w / 1000 ^ 2 / 9.81    '[kg]
        total_wght_on_walls = vacuum_wght + wght_total      '[kg]

        'Traverse compression stress
        total_wall_area = (box_l + box_w) * 2 * box_thick           '[mm2]
        comp_stress = total_wght_on_walls * 9.81 / total_wall_area  '[N/mm2]

        TextBox5.Text = Math.Round(box_thick, 0).ToString
        TextBox6.Text = Math.Round(w_roof, 0).ToString
        TextBox7.Text = Math.Round(w_longpanel, 0).ToString
        TextBox8.Text = Math.Round(w_shortpanel, 0).ToString
        TextBox9.Text = Math.Round(wght_total, 0).ToString
        TextBox10.Text = Math.Round(comp_stress, 1).ToString
        TextBox11.Text = Math.Round(vacuum_p, 0).ToString
        TextBox12.Text = Math.Round(vacuum_wght, 0).ToString
        TextBox14.Text = Math.Round(total_wght_on_walls, 0).ToString

        'Transfer result to next tab
        If comp_stress < NumericUpDown43.Maximum And comp_stress > NumericUpDown43.Minimum Then
            NumericUpDown43.Value = Math.Round(comp_stress, 0)
        End If
    End Sub
    Private Sub DNV_chapter5_0()
        'Çhapter 5, DNV-RP-C201, Oktober 2010 
        Dim t As Double
        Dim s As Double
        Dim length As Double

        Dim upsilon_X As Double
        Dim upsilon_Y As Double

        'sd stands for design stress or load
        Dim Psd As Double
        Dim sigma_Xsd As Double
        Dim sigma_Ysd As Double
        Dim sigma_Jsd As Double 'Von Misses
        Dim tau_sd As Double
        Dim Psd_check As Double
        Dim epsilon As Double

        s = NumericUpDown39.Value           'Stiffeners distance
        t = NumericUpDown36.Value           'Plate thickness
        length = NumericUpDown47.Value      'stiffeners length

        Psd = NumericUpDown45.Value / 1000 ^ 2  '[N/mm2]
        sigma_Xsd = NumericUpDown43.Value
        sigma_Ysd = NumericUpDown42.Value
        tau_sd = NumericUpDown44.Value

        'Calc Von misses (equation 5.4)
        sigma_Jsd = Math.Sqrt(sigma_Xsd ^ 2 + sigma_Ysd ^ 2 - sigma_Xsd * sigma_Ysd + 3 * tau_sd ^ 2)
        NumericUpDown46.Value = sigma_Jsd

        'Calc constant (equation 5.2)
        upsilon_Y = 1 - (sigma_Jsd / fy) ^ 2
        upsilon_Y = upsilon_Y / Math.Sqrt(1 - 3 / 4 * (sigma_Xsd / fy) ^ 2) - 3 * (tau_sd / fy) ^ 2

        'Calc constant (equation 5.3)
        upsilon_X = 1 - (sigma_Jsd / fy) ^ 2
        upsilon_X /= Math.Sqrt(1 - 3 / 4 * (sigma_Ysd / fy) ^ 2) - 3 * (tau_sd / fy) ^ 2


        'Psd Check (equation 5.1)
        Psd_check = 4 * fy / Ym * ((t / s) ^ 2) * (upsilon_Y + ((s / length) ^ 2 * upsilon_X))

        TextBox1.Text = Math.Round(upsilon_X, 4).ToString
        TextBox2.Text = Math.Round(upsilon_Y, 4).ToString
        TextBox4.Text = Math.Round(Psd_check, 4).ToString

        'Checks
        NumericUpDown39.BackColor = IIf(s < length, Color.Yellow, Color.LightCoral)
        NumericUpDown47.BackColor = IIf(s < length, Color.Yellow, Color.LightCoral)

        If Psd < Psd_check Then
            TextBox4.BackColor = Color.LightGreen
            Label118.Visible = False
        Else
            TextBox4.BackColor = Color.Red
            Label118.Visible = True
        End If

        epsilon = Math.Sqrt(235 / fy)
        Label18.Text = IIf(s / t < 5.4 * epsilon, "Buckling check NOT neccessary", "Buckling check is neccessary !!")
    End Sub

    Private Sub DNV_chapter6_2()
        '============ Longitudinal uniform compression chapter 6.2 ===========
        Dim t, S As Double

        Dim lambda_P, Cx, sigma_Xrd, sigma_Xsd As Double

        sigma_Xsd = NumericUpDown43.Value   'Design stress
        S = NumericUpDown12.Value           'Stiffeners distance
        t = NumericUpDown11.Value           'Plate thickness

        lambda_P = 0.525 * S / t * Math.Sqrt(fy / E_mod)    'formule 6.3

        If lambda_P < 0.673 Then            'formule 6.2
            Cx = 1
        Else
            Cx = (lambda_P - 0.22) / lambda_P ^ 2
        End If

        sigma_Xrd = Cx * fy / Ym           'formule 6.1

        TextBox3.Text = Math.Round(lambda_P, 2).ToString
        TextBox15.Text = Math.Round(Cx, 2).ToString
        TextBox16.Text = Math.Round(sigma_Xrd, 0).ToString
        TextBox63.Text = Math.Round(sigma_Xsd, 0).ToString
    End Sub

    Private Sub DNV_chapter6_3()
        '============ Transverse uniform compression chapter 6.3 ===========
        Dim t, S, psd, sigma_Ysd As Double
        Dim length As Double

        sigma_Ysd = NumericUpDown43.Value   'Design stress
        S = NumericUpDown12.Value           'Stiffeners distance
        t = NumericUpDown11.Value           'Plate thickness
        length = NumericUpDown8.Value      'stiffeners length
        psd = NumericUpDown24.Value / 1000 ^ 2  '[N/mm2]

        '============ Transverserse uniform compression chapter 6.3===========
        Dim lambda_C, kappa, mu As Double
        Dim Kp, h_alfa, sigma_YR, sigma_YRd As Double

        '============= h_alfa ============= Formule 6.11
        h_alfa = 0.05 * S / t - 0.75

        '========== kp  ============ Formule 6.10
        If psd <= 2 * (t / S) ^ 2 * fy Then
            Kp = 1
        Else
            Kp = 1 - h_alfa * (psd / fy - 2 * (t / S) ^ 2)
            If Kp < 0 Then Kp = 0
        End If

        '======== lambda_C ========= Formule 6.8
        lambda_C = 1.1 * S / t * Math.Sqrt(fy / E_mod)

        '=========== kappa======= Formule 6.7
        If lambda_C <= 0.2 Then kappa = 1
        If lambda_C > 0.2 And lambda_C < 0.2 Then
            kappa = 1 / (2 * lambda_C ^ 2)
            kappa *= 1 + mu + lambda_C ^ 2 - Math.Sqrt((1 + mu + lambda_C ^ 2) ^ 2 - 4 * lambda_C ^ 2)
        End If
        If lambda_C >= 2 Then kappa = 0.07 + 1 / (2 * lambda_C ^ 2)

        '========== mu========== Formule 6.9
        mu = 0.21 * (lambda_C - 0.2)

        '========== sigma_YR ========== Formule 6.6
        sigma_YR = 1.3 * t / length * Math.Sqrt(E_mod / fy)
        sigma_YR += kappa * (1 - 1.3 * t / length * Math.Sqrt(E_mod / fy))
        sigma_YR *= fy * Kp

        '========== sigma_YR ========== Formule 6.5
        sigma_YRd = sigma_YR / Ym

        TextBox65.Text = Math.Round(S, 1).ToString
        TextBox66.Text = Math.Round(length, 2).ToString
        TextBox67.Text = Math.Round(t, 1).ToString

        TextBox17.Text = Math.Round(h_alfa, 2).ToString     '6.11
        TextBox18.Text = Math.Round(Kp, 2).ToString         '6.10
        TextBox19.Text = Math.Round(mu, 0).ToString         '6.9
        TextBox20.Text = Math.Round(lambda_C, 2).ToString   '6.8
        TextBox21.Text = Math.Round(kappa, 2).ToString      '6.7
        TextBox22.Text = Math.Round(sigma_YR, 0).ToString   '6.6

        TextBox24.Text = Math.Round(sigma_YRd, 0).ToString  '6.5
        TextBox64.Text = Math.Round(sigma_Ysd, 0).ToString

        '---------------- Check stress------------------
        TextBox16.BackColor = IIf(sigma_YRd >= sigma_Ysd, Color.LightGreen, Color.Red)
    End Sub

    Private Sub DNV_chapter6_6()
        Dim t, S, sigma_Xsd, sigma_xRd, kappa_sigma, upsilon, lambda_P, Psi, cx As Double

        S = NumericUpDown21.Value           'Plate width
        t = NumericUpDown29.Value           'Plate thickness
        Psi = 1                             'Stress ratio is 1.0
        sigma_Xsd = NumericUpDown43.Value   'Design stress


        kappa_sigma = 8.2 / (1.05 + 1)      'equation 6.44

        '------------- slenderness calculation--------------------
        upsilon = Sqrt(235 / fy)
        lambda_P = S / t * 1 / (28.4 * upsilon * Sqrt(kappa_sigma)) 'equation 6.24

        If lambda_P <= 0.673 Then
            cx = 1                          'equation 6.22
        Else
            cx = (lambda_P - 0.055 * (3 + 1)) / lambda_P ^ 2     'equation 6.23
        End If

        sigma_xRd = cx * fy / Ym            'equation 6.21

        TextBox50.Text = Math.Round(kappa_sigma, 2).ToString
        TextBox51.Text = Math.Round(upsilon, 2).ToString
        TextBox52.Text = Math.Round(cx, 2).ToString
        TextBox53.Text = Math.Round(sigma_xRd, 1).ToString
        TextBox54.Text = Math.Round(lambda_P, 2).ToString
        TextBox61.Text = Math.Round(sigma_Xsd, 2).ToString

        '---------------- Check stress------------------
        TextBox24.BackColor = IIf(sigma_xRd >= sigma_Xsd, Color.LightGreen, Color.Red)
    End Sub
    Private Sub DNV_chapter6_7()
        Dim t, S, sigma_Xsd, sigma_xRd, kappa_sigma, upsilon, lambda_P, Psi, cx As Double

        S = NumericUpDown21.Value           'Plate width
        t = NumericUpDown29.Value           'Plate thickness
        Psi = 1                             'Stress ratio is 1.0
        sigma_Xsd = NumericUpDown43.Value   'Design stress

        If RadioButton3.Checked Then
            kappa_sigma = 0.57 - 0.21 * Psi + 0.07 * Psi ^ 2
        Else
            kappa_sigma = 0.578 / (0.34 + Psi)
        End If

        '------------- slenderness calculation--------------------
        upsilon = Sqrt(235 / fy)

        lambda_P = S / t * 1 / (28.4 * upsilon * Sqrt(kappa_sigma)) 'equation 6.29

        If lambda_P <= 0.749 Then
            cx = 1                          'equation 6.22
        Else
            cx = (lambda_P - 0.188) / lambda_P ^ 2     'equation 6.28
        End If

        sigma_xRd = cx * fy / Ym            'equation 6.21

        TextBox55.Text = Math.Round(kappa_sigma, 2).ToString
        TextBox56.Text = Math.Round(upsilon, 2).ToString
        TextBox57.Text = Math.Round(cx, 2).ToString
        TextBox58.Text = Math.Round(sigma_xRd, 1).ToString
        TextBox59.Text = Math.Round(t, 1).ToString
        TextBox60.Text = Math.Round(S, 0).ToString
        TextBox13.Text = Math.Round(lambda_P, 2).ToString
        TextBox62.Text = Math.Round(sigma_Xsd, 2).ToString

        '---------------- Check stress------------------
        TextBox58.BackColor = IIf(sigma_xRd >= sigma_Xsd, Color.LightGreen, Color.Red)

    End Sub

    Private Sub DNV_chapter7_2()
        '============ Forces in the idealised stiffener plate chapter 7.2 ===========
        Dim t, S, sigma_Xsd, sigma_Y1sd, N_Sd, A_s As Double
        Dim tau_tf, tau_crl, tau_crg As Double
        Dim length, LG, psi, I_s As Double
        Dim q_sd, psd, p_0, C0, Kc As Double
        Dim K_l, k_g As Double

        S = NumericUpDown14.Value           'Stiffeners distance
        t = NumericUpDown13.Value           'Plate thickness
        length = NumericUpDown23.Value      'stiffeners length
        LG = NumericUpDown28.Value          'Girder length
        psd = NumericUpDown22.Value / 1000 ^ 2  '[N/mm2]
        sigma_Xsd = NumericUpDown16.Value   'design stress axial stress in plate and stiffener
        sigma_Y1sd = NumericUpDown25.Value  'design stress transverse direction
        tau_tf = NumericUpDown20.Value      'shear sress in plate and stiffener
        A_s = NumericUpDown3.Value          'Area cross sectional stiffener

        'Moment of inertia stiffener+plate
        Dim I_stif, I_plate As Double
        Dim Stif_Thick, Stif_Heigh As Double
        Dim Wes, mc As Double

        Stif_Thick = NumericUpDown26.Value
        Stif_Heigh = NumericUpDown27.Value

        I_stif = Stif_Thick * Stif_Heigh ^ 3 / 3    'Stiffener
        I_plate = S * t ^ 3 / 3                     'Plate
        I_s = I_stif + I_plate
        Wes = I_s / (Stif_Heigh * 0.5)              'section modulud plate at flange tip ?????

        TextBox30.Text = Math.Round(I_stif, 0).ToString
        TextBox31.Text = Math.Round(I_plate, 0).ToString
        TextBox33.Text = Math.Round(I_s, 0).ToString


        '========== Equivalent lateral load ========== 
        Kc = 2 * (1 + Sqrt(1 + (10.9 * I_s / (t ^ 3 * S))))   'equation 7.12
        psi = 1                             'equation 7.11
        If RadioButton1.Checked Then
            mc = 13.3       'Continuous stiffeners
        Else
            mc = 8.9        'Snipped stiffeners
        End If

        C0 = Wes * fy * mc / (Kc * E_mod * t ^ 2 * S)   'equation 7.11 
        'p_0 equation is not applicable                  'equation 7.10
        p_0 = (0.6 + 0.4) * C0 * sigma_Y1sd             'equation 7.9
        q_sd = (psd + p_0) * t                          'equation 7.8
        If length >= S Then
            K_l = 5.34 + 4 * (t / S) ^ 2                'equation 7.7
        Else
            K_l = 5.34 * (t / S) ^ 2 + 4
        End If
        tau_crl = K_l * 0.904 * E_mod * (t / S) ^ 2     'equation 7.6
        If length <= LG Then
            k_g = 5.34 + 4 * (length / LG) ^ 2          'equation 7.5
        Else
            k_g = 5.34 * (length / LG) ^ 2 + 4
        End If
        tau_crg = k_g * 0.904 * E_mod * (t / length) ^ 2 'equation 7.4

        '========== Equivalent axial load ========== 
        N_Sd = sigma_Xsd * (A_s + S * t) + tau_tf * S * t 'Formule 7.1

        TextBox27.Text = Math.Round(C0, 7).ToString
        TextBox28.Text = Math.Round(psi, 2).ToString
        TextBox29.Text = Math.Round(Kc, 2).ToString
        TextBox40.Text = Math.Round(N_Sd * 0.001, 1).ToString    '[kN]
        TextBox34.Text = Math.Round(p_0, 6).ToString
        TextBox26.Text = Math.Round(q_sd, 2).ToString
        TextBox35.Text = Math.Round(tau_crl, 0).ToString
        TextBox36.Text = Math.Round(K_l, 2).ToString
        TextBox37.Text = Math.Round(k_g, 2).ToString
        TextBox38.Text = Math.Round(tau_crg, 0).ToString
    End Sub
    Private Sub DNV_chapter7_3()
        Dim t, S, sigma_Xsd, sigma_Ysd, sigma_Yr As Double
        Dim Cys, Cxs, Ci, lambda_p, Se As Double
        S = NumericUpDown14.Value           'Stiffeners distance
        t = NumericUpDown13.Value           'Plate thickness
        sigma_Xsd = NumericUpDown16.Value   'design stress axial stress in plate and stiffener
        sigma_Ysd = NumericUpDown25.Value   'design stress transverse direction

        Double.TryParse(TextBox22.Text, sigma_Yr)       'equation 6.6

        '--------------- equation 7.15---
        lambda_p = 0.525 * S / t * Sqrt(fy / E_mod)

        '--------------- equation 7.14---
        Cxs = (lambda_p - 0.22) / lambda_p ^ 2

        '--------------- equation 7.16---
        If (S / t <= 120) Then
            Ci = 1 - (S / (120 * t))
        Else
            Ci = 0
        End If

        Cys = (sigma_Ysd / sigma_Yr) ^ 2
        Cys = Cys + Ci * (sigma_Xsd * sigma_Ysd / (Cxs * fy * sigma_Yr))
        Cys = Sqrt(1 - Cys)

        '--------------- equation 7.13---
        Se = S * Cxs * Cys

        TextBox41.Text = Math.Round(S, 0).ToString
        TextBox42.Text = Math.Round(t, 0).ToString
        TextBox44.Text = Math.Round(sigma_Xsd, 0).ToString
        TextBox43.Text = Math.Round(sigma_Ysd, 0).ToString
        TextBox45.Text = Math.Round(fy, 0).ToString
        TextBox46.Text = Math.Round(Cxs, 2).ToString
        TextBox47.Text = Math.Round(Cys, 2).ToString
        TextBox48.Text = Math.Round(sigma_Yr, 0).ToString
        TextBox49.Text = Math.Round(Se, 0).ToString
        TextBox68.Text = Math.Round(Ci, 3).ToString
        TextBox69.Text = Math.Round(lambda_p, 2).ToString
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        get_material_data()

        TextBox23.Text =
        "Based on " & vbCrLf &
        "Recommended Practice DNV-RP-C201, October 2010" & vbCrLf &
        "Buckling Strength of plated structures" & vbCrLf &
        "https://rules.dnvgl.com/servicedocuments/dnv" & vbCrLf & vbCrLf &
        "See also" & vbCrLf &
        "http://www.steelconstruction.info/Stiffeners"

        TextBox25.Text =
      vbTab & "Construction steel S235JR" & vbTab & "@ 20c" & vbTab & "235 [N/mm2]" & vbCrLf &
      vbTab & "Construction steel S355JR" & vbTab & "@ 20c" & vbTab & "335 [N/mm2]" & vbCrLf &
      vbTab & "Stainless steel 304L" & vbTab & vbTab & "@ 20c" & vbTab & "220 [N/mm2]" & vbCrLf &
      vbTab & "Stainless steel 316L" & vbTab & vbTab & "@ 20c" & vbTab & "195 [N/mm2]"

        DNV_chapter6_3()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles NumericUpDown7.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown10.ValueChanged, Button4.Click
        get_material_data()
    End Sub
    Private Sub get_material_data()
        fy = NumericUpDown19.Value          'Yield stress
        E_mod = NumericUpDown7.Value        'Elasticity [N/mm2]
        Ym = NumericUpDown10.Value          'Safety
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged, RadioButton1.CheckedChanged, NumericUpDown3.ValueChanged, NumericUpDown27.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown22.ValueChanged, NumericUpDown20.VisibleChanged, NumericUpDown20.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged, GroupBox13.VisibleChanged, Button5.Click, NumericUpDown28.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown23.ValueChanged
        DNV_chapter7_2() 'Chapter 7.2
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabPage9.Enter
        DNV_chapter7_3() 'Chapter 7.3
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click, TabPage4.Enter, NumericUpDown29.ValueChanged, NumericUpDown21.ValueChanged, RadioButton3.CheckedChanged
        DNV_chapter6_6() 'Chapter 6.6
        DNV_chapter6_7() 'Chapter 6.7
    End Sub


End Class
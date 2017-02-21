Imports System.IO
Imports System.Text
Imports System.Math

Public Class Form1
    Public fy As Double         'Yield stress
    Public Ym As Double         'Safety factor
    Public E_mod As Double      'Elasticity [N/mm2]
    Public G As Double          'Shear modulus
    Public psd As Double        'Lateral load [N/mm2]
    Public _t As Double         'Plate thickness
    Public _s As Double         'Stiffeners distance
    Public _l As Double         'stiffeners length



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles TabPage1.Enter, NumericUpDown9.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, Button1.Click
        Calc_weight_and_loads()
    End Sub
    Private Sub Calc_weight_and_loads()
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
        Dim upsilon_X As Double
        Dim upsilon_Y As Double

        'sd stands for design stress or load
        Dim sigma_Xsd, sigma_Ysd, tau_sd As Double
        Dim sigma_Jsd As Double 'Von Misses
        Dim Psd_check As Double
        Dim epsilon As Double

        sigma_Xsd = NumericUpDown43.Value
        sigma_Ysd = NumericUpDown42.Value
        tau_sd = NumericUpDown44.Value

        'Calc Von misses (equation 5.4)
        sigma_Jsd = Math.Sqrt(sigma_Xsd ^ 2 + sigma_Ysd ^ 2 - sigma_Xsd * sigma_Ysd + 3 * tau_sd ^ 2)

        'Calc constant (equation 5.3)
        upsilon_X = 1 - (sigma_Jsd / fy) ^ 2
        upsilon_X /= Math.Sqrt(1 - (3 / 4 * (sigma_Ysd / fy) ^ 2)) - 3 * (tau_sd / fy) ^ 2

        'Calc constant (equation 5.2)
        upsilon_Y = 1 - (sigma_Jsd / fy) ^ 2
        upsilon_Y /= Math.Sqrt(1 - (3 / 4 * (sigma_Xsd / fy) ^ 2)) - 3 * (tau_sd / fy) ^ 2

        'Psd Check (equation 5.1)
        Psd_check = 4 * fy / Ym * ((_t / _s) ^ 2)
        Psd_check *= upsilon_Y + ((_s / _l) ^ 2 * upsilon_X)

        'Results
        TextBox130.Text = Math.Round(Psd, 4).ToString
        TextBox107.Text = Math.Round(sigma_Jsd, 1).ToString
        TextBox1.Text = Math.Round(upsilon_X, 4).ToString
        TextBox2.Text = Math.Round(upsilon_Y, 4).ToString
        TextBox4.Text = Math.Round(Psd_check, 4).ToString


        'Checks
        NumericUpDown39.BackColor = IIf(_s < _l, Color.Yellow, Color.LightCoral)
        NumericUpDown47.BackColor = IIf(_s < _l, Color.Yellow, Color.LightCoral)

        If Psd < Psd_check Then
            TextBox4.BackColor = Color.LightGreen
            Label118.Visible = False
        Else
            TextBox4.BackColor = Color.Red
            Label118.Visible = True
        End If

        epsilon = Math.Sqrt(235 / fy)
        Label18.Text = IIf(_s / _t < 5.4 * epsilon, "Buckling check NOT neccessary", "Buckling check is neccessary !!")
    End Sub

    Private Sub DNV_chapter6_2()
        '============ Longitudinal uniform compression chapter 6.2 ===========
        Dim lambda_P, Cx, sigma_Xrd, sigma_Xsd As Double
        sigma_Xsd = NumericUpDown43.Value   'Design stress
        lambda_P = 0.525 * _s / _t * Math.Sqrt(fy / E_mod)    'formule 6.3

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
        TextBox169.Text = Math.Round(_s, 0).ToString
        TextBox177.Text = Math.Round(_l, 0).ToString
        TextBox178.Text = Math.Round(_t, 0).ToString
    End Sub

    Private Sub DNV_chapter6_3()
        '============ Transverse uniform compression chapter 6.3 ===========
        Dim sigma_Ysd As Double
        sigma_Ysd = NumericUpDown43.Value   'Design stress

        '============ Transverserse uniform compression chapter 6.3===========
        Dim lambda_C, kappa, mu As Double
        Dim Kp, h_alfa, sigma_YR, sigma_YRd As Double

        '============= h_alfa ============= Formule 6.11
        h_alfa = 0.05 * _s / _t - 0.75

        '========== kp  ============ Formule 6.10
        If psd <= 2 * (_t / _s) ^ 2 * fy Then
            Kp = 1
        Else
            Kp = 1 - h_alfa * (psd / fy - 2 * (_t / _s) ^ 2)
            If Kp < 0 Then Kp = 0
        End If

        '======== lambda_C ========= Formule 6.8
        lambda_C = 1.1 * _s / _t * Math.Sqrt(fy / E_mod)

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
        sigma_YR = 1.3 * _t / _l * Math.Sqrt(E_mod / fy)
        sigma_YR += kappa * (1 - 1.3 * _t / _l * Math.Sqrt(E_mod / fy))
        sigma_YR *= fy * Kp

        '========== sigma_YR ========== Formule 6.5
        sigma_YRd = sigma_YR / Ym

        TextBox65.Text = Math.Round(_s, 1).ToString
        TextBox66.Text = Math.Round(_l, 2).ToString
        TextBox67.Text = Math.Round(_t, 1).ToString

        TextBox17.Text = Math.Round(h_alfa, 2).ToString     '6.11
        TextBox18.Text = Math.Round(Kp, 2).ToString         '6.10
        TextBox19.Text = Math.Round(mu, 0).ToString         '6.9
        TextBox20.Text = Math.Round(lambda_C, 2).ToString   '6.8
        TextBox21.Text = Math.Round(kappa, 2).ToString      '6.7
        TextBox22.Text = Math.Round(sigma_YR, 0).ToString   '6.6

        TextBox24.Text = Math.Round(sigma_YRd, 0).ToString  '6.5
        TextBox64.Text = Math.Round(sigma_Ysd, 0).ToString
        TextBox132.Text = Math.Round(psd, 0).ToString

        '---------------- Check stress------------------
        TextBox16.BackColor = IIf(sigma_YRd >= sigma_Ysd, Color.LightGreen, Color.Red)
    End Sub

    Private Sub DNV_chapter6_6()
        Dim sigma_Xsd, sigma_xRd, kappa_sigma, upsilon, lambda_P, Psi, cx As Double

        Psi = 1                             'Stress ratio is 1.0
        sigma_Xsd = NumericUpDown43.Value   'Design stress

        kappa_sigma = 8.2 / (1.05 + 1)      'equation 6.44

        '------------- slenderness calculation--------------------
        upsilon = Sqrt(235 / fy)
        lambda_P = _s / _t * 1 / (28.4 * upsilon * Sqrt(kappa_sigma)) 'equation 6.24

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

        TextBox179.Text = Math.Round(_s, 2).ToString
        TextBox180.Text = Math.Round(_t, 2).ToString
        '---------------- Check stress------------------
        TextBox24.BackColor = IIf(sigma_xRd >= sigma_Xsd, Color.LightGreen, Color.Red)
    End Sub
    Private Sub DNV_chapter6_7()
        Dim sigma_Xsd, sigma_xRd, kappa_sigma, upsilon, lambda_P, Psi, cx As Double

        Psi = 1                             'Stress ratio is 1.0
        sigma_Xsd = NumericUpDown43.Value   'Design stress

        If RadioButton3.Checked Then
            kappa_sigma = 0.57 - 0.21 * Psi + 0.07 * Psi ^ 2
        Else
            kappa_sigma = 0.578 / (0.34 + Psi)
        End If

        '------------- slenderness calculation--------------------
        upsilon = Sqrt(235 / fy)

        lambda_P = _s / _t * 1 / (28.4 * upsilon * Sqrt(kappa_sigma)) 'equation 6.29

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
        TextBox59.Text = Math.Round(_t, 1).ToString
        TextBox60.Text = Math.Round(_s, 0).ToString
        TextBox13.Text = Math.Round(lambda_P, 2).ToString
        TextBox62.Text = Math.Round(sigma_Xsd, 2).ToString

        '---------------- Check stress------------------
        TextBox58.BackColor = IIf(sigma_xRd >= sigma_Xsd, Color.LightGreen, Color.Red)

    End Sub

    Private Sub DNV_chapter7_2()
        '============ Forces in the idealised stiffener plate chapter 7.2 ===========
        Dim sigma_Xsd, sigma_Y1sd, N_Sd, A_s As Double
        Dim tau_tf, tau_crl, tau_crg As Double
        Dim LG, psi, I_s As Double
        Dim q_sd, p_0, C0, Kc As Double
        Dim K_l, k_g As Double

        LG = NumericUpDown28.Value          'Girder length

        sigma_Xsd = NumericUpDown16.Value   'design stress axial stress in plate and stiffener
        sigma_Y1sd = NumericUpDown25.Value  'design stress transverse direction
        tau_tf = NumericUpDown20.Value      'shear sress in plate and stiffener
        Double.TryParse(TextBox116.Text, A_s) 'Area cross sectional stiffener

        'Moment of inertia stiffener+plate
        Dim I_stif, I_plate As Double
        Dim Stif_Heigh As Double
        Dim Wes, mc As Double

        Stif_Heigh = NumericUpDown35.Value
        Double.TryParse(TextBox115.Text, I_stif)    'Moment of inertia stiffener

        I_plate = _s * _t ^ 3 / 3                     'Plate
        I_s = I_stif + I_plate
        Wes = I_s / (Stif_Heigh * 0.5)              'section modulud plate at flange tip ?????

        TextBox30.Text = Math.Round(I_stif, 0).ToString
        TextBox31.Text = Math.Round(I_plate, 0).ToString
        TextBox33.Text = Math.Round(I_s, 0).ToString

        '========== Equivalent lateral load ========== 
        Kc = 2 * (1 + Sqrt(1 + (10.9 * I_s / (_t ^ 3 * _s))))   'equation 7.12
        psi = 1                             'equation 7.11
        If RadioButton1.Checked Then
            mc = 13.3       'Continuous stiffeners
        Else
            mc = 8.9        'Snipped stiffeners
        End If

        C0 = Wes * fy * mc / (Kc * E_mod * _t ^ 2 * _s)     'equation 7.11 
        'p_0 equation is not applicable                     'equation 7.10
        p_0 = (0.6 + 0.4) * C0 * sigma_Y1sd                 'equation 7.9
        q_sd = (psd + p_0) * _t                             'equation 7.8
        If _l >= _s Then
            K_l = 5.34 + 4 * (_t / _s) ^ 2                  'equation 7.7
        Else
            K_l = 5.34 * (_t / _s) ^ 2 + 4
        End If
        tau_crl = K_l * 0.904 * E_mod * (_t / _s) ^ 2       'equation 7.6
        If _l <= LG Then
            k_g = 5.34 + 4 * (_l / LG) ^ 2                  'equation 7.5
        Else
            k_g = 5.34 * (_l / LG) ^ 2 + 4
        End If
        tau_crg = k_g * 0.904 * E_mod * (_t / _l) ^ 2       'equation 7.4

        '========== Equivalent axial load ========== 
        N_Sd = sigma_Xsd * (A_s + _s * _t) + tau_tf * _s * _t 'Formule 7.1

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
        TextBox133.Text = Math.Round(psd, 4).ToString

        TextBox181.Text = Math.Round(_s, 0).ToString
        TextBox182.Text = Math.Round(_t, 0).ToString
        TextBox183.Text = Math.Round(_l, 0).ToString
        TextBox184.Text = Math.Round(A_s, 0).ToString
    End Sub
    Private Sub DNV_chapter7_3()
        Dim sigma_Xsd, sigma_Ysd, sigma_Yr As Double
        Dim Cys, Cxs, Ci, lambda_p, Se As Double

        sigma_Xsd = NumericUpDown16.Value   'design stress axial stress in plate and stiffener
        sigma_Ysd = NumericUpDown25.Value   'design stress transverse direction

        Double.TryParse(TextBox22.Text, sigma_Yr)       'equation 6.6

        '--------------- equation 7.15---
        lambda_p = 0.525 * _s / _t * Sqrt(fy / E_mod)

        '--------------- equation 7.14---
        Cxs = (lambda_p - 0.22) / lambda_p ^ 2

        '--------------- equation 7.16---
        If (_s / _t <= 120) Then
            Ci = 1 - (_s / (120 * _t))
        Else
            Ci = 0
        End If

        Cys = (sigma_Ysd / sigma_Yr) ^ 2
        Cys = Cys + Ci * (sigma_Xsd * sigma_Ysd / (Cxs * fy * sigma_Yr))
        Cys = Sqrt(1 - Cys)

        '--------------- equation 7.13---
        Se = _s * Cxs * Cys

        TextBox41.Text = Math.Round(_s, 0).ToString
        TextBox42.Text = Math.Round(_t, 0).ToString
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
    Private Sub DNV_chapter7_4()
        Dim K_sp, tau_sd, tau_rd, sigma_YRd, KspsigmaYrd As Double

        Double.TryParse(TextBox24.Text, sigma_YRd)      'equation 6.5
        tau_sd = NumericUpDown44.Value

        K_sp = Sqrt(1 - 3 * (tau_sd / fy) ^ 2)          'equation 7.20
        KspsigmaYrd = K_sp * sigma_YRd                  'equation 7.19
        tau_rd = fy / ((Sqrt(3) * Ym))                  'equation 7.18

        TextBox70.Text = Math.Round(sigma_YRd, 0).ToString
        TextBox71.Text = Math.Round(K_sp, 3).ToString
        TextBox72.Text = Math.Round(tau_sd, 2).ToString
        TextBox73.Text = Math.Round(tau_rd, 2).ToString
        TextBox74.Text = Math.Round(KspsigmaYrd, 0).ToString
    End Sub
    Private Sub DNV_chapter7_51()  'General
        Dim mu1, mu2, Ie, iee, Ae, fe, fr, fT, lambda, lambdaT, fk_fr As Double
        Dim Zp, Zt, lk, fet As Double
        Dim pf, W, Wep, Wes As Double

        Double.TryParse(TextBox115.Text, Ie)
        Double.TryParse(TextBox103.Text, Zp)
        Double.TryParse(TextBox102.Text, Zt)
        Double.TryParse(TextBox116.Text, Ae)
        Double.TryParse(TextBox87.Text, iee)
        Double.TryParse(TextBox99.Text, fT)     'SECTION 7.52
        Double.TryParse(TextBox97.Text, fet)

        Wep = Ie / Zp                            '(equation 7.71a)
        Wes = Ie / Zt
        W = IIf(Wep < Wes, Wep, Wes)

        ' MessageBox.Show("W=" & W.ToString & ",fy=" & fy.ToString & ",L=" & _l.ToString & ",S=" & _S.ToString & ",Ym=" & Ym.ToString)
        pf = 12 * W * fy / (_l ^ 2 * _s * Ym)        '(equation 7.75)
        'lk = _l * (1 - 0.5 * Abs(psd / pf))        '(equation 7.74)
        lk = _l                                     'equ 7.75 only for simple supported stiffeners DNV_chapter7_51

        lambdaT = Sqrt(fy / fet)                            '(equation 7.30)
        'fr = fy                 'Plate side  MOET NOG UITGEWERKT WORDEN

        fr = IIf(lambdaT <= 0.6, fy, fT)
        fe = PI ^ 2 * E_mod * (iee / lk) ^ 2                'equation 7.24
        lambda = Sqrt(fr / fe)                              'equation 7.23

        If lambda <= 0.2 Then       '========== Lambda <= 0.2======
            fk_fr = 1                                           'equation 7.21
        Else                        '========== Lambda > 0.2======


            mu1 = (0.34 + 0.08 * Zt / iee) * (lambda - 0.2)     'equation 7.26
            mu2 = (0.34 + 0.08 * Zp / iee) * (lambda - 0.2)     'equation 7.25


            fk_fr = 1 + mu1 + lambda ^ 2           'equation 7.22 (opgelet mu-mu1,mu2 ??)
            fk_fr -= Sqrt((1 + mu1 + lambda ^ 2) ^ 2 - 4 * lambda ^ 2)
            fk_fr /= 2 * lambda ^ 2
        End If

        TextBox76.Text = Math.Round(mu1, 2).ToString
        TextBox75.Text = Math.Round(mu2, 2).ToString
        TextBox77.Text = Math.Round(fe, 0).ToString
        TextBox78.Text = Math.Round(lambda, 3).ToString
        TextBox79.Text = Math.Round(fk_fr, 2).ToString
        TextBox85.Text = Math.Round(lambdaT, 2).ToString
        TextBox86.Text = Math.Round(lk, 2).ToString
        TextBox88.Text = Math.Round(fT, 2).ToString
        TextBox119.Text = Math.Round(Zp, 3).ToString
        TextBox117.Text = Math.Round(Zt, 3).ToString
        TextBox121.Text = Math.Round(pf, 3).ToString
        TextBox122.Text = Math.Round(W, 0).ToString
        TextBox123.Text = Math.Round(Ie, 0).ToString
        TextBox124.Text = Math.Round(Wes, 0).ToString
        TextBox125.Text = Math.Round(Wep, 0).ToString
        TextBox126.Text = Math.Round(_l, 0).ToString
        TextBox127.Text = Math.Round(fr, 0).ToString

        If lambda <= 0.6 Then Label305.Text = "Note: this is a Slender structure"
        If lambda > 0.6 And lambda < 1.4 Then Label305.Text = "Note: this is Moderate slender structure"
        If lambda >= 1.4 Then Label305.Text = "Note: this is a Stocky structure"
    End Sub

    Private Sub DNV_chapter7_52()
        Dim F_Ept, F_Epy, F_Epx, c, lambda_e As Double
        Dim F_ep, Sigma_JSd, eta, CC, beta, F_ET, Iz As Double
        Dim hw, tw, lT, b, tf As Double
        Dim Af, Aw, ef As Double
        Dim Zp, Zt, hs As Double
        ' sd stands for design stress Or load
        Dim sigma_Xsd, sigma_Ysd, tau_sd As Double
        Dim lambdaT, mu, fT As Double

        sigma_Xsd = NumericUpDown43.Value
        sigma_Ysd = NumericUpDown42.Value
        tau_sd = NumericUpDown44.Value

        b = NumericUpDown33.Value           'Flange width
        tf = NumericUpDown40.Value          'Flange thickness
        hw = NumericUpDown35.Value          'stiffeners height
        tw = NumericUpDown40.Value          'stiffeners flange thickness

        Double.TryParse(TextBox114.Text, G) 'Shear modulus
        Double.TryParse(TextBox102.Text, Zt)
        Double.TryParse(TextBox103.Text, Zp)
        Double.TryParse(TextBox105.Text, hs)
        Double.TryParse(TextBox106.Text, ef)

        lT = _l         'torsional buckling _l

        Sigma_JSd = Math.Sqrt(sigma_Xsd ^ 2 + sigma_Ysd ^ 2 - sigma_Xsd * sigma_Ysd + 3 * tau_sd ^ 2) 'equation 7.38

        F_Ept = 5.0 * E_mod * (_t / _s) ^ 2       'equation 7.44
        F_Epy = 0.9 * E_mod * (_t / _s) ^ 2       'equation 7.43
        F_Epx = 3.62 * E_mod * (_t / _s) ^ 2      'equation 7.42
        c = 2 - (_s / _l)                    'equation 7.41

        lambda_e = (tau_sd / F_Ept) ^ c         'equation 7.40
        lambda_e += (sigma_Ysd / F_Epy) ^ c
        lambda_e += (sigma_Xsd / F_Epx) ^ c
        lambda_e = lambda_e ^ (1 / c)
        lambda_e = fy / Sigma_JSd
        lambda_e *= fy / Sigma_JSd
        lambda_e = lambda_e ^ -0.5

        F_ep = fy / Sqrt(1 + lambda_e ^ -4)     'equation 7.39

        eta = Sigma_JSd / F_ep                  'equation 7.37
        If eta > 1 Then eta = 1

        CC = hw / _s * (_t / tw) ^ 3 * Sqrt(1 - eta)  'equation 7.36

        beta = (3 * CC + 0.2) / (CC + 0.2)          'equation 7.35

        '--------- For NOW just flat bar-----------------
        F_ET = beta + 2 * (hw / lT) ^ 2             'equation 7.34 (Flat bar)
        F_ET *= G * (tw / hw) ^ 2

        lambdaT = Sqrt(fy / F_ET)                   'equation 7.30

        mu = 0.35 * (lambdaT - 0.6)                 'equation 7.29
        '---------------------------------------------------

        If lambdaT <= 0.6 Then
            fT = 1.0 * fy                           'equation 7.27
        Else
            fT = 1 + mu + lambdaT ^ 2               'equation 7.28
            fT -= Sqrt((1 + mu + lambdaT ^ 2) ^ 2 - 4 * lambdaT ^ 2)
            fT /= 2 * lambdaT ^ 2
            fT *= fy
        End If

        Af = hw * tw + b * tf                       'Cross section area flange
        Aw = b * tf                                 'Cross section area web

        'Iz = 1 / 12 * Af * beta ^ 2                 'equation 7.33
        'Iz += ef ^ 2 * Af / (1 + Af / Aw)

        'F_ET = PI ^ 2 * E_mod * Iz / ((Aw / 3 + Af) * lT ^ 2)    'equation 7.32
        'F_ET += beta * (Aw + (tf / tw) ^ 2 * Af * G * (tw / hw ^ 2) / (Aw + 3 * Af))

        TextBox89.Text = Math.Round(_s, 1).ToString
        TextBox90.Text = Math.Round(_t, 1).ToString
        TextBox96.Text = Math.Round(_l, 1).ToString
        TextBox80.Text = Math.Round(F_Ept, 1).ToString
        TextBox81.Text = Math.Round(F_Epy, 1).ToString
        TextBox82.Text = Math.Round(F_Epx, 1).ToString
        TextBox83.Text = Math.Round(c, 1).ToString
        TextBox84.Text = Math.Round(lambda_e, 4).ToString
        TextBox91.Text = Math.Round(F_ep, 1).ToString
        TextBox92.Text = Math.Round(Sigma_JSd, 1).ToString
        TextBox93.Text = Math.Round(eta, 1).ToString
        TextBox94.Text = Math.Round(CC, 3).ToString

        TextBox95.Text = Math.Round(beta, 1).ToString
        TextBox97.Text = Math.Round(F_ET, 1).ToString
        TextBox98.Text = Math.Round(Iz, 1).ToString
        TextBox104.Text = Math.Round(lT, 1).ToString

        TextBox108.Text = Math.Round(psd, 4).ToString
        TextBox109.Text = Math.Round(sigma_Xsd, 1).ToString
        TextBox110.Text = Math.Round(sigma_Ysd, 1).ToString
        TextBox111.Text = Math.Round(tau_sd, 1).ToString
        TextBox112.Text = Math.Round(lambdaT, 1).ToString
        TextBox113.Text = Math.Round(mu, 1).ToString
        TextBox99.Text = Math.Round(fT, 1).ToString
    End Sub
    Private Sub DNV_chapter7_6()

        Dim sigma_Xsd, sigma_Ysd, tau_sd As Double
        Dim tau_RdY, tau_Rdl, tau_Rds As Double
        Dim tau_crl, tau_crs As Double
        Dim Ip, Iss As Double

        sigma_Xsd = NumericUpDown43.Value
        sigma_Ysd = NumericUpDown42.Value
        tau_sd = NumericUpDown44.Value

        '------------ start calculation ----------
        Ip = _t ^ 3 * _s / 10.9                     'eq 7.49

        Double.TryParse(TextBox33.Text, Iss)
        Double.TryParse(TextBox35.Text, tau_crl)    'eq 7.6
        tau_crs = 36 * E_mod
        tau_crs /= (_s * _t * _l ^ 2)
        tau_crs *= (Ip * Iss ^ 3) ^ 0.25            'eq 7.48

        tau_RdY = fy / (Sqrt(3) * Ym)               'eq 7.45
        tau_Rdl = tau_crl / Ym                      'eq 7.46
        tau_Rds = tau_crs / Ym                      'eq 7.47

        '------------------ present----------------
        TextBox140.Text = Math.Round(_s, 1).ToString
        TextBox141.Text = Math.Round(_t, 1).ToString
        TextBox131.Text = Math.Round(_l, 1).ToString

        TextBox149.Text = Math.Round(psd, 4).ToString
        TextBox148.Text = Math.Round(sigma_Xsd, 1).ToString
        TextBox147.Text = Math.Round(sigma_Ysd, 1).ToString
        TextBox146.Text = Math.Round(tau_sd, 1).ToString

        TextBox128.Text = Math.Round(Ip, 0).ToString
        TextBox129.Text = Math.Round(Iss, 0).ToString

        TextBox120.Text = Math.Round(tau_crs, 0).ToString
        TextBox142.Text = Math.Round(tau_crl, 0).ToString
        TextBox143.Text = Math.Round(tau_Rds, 0).ToString
        TextBox144.Text = Math.Round(tau_Rdl, 0).ToString
        TextBox136.Text = Math.Round(tau_RdY, 0).ToString
    End Sub
    Private Sub DNV_chapter7_71()
        Dim u As Double
        Dim tau_sd, tau_Rd As Double
        Dim e750, e751, e752, e753 As Double
        Dim e754, e755, e756, e757 As Double

        tau_sd = 999
        tau_Rd = 888

        u = tau_sd / tau_Rd                   'equation 7.58

        e750 = 1
        e751 = 2
        e752 = 3
        e753 = 4

        e754 = 5
        e755 = 6
        e756 = 7
        e757 = 8

        TextBox134.Text = Math.Round(e750, 0).ToString
        TextBox135.Text = Math.Round(e751, 0).ToString
        TextBox137.Text = Math.Round(e752, 0).ToString
        TextBox138.Text = Math.Round(e753, 0).ToString
        TextBox158.Text = Math.Round(e754, 0).ToString
        TextBox159.Text = Math.Round(e755, 0).ToString
        TextBox160.Text = Math.Round(e756, 0).ToString
        TextBox161.Text = Math.Round(e757, 0).ToString
    End Sub
    Private Sub DNV_chapter7_73() 'Chapter 7.7.3
        Dim Nrd, Ac, A_s, Sc As Double

        Nrd = 99
        Ac = 99
        A_s = 99
        Sc = 99

        TextBox168.Text = Math.Round(Nrd, 0).ToString
        TextBox170.Text = Math.Round(Ac, 0).ToString
        TextBox172.Text = Math.Round(A_s, 0).ToString
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Calc_sequence()

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
      vbTab & "Stainless steel 316L" & vbTab & vbTab & "@ 20c" & vbTab & "195 [N/mm2]" & vbCrLf &
       vbTab & "Stainless steel 316L" & vbTab & vbTab & "@ 200c" & vbTab & "152 [N/mm2]"
        DNV_chapter6_3()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles NumericUpDown7.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown10.ValueChanged, Button4.Click
        Get_material_data()
    End Sub
    Private Sub Get_material_data()
        Dim pv, G As Double

        fy = NumericUpDown19.Value          'Yield stress
        E_mod = NumericUpDown7.Value        'Elasticity [N/mm2]
        Ym = NumericUpDown10.Value          'Safety
        pv = NumericUpDown17.Value          'Poissons ration

        G = E_mod / (2 * (1 + pv))          'Shear modulus
        psd = NumericUpDown45.Value / 1000 ^ 2  '[N/mm2]

        _s = NumericUpDown39.Value           'Stiffeners distance
        _t = NumericUpDown36.Value           'Plate thickness
        _l = NumericUpDown47.Value      'stiffeners length

        TextBox114.Text = Math.Round(G, 0).ToString
    End Sub
    Private Sub Stiffener_data()
        Dim Zp, Zt, hs, ef As Double
        Dim bf, tf As Double    'Dimensions flange
        Dim hw, tw As Double    'Dimensions web
        Dim Ie, Ie1, Ie2, Ae, ie_small_char As Double
        Dim A1, A2, dY As Double
        bf = NumericUpDown33.Value           'Flange width
        tf = NumericUpDown37.Value          'Flange thickness
        hw = NumericUpDown35.Value          'web height
        tw = NumericUpDown40.Value          'web thickness

        '----------  Centroid calculation ------------------
        A1 = bf * tf                     'Cross section area flange
        A2 = hw * tw                    'Cross section area web

        dY = hw / 2 + tw / 2
        dY = (dY * A1) / (A1 + A2)
        TextBox118.Text = Math.Round(dY, 1).ToString

        Zt = hw / 2 + dY
        Zp = hw / 2 - dY

        hs = hw / 2           '??????????
        ef = bf / 2

        Ae = A1 + A2        'Effective area

        '---------------- Moment of inertia ----------------
        Ie1 = tw * hw ^ 3 / 3   'Web part

        Ie2 = bf * tf / 12      'Flange part
        Ie2 += A1 * hw ^ 2      'Flange part verschuiving naar buiten

        Ie = Ie1 + Ie2

        ie_small_char = Sqrt(Ie / Ae) 'Effective radius of gyration

        TextBox100.Text = Math.Round(A1, 1).ToString
        TextBox101.Text = Math.Round(A2, 1).ToString
        TextBox102.Text = Math.Round(Zt, 1).ToString
        TextBox103.Text = Math.Round(Zp, 1).ToString
        TextBox105.Text = Math.Round(hs, 1).ToString
        TextBox106.Text = Math.Round(ef, 1).ToString
        TextBox87.Text = Math.Round(ie_small_char, 1).ToString
        TextBox115.Text = Math.Round(Ie, 0).ToString
        TextBox116.Text = Math.Round(Ae, 0).ToString
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged, RadioButton1.CheckedChanged, NumericUpDown20.VisibleChanged, NumericUpDown20.ValueChanged, NumericUpDown16.ValueChanged, GroupBox13.VisibleChanged, Button5.Click, NumericUpDown28.ValueChanged, NumericUpDown25.ValueChanged
        Calc_sequence()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabPage9.Enter
        Calc_sequence()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click, TabPage4.Enter, RadioButton3.CheckedChanged
        Calc_sequence()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, TabPage3.Enter
        Calc_sequence()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click, TabPage11.Enter
        Calc_sequence()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click, TabPage12.Enter, NumericUpDown40.ValueChanged, NumericUpDown37.ValueChanged, NumericUpDown35.ValueChanged, NumericUpDown33.ValueChanged
        Calc_sequence()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click, TabPage13.Enter
        Calc_sequence()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click, TabPage14.Enter
        Calc_sequence()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click, TabPage15.Click
        Calc_sequence()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles NumericUpDown47.ValueChanged, NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown43.ValueChanged, NumericUpDown42.ValueChanged, NumericUpDown39.ValueChanged, NumericUpDown36.ValueChanged, Button6.Click
        Calc_sequence()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter, Button2.Click
        Calc_sequence()
    End Sub


    Private Sub Calc_sequence()
        Get_material_data()
        Calc_weight_and_loads()
        Stiffener_data()
        DNV_chapter5_0()
        DNV_chapter7_52()
        DNV_chapter6_2()    'Unstiffened plate longitudinal uniform loading
        DNV_chapter6_3()    'Unstiffened plate transverse uniform loading
        DNV_chapter6_6() 'Chapter 6.6
        DNV_chapter6_7() 'Chapter 6.7
        DNV_chapter7_2() 'Chapter 7.2
        DNV_chapter7_3() 'Chapter 7.3
        DNV_chapter7_4() 'Chapter 7.4
        DNV_chapter7_6() 'Chapter 7.6
        DNV_chapter7_51() 'Chapter 7.5.1
        DNV_chapter7_52() 'Chapter 7.5.2
        DNV_chapter7_71() 'Chapter 7.7.1
        DNV_chapter7_73() 'Chapter 7.7.3
    End Sub
End Class

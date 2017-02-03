Public Class Form1


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown43.ValueChanged, NumericUpDown42.ValueChanged, NumericUpDown41.ValueChanged, NumericUpDown40.ValueChanged, NumericUpDown39.ValueChanged, NumericUpDown36.ValueChanged, NumericUpDown47.ValueChanged, NumericUpDown46.ValueChanged
        'Çhapter 5, DNV-RP-C201, Oktober 2010 
        Dim fy As Double
        Dim t As Double
        Dim s As Double
        Dim length As Double
        Dim upsilon_m As Double
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

        fy = NumericUpDown40.Value          'Yield stress
        s = NumericUpDown39.Value           'Stiffeners distance
        t = NumericUpDown36.Value           'Plate thickness
        length = NumericUpDown47.Value      'stiffeners length
        upsilon_m = NumericUpDown41.Value   'Safety factor

        Psd = NumericUpDown45.Value / 1000 ^ 2  '[N/mm2]
        sigma_Xsd = NumericUpDown43.Value
        sigma_Ysd = NumericUpDown42.Value
        tau_sd = NumericUpDown44.Value

        'Calc Von misses (formula 5.4)
        sigma_Jsd = Math.Sqrt(sigma_Xsd ^ 2 + sigma_Ysd ^ 2 - sigma_Xsd * sigma_Ysd + 3 * tau_sd ^ 2)
        NumericUpDown46.Value = sigma_Jsd

        'Calc constant (formula 5.2)
        upsilon_Y = 1 - (sigma_Jsd / fy) ^ 2
        upsilon_Y = upsilon_Y / Math.Sqrt(1 - 3 / 4 * (sigma_Xsd / fy) ^ 2) - 3 * (tau_sd / fy) ^ 2

        'Calc constant (formula 5.3)
        upsilon_X = 1 - (sigma_Jsd / fy) ^ 2
        upsilon_X /= Math.Sqrt(1 - 3 / 4 * (sigma_Ysd / fy) ^ 2) - 3 * (tau_sd / fy) ^ 2


        'Psd Check (formula 5.1)
        Psd_check = 4 * fy / upsilon_m * ((t / s) ^ 2) * (upsilon_Y + ((s / length) ^ 2 * upsilon_X))

        TextBox1.Text = Math.Round(upsilon_X, 6).ToString
        TextBox2.Text = Math.Round(upsilon_Y, 6).ToString
        TextBox4.Text = Math.Round(Psd_check, 4).ToString

        'Checks
        NumericUpDown39.BackColor = IIf(s < length, Color.Yellow, Color.LightCoral)
        NumericUpDown47.BackColor = IIf(s < length, Color.Yellow, Color.LightCoral)
        TextBox4.BackColor = IIf(Psd < Psd_check, Color.LightGreen, Color.LightCoral)

        epsilon = Math.Sqrt(235 / fy)
        Label18.Text = IIf(s / t < 5.4 * epsilon, "Buckling check NOT neccessary", "Buckling check is neccessary")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, TabPage1.Enter
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage2.Enter
        'Chapter 6.2
        Dim fy, t, s, ym, psd As Double
        Dim E_mod As Double
        Dim lambda_P, Cx, sigma_Xrd As Double

        fy = NumericUpDown40.Value          'Yield stress
        s = NumericUpDown39.Value           'Stiffeners distance
        t = NumericUpDown36.Value           'Plate thickness
        E_mod = NumericUpDown3.Value        'Elasticity [N/mm2]
        ym = NumericUpDown41.Value          'Safety
        Psd = NumericUpDown45.Value / 1000 ^ 2  '[N/mm2]

        '============ Longitudinal uniform compression===========
        lambda_P = 0.525 * s / t * Math.Sqrt(fy / E_mod)    'formule 6.3

        If lambda_P < 0.673 Then            'formule 6.2
            Cx = 1
        Else
            Cx = (lambda_P - 0.22) / lambda_P ^ 2
        End If

        sigma_Xrd = Cx * fy / ym            'formule 6.1

        TextBox3.Text = Math.Round(lambda_P, 2).ToString
        TextBox15.Text = Math.Round(Cx, 2).ToString
        TextBox16.Text = Math.Round(sigma_Xrd, 0).ToString

        '============ Transverserse uniform compression===========
        Dim lambda_C, kappa, mu As Double
        Dim Kp, h_alfa As Double

        '============= h_alfa ============= Formule 6.11
        h_alfa = 0.05 * s / t - 0.75

        '========== kp  ============ Formule 6.10
        If psd <= 2 * (t / s) ^ 2 * fy Then
            Kp = 1
        Else
            Kp = 1 - h_alfa * (psd / fy - 2 * (t / s) ^ 2)
            If Kp < 0 Then Kp = 0
        End If

        '======== lambda_C ========= Formule 6.8
        lambda_C = 1.1 * s / t * Math.Sqrt(fy / E_mod)

        '=========== kappa======= Formule 6.7
        If lambda_C <= 0.2 Then kappa = 1
        If lambda_C > 0.2 And lambda_C < 0.2 Then
            kappa = 1 / (2 * lambda_C ^ 2)
            kappa *= 1 + mu + lambda_C ^ 2 - Math.Sqrt((1 + mu + lambda_C ^ 2) ^ 2 - 4 * lambda_C ^ 2)
        End If
        If lambda_C >= 2 Then kappa = 0.07 + 1 / (2 * lambda_C ^ 2)

        '========== mu========== Formule 6.9
        mu = 0.21 * (lambda_C - 0.2)

    End Sub


End Class

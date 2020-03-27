Imports System.Math
Imports Microsoft.Office.Interop.Excel
Imports 船体梁极限强度计算程序.模块_公有过程_单元应力计算

Public Class 主窗体
    Private Sub 主窗体加载(sender As Object, e As EventArgs) Handles Me.Load

        DataGridView4.Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")

        Dim 输入提示 As String =
            "一: 模式船型代号(OT: ISSC Oil Tanker (Double Hull VLCC); BC: ISSC Bulk Carrier)" & vbCrLf &
            "二: 完整性代号(0: 完整; 1: 碰撞; 2: 非对称搁浅; 3: 对称搁浅)" & vbCrLf &
            "三: 计算方法代号(IIM: 增量迭代法; PIM: 纯增量法)" & vbCrLf &
            "四: 确定性/随机性方法代号(0: 确定性方法; 1: 随机性方法)" & vbCrLf &
            "五: 正/负向曲率代号(P: 正向; N: 负向)" & "水线倾角代号(000; 15P; 30P; ...; 75P; 090; 15N; 30N; ... ; 75N)" & "附加代号(1: IIM-ζ; 2: IIM-γ)"

        参数化输入 = InputBox(输入提示, "参数化输入(各部分以下划线连接)", "OT_0_IIM_0_P0001")

        Select Case 参数化输入
            Case "OT_0_IIM_0_P0001"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    .Rows.Add("α/(°)", "0.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N0001"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    .Rows.Add("α/(°)", "0.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P15P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "5.3", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P15N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-5.3", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N15P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "5.3", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N15N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-5.3", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P30P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "11.3", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P30P2"
                With DataGridView4
                    '.Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "11.3", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P30N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-11.3", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P30N2"
                With DataGridView4
                    '.Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-11.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N30P2"
                With DataGridView4
                    '.Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "11.3", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N30P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "11.3", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N30N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-11.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N30N2"
                With DataGridView4
                    '.Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-11.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P45P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "19.1", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P45P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-37.3", "0.01", "300")
                    .Rows.Add("α/(°)", "19.1", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P45N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.1", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P45N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "37.3", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.1", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N45P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "19.1", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N45P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-37.3", "0.01", "300")
                    .Rows.Add("α/(°)", "19.1", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N45N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.1", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N45N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "37.3", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.1", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P60P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    '.Rows.Add("γ/(m)", "-22.0", "0.01", "200")
                    .Rows.Add("α/(°)", "31.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P60P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-21.4", "0.01", "200")
                    .Rows.Add("α/(°)", "31.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P60N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    '.Rows.Add("γ/(m)", "22.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-31.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P60N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "21.4", "0.01", "200")
                    .Rows.Add("α/(°)", "-31.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N60P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    '.Rows.Add("γ/(m)", "-22.0", "0.01", "200")
                    .Rows.Add("α/(°)", "31.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N60P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-21.4", "0.01", "200")
                    .Rows.Add("α/(°)", "31.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N60N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    '.Rows.Add("γ/(m)", "22.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-31.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N60N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "21.4", "0.01", "200")
                    .Rows.Add("α/(°)", "-31.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P75P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-10.0", "0.01", "200")
                    .Rows.Add("α/(°)", "52.2", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P75N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "10.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-52.2", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N75P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-10.0", "0.01", "200")
                    .Rows.Add("α/(°)", "52.2", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N75N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "10.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-52.2", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P0902"
                With DataGridView4
                    .Rows.Add("γ/(m)", "0.0", "0.01", "200")
                    .Rows.Add("α/(°)", "90.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N0902"
                With DataGridView4
                    .Rows.Add("γ/(m)", "0.0", "0.01", "200")
                    .Rows.Add("α/(°)", "90.0", "0.1", "100")
                End With

            Case "OT_0_IIM_0_P80P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-6.6", "0.01", "200")
                    .Rows.Add("α/(°)", "63.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N80P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-6.6", "0.01", "200")
                    .Rows.Add("α/(°)", "63.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P80N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "6.6", "0.01", "200")
                    .Rows.Add("α/(°)", "-63.0", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N80N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "6.6", "0.01", "200")
                    .Rows.Add("α/(°)", "-63.0", "0.1", "100")
                End With

            Case "OT_0_IIM_0_P85P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-3.3", "0.01", "200")
                    .Rows.Add("α/(°)", "76.2", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N85P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-3.3", "0.01", "200")
                    .Rows.Add("α/(°)", "76.2", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P85N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "3.3", "0.01", "200")
                    .Rows.Add("α/(°)", "-76.2", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N85N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "3.3", "0.01", "200")
                    .Rows.Add("α/(°)", "-76.2", "0.1", "100")
                End With

            Case "OT_0_IIM_0_P70P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-13.7", "0.01", "200")
                    .Rows.Add("α/(°)", "43.5", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N70P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-13.7", "0.01", "200")
                    .Rows.Add("α/(°)", "43.5", "0.1", "100")
                End With
            Case "OT_0_IIM_0_P70N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "13.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-43.5", "0.1", "100")
                End With
            Case "OT_0_IIM_0_N70N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "13.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-43.5", "0.1", "100")
                End With

            Case "OT_1_IIM_0_P0001"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "11.4", "0.01", "200")
                    .Rows.Add("α/(°)", "-5.8", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N0001"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "11.4", "0.01", "200")
                    .Rows.Add("α/(°)", "-5.8", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P15P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.0", "0.01", "200")
                    .Rows.Add("α/(°)", "1.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N15P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.0", "0.01", "200")
                    .Rows.Add("α/(°)", "1.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P15N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "10.8", "0.01", "200")
                    .Rows.Add("α/(°)", "-12.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N15N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "10.8", "0.01", "200")
                    .Rows.Add("α/(°)", "-12.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P30P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "10.3", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P30P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-70.4", "0.01", "500")
                    .Rows.Add("α/(°)", "10.4", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P30N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "10.2", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P30N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "30.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N30P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "12.9", "0.005", "500")
                    .Rows.Add("α/(°)", "10.3", "0.05", "200")
                End With
            Case "OT_1_IIM_0_N30P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-70.4", "0.002", "800")
                    .Rows.Add("α/(°)", "10.4", "0.02", "500")
                End With
            Case "OT_1_IIM_0_N30N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "10.2", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N30N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "30.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P45P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.0", "0.01", "200")
                    .Rows.Add("α/(°)", "22.4", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P45P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-34.0", "0.01", "300")
                    .Rows.Add("α/(°)", "22.4", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P45N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "9.4", "0.01", "200")
                    .Rows.Add("α/(°)", "-27.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P45N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "18.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-27.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N45P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.0", "0.01", "200")
                    .Rows.Add("α/(°)", "22.4", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N45P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-34.0", "0.01", "300")
                    .Rows.Add("α/(°)", "22.4", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N45N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "9.4", "0.01", "200")
                    .Rows.Add("α/(°)", "-27.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N45N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "18.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-27.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P60P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "16.0", "0.01", "200")
                    .Rows.Add("α/(°)", "41.0", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P60P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-19.0", "0.01", "200")
                    .Rows.Add("α/(°)", "41.0", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N60P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "16.0", "0.01", "200")
                    .Rows.Add("α/(°)", "41.0", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N60P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-19.0", "0.01", "300")
                    .Rows.Add("α/(°)", "41.0", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P60N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "8.1", "0.01", "200")
                    .Rows.Add("α/(°)", "-38.0", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P60N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "10.3", "0.01", "200")
                    .Rows.Add("α/(°)", "-38.0", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N60N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "8.1", "0.01", "200")
                    .Rows.Add("α/(°)", "-38.0", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N60N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "10.3", "0.01", "200")
                    .Rows.Add("α/(°)", "-38.0", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P75P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-9.4", "0.01", "200")
                    .Rows.Add("α/(°)", "69.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P75N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "3.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-53.4", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N75P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-9.4", "0.01", "200")
                    .Rows.Add("α/(°)", "69.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N75N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "3.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-53.4", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P0902"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-2.4", "0.01", "200")
                    .Rows.Add("α/(°)", "102.0", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N0902"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-2.4", "0.01", "200")
                    .Rows.Add("α/(°)", "102.0", "0.1", "100")
                End With

            Case "OT_1_IIM_0_P80P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-7.0", "0.01", "200")
                    .Rows.Add("α/(°)", "80.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N80P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-7.0", "0.01", "200")
                    .Rows.Add("α/(°)", "80.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P80N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "1.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-60.4", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N80N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "1.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-60.4", "0.1", "100")
                End With

            Case "OT_1_IIM_0_P85P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-4.5", "0.01", "200")
                    .Rows.Add("α/(°)", "91.7", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N85P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-4.5", "0.01", "200")
                    .Rows.Add("α/(°)", "91.7", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P85N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "0.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-68.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N85N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "0.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-68.5", "0.1", "100")
                End With

            Case "OT_1_IIM_0_P70P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-12.1", "0.01", "200")
                    .Rows.Add("α/(°)", "58.8", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N70P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-12.1", "0.01", "200")
                    .Rows.Add("α/(°)", "58.8", "0.1", "100")
                End With
            Case "OT_1_IIM_0_P70N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "6.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-47.5", "0.1", "100")
                End With
            Case "OT_1_IIM_0_N70N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "6.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-47.5", "0.1", "100")
                End With

            Case "OT_2_IIM_0_P0001"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "1.2", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N0001"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "1.2", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P15P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "5.7", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N15P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "5.7", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P15N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-3.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N15N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-3.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P30P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.9", "0.01", "200")
                    .Rows.Add("α/(°)", "10.8", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N30P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.9", "0.01", "200")
                    .Rows.Add("α/(°)", "10.8", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P30N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-8.7", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N30N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-8.7", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P30P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-70.4", "0.01", "500")
                    .Rows.Add("α/(°)", "10.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P30N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "30.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.5", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N30P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-70.4", "0.002", "800")
                    .Rows.Add("α/(°)", "10.4", "0.02", "500")
                End With
            Case "OT_2_IIM_0_N30N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "30.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-19.5", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P45P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "15.0", "0.01", "200")
                    .Rows.Add("α/(°)", "17.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N45P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "15.0", "0.01", "200")
                    .Rows.Add("α/(°)", "17.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P45N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.5", "0.01", "200")
                    .Rows.Add("α/(°)", "-16.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N45N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "50.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-16.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P45P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-48.0", "0.01", "200")
                    .Rows.Add("α/(°)", "17.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N45P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-48.0", "0.01", "400")
                    .Rows.Add("α/(°)", "17.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P45N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "51.5", "0.01", "200")
                    .Rows.Add("α/(°)", "-15.8", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N45N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "51.5", "0.01", "200")
                    .Rows.Add("α/(°)", "-15.8", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P60P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "16.0", "0.01", "200")
                    .Rows.Add("α/(°)", "41.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N60P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "16.0", "0.01", "200")
                    .Rows.Add("α/(°)", "41.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P60N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "8.1", "0.01", "200")
                    .Rows.Add("α/(°)", "-38.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N60N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "8.1", "0.01", "200")
                    .Rows.Add("α/(°)", "-38.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P60P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-29.0", "0.01", "200")
                    .Rows.Add("α/(°)", "27.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N60P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-29.0", "0.01", "200")
                    .Rows.Add("α/(°)", "27.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P60N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "28.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-27.2", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N60N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "28.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-27.2", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P75P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-14.6", "0.01", "200")
                    .Rows.Add("α/(°)", "46.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N75P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-14.6", "0.01", "200")
                    .Rows.Add("α/(°)", "46.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P75N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "12.0", "0.01", "300")
                    .Rows.Add("α/(°)", "-50.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N75N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "12.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-50.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P0902"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-1.6", "0.01", "200")
                    .Rows.Add("α/(°)", "86.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N0902"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-1.6", "0.01", "200")
                    .Rows.Add("α/(°)", "86.0", "0.1", "100")
                End With

            Case "OT_2_IIM_0_P80P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-10.3", "0.01", "200")
                    .Rows.Add("α/(°)", "57.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N80P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-10.3", "0.01", "200")
                    .Rows.Add("α/(°)", "57.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P80N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "7.2", "0.01", "300")
                    .Rows.Add("α/(°)", "-62.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N80N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "7.2", "0.01", "200")
                    .Rows.Add("α/(°)", "-62.0", "0.1", "100")
                End With

            Case "OT_2_IIM_0_P85P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-6.0", "0.01", "200")
                    .Rows.Add("α/(°)", "70.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N85P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-6.0", "0.01", "200")
                    .Rows.Add("α/(°)", "70.0", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P85N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "2.7", "0.01", "300")
                    .Rows.Add("α/(°)", "-77.2", "0.1", "100")
                End With
            Case "OT_2_IIM_0_N85N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "2.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-77.2", "0.1", "100")
                End With

            Case "OT_2_IIM_0_P70P2", "OT_2_IIM_0_N70P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-19.2", "0.01", "200")
                    .Rows.Add("α/(°)", "38.4", "0.1", "100")
                End With
            Case "OT_2_IIM_0_P70N2", "OT_2_IIM_0_N70N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "16.9", "0.01", "200")
                    .Rows.Add("α/(°)", "-40.1", "0.1", "100")
                End With

            Case "OT_3_IIM_0_P0001"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "0.0", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N0001"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "0.0", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P15P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "4.6", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N15P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "4.6", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P15N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-4.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N15N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-4.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P30P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "9.7", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N30P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "9.7", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P30N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "-9.7", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N30N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "-9.7", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P30P2"
                With DataGridView4
                    '.Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "11.3", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P30N2"
                With DataGridView4
                    '.Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-11.0", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N30P2"
                With DataGridView4
                    '.Rows.Add("ζ/(m)", "12.9", "0.01", "200")
                    .Rows.Add("α/(°)", "11.3", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N30N2"
                With DataGridView4
                    '.Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-11.0", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P45P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "16.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N45P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.8", "0.01", "200")
                    .Rows.Add("α/(°)", "16.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P45N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-16.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N45N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "14.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-16.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P45P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-50.0", "0.01", "300")
                    .Rows.Add("α/(°)", "16.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N45P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-50.0", "0.01", "400")
                    .Rows.Add("α/(°)", "16.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P45N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "50.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-16.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N45N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "50.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-16.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P60P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    '.Rows.Add("γ/(m)", "-22.0", "0.01", "200")
                    .Rows.Add("α/(°)", "31.0", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P60P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-28.7", "0.01", "200")
                    .Rows.Add("α/(°)", "27.2", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N60P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-28.7", "0.01", "200")
                    .Rows.Add("α/(°)", "27.2", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P60N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    '.Rows.Add("γ/(m)", "22.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-31.0", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P60N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "28.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-27.2", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N60N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "28.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-27.2", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N60P1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    '.Rows.Add("γ/(m)", "-22.0", "0.01", "200")
                    .Rows.Add("α/(°)", "31.0", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N60N1"
                With DataGridView4
                    .Rows.Add("ζ/(m)", "13.0", "0.01", "200")
                    '.Rows.Add("γ/(m)", "22.0", "0.01", "200")
                    .Rows.Add("α/(°)", "-31.0", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P75P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-13.3", "0.01", "200")
                    .Rows.Add("α/(°)", "47.9", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N75P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-13.3", "0.01", "200")
                    .Rows.Add("α/(°)", "47.9", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P75N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "13.3", "0.01", "300")
                    .Rows.Add("α/(°)", "-47.9", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N75N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "13.3", "0.01", "200")
                    .Rows.Add("α/(°)", "-47.9", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P0902"
                With DataGridView4
                    .Rows.Add("γ/(m)", "0.0", "0.01", "200")
                    .Rows.Add("α/(°)", "90.0", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N0902"
                With DataGridView4
                    .Rows.Add("γ/(m)", "0.0", "0.01", "300")
                    .Rows.Add("α/(°)", "90.0", "0.1", "100")
                End With

            Case "OT_3_IIM_0_P80P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-8.7", "0.01", "200")
                    .Rows.Add("α/(°)", "59.1", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N80P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-8.7", "0.01", "200")
                    .Rows.Add("α/(°)", "59.1", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P80N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "8.7", "0.01", "300")
                    .Rows.Add("α/(°)", "-59.1", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N80N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "8.7", "0.01", "200")
                    .Rows.Add("α/(°)", "-59.1", "0.1", "100")
                End With

            Case "OT_3_IIM_0_P85P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-4.4", "0.01", "200")
                    .Rows.Add("α/(°)", "73.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N85P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-4.4", "0.01", "300")
                    .Rows.Add("α/(°)", "73.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P85N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "4.4", "0.01", "300")
                    .Rows.Add("α/(°)", "-73.5", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N85N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "4.4", "0.01", "200")
                    .Rows.Add("α/(°)", "-73.5", "0.1", "100")
                End With

            Case "OT_3_IIM_0_P70P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-18.1", "0.01", "200")
                    .Rows.Add("α/(°)", "39.1", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N70P2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "-18.1", "0.01", "200")
                    .Rows.Add("α/(°)", "39.1", "0.1", "100")
                End With
            Case "OT_3_IIM_0_P70N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "18.1", "0.01", "200")
                    .Rows.Add("α/(°)", "-39.1", "0.1", "100")
                End With
            Case "OT_3_IIM_0_N70N2"
                With DataGridView4
                    .Rows.Add("γ/(m)", "18.1", "0.01", "200")
                    .Rows.Add("α/(°)", "-39.1", "0.1", "100")
                End With

                'Case "BC"
                '    With DataGridView4
                '        .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
                '        .Rows.Add("ζ/(m)", "", "0.01", "200")
                '        .Rows.Add("α/(°)", "", "0.1", "100")
                '    End With
            Case Else
                With DataGridView4
                    .Rows.Add("ζ or γ/(m)", "", "0.01", "200")
                    .Rows.Add("α/(°)", "", "0.1", "100")
                End With
        End Select
        DataGridView4.AllowUserToAddRows = False
        DataGridView4.AllowUserToDeleteRows = False
        DataGridView4.AllowUserToOrderColumns = False

        第一部分 = Strings.Left(参数化输入, 2)
        参数化输入 = Mid(参数化输入, 4)
        第二部分 = Strings.Left(参数化输入, 1)
        参数化输入 = Mid(参数化输入, 3)
        第三部分 = Strings.Left(参数化输入, 3)
        参数化输入 = Mid(参数化输入, 5)
        第四部分 = Strings.Left(参数化输入, 1)
        参数化输入 = Mid(参数化输入, 3)
        第五部分 = Strings.Left(参数化输入, 5)

        Select Case 第一部分
            Case "OT"
                跨度_文本框.Text = 4950
                型宽_文本框.Text = 58000
                型深_文本框.Text = 32000
                节点总数_文本框.Text = 422
                加强筋总数_文本框.Text = 392
                面板总数_文本框.Text = 0
                板格总数_文本框.Text = 432
                特别硬角单元总数_文本框.Text = 4
            Case "BC"
                跨度_文本框.Text = 5220
                型宽_文本框.Text = 50000
                型深_文本框.Text = 27000
                节点总数_文本框.Text = 228
                加强筋总数_文本框.Text = 188
                面板总数_文本框.Text = 0
                板格总数_文本框.Text = 239
                特别硬角单元总数_文本框.Text = 4
        End Select

        Select Case Mid(第五部分, 4, 1)
            Case "P"
                水线倾角_α = Val(Mid(第五部分, 2, 2)) * PI / 180
            Case "N"
                水线倾角_α = -Val(Mid(第五部分, 2, 2)) * PI / 180
            Case "0"
                水线倾角_α = Val(Mid(第五部分, 2, 3)) * PI / 180
        End Select

        Refresh()
    End Sub

    Private Sub 生成表格(sender As Object, e As EventArgs) Handles Button1.Click
        Dim xlApp As Application
        Dim xlBook As Workbook
        Dim xlSheet As Worksheet

        xlApp = CType(CreateObject("Excel.Application"), Application)
        xlBook = xlApp.Workbooks.Add
        xlSheet = CType(xlBook.Worksheets(1), Worksheet)

        '调用 共有过程_通用型.读入界面参数
        读入界面参数()

        Dim 行号 As UShort = 0
        For 基本输入对象类型序数 As UShort = 1 To 5
            Select Case 基本输入对象类型序数
                Case 1
                    行号 += 1
                    xlSheet.Cells(行号, 1) = "节点"

                    行号 += 1
                    xlSheet.Cells(行号, 1) = "序数"
                    xlSheet.Cells(行号, 2) = "Y/(mm)"
                    xlSheet.Cells(行号, 3) = "Z/(mm)"
                    For 节点_序数 = 1 To 节点_总数
                        行号 += 1
                        xlSheet.Cells(行号, 1) = 节点_序数
                    Next
                Case 2
                    行号 += 1
                    xlSheet.Cells(行号, 1) = "加强筋"

                    行号 += 1
                    xlSheet.Cells(行号, 1) = "序数"
                    xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    xlSheet.Cells(行号, 4) = "l/(mm)"
                    xlSheet.Cells(行号, 5) = "hw/(mm)"
                    xlSheet.Cells(行号, 6) = "tw/(mm)"
                    xlSheet.Cells(行号, 7) = "αw/(rad)"
                    xlSheet.Cells(行号, 8) = "wf/(mm)"
                    xlSheet.Cells(行号, 9) = "tf/(mm)"
                    xlSheet.Cells(行号, 10) = "αf/(rad)"
                    xlSheet.Cells(行号, 11) = "dx/(mm)"
                    xlSheet.Cells(行号, 12) = "σY/(MPa)"
                    xlSheet.Cells(行号, 13) = "tp(F/B/T/L1/L2/L3)"
                    xlSheet.Cells(行号, 14) = "mk(TRUE/FALSE)"
                    For 加强筋_序数 = 1 To 加强筋_总数
                        行号 += 1
                        xlSheet.Cells(行号, 1) = 加强筋_序数
                    Next
                Case 3
                    行号 += 1
                    xlSheet.Cells(行号, 1) = "面板"

                    行号 += 1
                    xlSheet.Cells(行号, 1) = "序数"
                    xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    xlSheet.Cells(行号, 4) = "Y1/(mm)"
                    xlSheet.Cells(行号, 5) = "Z1/(mm)"
                    xlSheet.Cells(行号, 6) = "YL/(mm)"
                    xlSheet.Cells(行号, 7) = "ZL/(mm)"
                    xlSheet.Cells(行号, 8) = "l/(mm)"
                    xlSheet.Cells(行号, 9) = "t/(mm)"
                    xlSheet.Cells(行号, 10) = "σY/(MPa)"
                    xlSheet.Cells(行号, 11) = "PMA(TRUE/FALSE)"
                    For 面板_序数 = 1 To 面板_总数
                        行号 += 1
                        xlSheet.Cells(行号, 1) = 面板_序数
                    Next
                Case 4
                    行号 += 1
                    xlSheet.Cells(行号, 1) = "板格"

                    行号 += 1
                    xlSheet.Cells(行号, 1) = "序数"
                    xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    xlSheet.Cells(行号, 4) = "Y1/(mm)"
                    xlSheet.Cells(行号, 5) = "Z1/(mm)"
                    xlSheet.Cells(行号, 6) = "l(mm)"
                    xlSheet.Cells(行号, 7) = "tp(L/T)"
                    xlSheet.Cells(行号, 8) = "板格板数目"
                    xlSheet.Cells(行号, 9) = "w/(mm)"
                    xlSheet.Cells(行号, 10) = "t/(mm)"
                    xlSheet.Cells(行号, 11) = "σY/(MPa)"
                    xlSheet.Cells(行号, 12) = "w/(mm)"
                    xlSheet.Cells(行号, 13) = "t/(mm)"
                    xlSheet.Cells(行号, 14) = "σY/(MPa)"
                    xlSheet.Cells(行号, 15) = "......"
                    xlSheet.Cells(行号, 16) = "......"
                    xlSheet.Cells(行号, 17) = "......"
                    For 板格_序数 = 1 To 板格_总数
                        行号 += 1
                        xlSheet.Cells(行号, 1) = 板格_序数
                    Next
                Case 5
                    行号 += 1
                    xlSheet.Cells(行号, 1) = "特别硬角单元"

                    行号 += 1
                    xlSheet.Cells(行号, 1) = "序数"
                    xlSheet.Cells(行号, 2) = "子对象数目"
                    xlSheet.Cells(行号, 3) = "Yc/(mm)"
                    xlSheet.Cells(行号, 4) = "Zc/(mm)"
                    xlSheet.Cells(行号, 5) = "A/(mm^2)"
                    xlSheet.Cells(行号, 6) = "σY/(MPa)"
                    xlSheet.Cells(行号, 7) = "Yc/(mm)"
                    xlSheet.Cells(行号, 8) = "Zc/(mm)"
                    xlSheet.Cells(行号, 9) = "A/(mm^2)"
                    xlSheet.Cells(行号, 10) = "σY/(MPa)"
                    xlSheet.Cells(行号, 11) = "......"
                    xlSheet.Cells(行号, 12) = "......"
                    xlSheet.Cells(行号, 13) = "......"
                    xlSheet.Cells(行号, 14) = "......"
                    For 特别硬角单元_序数 = 1 To 特别硬角单元_总数
                        行号 += 1
                        xlSheet.Cells(行号, 1) = 特别硬角单元_序数
                    Next
            End Select
        Next
        'xlSheet.Cells(2, 2) = "Hello World!"
        SaveFileDialog1.ShowDialog()
        xlSheet.Application.Visible = True
        xlSheet.SaveAs(保存文件对话框.FileName)
        MsgBox("生成Excel输入模板的操作已完成...", vbOKOnly, "通知")
        'xlSheet.SaveAs("C:\Users\Yan Liyu\Documents\Test.xlsx")
    End Sub

    Private Sub 读入数据(sender As Object, e As EventArgs) Handles Button2.Click
        Dim xlApp As Application
        Dim xlBook As Workbook
        Dim xlSheet As Worksheet

        OpenFileDialog1.ShowDialog()

        xlApp = CType(CreateObject("Excel.Application"), Application)
        xlBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName,, True)
        xlSheet = CType(xlBook.Worksheets(1), Worksheet)

        '调用 共有过程_通用型.读入界面参数
        读入界面参数()

        Dim 行号 As UShort
        For 基本输入对象类型序数 As UShort = 1 To 5
            Select Case 基本输入对象类型序数
                Case 1
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "节点"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z/(mm)"
                    For 一级序数 As UShort = 1 To 节点_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        节点_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        节点_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 1 / 4
                Case 2
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "加强筋"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    'xlSheet.Cells(行号, 4) = "l/(mm)"
                    'xlSheet.Cells(行号, 5) = "hw/(mm)"
                    'xlSheet.Cells(行号, 6) = "tw/(mm)"
                    'xlSheet.Cells(行号, 7) = "αw/(rad)"
                    'xlSheet.Cells(行号, 8) = "wf/(mm)"
                    'xlSheet.Cells(行号, 9) = "tf/(mm)"
                    'xlSheet.Cells(行号, 10) = "αf/(rad)"
                    'xlSheet.Cells(行号, 11) = "dx/(mm)"
                    'xlSheet.Cells(行号, 12) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 13) = "tp(F/B/T/L1/L2/L3)"
                    'xlSheet.Cells(行号, 14) = "mk(TRUE/FALSE)"
                    For 一级序数 As UShort = 1 To 加强筋_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        加强筋_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        加强筋_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                        加强筋_l(一级序数) = xlSheet.Cells(行号, 4).value
                        加强筋_hw(一级序数) = xlSheet.Cells(行号, 5).value
                        加强筋_tw(一级序数) = xlSheet.Cells(行号, 6).value
                        加强筋_αw(一级序数) = xlSheet.Cells(行号, 7).value
                        加强筋_wf(一级序数) = xlSheet.Cells(行号, 8).value
                        加强筋_tf(一级序数) = xlSheet.Cells(行号, 9).value
                        加强筋_αf(一级序数) = xlSheet.Cells(行号, 10).value
                        加强筋_dx(一级序数) = xlSheet.Cells(行号, 11).value
                        加强筋_σY(一级序数) = xlSheet.Cells(行号, 12).value
                        加强筋_tp(一级序数) = xlSheet.Cells(行号, 13).value
                        加强筋_mk(一级序数) = xlSheet.Cells(行号, 14).value
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 2 / 4
                Case 3
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "面板"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    'xlSheet.Cells(行号, 4) = "Y1/(mm)"
                    'xlSheet.Cells(行号, 5) = "Z1/(mm)"
                    'xlSheet.Cells(行号, 6) = "YL/(mm)"
                    'xlSheet.Cells(行号, 7) = "ZL/(mm)"
                    'xlSheet.Cells(行号, 8) = "l/(mm)"
                    'xlSheet.Cells(行号, 9) = "t/(mm)"
                    'xlSheet.Cells(行号, 10) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 11) = "PMA(TRUE/FALSE)"
                    For 一级序数 As UShort = 1 To 面板_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        面板_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        面板_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                        面板_Y1(一级序数) = xlSheet.Cells(行号, 4).value
                        面板_Z1(一级序数) = xlSheet.Cells(行号, 5).value
                        面板_YL(一级序数) = xlSheet.Cells(行号, 6).value
                        面板_ZL(一级序数) = xlSheet.Cells(行号, 7).value
                        面板_l(一级序数) = xlSheet.Cells(行号, 8).value
                        面板_t(一级序数) = xlSheet.Cells(行号, 9).value
                        面板_σY(一级序数) = xlSheet.Cells(行号, 10).value
                        面板_PMA(一级序数) = xlSheet.Cells(行号, 11).value
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 3 / 4
                Case 4
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "板格"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    'xlSheet.Cells(行号, 4) = "Y1/(mm)"
                    'xlSheet.Cells(行号, 5) = "Z1/(mm)"
                    'xlSheet.Cells(行号, 6) = "l(mm)"
                    'xlSheet.Cells(行号, 7) = "tp(L/T)"
                    'xlSheet.Cells(行号, 8) = "板格板数目"
                    'xlSheet.Cells(行号, 9) = "w/(mm)"
                    'xlSheet.Cells(行号, 10) = "t/(mm)"
                    'xlSheet.Cells(行号, 11) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 12) = "w/(mm)"
                    'xlSheet.Cells(行号, 13) = "t/(mm)"
                    'xlSheet.Cells(行号, 14) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 15) = "......"
                    'xlSheet.Cells(行号, 16) = "......"
                    'xlSheet.Cells(行号, 17) = "......"
                    For 一级序数 As UShort = 1 To 板格_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        板格_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        板格_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                        板格_Y1(一级序数) = xlSheet.Cells(行号, 4).value
                        板格_Z1(一级序数) = xlSheet.Cells(行号, 5).value
                        板格_l(一级序数) = xlSheet.Cells(行号, 6).value
                        板格_tp(一级序数) = xlSheet.Cells(行号, 7).value
                        板格_板格板数目(一级序数) = xlSheet.Cells(行号, 8).value
                        For 二级序数 As UShort = 1 To 板格_板格板数目(一级序数)
                            板格板_w(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 9).value
                            板格板_t(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 10).value
                            板格板_σY(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 11).value
                        Next
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 4 / 4
                Case 5
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "特别硬角单元"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "子对象数目"
                    'xlSheet.Cells(行号, 3) = "Yc/(mm)"
                    'xlSheet.Cells(行号, 4) = "Zc/(mm)"
                    'xlSheet.Cells(行号, 5) = "A/(mm^2)"
                    'xlSheet.Cells(行号, 6) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 7) = "Yc/(mm)"
                    'xlSheet.Cells(行号, 8) = "Zc/(mm)"
                    'xlSheet.Cells(行号, 9) = "A/(mm^2)"
                    'xlSheet.Cells(行号, 10) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 11) = "......"
                    'xlSheet.Cells(行号, 12) = "......"
                    'xlSheet.Cells(行号, 13) = "......"
                    'xlSheet.Cells(行号, 14) = "......"
                    For 一级序数 As UShort = 1 To 特别硬角单元_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        子对象_数目(一级序数) = xlSheet.Cells(行号, 2).value
                        For 二级序数 As UShort = 1 To 子对象_数目(一级序数)
                            子对象_Yc(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 3).value
                            子对象_Zc(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 4).value
                            子对象_A(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 5).value
                            子对象_σY(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 6).value
                        Next
                    Next
            End Select
        Next

        xlBook.Close()
        xlApp.Quit() 'xlApp = Nothing

        '基本输入对象的属性计算及形心坐标输出
        For 基本输入对象类型序数 As UShort = 1 To 5
            Select Case 基本输入对象类型序数
                Case 1
                    Chart1.Series.Add("节点_形心")
                    Chart1.Series("节点_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 节点_总数
                        Chart1.Series("节点_形心").Points.AddXY(节点_Y0(一级序数), 节点_Z0(一级序数))
                    Next
                Case 2
                    Chart1.Series.Add("加强筋_形心")
                    Chart1.Series("加强筋_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 加强筋_总数
                        加强筋_Aw(一级序数) = 加强筋_hw(一级序数) * 加强筋_tw(一级序数)
                        加强筋_Icw(一级序数) = 加强筋_tw(一级序数) * 加强筋_hw(一级序数) ^ 3 / 12
                        加强筋_Iow(一级序数) = 加强筋_Icw(一级序数) + 加强筋_Aw(一级序数) * (加强筋_hw(一级序数) / 2) ^ 2

                        Select Case 加强筋_tp(一级序数)
                            Case "F"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数)
                            Case "B"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) - 加强筋_tf(一级序数) / 2
                            Case "T"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                            Case "L1"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                            Case "L2"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                            Case "L3"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) - 加强筋_dx(一级序数) - 加强筋_tf(一级序数) / 2
                            Case Else
                                MsgBox("加强筋_tp(" & 一级序数 & ")错误：类型不符")
                        End Select
                        加强筋_Af(一级序数) = 加强筋_wf(一级序数) * 加强筋_tf(一级序数)
                        加强筋_Icf(一级序数) = 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 12
                        加强筋_Iof(一级序数) = 加强筋_Icf(一级序数) + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2

                        加强筋_AS(一级序数) = 加强筋_Aw(一级序数) + 加强筋_Af(一级序数)
                        加强筋_hcS(一级序数) = 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_hw(一级序数) / 2 + 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_df(一级序数)
                        加强筋_IoS(一级序数) = 加强筋_Iow(一级序数) + 加强筋_Iof(一级序数)
                        加强筋_IcS(一级序数) = 加强筋_Icw(一级序数) + 加强筋_Aw(一级序数) * (加强筋_hw(一级序数) / 2 - 加强筋_hcS(一级序数)) ^ 2 + 加强筋_Icf(一级序数) + 加强筋_Af(一级序数) * (加强筋_df(一级序数) - 加强筋_hcS(一级序数)) ^ 2

                        Select Case 加强筋_tp(一级序数)
                            Case "F"
                                加强筋_IPS(一级序数) = 加强筋_hw(一级序数) ^ 3 * 加强筋_tw(一级序数) / 3
                                加强筋_ITS(一级序数) = 加强筋_hw(一级序数) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / 加强筋_hw(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_hw(一级序数) ^ 3 * 加强筋_tw(一级序数) ^ 3 / 36
                            Case "B"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                            Case "T"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_wf(一级序数) ^ 3 * 加强筋_tf(一级序数) * 加强筋_df(一级序数) ^ 2 / 12
                            Case "L1"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                            Case "L2"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                            Case "L3"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                            Case Else

                        End Select
                        '加强筋_ηS(一级序数) = 1 + 加强筋_l(一级序数) ^ 2 / PI ^ 2 / Sqrt(加强筋_IWS(一级序数) * (0.75 * 加强筋带板_w(一级序数) / 加强筋带板_t(一级序数) ^ 3 + (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) / 加强筋_tw(一级序数) ^ 3))
                        '加强筋_σETS(一级序数) = 标准弹性模量 / 加强筋_IPS(一级序数) * (加强筋_ηS(一级序数) * PI ^ 2 * 加强筋_IWS(一级序数) / 加强筋_lS(一级序数) ^ 2 + 0.385 * 加强筋_ITS(一级序数))
                        加强筋_σELS(一级序数) = 160000 * (加强筋_tw(一级序数) / 加强筋_hw(一级序数)) ^ 2

                        加强筋_Ycw(一级序数) = 加强筋_Y0(一级序数) + 加强筋_hw(一级序数) * Cos(加强筋_αw(一级序数)) / 2
                        加强筋_Zcw(一级序数) = 加强筋_Z0(一级序数) + 加强筋_hw(一级序数) * Sin(加强筋_αw(一级序数)) / 2

                        加强筋_Ycf(一级序数) = 加强筋_Y0(一级序数) + 加强筋_df(一级序数) * Cos(加强筋_αw(一级序数)) + 加强筋_wf(一级序数) * Cos(加强筋_αf(一级序数)) / 2
                        加强筋_Zcf(一级序数) = 加强筋_Z0(一级序数) + 加强筋_df(一级序数) * Sin(加强筋_αw(一级序数)) + 加强筋_wf(一级序数) * Sin(加强筋_αf(一级序数)) / 2

                        加强筋_YcS(一级序数) = 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_Ycf(一级序数) + 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_Ycw(一级序数)
                        加强筋_ZcS(一级序数) = 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_Zcf(一级序数) + 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_Zcw(一级序数)

                        Chart1.Series("加强筋_形心").Points.AddXY(加强筋_YcS(一级序数), 加强筋_ZcS(一级序数))
                    Next
                Case 3
                    Chart1.Series.Add("面板_形心")
                    Chart1.Series("面板_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 面板_总数
                        面板_w(一级序数) = Sqrt((面板_Y0(一级序数) - 面板_Y1(一级序数)) ^ 2 + (面板_Z0(一级序数) - 面板_Z1(一级序数)) ^ 2)
                        面板_A(一级序数) = 面板_w(一级序数) * 面板_t(一级序数)

                        面板_Yc(一级序数) = (面板_Y0(一级序数) + 面板_Y1(一级序数)) / 2
                        面板_Zc(一级序数) = (面板_Z0(一级序数) + 面板_Z1(一级序数)) / 2

                        Chart1.Series("面板_形心").Points.AddXY(面板_Yc(一级序数), 面板_Zc(一级序数))
                        Chart1.Series("面板_形心").Points.AddXY(面板_Y0(一级序数), 面板_Z0(一级序数))
                        Chart1.Series("面板_形心").Points.AddXY(面板_Y1(一级序数), 面板_Z1(一级序数))
                    Next
                Case 4
                    Chart1.Series.Add("板格_形心")
                    Chart1.Series("板格_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 板格_总数
                        板格_α(一级序数) = If(板格_Y0(一级序数) = 板格_Y1(一级序数), PI / 2, Atan((板格_Z0(一级序数) - 板格_Z1(一级序数)) / (板格_Y0(一级序数) - 板格_Y1(一级序数))))
                        板格_w(一级序数) = Sqrt((板格_Y0(一级序数) - 板格_Y1(一级序数)) ^ 2 + (板格_Z0(一级序数) - 板格_Z1(一级序数)) ^ 2)
                        For 二级序数 As UShort = 1 To 板格_板格板数目(一级序数)
                            板格板_Y0(一级序数, 二级序数) = If(二级序数 = 1, 板格_Y0(一级序数), 板格板_Y0(一级序数, 二级序数 - 1))
                            板格板_Z0(一级序数, 二级序数) = If(二级序数 = 1, 板格_Z0(一级序数), 板格板_Z0(一级序数, 二级序数 - 1))

                            板格板_Y1(一级序数, 二级序数) = If(二级序数 = 板格_板格板数目(一级序数), 板格_Y1(一级序数), 板格板_Y0(一级序数, 二级序数) + 板格板_w(一级序数, 二级序数) * Cos(板格_α(一级序数)))
                            板格板_Z1(一级序数, 二级序数) = If(二级序数 = 板格_板格板数目(一级序数), 板格_Z1(一级序数), 板格板_Z0(一级序数, 二级序数) + 板格板_w(一级序数, 二级序数) * Sin(板格_α(一级序数)))

                            板格板_Yc(一级序数, 二级序数) = (板格板_Y0(一级序数, 二级序数) + 板格板_Y1(一级序数, 二级序数)) / 2
                            板格板_Zc(一级序数, 二级序数) = (板格板_Z0(一级序数, 二级序数) + 板格板_Z1(一级序数, 二级序数)) / 2

                            板格板_A(一级序数, 二级序数) = 板格板_w(一级序数, 二级序数) * 板格板_t(一级序数, 二级序数)
                            板格板_AYc(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_Yc(一级序数, 二级序数)
                            板格板_AZc(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_Zc(一级序数, 二级序数)
                            板格板_AσY(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_σY(一级序数, 二级序数)

                            板格_A(一级序数) += 板格板_A(一级序数, 二级序数)
                            板格_AYc(一级序数) += 板格板_AYc(一级序数, 二级序数)
                            板格_AZc(一级序数) += 板格板_AZc(一级序数, 二级序数)
                            板格_AσY(一级序数) += 板格板_AσY(一级序数, 二级序数)
                        Next
                        板格_t(一级序数) = 板格_A(一级序数) / 板格_w(一级序数)
                        板格_Yc(一级序数) = 板格_AYc(一级序数) / 板格_A(一级序数)
                        板格_Zc(一级序数) = 板格_AZc(一级序数) / 板格_A(一级序数)
                        板格_σY(一级序数) = 板格_AσY(一级序数) / 板格_A(一级序数)

                        Chart1.Series("板格_形心").Points.AddXY(板格_Yc(一级序数), 板格_Zc(一级序数))
                    Next
                Case 5
                    Chart1.Series.Add("特别硬角单元_形心")
                    Chart1.Series("特别硬角单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 特别硬角单元_总数
                        For 二级序数 As UShort = 1 To 子对象_数目(一级序数)
                            子对象_AYc(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_Yc(一级序数, 二级序数)
                            子对象_AZc(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_Zc(一级序数, 二级序数)
                            子对象_AσY(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_σY(一级序数, 二级序数)

                            特别硬角单元_A(一级序数) += 子对象_A(一级序数, 二级序数)
                            特别硬角单元_AYc(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_Yc(一级序数, 二级序数)
                            特别硬角单元_AZc(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_Zc(一级序数, 二级序数)
                            特别硬角单元_AσY(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_σY(一级序数, 二级序数)
                        Next
                        特别硬角单元_Yc(一级序数) += 特别硬角单元_AYc(一级序数) / 特别硬角单元_A(一级序数)
                        特别硬角单元_Zc(一级序数) += 特别硬角单元_AZc(一级序数) / 特别硬角单元_A(一级序数)
                        特别硬角单元_σY(一级序数) += 特别硬角单元_AσY(一级序数) / 特别硬角单元_A(一级序数)

                        Chart1.Series("特别硬角单元_形心").Points.AddXY(特别硬角单元_Yc(一级序数), 特别硬角单元_Zc(一级序数))
                    Next
            End Select
        Next
    End Sub

    Private Sub 单元划分(sender As Object, e As EventArgs) Handles Button3.Click
        For 单元划分步骤序数 As UShort = 1 To 4
            Select Case 单元划分步骤序数
                Case 1      '板格-节点配对, 成立节点-分支
                    For 一级序数 As UShort = 1 To 板格_总数
                        For 二级序数 As UShort = 1 To 节点_总数
                            If 板格_Y0(一级序数) = 节点_Y0(二级序数) And 板格_Z0(一级序数) = 节点_Z0(二级序数) Then
                                节点_分支_数目(二级序数) += 1
                                节点_分支_板格_序数(二级序数, 节点_分支_数目(二级序数)) = 一级序数

                                节点_分支_α(二级序数, 节点_分支_数目(二级序数)) = 板格_α(一级序数)

                                板格_首端节点_序数(一级序数) = 二级序数
                                板格_首端节点_分支_序数(一级序数) = 节点_分支_数目(二级序数)
                            ElseIf 板格_Y1(一级序数) = 节点_Y0(二级序数) And 板格_Z1(一级序数) = 节点_Z0(二级序数) Then
                                节点_分支_数目(二级序数) += 1
                                节点_分支_板格_序数(二级序数, 节点_分支_数目(二级序数)) = 一级序数

                                节点_分支_α(二级序数, 节点_分支_数目(二级序数)) = If(板格_α(一级序数) <= 0, 板格_α(一级序数) + PI, 板格_α(一级序数) - PI)

                                板格_末端节点_序数(一级序数) = 二级序数
                                板格_末端节点_分支_序数(一级序数) = 节点_分支_数目(二级序数)
                            End If
                            If (Not 板格_首端节点_序数(一级序数) = 0) And (Not 板格_末端节点_序数(一级序数) = 0) Then
                                '节点_分支_首端节点_序数(二级序数) = 二级序数
                                节点_分支_首端节点_序数(板格_首端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_首端节点_序数(一级序数)
                                节点_分支_末端节点_序数(板格_首端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_末端节点_序数(一级序数)
                                节点_分支_首端节点_序数(板格_末端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_末端节点_序数(一级序数)
                                节点_分支_末端节点_序数(板格_末端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_首端节点_序数(一级序数)
                                '节点_分支_末端节点_序数(二级序数) = 二级序数
                                Exit For
                            End If
                        Next
                    Next
                Case 2      '节点-加强筋/面板配对
                    For 一级序数 As UShort = 1 To 节点_总数
                        For 二级序数 As UShort = 1 To 加强筋_总数
                            If 节点_Y0(一级序数) = 加强筋_Y0(二级序数) And 节点_Z0(一级序数) = 加强筋_Z0(二级序数) Then
                                节点_tp(一级序数) = "加强筋"

                                节点_加强筋_序数(一级序数) = 二级序数
                                加强筋_节点_序数(二级序数) = 一级序数

                                '属性继承：加强筋 → 节点_加强筋
                                Exit For
                            End If
                        Next

                        For 二级序数 As UShort = 1 To 面板_总数
                            If 节点_Y0(一级序数) = 面板_YL(二级序数) And 节点_Z0(一级序数) = 面板_ZL(二级序数) Then
                                节点_tp(一级序数) = "面板"

                                节点_面板_序数(一级序数) = 二级序数
                                面板_节点_序数(二级序数) = 一级序数

                                '属性继承：面板 → 节点_面板
                                Exit For
                            End If
                        Next
                    Next
                Case 3      '根据节点分支数目确定单元类型
                    For 一级序数 As UShort = 1 To 节点_总数
                        Select Case 节点_分支_数目(一级序数)
                            Case >= 3
                                节点_tp(一级序数) = "硬角单元"
                                硬角单元_总数 += 1

                                节点_硬角单元_序数(一级序数) = 硬角单元_总数
                                ReDim Preserve 硬角单元_节点_序数(硬角单元_总数)
                                硬角单元_节点_序数(硬角单元_总数) = 一级序数

                                Exit Select
                            Case 2
                                Dim 双分支夹角 As Single = 节点_分支_α(一级序数, 1) - 节点_分支_α(一级序数, 2)
                                Select Case 双分支夹角
                                    Case >= 2 * PI
                                        双分支夹角 -= 2 * PI
                                    Case < 0
                                        双分支夹角 += 2 * PI
                                    Case Else

                                End Select
                                Select Case 双分支夹角
                                    Case <= 5 / 6 * PI
                                        节点_tp(一级序数) = "硬角单元"
                                        硬角单元_总数 += 1

                                        节点_硬角单元_序数(一级序数) = 硬角单元_总数
                                        ReDim Preserve 硬角单元_节点_序数(硬角单元_总数)
                                        硬角单元_节点_序数(硬角单元_总数) = 一级序数

                                        Exit Select
                                    Case >= 7 / 6 * PI
                                        节点_tp(一级序数) = "硬角单元"
                                        硬角单元_总数 += 1

                                        节点_硬角单元_序数(一级序数) = 硬角单元_总数
                                        ReDim Preserve 硬角单元_节点_序数(硬角单元_总数)
                                        硬角单元_节点_序数(硬角单元_总数) = 一级序数

                                        Exit Select
                                    Case Else
                                        Select Case 节点_tp(一级序数)
                                            Case "加强筋"
                                                节点_tp(一级序数) = "加强筋单元"
                                                加强筋单元_总数 += 1

                                                节点_加强筋单元_序数(一级序数) = 加强筋单元_总数
                                                ReDim Preserve 加强筋单元_节点_序数(加强筋单元_总数)
                                                加强筋单元_节点_序数(加强筋单元_总数) = 一级序数

                                                Exit Select
                                            Case Else

                                        End Select
                                End Select
                            Case 1
                                Select Case 节点_tp(一级序数)
                                    Case "加强筋"
                                        节点_tp(一级序数) = "加强筋单元"
                                        加强筋单元_总数 += 1

                                        节点_加强筋单元_序数(一级序数) = 加强筋单元_总数
                                        ReDim Preserve 加强筋单元_节点_序数(加强筋单元_总数)
                                        加强筋单元_节点_序数(加强筋单元_总数) = 一级序数

                                        Exit Select
                                    Case "面板"
                                        Select Case 面板_PMA(节点_面板_序数(一级序数))
                                            Case True
                                                节点_tp(一级序数) = "面板加强筋单元"
                                                面板加强筋单元_总数 += 1

                                                节点_面板加强筋单元_序数(一级序数) = 面板加强筋单元_总数
                                                ReDim Preserve 面板加强筋单元_节点_序数(面板加强筋单元_总数)
                                                面板加强筋单元_节点_序数(面板加强筋单元_总数) = 一级序数

                                                Exit Select
                                            Case False
                                                节点_tp(一级序数) = "面板硬角单元"
                                                面板硬角单元_总数 += 1

                                                节点_面板硬角单元_序数(一级序数) = 面板硬角单元_总数
                                                ReDim Preserve 面板硬角单元_节点_序数(面板硬角单元_总数)
                                                面板硬角单元_节点_序数(面板硬角单元_总数) = 一级序数

                                                Exit Select
                                        End Select
                                    Case Else
                                        节点_tp(一级序数) = "自由端"

                                        Exit Select
                                End Select
                        End Select
                    Next
                Case 4      '确定单元属性
                    ReDim 硬角单元_A(硬角单元_总数), 硬角单元_Yc(硬角单元_总数), 硬角单元_Zc(硬角单元_总数), 硬角单元_σY(硬角单元_总数)
                    ReDim 加强筋单元_A(加强筋单元_总数), 加强筋单元_Yc(加强筋单元_总数), 加强筋单元_Zc(加强筋单元_总数), 加强筋单元_σY(加强筋单元_总数)
                    For 单元对象类型序数 As UShort = 1 To 6
                        Dim 通用单元数目 As UShort = 加强筋单元_总数
                        Dim 单元分支_w(通用单元数目, 通用分支数目) As Single,
                            单元分支_A(通用单元数目, 通用分支数目) As Single,
                            单元分支_AYc(通用单元数目, 通用分支数目) As Single, 单元分支_AZc(通用单元数目, 通用分支数目) As Single,
                            单元分支_AσY(通用单元数目, 通用分支数目) As Single

                        Select Case 单元对象类型序数
                            Case 1      '硬角单元
                                For 一级序数 As UShort = 1 To 硬角单元_总数
                                    Dim 硬角单元_AYc(硬角单元_总数) As Single, 硬角单元_AZc(硬角单元_总数) As Single,
                                        硬角单元_AσY(硬角单元_总数) As Single

                                    Dim 关联节点序数 As UShort = 硬角单元_节点_序数(一级序数)
                                    For 二级序数 As UShort = 1 To 节点_分支_数目(关联节点序数)
                                        Dim 关联板格序数 As UShort = 节点_分支_板格_序数(关联节点序数, 二级序数)

                                        Select Case 板格_tp(关联板格序数)
                                            Case "L"
                                                单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数) / 2
                                                If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 二级序数)) = "自由端" Then
                                                    单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数)
                                                End If
                                            Case "T"
                                                单元分支_w(一级序数, 二级序数) = 板格_t(关联板格序数) * 20
                                            Case Else

                                        End Select
                                        Dim 分支板_A(硬角单元_总数, 通用分支数目) As Single,
                                            分支板_AYc(硬角单元_总数, 通用分支数目) As Single, 分支板_AZc(硬角单元_总数, 通用分支数目) As Single,
                                            分支板_AσY(硬角单元_总数, 通用分支数目) As Single
                                        For 三级序数 As UShort = 1 To 板格_板格板数目(关联板格序数)
                                            If 关联节点序数 = 板格_首端节点_序数(关联板格序数) Then
                                                板格_首端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                Dim 分支板_w(硬角单元_总数, 通用分支数目) As Single
                                                分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 三级序数)
                                                Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                Select Case 分支板超出宽度
                                                    Case <= 0   '分支板宽度和小于分支所需宽度
                                                        分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 三级序数)
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 三级序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 三级序数)
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 三级序数)
                                                    Case > 0    '分支板宽度和大于分支所需宽度
                                                        Dim 通用板格板数目 As UShort = 3
                                                        Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                        '板格板实际取用宽度
                                                        板格板_w1(关联板格序数, 三级序数) = 板格板_w(关联板格序数, 三级序数) - 分支板超出宽度
                                                        '板格板宽度利用系数
                                                        Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 三级序数) / 板格板_w(关联板格序数, 三级序数)
                                                        Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y11(板格_总数, 通用板格板数目), 板格板_Z11(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                        '板格板实际取用面积
                                                        板格板_A1(关联板格序数, 三级序数) = 板格板_A(关联板格序数, 三级序数) * 板格板_ηw
                                                        '板格板实际末端坐标
                                                        板格板_Y11(关联板格序数, 三级序数) = 板格板_Y0(关联板格序数, 三级序数) + (板格板_Y1(关联板格序数, 三级序数) - 板格板_Y0(关联板格序数, 三级序数)) * 板格板_ηw
                                                        板格板_Z11(关联板格序数, 三级序数) = 板格板_Z0(关联板格序数, 三级序数) + (板格板_Z1(关联板格序数, 三级序数) - 板格板_Z0(关联板格序数, 三级序数)) * 板格板_ηw
                                                        '板格板实际形心坐标
                                                        板格板_Yc1(关联板格序数, 三级序数) = (板格板_Y0(关联板格序数, 三级序数) + 板格板_Y11(关联板格序数, 三级序数)） / 2
                                                        板格板_Zc1(关联板格序数, 三级序数) = (板格板_Z0(关联板格序数, 三级序数) + 板格板_Z11(关联板格序数, 三级序数)） / 2
                                                        '板格板实际面积坐标积数
                                                        板格板_AYc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Yc1(关联板格序数, 三级序数)
                                                        板格板_AZc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Zc1(关联板格序数, 三级序数)
                                                        '板格板实际面积强度积数
                                                        板格板_AσY1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_σY(关联板格序数, 三级序数)

                                                        'Select Case 三级序数
                                                        '    Case = 1

                                                        '    Case > 1

                                                        'End Select

                                                        '分支板实际面积
                                                        分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 三级序数)
                                                        '分支板实际面积坐标积数
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 三级序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 三级序数)
                                                        '分支板实际面积强度积数
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 三级序数)

                                                        Exit For
                                                End Select
                                            ElseIf 关联节点序数 = 板格_末端节点_序数(关联板格序数) Then
                                                板格_末端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                Dim 分支板_w(硬角单元_总数, 通用分支数目) As Single
                                                Dim 反向序数 As UShort = 板格_板格板数目(关联板格序数) + 1 - 三级序数
                                                分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 反向序数)
                                                Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                Select Case 分支板超出宽度
                                                    Case <= 0   '分支板宽度和小于分支所需宽度
                                                        分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 反向序数)
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 反向序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 反向序数)
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 反向序数)
                                                    Case > 0    '分支板宽度和大于分支所需宽度
                                                        Dim 通用板格板数目 As UShort = 3
                                                        Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                        '板格板实际取用宽度
                                                        板格板_w1(关联板格序数, 反向序数) = 板格板_w(关联板格序数, 反向序数) - 分支板超出宽度
                                                        '板格板宽度利用系数
                                                        Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 反向序数) / 板格板_w(关联板格序数, 反向序数)
                                                        Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y01(板格_总数, 通用板格板数目), 板格板_Z01(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                        '板格板实际取用面积
                                                        板格板_A1(关联板格序数, 反向序数) = 板格板_A(关联板格序数, 反向序数) * 板格板_ηw
                                                        '板格板实际末端坐标
                                                        板格板_Y01(关联板格序数, 反向序数) = 板格板_Y1(关联板格序数, 反向序数) + (板格板_Y0(关联板格序数, 反向序数) - 板格板_Y1(关联板格序数, 反向序数)) * 板格板_ηw
                                                        板格板_Z01(关联板格序数, 反向序数) = 板格板_Z1(关联板格序数, 反向序数) + (板格板_Z0(关联板格序数, 反向序数) - 板格板_Z1(关联板格序数, 反向序数)) * 板格板_ηw
                                                        '板格板实际形心坐标
                                                        板格板_Yc1(关联板格序数, 反向序数) = (板格板_Y01(关联板格序数, 反向序数) + 板格板_Y1(关联板格序数, 反向序数)） / 2
                                                        板格板_Zc1(关联板格序数, 反向序数) = (板格板_Z01(关联板格序数, 反向序数) + 板格板_Z1(关联板格序数, 反向序数)） / 2
                                                        '板格板实际面积坐标积数
                                                        板格板_AYc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Yc1(关联板格序数, 反向序数)
                                                        板格板_AZc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Zc1(关联板格序数, 反向序数)
                                                        '板格板实际面积强度积数
                                                        板格板_AσY1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_σY(关联板格序数, 反向序数)

                                                        'Select Case 三级序数
                                                        '    Case = 1

                                                        '    Case > 1

                                                        'End Select

                                                        '分支板实际面积
                                                        分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 反向序数)
                                                        '分支板实际面积坐标积数
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 反向序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 反向序数)
                                                        '分支板实际面积强度积数
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 反向序数)

                                                        Exit For
                                                End Select
                                            End If
                                        Next
                                        单元分支_A(一级序数, 二级序数) = 分支板_A(一级序数, 二级序数)
                                        单元分支_AYc(一级序数, 二级序数) = 分支板_AYc(一级序数, 二级序数)
                                        单元分支_AZc(一级序数, 二级序数) = 分支板_AZc(一级序数, 二级序数)
                                        单元分支_AσY(一级序数, 二级序数) = 分支板_AσY(一级序数, 二级序数)

                                        硬角单元_A(一级序数) += 分支板_A(一级序数, 二级序数)
                                        硬角单元_AYc(一级序数) += 分支板_AYc(一级序数, 二级序数)
                                        硬角单元_AZc(一级序数) += 分支板_AZc(一级序数, 二级序数)
                                        硬角单元_AσY(一级序数) += 分支板_AσY(一级序数, 二级序数)
                                    Next

                                    硬角单元_Yc(一级序数) = 硬角单元_AYc(一级序数) / 硬角单元_A(一级序数)
                                    硬角单元_Zc(一级序数) = 硬角单元_AZc(一级序数) / 硬角单元_A(一级序数)
                                    硬角单元_σY(一级序数) = 硬角单元_AσY(一级序数) / 硬角单元_A(一级序数)

                                    全截面_A += 硬角单元_A(一级序数)
                                    全截面_AYc += 硬角单元_AYc(一级序数)
                                    全截面_AZc += 硬角单元_AZc(一级序数)
                                    全截面_AσY += 硬角单元_AσY(一级序数)
                                Next
                            Case 2      '面板硬角单元
                                'MsgBox("面板硬角单元部分未完成!")
                            Case 3      '特别硬角单元
                                For 一级序数 As UShort = 1 To 特别硬角单元_总数
                                    全截面_A += 特别硬角单元_A(一级序数)
                                    全截面_AYc += 特别硬角单元_AYc(一级序数)
                                    全截面_AZc += 特别硬角单元_AZc(一级序数)
                                    全截面_AσY += 特别硬角单元_AσY(一级序数)
                                Next
                            Case 4      '加强筋单元
                                ReDim 加强筋单元_σYP(加强筋单元_总数),
                                    加强筋单元_lP(加强筋单元_总数),
                                    加强筋单元_wP(加强筋单元_总数), 加强筋单元_tP(加强筋单元_总数),
                                    加强筋单元_σYS(加强筋单元_总数),
                                    加强筋单元_lS(加强筋单元_总数),
                                    加强筋单元_hw(加强筋单元_总数), 加强筋单元_tw(加强筋单元_总数),
                                    加强筋单元_wf(加强筋单元_总数), 加强筋单元_tf(加强筋单元_总数),
                                    加强筋单元_dx(加强筋单元_总数),
                                    加强筋单元_tpS(加强筋单元_总数), 加强筋单元_mk(加强筋单元_总数)

                                ReDim 加强筋单元_ηS(加强筋单元_总数), 加强筋单元_σETS(加强筋单元_总数)

                                ReDim 加强筋单元_σELS(加强筋单元_总数)

                                Dim 加强筋单元_AYc(加强筋单元_总数) As Single, 加强筋单元_AZc(加强筋单元_总数) As Single,
                                    加强筋单元_AσY(加强筋单元_总数) As Single

                                For 一级序数 As UShort = 1 To 加强筋单元_总数
                                    Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(一级序数)
                                    Select Case 节点_分支_数目(关联节点序数)
                                        Case 1
                                            Select Case 板格_tp(节点_分支_板格_序数(关联节点序数, 1))
                                                Case "L"
                                                    单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 1)) = "自由端" Then
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1))
                                                    End If
                                                    加强筋单元_lP(一级序数) = 板格_l(节点_分支_板格_序数(关联节点序数, 1))
                                                    加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1)
                                                    'MsgBox("警告:单L分支")
                                                Case "T"
                                                    'MsgBox("错误:单T分支")
                                                Case Else

                                            End Select
                                        Case 2
                                            Dim 板格_联合_tp As String = 板格_tp(节点_分支_板格_序数(关联节点序数, 1)) & 板格_tp(节点_分支_板格_序数(关联节点序数, 2))
                                            Select Case 板格_联合_tp
                                                Case "LL"
                                                    单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                    单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2)) / 2
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 1)) = "自由端" Then
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1))
                                                    End If
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 2)) = "自由端" Then
                                                        单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2))
                                                    End If
                                                    加强筋单元_lP(一级序数) = (板格_l(节点_分支_板格_序数(关联节点序数, 1)) + 板格_l(节点_分支_板格_序数(关联节点序数, 2))) / 2
                                                    加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1) + 单元分支_w(一级序数, 2)
                                                Case "LT"
                                                    单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 1)) = "自由端" Then
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1))
                                                    End If
                                                    单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                    加强筋单元_lP(一级序数) = (板格_l(节点_分支_板格_序数(关联节点序数, 1)) + 板格_l(节点_分支_板格_序数(关联节点序数, 2))) / 2
                                                    加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1) + 单元分支_w(一级序数, 2)
                                                    MsgBox("警告:单L单T分支")
                                                Case "TL"
                                                    单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 2)) / 2
                                                    单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2)) / 2
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 2)) = "自由端" Then
                                                        单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2))
                                                    End If
                                                    加强筋单元_lP(一级序数) = (板格_l(节点_分支_板格_序数(关联节点序数, 1)) + 板格_l(节点_分支_板格_序数(关联节点序数, 2))) / 2
                                                    加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1) + 单元分支_w(一级序数, 2)
                                                    MsgBox("警告:单T单L分支")
                                                Case "TT"
                                                    MsgBox("错误:双T分支")
                                                Case Else

                                            End Select
                                    End Select

                                    Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)
                                    加强筋单元_A(一级序数) += 加强筋_AS(关联加强筋序数)
                                    加强筋单元_AYc(一级序数) += 加强筋_AS(关联加强筋序数) * 加强筋_YcS(关联加强筋序数)
                                    加强筋单元_AZc(一级序数) += 加强筋_AS(关联加强筋序数) * 加强筋_ZcS(关联加强筋序数)
                                    加强筋单元_AσY(一级序数) += 加强筋_AS(关联加强筋序数) * 加强筋_σY(关联加强筋序数)

                                    加强筋单元_σYS(一级序数) = 加强筋_σY(关联加强筋序数)

                                    加强筋单元_lS(一级序数) = 加强筋_l(关联加强筋序数)

                                    加强筋单元_hw(一级序数) = 加强筋_hw(关联加强筋序数)
                                    加强筋单元_tw(一级序数) = 加强筋_tw(关联加强筋序数)

                                    加强筋单元_wf(一级序数) = 加强筋_wf(关联加强筋序数)
                                    加强筋单元_tf(一级序数) = 加强筋_tf(关联加强筋序数)

                                    加强筋单元_dx(一级序数) = 加强筋_dx(关联加强筋序数)

                                    加强筋单元_tpS(一级序数) = 加强筋_tp(关联加强筋序数)
                                    加强筋单元_mk(一级序数) = 加强筋_mk(关联加强筋序数)

                                    加强筋单元_σELS(一级序数) = 加强筋_σELS(关联加强筋序数)

                                    For 二级序数 As UShort = 1 To 节点_分支_数目(关联节点序数)
                                        Dim 关联板格序数 As UShort = 节点_分支_板格_序数(关联节点序数, 二级序数)

                                        'Select Case 板格_tp(关联板格序数)
                                        '    Case "L"
                                        '        单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数) / 2
                                        '        If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 二级序数)) = "自由端" Then
                                        '            单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数)
                                        '        End If
                                        '    Case "T"
                                        '        单元分支_w(一级序数, 二级序数) = 板格_t(关联板格序数) * 20
                                        'End Select
                                        Dim 分支板_A(加强筋单元_总数, 通用分支数目) As Single,
                                            分支板_AYc(加强筋单元_总数, 通用分支数目) As Single, 分支板_AZc(加强筋单元_总数, 通用分支数目) As Single,
                                            分支板_AσY(加强筋单元_总数, 通用分支数目) As Single
                                        For 三级序数 As UShort = 1 To 板格_板格板数目(关联板格序数)
                                            If 关联节点序数 = 板格_首端节点_序数(关联板格序数) Then
                                                板格_首端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                Dim 分支板_w(加强筋单元_总数, 通用分支数目) As Single
                                                分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 三级序数)
                                                Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                Select Case 分支板超出宽度
                                                    Case <= 0   '分支板宽度和小于分支所需宽度
                                                        分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 三级序数)
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 三级序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 三级序数)
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 三级序数)
                                                    Case > 0    '分支板宽度和大于分支所需宽度
                                                        Dim 通用板格板数目 As UShort = 3
                                                        Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                        '板格板实际取用宽度
                                                        板格板_w1(关联板格序数, 三级序数) = 板格板_w(关联板格序数, 三级序数) - 分支板超出宽度
                                                        '板格板宽度利用系数
                                                        Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 三级序数) / 板格板_w(关联板格序数, 三级序数)
                                                        Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y11(板格_总数, 通用板格板数目), 板格板_Z11(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                        '板格板实际取用面积
                                                        板格板_A1(关联板格序数, 三级序数) = 板格板_A(关联板格序数, 三级序数) * 板格板_ηw
                                                        '板格板实际末端坐标
                                                        板格板_Y11(关联板格序数, 三级序数) = 板格板_Y0(关联板格序数, 三级序数) + (板格板_Y1(关联板格序数, 三级序数) - 板格板_Y0(关联板格序数, 三级序数)) * 板格板_ηw
                                                        板格板_Z11(关联板格序数, 三级序数) = 板格板_Z0(关联板格序数, 三级序数) + (板格板_Z1(关联板格序数, 三级序数) - 板格板_Z0(关联板格序数, 三级序数)) * 板格板_ηw
                                                        '板格板实际形心坐标
                                                        板格板_Yc1(关联板格序数, 三级序数) = (板格板_Y0(关联板格序数, 三级序数) + 板格板_Y11(关联板格序数, 三级序数)） / 2
                                                        板格板_Zc1(关联板格序数, 三级序数) = (板格板_Z0(关联板格序数, 三级序数) + 板格板_Z11(关联板格序数, 三级序数)） / 2
                                                        '板格板实际面积坐标积数
                                                        板格板_AYc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Yc1(关联板格序数, 三级序数)
                                                        板格板_AZc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Zc1(关联板格序数, 三级序数)
                                                        '板格板实际面积强度积数
                                                        板格板_AσY1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_σY(关联板格序数, 三级序数)

                                                        'Select Case 三级序数
                                                        '    Case = 1

                                                        '    Case > 1

                                                        'End Select

                                                        '分支板实际面积
                                                        分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 三级序数)
                                                        '分支板实际面积坐标积数
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 三级序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 三级序数)
                                                        '分支板实际面积强度积数
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 三级序数)

                                                        Exit For
                                                End Select
                                            ElseIf 关联节点序数 = 板格_末端节点_序数(关联板格序数) Then
                                                板格_末端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                Dim 分支板_w(加强筋单元_总数, 通用分支数目) As Single
                                                Dim 反向序数 As UShort = 板格_板格板数目(关联板格序数) + 1 - 三级序数
                                                分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 反向序数)
                                                Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                Select Case 分支板超出宽度
                                                    Case <= 0   '分支板宽度和小于分支所需宽度
                                                        分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 反向序数)
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 反向序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 反向序数)
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 反向序数)
                                                    Case > 0    '分支板宽度和大于分支所需宽度
                                                        Dim 通用板格板数目 As UShort = 3
                                                        Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                        '板格板实际取用宽度
                                                        板格板_w1(关联板格序数, 反向序数) = 板格板_w(关联板格序数, 反向序数) - 分支板超出宽度
                                                        '板格板宽度利用系数
                                                        Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 反向序数) / 板格板_w(关联板格序数, 反向序数)
                                                        Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y01(板格_总数, 通用板格板数目), 板格板_Z01(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                        '板格板实际取用面积
                                                        板格板_A1(关联板格序数, 反向序数) = 板格板_A(关联板格序数, 反向序数) * 板格板_ηw
                                                        '板格板实际末端坐标
                                                        板格板_Y01(关联板格序数, 反向序数) = 板格板_Y1(关联板格序数, 反向序数) + (板格板_Y0(关联板格序数, 反向序数) - 板格板_Y1(关联板格序数, 反向序数)) * 板格板_ηw
                                                        板格板_Z01(关联板格序数, 反向序数) = 板格板_Z1(关联板格序数, 反向序数) + (板格板_Z0(关联板格序数, 反向序数) - 板格板_Z1(关联板格序数, 反向序数)) * 板格板_ηw
                                                        '板格板实际形心坐标
                                                        板格板_Yc1(关联板格序数, 反向序数) = (板格板_Y01(关联板格序数, 反向序数) + 板格板_Y1(关联板格序数, 反向序数)） / 2
                                                        板格板_Zc1(关联板格序数, 反向序数) = (板格板_Z01(关联板格序数, 反向序数) + 板格板_Z1(关联板格序数, 反向序数)） / 2
                                                        '板格板实际面积坐标积数
                                                        板格板_AYc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Yc1(关联板格序数, 反向序数)
                                                        板格板_AZc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Zc1(关联板格序数, 反向序数)
                                                        '板格板实际面积强度积数
                                                        板格板_AσY1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_σY(关联板格序数, 反向序数)

                                                        'Select Case 三级序数
                                                        '    Case = 1

                                                        '    Case > 1

                                                        'End Select

                                                        '分支板实际面积
                                                        分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 反向序数)
                                                        '分支板实际面积坐标积数
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 反向序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 反向序数)
                                                        '分支板实际面积强度积数
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 反向序数)

                                                        Exit For
                                                End Select
                                            End If
                                        Next
                                        单元分支_A(一级序数, 二级序数) = 分支板_A(一级序数, 二级序数)
                                        单元分支_AYc(一级序数, 二级序数) = 分支板_AYc(一级序数, 二级序数)
                                        单元分支_AZc(一级序数, 二级序数) = 分支板_AZc(一级序数, 二级序数)
                                        单元分支_AσY(一级序数, 二级序数) = 分支板_AσY(一级序数, 二级序数)

                                        加强筋单元_A(一级序数) += 分支板_A(一级序数, 二级序数)
                                        加强筋单元_AYc(一级序数) += 分支板_AYc(一级序数, 二级序数)
                                        加强筋单元_AZc(一级序数) += 分支板_AZc(一级序数, 二级序数)
                                        加强筋单元_AσY(一级序数) += 分支板_AσY(一级序数, 二级序数)
                                    Next

                                    加强筋单元_tP(一级序数) = (加强筋单元_A(一级序数) - 加强筋_AS(关联加强筋序数)) / 加强筋单元_wP(一级序数)
                                    加强筋单元_σYP(一级序数) = (加强筋单元_AσY(一级序数) - 加强筋_AS(关联加强筋序数) * 加强筋_σY(关联加强筋序数)) / (加强筋单元_A(一级序数) - 加强筋_AS(关联加强筋序数))

                                    加强筋单元_ηS(一级序数) = 1 + 加强筋_l(关联加强筋序数) ^ 2 / PI ^ 2 / Sqrt(加强筋_IWS(关联加强筋序数) * (0.75 * 加强筋单元_wP(一级序数) / 加强筋单元_tP(一级序数) ^ 3 + (加强筋_df(关联加强筋序数) - 加强筋_tf(关联加强筋序数) / 2) / 加强筋_tw(关联加强筋序数) ^ 3))
                                    加强筋单元_σETS(一级序数) = 标准弹性模量 / 加强筋_IPS(关联加强筋序数) * (加强筋单元_ηS(一级序数) * PI ^ 2 * 加强筋_IWS(关联加强筋序数) / 加强筋单元_lS(一级序数) ^ 2 + 0.385 * 加强筋_ITS(关联加强筋序数))

                                    加强筋单元_Yc(一级序数) = 加强筋单元_AYc(一级序数) / 加强筋单元_A(一级序数)
                                    加强筋单元_Zc(一级序数) = 加强筋单元_AZc(一级序数) / 加强筋单元_A(一级序数)
                                    加强筋单元_σY(一级序数) = 加强筋单元_AσY(一级序数) / 加强筋单元_A(一级序数)

                                    全截面_A += 加强筋单元_A(一级序数)
                                    全截面_AYc += 加强筋单元_AYc(一级序数)
                                    全截面_AZc += 加强筋单元_AZc(一级序数)
                                    全截面_AσY += 加强筋单元_AσY(一级序数)
                                Next
                            Case 5      '面板加强筋单元
                                'MsgBox("面板加强筋单元部分未完成!")
                            Case 6      '加筋板单元
                                For 一级序数 As UShort = 1 To 板格_总数
                                    Dim 板格_原始_w As Single, 板格_剩余_w As Single
                                    Dim 板格_原始_A As Single, 板格_首端_A As Single, 板格_末端_A As Single, 板格_剩余_A As Single
                                    Dim 板格_原始_AYc As Single, 板格_首端_AYc As Single, 板格_末端_AYc As Single, 板格_剩余_AYc As Single
                                    Dim 板格_原始_AZc As Single, 板格_首端_AZc As Single, 板格_末端_AZc As Single, 板格_剩余_AZc As Single
                                    Dim 板格_原始_AσY As Single, 板格_首端_AσY As Single, 板格_末端_AσY As Single, 板格_剩余_AσY As Single

                                    Dim 首端关联节点序数 As UShort, 首端关联单元序数 As UShort, 首端关联分支序数 As UShort
                                    Dim 末端关联节点序数 As UShort, 末端关联单元序数 As UShort, 末端关联分支序数 As UShort

                                    首端关联节点序数 = 板格_首端节点_序数(一级序数)
                                    Select Case 节点_tp(首端关联节点序数)
                                        Case "硬角单元"
                                            首端关联单元序数 = 节点_硬角单元_序数(首端关联节点序数)
                                        Case "面板硬角单元"
                                            首端关联单元序数 = 节点_面板硬角单元_序数(首端关联节点序数)
                                        Case "加强筋单元"
                                            首端关联单元序数 = 节点_加强筋单元_序数(首端关联节点序数)
                                        Case "面板硬角单元"
                                            首端关联单元序数 = 节点_面板硬角单元_序数(首端关联节点序数)
                                        Case Else

                                    End Select
                                    首端关联分支序数 = 板格_首端节点_分支_序数(一级序数)

                                    末端关联节点序数 = 板格_末端节点_序数(一级序数)
                                    Select Case 节点_tp(末端关联节点序数)
                                        Case "硬角单元"
                                            末端关联单元序数 = 节点_硬角单元_序数(末端关联节点序数)
                                        Case "面板硬角单元"
                                            末端关联单元序数 = 节点_面板硬角单元_序数(末端关联节点序数)
                                        Case "加强筋单元"
                                            末端关联单元序数 = 节点_加强筋单元_序数(末端关联节点序数)
                                        Case "面板硬角单元"
                                            末端关联单元序数 = 节点_面板硬角单元_序数(末端关联节点序数)
                                        Case Else

                                    End Select
                                    末端关联分支序数 = 板格_末端节点_分支_序数(一级序数)

                                    板格_原始_w = 板格_w(一级序数)
                                    板格_剩余_w = 板格_原始_w - 板格_首端_w(一级序数) - 板格_末端_w(一级序数)

                                    板格_原始_A = 板格_A(一级序数)
                                    板格_首端_A = 单元分支_A(首端关联单元序数, 首端关联分支序数)
                                    板格_末端_A = 单元分支_A(末端关联单元序数, 末端关联分支序数)
                                    板格_剩余_A = 板格_原始_A - 板格_首端_A - 板格_末端_A

                                    板格_原始_AYc = 板格_AYc(一级序数)
                                    板格_首端_AYc = 单元分支_AYc(首端关联单元序数, 首端关联分支序数)
                                    板格_末端_AYc = 单元分支_AYc(末端关联单元序数, 末端关联分支序数)
                                    板格_剩余_AYc = 板格_原始_AYc - 板格_首端_AYc - 板格_末端_AYc

                                    板格_原始_AZc = 板格_AZc(一级序数)
                                    板格_首端_AZc = 单元分支_AZc(首端关联单元序数, 首端关联分支序数)
                                    板格_末端_AZc = 单元分支_AZc(末端关联单元序数, 末端关联分支序数)
                                    板格_剩余_AZc = 板格_原始_AZc - 板格_首端_AZc - 板格_末端_AZc

                                    板格_原始_AσY = 板格_AσY(一级序数)
                                    板格_首端_AσY = 单元分支_AσY(首端关联单元序数, 首端关联分支序数)
                                    板格_末端_AσY = 单元分支_AσY(末端关联单元序数, 末端关联分支序数)
                                    板格_剩余_AσY = 板格_原始_AσY - 板格_首端_AσY - 板格_末端_AσY

                                    Select Case 板格_剩余_w / 板格_原始_w
                                        Case >= 0.1
                                            加筋板单元_总数 += 1

                                            ReDim Preserve 加筋板单元_板格_序数(加筋板单元_总数)
                                            加筋板单元_板格_序数(加筋板单元_总数) = 一级序数

                                            ReDim Preserve 加筋板单元_原始_w(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_原始_A(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_原始_Yc(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_原始_Zc(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_原始_σY(加筋板单元_总数)

                                            ReDim Preserve 加筋板单元_剩余_w(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_剩余_A(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_剩余_Yc(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_剩余_Zc(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_剩余_σY(加筋板单元_总数)

                                            加筋板单元_原始_w(加筋板单元_总数) = 板格_原始_w
                                            加筋板单元_原始_A(加筋板单元_总数) = 板格_原始_A
                                            加筋板单元_原始_Yc(加筋板单元_总数) = 板格_原始_AYc / 板格_原始_A
                                            加筋板单元_原始_Zc(加筋板单元_总数) = 板格_原始_AZc / 板格_原始_A
                                            加筋板单元_原始_σY(加筋板单元_总数) = 板格_原始_AσY / 板格_原始_A

                                            加筋板单元_剩余_w(加筋板单元_总数) = 板格_剩余_w
                                            加筋板单元_剩余_A(加筋板单元_总数) = 板格_剩余_A
                                            加筋板单元_剩余_Yc(加筋板单元_总数) = 板格_剩余_AYc / 板格_剩余_A
                                            加筋板单元_剩余_Zc(加筋板单元_总数) = 板格_剩余_AZc / 板格_剩余_A
                                            加筋板单元_剩余_σY(加筋板单元_总数) = 板格_剩余_AσY / 板格_剩余_A

                                            全截面_A += 板格_剩余_A
                                            全截面_AYc += 板格_剩余_AYc
                                            全截面_AZc += 板格_剩余_AZc
                                            全截面_AσY += 板格_剩余_AσY
                                        Case Else

                                    End Select
                                Next
                            Case Else

                        End Select
                    Next
                    全截面_Yc = 全截面_AYc / 全截面_A
                    全截面_Zc = 全截面_AZc / 全截面_A
            End Select
        Next

        '单元对象的形心坐标输出
        For 单元对象类型序数 As UShort = 1 To 6
            Select Case 单元对象类型序数
                Case 1
                    Chart2.Series.Clear()
                    Chart2.Show()
                    Chart2.Series.Add("硬角单元_形心")
                    Chart2.Series("硬角单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 硬角单元_总数
                        Chart2.Series("硬角单元_形心").Points.AddXY(硬角单元_Yc(一级序数), 硬角单元_Zc(一级序数))
                    Next
                Case 2
                    'MsgBox("面板硬角单元部分未完成!")
                Case 3
                    Chart2.Series.Add("特别硬角单元_形心")
                    Chart2.Series("特别硬角单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 特别硬角单元_总数
                        Chart2.Series("特别硬角单元_形心").Points.AddXY(特别硬角单元_Yc(一级序数), 特别硬角单元_Zc(一级序数))
                    Next
                Case 4
                    Chart2.Series.Add("加强筋单元_形心")
                    Chart2.Series("加强筋单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 加强筋单元_总数
                        Chart2.Series("加强筋单元_形心").Points.AddXY(加强筋单元_Yc(一级序数), 加强筋单元_Zc(一级序数))
                    Next
                Case 5
                    'MsgBox("面板加强筋单元部分未完成!")
                Case 6
                    Chart2.Series.Add("加筋板单元_形心")
                    Chart2.Series("加筋板单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 加筋板单元_总数
                        Chart2.Series("加筋板单元_形心").Points.AddXY(加筋板单元_剩余_Yc(一级序数), 加筋板单元_剩余_Zc(一级序数))
                    Next
                Case Else

            End Select
        Next

    End Sub

    Private Sub 刚度计算(sender As Object, e As EventArgs) Handles Button4.Click
        Dim 插值点_总数 As UShort = InputBox("插值点_总数", "插值点_总数", "500")

        ReDim 硬角单元_D(硬角单元_总数, 插值点_总数), 面板硬角单元_D(面板硬角单元_总数, 插值点_总数), 特别硬角单元_D(特别硬角单元_总数, 插值点_总数), 加强筋单元_屈服_D(加强筋单元_总数, 插值点_总数), 加强筋单元_屈曲_D(加强筋单元_总数, 插值点_总数), 面板加强筋单元_D(面板加强筋单元_总数, 插值点_总数), 加筋板单元_剩余_D(加筋板单元_总数, 插值点_总数)

        'Dim EP_D(加强筋单元_总数, 插值点_总数) As Single
        'Dim BC_D(加强筋单元_总数, 插值点_总数) As Single
        'Dim FT_D(加强筋单元_总数, 插值点_总数) As Single
        'Dim WB_D(加强筋单元_总数, 插值点_总数) As Single
        Dim Min_D(加强筋单元_总数, 插值点_总数) As Single

        'Dim σEPEε(加强筋单元_总数, 插值点_总数) As Single
        Dim σBCEε(加强筋单元_总数, 插值点_总数) As Single
        Dim σFTEε(加强筋单元_总数, 插值点_总数) As Single
        Dim σWBEε(加强筋单元_总数, 插值点_总数) As Single
        Dim σMINEε(加强筋单元_总数, 插值点_总数) As Single

        Dim σPBEε(加筋板单元_总数, 插值点_总数) As Single

        For 单元对象类型序数 As UShort = 1 To 6
            Select Case 单元对象类型序数
                Case 1
                    For 一级序数 As UShort = 1 To 硬角单元_总数
                        'Dim 输出 As String = "硬角单元" & Space(8) & Format(一级序数, "000") & Space(4)
                        For 二级序数 As UShort = 0 To 插值点_总数 - 1
                            Select Case 二级序数
                                Case < 100
                                    硬角单元_D(一级序数, 二级序数) = 硬角单元_σY(一级序数)
                                Case >= 100
                                    硬角单元_D(一级序数, 二级序数) = 0
                                Case Else

                            End Select
                            '输出 = 输出 & Format(硬角单元_D(一级序数, 二级序数), "0.000E+000") & Space(4)
                        Next
                        'Debug.Print(输出)
                    Next
                Case 2
                    'MsgBox("面板硬角单元部分未完成!")
                Case 3
                    For 一级序数 As UShort = 1 To 特别硬角单元_总数
                        'Dim 输出 As String = "特别硬角单元" & Space(8) & Format(一级序数, "000") & Space(4)
                        For 二级序数 As UShort = 0 To 插值点_总数 - 1
                            Select Case 二级序数
                                Case < 100
                                    特别硬角单元_D(一级序数, 二级序数) = 特别硬角单元_σY(一级序数)
                                Case >= 100
                                    特别硬角单元_D(一级序数, 二级序数) = 0
                                Case Else

                            End Select
                            '输出 = 输出 & Format(特别硬角单元_D(一级序数, 二级序数), "0.000E+000") & Space(4)
                        Next
                        'Debug.Print(输出)
                    Next
                Case 4
                    For 一级序数 As UShort = 1 To 加强筋单元_总数
                        'Dim 输出σ As String = "加强筋单元" & Space(8) & Format(一级序数, "000") & Space(4)
                        'Dim 输出D As String = "加强筋单元" & Space(8) & Format(一级序数, "000") & Space(4)
                        For 二级序数 As UShort = 0 To 插值点_总数
                            '输出σ = 输出σ & Format(二级序数, "000") & Space(4)

                            Dim tpS As String = 加强筋单元_tpS(一级序数)
                            Dim σYP As Single = 加强筋单元_σYP(一级序数)
                            Dim lP As Single = 加强筋单元_lP(一级序数)
                            Dim wP As Single = 加强筋单元_wP(一级序数)
                            Dim tP As Single = 加强筋单元_tP(一级序数)

                            Dim εYP As Single = 加强筋单元_σYP(一级序数) / 标准弹性模量
                            Dim AP As Single = wP * tP
                            Dim ICP As Single = wP * tP ^ 3 / 12
                            Dim IOP As Single = ICP + AP * (-tP / 2) ^ 2
                            Dim βOP As Single = wP / tP * Sqrt(εYP)
                            Dim wEOP As Single = If(βOP >= 1.25, (2.25 / βOP - 1.25 / βOP ^ 2) * wP, wP)

                            Dim σYS As Single = 加强筋单元_σYS(一级序数)
                            Dim lS As Single = 加强筋单元_lS(一级序数)
                            Dim hw As Single = 加强筋单元_hw(一级序数)
                            Dim tw As Single = 加强筋单元_tw(一级序数)
                            Dim wf As Single = 加强筋单元_wf(一级序数)
                            Dim tf As Single = 加强筋单元_tf(一级序数)
                            Dim dx As Single = 加强筋单元_dx(一级序数)
                            Dim εYS As Single = σYS / 标准弹性模量
                            Dim Aw As Single = hw * tw
                            Dim ICw As Single = tw * hw ^ 3 / 12
                            Dim IOw As Single = ICw + Aw * (hw / 2) ^ 2
                            Dim βOw As Single = hw / tw * Sqrt(εYS)
                            Dim hEOw As Single = If(βOw >= 1.25, (2.25 / βOw - 1.5 / βOw ^ 2) * hw, hw)
                            Dim df As Single = If(tpS = "F", hw, If(tpS = "B", hw - tf / 2, If(tpS = "T" Or tpS = "L1" Or tpS = "L2", hw + tf / 2, hw - dx - tf / 2)))
                            Dim Af As Single = wf * tf
                            Dim ICf As Single = wf * tf ^ 3 / 12
                            Dim IOf As Single = ICf + Af * df ^ 2
                            Dim ASS As Single = Aw + Af
                            Dim hCS As Single = Aw / ASS * hw / 2 + Af / ASS * df
                            Dim ICS As Single = ICw + Aw * (hw / 2 - hCS) ^ 2 + ICf + Af * (df - hCS) ^ 2
                            Dim IPS As Single = If(tpS = "F", hw ^ 3 * tw / 3, Aw * (df - tf / 2) ^ 2 / 3 + Af * df ^ 2)
                            Dim ITS As Single = If(tpS = "F", hw * tw ^ 3 / 3 * (1 - 0.63 * tw / hw), (df - tf / 2) * tw ^ 3 / 3 * (1 - 0.63 * tw / (df - tf / 2)) + wf * tf ^ 3 / 3 * (1 - 0.63 * tf / wf))
                            Dim IWS As Single = If(tpS = "F", hw ^ 3 * tw ^ 3 / 36, If(tpS = "B" Or tpS = "L1" Or tpS = "L2" Or tpS = "L3", Af * df ^ 2 * wf ^ 2 / 12 * (Af + 2.6 * Aw) / (Af * Aw), wf ^ 3 * tf * df ^ 2 / 12))
                            Dim ηS As Single = 1 + (lS / PI) ^ 2 / Sqrt(IWS * (0.75 * wP / tP ^ 3 + (df - tf / 2) / tw ^ 3))
                            Dim σETS As Single = 标准弹性模量 / IPS * (ηS * PI ^ 2 * IWS / lS ^ 2 + 0.385 * ITS)
                            Dim σELS As Single = 160000 * (tw / hw) ^ 2

                            Dim AE As Single = ASS + AP
                            Dim hCE As Single = ASS / AE * hCS + AP / AE * (-tP / 2)
                            Dim ICE As Single = ICw + Aw * (hw / 2 - hCE) ^ 2 + ICf + Af * (df - hCE) ^ 2 + ICP + AP * (-tP / 2) ^ 2
                            Dim σYE As Single = 加强筋单元_σY(一级序数)
                            Dim εYE As Single = σYE / 标准弹性模量

                            Dim εOE As Single = 二级序数 * εYE / 100
                            Dim εRE As Single = εOE / εYE
                            Dim Φε As Single
                            'Dim Φε As Single = If(εRE < -1, -1, If(εRE > 1, 1, εRE))
                            Select Case εRE
                                Case < -1
                                    Φε = -1
                                Case > 1.1
                                    Φε = 1
                                Case -1 To 0.9
                                    Φε = εRE
                                Case 0.9 To 1
                                    Φε = 0.9 + (εRE - 0.9) * 2 / 3
                                Case 1 To 1.1
                                    Φε = 1 - (1.1 - εRE) / 3
                            End Select
                            'σEPEε(一级序数, 二级序数) = Φε * σYE

                            Dim βPε As Single = βOP * Sqrt(εRE)
                            Dim wEPε As Single = If(βPε >= 1.25, (2.25 / βPε - 1.25 / βPε ^ 2) * wP, wP)
                            Dim wE1Pε As Single = If(βPε >= 1, wP / βPε, wP)
                            Dim AEPε As Single = wEPε * tP
                            Dim AE1Pε As Single = wE1Pε * tP
                            Dim ICE1Pε As Single = wE1Pε * tP ^ 3 / 12

                            Dim βwε As Single = βOw * Sqrt(εRE)
                            Dim hEwε As Single = If(βwε >= 1.25, (2.25 / βwε - 1.25 / βwε ^ 2) * hw, hw)
                            Dim AEwε As Single = hEwε * tw
                            Dim AESε As Single = AEwε + Af

                            Dim AEEε As Single = AEPε + ASS
                            Dim AE1Eε As Single = AE1Eε + ASS
                            Dim hCE1Eε As Single = AE1Pε / AE1Eε * (-tP / 2) + ASS / AE1Eε * hCS
                            Dim ICE1Eε As Single = ICE1Pε + AE1Pε * (-tP / 2) ^ 2 + ICS + ASS * (hCS - hCE1Eε) ^ 2
                            Dim lE1Pε As Single = hCE1Eε - (-tP / 2)
                            Dim lE1Sε As Single = If(tpS = "F" Or tpS = "B" Or tpS = "L3", hw - hCE1Eε, hw + tf - hCE1Eε)
                            Dim σYE1Eε As Single = ASS * lE1Sε / (ASS * lE1Sε + AE1Pε * lE1Pε) * σYS + AE1Pε * lE1Pε / (ASS * lE1Sε + AE1Pε * lE1Pε) * σYP
                            Dim σECEε As Single = PI ^ 2 * 标准弹性模量 * ICE1Eε / AEEε / lS ^ 2
                            Dim σCCEε As Single = If(σECEε <= σYE1Eε / 2 * εRE, σECEε / εRE, σYE1Eε * (1 - σYE1Eε * εRE / 4 / σECEε))
                            σBCEε(一级序数, 二级序数) = Φε * σCCEε * AEEε / AE

                            Dim σCTSε As Single = If(σETS <= σYS / 2 * εRE, σETS / εRE, σYS * (1 - σYS * εRE / 4 / σETS))
                            Dim σCPε As Single = If(βPε >= 1.25, (2.25 / βPε - 1.25 / βPε ^ 2) * σYP, σYP)
                            σFTEε(一级序数, 二级序数) = Φε * (ASS * σCTSε + AP * σCPε) / AE

                            Dim σCLSε As Single = If(σELS <= σYS / 2 * εRE, σELS / εRE, σYS * (1 - σYS * εRE / 4 / σELS))
                            σWBEε(一级序数, 二级序数) = If(tpS = "F", Φε * (ASS * σCLSε + AP * σCPε) / AE, Φε * (AESε * σYS + AEPε * σYP) / AE)

                            If 二级序数 = 0 Then
                                σMINEε(一级序数, 二级序数) = 0
                            Else
                                σMINEε(一级序数, 二级序数) = Min(σBCEε(一级序数, 二级序数), Min(σFTEε(一级序数, 二级序数), σWBEε(一级序数, 二级序数)))
                            End If
                            '输出σ = 输出σ & Format(σMINEε(一级序数, 二级序数), "000.000") & Space(4)
                        Next
                        For 二级序数 As UShort = 0 To 插值点_总数 - 1
                            Min_D(一级序数, 二级序数) = (σMINEε(一级序数, 二级序数 + 1) - σMINEε(一级序数, 二级序数)) * 100
                            加强筋单元_屈曲_D(一级序数, 二级序数) = Min_D(一级序数, 二级序数)

                            Select Case 二级序数
                                Case < 100
                                    加强筋单元_屈服_D(一级序数, 二级序数) = 加强筋单元_σY(一级序数)
                                Case >= 100
                                    加强筋单元_屈服_D(一级序数, 二级序数) = 0
                                Case Else

                            End Select
                            '输出D = 输出D & Format(加强筋单元_屈曲_D(一级序数, 二级序数), "0.000E+000") & Space(4)
                        Next
                        Debug.Print(Format(一级序数, "000") & Space(4) & 加强筋单元_屈曲_D(一级序数, 0))   'Debug.Print(输出σ)
                        'Debug.Print(输出D)
                    Next
                Case 5
                    'MsgBox("面板加强筋单元部分未完成!")
                Case 6
                    For 一级序数 As UShort = 1 To 加筋板单元_总数
                        Dim 输出 As String = "加筋板单元" & Space(8) & Format(一级序数, "000") & Space(4)
                        For 二级序数 As UShort = 0 To 插值点_总数
                            Dim σYE As Single = 加筋板单元_剩余_σY(一级序数)
                            Dim εYE As Single = σYE / 标准弹性模量
                            Dim εOE As Single = 二级序数 * εYE / 100
                            Dim εRE As Single = εOE / εYE
                            Dim Φε As Single = If(εRE < -1, -1, If(εRE > 1, 1, εRE))
                            Dim s As Single = 板格_l(加筋板单元_板格_序数(一级序数))
                            Dim l As Single = 加筋板单元_原始_w(一级序数)
                            Dim t As Single = 加筋板单元_原始_A(一级序数) / l
                            Dim βOP As Single = l / t * Sqrt(εYE)
                            Dim βPε As Single = βOP * Sqrt(εRE)

                            'σEPEε(一级序数, 二级序数) = Φε * σYE
                            If 二级序数 = 0 Then
                                σPBEε(一级序数, 二级序数) = 0
                            Else
                                σPBEε(一级序数, 二级序数) = σYE * Φε * Min(1, s / l * (2.25 / βPε - 1.25 / βPε ^ 2 + 0.1 * (1 - s / l) * (1 + 1 / βPε ^ 2) ^ 2))
                            End If
                        Next
                        For 二级序数 As UShort = 0 To 插值点_总数 - 1
                            Dim σYE As Single = 加筋板单元_剩余_σY(一级序数)
                            Dim εYE As Single = σYE / 标准弹性模量
                            加筋板单元_剩余_D(一级序数, 二级序数) = (σPBEε(一级序数, 二级序数 + 1) - σPBEε(一级序数, 二级序数)) / (εYE / 100)
                            输出 = 输出 & Format(加筋板单元_剩余_D(一级序数, 二级序数), "0.000E+000") & Space(4)
                        Next
                        Debug.Print(输出)
                    Next
                Case Else

            End Select
        Next
    End Sub

    Private Sub 增量迭代法(sender As Object, e As EventArgs) Handles Button5.Click
        Dim xlApp As Application
        Dim xlBook As Workbook
        Dim xlSheet As Worksheet

        OpenFileDialog1.ShowDialog()

        xlApp = CType(CreateObject("Excel.Application"), Application)
        xlBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName,, True)
        xlSheet = CType(xlBook.Worksheets(1), Worksheet)

        '调用 共有过程_通用型.读入界面参数
        读入界面参数()

        Dim 行号 As UShort
        For 基本输入对象类型序数 As UShort = 1 To 5
            Select Case 基本输入对象类型序数
                Case 1
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "节点"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z/(mm)"
                    For 一级序数 As UShort = 1 To 节点_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        节点_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        节点_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 1 / 4
                Case 2
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "加强筋"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    'xlSheet.Cells(行号, 4) = "l/(mm)"
                    'xlSheet.Cells(行号, 5) = "hw/(mm)"
                    'xlSheet.Cells(行号, 6) = "tw/(mm)"
                    'xlSheet.Cells(行号, 7) = "αw/(rad)"
                    'xlSheet.Cells(行号, 8) = "wf/(mm)"
                    'xlSheet.Cells(行号, 9) = "tf/(mm)"
                    'xlSheet.Cells(行号, 10) = "αf/(rad)"
                    'xlSheet.Cells(行号, 11) = "dx/(mm)"
                    'xlSheet.Cells(行号, 12) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 13) = "tp(F/B/T/L1/L2/L3)"
                    'xlSheet.Cells(行号, 14) = "mk(TRUE/FALSE)"
                    For 一级序数 As UShort = 1 To 加强筋_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        加强筋_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        加强筋_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                        加强筋_l(一级序数) = xlSheet.Cells(行号, 4).value
                        加强筋_hw(一级序数) = xlSheet.Cells(行号, 5).value
                        加强筋_tw(一级序数) = xlSheet.Cells(行号, 6).value
                        加强筋_αw(一级序数) = xlSheet.Cells(行号, 7).value
                        加强筋_wf(一级序数) = xlSheet.Cells(行号, 8).value
                        加强筋_tf(一级序数) = xlSheet.Cells(行号, 9).value
                        加强筋_αf(一级序数) = xlSheet.Cells(行号, 10).value
                        加强筋_dx(一级序数) = xlSheet.Cells(行号, 11).value
                        加强筋_σY(一级序数) = xlSheet.Cells(行号, 12).value
                        加强筋_tp(一级序数) = xlSheet.Cells(行号, 13).value
                        加强筋_mk(一级序数) = xlSheet.Cells(行号, 14).value
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 2 / 4
                Case 3
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "面板"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    'xlSheet.Cells(行号, 4) = "Y1/(mm)"
                    'xlSheet.Cells(行号, 5) = "Z1/(mm)"
                    'xlSheet.Cells(行号, 6) = "YL/(mm)"
                    'xlSheet.Cells(行号, 7) = "ZL/(mm)"
                    'xlSheet.Cells(行号, 8) = "l/(mm)"
                    'xlSheet.Cells(行号, 9) = "t/(mm)"
                    'xlSheet.Cells(行号, 10) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 11) = "PMA(TRUE/FALSE)"
                    For 一级序数 As UShort = 1 To 面板_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        面板_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        面板_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                        面板_Y1(一级序数) = xlSheet.Cells(行号, 4).value
                        面板_Z1(一级序数) = xlSheet.Cells(行号, 5).value
                        面板_YL(一级序数) = xlSheet.Cells(行号, 6).value
                        面板_ZL(一级序数) = xlSheet.Cells(行号, 7).value
                        面板_l(一级序数) = xlSheet.Cells(行号, 8).value
                        面板_t(一级序数) = xlSheet.Cells(行号, 9).value
                        面板_σY(一级序数) = xlSheet.Cells(行号, 10).value
                        面板_PMA(一级序数) = xlSheet.Cells(行号, 11).value
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 3 / 4
                Case 4
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "板格"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    'xlSheet.Cells(行号, 4) = "Y1/(mm)"
                    'xlSheet.Cells(行号, 5) = "Z1/(mm)"
                    'xlSheet.Cells(行号, 6) = "l(mm)"
                    'xlSheet.Cells(行号, 7) = "tp(L/T)"
                    'xlSheet.Cells(行号, 8) = "板格板数目"
                    'xlSheet.Cells(行号, 9) = "w/(mm)"
                    'xlSheet.Cells(行号, 10) = "t/(mm)"
                    'xlSheet.Cells(行号, 11) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 12) = "w/(mm)"
                    'xlSheet.Cells(行号, 13) = "t/(mm)"
                    'xlSheet.Cells(行号, 14) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 15) = "......"
                    'xlSheet.Cells(行号, 16) = "......"
                    'xlSheet.Cells(行号, 17) = "......"
                    For 一级序数 As UShort = 1 To 板格_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        板格_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        板格_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                        板格_Y1(一级序数) = xlSheet.Cells(行号, 4).value
                        板格_Z1(一级序数) = xlSheet.Cells(行号, 5).value
                        板格_l(一级序数) = xlSheet.Cells(行号, 6).value
                        板格_tp(一级序数) = xlSheet.Cells(行号, 7).value
                        板格_板格板数目(一级序数) = xlSheet.Cells(行号, 8).value
                        For 二级序数 As UShort = 1 To 板格_板格板数目(一级序数)
                            板格板_w(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 9).value
                            板格板_t(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 10).value
                            板格板_σY(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 11).value
                        Next
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 4 / 4
                Case 5
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "特别硬角单元"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "子对象数目"
                    'xlSheet.Cells(行号, 3) = "Yc/(mm)"
                    'xlSheet.Cells(行号, 4) = "Zc/(mm)"
                    'xlSheet.Cells(行号, 5) = "A/(mm^2)"
                    'xlSheet.Cells(行号, 6) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 7) = "Yc/(mm)"
                    'xlSheet.Cells(行号, 8) = "Zc/(mm)"
                    'xlSheet.Cells(行号, 9) = "A/(mm^2)"
                    'xlSheet.Cells(行号, 10) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 11) = "......"
                    'xlSheet.Cells(行号, 12) = "......"
                    'xlSheet.Cells(行号, 13) = "......"
                    'xlSheet.Cells(行号, 14) = "......"
                    For 一级序数 As UShort = 1 To 特别硬角单元_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        子对象_数目(一级序数) = xlSheet.Cells(行号, 2).value
                        For 二级序数 As UShort = 1 To 子对象_数目(一级序数)
                            子对象_Yc(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 3).value
                            子对象_Zc(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 4).value
                            子对象_A(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 5).value
                            子对象_σY(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 6).value
                        Next
                    Next
            End Select
        Next

        xlBook.Close()
        xlApp.Quit() 'xlApp = Nothing

        '基本输入对象的属性计算及形心坐标输出

        For 基本输入对象类型序数 As UShort = 1 To 5
            Select Case 基本输入对象类型序数
                Case 1
                    Chart1.Series.Add("节点_形心")
                    Chart1.Series("节点_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 节点_总数
                        Chart1.Series("节点_形心").Points.AddXY(节点_Y0(一级序数), 节点_Z0(一级序数))
                    Next
                Case 2
                    Chart1.Series.Add("加强筋_形心")
                    Chart1.Series("加强筋_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 加强筋_总数
                        加强筋_Aw(一级序数) = 加强筋_hw(一级序数) * 加强筋_tw(一级序数)
                        加强筋_Icw(一级序数) = 加强筋_tw(一级序数) * 加强筋_hw(一级序数) ^ 3 / 12
                        加强筋_Iow(一级序数) = 加强筋_Icw(一级序数) + 加强筋_Aw(一级序数) * (加强筋_hw(一级序数) / 2) ^ 2

                        Select Case 加强筋_tp(一级序数)
                            Case "F"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数)
                            Case "B"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) - 加强筋_tf(一级序数) / 2
                            Case "T"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                            Case "L1"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                            Case "L2"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                            Case "L3"
                                加强筋_df(一级序数) = 加强筋_hw(一级序数) - 加强筋_dx(一级序数) - 加强筋_tf(一级序数) / 2
                            Case Else
                                MsgBox("加强筋_tp(" & 一级序数 & ")错误：类型不符")
                        End Select
                        加强筋_Af(一级序数) = 加强筋_wf(一级序数) * 加强筋_tf(一级序数)
                        加强筋_Icf(一级序数) = 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 12
                        加强筋_Iof(一级序数) = 加强筋_Icf(一级序数) + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2

                        加强筋_AS(一级序数) = 加强筋_Aw(一级序数) + 加强筋_Af(一级序数)
                        加强筋_hcS(一级序数) = 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_hw(一级序数) / 2 + 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_df(一级序数)
                        加强筋_IoS(一级序数) = 加强筋_Iow(一级序数) + 加强筋_Iof(一级序数)
                        加强筋_IcS(一级序数) = 加强筋_Icw(一级序数) + 加强筋_Aw(一级序数) * (加强筋_hw(一级序数) / 2 - 加强筋_hcS(一级序数)) ^ 2 + 加强筋_Icf(一级序数) + 加强筋_Af(一级序数) * (加强筋_df(一级序数) - 加强筋_hcS(一级序数)) ^ 2

                        Select Case 加强筋_tp(一级序数)
                            Case "F"
                                加强筋_IPS(一级序数) = 加强筋_hw(一级序数) ^ 3 * 加强筋_tw(一级序数) / 3
                                加强筋_ITS(一级序数) = 加强筋_hw(一级序数) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / 加强筋_hw(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_hw(一级序数) ^ 3 * 加强筋_tw(一级序数) ^ 3 / 36
                            Case "B"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                            Case "T"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_wf(一级序数) ^ 3 * 加强筋_tf(一级序数) * 加强筋_df(一级序数) ^ 2 / 12
                            Case "L1"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                            Case "L2"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                            Case "L3"
                                加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                            Case Else

                        End Select
                        '加强筋_ηS(一级序数) = 1 + 加强筋_l(一级序数) ^ 2 / PI ^ 2 / Sqrt(加强筋_IWS(一级序数) * (0.75 * 加强筋带板_w(一级序数) / 加强筋带板_t(一级序数) ^ 3 + (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) / 加强筋_tw(一级序数) ^ 3))
                        '加强筋_σETS(一级序数) = 标准弹性模量 / 加强筋_IPS(一级序数) * (加强筋_ηS(一级序数) * PI ^ 2 * 加强筋_IWS(一级序数) / 加强筋_lS(一级序数) ^ 2 + 0.385 * 加强筋_ITS(一级序数))
                        加强筋_σELS(一级序数) = 160000 * (加强筋_tw(一级序数) / 加强筋_hw(一级序数)) ^ 2

                        加强筋_Ycw(一级序数) = 加强筋_Y0(一级序数) + 加强筋_hw(一级序数) * Cos(加强筋_αw(一级序数)) / 2
                        加强筋_Zcw(一级序数) = 加强筋_Z0(一级序数) + 加强筋_hw(一级序数) * Sin(加强筋_αw(一级序数)) / 2

                        加强筋_Ycf(一级序数) = 加强筋_Y0(一级序数) + 加强筋_df(一级序数) * Cos(加强筋_αw(一级序数)) + 加强筋_wf(一级序数) * Cos(加强筋_αf(一级序数)) / 2
                        加强筋_Zcf(一级序数) = 加强筋_Z0(一级序数) + 加强筋_df(一级序数) * Sin(加强筋_αw(一级序数)) + 加强筋_wf(一级序数) * Sin(加强筋_αf(一级序数)) / 2

                        加强筋_YcS(一级序数) = 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_Ycf(一级序数) + 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_Ycw(一级序数)
                        加强筋_ZcS(一级序数) = 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_Zcf(一级序数) + 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_Zcw(一级序数)

                        Chart1.Series("加强筋_形心").Points.AddXY(加强筋_YcS(一级序数), 加强筋_ZcS(一级序数))
                    Next
                Case 3
                    Chart1.Series.Add("面板_形心")
                    Chart1.Series("面板_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 面板_总数
                        面板_w(一级序数) = Sqrt((面板_Y0(一级序数) - 面板_Y1(一级序数)) ^ 2 + (面板_Z0(一级序数) - 面板_Z1(一级序数)) ^ 2)
                        面板_A(一级序数) = 面板_w(一级序数) * 面板_t(一级序数)

                        面板_Yc(一级序数) = (面板_Y0(一级序数) + 面板_Y1(一级序数)) / 2
                        面板_Zc(一级序数) = (面板_Z0(一级序数) + 面板_Z1(一级序数)) / 2

                        Chart1.Series("面板_形心").Points.AddXY(面板_Yc(一级序数), 面板_Zc(一级序数))
                        Chart1.Series("面板_形心").Points.AddXY(面板_Y0(一级序数), 面板_Z0(一级序数))
                        Chart1.Series("面板_形心").Points.AddXY(面板_Y1(一级序数), 面板_Z1(一级序数))
                    Next
                Case 4
                    Chart1.Series.Add("板格_形心")
                    Chart1.Series("板格_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 板格_总数
                        板格_α(一级序数) = If(板格_Y0(一级序数) = 板格_Y1(一级序数), PI / 2, Atan((板格_Z0(一级序数) - 板格_Z1(一级序数)) / (板格_Y0(一级序数) - 板格_Y1(一级序数))))
                        板格_w(一级序数) = Sqrt((板格_Y0(一级序数) - 板格_Y1(一级序数)) ^ 2 + (板格_Z0(一级序数) - 板格_Z1(一级序数)) ^ 2)
                        For 二级序数 As UShort = 1 To 板格_板格板数目(一级序数)
                            板格板_Y0(一级序数, 二级序数) = If(二级序数 = 1, 板格_Y0(一级序数), 板格板_Y0(一级序数, 二级序数 - 1))
                            板格板_Z0(一级序数, 二级序数) = If(二级序数 = 1, 板格_Z0(一级序数), 板格板_Z0(一级序数, 二级序数 - 1))

                            板格板_Y1(一级序数, 二级序数) = If(二级序数 = 板格_板格板数目(一级序数), 板格_Y1(一级序数), 板格板_Y0(一级序数, 二级序数) + 板格板_w(一级序数, 二级序数) * Cos(板格_α(一级序数)))
                            板格板_Z1(一级序数, 二级序数) = If(二级序数 = 板格_板格板数目(一级序数), 板格_Z1(一级序数), 板格板_Z0(一级序数, 二级序数) + 板格板_w(一级序数, 二级序数) * Sin(板格_α(一级序数)))

                            板格板_Yc(一级序数, 二级序数) = (板格板_Y0(一级序数, 二级序数) + 板格板_Y1(一级序数, 二级序数)) / 2
                            板格板_Zc(一级序数, 二级序数) = (板格板_Z0(一级序数, 二级序数) + 板格板_Z1(一级序数, 二级序数)) / 2

                            板格板_A(一级序数, 二级序数) = 板格板_w(一级序数, 二级序数) * 板格板_t(一级序数, 二级序数)
                            板格板_AYc(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_Yc(一级序数, 二级序数)
                            板格板_AZc(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_Zc(一级序数, 二级序数)
                            板格板_AσY(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_σY(一级序数, 二级序数)

                            板格_A(一级序数) += 板格板_A(一级序数, 二级序数)
                            板格_AYc(一级序数) += 板格板_AYc(一级序数, 二级序数)
                            板格_AZc(一级序数) += 板格板_AZc(一级序数, 二级序数)
                            板格_AσY(一级序数) += 板格板_AσY(一级序数, 二级序数)
                        Next
                        板格_t(一级序数) = 板格_A(一级序数) / 板格_w(一级序数)
                        板格_Yc(一级序数) = 板格_AYc(一级序数) / 板格_A(一级序数)
                        板格_Zc(一级序数) = 板格_AZc(一级序数) / 板格_A(一级序数)
                        板格_σY(一级序数) = 板格_AσY(一级序数) / 板格_A(一级序数)

                        Chart1.Series("板格_形心").Points.AddXY(板格_Yc(一级序数), 板格_Zc(一级序数))
                    Next
                Case 5
                    Chart1.Series.Add("特别硬角单元_形心")
                    Chart1.Series("特别硬角单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 特别硬角单元_总数
                        For 二级序数 As UShort = 1 To 子对象_数目(一级序数)
                            子对象_AYc(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_Yc(一级序数, 二级序数)
                            子对象_AZc(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_Zc(一级序数, 二级序数)
                            子对象_AσY(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_σY(一级序数, 二级序数)

                            特别硬角单元_A(一级序数) += 子对象_A(一级序数, 二级序数)
                            特别硬角单元_AYc(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_Yc(一级序数, 二级序数)
                            特别硬角单元_AZc(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_Zc(一级序数, 二级序数)
                            特别硬角单元_AσY(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_σY(一级序数, 二级序数)
                        Next
                        特别硬角单元_Yc(一级序数) += 特别硬角单元_AYc(一级序数) / 特别硬角单元_A(一级序数)
                        特别硬角单元_Zc(一级序数) += 特别硬角单元_AZc(一级序数) / 特别硬角单元_A(一级序数)
                        特别硬角单元_σY(一级序数) += 特别硬角单元_AσY(一级序数) / 特别硬角单元_A(一级序数)

                        Chart1.Series("特别硬角单元_形心").Points.AddXY(特别硬角单元_Yc(一级序数), 特别硬角单元_Zc(一级序数))
                    Next
            End Select
        Next

        For 单元划分步骤序数 As UShort = 1 To 4
            Select Case 单元划分步骤序数
                Case 1      '板格-节点配对, 成立节点-分支
                    For 一级序数 As UShort = 1 To 板格_总数
                        For 二级序数 As UShort = 1 To 节点_总数
                            If 板格_Y0(一级序数) = 节点_Y0(二级序数) And 板格_Z0(一级序数) = 节点_Z0(二级序数) Then
                                节点_分支_数目(二级序数) += 1
                                节点_分支_板格_序数(二级序数, 节点_分支_数目(二级序数)) = 一级序数

                                节点_分支_α(二级序数, 节点_分支_数目(二级序数)) = 板格_α(一级序数)

                                板格_首端节点_序数(一级序数) = 二级序数
                                板格_首端节点_分支_序数(一级序数) = 节点_分支_数目(二级序数)
                            ElseIf 板格_Y1(一级序数) = 节点_Y0(二级序数) And 板格_Z1(一级序数) = 节点_Z0(二级序数) Then
                                节点_分支_数目(二级序数) += 1
                                节点_分支_板格_序数(二级序数, 节点_分支_数目(二级序数)) = 一级序数

                                节点_分支_α(二级序数, 节点_分支_数目(二级序数)) = If(板格_α(一级序数) <= 0, 板格_α(一级序数) + PI, 板格_α(一级序数) - PI)

                                板格_末端节点_序数(一级序数) = 二级序数
                                板格_末端节点_分支_序数(一级序数) = 节点_分支_数目(二级序数)
                            End If
                            If (Not 板格_首端节点_序数(一级序数) = 0) And (Not 板格_末端节点_序数(一级序数) = 0) Then
                                '节点_分支_首端节点_序数(二级序数) = 二级序数
                                节点_分支_首端节点_序数(板格_首端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_首端节点_序数(一级序数)
                                节点_分支_末端节点_序数(板格_首端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_末端节点_序数(一级序数)
                                节点_分支_首端节点_序数(板格_末端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_末端节点_序数(一级序数)
                                节点_分支_末端节点_序数(板格_末端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_首端节点_序数(一级序数)
                                '节点_分支_末端节点_序数(二级序数) = 二级序数
                                Exit For
                            End If
                        Next
                    Next
                Case 2      '节点-加强筋/面板配对
                    For 一级序数 As UShort = 1 To 节点_总数
                        For 二级序数 As UShort = 1 To 加强筋_总数
                            If 节点_Y0(一级序数) = 加强筋_Y0(二级序数) And 节点_Z0(一级序数) = 加强筋_Z0(二级序数) Then
                                节点_tp(一级序数) = "加强筋"

                                节点_加强筋_序数(一级序数) = 二级序数
                                加强筋_节点_序数(二级序数) = 一级序数

                                '属性继承：加强筋 → 节点_加强筋
                                Exit For
                            End If
                        Next

                        For 二级序数 As UShort = 1 To 面板_总数
                            If 节点_Y0(一级序数) = 面板_YL(二级序数) And 节点_Z0(一级序数) = 面板_ZL(二级序数) Then
                                节点_tp(一级序数) = "面板"

                                节点_面板_序数(一级序数) = 二级序数
                                面板_节点_序数(二级序数) = 一级序数

                                '属性继承：面板 → 节点_面板
                                Exit For
                            End If
                        Next
                    Next
                Case 3      '根据节点分支数目确定单元类型
                    For 一级序数 As UShort = 1 To 节点_总数
                        Select Case 节点_分支_数目(一级序数)
                            Case >= 3
                                节点_tp(一级序数) = "硬角单元"
                                硬角单元_总数 += 1

                                节点_硬角单元_序数(一级序数) = 硬角单元_总数
                                ReDim Preserve 硬角单元_节点_序数(硬角单元_总数)
                                硬角单元_节点_序数(硬角单元_总数) = 一级序数

                                Exit Select
                            Case 2
                                Dim 双分支夹角 As Single = 节点_分支_α(一级序数, 1) - 节点_分支_α(一级序数, 2)
                                Select Case 双分支夹角
                                    Case >= 2 * PI
                                        双分支夹角 -= 2 * PI
                                    Case < 0
                                        双分支夹角 += 2 * PI
                                    Case Else

                                End Select
                                Select Case 双分支夹角
                                    Case <= 5 / 6 * PI
                                        节点_tp(一级序数) = "硬角单元"
                                        硬角单元_总数 += 1

                                        节点_硬角单元_序数(一级序数) = 硬角单元_总数
                                        ReDim Preserve 硬角单元_节点_序数(硬角单元_总数)
                                        硬角单元_节点_序数(硬角单元_总数) = 一级序数

                                        Exit Select
                                    Case >= 7 / 6 * PI
                                        节点_tp(一级序数) = "硬角单元"
                                        硬角单元_总数 += 1

                                        节点_硬角单元_序数(一级序数) = 硬角单元_总数
                                        ReDim Preserve 硬角单元_节点_序数(硬角单元_总数)
                                        硬角单元_节点_序数(硬角单元_总数) = 一级序数

                                        Exit Select
                                    Case Else
                                        Select Case 节点_tp(一级序数)
                                            Case "加强筋"
                                                节点_tp(一级序数) = "加强筋单元"
                                                加强筋单元_总数 += 1

                                                节点_加强筋单元_序数(一级序数) = 加强筋单元_总数
                                                ReDim Preserve 加强筋单元_节点_序数(加强筋单元_总数)
                                                加强筋单元_节点_序数(加强筋单元_总数) = 一级序数

                                                Exit Select
                                            Case Else

                                        End Select
                                End Select
                            Case 1
                                Select Case 节点_tp(一级序数)
                                    Case "加强筋"
                                        节点_tp(一级序数) = "加强筋单元"
                                        加强筋单元_总数 += 1

                                        节点_加强筋单元_序数(一级序数) = 加强筋单元_总数
                                        ReDim Preserve 加强筋单元_节点_序数(加强筋单元_总数)
                                        加强筋单元_节点_序数(加强筋单元_总数) = 一级序数

                                        Exit Select
                                    Case "面板"
                                        Select Case 面板_PMA(节点_面板_序数(一级序数))
                                            Case True
                                                节点_tp(一级序数) = "面板加强筋单元"
                                                面板加强筋单元_总数 += 1

                                                节点_面板加强筋单元_序数(一级序数) = 面板加强筋单元_总数
                                                ReDim Preserve 面板加强筋单元_节点_序数(面板加强筋单元_总数)
                                                面板加强筋单元_节点_序数(面板加强筋单元_总数) = 一级序数

                                                Exit Select
                                            Case False
                                                节点_tp(一级序数) = "面板硬角单元"
                                                面板硬角单元_总数 += 1

                                                节点_面板硬角单元_序数(一级序数) = 面板硬角单元_总数
                                                ReDim Preserve 面板硬角单元_节点_序数(面板硬角单元_总数)
                                                面板硬角单元_节点_序数(面板硬角单元_总数) = 一级序数

                                                Exit Select
                                        End Select
                                    Case Else
                                        节点_tp(一级序数) = "自由端"

                                        Exit Select
                                End Select
                        End Select
                    Next
                Case 4      '确定单元属性
                    ReDim 硬角单元_A(硬角单元_总数), 硬角单元_Yc(硬角单元_总数), 硬角单元_Zc(硬角单元_总数), 硬角单元_σY(硬角单元_总数)
                    ReDim 加强筋单元_A(加强筋单元_总数), 加强筋单元_Yc(加强筋单元_总数), 加强筋单元_Zc(加强筋单元_总数), 加强筋单元_σY(加强筋单元_总数)
                    For 单元对象类型序数 As UShort = 1 To 6
                        Dim 通用单元数目 As UShort = 加强筋单元_总数
                        Dim 单元分支_w(通用单元数目, 通用分支数目) As Single,
                            单元分支_A(通用单元数目, 通用分支数目) As Single,
                            单元分支_AYc(通用单元数目, 通用分支数目) As Single, 单元分支_AZc(通用单元数目, 通用分支数目) As Single,
                            单元分支_AσY(通用单元数目, 通用分支数目) As Single

                        Select Case 单元对象类型序数
                            Case 1      '硬角单元
                                For 一级序数 As UShort = 1 To 硬角单元_总数
                                    Dim 硬角单元_AYc(硬角单元_总数) As Single, 硬角单元_AZc(硬角单元_总数) As Single,
                                        硬角单元_AσY(硬角单元_总数) As Single

                                    Dim 关联节点序数 As UShort = 硬角单元_节点_序数(一级序数)
                                    For 二级序数 As UShort = 1 To 节点_分支_数目(关联节点序数)
                                        Dim 关联板格序数 As UShort = 节点_分支_板格_序数(关联节点序数, 二级序数)

                                        Select Case 板格_tp(关联板格序数)
                                            Case "L"
                                                单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数) / 2
                                                If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 二级序数)) = "自由端" Then
                                                    单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数)
                                                End If
                                            Case "T"
                                                单元分支_w(一级序数, 二级序数) = 板格_t(关联板格序数) * 20
                                            Case Else

                                        End Select
                                        Dim 分支板_A(硬角单元_总数, 通用分支数目) As Single,
                                            分支板_AYc(硬角单元_总数, 通用分支数目) As Single, 分支板_AZc(硬角单元_总数, 通用分支数目) As Single,
                                            分支板_AσY(硬角单元_总数, 通用分支数目) As Single
                                        For 三级序数 As UShort = 1 To 板格_板格板数目(关联板格序数)
                                            If 关联节点序数 = 板格_首端节点_序数(关联板格序数) Then
                                                板格_首端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                Dim 分支板_w(硬角单元_总数, 通用分支数目) As Single
                                                分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 三级序数)
                                                Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                Select Case 分支板超出宽度
                                                    Case <= 0   '分支板宽度和小于分支所需宽度
                                                        分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 三级序数)
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 三级序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 三级序数)
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 三级序数)
                                                    Case > 0    '分支板宽度和大于分支所需宽度
                                                        Dim 通用板格板数目 As UShort = 3
                                                        Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                        '板格板实际取用宽度
                                                        板格板_w1(关联板格序数, 三级序数) = 板格板_w(关联板格序数, 三级序数) - 分支板超出宽度
                                                        '板格板宽度利用系数
                                                        Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 三级序数) / 板格板_w(关联板格序数, 三级序数)
                                                        Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y11(板格_总数, 通用板格板数目), 板格板_Z11(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                        '板格板实际取用面积
                                                        板格板_A1(关联板格序数, 三级序数) = 板格板_A(关联板格序数, 三级序数) * 板格板_ηw
                                                        '板格板实际末端坐标
                                                        板格板_Y11(关联板格序数, 三级序数) = 板格板_Y0(关联板格序数, 三级序数) + (板格板_Y1(关联板格序数, 三级序数) - 板格板_Y0(关联板格序数, 三级序数)) * 板格板_ηw
                                                        板格板_Z11(关联板格序数, 三级序数) = 板格板_Z0(关联板格序数, 三级序数) + (板格板_Z1(关联板格序数, 三级序数) - 板格板_Z0(关联板格序数, 三级序数)) * 板格板_ηw
                                                        '板格板实际形心坐标
                                                        板格板_Yc1(关联板格序数, 三级序数) = (板格板_Y0(关联板格序数, 三级序数) + 板格板_Y11(关联板格序数, 三级序数)） / 2
                                                        板格板_Zc1(关联板格序数, 三级序数) = (板格板_Z0(关联板格序数, 三级序数) + 板格板_Z11(关联板格序数, 三级序数)） / 2
                                                        '板格板实际面积坐标积数
                                                        板格板_AYc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Yc1(关联板格序数, 三级序数)
                                                        板格板_AZc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Zc1(关联板格序数, 三级序数)
                                                        '板格板实际面积强度积数
                                                        板格板_AσY1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_σY(关联板格序数, 三级序数)

                                                        'Select Case 三级序数
                                                        '    Case = 1

                                                        '    Case > 1

                                                        'End Select

                                                        '分支板实际面积
                                                        分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 三级序数)
                                                        '分支板实际面积坐标积数
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 三级序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 三级序数)
                                                        '分支板实际面积强度积数
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 三级序数)

                                                        Exit For
                                                End Select
                                            ElseIf 关联节点序数 = 板格_末端节点_序数(关联板格序数) Then
                                                板格_末端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                Dim 分支板_w(硬角单元_总数, 通用分支数目) As Single
                                                Dim 反向序数 As UShort = 板格_板格板数目(关联板格序数) + 1 - 三级序数
                                                分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 反向序数)
                                                Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                Select Case 分支板超出宽度
                                                    Case <= 0   '分支板宽度和小于分支所需宽度
                                                        分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 反向序数)
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 反向序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 反向序数)
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 反向序数)
                                                    Case > 0    '分支板宽度和大于分支所需宽度
                                                        Dim 通用板格板数目 As UShort = 3
                                                        Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                        '板格板实际取用宽度
                                                        板格板_w1(关联板格序数, 反向序数) = 板格板_w(关联板格序数, 反向序数) - 分支板超出宽度
                                                        '板格板宽度利用系数
                                                        Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 反向序数) / 板格板_w(关联板格序数, 反向序数)
                                                        Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y01(板格_总数, 通用板格板数目), 板格板_Z01(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                        '板格板实际取用面积
                                                        板格板_A1(关联板格序数, 反向序数) = 板格板_A(关联板格序数, 反向序数) * 板格板_ηw
                                                        '板格板实际末端坐标
                                                        板格板_Y01(关联板格序数, 反向序数) = 板格板_Y1(关联板格序数, 反向序数) + (板格板_Y0(关联板格序数, 反向序数) - 板格板_Y1(关联板格序数, 反向序数)) * 板格板_ηw
                                                        板格板_Z01(关联板格序数, 反向序数) = 板格板_Z1(关联板格序数, 反向序数) + (板格板_Z0(关联板格序数, 反向序数) - 板格板_Z1(关联板格序数, 反向序数)) * 板格板_ηw
                                                        '板格板实际形心坐标
                                                        板格板_Yc1(关联板格序数, 反向序数) = (板格板_Y01(关联板格序数, 反向序数) + 板格板_Y1(关联板格序数, 反向序数)） / 2
                                                        板格板_Zc1(关联板格序数, 反向序数) = (板格板_Z01(关联板格序数, 反向序数) + 板格板_Z1(关联板格序数, 反向序数)） / 2
                                                        '板格板实际面积坐标积数
                                                        板格板_AYc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Yc1(关联板格序数, 反向序数)
                                                        板格板_AZc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Zc1(关联板格序数, 反向序数)
                                                        '板格板实际面积强度积数
                                                        板格板_AσY1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_σY(关联板格序数, 反向序数)

                                                        'Select Case 三级序数
                                                        '    Case = 1

                                                        '    Case > 1

                                                        'End Select

                                                        '分支板实际面积
                                                        分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 反向序数)
                                                        '分支板实际面积坐标积数
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 反向序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 反向序数)
                                                        '分支板实际面积强度积数
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 反向序数)

                                                        Exit For
                                                End Select
                                            End If
                                        Next
                                        单元分支_A(一级序数, 二级序数) = 分支板_A(一级序数, 二级序数)
                                        单元分支_AYc(一级序数, 二级序数) = 分支板_AYc(一级序数, 二级序数)
                                        单元分支_AZc(一级序数, 二级序数) = 分支板_AZc(一级序数, 二级序数)
                                        单元分支_AσY(一级序数, 二级序数) = 分支板_AσY(一级序数, 二级序数)

                                        硬角单元_A(一级序数) += 分支板_A(一级序数, 二级序数)
                                        硬角单元_AYc(一级序数) += 分支板_AYc(一级序数, 二级序数)
                                        硬角单元_AZc(一级序数) += 分支板_AZc(一级序数, 二级序数)
                                        硬角单元_AσY(一级序数) += 分支板_AσY(一级序数, 二级序数)
                                    Next

                                    硬角单元_Yc(一级序数) = 硬角单元_AYc(一级序数) / 硬角单元_A(一级序数)
                                    硬角单元_Zc(一级序数) = 硬角单元_AZc(一级序数) / 硬角单元_A(一级序数)
                                    硬角单元_σY(一级序数) = 硬角单元_AσY(一级序数) / 硬角单元_A(一级序数)

                                    全截面_A += 硬角单元_A(一级序数)
                                    全截面_AYc += 硬角单元_AYc(一级序数)
                                    全截面_AZc += 硬角单元_AZc(一级序数)
                                    全截面_AσY += 硬角单元_AσY(一级序数)
                                Next
                            Case 2      '面板硬角单元
                                'MsgBox("面板硬角单元部分未完成!")
                            Case 3      '特别硬角单元
                                For 一级序数 As UShort = 1 To 特别硬角单元_总数
                                    全截面_A += 特别硬角单元_A(一级序数)
                                    全截面_AYc += 特别硬角单元_AYc(一级序数)
                                    全截面_AZc += 特别硬角单元_AZc(一级序数)
                                    全截面_AσY += 特别硬角单元_AσY(一级序数)
                                Next
                            Case 4      '加强筋单元
                                ReDim 加强筋单元_σYP(加强筋单元_总数),
                                    加强筋单元_lP(加强筋单元_总数),
                                    加强筋单元_wP(加强筋单元_总数), 加强筋单元_tP(加强筋单元_总数),
                                    加强筋单元_σYS(加强筋单元_总数),
                                    加强筋单元_lS(加强筋单元_总数),
                                    加强筋单元_hw(加强筋单元_总数), 加强筋单元_tw(加强筋单元_总数),
                                    加强筋单元_wf(加强筋单元_总数), 加强筋单元_tf(加强筋单元_总数),
                                    加强筋单元_dx(加强筋单元_总数),
                                    加强筋单元_tpS(加强筋单元_总数), 加强筋单元_mk(加强筋单元_总数)

                                ReDim 加强筋单元_ηS(加强筋单元_总数), 加强筋单元_σETS(加强筋单元_总数)

                                ReDim 加强筋单元_σELS(加强筋单元_总数)

                                Dim 加强筋单元_AYc(加强筋单元_总数) As Single, 加强筋单元_AZc(加强筋单元_总数) As Single,
                                    加强筋单元_AσY(加强筋单元_总数) As Single

                                For 一级序数 As UShort = 1 To 加强筋单元_总数
                                    Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(一级序数)
                                    Select Case 节点_分支_数目(关联节点序数)
                                        Case 1
                                            Select Case 板格_tp(节点_分支_板格_序数(关联节点序数, 1))
                                                Case "L"
                                                    单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 1)) = "自由端" Then
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1))
                                                    End If
                                                    加强筋单元_lP(一级序数) = 板格_l(节点_分支_板格_序数(关联节点序数, 1))
                                                    加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1)
                                                    'MsgBox("警告:单L分支")
                                                Case "T"
                                                    'MsgBox("错误:单T分支")
                                                Case Else

                                            End Select
                                        Case 2
                                            Dim 板格_联合_tp As String = 板格_tp(节点_分支_板格_序数(关联节点序数, 1)) & 板格_tp(节点_分支_板格_序数(关联节点序数, 2))
                                            Select Case 板格_联合_tp
                                                Case "LL"
                                                    单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                    单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2)) / 2
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 1)) = "自由端" Then
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1))
                                                    End If
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 2)) = "自由端" Then
                                                        单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2))
                                                    End If
                                                    加强筋单元_lP(一级序数) = (板格_l(节点_分支_板格_序数(关联节点序数, 1)) + 板格_l(节点_分支_板格_序数(关联节点序数, 2))) / 2
                                                    加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1) + 单元分支_w(一级序数, 2)
                                                Case "LT"
                                                    单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 1)) = "自由端" Then
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1))
                                                    End If
                                                    单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                    加强筋单元_lP(一级序数) = (板格_l(节点_分支_板格_序数(关联节点序数, 1)) + 板格_l(节点_分支_板格_序数(关联节点序数, 2))) / 2
                                                    加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1) + 单元分支_w(一级序数, 2)
                                                    'MsgBox("警告:单L单T分支")
                                                Case "TL"
                                                    单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 2)) / 2
                                                    单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2)) / 2
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 2)) = "自由端" Then
                                                        单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2))
                                                    End If
                                                    加强筋单元_lP(一级序数) = (板格_l(节点_分支_板格_序数(关联节点序数, 1)) + 板格_l(节点_分支_板格_序数(关联节点序数, 2))) / 2
                                                    加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1) + 单元分支_w(一级序数, 2)
                                                    'MsgBox("警告:单T单L分支")
                                                Case "TT"
                                                    'MsgBox("错误:双T分支")
                                                Case Else

                                            End Select
                                    End Select

                                    Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)
                                    加强筋单元_A(一级序数) += 加强筋_AS(关联加强筋序数)
                                    加强筋单元_AYc(一级序数) += 加强筋_AS(关联加强筋序数) * 加强筋_YcS(关联加强筋序数)
                                    加强筋单元_AZc(一级序数) += 加强筋_AS(关联加强筋序数) * 加强筋_ZcS(关联加强筋序数)
                                    加强筋单元_AσY(一级序数) += 加强筋_AS(关联加强筋序数) * 加强筋_σY(关联加强筋序数)

                                    加强筋单元_σYS(一级序数) = 加强筋_σY(关联加强筋序数)

                                    加强筋单元_lS(一级序数) = 加强筋_l(关联加强筋序数)

                                    加强筋单元_hw(一级序数) = 加强筋_hw(关联加强筋序数)
                                    加强筋单元_tw(一级序数) = 加强筋_tw(关联加强筋序数)

                                    加强筋单元_wf(一级序数) = 加强筋_wf(关联加强筋序数)
                                    加强筋单元_tf(一级序数) = 加强筋_tf(关联加强筋序数)

                                    加强筋单元_dx(一级序数) = 加强筋_dx(关联加强筋序数)

                                    加强筋单元_tpS(一级序数) = 加强筋_tp(关联加强筋序数)
                                    加强筋单元_mk(一级序数) = 加强筋_mk(关联加强筋序数)

                                    加强筋单元_σELS(一级序数) = 加强筋_σELS(关联加强筋序数)

                                    For 二级序数 As UShort = 1 To 节点_分支_数目(关联节点序数)
                                        Dim 关联板格序数 As UShort = 节点_分支_板格_序数(关联节点序数, 二级序数)

                                        'Select Case 板格_tp(关联板格序数)
                                        '    Case "L"
                                        '        单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数) / 2
                                        '        If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 二级序数)) = "自由端" Then
                                        '            单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数)
                                        '        End If
                                        '    Case "T"
                                        '        单元分支_w(一级序数, 二级序数) = 板格_t(关联板格序数) * 20
                                        'End Select
                                        Dim 分支板_A(加强筋单元_总数, 通用分支数目) As Single,
                                            分支板_AYc(加强筋单元_总数, 通用分支数目) As Single, 分支板_AZc(加强筋单元_总数, 通用分支数目) As Single,
                                            分支板_AσY(加强筋单元_总数, 通用分支数目) As Single
                                        For 三级序数 As UShort = 1 To 板格_板格板数目(关联板格序数)
                                            If 关联节点序数 = 板格_首端节点_序数(关联板格序数) Then
                                                板格_首端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                Dim 分支板_w(加强筋单元_总数, 通用分支数目) As Single
                                                分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 三级序数)
                                                Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                Select Case 分支板超出宽度
                                                    Case <= 0   '分支板宽度和小于分支所需宽度
                                                        分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 三级序数)
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 三级序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 三级序数)
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 三级序数)
                                                    Case > 0    '分支板宽度和大于分支所需宽度
                                                        Dim 通用板格板数目 As UShort = 3
                                                        Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                        '板格板实际取用宽度
                                                        板格板_w1(关联板格序数, 三级序数) = 板格板_w(关联板格序数, 三级序数) - 分支板超出宽度
                                                        '板格板宽度利用系数
                                                        Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 三级序数) / 板格板_w(关联板格序数, 三级序数)
                                                        Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y11(板格_总数, 通用板格板数目), 板格板_Z11(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                        '板格板实际取用面积
                                                        板格板_A1(关联板格序数, 三级序数) = 板格板_A(关联板格序数, 三级序数) * 板格板_ηw
                                                        '板格板实际末端坐标
                                                        板格板_Y11(关联板格序数, 三级序数) = 板格板_Y0(关联板格序数, 三级序数) + (板格板_Y1(关联板格序数, 三级序数) - 板格板_Y0(关联板格序数, 三级序数)) * 板格板_ηw
                                                        板格板_Z11(关联板格序数, 三级序数) = 板格板_Z0(关联板格序数, 三级序数) + (板格板_Z1(关联板格序数, 三级序数) - 板格板_Z0(关联板格序数, 三级序数)) * 板格板_ηw
                                                        '板格板实际形心坐标
                                                        板格板_Yc1(关联板格序数, 三级序数) = (板格板_Y0(关联板格序数, 三级序数) + 板格板_Y11(关联板格序数, 三级序数)） / 2
                                                        板格板_Zc1(关联板格序数, 三级序数) = (板格板_Z0(关联板格序数, 三级序数) + 板格板_Z11(关联板格序数, 三级序数)） / 2
                                                        '板格板实际面积坐标积数
                                                        板格板_AYc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Yc1(关联板格序数, 三级序数)
                                                        板格板_AZc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Zc1(关联板格序数, 三级序数)
                                                        '板格板实际面积强度积数
                                                        板格板_AσY1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_σY(关联板格序数, 三级序数)

                                                        'Select Case 三级序数
                                                        '    Case = 1

                                                        '    Case > 1

                                                        'End Select

                                                        '分支板实际面积
                                                        分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 三级序数)
                                                        '分支板实际面积坐标积数
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 三级序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 三级序数)
                                                        '分支板实际面积强度积数
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 三级序数)

                                                        Exit For
                                                End Select
                                            ElseIf 关联节点序数 = 板格_末端节点_序数(关联板格序数) Then
                                                板格_末端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                Dim 分支板_w(加强筋单元_总数, 通用分支数目) As Single
                                                Dim 反向序数 As UShort = 板格_板格板数目(关联板格序数) + 1 - 三级序数
                                                分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 反向序数)
                                                Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                Select Case 分支板超出宽度
                                                    Case <= 0   '分支板宽度和小于分支所需宽度
                                                        分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 反向序数)
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 反向序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 反向序数)
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 反向序数)
                                                    Case > 0    '分支板宽度和大于分支所需宽度
                                                        Dim 通用板格板数目 As UShort = 3
                                                        Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                        '板格板实际取用宽度
                                                        板格板_w1(关联板格序数, 反向序数) = 板格板_w(关联板格序数, 反向序数) - 分支板超出宽度
                                                        '板格板宽度利用系数
                                                        Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 反向序数) / 板格板_w(关联板格序数, 反向序数)
                                                        Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y01(板格_总数, 通用板格板数目), 板格板_Z01(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                        '板格板实际取用面积
                                                        板格板_A1(关联板格序数, 反向序数) = 板格板_A(关联板格序数, 反向序数) * 板格板_ηw
                                                        '板格板实际末端坐标
                                                        板格板_Y01(关联板格序数, 反向序数) = 板格板_Y1(关联板格序数, 反向序数) + (板格板_Y0(关联板格序数, 反向序数) - 板格板_Y1(关联板格序数, 反向序数)) * 板格板_ηw
                                                        板格板_Z01(关联板格序数, 反向序数) = 板格板_Z1(关联板格序数, 反向序数) + (板格板_Z0(关联板格序数, 反向序数) - 板格板_Z1(关联板格序数, 反向序数)) * 板格板_ηw
                                                        '板格板实际形心坐标
                                                        板格板_Yc1(关联板格序数, 反向序数) = (板格板_Y01(关联板格序数, 反向序数) + 板格板_Y1(关联板格序数, 反向序数)） / 2
                                                        板格板_Zc1(关联板格序数, 反向序数) = (板格板_Z01(关联板格序数, 反向序数) + 板格板_Z1(关联板格序数, 反向序数)） / 2
                                                        '板格板实际面积坐标积数
                                                        板格板_AYc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Yc1(关联板格序数, 反向序数)
                                                        板格板_AZc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Zc1(关联板格序数, 反向序数)
                                                        '板格板实际面积强度积数
                                                        板格板_AσY1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_σY(关联板格序数, 反向序数)

                                                        'Select Case 三级序数
                                                        '    Case = 1

                                                        '    Case > 1

                                                        'End Select

                                                        '分支板实际面积
                                                        分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 反向序数)
                                                        '分支板实际面积坐标积数
                                                        分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 反向序数)
                                                        分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 反向序数)
                                                        '分支板实际面积强度积数
                                                        分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 反向序数)

                                                        Exit For
                                                End Select
                                            End If
                                        Next
                                        单元分支_A(一级序数, 二级序数) = 分支板_A(一级序数, 二级序数)
                                        单元分支_AYc(一级序数, 二级序数) = 分支板_AYc(一级序数, 二级序数)
                                        单元分支_AZc(一级序数, 二级序数) = 分支板_AZc(一级序数, 二级序数)
                                        单元分支_AσY(一级序数, 二级序数) = 分支板_AσY(一级序数, 二级序数)

                                        加强筋单元_A(一级序数) += 分支板_A(一级序数, 二级序数)
                                        加强筋单元_AYc(一级序数) += 分支板_AYc(一级序数, 二级序数)
                                        加强筋单元_AZc(一级序数) += 分支板_AZc(一级序数, 二级序数)
                                        加强筋单元_AσY(一级序数) += 分支板_AσY(一级序数, 二级序数)
                                    Next

                                    加强筋单元_tP(一级序数) = (加强筋单元_A(一级序数) - 加强筋_AS(关联加强筋序数)) / 加强筋单元_wP(一级序数)
                                    加强筋单元_σYP(一级序数) = (加强筋单元_AσY(一级序数) - 加强筋_AS(关联加强筋序数) * 加强筋_σY(关联加强筋序数)) / (加强筋单元_A(一级序数) - 加强筋_AS(关联加强筋序数))

                                    加强筋单元_ηS(一级序数) = 1 + 加强筋_l(关联加强筋序数) ^ 2 / PI ^ 2 / Sqrt(加强筋_IWS(关联加强筋序数) * (0.75 * 加强筋单元_wP(一级序数) / 加强筋单元_tP(一级序数) ^ 3 + (加强筋_df(关联加强筋序数) - 加强筋_tf(关联加强筋序数) / 2) / 加强筋_tw(关联加强筋序数) ^ 3))
                                    加强筋单元_σETS(一级序数) = 标准弹性模量 / 加强筋_IPS(关联加强筋序数) * (加强筋单元_ηS(一级序数) * PI ^ 2 * 加强筋_IWS(关联加强筋序数) / 加强筋单元_lS(一级序数) ^ 2 + 0.385 * 加强筋_ITS(关联加强筋序数))

                                    加强筋单元_Yc(一级序数) = 加强筋单元_AYc(一级序数) / 加强筋单元_A(一级序数)
                                    加强筋单元_Zc(一级序数) = 加强筋单元_AZc(一级序数) / 加强筋单元_A(一级序数)
                                    加强筋单元_σY(一级序数) = 加强筋单元_AσY(一级序数) / 加强筋单元_A(一级序数)

                                    全截面_A += 加强筋单元_A(一级序数)
                                    全截面_AYc += 加强筋单元_AYc(一级序数)
                                    全截面_AZc += 加强筋单元_AZc(一级序数)
                                    全截面_AσY += 加强筋单元_AσY(一级序数)
                                Next
                            Case 5      '面板加强筋单元
                                'MsgBox("面板加强筋单元部分未完成!")
                            Case 6      '加筋板单元
                                For 一级序数 As UShort = 1 To 板格_总数
                                    Dim 板格_原始_w As Single, 板格_剩余_w As Single
                                    Dim 板格_原始_A As Single, 板格_首端_A As Single, 板格_末端_A As Single, 板格_剩余_A As Single
                                    Dim 板格_原始_AYc As Single, 板格_首端_AYc As Single, 板格_末端_AYc As Single, 板格_剩余_AYc As Single
                                    Dim 板格_原始_AZc As Single, 板格_首端_AZc As Single, 板格_末端_AZc As Single, 板格_剩余_AZc As Single
                                    Dim 板格_原始_AσY As Single, 板格_首端_AσY As Single, 板格_末端_AσY As Single, 板格_剩余_AσY As Single

                                    Dim 首端关联节点序数 As UShort, 首端关联单元序数 As UShort, 首端关联分支序数 As UShort
                                    Dim 末端关联节点序数 As UShort, 末端关联单元序数 As UShort, 末端关联分支序数 As UShort

                                    首端关联节点序数 = 板格_首端节点_序数(一级序数)
                                    Select Case 节点_tp(首端关联节点序数)
                                        Case "硬角单元"
                                            首端关联单元序数 = 节点_硬角单元_序数(首端关联节点序数)
                                        Case "面板硬角单元"
                                            首端关联单元序数 = 节点_面板硬角单元_序数(首端关联节点序数)
                                        Case "加强筋单元"
                                            首端关联单元序数 = 节点_加强筋单元_序数(首端关联节点序数)
                                        Case "面板硬角单元"
                                            首端关联单元序数 = 节点_面板硬角单元_序数(首端关联节点序数)
                                        Case Else

                                    End Select
                                    首端关联分支序数 = 板格_首端节点_分支_序数(一级序数)

                                    末端关联节点序数 = 板格_末端节点_序数(一级序数)
                                    Select Case 节点_tp(末端关联节点序数)
                                        Case "硬角单元"
                                            末端关联单元序数 = 节点_硬角单元_序数(末端关联节点序数)
                                        Case "面板硬角单元"
                                            末端关联单元序数 = 节点_面板硬角单元_序数(末端关联节点序数)
                                        Case "加强筋单元"
                                            末端关联单元序数 = 节点_加强筋单元_序数(末端关联节点序数)
                                        Case "面板硬角单元"
                                            末端关联单元序数 = 节点_面板硬角单元_序数(末端关联节点序数)
                                        Case Else

                                    End Select
                                    末端关联分支序数 = 板格_末端节点_分支_序数(一级序数)

                                    板格_原始_w = 板格_w(一级序数)
                                    板格_剩余_w = 板格_原始_w - 板格_首端_w(一级序数) - 板格_末端_w(一级序数)

                                    板格_原始_A = 板格_A(一级序数)
                                    板格_首端_A = 单元分支_A(首端关联单元序数, 首端关联分支序数)
                                    板格_末端_A = 单元分支_A(末端关联单元序数, 末端关联分支序数)
                                    板格_剩余_A = 板格_原始_A - 板格_首端_A - 板格_末端_A

                                    板格_原始_AYc = 板格_AYc(一级序数)
                                    板格_首端_AYc = 单元分支_AYc(首端关联单元序数, 首端关联分支序数)
                                    板格_末端_AYc = 单元分支_AYc(末端关联单元序数, 末端关联分支序数)
                                    板格_剩余_AYc = 板格_原始_AYc - 板格_首端_AYc - 板格_末端_AYc

                                    板格_原始_AZc = 板格_AZc(一级序数)
                                    板格_首端_AZc = 单元分支_AZc(首端关联单元序数, 首端关联分支序数)
                                    板格_末端_AZc = 单元分支_AZc(末端关联单元序数, 末端关联分支序数)
                                    板格_剩余_AZc = 板格_原始_AZc - 板格_首端_AZc - 板格_末端_AZc

                                    板格_原始_AσY = 板格_AσY(一级序数)
                                    板格_首端_AσY = 单元分支_AσY(首端关联单元序数, 首端关联分支序数)
                                    板格_末端_AσY = 单元分支_AσY(末端关联单元序数, 末端关联分支序数)
                                    板格_剩余_AσY = 板格_原始_AσY - 板格_首端_AσY - 板格_末端_AσY

                                    Select Case 板格_剩余_w / 板格_原始_w
                                        Case >= 0.1
                                            加筋板单元_总数 += 1

                                            ReDim Preserve 加筋板单元_板格_序数(加筋板单元_总数)
                                            加筋板单元_板格_序数(加筋板单元_总数) = 一级序数

                                            ReDim Preserve 加筋板单元_原始_w(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_原始_A(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_原始_Yc(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_原始_Zc(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_原始_σY(加筋板单元_总数)

                                            ReDim Preserve 加筋板单元_剩余_w(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_剩余_A(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_剩余_Yc(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_剩余_Zc(加筋板单元_总数)
                                            ReDim Preserve 加筋板单元_剩余_σY(加筋板单元_总数)

                                            加筋板单元_原始_w(加筋板单元_总数) = 板格_原始_w
                                            加筋板单元_原始_A(加筋板单元_总数) = 板格_原始_A
                                            加筋板单元_原始_Yc(加筋板单元_总数) = 板格_原始_AYc / 板格_原始_A
                                            加筋板单元_原始_Zc(加筋板单元_总数) = 板格_原始_AZc / 板格_原始_A
                                            加筋板单元_原始_σY(加筋板单元_总数) = 板格_原始_AσY / 板格_原始_A

                                            加筋板单元_剩余_w(加筋板单元_总数) = 板格_剩余_w
                                            加筋板单元_剩余_A(加筋板单元_总数) = 板格_剩余_A
                                            加筋板单元_剩余_Yc(加筋板单元_总数) = 板格_剩余_AYc / 板格_剩余_A
                                            加筋板单元_剩余_Zc(加筋板单元_总数) = 板格_剩余_AZc / 板格_剩余_A
                                            加筋板单元_剩余_σY(加筋板单元_总数) = 板格_剩余_AσY / 板格_剩余_A

                                            全截面_A += 板格_剩余_A
                                            全截面_AYc += 板格_剩余_AYc
                                            全截面_AZc += 板格_剩余_AZc
                                            全截面_AσY += 板格_剩余_AσY
                                        Case Else

                                    End Select
                                Next
                            Case Else

                        End Select
                    Next
                    全截面_Yc = 全截面_AYc / 全截面_A
                    全截面_Zc = 全截面_AZc / 全截面_A
            End Select
        Next

        '单元对象的形心坐标输出
        For 单元对象类型序数 As UShort = 1 To 6
            Select Case 单元对象类型序数
                Case 1
                    Chart2.Series.Clear()
                    Chart2.Show()
                    Chart2.Series.Add("硬角单元_形心")
                    Chart2.Series("硬角单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 硬角单元_总数
                        Chart2.Series("硬角单元_形心").Points.AddXY(硬角单元_Yc(一级序数), 硬角单元_Zc(一级序数))
                    Next
                Case 2
                    'MsgBox("面板硬角单元部分未完成!")
                Case 3
                    Chart2.Series.Add("特别硬角单元_形心")
                    Chart2.Series("特别硬角单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 特别硬角单元_总数
                        Chart2.Series("特别硬角单元_形心").Points.AddXY(特别硬角单元_Yc(一级序数), 特别硬角单元_Zc(一级序数))
                    Next
                Case 4
                    Chart2.Series.Add("加强筋单元_形心")
                    Chart2.Series("加强筋单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 加强筋单元_总数
                        Chart2.Series("加强筋单元_形心").Points.AddXY(加强筋单元_Yc(一级序数), 加强筋单元_Zc(一级序数))
                    Next
                Case 5
                    'MsgBox("面板加强筋单元部分未完成!")
                Case 6
                    Chart2.Series.Add("加筋板单元_形心")
                    Chart2.Series("加筋板单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                    For 一级序数 As UShort = 1 To 加筋板单元_总数
                        Chart2.Series("加筋板单元_形心").Points.AddXY(加筋板单元_剩余_Yc(一级序数), 加筋板单元_剩余_Zc(一级序数))
                    Next
                Case Else

            End Select
        Next

        Select Case Mid(第五部分, 5, 1)
            Case 1
                Select Case 第三部分
                    Case "IIM"
                        多角度增量迭代法()
                    Case "MIM"
                        基于二分法的多角度增量迭代法()
                End Select
            Case 2
                Select Case 第三部分
                    Case "IIM"
                        多角度第二增量迭代法()
                    Case "MIM"
                        基于二分法的第二多角度增量迭代法()
                End Select
        End Select
        End
    End Sub

    Private Sub 增量解析法(sender As Object, e As EventArgs) Handles Button7.Click
        Dim χy_总数 As UShort = χ_总数
        Dim χz_总数 As UShort = χ_总数
        增量解析法双轴计算()
    End Sub

    Private Sub 多角度增量迭代法()
        '中拱为正, 单元在中性轴(α=0)之上为正, 单元受拉为正
        '弯矩方向
        '水线倾角_α = Val(InputBox("水线倾角_α(deg)", "水线倾角", "0")) / 180 * PI

        Dim ΣM(χ_总数, ζ_总数, α_总数) As Single
        Dim ΣFO(χ_总数, ζ_总数, α_总数) As Single, ΣFA(χ_总数, ζ_总数, α_总数) As Single
        Dim ΣFP(χ_总数, ζ_总数, α_总数) As Single, ΣFN(χ_总数, ζ_总数, α_总数) As Single
        Dim ΣFYP(χ_总数, ζ_总数, α_总数) As Single, ΣFYN(χ_总数, ζ_总数, α_总数) As Single
        Dim ΣFZP(χ_总数, ζ_总数, α_总数) As Single, ΣFZN(χ_总数, ζ_总数, α_总数) As Single
        Dim FYP(χ_总数, ζ_总数, α_总数) As Single, FYN(χ_总数, ζ_总数, α_总数) As Single
        Dim FZP(χ_总数, ζ_总数, α_总数) As Single, FZN(χ_总数, ζ_总数, α_总数) As Single

        Chart3.Series.Clear()
        Chart3.Series.Add("弯矩曲率曲线")
        Chart3.Series("弯矩曲率曲线").ChartType = DataVisualization.Charting.SeriesChartType.Line

        Chart3.Series("弯矩曲率曲线").Points.AddXY(0, 0)

        For 一级序数 As UShort = 1 To χ_总数
            '[注意正负!]
            Select Case Mid(第五部分, 1, 1)
                Case "P"
                    χ_瞬时 = Abs(χ_初值) + 一级序数 * Abs(χ_增量)
                Case "N"
                    χ_瞬时 = -Abs(χ_初值) - 一级序数 * Abs(χ_增量)
            End Select

            Dim 调试输出 As String = ""
            Dim α_候选(ζ_总数) As Single

            Dim 轴力相对差值(ζ_总数) As Single
            For 二级序数 As UShort = 1 To ζ_总数
                ζ_瞬时 = ζ_初值 + (-1) ^ 二级序数 * ζ_增量 * (二级序数 \ 2)

                Dim 向量积(α_总数) As Single
                For 三级序数 As UShort = 1 To α_总数
                    α_瞬时 = α_初值 + (-1) ^ 三级序数 * α_增量 * (三级序数 \ 2)

                    For 单元对象类型序数 As UShort = 1 To 6
                        Select Case 单元对象类型序数
                            Case 1
                                For 四级序数 As UShort = 1 To 硬角单元_总数
                                    单元_Yc = 硬角单元_Yc(四级序数) / 1000
                                    单元_Zc = 硬角单元_Zc(四级序数) / 1000

                                    Select Case 第一部分
                                        Case "OT"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                            End Select
                                        Case "BC"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                            End Select
                                    End Select

                                    单元_A = 硬角单元_A(四级序数) / 1000000
                                    单元_σY = 硬角单元_σY(四级序数)
                                    '[注意正负!]
                                    单元_L = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)

                                    '[注意正负!]
                                    单元_εO = 单元_L * χ_瞬时

                                    单元_εY = 单元_σY / 标准弹性模量

                                    '[注意正负!]
                                    单元_εR = 单元_εO / 单元_εY

                                    Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                                    Dim EP As Single = Φ * 单元_σY
                                    单元_σO = EP
                                    单元_FO = 单元_σO * 单元_A
                                    单元_MO = 单元_FO * 单元_L

                                    Select Case 单元_FO
                                        Case > 0
                                            ΣFP(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case < 0
                                            ΣFN(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case = 0

                                        Case Else

                                    End Select

                                    ΣFO(一级序数, 二级序数, 三级序数) += 单元_FO
                                    ΣFA(一级序数, 二级序数, 三级序数) += Abs(单元_FO)
                                    ΣM(一级序数, 二级序数, 三级序数) += 单元_MO
                                Next
                            Case 2
                                'MsgBox("面板硬角单元部分未完成!")
                            Case 3
                                For 四级序数 As UShort = 1 To 特别硬角单元_总数
                                    单元_Yc = 特别硬角单元_Yc(四级序数) / 1000
                                    单元_Zc = 特别硬角单元_Zc(四级序数) / 1000

                                    Select Case 第一部分
                                        Case "OT"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                            End Select
                                        Case "BC"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                            End Select
                                    End Select

                                    单元_A = 特别硬角单元_A(四级序数) / 1000000
                                    单元_σY = 特别硬角单元_σY(四级序数)
                                    '[注意正负!]
                                    单元_L = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                                    '[注意正负!]
                                    单元_εO = 单元_L * χ_瞬时

                                    单元_εY = 单元_σY / 标准弹性模量

                                    '[注意正负!]
                                    单元_εR = 单元_εO / 单元_εY

                                    Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                                    Dim EP As Single = Φ * 单元_σY
                                    单元_σO = EP
                                    单元_FO = 单元_σO * 单元_A
                                    单元_MO = 单元_FO * 单元_L

                                    Select Case 单元_FO
                                        Case > 0
                                            ΣFP(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case < 0
                                            ΣFN(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case = 0

                                        Case Else

                                    End Select

                                    ΣFO(一级序数, 二级序数, 三级序数) += 单元_FO
                                    ΣFA(一级序数, 二级序数, 三级序数) += Abs(单元_FO)
                                    ΣM(一级序数, 二级序数, 三级序数) += 单元_MO
                                Next
                            Case 4
                                For 四级序数 As UShort = 1 To 加强筋单元_总数
                                    Dim 单元_Yc As Single = 加强筋单元_Yc(四级序数) / 1000
                                    Dim 单元_Zc As Single = 加强筋单元_Zc(四级序数) / 1000

                                    Select Case 第一部分
                                        Case "OT"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                            End Select
                                        Case "BC"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                            End Select
                                    End Select

                                    Dim 单元_A As Single = 加强筋单元_A(四级序数)
                                    Dim 单元_σY As Single = 加强筋单元_σY(四级序数)

                                    Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(四级序数)
                                    Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)

                                    Dim 单元_σYP As Single = 加强筋单元_σYP(四级序数)
                                    Dim 单元_lP As Single = 加强筋单元_lP(四级序数)
                                    Dim 单元_wP As Single = 加强筋单元_wP(四级序数)
                                    Dim 单元_tP As Single = 加强筋单元_tP(四级序数)
                                    Dim 单元_σYS As Single = 加强筋单元_σYS(四级序数)
                                    Dim 单元_lS As Single = 加强筋单元_lS(四级序数)
                                    Dim 单元_hw As Single = 加强筋单元_hw(四级序数)
                                    Dim 单元_tw As Single = 加强筋单元_tw(四级序数)
                                    Dim 单元_wf As Single = 加强筋单元_wf(四级序数)
                                    Dim 单元_tf As Single = 加强筋单元_tf(四级序数)
                                    Dim 单元_dx As Single = 加强筋单元_dx(四级序数)
                                    Dim 单元_tpS As String = 加强筋单元_tpS(四级序数)
                                    Dim 单元_mk As Boolean = 加强筋单元_mk(四级序数)
                                    'Dim 单元_σETS As Single = 加强筋单元_σETS(四级序数)
                                    'Dim 单元_σELS As Single = 加强筋单元_σELS(四级序数)

                                    Dim 单元_εYP As Single = 单元_σYP / 标准弹性模量
                                    Dim 单元_AP As Single = 单元_wP * 单元_tP
                                    Dim 单元_ICP As Single = 单元_wP * 单元_tP ^ 3 / 12
                                    Dim 单元_IOP As Single = 单元_ICP + 单元_AP * (-单元_tP / 2) ^ 2
                                    Dim 单元_βOP As Single = 单元_wP / 单元_tP * 单元_εYP ^ (1 / 2)
                                    Dim 单元_wEoP As Single = If(单元_βOP >= 1.25, (2.25 / 单元_βOP - 1.25 / 单元_βOP ^ 2) * 单元_wP, 单元_wP)

                                    Dim 单元_εYS As Single = 单元_σYS / 标准弹性模量
                                    Dim 单元_Aw As Single = 单元_hw * 单元_tw
                                    Dim 单元_Icw As Single = 单元_tw * 单元_hw ^ 3 / 12
                                    Dim 单元_Iow As Single = 单元_Icw + 单元_Aw * (单元_hw / 2) ^ 2
                                    Dim 单元_βOw As Single = 单元_hw / 单元_tw * 单元_εYS ^ (1 / 2)
                                    Dim 单元_hEow As Single = If(单元_βOw >= 1.25, (2.25 / 单元_βOw - 1.25 / 单元_βOw ^ 2) * 单元_hw, 单元_hw)
                                    Dim 单元_df As Single = If(单元_tpS = "F", 单元_hw, If(单元_tpS = "B", 单元_hw - 单元_tf / 2, If(单元_tpS = "T" Or 单元_tpS = "L1" Or 单元_tpS = "L2", 单元_hw + 单元_tf / 2, 单元_hw - 单元_dx - 单元_tf / 2)))
                                    Dim 单元_Af As Single = 单元_wf * 单元_tf
                                    Dim 单元_Icf As Single = 单元_wf * 单元_tf ^ 3 / 12
                                    Dim 单元_Iof As Single = 单元_Icf + 单元_Af * 单元_df ^ 2
                                    Dim 单元_AS As Single = 单元_Aw + 单元_Af
                                    Dim 单元_hcS As Single = 单元_Aw / 单元_AS * 单元_hw / 2 + 单元_Af / 单元_AS * 单元_df
                                    Dim 单元_IOS As Single = 单元_Iow + 单元_Iof
                                    Dim 单元_ICS As Single = 单元_Icw + 单元_Aw * (单元_hw / 2 - 单元_hcS) ^ 2 + 单元_Icf + 单元_Af * (单元_df - 单元_hcS) ^ 2
                                    Dim 单元_IpS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw / 3, 单元_Aw * (单元_df - 单元_tf / 2) ^ 2 / 3 + 单元_Af * 单元_df ^ 2)
                                    Dim 单元_ItS As Single = If(单元_tpS = "F", 单元_hw * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / 单元_hw), (单元_df - 单元_tf / 2) * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / (单元_df - 单元_tf / 2)) + 单元_wf * 单元_tf ^ 3 / 3 * (1 - 0.63 * 单元_tf / 单元_wf))
                                    Dim 单元_IwS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw ^ 3 / 36, If(单元_tpS = "B" Or 单元_tpS = "L1" Or 单元_tpS = "L2" Or 单元_tpS = "L3", 单元_Af * 单元_df ^ 2 * 单元_wf ^ 2 / 12 * (单元_Af + 2.6 * 单元_Aw) / (单元_Af * 单元_Aw), 单元_wf ^ 3 * 单元_tf * 单元_df ^ 2 / 12))
                                    Dim 单元_ηS As Single = 1 + (单元_lS / PI) ^ 2 / (单元_IwS * (0.75 * 单元_wP / 单元_tP ^ 3 + (单元_df - 单元_tf / 2) / 单元_tw ^ 3)) ^ (1 / 2)
                                    Dim 单元_σETS As Single = 标准弹性模量 / 单元_IpS * (单元_ηS * PI ^ 2 * 单元_IwS / 单元_lS ^ 2 + 0.385 * 单元_ItS)
                                    Dim 单元_σELS As Single = 160000 * (单元_tw / 单元_hw) ^ 2

                                    'Dim 单元_A As Single = 单元_AS + 单元_AP
                                    Dim 单元_hc As Single = 单元_AS / 单元_A * 单元_hcS + 单元_AP / 单元_A * (-单元_tP / 2)
                                    Dim 单元_Io As Single = 单元_IOP + 单元_IOS
                                    Dim 单元_Ic As Single = 单元_Io - 单元_A * 单元_hc ^ 2
                                    'Dim 单元_σY As Single = 单元_AS / A * 单元_σYS + 单元_AP / 单元_A * 单元_σYP
                                    'Dim 单元_εY As Single = 单元_σY / 标准弹性模量

                                    '[注意正负!]
                                    Dim 单元_L As Single = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                                    '[注意正负!]
                                    单元_εO = 单元_L * χ_瞬时

                                    单元_εY = 单元_σY / 标准弹性模量

                                    '[注意正负!]
                                    单元_εR = 单元_εO / 单元_εY
                                    '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING

                                    'Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                                    'Dim EP As Single = Φ * 单元_σY

                                    'wP / lP
                                    'Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)
                                    'Dim βPε As Single = βOP * 单元_εR ^ (1 / 2)
                                    'Dim BC As Single, FT As Single, WB As Single

                                    ''''''
                                    Dim 单元_Φε As Single = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))
                                    Dim 单元_σEPε As Single = 单元_Φε * 单元_σY

                                    Dim 单元_Rlt_EP As String = Format(单元_σEPε, "000.000")

                                    Select Case 单元_εR
                                        Case < 0
                                            Dim 单元_βPε As Single = 单元_βOP * (-单元_εR) ^ (1 / 2)
                                            Dim 单元_wEPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_wP, 单元_wP)
                                            Dim 单元_wE1Pε As Single = If(单元_βPε >= 1, 单元_wP / 单元_βPε, 单元_wP)
                                            Dim 单元_AEPε As Single = 单元_wEPε * 单元_tP
                                            Dim 单元_AE1Pε As Single = 单元_wE1Pε * 单元_tP
                                            Dim 单元_IcE1Pε As Single = 单元_wE1Pε * 单元_tP ^ 3 / 12
                                            Dim 单元_IoE1Pε As Single = 单元_IcE1Pε + 单元_AE1Pε * (-单元_tP / 2) ^ 2

                                            Dim 单元_βwε As Single = 单元_βOw * (-单元_εR) ^ (1 / 2)
                                            Dim 单元_hEwε As Single = If(单元_βwε >= 1.25, (2.25 / 单元_βwε - 1.25 / 单元_βwε ^ 2) * 单元_hw, 单元_hw)
                                            Dim 单元_AEwε As Single = 单元_hEwε * 单元_tw
                                            Dim 单元_AESε As Single = 单元_AEwε + 单元_Af

                                            Dim 单元_AEε As Single = 单元_AEPε + 单元_AS
                                            Dim 单元_AE1ε As Single = 单元_AE1Pε + 单元_AS
                                            Dim 单元_hcE1ε As Single = 单元_AE1Pε / 单元_AE1ε * (-单元_tP / 2) + 单元_AS / 单元_AE1ε * 单元_hcS
                                            Dim 单元_IoE1ε As Single = 单元_IoE1Pε + 单元_IOS
                                            Dim 单元_IcE1ε As Single = 单元_IoE1ε - 单元_AE1ε * 单元_hcE1ε ^ 2
                                            Dim 单元_lE1Pε As Single = 单元_hcE1ε - (-单元_tP / 2)
                                            Dim 单元_lE1Sε As Single = If(单元_tpS = "F" Or 单元_tpS = "B" Or 单元_tpS = "L3", 单元_hw - 单元_hcE1ε, 单元_hw + 单元_tf - 单元_hcE1ε)
                                            Dim 单元_σYE1ε As Single = 单元_AS * 单元_lE1Sε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYS + 单元_AE1Pε * 单元_lE1Pε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYP
                                            Dim 单元_σECε As Single = PI ^ 2 * 标准弹性模量 * 单元_IcE1ε / 单元_AEε / 单元_lS ^ 2
                                            Dim 单元_σCCε As Single = If(单元_σECε <= 单元_σYE1ε / 2 * (-单元_εR), 单元_σECε / (-单元_εR), 单元_σYE1ε * (1 - 单元_σYE1ε * (-单元_εR) / 4 / 单元_σECε))
                                            Dim 单元_σBCε As Single = 单元_Φε * 单元_σCCε * 单元_AEε / 单元_A

                                            Dim 单元_Rlt_BC As String = Format(单元_σBCε, "000.000")

                                            Dim 单元_σCTSε As Single = If(单元_σETS <= 单元_σYS / 2 * (-单元_εR), 单元_σETS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σETS))
                                            Dim 单元_σCPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_σYP, 单元_σYP)
                                            Dim 单元_σFTε As Single = 单元_Φε * (单元_AS * 单元_σCTSε + 单元_AP * 单元_σCPε) / 单元_A

                                            Dim 单元_Rlt_FT As String = Format(单元_σFTε, "000.000")

                                            Dim 单元_σCLSε As Single = If(单元_σELS <= 单元_σYS / 2 * (-单元_εR), 单元_σELS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σELS))
                                            Dim 单元_σWBε As Single = If(单元_tpS = "F", 单元_Φε * (单元_AS * 单元_σCLSε + 单元_AP * 单元_σCPε) / 单元_A, 单元_Φε * (单元_AESε * 单元_σYS + 单元_AEPε * 单元_σYP) / 单元_A)

                                            Dim 单元_Rlt_WB As String = Format(单元_σWBε, "000.000")

                                            ''''''
                                            单元_σO = Max(单元_σBCε, Max(单元_σFTε, 单元_σWBε)) ' 单元_σO = 单元_σEPε   ''''''NO BUCKLING
                                            'If 四级序数 = 1 Then Debug.Print(单元_Rlt_EP & ", " & 单元_Rlt_BC & ", " & 单元_Rlt_FT & ", " & 单元_Rlt_WB)
                                        Case > 0
                                            单元_σO = 单元_σEPε
                                        Case = 0
                                            单元_σO = 0
                                        Case Else

                                    End Select

                                    '考虑[单元_剩余_A]!
                                    单元_FO = 单元_σO * 单元_A / 1000000
                                    单元_MO = 单元_FO * 单元_L

                                    Select Case 单元_FO
                                        Case > 0
                                            ΣFP(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case < 0
                                            ΣFN(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case = 0

                                        Case Else

                                    End Select

                                    ΣFO(一级序数, 二级序数, 三级序数) += 单元_FO
                                    ΣFA(一级序数, 二级序数, 三级序数) += Abs(单元_FO)
                                    ΣM(一级序数, 二级序数, 三级序数) += 单元_MO
                                Next
                            Case 5
                                'MsgBox("面板加强筋单元部分未完成!")
                            Case 6
                                For 四级序数 As UShort = 1 To 加筋板单元_总数
                                    Dim 单元_原始_Yc As Single = 加筋板单元_原始_Yc(四级序数) / 1000
                                    Dim 单元_原始_Zc As Single = 加筋板单元_原始_Zc(四级序数) / 1000

                                    Select Case 第一部分
                                        Case "OT"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                            End Select
                                        Case "BC"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                            End Select
                                    End Select

                                    Dim 单元_原始_A As Single = 加筋板单元_原始_A(四级序数)
                                    Dim 单元_原始_σY As Single = 加筋板单元_原始_σY(四级序数)

                                    Dim 单元_剩余_Yc As Single = 加筋板单元_剩余_Yc(四级序数) / 1000
                                    Dim 单元_剩余_Zc As Single = 加筋板单元_剩余_Zc(四级序数) / 1000
                                    Dim 单元_剩余_A As Single = 加筋板单元_剩余_A(四级序数)
                                    Dim 单元_剩余_σY As Single = 加筋板单元_剩余_σY(四级序数)

                                    Dim 关联板格序数 As UShort = 加筋板单元_板格_序数(四级序数)

                                    '[注意正负!]
                                    单元_L = -单元_原始_Yc * Sin(α_瞬时) + (单元_原始_Zc - ζ_瞬时) * Cos(α_瞬时)
                                    '[注意正负!]
                                    单元_εO = 单元_L * χ_瞬时

                                    单元_εY = 单元_原始_σY / 标准弹性模量

                                    '[注意正负!]
                                    单元_εR = 单元_εO / 单元_εY
                                    '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING

                                    Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                                    Dim EP As Single = Φ * 单元_原始_σY

                                    'wP / lP
                                    Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)

                                    Select Case 单元_εR
                                        Case < 0
                                            Dim βPε As Single = βOP * (-单元_εR) ^ (1 / 2)
                                            Dim FB As Single = Φ * 单元_原始_σY * Min(1, 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数) * (2.25 / βPε - 1.25 / βPε ^ 2) + 0.1 * (1 - 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数)) * (1 + 1 / βPε ^ 2))
                                            单元_σO = FB'单元_σO = EP   ' '''''NO BUCKLING
                                        Case > 0
                                            单元_σO = EP
                                        Case = 0
                                            单元_σO = 0
                                        Case Else

                                    End Select

                                    '考虑[单元_剩余_A]!
                                    单元_FO = 单元_σO * 单元_剩余_A / 1000000
                                    单元_MO = 单元_FO * 单元_L

                                    Select Case 单元_FO
                                        Case > 0
                                            ΣFP(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_原始_Yc
                                            ΣFZP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_原始_Zc
                                        Case < 0
                                            ΣFN(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_原始_Yc
                                            ΣFZN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_原始_Zc
                                        Case = 0

                                        Case Else

                                    End Select

                                    ΣFO(一级序数, 二级序数, 三级序数) += 单元_FO
                                    ΣFA(一级序数, 二级序数, 三级序数) += Abs(单元_FO)
                                    ΣM(一级序数, 二级序数, 三级序数) += 单元_MO
                                Next
                            Case Else

                        End Select
                    Next

                    FYP(一级序数, 二级序数, 三级序数) = ΣFYP(一级序数, 二级序数, 三级序数) / ΣFP(一级序数, 二级序数, 三级序数)
                    FZP(一级序数, 二级序数, 三级序数) = ΣFZP(一级序数, 二级序数, 三级序数) / ΣFP(一级序数, 二级序数, 三级序数)

                    FYN(一级序数, 二级序数, 三级序数) = ΣFYN(一级序数, 二级序数, 三级序数) / ΣFN(一级序数, 二级序数, 三级序数)
                    FZN(一级序数, 二级序数, 三级序数) = ΣFZN(一级序数, 二级序数, 三级序数) / ΣFN(一级序数, 二级序数, 三级序数)

                    合力倾角_α = Atan((FZP(一级序数, 二级序数, 三级序数) - FZN(一级序数, 二级序数, 三级序数)) / (FYP(一级序数, 二级序数, 三级序数) - FYN(一级序数, 二级序数, 三级序数)))
                    向量积(三级序数) = Abs(Cos(水线倾角_α) * Cos(合力倾角_α) + Sin(水线倾角_α) * Sin(合力倾角_α))
                    Select Case 三级序数
                        Case 1
                            向量积(0) = 向量积(三级序数)
                            α_临界 = 三级序数

                            ΣFO(一级序数, 二级序数, 0) = ΣFO(一级序数, 二级序数, α_临界)
                            ΣFA(一级序数, 二级序数, 0) = ΣFA(一级序数, 二级序数, α_临界)
                            ΣM(一级序数, 二级序数, 0) = ΣM(一级序数, 二级序数, α_临界)
                        Case > 1
                            If 向量积(0) > 向量积(三级序数) Then
                                向量积(0) = 向量积(三级序数)
                                α_临界 = 三级序数

                                ΣFO(一级序数, 二级序数, 0) = ΣFO(一级序数, 二级序数, α_临界)
                                ΣFA(一级序数, 二级序数, 0) = ΣFA(一级序数, 二级序数, α_临界)
                                ΣM(一级序数, 二级序数, 0) = ΣM(一级序数, 二级序数, α_临界)
                            End If
                        Case Else

                    End Select
                Next    'α_循环
                α_初值 += (-1) ^ α_临界 * α_增量 * (α_临界 \ 2)

                α_候选(二级序数) = α_初值 * 180 / PI

                轴力相对差值(二级序数) = Abs(ΣFO(一级序数, 二级序数, 0)) '/ ΣFA(一级序数, 二级序数, 0)
                Select Case 二级序数
                    Case 1
                        轴力相对差值(0) = 轴力相对差值(二级序数)
                        ζ_临界 = 二级序数

                        ΣM(一级序数, 0, 0) = ΣM(一级序数, ζ_临界, 0)
                    Case > 1
                        If 轴力相对差值(0) > 轴力相对差值(二级序数) Then
                            轴力相对差值(0) = 轴力相对差值(二级序数)
                            ζ_临界 = 二级序数

                            ΣM(一级序数, 0, 0) = ΣM(一级序数, ζ_临界, 0)
                        End If
                    Case Else

                End Select
                '### Debug.Print(ΣFO(一级序数, 二级序数, 0))  
            Next    'ζ_循环
            ζ_初值 += (-1) ^ ζ_临界 * ζ_增量 * (ζ_临界 \ 2)

            调试输出 = 调试输出 & Format(χ_瞬时, "0.000000") & Space(8) & Format(ΣM(一级序数, 0, 0), "00000.000") & Space(8) & Format(ζ_初值, "000.000") & Space(8) & Format(α_候选(ζ_临界), "000.000")

            Debug.Print(调试输出) 'Debug.Print(一级序数) '

            Chart3.Series("弯矩曲率曲线").Points.AddXY(χ_瞬时, ΣM(一级序数, 0, 0))
            DataGridView3.Rows.Add(一级序数, χ_瞬时, ΣM(一级序数, 0, 0))

            If Abs(ΣM(0, 0, 0)) < Abs(ΣM(一级序数, 0, 0)) Then
                ΣM(0, 0, 0) = ΣM(一级序数, 0, 0)
                χ_临界 = 一级序数
            End If
            'Select Case 一级序数
            '    Case 1
            '        ΣM(0, 0, 0) = ΣM(一级序数, 0, 0)
            '        χ_临界 = 一级序数
            '    Case > 1
            '        If Abs(ΣM(0, 0, 0)) < Abs(ΣM(一级序数, 0, 0)) Then
            '            ΣM(0, 0, 0) = ΣM(一级序数, 0, 0)
            '            χ_临界 = 一级序数
            '        End If
            'End Select
        Next    'χ_循环

        χ_初值 = Abs(χ_初值) + χ_临界 * Abs(χ_增量)
        Debug.Print("全截面_ΣM(0, 0, 0) = " & ΣM(0, 0, 0) & " [MN.m]")
        Debug.Print("χ_临界 = " & χ_初值 & " [1/m]")
    End Sub

    Private Sub 多角度第二增量迭代法()
        '中拱为正, 单元在中性轴(α=0)之上为正, 单元受拉为正
        '弯矩方向
        '水线倾角_α = Val(InputBox("水线倾角_α(deg)", "水线倾角", "0")) / 180 * PI

        Dim ΣM(χ_总数, γ_总数, α_总数) As Single
        Dim ΣFO(χ_总数, γ_总数, α_总数) As Single, ΣFA(χ_总数, γ_总数, α_总数) As Single
        Dim ΣFP(χ_总数, γ_总数, α_总数) As Single, ΣFN(χ_总数, γ_总数, α_总数) As Single
        Dim ΣFYP(χ_总数, γ_总数, α_总数) As Single, ΣFYN(χ_总数, γ_总数, α_总数) As Single
        Dim ΣFZP(χ_总数, γ_总数, α_总数) As Single, ΣFZN(χ_总数, γ_总数, α_总数) As Single
        Dim FYP(χ_总数, γ_总数, α_总数) As Single, FYN(χ_总数, γ_总数, α_总数) As Single
        Dim FZP(χ_总数, γ_总数, α_总数) As Single, FZN(χ_总数, γ_总数, α_总数) As Single

        Chart3.Series.Clear()
        Chart3.Series.Add("弯矩曲率曲线")
        Chart3.Series("弯矩曲率曲线").ChartType = DataVisualization.Charting.SeriesChartType.Line

        Chart3.Series("弯矩曲率曲线").Points.AddXY(0, 0)

        For 一级序数 As UShort = 1 To χ_总数
            '[注意正负!]
            Select Case Mid(第五部分, 1, 1)
                Case "P"
                    χ_瞬时 = Abs(χ_初值) + 一级序数 * Abs(χ_增量)
                Case "N"
                    χ_瞬时 = -Abs(χ_初值) - 一级序数 * Abs(χ_增量)
            End Select

            Dim 调试输出 As String = ""
            Dim α_候选(γ_总数) As Single

            Dim 轴力相对差值(γ_总数) As Single
            For 二级序数 As UShort = 1 To γ_总数
                γ_瞬时 = γ_初值 + (-1) ^ 二级序数 * γ_增量 * (二级序数 \ 2)
                Dim 向量积(α_总数) As Single
                For 三级序数 As UShort = 1 To α_总数
                    α_瞬时 = α_初值 + (-1) ^ 三级序数 * α_增量 * (三级序数 \ 2)

                    For 单元对象类型序数 As UShort = 1 To 6
                        Select Case 单元对象类型序数
                            Case 1
                                For 四级序数 As UShort = 1 To 硬角单元_总数
                                    单元_Yc = 硬角单元_Yc(四级序数) / 1000
                                    单元_Zc = 硬角单元_Zc(四级序数) / 1000

                                    Select Case 第一部分
                                        Case "OT"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                            End Select
                                        Case "BC"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                            End Select
                                    End Select

                                    单元_A = 硬角单元_A(四级序数) / 1000000
                                    单元_σY = 硬角单元_σY(四级序数)
                                    '[注意正负!]
                                    单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)

                                    '[注意正负!]
                                    单元_εO = 单元_L * χ_瞬时

                                    单元_εY = 单元_σY / 标准弹性模量

                                    '[注意正负!]
                                    单元_εR = 单元_εO / 单元_εY

                                    Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                                    Dim EP As Single = Φ * 单元_σY
                                    单元_σO = EP
                                    单元_FO = 单元_σO * 单元_A
                                    单元_MO = 单元_FO * 单元_L

                                    Select Case 单元_FO
                                        Case > 0
                                            ΣFP(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case < 0
                                            ΣFN(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case = 0

                                        Case Else

                                    End Select

                                    ΣFO(一级序数, 二级序数, 三级序数) += 单元_FO
                                    ΣFA(一级序数, 二级序数, 三级序数) += Abs(单元_FO)
                                    ΣM(一级序数, 二级序数, 三级序数) += 单元_MO
                                Next
                            Case 2
                                'MsgBox("面板硬角单元部分未完成!")
                            Case 3
                                For 四级序数 As UShort = 1 To 特别硬角单元_总数
                                    单元_Yc = 特别硬角单元_Yc(四级序数) / 1000
                                    单元_Zc = 特别硬角单元_Zc(四级序数) / 1000

                                    Select Case 第一部分
                                        Case "OT"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                            End Select
                                        Case "BC"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                            End Select
                                    End Select

                                    单元_A = 特别硬角单元_A(四级序数) / 1000000
                                    单元_σY = 特别硬角单元_σY(四级序数)
                                    '[注意正负!]
                                    单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                                    '[注意正负!]
                                    单元_εO = 单元_L * χ_瞬时

                                    单元_εY = 单元_σY / 标准弹性模量

                                    '[注意正负!]
                                    单元_εR = 单元_εO / 单元_εY

                                    Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                                    Dim EP As Single = Φ * 单元_σY
                                    单元_σO = EP
                                    单元_FO = 单元_σO * 单元_A
                                    单元_MO = 单元_FO * 单元_L

                                    Select Case 单元_FO
                                        Case > 0
                                            ΣFP(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case < 0
                                            ΣFN(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case = 0

                                        Case Else

                                    End Select

                                    ΣFO(一级序数, 二级序数, 三级序数) += 单元_FO
                                    ΣFA(一级序数, 二级序数, 三级序数) += Abs(单元_FO)
                                    ΣM(一级序数, 二级序数, 三级序数) += 单元_MO
                                Next
                            Case 4
                                For 四级序数 As UShort = 1 To 加强筋单元_总数
                                    Dim 单元_Yc As Single = 加强筋单元_Yc(四级序数) / 1000
                                    Dim 单元_Zc As Single = 加强筋单元_Zc(四级序数) / 1000

                                    Select Case 第一部分
                                        Case "OT"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                            End Select
                                        Case "BC"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                            End Select
                                    End Select

                                    Dim 单元_A As Single = 加强筋单元_A(四级序数)
                                    Dim 单元_σY As Single = 加强筋单元_σY(四级序数)

                                    Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(四级序数)
                                    Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)

                                    Dim 单元_σYP As Single = 加强筋单元_σYP(四级序数)
                                    Dim 单元_lP As Single = 加强筋单元_lP(四级序数)
                                    Dim 单元_wP As Single = 加强筋单元_wP(四级序数)
                                    Dim 单元_tP As Single = 加强筋单元_tP(四级序数)
                                    Dim 单元_σYS As Single = 加强筋单元_σYS(四级序数)
                                    Dim 单元_lS As Single = 加强筋单元_lS(四级序数)
                                    Dim 单元_hw As Single = 加强筋单元_hw(四级序数)
                                    Dim 单元_tw As Single = 加强筋单元_tw(四级序数)
                                    Dim 单元_wf As Single = 加强筋单元_wf(四级序数)
                                    Dim 单元_tf As Single = 加强筋单元_tf(四级序数)
                                    Dim 单元_dx As Single = 加强筋单元_dx(四级序数)
                                    Dim 单元_tpS As String = 加强筋单元_tpS(四级序数)
                                    Dim 单元_mk As Boolean = 加强筋单元_mk(四级序数)
                                    'Dim 单元_σETS As Single = 加强筋单元_σETS(四级序数)
                                    'Dim 单元_σELS As Single = 加强筋单元_σELS(四级序数)

                                    Dim 单元_εYP As Single = 单元_σYP / 标准弹性模量
                                    Dim 单元_AP As Single = 单元_wP * 单元_tP
                                    Dim 单元_ICP As Single = 单元_wP * 单元_tP ^ 3 / 12
                                    Dim 单元_IOP As Single = 单元_ICP + 单元_AP * (-单元_tP / 2) ^ 2
                                    Dim 单元_βOP As Single = 单元_wP / 单元_tP * 单元_εYP ^ (1 / 2)
                                    Dim 单元_wEoP As Single = If(单元_βOP >= 1.25, (2.25 / 单元_βOP - 1.25 / 单元_βOP ^ 2) * 单元_wP, 单元_wP)

                                    Dim 单元_εYS As Single = 单元_σYS / 标准弹性模量
                                    Dim 单元_Aw As Single = 单元_hw * 单元_tw
                                    Dim 单元_Icw As Single = 单元_tw * 单元_hw ^ 3 / 12
                                    Dim 单元_Iow As Single = 单元_Icw + 单元_Aw * (单元_hw / 2) ^ 2
                                    Dim 单元_βOw As Single = 单元_hw / 单元_tw * 单元_εYS ^ (1 / 2)
                                    Dim 单元_hEow As Single = If(单元_βOw >= 1.25, (2.25 / 单元_βOw - 1.25 / 单元_βOw ^ 2) * 单元_hw, 单元_hw)
                                    Dim 单元_df As Single = If(单元_tpS = "F", 单元_hw, If(单元_tpS = "B", 单元_hw - 单元_tf / 2, If(单元_tpS = "T" Or 单元_tpS = "L1" Or 单元_tpS = "L2", 单元_hw + 单元_tf / 2, 单元_hw - 单元_dx - 单元_tf / 2)))
                                    Dim 单元_Af As Single = 单元_wf * 单元_tf
                                    Dim 单元_Icf As Single = 单元_wf * 单元_tf ^ 3 / 12
                                    Dim 单元_Iof As Single = 单元_Icf + 单元_Af * 单元_df ^ 2
                                    Dim 单元_AS As Single = 单元_Aw + 单元_Af
                                    Dim 单元_hcS As Single = 单元_Aw / 单元_AS * 单元_hw / 2 + 单元_Af / 单元_AS * 单元_df
                                    Dim 单元_IOS As Single = 单元_Iow + 单元_Iof
                                    Dim 单元_ICS As Single = 单元_Icw + 单元_Aw * (单元_hw / 2 - 单元_hcS) ^ 2 + 单元_Icf + 单元_Af * (单元_df - 单元_hcS) ^ 2
                                    Dim 单元_IpS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw / 3, 单元_Aw * (单元_df - 单元_tf / 2) ^ 2 / 3 + 单元_Af * 单元_df ^ 2)
                                    Dim 单元_ItS As Single = If(单元_tpS = "F", 单元_hw * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / 单元_hw), (单元_df - 单元_tf / 2) * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / (单元_df - 单元_tf / 2)) + 单元_wf * 单元_tf ^ 3 / 3 * (1 - 0.63 * 单元_tf / 单元_wf))
                                    Dim 单元_IwS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw ^ 3 / 36, If(单元_tpS = "B" Or 单元_tpS = "L1" Or 单元_tpS = "L2" Or 单元_tpS = "L3", 单元_Af * 单元_df ^ 2 * 单元_wf ^ 2 / 12 * (单元_Af + 2.6 * 单元_Aw) / (单元_Af * 单元_Aw), 单元_wf ^ 3 * 单元_tf * 单元_df ^ 2 / 12))
                                    Dim 单元_ηS As Single = 1 + (单元_lS / PI) ^ 2 / (单元_IwS * (0.75 * 单元_wP / 单元_tP ^ 3 + (单元_df - 单元_tf / 2) / 单元_tw ^ 3)) ^ (1 / 2)
                                    Dim 单元_σETS As Single = 标准弹性模量 / 单元_IpS * (单元_ηS * PI ^ 2 * 单元_IwS / 单元_lS ^ 2 + 0.385 * 单元_ItS)
                                    Dim 单元_σELS As Single = 160000 * (单元_tw / 单元_hw) ^ 2

                                    'Dim 单元_A As Single = 单元_AS + 单元_AP
                                    Dim 单元_hc As Single = 单元_AS / 单元_A * 单元_hcS + 单元_AP / 单元_A * (-单元_tP / 2)
                                    Dim 单元_Io As Single = 单元_IOP + 单元_IOS
                                    Dim 单元_Ic As Single = 单元_Io - 单元_A * 单元_hc ^ 2
                                    'Dim 单元_σY As Single = 单元_AS / A * 单元_σYS + 单元_AP / 单元_A * 单元_σYP
                                    'Dim 单元_εY As Single = 单元_σY / 标准弹性模量

                                    '[注意正负!]
                                    Dim 单元_L As Single = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                                    '[注意正负!]
                                    单元_εO = 单元_L * χ_瞬时

                                    单元_εY = 单元_σY / 标准弹性模量

                                    '[注意正负!]
                                    单元_εR = 单元_εO / 单元_εY

                                    'Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                                    'Dim EP As Single = Φ * 单元_σY

                                    'wP / lP
                                    'Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)
                                    'Dim βPε As Single = βOP * 单元_εR ^ (1 / 2)
                                    'Dim BC As Single, FT As Single, WB As Single

                                    ''''''
                                    Dim 单元_Φε As Single = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))
                                    Dim 单元_σEPε As Single = 单元_Φε * 单元_σY

                                    Dim 单元_Rlt_EP As String = Format(单元_σEPε, "000.000")

                                    Select Case 单元_εR
                                        Case < 0
                                            Dim 单元_βPε As Single = 单元_βOP * (-单元_εR) ^ (1 / 2)
                                            Dim 单元_wEPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_wP, 单元_wP)
                                            Dim 单元_wE1Pε As Single = If(单元_βPε >= 1, 单元_wP / 单元_βPε, 单元_wP)
                                            Dim 单元_AEPε As Single = 单元_wEPε * 单元_tP
                                            Dim 单元_AE1Pε As Single = 单元_wE1Pε * 单元_tP
                                            Dim 单元_IcE1Pε As Single = 单元_wE1Pε * 单元_tP ^ 3 / 12
                                            Dim 单元_IoE1Pε As Single = 单元_IcE1Pε + 单元_AE1Pε * (-单元_tP / 2) ^ 2

                                            Dim 单元_βwε As Single = 单元_βOw * (-单元_εR) ^ (1 / 2)
                                            Dim 单元_hEwε As Single = If(单元_βwε >= 1.25, (2.25 / 单元_βwε - 1.25 / 单元_βwε ^ 2) * 单元_hw, 单元_hw)
                                            Dim 单元_AEwε As Single = 单元_hEwε * 单元_tw
                                            Dim 单元_AESε As Single = 单元_AEwε + 单元_Af

                                            Dim 单元_AEε As Single = 单元_AEPε + 单元_AS
                                            Dim 单元_AE1ε As Single = 单元_AE1Pε + 单元_AS
                                            Dim 单元_hcE1ε As Single = 单元_AE1Pε / 单元_AE1ε * (-单元_tP / 2) + 单元_AS / 单元_AE1ε * 单元_hcS
                                            Dim 单元_IoE1ε As Single = 单元_IoE1Pε + 单元_IOS
                                            Dim 单元_IcE1ε As Single = 单元_IoE1ε - 单元_AE1ε * 单元_hcE1ε ^ 2
                                            Dim 单元_lE1Pε As Single = 单元_hcE1ε - (-单元_tP / 2)
                                            Dim 单元_lE1Sε As Single = If(单元_tpS = "F" Or 单元_tpS = "B" Or 单元_tpS = "L3", 单元_hw - 单元_hcE1ε, 单元_hw + 单元_tf - 单元_hcE1ε)
                                            Dim 单元_σYE1ε As Single = 单元_AS * 单元_lE1Sε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYS + 单元_AE1Pε * 单元_lE1Pε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYP
                                            Dim 单元_σECε As Single = PI ^ 2 * 标准弹性模量 * 单元_IcE1ε / 单元_AEε / 单元_lS ^ 2
                                            Dim 单元_σCCε As Single = If(单元_σECε <= 单元_σYE1ε / 2 * (-单元_εR), 单元_σECε / (-单元_εR), 单元_σYE1ε * (1 - 单元_σYE1ε * (-单元_εR) / 4 / 单元_σECε))
                                            Dim 单元_σBCε As Single = 单元_Φε * 单元_σCCε * 单元_AEε / 单元_A

                                            Dim 单元_Rlt_BC As String = Format(单元_σBCε, "000.000")

                                            Dim 单元_σCTSε As Single = If(单元_σETS <= 单元_σYS / 2 * (-单元_εR), 单元_σETS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σETS))
                                            Dim 单元_σCPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_σYP, 单元_σYP)
                                            Dim 单元_σFTε As Single = 单元_Φε * (单元_AS * 单元_σCTSε + 单元_AP * 单元_σCPε) / 单元_A

                                            Dim 单元_Rlt_FT As String = Format(单元_σFTε, "000.000")

                                            Dim 单元_σCLSε As Single = If(单元_σELS <= 单元_σYS / 2 * (-单元_εR), 单元_σELS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σELS))
                                            Dim 单元_σWBε As Single = If(单元_tpS = "F", 单元_Φε * (单元_AS * 单元_σCLSε + 单元_AP * 单元_σCPε) / 单元_A, 单元_Φε * (单元_AESε * 单元_σYS + 单元_AEPε * 单元_σYP) / 单元_A)

                                            Dim 单元_Rlt_WB As String = Format(单元_σWBε, "000.000")

                                            ''''''
                                            单元_σO = Max(单元_σBCε, Max(单元_σFTε, 单元_σWBε)) ' 单元_σO = 单元_σEPε   ''''''NO BUCKLING
                                            'If 四级序数 = 1 Then Debug.Print(单元_Rlt_EP & ", " & 单元_Rlt_BC & ", " & 单元_Rlt_FT & ", " & 单元_Rlt_WB)
                                        Case > 0
                                            单元_σO = 单元_σEPε
                                        Case = 0
                                            单元_σO = 0
                                        Case Else

                                    End Select

                                    '考虑[单元_剩余_A]!
                                    单元_FO = 单元_σO * 单元_A / 1000000
                                    单元_MO = 单元_FO * 单元_L

                                    Select Case 单元_FO
                                        Case > 0
                                            ΣFP(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case < 0
                                            ΣFN(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Yc
                                            ΣFZN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_Zc
                                        Case = 0

                                        Case Else

                                    End Select

                                    ΣFO(一级序数, 二级序数, 三级序数) += 单元_FO
                                    ΣFA(一级序数, 二级序数, 三级序数) += Abs(单元_FO)
                                    ΣM(一级序数, 二级序数, 三级序数) += 单元_MO
                                Next
                            Case 5
                                'MsgBox("面板加强筋单元部分未完成!")
                            Case 6
                                For 四级序数 As UShort = 1 To 加筋板单元_总数
                                    Dim 单元_原始_Yc As Single = 加筋板单元_原始_Yc(四级序数) / 1000
                                    Dim 单元_原始_Zc As Single = 加筋板单元_原始_Zc(四级序数) / 1000

                                    Select Case 第一部分
                                        Case "OT"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                            End Select
                                        Case "BC"
                                            Select Case 第二部分
                                                Case "0"
                                                    Exit Select
                                                Case "1"
                                                    If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                                Case "2"
                                                    If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                                Case "3"
                                                    If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                            End Select
                                    End Select

                                    Dim 单元_原始_A As Single = 加筋板单元_原始_A(四级序数)
                                    Dim 单元_原始_σY As Single = 加筋板单元_原始_σY(四级序数)

                                    Dim 单元_剩余_Yc As Single = 加筋板单元_剩余_Yc(四级序数) / 1000
                                    Dim 单元_剩余_Zc As Single = 加筋板单元_剩余_Zc(四级序数) / 1000
                                    Dim 单元_剩余_A As Single = 加筋板单元_剩余_A(四级序数)
                                    Dim 单元_剩余_σY As Single = 加筋板单元_剩余_σY(四级序数)

                                    Dim 关联板格序数 As UShort = 加筋板单元_板格_序数(四级序数)

                                    '[注意正负!]
                                    单元_L = -(单元_原始_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_原始_Zc * Cos(α_瞬时)
                                    '[注意正负!]
                                    单元_εO = 单元_L * χ_瞬时

                                    单元_εY = 单元_原始_σY / 标准弹性模量

                                    '[注意正负!]
                                    单元_εR = 单元_εO / 单元_εY

                                    Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                                    Dim EP As Single = Φ * 单元_原始_σY

                                    'wP / lP
                                    Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)

                                    Select Case 单元_εR
                                        Case < 0
                                            Dim βPε As Single = βOP * (-单元_εR) ^ (1 / 2)
                                            Dim FB As Single = Φ * 单元_原始_σY * Min(1, 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数) * (2.25 / βPε - 1.25 / βPε ^ 2) + 0.1 * (1 - 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数)) * (1 + 1 / βPε ^ 2))
                                            单元_σO = FB
                                        Case > 0
                                            单元_σO = EP
                                        Case = 0
                                            单元_σO = 0
                                        Case Else

                                    End Select

                                    '考虑[单元_剩余_A]!
                                    单元_FO = 单元_σO * 单元_剩余_A / 1000000
                                    单元_MO = 单元_FO * 单元_L

                                    Select Case 单元_FO
                                        Case > 0
                                            ΣFP(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_原始_Yc
                                            ΣFZP(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_原始_Zc
                                        Case < 0
                                            ΣFN(一级序数, 二级序数, 三级序数) += 单元_FO
                                            ΣFYN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_原始_Yc
                                            ΣFZN(一级序数, 二级序数, 三级序数) += 单元_FO * 单元_原始_Zc
                                        Case = 0

                                        Case Else

                                    End Select

                                    ΣFO(一级序数, 二级序数, 三级序数) += 单元_FO
                                    ΣFA(一级序数, 二级序数, 三级序数) += Abs(单元_FO)
                                    ΣM(一级序数, 二级序数, 三级序数) += 单元_MO
                                Next
                            Case Else

                        End Select
                    Next

                    FYP(一级序数, 二级序数, 三级序数) = ΣFYP(一级序数, 二级序数, 三级序数) / ΣFP(一级序数, 二级序数, 三级序数)
                    FZP(一级序数, 二级序数, 三级序数) = ΣFZP(一级序数, 二级序数, 三级序数) / ΣFP(一级序数, 二级序数, 三级序数)

                    FYN(一级序数, 二级序数, 三级序数) = ΣFYN(一级序数, 二级序数, 三级序数) / ΣFN(一级序数, 二级序数, 三级序数)
                    FZN(一级序数, 二级序数, 三级序数) = ΣFZN(一级序数, 二级序数, 三级序数) / ΣFN(一级序数, 二级序数, 三级序数)

                    合力倾角_α = Atan((FZP(一级序数, 二级序数, 三级序数) - FZN(一级序数, 二级序数, 三级序数)) / (FYP(一级序数, 二级序数, 三级序数) - FYN(一级序数, 二级序数, 三级序数)))
                    向量积(三级序数) = Abs(Cos(水线倾角_α) * Cos(合力倾角_α) + Sin(水线倾角_α) * Sin(合力倾角_α))
                    Select Case 三级序数
                        Case 1
                            向量积(0) = 向量积(三级序数)
                            α_临界 = 三级序数

                            ΣFO(一级序数, 二级序数, 0) = ΣFO(一级序数, 二级序数, α_临界)
                            ΣFA(一级序数, 二级序数, 0) = ΣFA(一级序数, 二级序数, α_临界)
                            ΣM(一级序数, 二级序数, 0) = ΣM(一级序数, 二级序数, α_临界)
                        Case > 1
                            If 向量积(0) > 向量积(三级序数) Then
                                向量积(0) = 向量积(三级序数)
                                α_临界 = 三级序数

                                ΣFO(一级序数, 二级序数, 0) = ΣFO(一级序数, 二级序数, α_临界)
                                ΣFA(一级序数, 二级序数, 0) = ΣFA(一级序数, 二级序数, α_临界)
                                ΣM(一级序数, 二级序数, 0) = ΣM(一级序数, 二级序数, α_临界)
                            End If
                        Case Else

                    End Select
                Next    'α_循环
                α_初值 += (-1) ^ α_临界 * α_增量 * (α_临界 \ 2)

                α_候选(二级序数) = α_初值 * 180 / PI

                轴力相对差值(二级序数) = Abs(ΣFO(一级序数, 二级序数, 0)) / ΣFA(一级序数, 二级序数, 0)
                Select Case 二级序数
                    Case 1
                        轴力相对差值(0) = 轴力相对差值(二级序数)
                        γ_临界 = 二级序数

                        ΣM(一级序数, 0, 0) = ΣM(一级序数, γ_临界, 0)
                    Case > 1
                        If 轴力相对差值(0) > 轴力相对差值(二级序数) Then
                            轴力相对差值(0) = 轴力相对差值(二级序数)
                            γ_临界 = 二级序数

                            ΣM(一级序数, 0, 0) = ΣM(一级序数, γ_临界, 0)
                        End If
                    Case Else

                End Select

            Next    'γ_循环
            γ_初值 += (-1) ^ γ_临界 * γ_增量 * (γ_临界 \ 2)

            调试输出 = 调试输出 & Format(χ_瞬时, "0.000000") & Space(8) & Format(ΣM(一级序数, 0, 0), "00000.000") & Space(8) & Format(γ_初值, "000.000") & Space(8) & Format(α_候选(γ_临界), "000.000")

            Debug.Print(调试输出)

            Chart3.Series("弯矩曲率曲线").Points.AddXY(χ_瞬时, ΣM(一级序数, 0, 0))
            DataGridView3.Rows.Add(一级序数, χ_瞬时, ΣM(一级序数, 0, 0))

            If Abs(ΣM(0, 0, 0)) < Abs(ΣM(一级序数, 0, 0)) Then
                ΣM(0, 0, 0) = ΣM(一级序数, 0, 0)
                χ_临界 = 一级序数
            End If
            'Select Case 一级序数
            '    Case 1
            '        ΣM(0, 0, 0) = ΣM(一级序数, 0, 0)
            '        χ_临界 = 一级序数
            '    Case > 1
            '        If Abs(ΣM(0, 0, 0)) < Abs(ΣM(一级序数, 0, 0)) Then
            '            ΣM(0, 0, 0) = ΣM(一级序数, 0, 0)
            '            χ_临界 = 一级序数
            '        End If
            'End Select
        Next    'χ_循环

        χ_初值 = Abs(χ_初值) + χ_临界 * Abs(χ_增量)
        Debug.Print("全截面_ΣM(0, 0, 0) = " & ΣM(0, 0, 0) & " [MN.m]")
        Debug.Print("χ_临界 = " & χ_初值 & " [1/m]")
    End Sub

    Private Sub 增量解析法中拱计算()
        Dim Δχy As Single = 0.000001
        Dim Δχz As Single = 0.000001
        Dim ΣDA(500) As Single
        Dim ΣDAY(500) As Single, ΣDAZ(500) As Single
        Dim Yc(500) As Single, Zc(500) As Single
        For 一级序数 As UShort = 0 To χ_总数 * 4 - 1
            For 单元对象类型序数 As UShort = 1 To 6
                Select Case 单元对象类型序数
                    Case 1
                        ReDim 硬角单元_L(硬角单元_总数), 硬角单元_εO(硬角单元_总数), 硬角单元_εY(硬角单元_总数), 硬角单元_εR(硬角单元_总数)
                        For 二级序数 As UShort = 1 To 硬角单元_总数
                            Select Case 一级序数
                                Case 0
                                    硬角单元_L(二级序数) = (-硬角单元_Yc(二级序数) * Sin(γ_瞬时) + (硬角单元_Zc(二级序数) - 全截面_Zc) * Cos(γ_瞬时)) / 1000
                                Case Else
                                    硬角单元_L(二级序数) = (硬角单元_Zc(二级序数) - Zc(一级序数 - 1)) / 1000
                            End Select

                            硬角单元_εO(二级序数) = 硬角单元_L(二级序数) * Δχy * (一级序数 + 1)
                            硬角单元_εY(二级序数) = 硬角单元_σY(二级序数) / 标准弹性模量
                            硬角单元_εR(二级序数) = 硬角单元_εO(二级序数) / 硬角单元_εY(二级序数)

                            Select Case 硬角单元_L(二级序数)
                                Case > 0
                                    ΣDA(一级序数) += 硬角单元_D(二级序数, Int(Abs(硬角单元_εR(二级序数)) * 100) \ 1) * 硬角单元_A(二级序数)
                                    ΣDAY(一级序数) += 硬角单元_D(二级序数, Int(Abs(硬角单元_εR(二级序数)) * 100) \ 1) * 硬角单元_A(二级序数) * 硬角单元_Yc(二级序数)
                                    ΣDAZ(一级序数) += 硬角单元_D(二级序数, Int(Abs(硬角单元_εR(二级序数)) * 100) \ 1) * 硬角单元_A(二级序数) * 硬角单元_Zc(二级序数)
                                Case < 0
                                    ΣDA(一级序数) += 硬角单元_D(二级序数, Int(Abs(硬角单元_εR(二级序数)) * 100) \ 1) * 硬角单元_A(二级序数)
                                    ΣDAY(一级序数) += 硬角单元_D(二级序数, Int(Abs(硬角单元_εR(二级序数)) * 100) \ 1) * 硬角单元_A(二级序数) * 硬角单元_Yc(二级序数)
                                    ΣDAZ(一级序数) += 硬角单元_D(二级序数, Int(Abs(硬角单元_εR(二级序数)) * 100) \ 1) * 硬角单元_A(二级序数) * 硬角单元_Zc(二级序数)
                                Case Else

                            End Select
                        Next
                    Case 2
                        '
                    Case 3
                        ReDim 特别硬角单元_L(特别硬角单元_总数), 特别硬角单元_εO(特别硬角单元_总数), 特别硬角单元_εY(特别硬角单元_总数), 特别硬角单元_εR(特别硬角单元_总数)
                        For 二级序数 As UShort = 1 To 特别硬角单元_总数
                            Select Case 一级序数
                                Case 0
                                    特别硬角单元_L(二级序数) = (特别硬角单元_Zc(二级序数) - 全截面_Zc) / 1000
                                Case Else
                                    特别硬角单元_L(二级序数) = (特别硬角单元_Zc(二级序数) - Zc(一级序数 - 1)) / 1000
                            End Select

                            特别硬角单元_εO(二级序数) = 特别硬角单元_L(二级序数) * Δχy * (一级序数 + 1)
                            特别硬角单元_εY(二级序数) = 特别硬角单元_σY(二级序数) / 标准弹性模量
                            特别硬角单元_εR(二级序数) = 特别硬角单元_εO(二级序数) / 特别硬角单元_εY(二级序数)

                            Select Case 特别硬角单元_L(二级序数)
                                Case > 0
                                    ΣDA(一级序数) += 特别硬角单元_D(二级序数, Int(Abs(特别硬角单元_εR(二级序数)) * 100) \ 1) * 特别硬角单元_A(二级序数)
                                    ΣDAY(一级序数) += 特别硬角单元_D(二级序数, Int(Abs(特别硬角单元_εR(二级序数)) * 100) \ 1) * 特别硬角单元_A(二级序数) * 特别硬角单元_Yc(二级序数)
                                    ΣDAZ(一级序数) += 特别硬角单元_D(二级序数, Int(Abs(特别硬角单元_εR(二级序数)) * 100) \ 1) * 特别硬角单元_A(二级序数) * 特别硬角单元_Zc(二级序数)
                                Case < 0
                                    ΣDA(一级序数) += 特别硬角单元_D(二级序数, Int(Abs(特别硬角单元_εR(二级序数)) * 100) \ 1) * 特别硬角单元_A(二级序数)
                                    ΣDAY(一级序数) += 特别硬角单元_D(二级序数, Int(Abs(特别硬角单元_εR(二级序数)) * 100) \ 1) * 特别硬角单元_A(二级序数) * 特别硬角单元_Yc(二级序数)
                                    ΣDAZ(一级序数) += 特别硬角单元_D(二级序数, Int(Abs(特别硬角单元_εR(二级序数)) * 100) \ 1) * 特别硬角单元_A(二级序数) * 特别硬角单元_Zc(二级序数)
                                Case Else

                            End Select
                        Next
                    Case 4
                        ReDim 加强筋单元_L(加强筋单元_总数), 加强筋单元_εO(加强筋单元_总数), 加强筋单元_εY(加强筋单元_总数), 加强筋单元_εR(加强筋单元_总数)
                        For 二级序数 As UShort = 1 To 加强筋单元_总数
                            Select Case 一级序数
                                Case 0
                                    加强筋单元_L(二级序数) = (加强筋单元_Zc(二级序数) - 全截面_Zc) / 1000
                                Case Else
                                    加强筋单元_L(二级序数) = (加强筋单元_Zc(二级序数) - Zc(一级序数 - 1)) / 1000
                            End Select

                            加强筋单元_εO(二级序数) = 加强筋单元_L(二级序数) * Δχy * (一级序数 + 1)
                            加强筋单元_εY(二级序数) = 加强筋单元_σY(二级序数) / 标准弹性模量
                            加强筋单元_εR(二级序数) = 加强筋单元_εO(二级序数) / 加强筋单元_εY(二级序数)

                            Select Case 加强筋单元_L(二级序数)
                                Case > 0
                                    ΣDA(一级序数) += 加强筋单元_屈服_D(二级序数, Int(Abs(加强筋单元_εR(二级序数)) * 100) \ 1) * 加强筋单元_A(二级序数)
                                    ΣDAY(一级序数) += 加强筋单元_屈服_D(二级序数, Int(Abs(加强筋单元_εR(二级序数)) * 100) \ 1) * 加强筋单元_A(二级序数) * 加强筋单元_Yc(二级序数)
                                    ΣDAZ(一级序数) += 加强筋单元_屈服_D(二级序数, Int(Abs(加强筋单元_εR(二级序数)) * 100) \ 1) * 加强筋单元_A(二级序数) * 加强筋单元_Zc(二级序数)
                                Case < 0
                                    ΣDA(一级序数) += 加强筋单元_屈曲_D(二级序数, Int(Abs(加强筋单元_εR(二级序数)) * 100) \ 1) * 加强筋单元_A(二级序数)
                                    ΣDAY(一级序数) += 加强筋单元_屈曲_D(二级序数, Int(Abs(加强筋单元_εR(二级序数)) * 100) \ 1) * 加强筋单元_A(二级序数) * 加强筋单元_Yc(二级序数)
                                    ΣDAZ(一级序数) += 加强筋单元_屈曲_D(二级序数, Int(Abs(加强筋单元_εR(二级序数)) * 100) \ 1) * 加强筋单元_A(二级序数) * 加强筋单元_Zc(二级序数)
                                Case Else

                            End Select
                        Next
                    Case 5
                        '
                    Case 6
                        '
                End Select
            Next
            Yc(一级序数) = ΣDAY(一级序数) / ΣDA(一级序数)
            Zc(一级序数) = ΣDAZ(一级序数) / ΣDA(一级序数)
        Next
    End Sub

    Private Sub 增量解析法双轴计算()
        'case 1
        Dim Δχy As Single = InputBox("Δχy", "χy_增量", "0.000001")
        Dim Δχz As Single
        Dim Δεoo As Single
        Dim χy As Single = 0
        Dim χz As Single = 0
        Dim εoo As Single = 0

        Dim α As Single

        For 一级序数 As UShort = 1 To 200
            Dim Yi As Single, Zi As Single, Ai As Single, σiY As Single, εiY As Single, εiO As Single, εiR As Single, Di As Single
            Dim ΣDA As Single = 0
            Dim ΣDAY As Single = 0
            Dim ΣDAZ As Single = 0
            Dim γc As Single
            Dim ζc As Single
            Dim ΣDAΔYΔY As Single = 0
            Dim ΣDAΔYΔZ As Single = 0
            Dim ΣDAΔZΔZ As Single = 0
            If 一级序数 = 1 Then
                χy = 0
                χz = 0
            End If
            For 单元对象类型序数 As UShort = 1 To 6
                Select Case 单元对象类型序数
                    Case 1
                        For 二级序数 As UShort = 1 To 硬角单元_总数
                            Yi = 硬角单元_Yc(二级序数) / 1000
                            Zi = 硬角单元_Zc(二级序数) / 1000
                            If Yi > 25 And Zi > 6 Then Continue For
                            Ai = 硬角单元_A(二级序数) / 1000000
                            σiY = 硬角单元_σY(二级序数)
                            εiY = σiY / 标准弹性模量
                            εiO = εoo + Zi * χy + Yi * χz
                            εiR = εiO / εiY
                            Di = 硬角单元_D(二级序数, Int(Abs(εiR) * 100) \ 1) / εiY

                            ΣDA += Di * Ai
                            ΣDAY += Di * Ai * Yi
                            ΣDAZ += Di * Ai * Zi
                        Next
                    Case 2

                    Case 3
                        For 二级序数 As UShort = 1 To 特别硬角单元_总数
                            Yi = 特别硬角单元_Yc(二级序数) / 1000
                            Zi = 特别硬角单元_Zc(二级序数) / 1000
                            If Yi > 25 And Zi > 6 Then Continue For
                            Ai = 特别硬角单元_A(二级序数) / 1000000
                            σiY = 特别硬角单元_σY(二级序数)
                            εiY = σiY / 标准弹性模量
                            εiO = εoo + Zi * χy + Yi * χz
                            εiR = εiO / εiY
                            Di = 特别硬角单元_D(二级序数, Int(Abs(εiR) * 100) \ 1) / εiY

                            ΣDA += Di * Ai
                            ΣDAY += Di * Ai * Yi
                            ΣDAZ += Di * Ai * Zi
                        Next
                    Case 4
                        For 二级序数 As UShort = 1 To 加强筋单元_总数
                            Yi = 加强筋单元_Yc(二级序数) / 1000
                            Zi = 加强筋单元_Zc(二级序数) / 1000
                            If Yi > 25 And Zi > 6 Then Continue For
                            Ai = 加强筋单元_A(二级序数) / 1000000
                            σiY = 加强筋单元_σY(二级序数)
                            εiY = σiY / 标准弹性模量
                            εiO = εoo + Zi * χy + Yi * χz
                            εiR = εiO / εiY
                            Select Case εiR
                                Case > 0
                                    Di = 加强筋单元_屈服_D(二级序数, Int(Abs(εiR) * 100) \ 1) / εiY
                                Case < 0
                                    Di = 加强筋单元_屈曲_D(二级序数, Int(Abs(εiR) * 100) \ 1) / εiY
                            End Select

                            ΣDA += Di * Ai
                            ΣDAY += Di * Ai * Yi
                            ΣDAZ += Di * Ai * Zi
                        Next
                End Select
            Next
            γc = ΣDAY / ΣDA
            ζc = ΣDAZ / ΣDA
            For 单元对象类型序数 As UShort = 1 To 6
                Select Case 单元对象类型序数
                    Case 1
                        For 二级序数 As UShort = 1 To 硬角单元_总数
                            Yi = 硬角单元_Yc(二级序数) / 1000
                            Zi = 硬角单元_Zc(二级序数) / 1000
                            If Yi > 25 And Zi > 6 Then Continue For
                            Ai = 硬角单元_A(二级序数) / 1000000
                            σiY = 硬角单元_σY(二级序数)
                            εiY = σiY / 标准弹性模量
                            εiO = εoo + Zi * χy + Yi * χz
                            εiR = εiO / εiY
                            Di = 硬角单元_D(二级序数, Int(Abs(εiR) * 100) \ 1) / εiY

                            ΣDAΔYΔY += Di * Ai * (Yi - γc) * (Yi - γc)
                            ΣDAΔYΔZ += Di * Ai * (Yi - γc) * (Zi - ζc)
                            ΣDAΔZΔZ += Di * Ai * (Zi - ζc) * (Zi - ζc)
                        Next
                    Case 2

                    Case 3
                        For 二级序数 As UShort = 1 To 特别硬角单元_总数
                            Yi = 特别硬角单元_Yc(二级序数) / 1000
                            Zi = 特别硬角单元_Zc(二级序数) / 1000
                            If Yi > 25 And Zi > 6 Then Continue For
                            Ai = 特别硬角单元_A(二级序数) / 1000000
                            σiY = 特别硬角单元_σY(二级序数)
                            εiY = σiY / 标准弹性模量
                            εiO = εoo + Zi * χy + Yi * χz
                            εiR = εiO / εiY
                            Di = 特别硬角单元_D(二级序数, Int(Abs(εiR) * 100) \ 1) / εiY

                            ΣDAΔYΔY += Di * Ai * (Yi - γc) * (Yi - γc)
                            ΣDAΔYΔZ += Di * Ai * (Yi - γc) * (Zi - ζc)
                            ΣDAΔZΔZ += Di * Ai * (Zi - ζc) * (Zi - ζc)
                        Next
                    Case 4
                        For 二级序数 As UShort = 1 To 加强筋单元_总数
                            Yi = 加强筋单元_Yc(二级序数) / 1000
                            Zi = 加强筋单元_Zc(二级序数) / 1000
                            If Yi > 25 And Zi > 6 Then Continue For
                            Ai = 加强筋单元_A(二级序数) / 1000000
                            σiY = 加强筋单元_σY(二级序数)
                            εiY = σiY / 标准弹性模量
                            εiO = εoo + Zi * χy + Yi * χz
                            εiR = εiO / εiY
                            Select Case εiR
                                Case > 0
                                    Di = 加强筋单元_屈服_D(二级序数, Int(Abs(εiR) * 100) \ 1) / εiY
                                Case < 0
                                    Di = 加强筋单元_屈曲_D(二级序数, Int(Abs(εiR) * 100) \ 1) / εiY
                            End Select

                            ΣDAΔYΔY += Di * Ai * (Yi - γc) * (Yi - γc)
                            ΣDAΔYΔZ += Di * Ai * (Yi - γc) * (Zi - ζc)
                            ΣDAΔZΔZ += Di * Ai * (Zi - ζc) * (Zi - ζc)
                        Next
                End Select
            Next
            Δχz = Δχy * (-ΣDAΔYΔZ / ΣDAΔYΔY)
            Δεoo = -ζc * Δχy - γc * Δχz
            χy += Δχy
            χz += Δχz
            εoo += Δεoo
            α = Atan(-χz / χy)
            Debug.Print(Format(一级序数, "000") & Space(4) & εoo & " + " & χz & " * y + " & χy & " * z = 0" & Space(4) & " OR " & "z - " & ζc & " = " & -χz / χy & " * (y - " & γc & ")")
        Next
        'case 2 & 3

    End Sub

    Private Sub 基于二分法的多角度增量迭代法()
        Dim Δζ As Single, Δα As Single
        Dim ζL As Single, ζU As Single, ζM As Single
        Dim αLL As Single, αLU As Single, αLM As Single,
            αUL As Single, αUU As Single, αUM As Single,
            αML As Single, αMU As Single, αMM As Single

        Dim ΣM As Single

        Δα = 0.01 * PI / 180
        Δζ = 0.01

        ζM = 全截面_Zc / 1000
        ζL = ζM - 5 'Δζ * 200 '
        ζU = ζM + 5 'Δζ * 200 '

        αLM = 水线倾角_α '0 '
        αLL = 水线倾角_α - 0 'PI / 4 '
        αLU = 水线倾角_α + 0 'PI / 4 '

        αUM = 水线倾角_α '0 '
        αUL = 水线倾角_α - 0 'PI / 4 '
        αUU = 水线倾角_α + 0 'PI / 4 '

        αMM = 水线倾角_α '0 '
        αML = 水线倾角_α - 0 'PI / 4 '
        αMU = 水线倾角_α + 0 'PI / 4 '

        For 一级序数 As UShort = 1 To χ_总数
            '[注意正负!]
            Select Case Mid(第五部分, 1, 1)
                Case "P"
                    χ_瞬时 = Abs(χ_初值) + 一级序数 * Abs(χ_增量)
                Case "N"
                    χ_瞬时 = -Abs(χ_初值) - 一级序数 * Abs(χ_增量)
            End Select

            Dim 轴力差值_L As Single, 轴力差值_U As Single, 轴力差值_M As Single

            Dim Flag_ζL As Boolean, Flag_ζU As Boolean
            Dim Flag_αLL As Boolean, Flag_αLU As Boolean
            Dim Flag_αUL As Boolean, Flag_αUU As Boolean
            Dim Flag_αML As Boolean, Flag_αMU As Boolean

            Flag_ζL = True
            Flag_ζU = True
            Flag_αLL = True
            Flag_αLU = True
            Flag_αUL = True
            Flag_αUU = True
            Flag_αML = True
            Flag_αMU = True

            Do
                If Flag_ζL = True Then
                    Do
                        Dim 向量积_LL As Single, 向量积_LU As Single, 向量积_LM As Single
                        If Flag_αLL = True Then
                            αζ判据(χ_瞬时, ζL, αLL, 向量积_LL)
                            Flag_αLL = Not Flag_αLL
                        End If

                        If Flag_αLU = True Then
                            αζ判据(χ_瞬时, ζL, αLU, 向量积_LU)
                            Flag_αLU = Not Flag_αLU
                        End If

                        αζ判据(χ_瞬时, ζL, αLM, 向量积_LM)

                        If 向量积_LL * 向量积_LU < 0 Then
                            If 向量积_LL * 向量积_LM < 0 Then
                                αLU = αLM
                                向量积_LU = 向量积_LM
                            ElseIf 向量积_LU * 向量积_LM < 0 Then
                                αLL = αLM
                                向量积_LL = 向量积_LM
                            Else
                                'Error
                            End If
                            αLM = (αLL + αLU) / 2
                        Else
                            'Error
                        End If
                    Loop Until (Abs(αLL - αLU) < Δα)
                    αLL = αLM - 0 'PI / 36 'PI / 18 'PI / 9 '
                    αLU = αLM + 0 'PI / 36 'PI / 18 'PI / 9 '

                    Flag_αLL = Not Flag_αLL
                    Flag_αLU = Not Flag_αLU

                    ζ判据(χ_瞬时, ζL, αLM, 轴力差值_L)
                    Flag_ζL = Not Flag_ζL
                End If

                If Flag_ζU = True Then
                    Do
                        Dim 向量积_UL As Single, 向量积_UU As Single, 向量积_UM As Single

                        If Flag_αUL = True Then
                            αζ判据(χ_瞬时, ζU, αUL, 向量积_UL)
                            Flag_αUL = Not Flag_αUL
                        End If

                        If Flag_αUU = True Then
                            αζ判据(χ_瞬时, ζU, αUU, 向量积_UU)
                            Flag_αUU = Not Flag_αUU
                        End If

                        αζ判据(χ_瞬时, ζU, αUM, 向量积_UM)

                        If 向量积_UL * 向量积_UU < 0 Then
                            If 向量积_UL * 向量积_UM < 0 Then
                                αUU = αUM
                                向量积_UU = 向量积_UM
                            ElseIf 向量积_UU * 向量积_UM < 0 Then
                                αUL = αUM
                                向量积_UL = 向量积_UM
                            Else
                                'Error
                            End If
                            αUM = (αUL + αUU) / 2
                        Else
                            'Error
                        End If
                    Loop Until (Abs(αUL - αUU) < Δα)
                    αUL = αUM - 0 'PI / 36 'PI / 18 'PI / 9 '
                    αUU = αUM + 0 'PI / 36 'PI / 18 'PI / 9 '

                    Flag_αUL = Not Flag_αUL
                    Flag_αUU = Not Flag_αUU

                    ζ判据(χ_瞬时, ζU, αUM, 轴力差值_U)
                    Flag_ζU = Not Flag_ζU
                End If

                Do
                    Dim 向量积_ML As Single, 向量积_MU As Single, 向量积_MM As Single

                    If Flag_αML = True Then
                        αζ判据(χ_瞬时, ζM, αML, 向量积_ML)
                        Flag_αML = Not Flag_αML
                    End If

                    If Flag_αMU = True Then
                        αζ判据(χ_瞬时, ζM, αMU, 向量积_MU)
                        Flag_αMU = Not Flag_αMU
                    End If

                    αζ判据(χ_瞬时, ζM, αMM, 向量积_MM)

                    If 向量积_ML * 向量积_MU < 0 Then
                        If 向量积_ML * 向量积_MM < 0 Then
                            αMU = αMM
                            向量积_MU = 向量积_MM
                        ElseIf 向量积_MU * 向量积_MM < 0 Then
                            αML = αMM
                            向量积_ML = 向量积_MM
                        Else
                            If 向量积_MM = 0 Then Exit Do 'Error
                        End If
                        αMM = (αML + αMU) / 2
                    Else
                        'Error
                    End If
                Loop Until (Abs(αML - αMU) < Δα)
                αML = αMM - 0 'PI / 36 'PI / 18 'PI / 9 '
                αMU = αMM + 0 'PI / 36 'PI / 18 'PI / 9 '

                Flag_αML = Not Flag_αML
                Flag_αMU = Not Flag_αMU

                ζ判据(χ_瞬时, ζM, αMM, 轴力差值_M)

                If 轴力差值_L * 轴力差值_U < 0 Then
                    If 轴力差值_U * 轴力差值_M < 0 Then
                        ζL = ζM
                        αLM = αMM
                        轴力差值_L = 轴力差值_M
                    ElseIf 轴力差值_L * 轴力差值_M < 0 Then
                        ζU = ζM
                        αUM = αMM
                        轴力差值_U = 轴力差值_M
                    End If
                    ζM = (ζL + ζU) / 2
                Else
                    'Error
                    If Abs(轴力差值_L) < Abs(轴力差值_U) Then
                        If 轴力差值_L * 轴力差值_M < 0 Then
                            ζU = ζM
                            αUM = αMM
                            轴力差值_U = 轴力差值_M
                        End If
                    Else
                        If 轴力差值_U * 轴力差值_M < 0 Then
                            ζL = ζM
                            αLM = αMM
                            轴力差值_L = 轴力差值_M
                        End If
                    End If
                    ζM = (ζL + ζU) / 2
                End If
            Loop Until Abs(ζL - ζU) < Δζ
            最终弯矩ζ(χ_瞬时, ζM, αMM, ΣM)
            Debug.Print(Format(χ_瞬时, "0.000000") & ", " & Format(ζM, "000.000") & ", " & Format(αMM / PI * 180, "000.000") & ", " & Format(ΣM, "00000.000"))
            'Debug.Print("")

            ζL = ζM - 5 'Δζ * 100 'Δζ * 200 '10 '
            If ζL < 1 Then ζL = 1
            ζU = ζM + 5 'Δζ * 100 'Δζ * 200 '10 '
            Flag_ζL = Not Flag_ζL
            Flag_ζU = Not Flag_ζU
        Next

    End Sub

    Private Sub 基于二分法的第二多角度增量迭代法()
        Dim Δγ As Single, Δα As Single
        Dim γL As Single, γU As Single, γM As Single
        Dim αLL As Single, αLU As Single, αLM As Single,
            αUL As Single, αUU As Single, αUM As Single,
            αML As Single, αMU As Single, αMM As Single

        Dim ΣM As Single

        Δα = 0.01 * PI / 180
        Δγ = 0.01

        γM = 全截面_Yc / 1000
        γL = γM - 20 '5 '10
        γU = γM + 20 '5 '10

        αLM = 水线倾角_α
        αLL = 水线倾角_α - PI / 4
        αLU = 水线倾角_α + PI / 4

        αUM = 水线倾角_α
        αUL = 水线倾角_α - PI / 4
        αUU = 水线倾角_α + PI / 4

        αMM = 水线倾角_α
        αML = 水线倾角_α - PI / 4
        αMU = 水线倾角_α + PI / 4

        For 一级序数 As UShort = 1 To χ_总数
            '[注意正负!]
            Select Case Mid(第五部分, 1, 1)
                Case "P"
                    χ_瞬时 = Abs(χ_初值) + 一级序数 * Abs(χ_增量)
                Case "N"
                    χ_瞬时 = -Abs(χ_初值) - 一级序数 * Abs(χ_增量)
            End Select

            Dim 轴力差值_L As Single, 轴力差值_U As Single, 轴力差值_M As Single

            Dim Flag_γL As Boolean, Flag_γU As Boolean
            Dim Flag_αLL As Boolean, Flag_αLU As Boolean
            Dim Flag_αUL As Boolean, Flag_αUU As Boolean
            Dim Flag_αML As Boolean, Flag_αMU As Boolean

            Flag_γL = True
            Flag_γU = True
            Flag_αLL = True
            Flag_αLU = True
            Flag_αUL = True
            Flag_αUU = True
            Flag_αML = True
            Flag_αMU = True

            Do
                If Flag_γL = True Then
                    Do
                        Dim 向量积_LL As Single, 向量积_LU As Single, 向量积_LM As Single
                        If Flag_αLL = True Then
                            αγ判据(χ_瞬时, γL, αLL, 向量积_LL)
                            Flag_αLL = Not Flag_αLL
                        End If

                        If Flag_αLU = True Then
                            αγ判据(χ_瞬时, γL, αLU, 向量积_LU)
                            Flag_αLU = Not Flag_αLU
                        End If

                        αγ判据(χ_瞬时, γL, αLM, 向量积_LM)

                        If 向量积_LL * 向量积_LU < 0 Then
                            If 向量积_LL * 向量积_LM < 0 Then
                                αLU = αLM
                                向量积_LU = 向量积_LM
                            ElseIf 向量积_LU * 向量积_LM < 0 Then
                                αLL = αLM
                                向量积_LL = 向量积_LM
                            Else
                                'Error
                            End If
                            αLM = (αLL + αLU) / 2
                        Else
                            'Error
                        End If
                    Loop Until (Abs(αLL - αLU) < Δα)
                    αLL = αLM - PI / 12 'PI / 18
                    αLU = αLM + PI / 12 'PI / 18

                    Flag_αLL = Not Flag_αLL
                    Flag_αLU = Not Flag_αLU

                    γ判据(χ_瞬时, γL, αLM, 轴力差值_L)
                    Flag_γL = Not Flag_γL
                End If

                If Flag_γU = True Then
                    Do
                        Dim 向量积_UL As Single, 向量积_UU As Single, 向量积_UM As Single

                        If Flag_αUL = True Then
                            αγ判据(χ_瞬时, γU, αUL, 向量积_UL)
                            Flag_αUL = Not Flag_αUL
                        End If

                        If Flag_αUU = True Then
                            αγ判据(χ_瞬时, γU, αUU, 向量积_UU)
                            Flag_αUU = Not Flag_αUU
                        End If

                        αγ判据(χ_瞬时, γU, αUM, 向量积_UM)

                        If 向量积_UL * 向量积_UU < 0 Then
                            If 向量积_UL * 向量积_UM < 0 Then
                                αUU = αUM
                                向量积_UU = 向量积_UM
                            ElseIf 向量积_UU * 向量积_UM < 0 Then
                                αUL = αUM
                                向量积_UL = 向量积_UM
                            Else
                                'Error
                            End If
                            αUM = (αUL + αUU) / 2
                        Else
                            'Error
                        End If
                    Loop Until (Abs(αUL - αUU) < Δα)
                    αUL = αUM - PI / 12 'PI / 18
                    αUU = αUM + PI / 12 'PI / 18

                    Flag_αUL = Not Flag_αUL
                    Flag_αUU = Not Flag_αUU

                    γ判据(χ_瞬时, γU, αUM, 轴力差值_U)
                    Flag_γU = Not Flag_γU
                End If

                Do
                    Dim 向量积_ML As Single, 向量积_MU As Single, 向量积_MM As Single

                    If Flag_αML = True Then
                        αγ判据(χ_瞬时, γM, αML, 向量积_ML)
                        Flag_αML = Not Flag_αML
                    End If

                    If Flag_αMU = True Then
                        αγ判据(χ_瞬时, γM, αMU, 向量积_MU)
                        Flag_αMU = Not Flag_αMU
                    End If

                    αγ判据(χ_瞬时, γM, αMM, 向量积_MM)

                    If 向量积_ML * 向量积_MU < 0 Then
                        If 向量积_ML * 向量积_MM < 0 Then
                            αMU = αMM
                            向量积_MU = 向量积_MM
                        ElseIf 向量积_MU * 向量积_MM < 0 Then
                            αML = αMM
                            向量积_ML = 向量积_MM
                        Else
                            If 向量积_MM = 0 Then Exit Do 'Error
                        End If
                        αMM = (αML + αMU) / 2
                    Else
                        'Error
                    End If
                Loop Until (Abs(αML - αMU) < Δα)
                αML = αMM - PI / 12 'PI / 18
                αMU = αMM + PI / 12 'PI / 18

                Flag_αML = Not Flag_αML
                Flag_αMU = Not Flag_αMU

                γ判据(χ_瞬时, γM, αMM, 轴力差值_M)

                If 轴力差值_L * 轴力差值_U < 0 Then
                    If 轴力差值_L * 轴力差值_M < 0 Then
                        γU = γM
                        αUM = αMM
                        轴力差值_U = 轴力差值_M
                    ElseIf 轴力差值_U * 轴力差值_M < 0 Then
                        γL = γM
                        αLM = αMM
                        轴力差值_L = 轴力差值_M
                    End If
                    γM = (γL + γU) / 2
                Else
                    'Error
                    If Abs(轴力差值_L) < Abs(轴力差值_U) Then
                        If 轴力差值_L * 轴力差值_M < 0 Then
                            γU = γM
                            αUM = αMM
                            轴力差值_U = 轴力差值_M
                        End If
                    Else
                        If 轴力差值_U * 轴力差值_M < 0 Then
                            γL = γM
                            αLM = αMM
                            轴力差值_L = 轴力差值_M
                        End If
                    End If
                    γM = (γL + γU) / 2
                End If
            Loop Until Abs(γL - γU) < Δγ
            最终弯矩γ(χ_瞬时, γM, αMM, ΣM)
            Debug.Print(Format(χ_瞬时, "0.000000") & ", " & Format(γM, "000.000") & ", " & Format(αMM / PI * 180, "000.000") & ", " & Format(ΣM, "00000.000"))

            γL = γM - 5
            'If γL < -28 Then γL = -28
            γU = γM + 5
            'If γU > 28 Then γU = 28
            Flag_γL = Not Flag_γL
            Flag_γU = Not Flag_γU
        Next

    End Sub

    Private Sub ζ判据(ByVal χ_瞬时 As Single, ByVal ζ_瞬时 As Single, ByVal α_瞬时 As Single, ByRef 判据参数1 As Single)
        Dim 截面_FO As Single, 截面_FA As Single,
            截面_FP As Single, 截面_FN As Single,
            截面_FYP As Single, 截面_FYN As Single,
            截面_FZP As Single, 截面_FZN As Single,
            截面_MO As Single

        截面_FO = 0
        截面_FA = 0
        截面_FP = 0
        截面_FN = 0
        截面_FYP = 0
        截面_FYN = 0
        截面_FZP = 0
        截面_FZN = 0
        截面_MO = 0

        For 单元类型 As UShort = 1 To 6
            Select Case 单元类型
                Case 1
                    For 四级序数 As UShort = 1 To 硬角单元_总数
                        单元_Yc = 硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 硬角单元_A(四级序数) / 1000000
                        单元_σY = 硬角单元_σY(四级序数)
                        单元_L = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 2
                    '
                Case 3
                    For 四级序数 As UShort = 1 To 特别硬角单元_总数
                        单元_Yc = 特别硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 特别硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 特别硬角单元_A(四级序数) / 1000000
                        单元_σY = 特别硬角单元_σY(四级序数)
                        单元_L = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 4
                    For 四级序数 As UShort = 1 To 加强筋单元_总数
                        Dim 单元_Yc As Single = 加强筋单元_Yc(四级序数) / 1000
                        Dim 单元_Zc As Single = 加强筋单元_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_A As Single = 加强筋单元_A(四级序数)
                        Dim 单元_σY As Single = 加强筋单元_σY(四级序数)

                        Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(四级序数)
                        Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)

                        Dim 单元_σYP As Single = 加强筋单元_σYP(四级序数)
                        Dim 单元_lP As Single = 加强筋单元_lP(四级序数)
                        Dim 单元_wP As Single = 加强筋单元_wP(四级序数)
                        Dim 单元_tP As Single = 加强筋单元_tP(四级序数)
                        Dim 单元_σYS As Single = 加强筋单元_σYS(四级序数)
                        Dim 单元_lS As Single = 加强筋单元_lS(四级序数)
                        Dim 单元_hw As Single = 加强筋单元_hw(四级序数)
                        Dim 单元_tw As Single = 加强筋单元_tw(四级序数)
                        Dim 单元_wf As Single = 加强筋单元_wf(四级序数)
                        Dim 单元_tf As Single = 加强筋单元_tf(四级序数)
                        Dim 单元_dx As Single = 加强筋单元_dx(四级序数)
                        Dim 单元_tpS As String = 加强筋单元_tpS(四级序数)
                        Dim 单元_mk As Boolean = 加强筋单元_mk(四级序数)
                        'Dim 单元_σETS As Single = 加强筋单元_σETS(四级序数)
                        'Dim 单元_σELS As Single = 加强筋单元_σELS(四级序数)

                        Dim 单元_εYP As Single = 单元_σYP / 标准弹性模量
                        Dim 单元_AP As Single = 单元_wP * 单元_tP
                        Dim 单元_ICP As Single = 单元_wP * 单元_tP ^ 3 / 12
                        Dim 单元_IOP As Single = 单元_ICP + 单元_AP * (-单元_tP / 2) ^ 2
                        Dim 单元_βOP As Single = 单元_wP / 单元_tP * 单元_εYP ^ (1 / 2)
                        Dim 单元_wEoP As Single = If(单元_βOP >= 1.25, (2.25 / 单元_βOP - 1.25 / 单元_βOP ^ 2) * 单元_wP, 单元_wP)

                        Dim 单元_εYS As Single = 单元_σYS / 标准弹性模量
                        Dim 单元_Aw As Single = 单元_hw * 单元_tw
                        Dim 单元_Icw As Single = 单元_tw * 单元_hw ^ 3 / 12
                        Dim 单元_Iow As Single = 单元_Icw + 单元_Aw * (单元_hw / 2) ^ 2
                        Dim 单元_βOw As Single = 单元_hw / 单元_tw * 单元_εYS ^ (1 / 2)
                        Dim 单元_hEow As Single = If(单元_βOw >= 1.25, (2.25 / 单元_βOw - 1.25 / 单元_βOw ^ 2) * 单元_hw, 单元_hw)
                        Dim 单元_df As Single = If(单元_tpS = "F", 单元_hw, If(单元_tpS = "B", 单元_hw - 单元_tf / 2, If(单元_tpS = "T" Or 单元_tpS = "L1" Or 单元_tpS = "L2", 单元_hw + 单元_tf / 2, 单元_hw - 单元_dx - 单元_tf / 2)))
                        Dim 单元_Af As Single = 单元_wf * 单元_tf
                        Dim 单元_Icf As Single = 单元_wf * 单元_tf ^ 3 / 12
                        Dim 单元_Iof As Single = 单元_Icf + 单元_Af * 单元_df ^ 2
                        Dim 单元_AS As Single = 单元_Aw + 单元_Af
                        Dim 单元_hcS As Single = 单元_Aw / 单元_AS * 单元_hw / 2 + 单元_Af / 单元_AS * 单元_df
                        Dim 单元_IOS As Single = 单元_Iow + 单元_Iof
                        Dim 单元_ICS As Single = 单元_Icw + 单元_Aw * (单元_hw / 2 - 单元_hcS) ^ 2 + 单元_Icf + 单元_Af * (单元_df - 单元_hcS) ^ 2
                        Dim 单元_IpS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw / 3, 单元_Aw * (单元_df - 单元_tf / 2) ^ 2 / 3 + 单元_Af * 单元_df ^ 2)
                        Dim 单元_ItS As Single = If(单元_tpS = "F", 单元_hw * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / 单元_hw), (单元_df - 单元_tf / 2) * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / (单元_df - 单元_tf / 2)) + 单元_wf * 单元_tf ^ 3 / 3 * (1 - 0.63 * 单元_tf / 单元_wf))
                        Dim 单元_IwS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw ^ 3 / 36, If(单元_tpS = "B" Or 单元_tpS = "L1" Or 单元_tpS = "L2" Or 单元_tpS = "L3", 单元_Af * 单元_df ^ 2 * 单元_wf ^ 2 / 12 * (单元_Af + 2.6 * 单元_Aw) / (单元_Af * 单元_Aw), 单元_wf ^ 3 * 单元_tf * 单元_df ^ 2 / 12))
                        Dim 单元_ηS As Single = 1 + (单元_lS / PI) ^ 2 / (单元_IwS * (0.75 * 单元_wP / 单元_tP ^ 3 + (单元_df - 单元_tf / 2) / 单元_tw ^ 3)) ^ (1 / 2)
                        Dim 单元_σETS As Single = 标准弹性模量 / 单元_IpS * (单元_ηS * PI ^ 2 * 单元_IwS / 单元_lS ^ 2 + 0.385 * 单元_ItS)
                        Dim 单元_σELS As Single = 160000 * (单元_tw / 单元_hw) ^ 2

                        'Dim 单元_A As Single = 单元_AS + 单元_AP
                        Dim 单元_hc As Single = 单元_AS / 单元_A * 单元_hcS + 单元_AP / 单元_A * (-单元_tP / 2)
                        Dim 单元_Io As Single = 单元_IOP + 单元_IOS
                        Dim 单元_Ic As Single = 单元_Io - 单元_A * 单元_hc ^ 2
                        'Dim 单元_σY As Single = 单元_AS / A * 单元_σYS + 单元_AP / 单元_A * 单元_σYP
                        'Dim 单元_εY As Single = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        Dim 单元_L As Single = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY

                        'Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        'Dim EP As Single = Φ * 单元_σY

                        'wP / lP
                        'Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)
                        'Dim βPε As Single = βOP * 单元_εR ^ (1 / 2)
                        'Dim BC As Single, FT As Single, WB As Single

                        '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING
                        Dim 单元_Φε As Single = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))
                        Dim 单元_σEPε As Single = 单元_Φε * 单元_σY

                        Dim 单元_Rlt_EP As String = Format(单元_σEPε, "000.000")

                        Select Case 单元_εR
                            Case < 0
                                Dim 单元_βPε As Single = 单元_βOP * (-单元_εR) ^ (1 / 2)
                                Dim 单元_wEPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_wP, 单元_wP)
                                Dim 单元_wE1Pε As Single = If(单元_βPε >= 1, 单元_wP / 单元_βPε, 单元_wP)
                                Dim 单元_AEPε As Single = 单元_wEPε * 单元_tP
                                Dim 单元_AE1Pε As Single = 单元_wE1Pε * 单元_tP
                                Dim 单元_IcE1Pε As Single = 单元_wE1Pε * 单元_tP ^ 3 / 12
                                Dim 单元_IoE1Pε As Single = 单元_IcE1Pε + 单元_AE1Pε * (-单元_tP / 2) ^ 2

                                Dim 单元_βwε As Single = 单元_βOw * (-单元_εR) ^ (1 / 2)
                                Dim 单元_hEwε As Single = If(单元_βwε >= 1.25, (2.25 / 单元_βwε - 1.25 / 单元_βwε ^ 2) * 单元_hw, 单元_hw)
                                Dim 单元_AEwε As Single = 单元_hEwε * 单元_tw
                                Dim 单元_AESε As Single = 单元_AEwε + 单元_Af

                                Dim 单元_AEε As Single = 单元_AEPε + 单元_AS
                                Dim 单元_AE1ε As Single = 单元_AE1Pε + 单元_AS
                                Dim 单元_hcE1ε As Single = 单元_AE1Pε / 单元_AE1ε * (-单元_tP / 2) + 单元_AS / 单元_AE1ε * 单元_hcS
                                Dim 单元_IoE1ε As Single = 单元_IoE1Pε + 单元_IOS
                                Dim 单元_IcE1ε As Single = 单元_IoE1ε - 单元_AE1ε * 单元_hcE1ε ^ 2
                                Dim 单元_lE1Pε As Single = 单元_hcE1ε - (-单元_tP / 2)
                                Dim 单元_lE1Sε As Single = If(单元_tpS = "F" Or 单元_tpS = "B" Or 单元_tpS = "L3", 单元_hw - 单元_hcE1ε, 单元_hw + 单元_tf - 单元_hcE1ε)
                                Dim 单元_σYE1ε As Single = 单元_AS * 单元_lE1Sε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYS + 单元_AE1Pε * 单元_lE1Pε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYP
                                Dim 单元_σECε As Single = PI ^ 2 * 标准弹性模量 * 单元_IcE1ε / 单元_AEε / 单元_lS ^ 2
                                Dim 单元_σCCε As Single = If(单元_σECε <= 单元_σYE1ε / 2 * (-单元_εR), 单元_σECε / (-单元_εR), 单元_σYE1ε * (1 - 单元_σYE1ε * (-单元_εR) / 4 / 单元_σECε))
                                Dim 单元_σBCε As Single = 单元_Φε * 单元_σCCε * 单元_AEε / 单元_A

                                Dim 单元_Rlt_BC As String = Format(单元_σBCε, "000.000")

                                Dim 单元_σCTSε As Single = If(单元_σETS <= 单元_σYS / 2 * (-单元_εR), 单元_σETS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σETS))
                                Dim 单元_σCPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_σYP, 单元_σYP)
                                Dim 单元_σFTε As Single = 单元_Φε * (单元_AS * 单元_σCTSε + 单元_AP * 单元_σCPε) / 单元_A

                                Dim 单元_Rlt_FT As String = Format(单元_σFTε, "000.000")

                                Dim 单元_σCLSε As Single = If(单元_σELS <= 单元_σYS / 2 * (-单元_εR), 单元_σELS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σELS))
                                Dim 单元_σWBε As Single = If(单元_tpS = "F", 单元_Φε * (单元_AS * 单元_σCLSε + 单元_AP * 单元_σCPε) / 单元_A, 单元_Φε * (单元_AESε * 单元_σYS + 单元_AEPε * 单元_σYP) / 单元_A)

                                Dim 单元_Rlt_WB As String = Format(单元_σWBε, "000.000")

                                ''''''
                                单元_σO = Max(单元_σBCε, Max(单元_σFTε, 单元_σWBε)) ' 单元_σO = 单元_σEPε   ''''''NO BUCKLING
                                            'If 四级序数 = 1 Then Debug.Print(单元_Rlt_EP & ", " & 单元_Rlt_BC & ", " & 单元_Rlt_FT & ", " & 单元_Rlt_WB)
                            Case > 0
                                单元_σO = 单元_σEPε
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 5
                    '
                Case 6
                    For 四级序数 As UShort = 1 To 加筋板单元_总数
                        Dim 单元_原始_Yc As Single = 加筋板单元_原始_Yc(四级序数) / 1000
                        Dim 单元_原始_Zc As Single = 加筋板单元_原始_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_原始_A As Single = 加筋板单元_原始_A(四级序数)
                        Dim 单元_原始_σY As Single = 加筋板单元_原始_σY(四级序数)

                        Dim 单元_剩余_Yc As Single = 加筋板单元_剩余_Yc(四级序数) / 1000
                        Dim 单元_剩余_Zc As Single = 加筋板单元_剩余_Zc(四级序数) / 1000
                        Dim 单元_剩余_A As Single = 加筋板单元_剩余_A(四级序数)
                        Dim 单元_剩余_σY As Single = 加筋板单元_剩余_σY(四级序数)

                        Dim 关联板格序数 As UShort = 加筋板单元_板格_序数(四级序数)

                        '[注意正负!]
                        单元_L = -单元_原始_Yc * Sin(α_瞬时) + (单元_原始_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_原始_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY
                        '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING

                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        Dim EP As Single = Φ * 单元_原始_σY

                        'wP / lP
                        Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)

                        Select Case 单元_εR
                            Case < 0
                                Dim βPε As Single = βOP * (-单元_εR) ^ (1 / 2)
                                Dim FB As Single = Φ * 单元_原始_σY * Min(1, 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数) * (2.25 / βPε - 1.25 / βPε ^ 2) + 0.1 * (1 - 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数)) * (1 + 1 / βPε ^ 2))
                                单元_σO = FB' 单元_σO = EP   ''''''NO BUCKLING
                            Case > 0
                                单元_σO = EP
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_剩余_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case Else
                    '
            End Select
        Next

        'Dim 截面_YFP As Single, 截面_YFN As Single,
        '    截面_ZFP As Single, 截面_ZFN As Single
        'Dim α_RF As Single
        'Dim VP As Single
        '截面_YFP = 截面_FYP / 截面_FP
        '截面_YFN = 截面_FYN / 截面_FN
        '截面_ZFP = 截面_FZP / 截面_FP
        '截面_ZFN = 截面_FZN / 截面_FN
        'α_RF = Atan((截面_ZFP - 截面_ZFN) / (截面_YFP - 截面_YFN))
        'VP = Cos(水线倾角_α) * Cos(α_RF) + Sin(水线倾角_α) * Sin(α_RF)
        判据参数1 = 截面_FO
    End Sub

    Private Sub αζ判据(ByVal χ_瞬时 As Single, ByVal ζ_瞬时 As Single, ByVal α_瞬时 As Single, ByRef 判据参数2 As Single)
        Dim 截面_FO As Single, 截面_FA As Single,
            截面_FP As Single, 截面_FN As Single,
            截面_FYP As Single, 截面_FYN As Single,
            截面_FZP As Single, 截面_FZN As Single,
            截面_MO As Single

        截面_FO = 0
        截面_FA = 0
        截面_FP = 0
        截面_FN = 0
        截面_FYP = 0
        截面_FYN = 0
        截面_FZP = 0
        截面_FZN = 0
        截面_MO = 0

        For 单元类型 As UShort = 1 To 6
            Select Case 单元类型
                Case 1
                    For 四级序数 As UShort = 1 To 硬角单元_总数
                        单元_Yc = 硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 硬角单元_A(四级序数) / 1000000
                        单元_σY = 硬角单元_σY(四级序数)
                        单元_L = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 2
                    '
                Case 3
                    For 四级序数 As UShort = 1 To 特别硬角单元_总数
                        单元_Yc = 特别硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 特别硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 特别硬角单元_A(四级序数) / 1000000
                        单元_σY = 特别硬角单元_σY(四级序数)
                        单元_L = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 4
                    For 四级序数 As UShort = 1 To 加强筋单元_总数
                        Dim 单元_Yc As Single = 加强筋单元_Yc(四级序数) / 1000
                        Dim 单元_Zc As Single = 加强筋单元_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_A As Single = 加强筋单元_A(四级序数)
                        Dim 单元_σY As Single = 加强筋单元_σY(四级序数)

                        Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(四级序数)
                        Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)

                        Dim 单元_σYP As Single = 加强筋单元_σYP(四级序数)
                        Dim 单元_lP As Single = 加强筋单元_lP(四级序数)
                        Dim 单元_wP As Single = 加强筋单元_wP(四级序数)
                        Dim 单元_tP As Single = 加强筋单元_tP(四级序数)
                        Dim 单元_σYS As Single = 加强筋单元_σYS(四级序数)
                        Dim 单元_lS As Single = 加强筋单元_lS(四级序数)
                        Dim 单元_hw As Single = 加强筋单元_hw(四级序数)
                        Dim 单元_tw As Single = 加强筋单元_tw(四级序数)
                        Dim 单元_wf As Single = 加强筋单元_wf(四级序数)
                        Dim 单元_tf As Single = 加强筋单元_tf(四级序数)
                        Dim 单元_dx As Single = 加强筋单元_dx(四级序数)
                        Dim 单元_tpS As String = 加强筋单元_tpS(四级序数)
                        Dim 单元_mk As Boolean = 加强筋单元_mk(四级序数)
                        'Dim 单元_σETS As Single = 加强筋单元_σETS(四级序数)
                        'Dim 单元_σELS As Single = 加强筋单元_σELS(四级序数)

                        Dim 单元_εYP As Single = 单元_σYP / 标准弹性模量
                        Dim 单元_AP As Single = 单元_wP * 单元_tP
                        Dim 单元_ICP As Single = 单元_wP * 单元_tP ^ 3 / 12
                        Dim 单元_IOP As Single = 单元_ICP + 单元_AP * (-单元_tP / 2) ^ 2
                        Dim 单元_βOP As Single = 单元_wP / 单元_tP * 单元_εYP ^ (1 / 2)
                        Dim 单元_wEoP As Single = If(单元_βOP >= 1.25, (2.25 / 单元_βOP - 1.25 / 单元_βOP ^ 2) * 单元_wP, 单元_wP)

                        Dim 单元_εYS As Single = 单元_σYS / 标准弹性模量
                        Dim 单元_Aw As Single = 单元_hw * 单元_tw
                        Dim 单元_Icw As Single = 单元_tw * 单元_hw ^ 3 / 12
                        Dim 单元_Iow As Single = 单元_Icw + 单元_Aw * (单元_hw / 2) ^ 2
                        Dim 单元_βOw As Single = 单元_hw / 单元_tw * 单元_εYS ^ (1 / 2)
                        Dim 单元_hEow As Single = If(单元_βOw >= 1.25, (2.25 / 单元_βOw - 1.25 / 单元_βOw ^ 2) * 单元_hw, 单元_hw)
                        Dim 单元_df As Single = If(单元_tpS = "F", 单元_hw, If(单元_tpS = "B", 单元_hw - 单元_tf / 2, If(单元_tpS = "T" Or 单元_tpS = "L1" Or 单元_tpS = "L2", 单元_hw + 单元_tf / 2, 单元_hw - 单元_dx - 单元_tf / 2)))
                        Dim 单元_Af As Single = 单元_wf * 单元_tf
                        Dim 单元_Icf As Single = 单元_wf * 单元_tf ^ 3 / 12
                        Dim 单元_Iof As Single = 单元_Icf + 单元_Af * 单元_df ^ 2
                        Dim 单元_AS As Single = 单元_Aw + 单元_Af
                        Dim 单元_hcS As Single = 单元_Aw / 单元_AS * 单元_hw / 2 + 单元_Af / 单元_AS * 单元_df
                        Dim 单元_IOS As Single = 单元_Iow + 单元_Iof
                        Dim 单元_ICS As Single = 单元_Icw + 单元_Aw * (单元_hw / 2 - 单元_hcS) ^ 2 + 单元_Icf + 单元_Af * (单元_df - 单元_hcS) ^ 2
                        Dim 单元_IpS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw / 3, 单元_Aw * (单元_df - 单元_tf / 2) ^ 2 / 3 + 单元_Af * 单元_df ^ 2)
                        Dim 单元_ItS As Single = If(单元_tpS = "F", 单元_hw * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / 单元_hw), (单元_df - 单元_tf / 2) * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / (单元_df - 单元_tf / 2)) + 单元_wf * 单元_tf ^ 3 / 3 * (1 - 0.63 * 单元_tf / 单元_wf))
                        Dim 单元_IwS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw ^ 3 / 36, If(单元_tpS = "B" Or 单元_tpS = "L1" Or 单元_tpS = "L2" Or 单元_tpS = "L3", 单元_Af * 单元_df ^ 2 * 单元_wf ^ 2 / 12 * (单元_Af + 2.6 * 单元_Aw) / (单元_Af * 单元_Aw), 单元_wf ^ 3 * 单元_tf * 单元_df ^ 2 / 12))
                        Dim 单元_ηS As Single = 1 + (单元_lS / PI) ^ 2 / (单元_IwS * (0.75 * 单元_wP / 单元_tP ^ 3 + (单元_df - 单元_tf / 2) / 单元_tw ^ 3)) ^ (1 / 2)
                        Dim 单元_σETS As Single = 标准弹性模量 / 单元_IpS * (单元_ηS * PI ^ 2 * 单元_IwS / 单元_lS ^ 2 + 0.385 * 单元_ItS)
                        Dim 单元_σELS As Single = 160000 * (单元_tw / 单元_hw) ^ 2

                        'Dim 单元_A As Single = 单元_AS + 单元_AP
                        Dim 单元_hc As Single = 单元_AS / 单元_A * 单元_hcS + 单元_AP / 单元_A * (-单元_tP / 2)
                        Dim 单元_Io As Single = 单元_IOP + 单元_IOS
                        Dim 单元_Ic As Single = 单元_Io - 单元_A * 单元_hc ^ 2
                        'Dim 单元_σY As Single = 单元_AS / A * 单元_σYS + 单元_AP / 单元_A * 单元_σYP
                        'Dim 单元_εY As Single = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        Dim 单元_L As Single = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY

                        'Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        'Dim EP As Single = Φ * 单元_σY

                        'wP / lP
                        'Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)
                        'Dim βPε As Single = βOP * 单元_εR ^ (1 / 2)
                        'Dim BC As Single, FT As Single, WB As Single

                        '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING
                        Dim 单元_Φε As Single = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))
                        Dim 单元_σEPε As Single = 单元_Φε * 单元_σY

                        Dim 单元_Rlt_EP As String = Format(单元_σEPε, "000.000")

                        Select Case 单元_εR
                            Case < 0
                                Dim 单元_βPε As Single = 单元_βOP * (-单元_εR) ^ (1 / 2)
                                Dim 单元_wEPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_wP, 单元_wP)
                                Dim 单元_wE1Pε As Single = If(单元_βPε >= 1, 单元_wP / 单元_βPε, 单元_wP)
                                Dim 单元_AEPε As Single = 单元_wEPε * 单元_tP
                                Dim 单元_AE1Pε As Single = 单元_wE1Pε * 单元_tP
                                Dim 单元_IcE1Pε As Single = 单元_wE1Pε * 单元_tP ^ 3 / 12
                                Dim 单元_IoE1Pε As Single = 单元_IcE1Pε + 单元_AE1Pε * (-单元_tP / 2) ^ 2

                                Dim 单元_βwε As Single = 单元_βOw * (-单元_εR) ^ (1 / 2)
                                Dim 单元_hEwε As Single = If(单元_βwε >= 1.25, (2.25 / 单元_βwε - 1.25 / 单元_βwε ^ 2) * 单元_hw, 单元_hw)
                                Dim 单元_AEwε As Single = 单元_hEwε * 单元_tw
                                Dim 单元_AESε As Single = 单元_AEwε + 单元_Af

                                Dim 单元_AEε As Single = 单元_AEPε + 单元_AS
                                Dim 单元_AE1ε As Single = 单元_AE1Pε + 单元_AS
                                Dim 单元_hcE1ε As Single = 单元_AE1Pε / 单元_AE1ε * (-单元_tP / 2) + 单元_AS / 单元_AE1ε * 单元_hcS
                                Dim 单元_IoE1ε As Single = 单元_IoE1Pε + 单元_IOS
                                Dim 单元_IcE1ε As Single = 单元_IoE1ε - 单元_AE1ε * 单元_hcE1ε ^ 2
                                Dim 单元_lE1Pε As Single = 单元_hcE1ε - (-单元_tP / 2)
                                Dim 单元_lE1Sε As Single = If(单元_tpS = "F" Or 单元_tpS = "B" Or 单元_tpS = "L3", 单元_hw - 单元_hcE1ε, 单元_hw + 单元_tf - 单元_hcE1ε)
                                Dim 单元_σYE1ε As Single = 单元_AS * 单元_lE1Sε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYS + 单元_AE1Pε * 单元_lE1Pε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYP
                                Dim 单元_σECε As Single = PI ^ 2 * 标准弹性模量 * 单元_IcE1ε / 单元_AEε / 单元_lS ^ 2
                                Dim 单元_σCCε As Single = If(单元_σECε <= 单元_σYE1ε / 2 * (-单元_εR), 单元_σECε / (-单元_εR), 单元_σYE1ε * (1 - 单元_σYE1ε * (-单元_εR) / 4 / 单元_σECε))
                                Dim 单元_σBCε As Single = 单元_Φε * 单元_σCCε * 单元_AEε / 单元_A

                                Dim 单元_Rlt_BC As String = Format(单元_σBCε, "000.000")

                                Dim 单元_σCTSε As Single = If(单元_σETS <= 单元_σYS / 2 * (-单元_εR), 单元_σETS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σETS))
                                Dim 单元_σCPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_σYP, 单元_σYP)
                                Dim 单元_σFTε As Single = 单元_Φε * (单元_AS * 单元_σCTSε + 单元_AP * 单元_σCPε) / 单元_A

                                Dim 单元_Rlt_FT As String = Format(单元_σFTε, "000.000")

                                Dim 单元_σCLSε As Single = If(单元_σELS <= 单元_σYS / 2 * (-单元_εR), 单元_σELS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σELS))
                                Dim 单元_σWBε As Single = If(单元_tpS = "F", 单元_Φε * (单元_AS * 单元_σCLSε + 单元_AP * 单元_σCPε) / 单元_A, 单元_Φε * (单元_AESε * 单元_σYS + 单元_AEPε * 单元_σYP) / 单元_A)

                                Dim 单元_Rlt_WB As String = Format(单元_σWBε, "000.000")

                                ''''''
                                单元_σO = Max(单元_σBCε, Max(单元_σFTε, 单元_σWBε)) ' 单元_σO = 单元_σEPε   ''''''NO BUCKLING
                                            'If 四级序数 = 1 Then Debug.Print(单元_Rlt_EP & ", " & 单元_Rlt_BC & ", " & 单元_Rlt_FT & ", " & 单元_Rlt_WB)
                            Case > 0
                                单元_σO = 单元_σEPε
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 5
                    '
                Case 6
                    For 四级序数 As UShort = 1 To 加筋板单元_总数
                        Dim 单元_原始_Yc As Single = 加筋板单元_原始_Yc(四级序数) / 1000
                        Dim 单元_原始_Zc As Single = 加筋板单元_原始_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_原始_A As Single = 加筋板单元_原始_A(四级序数)
                        Dim 单元_原始_σY As Single = 加筋板单元_原始_σY(四级序数)

                        Dim 单元_剩余_Yc As Single = 加筋板单元_剩余_Yc(四级序数) / 1000
                        Dim 单元_剩余_Zc As Single = 加筋板单元_剩余_Zc(四级序数) / 1000
                        Dim 单元_剩余_A As Single = 加筋板单元_剩余_A(四级序数)
                        Dim 单元_剩余_σY As Single = 加筋板单元_剩余_σY(四级序数)

                        Dim 关联板格序数 As UShort = 加筋板单元_板格_序数(四级序数)

                        '[注意正负!]
                        单元_L = -单元_原始_Yc * Sin(α_瞬时) + (单元_原始_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_原始_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY
                        '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING

                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        Dim EP As Single = Φ * 单元_原始_σY

                        'wP / lP
                        Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)

                        Select Case 单元_εR
                            Case < 0
                                Dim βPε As Single = βOP * (-单元_εR) ^ (1 / 2)
                                Dim FB As Single = Φ * 单元_原始_σY * Min(1, 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数) * (2.25 / βPε - 1.25 / βPε ^ 2) + 0.1 * (1 - 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数)) * (1 + 1 / βPε ^ 2))
                                单元_σO = FB' 单元_σO = EP   ''''''NO BUCKLING
                            Case > 0
                                单元_σO = EP
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_剩余_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case Else
                    '
            End Select
        Next

        Dim 截面_YFP As Single, 截面_YFN As Single,
            截面_ZFP As Single, 截面_ZFN As Single
        Dim VP As Single
        截面_YFP = 截面_FYP / 截面_FP
        截面_YFN = 截面_FYN / 截面_FN
        截面_ZFP = 截面_FZP / 截面_FP
        截面_ZFN = 截面_FZN / 截面_FN
        VP = Cos(水线倾角_α) * (截面_YFP - 截面_YFN) / Sqrt((截面_ZFP - 截面_ZFN) ^ 2 + (截面_YFP - 截面_YFN) ^ 2) + Sin(水线倾角_α) * (截面_ZFP - 截面_ZFN) / Sqrt((截面_ZFP - 截面_ZFN) ^ 2 + (截面_YFP - 截面_YFN) ^ 2)
        判据参数2 = VP
    End Sub

    Private Sub 最终弯矩ζ(ByVal χ_瞬时 As Single, ByVal ζ_瞬时 As Single, ByVal α_瞬时 As Single, ByRef 结果 As Single)
        Dim 截面_FO As Single, 截面_FA As Single,
            截面_FP As Single, 截面_FN As Single,
            截面_FYP As Single, 截面_FYN As Single,
            截面_FZP As Single, 截面_FZN As Single,
            截面_MO As Single

        截面_FO = 0
        截面_FA = 0
        截面_FP = 0
        截面_FN = 0
        截面_FYP = 0
        截面_FYN = 0
        截面_FZP = 0
        截面_FZN = 0
        截面_MO = 0

        For 单元类型 As UShort = 1 To 6
            Select Case 单元类型
                Case 1
                    For 四级序数 As UShort = 1 To 硬角单元_总数
                        单元_Yc = 硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        单元_A = 硬角单元_A(四级序数) / 1000000
                        单元_σY = 硬角单元_σY(四级序数)
                        单元_L = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        'Debug.Print("#" & Format(单元类型, "0") & "," & Format(四级序数, "000") & "," & 单元_Yc & "," & 单元_Zc & "," & 单元_σO)
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 2
                    '
                Case 3
                    For 四级序数 As UShort = 1 To 特别硬角单元_总数
                        单元_Yc = 特别硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 特别硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 特别硬角单元_A(四级序数) / 1000000
                        单元_σY = 特别硬角单元_σY(四级序数)
                        单元_L = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        'Debug.Print("#" & Format(单元类型, "0") & "," & Format(四级序数, "000") & "," & 单元_Yc & "," & 单元_Zc & "," & 单元_σO)
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 4
                    For 四级序数 As UShort = 1 To 加强筋单元_总数
                        Dim 单元_Yc As Single = 加强筋单元_Yc(四级序数) / 1000
                        Dim 单元_Zc As Single = 加强筋单元_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_A As Single = 加强筋单元_A(四级序数)
                        Dim 单元_σY As Single = 加强筋单元_σY(四级序数)

                        Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(四级序数)
                        Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)

                        Dim 单元_σYP As Single = 加强筋单元_σYP(四级序数)
                        Dim 单元_lP As Single = 加强筋单元_lP(四级序数)
                        Dim 单元_wP As Single = 加强筋单元_wP(四级序数)
                        Dim 单元_tP As Single = 加强筋单元_tP(四级序数)
                        Dim 单元_σYS As Single = 加强筋单元_σYS(四级序数)
                        Dim 单元_lS As Single = 加强筋单元_lS(四级序数)
                        Dim 单元_hw As Single = 加强筋单元_hw(四级序数)
                        Dim 单元_tw As Single = 加强筋单元_tw(四级序数)
                        Dim 单元_wf As Single = 加强筋单元_wf(四级序数)
                        Dim 单元_tf As Single = 加强筋单元_tf(四级序数)
                        Dim 单元_dx As Single = 加强筋单元_dx(四级序数)
                        Dim 单元_tpS As String = 加强筋单元_tpS(四级序数)
                        Dim 单元_mk As Boolean = 加强筋单元_mk(四级序数)
                        'Dim 单元_σETS As Single = 加强筋单元_σETS(四级序数)
                        'Dim 单元_σELS As Single = 加强筋单元_σELS(四级序数)

                        Dim 单元_εYP As Single = 单元_σYP / 标准弹性模量
                        Dim 单元_AP As Single = 单元_wP * 单元_tP
                        Dim 单元_ICP As Single = 单元_wP * 单元_tP ^ 3 / 12
                        Dim 单元_IOP As Single = 单元_ICP + 单元_AP * (-单元_tP / 2) ^ 2
                        Dim 单元_βOP As Single = 单元_wP / 单元_tP * 单元_εYP ^ (1 / 2)
                        Dim 单元_wEoP As Single = If(单元_βOP >= 1.25, (2.25 / 单元_βOP - 1.25 / 单元_βOP ^ 2) * 单元_wP, 单元_wP)

                        Dim 单元_εYS As Single = 单元_σYS / 标准弹性模量
                        Dim 单元_Aw As Single = 单元_hw * 单元_tw
                        Dim 单元_Icw As Single = 单元_tw * 单元_hw ^ 3 / 12
                        Dim 单元_Iow As Single = 单元_Icw + 单元_Aw * (单元_hw / 2) ^ 2
                        Dim 单元_βOw As Single = 单元_hw / 单元_tw * 单元_εYS ^ (1 / 2)
                        Dim 单元_hEow As Single = If(单元_βOw >= 1.25, (2.25 / 单元_βOw - 1.25 / 单元_βOw ^ 2) * 单元_hw, 单元_hw)
                        Dim 单元_df As Single = If(单元_tpS = "F", 单元_hw, If(单元_tpS = "B", 单元_hw - 单元_tf / 2, If(单元_tpS = "T" Or 单元_tpS = "L1" Or 单元_tpS = "L2", 单元_hw + 单元_tf / 2, 单元_hw - 单元_dx - 单元_tf / 2)))
                        Dim 单元_Af As Single = 单元_wf * 单元_tf
                        Dim 单元_Icf As Single = 单元_wf * 单元_tf ^ 3 / 12
                        Dim 单元_Iof As Single = 单元_Icf + 单元_Af * 单元_df ^ 2
                        Dim 单元_AS As Single = 单元_Aw + 单元_Af
                        Dim 单元_hcS As Single = 单元_Aw / 单元_AS * 单元_hw / 2 + 单元_Af / 单元_AS * 单元_df
                        Dim 单元_IOS As Single = 单元_Iow + 单元_Iof
                        Dim 单元_ICS As Single = 单元_Icw + 单元_Aw * (单元_hw / 2 - 单元_hcS) ^ 2 + 单元_Icf + 单元_Af * (单元_df - 单元_hcS) ^ 2
                        Dim 单元_IpS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw / 3, 单元_Aw * (单元_df - 单元_tf / 2) ^ 2 / 3 + 单元_Af * 单元_df ^ 2)
                        Dim 单元_ItS As Single = If(单元_tpS = "F", 单元_hw * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / 单元_hw), (单元_df - 单元_tf / 2) * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / (单元_df - 单元_tf / 2)) + 单元_wf * 单元_tf ^ 3 / 3 * (1 - 0.63 * 单元_tf / 单元_wf))
                        Dim 单元_IwS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw ^ 3 / 36, If(单元_tpS = "B" Or 单元_tpS = "L1" Or 单元_tpS = "L2" Or 单元_tpS = "L3", 单元_Af * 单元_df ^ 2 * 单元_wf ^ 2 / 12 * (单元_Af + 2.6 * 单元_Aw) / (单元_Af * 单元_Aw), 单元_wf ^ 3 * 单元_tf * 单元_df ^ 2 / 12))
                        Dim 单元_ηS As Single = 1 + (单元_lS / PI) ^ 2 / (单元_IwS * (0.75 * 单元_wP / 单元_tP ^ 3 + (单元_df - 单元_tf / 2) / 单元_tw ^ 3)) ^ (1 / 2)
                        Dim 单元_σETS As Single = 标准弹性模量 / 单元_IpS * (单元_ηS * PI ^ 2 * 单元_IwS / 单元_lS ^ 2 + 0.385 * 单元_ItS)
                        Dim 单元_σELS As Single = 160000 * (单元_tw / 单元_hw) ^ 2

                        'Dim 单元_A As Single = 单元_AS + 单元_AP
                        Dim 单元_hc As Single = 单元_AS / 单元_A * 单元_hcS + 单元_AP / 单元_A * (-单元_tP / 2)
                        Dim 单元_Io As Single = 单元_IOP + 单元_IOS
                        Dim 单元_Ic As Single = 单元_Io - 单元_A * 单元_hc ^ 2
                        'Dim 单元_σY As Single = 单元_AS / A * 单元_σYS + 单元_AP / 单元_A * 单元_σYP
                        'Dim 单元_εY As Single = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        Dim 单元_L As Single = -单元_Yc * Sin(α_瞬时) + (单元_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY
                        '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING

                        'Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        'Dim EP As Single = Φ * 单元_σY

                        'wP / lP
                        'Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)
                        'Dim βPε As Single = βOP * 单元_εR ^ (1 / 2)
                        'Dim BC As Single, FT As Single, WB As Single

                        '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING
                        Dim 单元_Φε As Single = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))
                        Dim 单元_σEPε As Single = 单元_Φε * 单元_σY

                        Dim 单元_Rlt_EP As String = Format(单元_σEPε, "000.000")

                        Select Case 单元_εR
                            Case < 0
                                Dim 单元_βPε As Single = 单元_βOP * (-单元_εR) ^ (1 / 2)
                                Dim 单元_wEPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_wP, 单元_wP)
                                Dim 单元_wE1Pε As Single = If(单元_βPε >= 1, 单元_wP / 单元_βPε, 单元_wP)
                                Dim 单元_AEPε As Single = 单元_wEPε * 单元_tP
                                Dim 单元_AE1Pε As Single = 单元_wE1Pε * 单元_tP
                                Dim 单元_IcE1Pε As Single = 单元_wE1Pε * 单元_tP ^ 3 / 12
                                Dim 单元_IoE1Pε As Single = 单元_IcE1Pε + 单元_AE1Pε * (-单元_tP / 2) ^ 2

                                Dim 单元_βwε As Single = 单元_βOw * (-单元_εR) ^ (1 / 2)
                                Dim 单元_hEwε As Single = If(单元_βwε >= 1.25, (2.25 / 单元_βwε - 1.25 / 单元_βwε ^ 2) * 单元_hw, 单元_hw)
                                Dim 单元_AEwε As Single = 单元_hEwε * 单元_tw
                                Dim 单元_AESε As Single = 单元_AEwε + 单元_Af

                                Dim 单元_AEε As Single = 单元_AEPε + 单元_AS
                                Dim 单元_AE1ε As Single = 单元_AE1Pε + 单元_AS
                                Dim 单元_hcE1ε As Single = 单元_AE1Pε / 单元_AE1ε * (-单元_tP / 2) + 单元_AS / 单元_AE1ε * 单元_hcS
                                Dim 单元_IoE1ε As Single = 单元_IoE1Pε + 单元_IOS
                                Dim 单元_IcE1ε As Single = 单元_IoE1ε - 单元_AE1ε * 单元_hcE1ε ^ 2
                                Dim 单元_lE1Pε As Single = 单元_hcE1ε - (-单元_tP / 2)
                                Dim 单元_lE1Sε As Single = If(单元_tpS = "F" Or 单元_tpS = "B" Or 单元_tpS = "L3", 单元_hw - 单元_hcE1ε, 单元_hw + 单元_tf - 单元_hcE1ε)
                                Dim 单元_σYE1ε As Single = 单元_AS * 单元_lE1Sε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYS + 单元_AE1Pε * 单元_lE1Pε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYP
                                Dim 单元_σECε As Single = PI ^ 2 * 标准弹性模量 * 单元_IcE1ε / 单元_AEε / 单元_lS ^ 2
                                Dim 单元_σCCε As Single = If(单元_σECε <= 单元_σYE1ε / 2 * (-单元_εR), 单元_σECε / (-单元_εR), 单元_σYE1ε * (1 - 单元_σYE1ε * (-单元_εR) / 4 / 单元_σECε))
                                Dim 单元_σBCε As Single = 单元_Φε * 单元_σCCε * 单元_AEε / 单元_A

                                Dim 单元_Rlt_BC As String = Format(单元_σBCε, "000.000")

                                Dim 单元_σCTSε As Single = If(单元_σETS <= 单元_σYS / 2 * (-单元_εR), 单元_σETS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σETS))
                                Dim 单元_σCPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_σYP, 单元_σYP)
                                Dim 单元_σFTε As Single = 单元_Φε * (单元_AS * 单元_σCTSε + 单元_AP * 单元_σCPε) / 单元_A

                                Dim 单元_Rlt_FT As String = Format(单元_σFTε, "000.000")

                                Dim 单元_σCLSε As Single = If(单元_σELS <= 单元_σYS / 2 * (-单元_εR), 单元_σELS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σELS))
                                Dim 单元_σWBε As Single = If(单元_tpS = "F", 单元_Φε * (单元_AS * 单元_σCLSε + 单元_AP * 单元_σCPε) / 单元_A, 单元_Φε * (单元_AESε * 单元_σYS + 单元_AEPε * 单元_σYP) / 单元_A)

                                Dim 单元_Rlt_WB As String = Format(单元_σWBε, "000.000")

                                ''''''
                                单元_σO = Max(单元_σBCε, Max(单元_σFTε, 单元_σWBε)) ' 单元_σO = 单元_σEPε   ''''''NO BUCKLING
                                            'If 四级序数 = 1 Then Debug.Print(单元_Rlt_EP & ", " & 单元_Rlt_BC & ", " & 单元_Rlt_FT & ", " & 单元_Rlt_WB)
                            Case > 0
                                单元_σO = 单元_σEPε
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select
                        'Debug.Print("#" & Format(单元类型, "0") & "," & Format(四级序数, "000") & "," & 单元_Yc & "," & 单元_Zc & "," & 单元_σO)

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 5
                    '
                Case 6
                    For 四级序数 As UShort = 1 To 加筋板单元_总数
                        Dim 单元_原始_Yc As Single = 加筋板单元_原始_Yc(四级序数) / 1000
                        Dim 单元_原始_Zc As Single = 加筋板单元_原始_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_原始_A As Single = 加筋板单元_原始_A(四级序数)
                        Dim 单元_原始_σY As Single = 加筋板单元_原始_σY(四级序数)

                        Dim 单元_剩余_Yc As Single = 加筋板单元_剩余_Yc(四级序数) / 1000
                        Dim 单元_剩余_Zc As Single = 加筋板单元_剩余_Zc(四级序数) / 1000
                        Dim 单元_剩余_A As Single = 加筋板单元_剩余_A(四级序数)
                        Dim 单元_剩余_σY As Single = 加筋板单元_剩余_σY(四级序数)

                        Dim 关联板格序数 As UShort = 加筋板单元_板格_序数(四级序数)

                        '[注意正负!]
                        单元_L = -单元_原始_Yc * Sin(α_瞬时) + (单元_原始_Zc - ζ_瞬时) * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_原始_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY

                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        Dim EP As Single = Φ * 单元_原始_σY

                        'wP / lP
                        Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)

                        Select Case 单元_εR
                            Case < 0
                                Dim βPε As Single = βOP * (-单元_εR) ^ (1 / 2)
                                Dim FB As Single = Φ * 单元_原始_σY * Min(1, 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数) * (2.25 / βPε - 1.25 / βPε ^ 2) + 0.1 * (1 - 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数)) * (1 + 1 / βPε ^ 2))
                                单元_σO = FB' 单元_σO = EP   ''''''NO BUCKLING
                            Case > 0
                                单元_σO = EP
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select
                        'Debug.Print("#" & Format(单元类型, "0") & "," & Format(四级序数, "000") & "," & 单元_剩余_Yc & "," & 单元_剩余_Zc & "," & 单元_σO)

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_剩余_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case Else
                    '
            End Select
        Next

        'Dim 截面_YFP As Single, 截面_YFN As Single,
        '    截面_ZFP As Single, 截面_ZFN As Single
        'Dim α_RF As Single
        'Dim VP As Single
        '截面_YFP = 截面_FYP / 截面_FP
        '截面_YFN = 截面_FYN / 截面_FN
        '截面_ZFP = 截面_FZP / 截面_FP
        '截面_ZFN = 截面_FZN / 截面_FN
        'α_RF = Atan((截面_ZFP - 截面_ZFN) / (截面_YFP - 截面_YFN))
        'VP = Cos(水线倾角_α) * Cos(α_RF) + Sin(水线倾角_α) * Sin(α_RF)
        结果 = 截面_MO
    End Sub

    Private Sub γ判据(ByVal χ_瞬时 As Single, ByVal γ_瞬时 As Single, ByVal α_瞬时 As Single, ByRef 判据参数3 As Single)
        Dim 截面_FO As Single, 截面_FA As Single,
            截面_FP As Single, 截面_FN As Single,
            截面_FYP As Single, 截面_FYN As Single,
            截面_FZP As Single, 截面_FZN As Single,
            截面_MO As Single

        截面_FO = 0
        截面_FA = 0
        截面_FP = 0
        截面_FN = 0
        截面_FYP = 0
        截面_FYN = 0
        截面_FZP = 0
        截面_FZN = 0
        截面_MO = 0

        For 单元类型 As UShort = 1 To 6
            Select Case 单元类型
                Case 1
                    For 四级序数 As UShort = 1 To 硬角单元_总数
                        单元_Yc = 硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 硬角单元_A(四级序数) / 1000000
                        单元_σY = 硬角单元_σY(四级序数)
                        单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 2
                    '
                Case 3
                    For 四级序数 As UShort = 1 To 特别硬角单元_总数
                        单元_Yc = 特别硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 特别硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 特别硬角单元_A(四级序数) / 1000000
                        单元_σY = 特别硬角单元_σY(四级序数)
                        单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 4
                    For 四级序数 As UShort = 1 To 加强筋单元_总数
                        Dim 单元_Yc As Single = 加强筋单元_Yc(四级序数) / 1000
                        Dim 单元_Zc As Single = 加强筋单元_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_A As Single = 加强筋单元_A(四级序数)
                        Dim 单元_σY As Single = 加强筋单元_σY(四级序数)

                        Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(四级序数)
                        Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)

                        Dim 单元_σYP As Single = 加强筋单元_σYP(四级序数)
                        Dim 单元_lP As Single = 加强筋单元_lP(四级序数)
                        Dim 单元_wP As Single = 加强筋单元_wP(四级序数)
                        Dim 单元_tP As Single = 加强筋单元_tP(四级序数)
                        Dim 单元_σYS As Single = 加强筋单元_σYS(四级序数)
                        Dim 单元_lS As Single = 加强筋单元_lS(四级序数)
                        Dim 单元_hw As Single = 加强筋单元_hw(四级序数)
                        Dim 单元_tw As Single = 加强筋单元_tw(四级序数)
                        Dim 单元_wf As Single = 加强筋单元_wf(四级序数)
                        Dim 单元_tf As Single = 加强筋单元_tf(四级序数)
                        Dim 单元_dx As Single = 加强筋单元_dx(四级序数)
                        Dim 单元_tpS As String = 加强筋单元_tpS(四级序数)
                        Dim 单元_mk As Boolean = 加强筋单元_mk(四级序数)
                        'Dim 单元_σETS As Single = 加强筋单元_σETS(四级序数)
                        'Dim 单元_σELS As Single = 加强筋单元_σELS(四级序数)

                        Dim 单元_εYP As Single = 单元_σYP / 标准弹性模量
                        Dim 单元_AP As Single = 单元_wP * 单元_tP
                        Dim 单元_ICP As Single = 单元_wP * 单元_tP ^ 3 / 12
                        Dim 单元_IOP As Single = 单元_ICP + 单元_AP * (-单元_tP / 2) ^ 2
                        Dim 单元_βOP As Single = 单元_wP / 单元_tP * 单元_εYP ^ (1 / 2)
                        Dim 单元_wEoP As Single = If(单元_βOP >= 1.25, (2.25 / 单元_βOP - 1.25 / 单元_βOP ^ 2) * 单元_wP, 单元_wP)

                        Dim 单元_εYS As Single = 单元_σYS / 标准弹性模量
                        Dim 单元_Aw As Single = 单元_hw * 单元_tw
                        Dim 单元_Icw As Single = 单元_tw * 单元_hw ^ 3 / 12
                        Dim 单元_Iow As Single = 单元_Icw + 单元_Aw * (单元_hw / 2) ^ 2
                        Dim 单元_βOw As Single = 单元_hw / 单元_tw * 单元_εYS ^ (1 / 2)
                        Dim 单元_hEow As Single = If(单元_βOw >= 1.25, (2.25 / 单元_βOw - 1.25 / 单元_βOw ^ 2) * 单元_hw, 单元_hw)
                        Dim 单元_df As Single = If(单元_tpS = "F", 单元_hw, If(单元_tpS = "B", 单元_hw - 单元_tf / 2, If(单元_tpS = "T" Or 单元_tpS = "L1" Or 单元_tpS = "L2", 单元_hw + 单元_tf / 2, 单元_hw - 单元_dx - 单元_tf / 2)))
                        Dim 单元_Af As Single = 单元_wf * 单元_tf
                        Dim 单元_Icf As Single = 单元_wf * 单元_tf ^ 3 / 12
                        Dim 单元_Iof As Single = 单元_Icf + 单元_Af * 单元_df ^ 2
                        Dim 单元_AS As Single = 单元_Aw + 单元_Af
                        Dim 单元_hcS As Single = 单元_Aw / 单元_AS * 单元_hw / 2 + 单元_Af / 单元_AS * 单元_df
                        Dim 单元_IOS As Single = 单元_Iow + 单元_Iof
                        Dim 单元_ICS As Single = 单元_Icw + 单元_Aw * (单元_hw / 2 - 单元_hcS) ^ 2 + 单元_Icf + 单元_Af * (单元_df - 单元_hcS) ^ 2
                        Dim 单元_IpS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw / 3, 单元_Aw * (单元_df - 单元_tf / 2) ^ 2 / 3 + 单元_Af * 单元_df ^ 2)
                        Dim 单元_ItS As Single = If(单元_tpS = "F", 单元_hw * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / 单元_hw), (单元_df - 单元_tf / 2) * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / (单元_df - 单元_tf / 2)) + 单元_wf * 单元_tf ^ 3 / 3 * (1 - 0.63 * 单元_tf / 单元_wf))
                        Dim 单元_IwS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw ^ 3 / 36, If(单元_tpS = "B" Or 单元_tpS = "L1" Or 单元_tpS = "L2" Or 单元_tpS = "L3", 单元_Af * 单元_df ^ 2 * 单元_wf ^ 2 / 12 * (单元_Af + 2.6 * 单元_Aw) / (单元_Af * 单元_Aw), 单元_wf ^ 3 * 单元_tf * 单元_df ^ 2 / 12))
                        Dim 单元_ηS As Single = 1 + (单元_lS / PI) ^ 2 / (单元_IwS * (0.75 * 单元_wP / 单元_tP ^ 3 + (单元_df - 单元_tf / 2) / 单元_tw ^ 3)) ^ (1 / 2)
                        Dim 单元_σETS As Single = 标准弹性模量 / 单元_IpS * (单元_ηS * PI ^ 2 * 单元_IwS / 单元_lS ^ 2 + 0.385 * 单元_ItS)
                        Dim 单元_σELS As Single = 160000 * (单元_tw / 单元_hw) ^ 2

                        'Dim 单元_A As Single = 单元_AS + 单元_AP
                        Dim 单元_hc As Single = 单元_AS / 单元_A * 单元_hcS + 单元_AP / 单元_A * (-单元_tP / 2)
                        Dim 单元_Io As Single = 单元_IOP + 单元_IOS
                        Dim 单元_Ic As Single = 单元_Io - 单元_A * 单元_hc ^ 2
                        'Dim 单元_σY As Single = 单元_AS / A * 单元_σYS + 单元_AP / 单元_A * 单元_σYP
                        'Dim 单元_εY As Single = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        Dim 单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY

                        'Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        'Dim EP As Single = Φ * 单元_σY

                        'wP / lP
                        'Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)
                        'Dim βPε As Single = βOP * 单元_εR ^ (1 / 2)
                        'Dim BC As Single, FT As Single, WB As Single

                        '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING
                        Dim 单元_Φε As Single = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))
                        Dim 单元_σEPε As Single = 单元_Φε * 单元_σY

                        Dim 单元_Rlt_EP As String = Format(单元_σEPε, "000.000")

                        Select Case 单元_εR
                            Case < 0
                                Dim 单元_βPε As Single = 单元_βOP * (-单元_εR) ^ (1 / 2)
                                Dim 单元_wEPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_wP, 单元_wP)
                                Dim 单元_wE1Pε As Single = If(单元_βPε >= 1, 单元_wP / 单元_βPε, 单元_wP)
                                Dim 单元_AEPε As Single = 单元_wEPε * 单元_tP
                                Dim 单元_AE1Pε As Single = 单元_wE1Pε * 单元_tP
                                Dim 单元_IcE1Pε As Single = 单元_wE1Pε * 单元_tP ^ 3 / 12
                                Dim 单元_IoE1Pε As Single = 单元_IcE1Pε + 单元_AE1Pε * (-单元_tP / 2) ^ 2

                                Dim 单元_βwε As Single = 单元_βOw * (-单元_εR) ^ (1 / 2)
                                Dim 单元_hEwε As Single = If(单元_βwε >= 1.25, (2.25 / 单元_βwε - 1.25 / 单元_βwε ^ 2) * 单元_hw, 单元_hw)
                                Dim 单元_AEwε As Single = 单元_hEwε * 单元_tw
                                Dim 单元_AESε As Single = 单元_AEwε + 单元_Af

                                Dim 单元_AEε As Single = 单元_AEPε + 单元_AS
                                Dim 单元_AE1ε As Single = 单元_AE1Pε + 单元_AS
                                Dim 单元_hcE1ε As Single = 单元_AE1Pε / 单元_AE1ε * (-单元_tP / 2) + 单元_AS / 单元_AE1ε * 单元_hcS
                                Dim 单元_IoE1ε As Single = 单元_IoE1Pε + 单元_IOS
                                Dim 单元_IcE1ε As Single = 单元_IoE1ε - 单元_AE1ε * 单元_hcE1ε ^ 2
                                Dim 单元_lE1Pε As Single = 单元_hcE1ε - (-单元_tP / 2)
                                Dim 单元_lE1Sε As Single = If(单元_tpS = "F" Or 单元_tpS = "B" Or 单元_tpS = "L3", 单元_hw - 单元_hcE1ε, 单元_hw + 单元_tf - 单元_hcE1ε)
                                Dim 单元_σYE1ε As Single = 单元_AS * 单元_lE1Sε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYS + 单元_AE1Pε * 单元_lE1Pε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYP
                                Dim 单元_σECε As Single = PI ^ 2 * 标准弹性模量 * 单元_IcE1ε / 单元_AEε / 单元_lS ^ 2
                                Dim 单元_σCCε As Single = If(单元_σECε <= 单元_σYE1ε / 2 * (-单元_εR), 单元_σECε / (-单元_εR), 单元_σYE1ε * (1 - 单元_σYE1ε * (-单元_εR) / 4 / 单元_σECε))
                                Dim 单元_σBCε As Single = 单元_Φε * 单元_σCCε * 单元_AEε / 单元_A

                                Dim 单元_Rlt_BC As String = Format(单元_σBCε, "000.000")

                                Dim 单元_σCTSε As Single = If(单元_σETS <= 单元_σYS / 2 * (-单元_εR), 单元_σETS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σETS))
                                Dim 单元_σCPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_σYP, 单元_σYP)
                                Dim 单元_σFTε As Single = 单元_Φε * (单元_AS * 单元_σCTSε + 单元_AP * 单元_σCPε) / 单元_A

                                Dim 单元_Rlt_FT As String = Format(单元_σFTε, "000.000")

                                Dim 单元_σCLSε As Single = If(单元_σELS <= 单元_σYS / 2 * (-单元_εR), 单元_σELS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σELS))
                                Dim 单元_σWBε As Single = If(单元_tpS = "F", 单元_Φε * (单元_AS * 单元_σCLSε + 单元_AP * 单元_σCPε) / 单元_A, 单元_Φε * (单元_AESε * 单元_σYS + 单元_AEPε * 单元_σYP) / 单元_A)

                                Dim 单元_Rlt_WB As String = Format(单元_σWBε, "000.000")

                                ''''''
                                单元_σO = Max(单元_σBCε, Max(单元_σFTε, 单元_σWBε)) ' 单元_σO = 单元_σEPε   ''''''NO BUCKLING
                                            'If 四级序数 = 1 Then Debug.Print(单元_Rlt_EP & ", " & 单元_Rlt_BC & ", " & 单元_Rlt_FT & ", " & 单元_Rlt_WB)
                            Case > 0
                                单元_σO = 单元_σEPε
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 5
                    '
                Case 6
                    For 四级序数 As UShort = 1 To 加筋板单元_总数
                        Dim 单元_原始_Yc As Single = 加筋板单元_原始_Yc(四级序数) / 1000
                        Dim 单元_原始_Zc As Single = 加筋板单元_原始_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_原始_A As Single = 加筋板单元_原始_A(四级序数)
                        Dim 单元_原始_σY As Single = 加筋板单元_原始_σY(四级序数)

                        Dim 单元_剩余_Yc As Single = 加筋板单元_剩余_Yc(四级序数) / 1000
                        Dim 单元_剩余_Zc As Single = 加筋板单元_剩余_Zc(四级序数) / 1000
                        Dim 单元_剩余_A As Single = 加筋板单元_剩余_A(四级序数)
                        Dim 单元_剩余_σY As Single = 加筋板单元_剩余_σY(四级序数)

                        Dim 关联板格序数 As UShort = 加筋板单元_板格_序数(四级序数)

                        '[注意正负!]
                        单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_原始_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY

                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        Dim EP As Single = Φ * 单元_原始_σY

                        'wP / lP
                        Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)

                        Select Case 单元_εR
                            Case < 0
                                Dim βPε As Single = βOP * (-单元_εR) ^ (1 / 2)
                                Dim FB As Single = Φ * 单元_原始_σY * Min(1, 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数) * (2.25 / βPε - 1.25 / βPε ^ 2) + 0.1 * (1 - 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数)) * (1 + 1 / βPε ^ 2))
                                单元_σO = FB
                            Case > 0
                                单元_σO = EP
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_剩余_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case Else
                    '
            End Select
        Next

        'Dim 截面_YFP As Single, 截面_YFN As Single,
        '    截面_ZFP As Single, 截面_ZFN As Single
        'Dim α_RF As Single
        'Dim VP As Single
        '截面_YFP = 截面_FYP / 截面_FP
        '截面_YFN = 截面_FYN / 截面_FN
        '截面_ZFP = 截面_FZP / 截面_FP
        '截面_ZFN = 截面_FZN / 截面_FN
        'α_RF = Atan((截面_ZFP - 截面_ZFN) / (截面_YFP - 截面_YFN))
        'VP = Cos(水线倾角_α) * Cos(α_RF) + Sin(水线倾角_α) * Sin(α_RF)
        判据参数3 = 截面_FO
    End Sub

    Private Sub αγ判据(ByVal χ_瞬时 As Single, ByVal γ_瞬时 As Single, ByVal α_瞬时 As Single, ByRef 判据参数4 As Single)
        Dim 截面_FO As Single, 截面_FA As Single,
            截面_FP As Single, 截面_FN As Single,
            截面_FYP As Single, 截面_FYN As Single,
            截面_FZP As Single, 截面_FZN As Single,
            截面_MO As Single

        截面_FO = 0
        截面_FA = 0
        截面_FP = 0
        截面_FN = 0
        截面_FYP = 0
        截面_FYN = 0
        截面_FZP = 0
        截面_FZN = 0
        截面_MO = 0

        For 单元类型 As UShort = 1 To 6
            Select Case 单元类型
                Case 1
                    For 四级序数 As UShort = 1 To 硬角单元_总数
                        单元_Yc = 硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 硬角单元_A(四级序数) / 1000000
                        单元_σY = 硬角单元_σY(四级序数)
                        单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 2
                    '
                Case 3
                    For 四级序数 As UShort = 1 To 特别硬角单元_总数
                        单元_Yc = 特别硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 特别硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 特别硬角单元_A(四级序数) / 1000000
                        单元_σY = 特别硬角单元_σY(四级序数)
                        单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 4
                    For 四级序数 As UShort = 1 To 加强筋单元_总数
                        Dim 单元_Yc As Single = 加强筋单元_Yc(四级序数) / 1000
                        Dim 单元_Zc As Single = 加强筋单元_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_A As Single = 加强筋单元_A(四级序数)
                        Dim 单元_σY As Single = 加强筋单元_σY(四级序数)

                        Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(四级序数)
                        Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)

                        Dim 单元_σYP As Single = 加强筋单元_σYP(四级序数)
                        Dim 单元_lP As Single = 加强筋单元_lP(四级序数)
                        Dim 单元_wP As Single = 加强筋单元_wP(四级序数)
                        Dim 单元_tP As Single = 加强筋单元_tP(四级序数)
                        Dim 单元_σYS As Single = 加强筋单元_σYS(四级序数)
                        Dim 单元_lS As Single = 加强筋单元_lS(四级序数)
                        Dim 单元_hw As Single = 加强筋单元_hw(四级序数)
                        Dim 单元_tw As Single = 加强筋单元_tw(四级序数)
                        Dim 单元_wf As Single = 加强筋单元_wf(四级序数)
                        Dim 单元_tf As Single = 加强筋单元_tf(四级序数)
                        Dim 单元_dx As Single = 加强筋单元_dx(四级序数)
                        Dim 单元_tpS As String = 加强筋单元_tpS(四级序数)
                        Dim 单元_mk As Boolean = 加强筋单元_mk(四级序数)
                        'Dim 单元_σETS As Single = 加强筋单元_σETS(四级序数)
                        'Dim 单元_σELS As Single = 加强筋单元_σELS(四级序数)

                        Dim 单元_εYP As Single = 单元_σYP / 标准弹性模量
                        Dim 单元_AP As Single = 单元_wP * 单元_tP
                        Dim 单元_ICP As Single = 单元_wP * 单元_tP ^ 3 / 12
                        Dim 单元_IOP As Single = 单元_ICP + 单元_AP * (-单元_tP / 2) ^ 2
                        Dim 单元_βOP As Single = 单元_wP / 单元_tP * 单元_εYP ^ (1 / 2)
                        Dim 单元_wEoP As Single = If(单元_βOP >= 1.25, (2.25 / 单元_βOP - 1.25 / 单元_βOP ^ 2) * 单元_wP, 单元_wP)

                        Dim 单元_εYS As Single = 单元_σYS / 标准弹性模量
                        Dim 单元_Aw As Single = 单元_hw * 单元_tw
                        Dim 单元_Icw As Single = 单元_tw * 单元_hw ^ 3 / 12
                        Dim 单元_Iow As Single = 单元_Icw + 单元_Aw * (单元_hw / 2) ^ 2
                        Dim 单元_βOw As Single = 单元_hw / 单元_tw * 单元_εYS ^ (1 / 2)
                        Dim 单元_hEow As Single = If(单元_βOw >= 1.25, (2.25 / 单元_βOw - 1.25 / 单元_βOw ^ 2) * 单元_hw, 单元_hw)
                        Dim 单元_df As Single = If(单元_tpS = "F", 单元_hw, If(单元_tpS = "B", 单元_hw - 单元_tf / 2, If(单元_tpS = "T" Or 单元_tpS = "L1" Or 单元_tpS = "L2", 单元_hw + 单元_tf / 2, 单元_hw - 单元_dx - 单元_tf / 2)))
                        Dim 单元_Af As Single = 单元_wf * 单元_tf
                        Dim 单元_Icf As Single = 单元_wf * 单元_tf ^ 3 / 12
                        Dim 单元_Iof As Single = 单元_Icf + 单元_Af * 单元_df ^ 2
                        Dim 单元_AS As Single = 单元_Aw + 单元_Af
                        Dim 单元_hcS As Single = 单元_Aw / 单元_AS * 单元_hw / 2 + 单元_Af / 单元_AS * 单元_df
                        Dim 单元_IOS As Single = 单元_Iow + 单元_Iof
                        Dim 单元_ICS As Single = 单元_Icw + 单元_Aw * (单元_hw / 2 - 单元_hcS) ^ 2 + 单元_Icf + 单元_Af * (单元_df - 单元_hcS) ^ 2
                        Dim 单元_IpS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw / 3, 单元_Aw * (单元_df - 单元_tf / 2) ^ 2 / 3 + 单元_Af * 单元_df ^ 2)
                        Dim 单元_ItS As Single = If(单元_tpS = "F", 单元_hw * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / 单元_hw), (单元_df - 单元_tf / 2) * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / (单元_df - 单元_tf / 2)) + 单元_wf * 单元_tf ^ 3 / 3 * (1 - 0.63 * 单元_tf / 单元_wf))
                        Dim 单元_IwS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw ^ 3 / 36, If(单元_tpS = "B" Or 单元_tpS = "L1" Or 单元_tpS = "L2" Or 单元_tpS = "L3", 单元_Af * 单元_df ^ 2 * 单元_wf ^ 2 / 12 * (单元_Af + 2.6 * 单元_Aw) / (单元_Af * 单元_Aw), 单元_wf ^ 3 * 单元_tf * 单元_df ^ 2 / 12))
                        Dim 单元_ηS As Single = 1 + (单元_lS / PI) ^ 2 / (单元_IwS * (0.75 * 单元_wP / 单元_tP ^ 3 + (单元_df - 单元_tf / 2) / 单元_tw ^ 3)) ^ (1 / 2)
                        Dim 单元_σETS As Single = 标准弹性模量 / 单元_IpS * (单元_ηS * PI ^ 2 * 单元_IwS / 单元_lS ^ 2 + 0.385 * 单元_ItS)
                        Dim 单元_σELS As Single = 160000 * (单元_tw / 单元_hw) ^ 2

                        'Dim 单元_A As Single = 单元_AS + 单元_AP
                        Dim 单元_hc As Single = 单元_AS / 单元_A * 单元_hcS + 单元_AP / 单元_A * (-单元_tP / 2)
                        Dim 单元_Io As Single = 单元_IOP + 单元_IOS
                        Dim 单元_Ic As Single = 单元_Io - 单元_A * 单元_hc ^ 2
                        'Dim 单元_σY As Single = 单元_AS / A * 单元_σYS + 单元_AP / 单元_A * 单元_σYP
                        'Dim 单元_εY As Single = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        Dim 单元_L As Single = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY

                        'Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        'Dim EP As Single = Φ * 单元_σY

                        'wP / lP
                        'Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)
                        'Dim βPε As Single = βOP * 单元_εR ^ (1 / 2)
                        'Dim BC As Single, FT As Single, WB As Single

                        '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING
                        Dim 单元_Φε As Single = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))
                        Dim 单元_σEPε As Single = 单元_Φε * 单元_σY

                        Dim 单元_Rlt_EP As String = Format(单元_σEPε, "000.000")

                        Select Case 单元_εR
                            Case < 0
                                Dim 单元_βPε As Single = 单元_βOP * (-单元_εR) ^ (1 / 2)
                                Dim 单元_wEPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_wP, 单元_wP)
                                Dim 单元_wE1Pε As Single = If(单元_βPε >= 1, 单元_wP / 单元_βPε, 单元_wP)
                                Dim 单元_AEPε As Single = 单元_wEPε * 单元_tP
                                Dim 单元_AE1Pε As Single = 单元_wE1Pε * 单元_tP
                                Dim 单元_IcE1Pε As Single = 单元_wE1Pε * 单元_tP ^ 3 / 12
                                Dim 单元_IoE1Pε As Single = 单元_IcE1Pε + 单元_AE1Pε * (-单元_tP / 2) ^ 2

                                Dim 单元_βwε As Single = 单元_βOw * (-单元_εR) ^ (1 / 2)
                                Dim 单元_hEwε As Single = If(单元_βwε >= 1.25, (2.25 / 单元_βwε - 1.25 / 单元_βwε ^ 2) * 单元_hw, 单元_hw)
                                Dim 单元_AEwε As Single = 单元_hEwε * 单元_tw
                                Dim 单元_AESε As Single = 单元_AEwε + 单元_Af

                                Dim 单元_AEε As Single = 单元_AEPε + 单元_AS
                                Dim 单元_AE1ε As Single = 单元_AE1Pε + 单元_AS
                                Dim 单元_hcE1ε As Single = 单元_AE1Pε / 单元_AE1ε * (-单元_tP / 2) + 单元_AS / 单元_AE1ε * 单元_hcS
                                Dim 单元_IoE1ε As Single = 单元_IoE1Pε + 单元_IOS
                                Dim 单元_IcE1ε As Single = 单元_IoE1ε - 单元_AE1ε * 单元_hcE1ε ^ 2
                                Dim 单元_lE1Pε As Single = 单元_hcE1ε - (-单元_tP / 2)
                                Dim 单元_lE1Sε As Single = If(单元_tpS = "F" Or 单元_tpS = "B" Or 单元_tpS = "L3", 单元_hw - 单元_hcE1ε, 单元_hw + 单元_tf - 单元_hcE1ε)
                                Dim 单元_σYE1ε As Single = 单元_AS * 单元_lE1Sε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYS + 单元_AE1Pε * 单元_lE1Pε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYP
                                Dim 单元_σECε As Single = PI ^ 2 * 标准弹性模量 * 单元_IcE1ε / 单元_AEε / 单元_lS ^ 2
                                Dim 单元_σCCε As Single = If(单元_σECε <= 单元_σYE1ε / 2 * (-单元_εR), 单元_σECε / (-单元_εR), 单元_σYE1ε * (1 - 单元_σYE1ε * (-单元_εR) / 4 / 单元_σECε))
                                Dim 单元_σBCε As Single = 单元_Φε * 单元_σCCε * 单元_AEε / 单元_A

                                Dim 单元_Rlt_BC As String = Format(单元_σBCε, "000.000")

                                Dim 单元_σCTSε As Single = If(单元_σETS <= 单元_σYS / 2 * (-单元_εR), 单元_σETS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σETS))
                                Dim 单元_σCPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_σYP, 单元_σYP)
                                Dim 单元_σFTε As Single = 单元_Φε * (单元_AS * 单元_σCTSε + 单元_AP * 单元_σCPε) / 单元_A

                                Dim 单元_Rlt_FT As String = Format(单元_σFTε, "000.000")

                                Dim 单元_σCLSε As Single = If(单元_σELS <= 单元_σYS / 2 * (-单元_εR), 单元_σELS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σELS))
                                Dim 单元_σWBε As Single = If(单元_tpS = "F", 单元_Φε * (单元_AS * 单元_σCLSε + 单元_AP * 单元_σCPε) / 单元_A, 单元_Φε * (单元_AESε * 单元_σYS + 单元_AEPε * 单元_σYP) / 单元_A)

                                Dim 单元_Rlt_WB As String = Format(单元_σWBε, "000.000")

                                ''''''
                                单元_σO = Max(单元_σBCε, Max(单元_σFTε, 单元_σWBε)) ' 单元_σO = 单元_σEPε   ''''''NO BUCKLING
                                            'If 四级序数 = 1 Then Debug.Print(单元_Rlt_EP & ", " & 单元_Rlt_BC & ", " & 单元_Rlt_FT & ", " & 单元_Rlt_WB)
                            Case > 0
                                单元_σO = 单元_σEPε
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 5
                    '
                Case 6
                    For 四级序数 As UShort = 1 To 加筋板单元_总数
                        Dim 单元_原始_Yc As Single = 加筋板单元_原始_Yc(四级序数) / 1000
                        Dim 单元_原始_Zc As Single = 加筋板单元_原始_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_原始_A As Single = 加筋板单元_原始_A(四级序数)
                        Dim 单元_原始_σY As Single = 加筋板单元_原始_σY(四级序数)

                        Dim 单元_剩余_Yc As Single = 加筋板单元_剩余_Yc(四级序数) / 1000
                        Dim 单元_剩余_Zc As Single = 加筋板单元_剩余_Zc(四级序数) / 1000
                        Dim 单元_剩余_A As Single = 加筋板单元_剩余_A(四级序数)
                        Dim 单元_剩余_σY As Single = 加筋板单元_剩余_σY(四级序数)

                        Dim 关联板格序数 As UShort = 加筋板单元_板格_序数(四级序数)

                        '[注意正负!]
                        单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_原始_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY

                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        Dim EP As Single = Φ * 单元_原始_σY

                        'wP / lP
                        Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)

                        Select Case 单元_εR
                            Case < 0
                                Dim βPε As Single = βOP * (-单元_εR) ^ (1 / 2)
                                Dim FB As Single = Φ * 单元_原始_σY * Min(1, 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数) * (2.25 / βPε - 1.25 / βPε ^ 2) + 0.1 * (1 - 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数)) * (1 + 1 / βPε ^ 2))
                                单元_σO = FB
                            Case > 0
                                单元_σO = EP
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_剩余_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case Else
                    '
            End Select
        Next

        Dim 截面_YFP As Single, 截面_YFN As Single,
            截面_ZFP As Single, 截面_ZFN As Single
        Dim VP As Single
        截面_YFP = 截面_FYP / 截面_FP
        截面_YFN = 截面_FYN / 截面_FN
        截面_ZFP = 截面_FZP / 截面_FP
        截面_ZFN = 截面_FZN / 截面_FN
        VP = Cos(水线倾角_α) * (截面_YFP - 截面_YFN) / Sqrt((截面_ZFP - 截面_ZFN) ^ 2 + (截面_YFP - 截面_YFN) ^ 2) + Sin(水线倾角_α) * (截面_ZFP - 截面_ZFN) / Sqrt((截面_ZFP - 截面_ZFN) ^ 2 + (截面_YFP - 截面_YFN) ^ 2)
        判据参数4 = VP
    End Sub

    Private Sub 最终弯矩γ(ByVal χ_瞬时 As Single, ByVal γ_瞬时 As Single, ByVal α_瞬时 As Single, ByRef 结果 As Single)
        Dim 截面_FO As Single, 截面_FA As Single,
            截面_FP As Single, 截面_FN As Single,
            截面_FYP As Single, 截面_FYN As Single,
            截面_FZP As Single, 截面_FZN As Single,
            截面_MO As Single

        截面_FO = 0
        截面_FA = 0
        截面_FP = 0
        截面_FN = 0
        截面_FYP = 0
        截面_FYN = 0
        截面_FZP = 0
        截面_FZN = 0
        截面_MO = 0

        For 单元类型 As UShort = 1 To 6
            Select Case 单元类型
                Case 1
                    For 四级序数 As UShort = 1 To 硬角单元_总数
                        单元_Yc = 硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 硬角单元_A(四级序数) / 1000000
                        单元_σY = 硬角单元_σY(四级序数)
                        单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 2
                    '
                Case 3
                    For 四级序数 As UShort = 1 To 特别硬角单元_总数
                        单元_Yc = 特别硬角单元_Yc(四级序数) / 1000
                        单元_Zc = 特别硬角单元_Zc(四级序数) / 1000
                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select
                        单元_A = 特别硬角单元_A(四级序数) / 1000000
                        单元_σY = 特别硬角单元_σY(四级序数)
                        单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时
                        单元_εY = 单元_σY / 标准弹性模量
                        单元_εR = 单元_εO / 单元_εY
                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))
                        Dim EP As Single = Φ * 单元_σY
                        单元_σO = EP
                        单元_FO = 单元_σO * 单元_A
                        单元_MO = 单元_FO * 单元_L
                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 4
                    For 四级序数 As UShort = 1 To 加强筋单元_总数
                        Dim 单元_Yc As Single = 加强筋单元_Yc(四级序数) / 1000
                        Dim 单元_Zc As Single = 加强筋单元_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_A As Single = 加强筋单元_A(四级序数)
                        Dim 单元_σY As Single = 加强筋单元_σY(四级序数)

                        Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(四级序数)
                        Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)

                        Dim 单元_σYP As Single = 加强筋单元_σYP(四级序数)
                        Dim 单元_lP As Single = 加强筋单元_lP(四级序数)
                        Dim 单元_wP As Single = 加强筋单元_wP(四级序数)
                        Dim 单元_tP As Single = 加强筋单元_tP(四级序数)
                        Dim 单元_σYS As Single = 加强筋单元_σYS(四级序数)
                        Dim 单元_lS As Single = 加强筋单元_lS(四级序数)
                        Dim 单元_hw As Single = 加强筋单元_hw(四级序数)
                        Dim 单元_tw As Single = 加强筋单元_tw(四级序数)
                        Dim 单元_wf As Single = 加强筋单元_wf(四级序数)
                        Dim 单元_tf As Single = 加强筋单元_tf(四级序数)
                        Dim 单元_dx As Single = 加强筋单元_dx(四级序数)
                        Dim 单元_tpS As String = 加强筋单元_tpS(四级序数)
                        Dim 单元_mk As Boolean = 加强筋单元_mk(四级序数)
                        'Dim 单元_σETS As Single = 加强筋单元_σETS(四级序数)
                        'Dim 单元_σELS As Single = 加强筋单元_σELS(四级序数)

                        Dim 单元_εYP As Single = 单元_σYP / 标准弹性模量
                        Dim 单元_AP As Single = 单元_wP * 单元_tP
                        Dim 单元_ICP As Single = 单元_wP * 单元_tP ^ 3 / 12
                        Dim 单元_IOP As Single = 单元_ICP + 单元_AP * (-单元_tP / 2) ^ 2
                        Dim 单元_βOP As Single = 单元_wP / 单元_tP * 单元_εYP ^ (1 / 2)
                        Dim 单元_wEoP As Single = If(单元_βOP >= 1.25, (2.25 / 单元_βOP - 1.25 / 单元_βOP ^ 2) * 单元_wP, 单元_wP)

                        Dim 单元_εYS As Single = 单元_σYS / 标准弹性模量
                        Dim 单元_Aw As Single = 单元_hw * 单元_tw
                        Dim 单元_Icw As Single = 单元_tw * 单元_hw ^ 3 / 12
                        Dim 单元_Iow As Single = 单元_Icw + 单元_Aw * (单元_hw / 2) ^ 2
                        Dim 单元_βOw As Single = 单元_hw / 单元_tw * 单元_εYS ^ (1 / 2)
                        Dim 单元_hEow As Single = If(单元_βOw >= 1.25, (2.25 / 单元_βOw - 1.25 / 单元_βOw ^ 2) * 单元_hw, 单元_hw)
                        Dim 单元_df As Single = If(单元_tpS = "F", 单元_hw, If(单元_tpS = "B", 单元_hw - 单元_tf / 2, If(单元_tpS = "T" Or 单元_tpS = "L1" Or 单元_tpS = "L2", 单元_hw + 单元_tf / 2, 单元_hw - 单元_dx - 单元_tf / 2)))
                        Dim 单元_Af As Single = 单元_wf * 单元_tf
                        Dim 单元_Icf As Single = 单元_wf * 单元_tf ^ 3 / 12
                        Dim 单元_Iof As Single = 单元_Icf + 单元_Af * 单元_df ^ 2
                        Dim 单元_AS As Single = 单元_Aw + 单元_Af
                        Dim 单元_hcS As Single = 单元_Aw / 单元_AS * 单元_hw / 2 + 单元_Af / 单元_AS * 单元_df
                        Dim 单元_IOS As Single = 单元_Iow + 单元_Iof
                        Dim 单元_ICS As Single = 单元_Icw + 单元_Aw * (单元_hw / 2 - 单元_hcS) ^ 2 + 单元_Icf + 单元_Af * (单元_df - 单元_hcS) ^ 2
                        Dim 单元_IpS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw / 3, 单元_Aw * (单元_df - 单元_tf / 2) ^ 2 / 3 + 单元_Af * 单元_df ^ 2)
                        Dim 单元_ItS As Single = If(单元_tpS = "F", 单元_hw * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / 单元_hw), (单元_df - 单元_tf / 2) * 单元_tw ^ 3 / 3 * (1 - 0.63 * 单元_tw / (单元_df - 单元_tf / 2)) + 单元_wf * 单元_tf ^ 3 / 3 * (1 - 0.63 * 单元_tf / 单元_wf))
                        Dim 单元_IwS As Single = If(单元_tpS = "F", 单元_hw ^ 3 * 单元_tw ^ 3 / 36, If(单元_tpS = "B" Or 单元_tpS = "L1" Or 单元_tpS = "L2" Or 单元_tpS = "L3", 单元_Af * 单元_df ^ 2 * 单元_wf ^ 2 / 12 * (单元_Af + 2.6 * 单元_Aw) / (单元_Af * 单元_Aw), 单元_wf ^ 3 * 单元_tf * 单元_df ^ 2 / 12))
                        Dim 单元_ηS As Single = 1 + (单元_lS / PI) ^ 2 / (单元_IwS * (0.75 * 单元_wP / 单元_tP ^ 3 + (单元_df - 单元_tf / 2) / 单元_tw ^ 3)) ^ (1 / 2)
                        Dim 单元_σETS As Single = 标准弹性模量 / 单元_IpS * (单元_ηS * PI ^ 2 * 单元_IwS / 单元_lS ^ 2 + 0.385 * 单元_ItS)
                        Dim 单元_σELS As Single = 160000 * (单元_tw / 单元_hw) ^ 2

                        'Dim 单元_A As Single = 单元_AS + 单元_AP
                        Dim 单元_hc As Single = 单元_AS / 单元_A * 单元_hcS + 单元_AP / 单元_A * (-单元_tP / 2)
                        Dim 单元_Io As Single = 单元_IOP + 单元_IOS
                        Dim 单元_Ic As Single = 单元_Io - 单元_A * 单元_hc ^ 2
                        'Dim 单元_σY As Single = 单元_AS / A * 单元_σYS + 单元_AP / 单元_A * 单元_σYP
                        'Dim 单元_εY As Single = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        Dim 单元_L As Single = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY

                        'Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        'Dim EP As Single = Φ * 单元_σY

                        'wP / lP
                        'Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)
                        'Dim βPε As Single = βOP * 单元_εR ^ (1 / 2)
                        'Dim BC As Single, FT As Single, WB As Single

                        '单元_εR = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))    ''''''NO STRENGTH REDUCTION AFTER BUCKLING
                        Dim 单元_Φε As Single = If(单元_εR < -1, -1, If(单元_εR >= -1 And 单元_εR <= 1, 单元_εR, 1))
                        Dim 单元_σEPε As Single = 单元_Φε * 单元_σY

                        Dim 单元_Rlt_EP As String = Format(单元_σEPε, "000.000")

                        Select Case 单元_εR
                            Case < 0
                                Dim 单元_βPε As Single = 单元_βOP * (-单元_εR) ^ (1 / 2)
                                Dim 单元_wEPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_wP, 单元_wP)
                                Dim 单元_wE1Pε As Single = If(单元_βPε >= 1, 单元_wP / 单元_βPε, 单元_wP)
                                Dim 单元_AEPε As Single = 单元_wEPε * 单元_tP
                                Dim 单元_AE1Pε As Single = 单元_wE1Pε * 单元_tP
                                Dim 单元_IcE1Pε As Single = 单元_wE1Pε * 单元_tP ^ 3 / 12
                                Dim 单元_IoE1Pε As Single = 单元_IcE1Pε + 单元_AE1Pε * (-单元_tP / 2) ^ 2

                                Dim 单元_βwε As Single = 单元_βOw * (-单元_εR) ^ (1 / 2)
                                Dim 单元_hEwε As Single = If(单元_βwε >= 1.25, (2.25 / 单元_βwε - 1.25 / 单元_βwε ^ 2) * 单元_hw, 单元_hw)
                                Dim 单元_AEwε As Single = 单元_hEwε * 单元_tw
                                Dim 单元_AESε As Single = 单元_AEwε + 单元_Af

                                Dim 单元_AEε As Single = 单元_AEPε + 单元_AS
                                Dim 单元_AE1ε As Single = 单元_AE1Pε + 单元_AS
                                Dim 单元_hcE1ε As Single = 单元_AE1Pε / 单元_AE1ε * (-单元_tP / 2) + 单元_AS / 单元_AE1ε * 单元_hcS
                                Dim 单元_IoE1ε As Single = 单元_IoE1Pε + 单元_IOS
                                Dim 单元_IcE1ε As Single = 单元_IoE1ε - 单元_AE1ε * 单元_hcE1ε ^ 2
                                Dim 单元_lE1Pε As Single = 单元_hcE1ε - (-单元_tP / 2)
                                Dim 单元_lE1Sε As Single = If(单元_tpS = "F" Or 单元_tpS = "B" Or 单元_tpS = "L3", 单元_hw - 单元_hcE1ε, 单元_hw + 单元_tf - 单元_hcE1ε)
                                Dim 单元_σYE1ε As Single = 单元_AS * 单元_lE1Sε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYS + 单元_AE1Pε * 单元_lE1Pε / (单元_AS * 单元_lE1Sε + 单元_AE1Pε * 单元_lE1Pε) * 单元_σYP
                                Dim 单元_σECε As Single = PI ^ 2 * 标准弹性模量 * 单元_IcE1ε / 单元_AEε / 单元_lS ^ 2
                                Dim 单元_σCCε As Single = If(单元_σECε <= 单元_σYE1ε / 2 * (-单元_εR), 单元_σECε / (-单元_εR), 单元_σYE1ε * (1 - 单元_σYE1ε * (-单元_εR) / 4 / 单元_σECε))
                                Dim 单元_σBCε As Single = 单元_Φε * 单元_σCCε * 单元_AEε / 单元_A

                                Dim 单元_Rlt_BC As String = Format(单元_σBCε, "000.000")

                                Dim 单元_σCTSε As Single = If(单元_σETS <= 单元_σYS / 2 * (-单元_εR), 单元_σETS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σETS))
                                Dim 单元_σCPε As Single = If(单元_βPε >= 1.25, (2.25 / 单元_βPε - 1.25 / 单元_βPε ^ 2) * 单元_σYP, 单元_σYP)
                                Dim 单元_σFTε As Single = 单元_Φε * (单元_AS * 单元_σCTSε + 单元_AP * 单元_σCPε) / 单元_A

                                Dim 单元_Rlt_FT As String = Format(单元_σFTε, "000.000")

                                Dim 单元_σCLSε As Single = If(单元_σELS <= 单元_σYS / 2 * (-单元_εR), 单元_σELS / (-单元_εR), 单元_σYS * (1 - 单元_σYS * (-单元_εR) / 4 / 单元_σELS))
                                Dim 单元_σWBε As Single = If(单元_tpS = "F", 单元_Φε * (单元_AS * 单元_σCLSε + 单元_AP * 单元_σCPε) / 单元_A, 单元_Φε * (单元_AESε * 单元_σYS + 单元_AEPε * 单元_σYP) / 单元_A)

                                Dim 单元_Rlt_WB As String = Format(单元_σWBε, "000.000")

                                ''''''
                                单元_σO = Max(单元_σBCε, Max(单元_σFTε, 单元_σWBε)) ' 单元_σO = 单元_σEPε   ''''''NO BUCKLING
                                            'If 四级序数 = 1 Then Debug.Print(单元_Rlt_EP & ", " & 单元_Rlt_BC & ", " & 单元_Rlt_FT & ", " & 单元_Rlt_WB)
                            Case > 0
                                单元_σO = 单元_σEPε
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case 5
                    '
                Case 6
                    For 四级序数 As UShort = 1 To 加筋板单元_总数
                        Dim 单元_原始_Yc As Single = 加筋板单元_原始_Yc(四级序数) / 1000
                        Dim 单元_原始_Zc As Single = 加筋板单元_原始_Zc(四级序数) / 1000

                        Select Case 第一部分
                            Case "OT"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 25.375 And 单元_Zc > 12 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -12.4 And 单元_Yc < 22.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -17.4 And 单元_Yc < 17.4 And 单元_Zc < 2 Then Continue For
                                End Select
                            Case "BC"
                                Select Case 第二部分
                                    Case "0"
                                        Exit Select
                                    Case "1"
                                        If 单元_Yc > 21.875 And 单元_Zc > 6.75 Then Continue For
                                    Case "2"
                                        If 单元_Yc > -11.6 And 单元_Yc < 18.4 And 单元_Zc < 2 Then Continue For
                                    Case "3"
                                        If 单元_Yc > -15 And 单元_Yc < 15 And 单元_Zc < 2 Then Continue For
                                End Select
                        End Select

                        Dim 单元_原始_A As Single = 加筋板单元_原始_A(四级序数)
                        Dim 单元_原始_σY As Single = 加筋板单元_原始_σY(四级序数)

                        Dim 单元_剩余_Yc As Single = 加筋板单元_剩余_Yc(四级序数) / 1000
                        Dim 单元_剩余_Zc As Single = 加筋板单元_剩余_Zc(四级序数) / 1000
                        Dim 单元_剩余_A As Single = 加筋板单元_剩余_A(四级序数)
                        Dim 单元_剩余_σY As Single = 加筋板单元_剩余_σY(四级序数)

                        Dim 关联板格序数 As UShort = 加筋板单元_板格_序数(四级序数)

                        '[注意正负!]
                        单元_L = -(单元_Yc - γ_瞬时) * Sin(α_瞬时) + 单元_Zc * Cos(α_瞬时)
                        '[注意正负!]
                        单元_εO = 单元_L * χ_瞬时

                        单元_εY = 单元_原始_σY / 标准弹性模量

                        '[注意正负!]
                        单元_εR = 单元_εO / 单元_εY

                        Dim Φ As Single = If(单元_εR > 1, 1, If(单元_εR < -1, -1, 单元_εR))

                        Dim EP As Single = Φ * 单元_原始_σY

                        'wP / lP
                        Dim βOP As Single = 板格_l(关联板格序数) / 板格_t(关联板格序数) * 单元_εY ^ (1 / 2)

                        Select Case 单元_εR
                            Case < 0
                                Dim βPε As Single = βOP * (-单元_εR) ^ (1 / 2)
                                Dim FB As Single = Φ * 单元_原始_σY * Min(1, 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数) * (2.25 / βPε - 1.25 / βPε ^ 2) + 0.1 * (1 - 板格_l(关联板格序数) / 加筋板单元_原始_w(四级序数)) * (1 + 1 / βPε ^ 2))
                                单元_σO = FB
                            Case > 0
                                单元_σO = EP
                            Case = 0
                                单元_σO = 0
                            Case Else

                        End Select

                        '考虑[单元_剩余_A]!
                        单元_FO = 单元_σO * 单元_剩余_A / 1000000
                        单元_MO = 单元_FO * 单元_L

                        Select Case 单元_FO
                            Case > 0
                                截面_FP += 单元_FO
                                截面_FYP += 单元_FO * 单元_Yc
                                截面_FZP += 单元_FO * 单元_Zc
                            Case < 0
                                截面_FN += 单元_FO
                                截面_FYN += 单元_FO * 单元_Yc
                                截面_FZN += 单元_FO * 单元_Zc
                            Case Else
                                '
                        End Select
                        截面_FO += 单元_FO
                        截面_FA += Abs(单元_FO)
                        截面_MO += 单元_MO
                    Next
                Case Else
                    '
            End Select
        Next

        'Dim 截面_YFP As Single, 截面_YFN As Single,
        '    截面_ZFP As Single, 截面_ZFN As Single
        'Dim α_RF As Single
        'Dim VP As Single
        '截面_YFP = 截面_FYP / 截面_FP
        '截面_YFN = 截面_FYN / 截面_FN
        '截面_ZFP = 截面_FZP / 截面_FP
        '截面_ZFN = 截面_FZN / 截面_FN
        'α_RF = Atan((截面_ZFP - 截面_ZFN) / (截面_YFP - 截面_YFN))
        'VP = Cos(水线倾角_α) * Cos(α_RF) + Sin(水线倾角_α) * Sin(α_RF)
        结果 = 截面_MO
    End Sub

    Private Sub 随机性计算() Handles Button8.Click
        样本总数 = InputBox("样本总数", "样本总数", "2000")

        Dim xlApp As Application
        Dim xlBook As Workbook
        Dim xlSheet As Worksheet

        OpenFileDialog1.ShowDialog()

        xlApp = CType(CreateObject("Excel.Application"), Application)
        xlBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName,, True)
        xlSheet = CType(xlBook.Worksheets(1), Worksheet)

        '调用 共有过程_通用型.读入界面参数
        读入界面参数()

        Dim 行号 As UShort
        For 基本输入对象类型序数 As UShort = 1 To 5
            Select Case 基本输入对象类型序数
                Case 1
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "节点"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z/(mm)"
                    For 一级序数 As UShort = 1 To 节点_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        节点_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        节点_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 1 / 4
                Case 2
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "加强筋"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    'xlSheet.Cells(行号, 4) = "l/(mm)"
                    'xlSheet.Cells(行号, 5) = "hw/(mm)"
                    'xlSheet.Cells(行号, 6) = "tw/(mm)"
                    'xlSheet.Cells(行号, 7) = "αw/(rad)"
                    'xlSheet.Cells(行号, 8) = "wf/(mm)"
                    'xlSheet.Cells(行号, 9) = "tf/(mm)"
                    'xlSheet.Cells(行号, 10) = "αf/(rad)"
                    'xlSheet.Cells(行号, 11) = "dx/(mm)"
                    'xlSheet.Cells(行号, 12) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 13) = "tp(F/B/T/L1/L2/L3)"
                    'xlSheet.Cells(行号, 14) = "mk(TRUE/FALSE)"
                    For 一级序数 As UShort = 1 To 加强筋_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        加强筋_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        加强筋_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                        加强筋_l(一级序数) = xlSheet.Cells(行号, 4).value
                        加强筋_hw(一级序数) = xlSheet.Cells(行号, 5).value
                        加强筋_tw(一级序数) = xlSheet.Cells(行号, 6).value
                        加强筋_twX(一级序数) = 加强筋_tw(一级序数)    '储存名义值, 用于后续的随机变量的产生
                        'BoxMuller(加强筋_twX(一级序数), 加强筋_tw(一级序数))
                        加强筋_αw(一级序数) = xlSheet.Cells(行号, 7).value
                        加强筋_wf(一级序数) = xlSheet.Cells(行号, 8).value
                        加强筋_tf(一级序数) = xlSheet.Cells(行号, 9).value
                        加强筋_tfX(一级序数) = 加强筋_tf(一级序数)    '储存名义值, 用于后续的随机变量的产生
                        'BoxMuller(加强筋_tfX(一级序数), 加强筋_tf(一级序数))
                        加强筋_αf(一级序数) = xlSheet.Cells(行号, 10).value
                        加强筋_dx(一级序数) = xlSheet.Cells(行号, 11).value
                        加强筋_σY(一级序数) = xlSheet.Cells(行号, 12).value
                        '加强筋_σYX(一级序数) = 加强筋_σY(一级序数)    '储存名义值, 用于后续的随机变量的产生
                        加强筋_tp(一级序数) = xlSheet.Cells(行号, 13).value
                        加强筋_mk(一级序数) = xlSheet.Cells(行号, 14).value
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 2 / 4
                Case 3
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "面板"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    'xlSheet.Cells(行号, 4) = "Y1/(mm)"
                    'xlSheet.Cells(行号, 5) = "Z1/(mm)"
                    'xlSheet.Cells(行号, 6) = "YL/(mm)"
                    'xlSheet.Cells(行号, 7) = "ZL/(mm)"
                    'xlSheet.Cells(行号, 8) = "l/(mm)"
                    'xlSheet.Cells(行号, 9) = "t/(mm)"
                    'xlSheet.Cells(行号, 10) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 11) = "PMA(TRUE/FALSE)"
                    For 一级序数 As UShort = 1 To 面板_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        面板_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        面板_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                        面板_Y1(一级序数) = xlSheet.Cells(行号, 4).value
                        面板_Z1(一级序数) = xlSheet.Cells(行号, 5).value
                        面板_YL(一级序数) = xlSheet.Cells(行号, 6).value
                        面板_ZL(一级序数) = xlSheet.Cells(行号, 7).value
                        面板_l(一级序数) = xlSheet.Cells(行号, 8).value
                        面板_t(一级序数) = xlSheet.Cells(行号, 9).value
                        面板_tX(一级序数) = 面板_t(一级序数)    '储存名义值, 用于后续的随机变量的产生
                        'BoxMuller(面板_tX(一级序数), 面板_t(一级序数))
                        面板_σY(一级序数) = xlSheet.Cells(行号, 10).value
                        '面板_σYX(一级序数) = 面板_σY(一级序数)    '储存名义值, 用于后续的随机变量的产生
                        面板_PMA(一级序数) = xlSheet.Cells(行号, 11).value
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 3 / 4
                Case 4
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "板格"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "Y0/(mm)"
                    'xlSheet.Cells(行号, 3) = "Z0/(mm)"
                    'xlSheet.Cells(行号, 4) = "Y1/(mm)"
                    'xlSheet.Cells(行号, 5) = "Z1/(mm)"
                    'xlSheet.Cells(行号, 6) = "l(mm)"
                    'xlSheet.Cells(行号, 7) = "tp(L/T)"
                    'xlSheet.Cells(行号, 8) = "板格板数目"
                    'xlSheet.Cells(行号, 9) = "w/(mm)"
                    'xlSheet.Cells(行号, 10) = "t/(mm)"
                    'xlSheet.Cells(行号, 11) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 12) = "w/(mm)"
                    'xlSheet.Cells(行号, 13) = "t/(mm)"
                    'xlSheet.Cells(行号, 14) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 15) = "......"
                    'xlSheet.Cells(行号, 16) = "......"
                    'xlSheet.Cells(行号, 17) = "......"
                    For 一级序数 As UShort = 1 To 板格_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        板格_Y0(一级序数) = xlSheet.Cells(行号, 2).value
                        板格_Z0(一级序数) = xlSheet.Cells(行号, 3).value
                        板格_Y1(一级序数) = xlSheet.Cells(行号, 4).value
                        板格_Z1(一级序数) = xlSheet.Cells(行号, 5).value
                        板格_l(一级序数) = xlSheet.Cells(行号, 6).value
                        板格_tp(一级序数) = xlSheet.Cells(行号, 7).value
                        板格_板格板数目(一级序数) = xlSheet.Cells(行号, 8).value
                        For 二级序数 As UShort = 1 To 板格_板格板数目(一级序数)
                            板格板_w(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 9).value
                            板格板_t(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 10).value
                            板格板_tX(一级序数, 二级序数) = 板格板_t(一级序数, 二级序数)    '储存名义值, 用于后续的随机变量的产生
                            'BoxMuller(板格板_tX(一级序数, 二级序数), 板格板_t(一级序数, 二级序数))
                            板格板_σY(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 11).value
                            '板格板_σYX(一级序数, 二级序数) = 板格板_σY(一级序数, 二级序数)    '储存名义值, 用于后续的随机变量的产生
                        Next
                    Next
                    ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * 4 / 4
                Case 5
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "特别硬角单元"
                    行号 += 1
                    'xlSheet.Cells(行号, 1) = "序数"
                    'xlSheet.Cells(行号, 2) = "子对象数目"
                    'xlSheet.Cells(行号, 3) = "Yc/(mm)"
                    'xlSheet.Cells(行号, 4) = "Zc/(mm)"
                    'xlSheet.Cells(行号, 5) = "A/(mm^2)"
                    'xlSheet.Cells(行号, 6) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 7) = "Yc/(mm)"
                    'xlSheet.Cells(行号, 8) = "Zc/(mm)"
                    'xlSheet.Cells(行号, 9) = "A/(mm^2)"
                    'xlSheet.Cells(行号, 10) = "σY/(MPa)"
                    'xlSheet.Cells(行号, 11) = "......"
                    'xlSheet.Cells(行号, 12) = "......"
                    'xlSheet.Cells(行号, 13) = "......"
                    'xlSheet.Cells(行号, 14) = "......"
                    For 一级序数 As UShort = 1 To 特别硬角单元_总数
                        行号 += 1
                        'xlSheet.Cells(行号, 1) = 一级序数
                        子对象_数目(一级序数) = xlSheet.Cells(行号, 2).value
                        For 二级序数 As UShort = 1 To 子对象_数目(一级序数)
                            子对象_Yc(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 3).value
                            子对象_Zc(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 4).value
                            子对象_A(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 5).value
                            子对象_AX(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数)
                            子对象_σY(一级序数, 二级序数) = xlSheet.Cells(行号, (二级序数 - 1) * 3 + 6).value
                            '子对象_σYX(一级序数, 二级序数) = 子对象_σY(一级序数, 二级序数)    '储存名义值, 用于后续的随机变量的产生
                        Next
                    Next
            End Select
        Next

        xlBook.Close()
        xlApp.Quit() 'xlApp = Nothing

        ''''''
        For 循环序数 As UShort = 0 To 样本总数
            '基本输入对象的属性计算及形心坐标输出
            ReDim 板格_A(板格_总数), 板格_AYc(板格_总数), 板格_AZc(板格_总数), 板格_AσY(板格_总数)
            ReDim 特别硬角单元_A(特别硬角单元_总数), 特别硬角单元_AYc(特别硬角单元_总数),
                特别硬角单元_AZc(特别硬角单元_总数), 特别硬角单元_AσY(特别硬角单元_总数)
            If 循环序数 = 0 Then
                For 基本输入对象类型序数 As UShort = 1 To 5
                    Select Case 基本输入对象类型序数
                        Case 1
                            Chart1.Series.Add("节点_形心")
                            Chart1.Series("节点_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                            For 一级序数 As UShort = 1 To 节点_总数
                                Chart1.Series("节点_形心").Points.AddXY(节点_Y0(一级序数), 节点_Z0(一级序数))
                            Next
                        Case 2
                            Chart1.Series.Add("加强筋_形心")
                            Chart1.Series("加强筋_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                            For 一级序数 As UShort = 1 To 加强筋_总数
                                加强筋_tw(一级序数) = 加强筋_twX(一级序数) : 加强筋_tf(一级序数) = 加强筋_tfX(一级序数) '加强筋_σY(一级序数) = 加强筋_σYX(一级序数)

                                加强筋_Aw(一级序数) = 加强筋_hw(一级序数) * 加强筋_tw(一级序数)
                                加强筋_Icw(一级序数) = 加强筋_tw(一级序数) * 加强筋_hw(一级序数) ^ 3 / 12
                                加强筋_Iow(一级序数) = 加强筋_Icw(一级序数) + 加强筋_Aw(一级序数) * (加强筋_hw(一级序数) / 2) ^ 2

                                Select Case 加强筋_tp(一级序数)
                                    Case "F"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数)
                                    Case "B"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) - 加强筋_tf(一级序数) / 2
                                    Case "T"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                                    Case "L1"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                                    Case "L2"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                                    Case "L3"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) - 加强筋_dx(一级序数) - 加强筋_tf(一级序数) / 2
                                    Case Else
                                        MsgBox("加强筋_tp(" & 一级序数 & ")错误：类型不符")
                                End Select
                                加强筋_Af(一级序数) = 加强筋_wf(一级序数) * 加强筋_tf(一级序数)
                                加强筋_Icf(一级序数) = 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 12
                                加强筋_Iof(一级序数) = 加强筋_Icf(一级序数) + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2

                                加强筋_AS(一级序数) = 加强筋_Aw(一级序数) + 加强筋_Af(一级序数)
                                加强筋_hcS(一级序数) = 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_hw(一级序数) / 2 + 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_df(一级序数)
                                加强筋_IoS(一级序数) = 加强筋_Iow(一级序数) + 加强筋_Iof(一级序数)
                                加强筋_IcS(一级序数) = 加强筋_Icw(一级序数) + 加强筋_Aw(一级序数) * (加强筋_hw(一级序数) / 2 - 加强筋_hcS(一级序数)) ^ 2 + 加强筋_Icf(一级序数) + 加强筋_Af(一级序数) * (加强筋_df(一级序数) - 加强筋_hcS(一级序数)) ^ 2

                                Select Case 加强筋_tp(一级序数)
                                    Case "F"
                                        加强筋_IPS(一级序数) = 加强筋_hw(一级序数) ^ 3 * 加强筋_tw(一级序数) / 3
                                        加强筋_ITS(一级序数) = 加强筋_hw(一级序数) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / 加强筋_hw(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_hw(一级序数) ^ 3 * 加强筋_tw(一级序数) ^ 3 / 36
                                    Case "B"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                                    Case "T"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_wf(一级序数) ^ 3 * 加强筋_tf(一级序数) * 加强筋_df(一级序数) ^ 2 / 12
                                    Case "L1"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                                    Case "L2"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                                    Case "L3"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                                    Case Else

                                End Select
                                '加强筋_ηS(一级序数) = 1 + 加强筋_l(一级序数) ^ 2 / PI ^ 2 / Sqrt(加强筋_IWS(一级序数) * (0.75 * 加强筋带板_w(一级序数) / 加强筋带板_t(一级序数) ^ 3 + (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) / 加强筋_tw(一级序数) ^ 3))
                                '加强筋_σETS(一级序数) = 标准弹性模量 / 加强筋_IPS(一级序数) * (加强筋_ηS(一级序数) * PI ^ 2 * 加强筋_IWS(一级序数) / 加强筋_lS(一级序数) ^ 2 + 0.385 * 加强筋_ITS(一级序数))
                                加强筋_σELS(一级序数) = 160000 * (加强筋_tw(一级序数) / 加强筋_hw(一级序数)) ^ 2

                                加强筋_Ycw(一级序数) = 加强筋_Y0(一级序数) + 加强筋_hw(一级序数) * Cos(加强筋_αw(一级序数)) / 2
                                加强筋_Zcw(一级序数) = 加强筋_Z0(一级序数) + 加强筋_hw(一级序数) * Sin(加强筋_αw(一级序数)) / 2

                                加强筋_Ycf(一级序数) = 加强筋_Y0(一级序数) + 加强筋_df(一级序数) * Cos(加强筋_αw(一级序数)) + 加强筋_wf(一级序数) * Cos(加强筋_αf(一级序数)) / 2
                                加强筋_Zcf(一级序数) = 加强筋_Z0(一级序数) + 加强筋_df(一级序数) * Sin(加强筋_αw(一级序数)) + 加强筋_wf(一级序数) * Sin(加强筋_αf(一级序数)) / 2

                                加强筋_YcS(一级序数) = 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_Ycf(一级序数) + 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_Ycw(一级序数)
                                加强筋_ZcS(一级序数) = 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_Zcf(一级序数) + 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_Zcw(一级序数)

                                Chart1.Series("加强筋_形心").Points.AddXY(加强筋_YcS(一级序数), 加强筋_ZcS(一级序数))
                            Next
                        Case 3
                            Chart1.Series.Add("面板_形心")
                            Chart1.Series("面板_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                            For 一级序数 As UShort = 1 To 面板_总数
                                面板_t(一级序数) = 面板_tX(一级序数) '面板_σY(一级序数) = 面板_σYX(一级序数)

                                面板_w(一级序数) = Sqrt((面板_Y0(一级序数) - 面板_Y1(一级序数)) ^ 2 + (面板_Z0(一级序数) - 面板_Z1(一级序数)) ^ 2)
                                面板_A(一级序数) = 面板_w(一级序数) * 面板_t(一级序数)

                                面板_Yc(一级序数) = (面板_Y0(一级序数) + 面板_Y1(一级序数)) / 2
                                面板_Zc(一级序数) = (面板_Z0(一级序数) + 面板_Z1(一级序数)) / 2

                                Chart1.Series("面板_形心").Points.AddXY(面板_Yc(一级序数), 面板_Zc(一级序数))
                                Chart1.Series("面板_形心").Points.AddXY(面板_Y0(一级序数), 面板_Z0(一级序数))
                                Chart1.Series("面板_形心").Points.AddXY(面板_Y1(一级序数), 面板_Z1(一级序数))
                            Next
                        Case 4
                            Chart1.Series.Add("板格_形心")
                            Chart1.Series("板格_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                            For 一级序数 As UShort = 1 To 板格_总数
                                板格_α(一级序数) = If(板格_Y0(一级序数) = 板格_Y1(一级序数), PI / 2, Atan((板格_Z0(一级序数) - 板格_Z1(一级序数)) / (板格_Y0(一级序数) - 板格_Y1(一级序数))))
                                板格_w(一级序数) = Sqrt((板格_Y0(一级序数) - 板格_Y1(一级序数)) ^ 2 + (板格_Z0(一级序数) - 板格_Z1(一级序数)) ^ 2)
                                For 二级序数 As UShort = 1 To 板格_板格板数目(一级序数)
                                    板格板_t(一级序数, 二级序数) = 板格板_tX(一级序数, 二级序数) '板格板_σY(一级序数, 二级序数) = 板格板_σYX(一级序数, 二级序数)

                                    板格板_Y0(一级序数, 二级序数) = If(二级序数 = 1, 板格_Y0(一级序数), 板格板_Y0(一级序数, 二级序数 - 1))
                                    板格板_Z0(一级序数, 二级序数) = If(二级序数 = 1, 板格_Z0(一级序数), 板格板_Z0(一级序数, 二级序数 - 1))

                                    板格板_Y1(一级序数, 二级序数) = If(二级序数 = 板格_板格板数目(一级序数), 板格_Y1(一级序数), 板格板_Y0(一级序数, 二级序数) + 板格板_w(一级序数, 二级序数) * Cos(板格_α(一级序数)))
                                    板格板_Z1(一级序数, 二级序数) = If(二级序数 = 板格_板格板数目(一级序数), 板格_Z1(一级序数), 板格板_Z0(一级序数, 二级序数) + 板格板_w(一级序数, 二级序数) * Sin(板格_α(一级序数)))

                                    板格板_Yc(一级序数, 二级序数) = (板格板_Y0(一级序数, 二级序数) + 板格板_Y1(一级序数, 二级序数)) / 2
                                    板格板_Zc(一级序数, 二级序数) = (板格板_Z0(一级序数, 二级序数) + 板格板_Z1(一级序数, 二级序数)) / 2

                                    板格板_A(一级序数, 二级序数) = 板格板_w(一级序数, 二级序数) * 板格板_t(一级序数, 二级序数)
                                    板格板_AYc(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_Yc(一级序数, 二级序数)
                                    板格板_AZc(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_Zc(一级序数, 二级序数)
                                    板格板_AσY(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_σY(一级序数, 二级序数)

                                    板格_A(一级序数) += 板格板_A(一级序数, 二级序数)
                                    板格_AYc(一级序数) += 板格板_AYc(一级序数, 二级序数)
                                    板格_AZc(一级序数) += 板格板_AZc(一级序数, 二级序数)
                                    板格_AσY(一级序数) += 板格板_AσY(一级序数, 二级序数)
                                Next
                                板格_t(一级序数) = 板格_A(一级序数) / 板格_w(一级序数)
                                板格_Yc(一级序数) = 板格_AYc(一级序数) / 板格_A(一级序数)
                                板格_Zc(一级序数) = 板格_AZc(一级序数) / 板格_A(一级序数)
                                板格_σY(一级序数) = 板格_AσY(一级序数) / 板格_A(一级序数)

                                Chart1.Series("板格_形心").Points.AddXY(板格_Yc(一级序数), 板格_Zc(一级序数))
                            Next
                        Case 5
                            Chart1.Series.Add("特别硬角单元_形心")
                            Chart1.Series("特别硬角单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                            For 一级序数 As UShort = 1 To 特别硬角单元_总数
                                For 二级序数 As UShort = 1 To 子对象_数目(一级序数)
                                    子对象_A(一级序数, 二级序数) = 子对象_AX(一级序数, 二级序数) '子对象_σY(一级序数, 二级序数) = 子对象_σYX(一级序数, 二级序数)

                                    子对象_AYc(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_Yc(一级序数, 二级序数)
                                    子对象_AZc(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_Zc(一级序数, 二级序数)
                                    子对象_AσY(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_σY(一级序数, 二级序数)

                                    特别硬角单元_A(一级序数) += 子对象_A(一级序数, 二级序数)
                                    特别硬角单元_AYc(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_Yc(一级序数, 二级序数)
                                    特别硬角单元_AZc(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_Zc(一级序数, 二级序数)
                                    特别硬角单元_AσY(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_σY(一级序数, 二级序数)
                                Next
                                特别硬角单元_Yc(一级序数) = 特别硬角单元_AYc(一级序数) / 特别硬角单元_A(一级序数)
                                特别硬角单元_Zc(一级序数) = 特别硬角单元_AZc(一级序数) / 特别硬角单元_A(一级序数)
                                特别硬角单元_σY(一级序数) = 特别硬角单元_AσY(一级序数) / 特别硬角单元_A(一级序数)

                                Chart1.Series("特别硬角单元_形心").Points.AddXY(特别硬角单元_Yc(一级序数), 特别硬角单元_Zc(一级序数))
                            Next
                    End Select
                Next
            Else
                For 基本输入对象类型序数 As UShort = 1 To 5
                    Select Case 基本输入对象类型序数
                        Case 2
                            For 一级序数 As UShort = 1 To 加强筋_总数
                                BoxMuller(加强筋_twX(一级序数), 加强筋_tw(一级序数)) : BoxMuller(加强筋_tfX(一级序数), 加强筋_tf(一级序数)) 'BoxMuller(加强筋_σYX(一级序数), 加强筋_σY(一级序数))

                                加强筋_Aw(一级序数) = 加强筋_hw(一级序数) * 加强筋_tw(一级序数)
                                加强筋_Icw(一级序数) = 加强筋_tw(一级序数) * 加强筋_hw(一级序数) ^ 3 / 12
                                加强筋_Iow(一级序数) = 加强筋_Icw(一级序数) + 加强筋_Aw(一级序数) * (加强筋_hw(一级序数) / 2) ^ 2

                                Select Case 加强筋_tp(一级序数)
                                    Case "F"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数)
                                    Case "B"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) - 加强筋_tf(一级序数) / 2
                                    Case "T"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                                    Case "L1"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                                    Case "L2"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) + 加强筋_tf(一级序数) / 2
                                    Case "L3"
                                        加强筋_df(一级序数) = 加强筋_hw(一级序数) - 加强筋_dx(一级序数) - 加强筋_tf(一级序数) / 2
                                    Case Else
                                        MsgBox("加强筋_tp(" & 一级序数 & ")错误：类型不符")
                                End Select
                                加强筋_Af(一级序数) = 加强筋_wf(一级序数) * 加强筋_tf(一级序数)
                                加强筋_Icf(一级序数) = 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 12
                                加强筋_Iof(一级序数) = 加强筋_Icf(一级序数) + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2

                                加强筋_AS(一级序数) = 加强筋_Aw(一级序数) + 加强筋_Af(一级序数)
                                加强筋_hcS(一级序数) = 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_hw(一级序数) / 2 + 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_df(一级序数)
                                加强筋_IoS(一级序数) = 加强筋_Iow(一级序数) + 加强筋_Iof(一级序数)
                                加强筋_IcS(一级序数) = 加强筋_Icw(一级序数) + 加强筋_Aw(一级序数) * (加强筋_hw(一级序数) / 2 - 加强筋_hcS(一级序数)) ^ 2 + 加强筋_Icf(一级序数) + 加强筋_Af(一级序数) * (加强筋_df(一级序数) - 加强筋_hcS(一级序数)) ^ 2

                                Select Case 加强筋_tp(一级序数)
                                    Case "F"
                                        加强筋_IPS(一级序数) = 加强筋_hw(一级序数) ^ 3 * 加强筋_tw(一级序数) / 3
                                        加强筋_ITS(一级序数) = 加强筋_hw(一级序数) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / 加强筋_hw(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_hw(一级序数) ^ 3 * 加强筋_tw(一级序数) ^ 3 / 36
                                    Case "B"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                                    Case "T"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_wf(一级序数) ^ 3 * 加强筋_tf(一级序数) * 加强筋_df(一级序数) ^ 2 / 12
                                    Case "L1"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                                    Case "L2"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                                    Case "L3"
                                        加强筋_IPS(一级序数) = 加强筋_Aw(一级序数) * (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) ^ 2 / 3 + 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2
                                        加强筋_ITS(一级序数) = (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) * 加强筋_tw(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tw(一级序数) / (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2)) + 加强筋_wf(一级序数) * 加强筋_tf(一级序数) ^ 3 / 3 * (1 - 0.63 * 加强筋_tf(一级序数) / 加强筋_wf(一级序数))
                                        加强筋_IWS(一级序数) = 加强筋_Af(一级序数) * 加强筋_df(一级序数) ^ 2 * 加强筋_wf(一级序数) ^ 2 / 12 * (加强筋_Af(一级序数) + 2.6 * 加强筋_Aw(一级序数)) / (加强筋_Af(一级序数) + 加强筋_Aw(一级序数))
                                    Case Else

                                End Select
                                '加强筋_ηS(一级序数) = 1 + 加强筋_l(一级序数) ^ 2 / PI ^ 2 / Sqrt(加强筋_IWS(一级序数) * (0.75 * 加强筋带板_w(一级序数) / 加强筋带板_t(一级序数) ^ 3 + (加强筋_df(一级序数) - 加强筋_tf(一级序数) / 2) / 加强筋_tw(一级序数) ^ 3))
                                '加强筋_σETS(一级序数) = 标准弹性模量 / 加强筋_IPS(一级序数) * (加强筋_ηS(一级序数) * PI ^ 2 * 加强筋_IWS(一级序数) / 加强筋_lS(一级序数) ^ 2 + 0.385 * 加强筋_ITS(一级序数))
                                加强筋_σELS(一级序数) = 160000 * (加强筋_tw(一级序数) / 加强筋_hw(一级序数)) ^ 2

                                加强筋_Ycw(一级序数) = 加强筋_Y0(一级序数) + 加强筋_hw(一级序数) * Cos(加强筋_αw(一级序数)) / 2
                                加强筋_Zcw(一级序数) = 加强筋_Z0(一级序数) + 加强筋_hw(一级序数) * Sin(加强筋_αw(一级序数)) / 2

                                加强筋_Ycf(一级序数) = 加强筋_Y0(一级序数) + 加强筋_df(一级序数) * Cos(加强筋_αw(一级序数)) + 加强筋_wf(一级序数) * Cos(加强筋_αf(一级序数)) / 2
                                加强筋_Zcf(一级序数) = 加强筋_Z0(一级序数) + 加强筋_df(一级序数) * Sin(加强筋_αw(一级序数)) + 加强筋_wf(一级序数) * Sin(加强筋_αf(一级序数)) / 2

                                加强筋_YcS(一级序数) = 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_Ycf(一级序数) + 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_Ycw(一级序数)
                                加强筋_ZcS(一级序数) = 加强筋_Af(一级序数) / 加强筋_AS(一级序数) * 加强筋_Zcf(一级序数) + 加强筋_Aw(一级序数) / 加强筋_AS(一级序数) * 加强筋_Zcw(一级序数)
                            Next
                        Case 3
                            For 一级序数 As UShort = 1 To 面板_总数
                                BoxMuller(面板_tX(一级序数), 面板_t(一级序数)) 'BoxMuller(面板_σYX(一级序数), 面板_σY(一级序数))

                                面板_w(一级序数) = Sqrt((面板_Y0(一级序数) - 面板_Y1(一级序数)) ^ 2 + (面板_Z0(一级序数) - 面板_Z1(一级序数)) ^ 2)
                                面板_A(一级序数) = 面板_w(一级序数) * 面板_t(一级序数)

                                面板_Yc(一级序数) = (面板_Y0(一级序数) + 面板_Y1(一级序数)) / 2
                                面板_Zc(一级序数) = (面板_Z0(一级序数) + 面板_Z1(一级序数)) / 2
                            Next
                        Case 4
                            For 一级序数 As UShort = 1 To 板格_总数
                                板格_α(一级序数) = If(板格_Y0(一级序数) = 板格_Y1(一级序数), PI / 2, Atan((板格_Z0(一级序数) - 板格_Z1(一级序数)) / (板格_Y0(一级序数) - 板格_Y1(一级序数))))
                                板格_w(一级序数) = Sqrt((板格_Y0(一级序数) - 板格_Y1(一级序数)) ^ 2 + (板格_Z0(一级序数) - 板格_Z1(一级序数)) ^ 2)
                                For 二级序数 As UShort = 1 To 板格_板格板数目(一级序数)
                                    BoxMuller(板格板_tX(一级序数, 二级序数), 板格板_t(一级序数, 二级序数)) 'BoxMuller(板格板_σYX(一级序数, 二级序数), 板格板_σY(一级序数, 二级序数))

                                    板格板_Y0(一级序数, 二级序数) = If(二级序数 = 1, 板格_Y0(一级序数), 板格板_Y0(一级序数, 二级序数 - 1))
                                    板格板_Z0(一级序数, 二级序数) = If(二级序数 = 1, 板格_Z0(一级序数), 板格板_Z0(一级序数, 二级序数 - 1))

                                    板格板_Y1(一级序数, 二级序数) = If(二级序数 = 板格_板格板数目(一级序数), 板格_Y1(一级序数), 板格板_Y0(一级序数, 二级序数) + 板格板_w(一级序数, 二级序数) * Cos(板格_α(一级序数)))
                                    板格板_Z1(一级序数, 二级序数) = If(二级序数 = 板格_板格板数目(一级序数), 板格_Z1(一级序数), 板格板_Z0(一级序数, 二级序数) + 板格板_w(一级序数, 二级序数) * Sin(板格_α(一级序数)))

                                    板格板_Yc(一级序数, 二级序数) = (板格板_Y0(一级序数, 二级序数) + 板格板_Y1(一级序数, 二级序数)) / 2
                                    板格板_Zc(一级序数, 二级序数) = (板格板_Z0(一级序数, 二级序数) + 板格板_Z1(一级序数, 二级序数)) / 2

                                    板格板_A(一级序数, 二级序数) = 板格板_w(一级序数, 二级序数) * 板格板_t(一级序数, 二级序数)
                                    板格板_AYc(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_Yc(一级序数, 二级序数)
                                    板格板_AZc(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_Zc(一级序数, 二级序数)
                                    板格板_AσY(一级序数, 二级序数) = 板格板_A(一级序数, 二级序数) * 板格板_σY(一级序数, 二级序数)

                                    板格_A(一级序数) += 板格板_A(一级序数, 二级序数)
                                    板格_AYc(一级序数) += 板格板_AYc(一级序数, 二级序数)
                                    板格_AZc(一级序数) += 板格板_AZc(一级序数, 二级序数)
                                    板格_AσY(一级序数) += 板格板_AσY(一级序数, 二级序数)
                                Next
                                板格_t(一级序数) = 板格_A(一级序数) / 板格_w(一级序数)
                                板格_Yc(一级序数) = 板格_AYc(一级序数) / 板格_A(一级序数)
                                板格_Zc(一级序数) = 板格_AZc(一级序数) / 板格_A(一级序数)
                                板格_σY(一级序数) = 板格_AσY(一级序数) / 板格_A(一级序数)
                            Next
                        Case 5
                            ReDim 特别硬角单元_A(特别硬角单元_总数), 特别硬角单元_AYc(特别硬角单元_总数), 特别硬角单元_AZc(特别硬角单元_总数), 特别硬角单元_AσY(特别硬角单元_总数)
                            For 一级序数 As UShort = 1 To 特别硬角单元_总数
                                For 二级序数 As UShort = 1 To 子对象_数目(一级序数)
                                    BoxMuller(子对象_AX(一级序数, 二级序数), 子对象_A(一级序数, 二级序数)) 'BoxMuller(子对象_σYX(一级序数, 二级序数), 子对象_σY(一级序数, 二级序数))

                                    子对象_AYc(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_Yc(一级序数, 二级序数)
                                    子对象_AZc(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_Zc(一级序数, 二级序数)
                                    子对象_AσY(一级序数, 二级序数) = 子对象_A(一级序数, 二级序数) * 子对象_σY(一级序数, 二级序数)

                                    特别硬角单元_A(一级序数) += 子对象_A(一级序数, 二级序数)
                                    特别硬角单元_AYc(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_Yc(一级序数, 二级序数)
                                    特别硬角单元_AZc(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_Zc(一级序数, 二级序数)
                                    特别硬角单元_AσY(一级序数) += 子对象_A(一级序数, 二级序数) * 子对象_σY(一级序数, 二级序数)
                                Next
                                特别硬角单元_Yc(一级序数) = 特别硬角单元_AYc(一级序数) / 特别硬角单元_A(一级序数)
                                特别硬角单元_Zc(一级序数) = 特别硬角单元_AZc(一级序数) / 特别硬角单元_A(一级序数)
                                特别硬角单元_σY(一级序数) = 特别硬角单元_AσY(一级序数) / 特别硬角单元_A(一级序数)
                            Next
                    End Select
                Next
            End If

            For 单元划分步骤序数 As UShort = 1 To 4
                Select Case 单元划分步骤序数
                    Case 1      '板格-节点配对, 成立节点-分支
                        ReDim 节点_分支_数目(节点_总数)
                        ReDim 节点_分支_板格_序数(节点_总数, 通用分支数目)

                        ReDim 节点_分支_α(节点_总数, 通用分支数目)

                        ReDim 板格_首端节点_序数(板格_总数)
                        ReDim 板格_首端节点_分支_序数(板格_总数)
                        ReDim 板格_末端节点_序数(板格_总数)
                        ReDim 板格_末端节点_分支_序数(板格_总数)

                        ReDim 节点_tp(节点_总数)

                        ReDim 节点_加强筋_序数(节点_总数)
                        ReDim 加强筋_节点_序数(加强筋_总数)

                        ReDim 节点_面板_序数(节点_总数)
                        ReDim 面板_节点_序数(面板_总数)

                        ReDim 节点_硬角单元_序数(节点_总数)
                        'ReDim 硬角单元_节点_序数()

                        ReDim 节点_面板硬角单元_序数(节点_总数)
                        'ReDim 面板硬角单元_节点_序数()

                        ReDim 节点_加强筋单元_序数(节点_总数)
                        'ReDim 加强筋单元_节点_序数()

                        ReDim 节点_面板加强筋单元_序数(节点_总数)
                        'ReDim 面板加强筋单元_节点_序数()

                        ReDim 节点_分支_首端节点_序数(节点_总数, 通用分支数目), 节点_分支_末端节点_序数(节点_总数, 通用分支数目)

                        ReDim 板格_首端_w(板格_总数), 板格_末端_w(板格_总数)

                        For 一级序数 As UShort = 1 To 板格_总数
                            For 二级序数 As UShort = 1 To 节点_总数
                                If 板格_Y0(一级序数) = 节点_Y0(二级序数) And 板格_Z0(一级序数) = 节点_Z0(二级序数) Then
                                    节点_分支_数目(二级序数) += 1
                                    节点_分支_板格_序数(二级序数, 节点_分支_数目(二级序数)) = 一级序数

                                    节点_分支_α(二级序数, 节点_分支_数目(二级序数)) = 板格_α(一级序数)

                                    板格_首端节点_序数(一级序数) = 二级序数
                                    板格_首端节点_分支_序数(一级序数) = 节点_分支_数目(二级序数)
                                ElseIf 板格_Y1(一级序数) = 节点_Y0(二级序数) And 板格_Z1(一级序数) = 节点_Z0(二级序数) Then
                                    节点_分支_数目(二级序数) += 1
                                    节点_分支_板格_序数(二级序数, 节点_分支_数目(二级序数)) = 一级序数

                                    节点_分支_α(二级序数, 节点_分支_数目(二级序数)) = If(板格_α(一级序数) <= 0, 板格_α(一级序数) + PI, 板格_α(一级序数) - PI)

                                    板格_末端节点_序数(一级序数) = 二级序数
                                    板格_末端节点_分支_序数(一级序数) = 节点_分支_数目(二级序数)
                                End If
                                If (Not 板格_首端节点_序数(一级序数) = 0) And (Not 板格_末端节点_序数(一级序数) = 0) Then
                                    '节点_分支_首端节点_序数(二级序数) = 二级序数
                                    节点_分支_首端节点_序数(板格_首端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_首端节点_序数(一级序数)
                                    节点_分支_末端节点_序数(板格_首端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_末端节点_序数(一级序数)
                                    节点_分支_首端节点_序数(板格_末端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_末端节点_序数(一级序数)
                                    节点_分支_末端节点_序数(板格_末端节点_序数(一级序数), 节点_分支_数目(板格_首端节点_序数(一级序数))) = 板格_首端节点_序数(一级序数)
                                    '节点_分支_末端节点_序数(二级序数) = 二级序数
                                    Exit For
                                End If
                            Next
                        Next
                    Case 2      '节点-加强筋/面板配对
                        For 一级序数 As UShort = 1 To 节点_总数
                            For 二级序数 As UShort = 1 To 加强筋_总数
                                If 节点_Y0(一级序数) = 加强筋_Y0(二级序数) And 节点_Z0(一级序数) = 加强筋_Z0(二级序数) Then
                                    节点_tp(一级序数) = "加强筋"

                                    节点_加强筋_序数(一级序数) = 二级序数
                                    加强筋_节点_序数(二级序数) = 一级序数

                                    '属性继承：加强筋 → 节点_加强筋
                                    Exit For
                                End If
                            Next

                            For 二级序数 As UShort = 1 To 面板_总数
                                If 节点_Y0(一级序数) = 面板_YL(二级序数) And 节点_Z0(一级序数) = 面板_ZL(二级序数) Then
                                    节点_tp(一级序数) = "面板"

                                    节点_面板_序数(一级序数) = 二级序数
                                    面板_节点_序数(二级序数) = 一级序数

                                    '属性继承：面板 → 节点_面板
                                    Exit For
                                End If
                            Next
                        Next
                    Case 3      '根据节点分支数目确定单元类型
                        硬角单元_总数 = 0
                        加强筋单元_总数 = 0
                        加筋板单元_总数 = 0

                        For 一级序数 As UShort = 1 To 节点_总数
                            Select Case 节点_分支_数目(一级序数)
                                Case >= 3
                                    节点_tp(一级序数) = "硬角单元"
                                    硬角单元_总数 += 1

                                    节点_硬角单元_序数(一级序数) = 硬角单元_总数
                                    ReDim Preserve 硬角单元_节点_序数(硬角单元_总数)
                                    硬角单元_节点_序数(硬角单元_总数) = 一级序数

                                    Exit Select
                                Case 2
                                    Dim 双分支夹角 As Single = 节点_分支_α(一级序数, 1) - 节点_分支_α(一级序数, 2)
                                    Select Case 双分支夹角
                                        Case >= 2 * PI
                                            双分支夹角 -= 2 * PI
                                        Case < 0
                                            双分支夹角 += 2 * PI
                                        Case Else

                                    End Select
                                    Select Case 双分支夹角
                                        Case <= 5 / 6 * PI
                                            节点_tp(一级序数) = "硬角单元"
                                            硬角单元_总数 += 1

                                            节点_硬角单元_序数(一级序数) = 硬角单元_总数
                                            ReDim Preserve 硬角单元_节点_序数(硬角单元_总数)
                                            硬角单元_节点_序数(硬角单元_总数) = 一级序数

                                            Exit Select
                                        Case >= 7 / 6 * PI
                                            节点_tp(一级序数) = "硬角单元"
                                            硬角单元_总数 += 1

                                            节点_硬角单元_序数(一级序数) = 硬角单元_总数
                                            ReDim Preserve 硬角单元_节点_序数(硬角单元_总数)
                                            硬角单元_节点_序数(硬角单元_总数) = 一级序数

                                            Exit Select
                                        Case Else
                                            Select Case 节点_tp(一级序数)
                                                Case "加强筋"
                                                    节点_tp(一级序数) = "加强筋单元"
                                                    加强筋单元_总数 += 1

                                                    节点_加强筋单元_序数(一级序数) = 加强筋单元_总数
                                                    ReDim Preserve 加强筋单元_节点_序数(加强筋单元_总数)
                                                    加强筋单元_节点_序数(加强筋单元_总数) = 一级序数

                                                    Exit Select
                                                Case Else

                                            End Select
                                    End Select
                                Case 1
                                    Select Case 节点_tp(一级序数)
                                        Case "加强筋"
                                            节点_tp(一级序数) = "加强筋单元"
                                            加强筋单元_总数 += 1

                                            节点_加强筋单元_序数(一级序数) = 加强筋单元_总数
                                            ReDim Preserve 加强筋单元_节点_序数(加强筋单元_总数)
                                            加强筋单元_节点_序数(加强筋单元_总数) = 一级序数

                                            Exit Select
                                        Case "面板"
                                            Select Case 面板_PMA(节点_面板_序数(一级序数))
                                                Case True
                                                    节点_tp(一级序数) = "面板加强筋单元"
                                                    面板加强筋单元_总数 += 1

                                                    节点_面板加强筋单元_序数(一级序数) = 面板加强筋单元_总数
                                                    ReDim Preserve 面板加强筋单元_节点_序数(面板加强筋单元_总数)
                                                    面板加强筋单元_节点_序数(面板加强筋单元_总数) = 一级序数

                                                    Exit Select
                                                Case False
                                                    节点_tp(一级序数) = "面板硬角单元"
                                                    面板硬角单元_总数 += 1

                                                    节点_面板硬角单元_序数(一级序数) = 面板硬角单元_总数
                                                    ReDim Preserve 面板硬角单元_节点_序数(面板硬角单元_总数)
                                                    面板硬角单元_节点_序数(面板硬角单元_总数) = 一级序数

                                                    Exit Select
                                            End Select
                                        Case Else
                                            节点_tp(一级序数) = "自由端"

                                            Exit Select
                                    End Select
                            End Select
                        Next
                    Case 4      '确定单元属性
                        全截面_A = 0
                        全截面_AYc = 0
                        全截面_AZc = 0
                        全截面_AσY = 0
                        ReDim 硬角单元_A(硬角单元_总数), 硬角单元_Yc(硬角单元_总数), 硬角单元_Zc(硬角单元_总数), 硬角单元_σY(硬角单元_总数)
                        ReDim 加强筋单元_A(加强筋单元_总数), 加强筋单元_Yc(加强筋单元_总数), 加强筋单元_Zc(加强筋单元_总数), 加强筋单元_σY(加强筋单元_总数)

                        For 单元对象类型序数 As UShort = 1 To 6
                            Dim 通用单元数目 As UShort = 加强筋单元_总数
                            Dim 单元分支_w(通用单元数目, 通用分支数目) As Single,
                            单元分支_A(通用单元数目, 通用分支数目) As Single,
                            单元分支_AYc(通用单元数目, 通用分支数目) As Single, 单元分支_AZc(通用单元数目, 通用分支数目) As Single,
                            单元分支_AσY(通用单元数目, 通用分支数目) As Single

                            Select Case 单元对象类型序数
                                Case 1      '硬角单元
                                    For 一级序数 As UShort = 1 To 硬角单元_总数
                                        Dim 硬角单元_AYc(硬角单元_总数) As Single, 硬角单元_AZc(硬角单元_总数) As Single,
                                        硬角单元_AσY(硬角单元_总数) As Single

                                        Dim 关联节点序数 As UShort = 硬角单元_节点_序数(一级序数)
                                        For 二级序数 As UShort = 1 To 节点_分支_数目(关联节点序数)
                                            Dim 关联板格序数 As UShort = 节点_分支_板格_序数(关联节点序数, 二级序数)

                                            Select Case 板格_tp(关联板格序数)
                                                Case "L"
                                                    单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数) / 2
                                                    If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 二级序数)) = "自由端" Then
                                                        单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数)
                                                    End If
                                                Case "T"
                                                    单元分支_w(一级序数, 二级序数) = 板格_t(关联板格序数) * 20
                                                Case Else

                                            End Select
                                            Dim 分支板_A(硬角单元_总数, 通用分支数目) As Single,
                                            分支板_AYc(硬角单元_总数, 通用分支数目) As Single, 分支板_AZc(硬角单元_总数, 通用分支数目) As Single,
                                            分支板_AσY(硬角单元_总数, 通用分支数目) As Single
                                            For 三级序数 As UShort = 1 To 板格_板格板数目(关联板格序数)
                                                If 关联节点序数 = 板格_首端节点_序数(关联板格序数) Then
                                                    板格_首端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                    Dim 分支板_w(硬角单元_总数, 通用分支数目) As Single
                                                    分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 三级序数)
                                                    Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                    Select Case 分支板超出宽度
                                                        Case <= 0   '分支板宽度和小于分支所需宽度
                                                            分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 三级序数)
                                                            分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 三级序数)
                                                            分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 三级序数)
                                                            分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 三级序数)
                                                        Case > 0    '分支板宽度和大于分支所需宽度
                                                            Dim 通用板格板数目 As UShort = 3
                                                            Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                            '板格板实际取用宽度
                                                            板格板_w1(关联板格序数, 三级序数) = 板格板_w(关联板格序数, 三级序数) - 分支板超出宽度
                                                            '板格板宽度利用系数
                                                            Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 三级序数) / 板格板_w(关联板格序数, 三级序数)
                                                            Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y11(板格_总数, 通用板格板数目), 板格板_Z11(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                            '板格板实际取用面积
                                                            板格板_A1(关联板格序数, 三级序数) = 板格板_A(关联板格序数, 三级序数) * 板格板_ηw
                                                            '板格板实际末端坐标
                                                            板格板_Y11(关联板格序数, 三级序数) = 板格板_Y0(关联板格序数, 三级序数) + (板格板_Y1(关联板格序数, 三级序数) - 板格板_Y0(关联板格序数, 三级序数)) * 板格板_ηw
                                                            板格板_Z11(关联板格序数, 三级序数) = 板格板_Z0(关联板格序数, 三级序数) + (板格板_Z1(关联板格序数, 三级序数) - 板格板_Z0(关联板格序数, 三级序数)) * 板格板_ηw
                                                            '板格板实际形心坐标
                                                            板格板_Yc1(关联板格序数, 三级序数) = (板格板_Y0(关联板格序数, 三级序数) + 板格板_Y11(关联板格序数, 三级序数)） / 2
                                                            板格板_Zc1(关联板格序数, 三级序数) = (板格板_Z0(关联板格序数, 三级序数) + 板格板_Z11(关联板格序数, 三级序数)） / 2
                                                            '板格板实际面积坐标积数
                                                            板格板_AYc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Yc1(关联板格序数, 三级序数)
                                                            板格板_AZc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Zc1(关联板格序数, 三级序数)
                                                            '板格板实际面积强度积数
                                                            板格板_AσY1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_σY(关联板格序数, 三级序数)

                                                            'Select Case 三级序数
                                                            '    Case = 1

                                                            '    Case > 1

                                                            'End Select

                                                            '分支板实际面积
                                                            分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 三级序数)
                                                            '分支板实际面积坐标积数
                                                            分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 三级序数)
                                                            分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 三级序数)
                                                            '分支板实际面积强度积数
                                                            分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 三级序数)

                                                            Exit For
                                                    End Select
                                                ElseIf 关联节点序数 = 板格_末端节点_序数(关联板格序数) Then
                                                    板格_末端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                    Dim 分支板_w(硬角单元_总数, 通用分支数目) As Single
                                                    Dim 反向序数 As UShort = 板格_板格板数目(关联板格序数) + 1 - 三级序数
                                                    分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 反向序数)
                                                    Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                    Select Case 分支板超出宽度
                                                        Case <= 0   '分支板宽度和小于分支所需宽度
                                                            分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 反向序数)
                                                            分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 反向序数)
                                                            分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 反向序数)
                                                            分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 反向序数)
                                                        Case > 0    '分支板宽度和大于分支所需宽度
                                                            Dim 通用板格板数目 As UShort = 3
                                                            Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                            '板格板实际取用宽度
                                                            板格板_w1(关联板格序数, 反向序数) = 板格板_w(关联板格序数, 反向序数) - 分支板超出宽度
                                                            '板格板宽度利用系数
                                                            Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 反向序数) / 板格板_w(关联板格序数, 反向序数)
                                                            Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y01(板格_总数, 通用板格板数目), 板格板_Z01(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                            '板格板实际取用面积
                                                            板格板_A1(关联板格序数, 反向序数) = 板格板_A(关联板格序数, 反向序数) * 板格板_ηw
                                                            '板格板实际末端坐标
                                                            板格板_Y01(关联板格序数, 反向序数) = 板格板_Y1(关联板格序数, 反向序数) + (板格板_Y0(关联板格序数, 反向序数) - 板格板_Y1(关联板格序数, 反向序数)) * 板格板_ηw
                                                            板格板_Z01(关联板格序数, 反向序数) = 板格板_Z1(关联板格序数, 反向序数) + (板格板_Z0(关联板格序数, 反向序数) - 板格板_Z1(关联板格序数, 反向序数)) * 板格板_ηw
                                                            '板格板实际形心坐标
                                                            板格板_Yc1(关联板格序数, 反向序数) = (板格板_Y01(关联板格序数, 反向序数) + 板格板_Y1(关联板格序数, 反向序数)） / 2
                                                            板格板_Zc1(关联板格序数, 反向序数) = (板格板_Z01(关联板格序数, 反向序数) + 板格板_Z1(关联板格序数, 反向序数)） / 2
                                                            '板格板实际面积坐标积数
                                                            板格板_AYc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Yc1(关联板格序数, 反向序数)
                                                            板格板_AZc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Zc1(关联板格序数, 反向序数)
                                                            '板格板实际面积强度积数
                                                            板格板_AσY1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_σY(关联板格序数, 反向序数)

                                                            'Select Case 三级序数
                                                            '    Case = 1

                                                            '    Case > 1

                                                            'End Select

                                                            '分支板实际面积
                                                            分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 反向序数)
                                                            '分支板实际面积坐标积数
                                                            分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 反向序数)
                                                            分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 反向序数)
                                                            '分支板实际面积强度积数
                                                            分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 反向序数)

                                                            Exit For
                                                    End Select
                                                End If
                                            Next
                                            单元分支_A(一级序数, 二级序数) = 分支板_A(一级序数, 二级序数)
                                            单元分支_AYc(一级序数, 二级序数) = 分支板_AYc(一级序数, 二级序数)
                                            单元分支_AZc(一级序数, 二级序数) = 分支板_AZc(一级序数, 二级序数)
                                            单元分支_AσY(一级序数, 二级序数) = 分支板_AσY(一级序数, 二级序数)

                                            硬角单元_A(一级序数) += 分支板_A(一级序数, 二级序数)
                                            硬角单元_AYc(一级序数) += 分支板_AYc(一级序数, 二级序数)
                                            硬角单元_AZc(一级序数) += 分支板_AZc(一级序数, 二级序数)
                                            硬角单元_AσY(一级序数) += 分支板_AσY(一级序数, 二级序数)
                                        Next

                                        硬角单元_Yc(一级序数) = 硬角单元_AYc(一级序数) / 硬角单元_A(一级序数)
                                        硬角单元_Zc(一级序数) = 硬角单元_AZc(一级序数) / 硬角单元_A(一级序数)
                                        硬角单元_σY(一级序数) = 硬角单元_AσY(一级序数) / 硬角单元_A(一级序数)

                                        全截面_A += 硬角单元_A(一级序数)
                                        全截面_AYc += 硬角单元_AYc(一级序数)
                                        全截面_AZc += 硬角单元_AZc(一级序数)
                                        全截面_AσY += 硬角单元_AσY(一级序数)
                                    Next
                                Case 2      '面板硬角单元
                                'MsgBox("面板硬角单元部分未完成!")
                                Case 3      '特别硬角单元
                                    For 一级序数 As UShort = 1 To 特别硬角单元_总数
                                        全截面_A += 特别硬角单元_A(一级序数)
                                        全截面_AYc += 特别硬角单元_AYc(一级序数)
                                        全截面_AZc += 特别硬角单元_AZc(一级序数)
                                        全截面_AσY += 特别硬角单元_AσY(一级序数)
                                    Next
                                    'Debug.Print(Format(全截面_AYc / 全截面_A / 1000, "000.000") & " - " & Format(全截面_AZc / 全截面_A / 1000, "000.000") & " - " & Format(全截面_A / 1000000, "000.000"))
                                Case 4      '加强筋单元
                                    ReDim 加强筋单元_σYP(加强筋单元_总数),
                                    加强筋单元_lP(加强筋单元_总数),
                                    加强筋单元_wP(加强筋单元_总数), 加强筋单元_tP(加强筋单元_总数),
                                    加强筋单元_σYS(加强筋单元_总数),
                                    加强筋单元_lS(加强筋单元_总数),
                                    加强筋单元_hw(加强筋单元_总数), 加强筋单元_tw(加强筋单元_总数),
                                    加强筋单元_wf(加强筋单元_总数), 加强筋单元_tf(加强筋单元_总数),
                                    加强筋单元_dx(加强筋单元_总数),
                                    加强筋单元_tpS(加强筋单元_总数), 加强筋单元_mk(加强筋单元_总数)

                                    ReDim 加强筋单元_ηS(加强筋单元_总数), 加强筋单元_σETS(加强筋单元_总数)

                                    ReDim 加强筋单元_σELS(加强筋单元_总数)

                                    Dim 加强筋单元_AYc(加强筋单元_总数) As Single, 加强筋单元_AZc(加强筋单元_总数) As Single,
                                    加强筋单元_AσY(加强筋单元_总数) As Single

                                    For 一级序数 As UShort = 1 To 加强筋单元_总数
                                        Dim 关联节点序数 As UShort = 加强筋单元_节点_序数(一级序数)
                                        Select Case 节点_分支_数目(关联节点序数)
                                            Case 1
                                                Select Case 板格_tp(节点_分支_板格_序数(关联节点序数, 1))
                                                    Case "L"
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                        If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 1)) = "自由端" Then
                                                            单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1))
                                                        End If
                                                        加强筋单元_lP(一级序数) = 板格_l(节点_分支_板格_序数(关联节点序数, 1))
                                                        加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1)
                                                    'MsgBox("警告:单L分支")
                                                    Case "T"
                                                        'MsgBox("错误:单T分支")
                                                    Case Else

                                                End Select
                                            Case 2
                                                Dim 板格_联合_tp As String = 板格_tp(节点_分支_板格_序数(关联节点序数, 1)) & 板格_tp(节点_分支_板格_序数(关联节点序数, 2))
                                                Select Case 板格_联合_tp
                                                    Case "LL"
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                        单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2)) / 2
                                                        If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 1)) = "自由端" Then
                                                            单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1))
                                                        End If
                                                        If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 2)) = "自由端" Then
                                                            单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2))
                                                        End If
                                                        加强筋单元_lP(一级序数) = (板格_l(节点_分支_板格_序数(关联节点序数, 1)) + 板格_l(节点_分支_板格_序数(关联节点序数, 2))) / 2
                                                        加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1) + 单元分支_w(一级序数, 2)
                                                    Case "LT"
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                        If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 1)) = "自由端" Then
                                                            单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 1))
                                                        End If
                                                        单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 1)) / 2
                                                        加强筋单元_lP(一级序数) = (板格_l(节点_分支_板格_序数(关联节点序数, 1)) + 板格_l(节点_分支_板格_序数(关联节点序数, 2))) / 2
                                                        加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1) + 单元分支_w(一级序数, 2)
                                                        'MsgBox("警告:单L单T分支")
                                                    Case "TL"
                                                        单元分支_w(一级序数, 1) = 板格_w(节点_分支_板格_序数(关联节点序数, 2)) / 2
                                                        单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2)) / 2
                                                        If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 2)) = "自由端" Then
                                                            单元分支_w(一级序数, 2) = 板格_w(节点_分支_板格_序数(关联节点序数, 2))
                                                        End If
                                                        加强筋单元_lP(一级序数) = (板格_l(节点_分支_板格_序数(关联节点序数, 1)) + 板格_l(节点_分支_板格_序数(关联节点序数, 2))) / 2
                                                        加强筋单元_wP(一级序数) = 单元分支_w(一级序数, 1) + 单元分支_w(一级序数, 2)
                                                        'MsgBox("警告:单T单L分支")
                                                    Case "TT"
                                                        MsgBox("错误:双T分支")
                                                    Case Else

                                                End Select
                                        End Select

                                        Dim 关联加强筋序数 As UShort = 节点_加强筋_序数(关联节点序数)
                                        加强筋单元_A(一级序数) += 加强筋_AS(关联加强筋序数)
                                        加强筋单元_AYc(一级序数) += 加强筋_AS(关联加强筋序数) * 加强筋_YcS(关联加强筋序数)
                                        加强筋单元_AZc(一级序数) += 加强筋_AS(关联加强筋序数) * 加强筋_ZcS(关联加强筋序数)
                                        加强筋单元_AσY(一级序数) += 加强筋_AS(关联加强筋序数) * 加强筋_σY(关联加强筋序数)

                                        加强筋单元_σYS(一级序数) = 加强筋_σY(关联加强筋序数)

                                        加强筋单元_lS(一级序数) = 加强筋_l(关联加强筋序数)

                                        加强筋单元_hw(一级序数) = 加强筋_hw(关联加强筋序数)
                                        加强筋单元_tw(一级序数) = 加强筋_tw(关联加强筋序数)

                                        加强筋单元_wf(一级序数) = 加强筋_wf(关联加强筋序数)
                                        加强筋单元_tf(一级序数) = 加强筋_tf(关联加强筋序数)

                                        加强筋单元_dx(一级序数) = 加强筋_dx(关联加强筋序数)

                                        加强筋单元_tpS(一级序数) = 加强筋_tp(关联加强筋序数)
                                        加强筋单元_mk(一级序数) = 加强筋_mk(关联加强筋序数)

                                        加强筋单元_σELS(一级序数) = 加强筋_σELS(关联加强筋序数)

                                        For 二级序数 As UShort = 1 To 节点_分支_数目(关联节点序数)
                                            Dim 关联板格序数 As UShort = 节点_分支_板格_序数(关联节点序数, 二级序数)

                                            'Select Case 板格_tp(关联板格序数)
                                            '    Case "L"
                                            '        单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数) / 2
                                            '        If 节点_tp(节点_分支_末端节点_序数(关联节点序数, 二级序数)) = "自由端" Then
                                            '            单元分支_w(一级序数, 二级序数) = 板格_w(关联板格序数)
                                            '        End If
                                            '    Case "T"
                                            '        单元分支_w(一级序数, 二级序数) = 板格_t(关联板格序数) * 20
                                            'End Select
                                            Dim 分支板_A(加强筋单元_总数, 通用分支数目) As Single,
                                            分支板_AYc(加强筋单元_总数, 通用分支数目) As Single, 分支板_AZc(加强筋单元_总数, 通用分支数目) As Single,
                                            分支板_AσY(加强筋单元_总数, 通用分支数目) As Single
                                            For 三级序数 As UShort = 1 To 板格_板格板数目(关联板格序数)
                                                If 关联节点序数 = 板格_首端节点_序数(关联板格序数) Then
                                                    板格_首端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                    Dim 分支板_w(加强筋单元_总数, 通用分支数目) As Single
                                                    分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 三级序数)
                                                    Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                    Select Case 分支板超出宽度
                                                        Case <= 0   '分支板宽度和小于分支所需宽度
                                                            分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 三级序数)
                                                            分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 三级序数)
                                                            分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 三级序数)
                                                            分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 三级序数)
                                                        Case > 0    '分支板宽度和大于分支所需宽度
                                                            Dim 通用板格板数目 As UShort = 3
                                                            Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                            '板格板实际取用宽度
                                                            板格板_w1(关联板格序数, 三级序数) = 板格板_w(关联板格序数, 三级序数) - 分支板超出宽度
                                                            '板格板宽度利用系数
                                                            Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 三级序数) / 板格板_w(关联板格序数, 三级序数)
                                                            Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y11(板格_总数, 通用板格板数目), 板格板_Z11(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                            '板格板实际取用面积
                                                            板格板_A1(关联板格序数, 三级序数) = 板格板_A(关联板格序数, 三级序数) * 板格板_ηw
                                                            '板格板实际末端坐标
                                                            板格板_Y11(关联板格序数, 三级序数) = 板格板_Y0(关联板格序数, 三级序数) + (板格板_Y1(关联板格序数, 三级序数) - 板格板_Y0(关联板格序数, 三级序数)) * 板格板_ηw
                                                            板格板_Z11(关联板格序数, 三级序数) = 板格板_Z0(关联板格序数, 三级序数) + (板格板_Z1(关联板格序数, 三级序数) - 板格板_Z0(关联板格序数, 三级序数)) * 板格板_ηw
                                                            '板格板实际形心坐标
                                                            板格板_Yc1(关联板格序数, 三级序数) = (板格板_Y0(关联板格序数, 三级序数) + 板格板_Y11(关联板格序数, 三级序数)） / 2
                                                            板格板_Zc1(关联板格序数, 三级序数) = (板格板_Z0(关联板格序数, 三级序数) + 板格板_Z11(关联板格序数, 三级序数)） / 2
                                                            '板格板实际面积坐标积数
                                                            板格板_AYc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Yc1(关联板格序数, 三级序数)
                                                            板格板_AZc1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_Zc1(关联板格序数, 三级序数)
                                                            '板格板实际面积强度积数
                                                            板格板_AσY1(关联板格序数, 三级序数) = 板格板_A1(关联板格序数, 三级序数) * 板格板_σY(关联板格序数, 三级序数)

                                                            'Select Case 三级序数
                                                            '    Case = 1

                                                            '    Case > 1

                                                            'End Select

                                                            '分支板实际面积
                                                            分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 三级序数)
                                                            '分支板实际面积坐标积数
                                                            分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 三级序数)
                                                            分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 三级序数)
                                                            '分支板实际面积强度积数
                                                            分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 三级序数)

                                                            Exit For
                                                    End Select
                                                ElseIf 关联节点序数 = 板格_末端节点_序数(关联板格序数) Then
                                                    板格_末端_w(关联板格序数) = 单元分支_w(一级序数, 二级序数)
                                                    Dim 分支板_w(加强筋单元_总数, 通用分支数目) As Single
                                                    Dim 反向序数 As UShort = 板格_板格板数目(关联板格序数) + 1 - 三级序数
                                                    分支板_w(一级序数, 二级序数) += 板格板_w(关联板格序数, 反向序数)
                                                    Dim 分支板超出宽度 As Single = 分支板_w(一级序数, 二级序数) - 单元分支_w(一级序数, 二级序数)
                                                    Select Case 分支板超出宽度
                                                        Case <= 0   '分支板宽度和小于分支所需宽度
                                                            分支板_A(一级序数, 二级序数) += 板格板_A(关联板格序数, 反向序数)
                                                            分支板_AYc(一级序数, 二级序数) += 板格板_AYc(关联板格序数, 反向序数)
                                                            分支板_AZc(一级序数, 二级序数) += 板格板_AZc(关联板格序数, 反向序数)
                                                            分支板_AσY(一级序数, 二级序数) += 板格板_AσY(关联板格序数, 反向序数)
                                                        Case > 0    '分支板宽度和大于分支所需宽度
                                                            Dim 通用板格板数目 As UShort = 3
                                                            Dim 板格板_w1(板格_总数, 通用板格板数目) As Single
                                                            '板格板实际取用宽度
                                                            板格板_w1(关联板格序数, 反向序数) = 板格板_w(关联板格序数, 反向序数) - 分支板超出宽度
                                                            '板格板宽度利用系数
                                                            Dim 板格板_ηw As Single = 板格板_w1(关联板格序数, 反向序数) / 板格板_w(关联板格序数, 反向序数)
                                                            Dim 板格板_A1(板格_总数, 通用板格板数目),
                                                            板格板_Y01(板格_总数, 通用板格板数目), 板格板_Z01(板格_总数, 通用板格板数目),
                                                            板格板_Yc1(板格_总数, 通用板格板数目), 板格板_Zc1(板格_总数, 通用板格板数目),
                                                            板格板_AYc1(板格_总数, 通用板格板数目), 板格板_AZc1(板格_总数, 通用板格板数目),
                                                            板格板_AσY1(板格_总数, 通用板格板数目)
                                                            '板格板实际取用面积
                                                            板格板_A1(关联板格序数, 反向序数) = 板格板_A(关联板格序数, 反向序数) * 板格板_ηw
                                                            '板格板实际末端坐标
                                                            板格板_Y01(关联板格序数, 反向序数) = 板格板_Y1(关联板格序数, 反向序数) + (板格板_Y0(关联板格序数, 反向序数) - 板格板_Y1(关联板格序数, 反向序数)) * 板格板_ηw
                                                            板格板_Z01(关联板格序数, 反向序数) = 板格板_Z1(关联板格序数, 反向序数) + (板格板_Z0(关联板格序数, 反向序数) - 板格板_Z1(关联板格序数, 反向序数)) * 板格板_ηw
                                                            '板格板实际形心坐标
                                                            板格板_Yc1(关联板格序数, 反向序数) = (板格板_Y01(关联板格序数, 反向序数) + 板格板_Y1(关联板格序数, 反向序数)） / 2
                                                            板格板_Zc1(关联板格序数, 反向序数) = (板格板_Z01(关联板格序数, 反向序数) + 板格板_Z1(关联板格序数, 反向序数)） / 2
                                                            '板格板实际面积坐标积数
                                                            板格板_AYc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Yc1(关联板格序数, 反向序数)
                                                            板格板_AZc1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_Zc1(关联板格序数, 反向序数)
                                                            '板格板实际面积强度积数
                                                            板格板_AσY1(关联板格序数, 反向序数) = 板格板_A1(关联板格序数, 反向序数) * 板格板_σY(关联板格序数, 反向序数)

                                                            'Select Case 三级序数
                                                            '    Case = 1

                                                            '    Case > 1

                                                            'End Select

                                                            '分支板实际面积
                                                            分支板_A(一级序数, 二级序数) += 板格板_A1(关联板格序数, 反向序数)
                                                            '分支板实际面积坐标积数
                                                            分支板_AYc(一级序数, 二级序数) += 板格板_AYc1(关联板格序数, 反向序数)
                                                            分支板_AZc(一级序数, 二级序数) += 板格板_AZc1(关联板格序数, 反向序数)
                                                            '分支板实际面积强度积数
                                                            分支板_AσY(一级序数, 二级序数) += 板格板_AσY1(关联板格序数, 反向序数)

                                                            Exit For
                                                    End Select
                                                End If
                                            Next
                                            单元分支_A(一级序数, 二级序数) = 分支板_A(一级序数, 二级序数)
                                            单元分支_AYc(一级序数, 二级序数) = 分支板_AYc(一级序数, 二级序数)
                                            单元分支_AZc(一级序数, 二级序数) = 分支板_AZc(一级序数, 二级序数)
                                            单元分支_AσY(一级序数, 二级序数) = 分支板_AσY(一级序数, 二级序数)

                                            加强筋单元_A(一级序数) += 分支板_A(一级序数, 二级序数)
                                            加强筋单元_AYc(一级序数) += 分支板_AYc(一级序数, 二级序数)
                                            加强筋单元_AZc(一级序数) += 分支板_AZc(一级序数, 二级序数)
                                            加强筋单元_AσY(一级序数) += 分支板_AσY(一级序数, 二级序数)
                                        Next

                                        加强筋单元_tP(一级序数) = (加强筋单元_A(一级序数) - 加强筋_AS(关联加强筋序数)) / 加强筋单元_wP(一级序数)
                                        加强筋单元_σYP(一级序数) = (加强筋单元_AσY(一级序数) - 加强筋_AS(关联加强筋序数) * 加强筋_σY(关联加强筋序数)) / (加强筋单元_A(一级序数) - 加强筋_AS(关联加强筋序数))

                                        加强筋单元_ηS(一级序数) = 1 + 加强筋_l(关联加强筋序数) ^ 2 / PI ^ 2 / Sqrt(加强筋_IWS(关联加强筋序数) * (0.75 * 加强筋单元_wP(一级序数) / 加强筋单元_tP(一级序数) ^ 3 + (加强筋_df(关联加强筋序数) - 加强筋_tf(关联加强筋序数) / 2) / 加强筋_tw(关联加强筋序数) ^ 3))
                                        加强筋单元_σETS(一级序数) = 标准弹性模量 / 加强筋_IPS(关联加强筋序数) * (加强筋单元_ηS(一级序数) * PI ^ 2 * 加强筋_IWS(关联加强筋序数) / 加强筋单元_lS(一级序数) ^ 2 + 0.385 * 加强筋_ITS(关联加强筋序数))

                                        加强筋单元_Yc(一级序数) = 加强筋单元_AYc(一级序数) / 加强筋单元_A(一级序数)
                                        加强筋单元_Zc(一级序数) = 加强筋单元_AZc(一级序数) / 加强筋单元_A(一级序数)
                                        加强筋单元_σY(一级序数) = 加强筋单元_AσY(一级序数) / 加强筋单元_A(一级序数)

                                        全截面_A += 加强筋单元_A(一级序数)
                                        全截面_AYc += 加强筋单元_AYc(一级序数)
                                        全截面_AZc += 加强筋单元_AZc(一级序数)
                                        全截面_AσY += 加强筋单元_AσY(一级序数)
                                    Next
                                Case 5      '面板加强筋单元
                                'MsgBox("面板加强筋单元部分未完成!")
                                Case 6      '加筋板单元
                                    For 一级序数 As UShort = 1 To 板格_总数
                                        Dim 板格_原始_w As Single, 板格_剩余_w As Single
                                        Dim 板格_原始_A As Single, 板格_首端_A As Single, 板格_末端_A As Single, 板格_剩余_A As Single
                                        Dim 板格_原始_AYc As Single, 板格_首端_AYc As Single, 板格_末端_AYc As Single, 板格_剩余_AYc As Single
                                        Dim 板格_原始_AZc As Single, 板格_首端_AZc As Single, 板格_末端_AZc As Single, 板格_剩余_AZc As Single
                                        Dim 板格_原始_AσY As Single, 板格_首端_AσY As Single, 板格_末端_AσY As Single, 板格_剩余_AσY As Single

                                        Dim 首端关联节点序数 As UShort, 首端关联单元序数 As UShort, 首端关联分支序数 As UShort
                                        Dim 末端关联节点序数 As UShort, 末端关联单元序数 As UShort, 末端关联分支序数 As UShort

                                        首端关联节点序数 = 板格_首端节点_序数(一级序数)
                                        Select Case 节点_tp(首端关联节点序数)
                                            Case "硬角单元"
                                                首端关联单元序数 = 节点_硬角单元_序数(首端关联节点序数)
                                            Case "面板硬角单元"
                                                首端关联单元序数 = 节点_面板硬角单元_序数(首端关联节点序数)
                                            Case "加强筋单元"
                                                首端关联单元序数 = 节点_加强筋单元_序数(首端关联节点序数)
                                            Case "面板硬角单元"
                                                首端关联单元序数 = 节点_面板硬角单元_序数(首端关联节点序数)
                                            Case Else

                                        End Select
                                        首端关联分支序数 = 板格_首端节点_分支_序数(一级序数)

                                        末端关联节点序数 = 板格_末端节点_序数(一级序数)
                                        Select Case 节点_tp(末端关联节点序数)
                                            Case "硬角单元"
                                                末端关联单元序数 = 节点_硬角单元_序数(末端关联节点序数)
                                            Case "面板硬角单元"
                                                末端关联单元序数 = 节点_面板硬角单元_序数(末端关联节点序数)
                                            Case "加强筋单元"
                                                末端关联单元序数 = 节点_加强筋单元_序数(末端关联节点序数)
                                            Case "面板硬角单元"
                                                末端关联单元序数 = 节点_面板硬角单元_序数(末端关联节点序数)
                                            Case Else

                                        End Select
                                        末端关联分支序数 = 板格_末端节点_分支_序数(一级序数)

                                        板格_原始_w = 板格_w(一级序数)
                                        板格_剩余_w = 板格_原始_w - 板格_首端_w(一级序数) - 板格_末端_w(一级序数)

                                        板格_原始_A = 板格_A(一级序数)
                                        板格_首端_A = 单元分支_A(首端关联单元序数, 首端关联分支序数)
                                        板格_末端_A = 单元分支_A(末端关联单元序数, 末端关联分支序数)
                                        板格_剩余_A = 板格_原始_A - 板格_首端_A - 板格_末端_A

                                        板格_原始_AYc = 板格_AYc(一级序数)
                                        板格_首端_AYc = 单元分支_AYc(首端关联单元序数, 首端关联分支序数)
                                        板格_末端_AYc = 单元分支_AYc(末端关联单元序数, 末端关联分支序数)
                                        板格_剩余_AYc = 板格_原始_AYc - 板格_首端_AYc - 板格_末端_AYc

                                        板格_原始_AZc = 板格_AZc(一级序数)
                                        板格_首端_AZc = 单元分支_AZc(首端关联单元序数, 首端关联分支序数)
                                        板格_末端_AZc = 单元分支_AZc(末端关联单元序数, 末端关联分支序数)
                                        板格_剩余_AZc = 板格_原始_AZc - 板格_首端_AZc - 板格_末端_AZc

                                        板格_原始_AσY = 板格_AσY(一级序数)
                                        板格_首端_AσY = 单元分支_AσY(首端关联单元序数, 首端关联分支序数)
                                        板格_末端_AσY = 单元分支_AσY(末端关联单元序数, 末端关联分支序数)
                                        板格_剩余_AσY = 板格_原始_AσY - 板格_首端_AσY - 板格_末端_AσY

                                        Select Case 板格_剩余_w / 板格_原始_w
                                            Case >= 0.1
                                                加筋板单元_总数 += 1

                                                ReDim Preserve 加筋板单元_板格_序数(加筋板单元_总数)
                                                加筋板单元_板格_序数(加筋板单元_总数) = 一级序数

                                                ReDim Preserve 加筋板单元_原始_w(加筋板单元_总数)
                                                ReDim Preserve 加筋板单元_原始_A(加筋板单元_总数)
                                                ReDim Preserve 加筋板单元_原始_Yc(加筋板单元_总数)
                                                ReDim Preserve 加筋板单元_原始_Zc(加筋板单元_总数)
                                                ReDim Preserve 加筋板单元_原始_σY(加筋板单元_总数)

                                                ReDim Preserve 加筋板单元_剩余_w(加筋板单元_总数)
                                                ReDim Preserve 加筋板单元_剩余_A(加筋板单元_总数)
                                                ReDim Preserve 加筋板单元_剩余_Yc(加筋板单元_总数)
                                                ReDim Preserve 加筋板单元_剩余_Zc(加筋板单元_总数)
                                                ReDim Preserve 加筋板单元_剩余_σY(加筋板单元_总数)

                                                加筋板单元_原始_w(加筋板单元_总数) = 板格_原始_w
                                                加筋板单元_原始_A(加筋板单元_总数) = 板格_原始_A
                                                加筋板单元_原始_Yc(加筋板单元_总数) = 板格_原始_AYc / 板格_原始_A
                                                加筋板单元_原始_Zc(加筋板单元_总数) = 板格_原始_AZc / 板格_原始_A
                                                加筋板单元_原始_σY(加筋板单元_总数) = 板格_原始_AσY / 板格_原始_A

                                                加筋板单元_剩余_w(加筋板单元_总数) = 板格_剩余_w
                                                加筋板单元_剩余_A(加筋板单元_总数) = 板格_剩余_A
                                                加筋板单元_剩余_Yc(加筋板单元_总数) = 板格_剩余_AYc / 板格_剩余_A
                                                加筋板单元_剩余_Zc(加筋板单元_总数) = 板格_剩余_AZc / 板格_剩余_A
                                                加筋板单元_剩余_σY(加筋板单元_总数) = 板格_剩余_AσY / 板格_剩余_A

                                                全截面_A += 板格_剩余_A
                                                全截面_AYc += 板格_剩余_AYc
                                                全截面_AZc += 板格_剩余_AZc
                                                全截面_AσY += 板格_剩余_AσY
                                            Case Else

                                        End Select
                                    Next
                                Case Else

                            End Select
                        Next
                        全截面_Yc = 全截面_AYc / 全截面_A
                        全截面_Zc = 全截面_AZc / 全截面_A
                End Select
            Next

            '单元对象的形心坐标输出
            If 循环序数 = 0 Then
                For 单元对象类型序数 As UShort = 1 To 6
                    Select Case 单元对象类型序数
                        Case 1
                            Chart2.Series.Clear()
                            Chart2.Show()
                            Chart2.Series.Add("硬角单元_形心")
                            Chart2.Series("硬角单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                            For 一级序数 As UShort = 1 To 硬角单元_总数
                                Chart2.Series("硬角单元_形心").Points.AddXY(硬角单元_Yc(一级序数), 硬角单元_Zc(一级序数))
                            Next
                        Case 2
                    'MsgBox("面板硬角单元部分未完成!")
                        Case 3
                            Chart2.Series.Add("特别硬角单元_形心")
                            Chart2.Series("特别硬角单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                            For 一级序数 As UShort = 1 To 特别硬角单元_总数
                                Chart2.Series("特别硬角单元_形心").Points.AddXY(特别硬角单元_Yc(一级序数), 特别硬角单元_Zc(一级序数))
                            Next
                        Case 4
                            Chart2.Series.Add("加强筋单元_形心")
                            Chart2.Series("加强筋单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                            For 一级序数 As UShort = 1 To 加强筋单元_总数
                                Chart2.Series("加强筋单元_形心").Points.AddXY(加强筋单元_Yc(一级序数), 加强筋单元_Zc(一级序数))
                            Next
                        Case 5
                    'MsgBox("面板加强筋单元部分未完成!")
                        Case 6
                            Chart2.Series.Add("加筋板单元_形心")
                            Chart2.Series("加筋板单元_形心").ChartType = DataVisualization.Charting.SeriesChartType.Point
                            For 一级序数 As UShort = 1 To 加筋板单元_总数
                                Chart2.Series("加筋板单元_形心").Points.AddXY(加筋板单元_剩余_Yc(一级序数), 加筋板单元_剩余_Zc(一级序数))
                            Next
                        Case Else

                    End Select
                Next
            End If

            'Debug.Print(Format(循环序数, "0000") & " - " & Format(全截面_Yc / 1000, "000.000") & " - " & Format(全截面_Zc / 1000, "000.000") & " - " & Format(全截面_A / 1000000, "000.000"))
            基于二分法的多角度增量迭代法()
            '多角度增量迭代法()
            Debug.Print("")
        Next

        End
    End Sub

    Private Sub BoxMuller(ByVal 名义值 As Single, ByRef 随机变量 As Single)
        Randomize()
        Dim u As Single, v As Single
        u = Rnd()
        Randomize()
        v = Rnd()
        Dim sd As Single = 名义值 * 0.1
        随机变量 = sd * (Sqrt(-2 * Math.Log(u)) * Cos(2 * PI * v)) + 名义值
    End Sub

    Private Sub 关于() Handles Button9.Click
        关于框.Show()
    End Sub

    Private Sub 退出() Handles Button10.Click
        If MsgBox("是否退出程序?", vbYesNo, "退出确认") = MsgBoxResult.Yes Then
            End
        Else
            Exit Sub
        End If
    End Sub
End Class
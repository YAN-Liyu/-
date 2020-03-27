Imports System.Math
Module 模块_公有过程_通用型
    Public Sub 读入界面参数()
        '读入界面参数并执行动作
        '   第一部分   读入界面参数
        '       第一小点    读入主尺度
        跨度 = Val(主窗体.跨度_文本框.Text)
        型宽 = Val(主窗体.型宽_文本框.Text)
        型深 = Val(主窗体.型深_文本框.Text)
        '       第二小点    读入基本输入对象
        节点_总数 = Val(主窗体.节点总数_文本框.Text)
        加强筋_总数 = Val(主窗体.加强筋总数_文本框.Text)
        面板_总数 = Val(主窗体.面板总数_文本框.Text)
        板格_总数 = Val(主窗体.板格总数_文本框.Text)
        特别硬角单元_总数 = Val(主窗体.特别硬角单元总数_文本框.Text)
        '       后续处理    数组重定义
        '           输入参数部分
        ReDim 节点_Y0(节点_总数), 节点_Z0(节点_总数)

        ReDim 加强筋_Y0(加强筋_总数), 加强筋_Z0(加强筋_总数),
              加强筋_l(加强筋_总数),
              加强筋_hw(加强筋_总数), 加强筋_tw(加强筋_总数), 加强筋_αw(加强筋_总数),
              加强筋_wf(加强筋_总数), 加强筋_tf(加强筋_总数), 加强筋_αf(加强筋_总数),
              加强筋_dx(加强筋_总数), 加强筋_σY(加强筋_总数), 加强筋_tp(加强筋_总数), 加强筋_mk(加强筋_总数)
        ReDim 加强筋_σYX(加强筋_总数), 加强筋_EX(加强筋_总数), 加强筋_twX(加强筋_总数), 加强筋_tfX(加强筋_总数)

        ReDim 面板_Y0(面板_总数), 面板_Z0(面板_总数),
            面板_Y1(面板_总数), 面板_Z1(面板_总数),
            面板_YL(面板_总数), 面板_ZL(面板_总数),
            面板_l(面板_总数),
            面板_t(面板_总数), 面板_σY(面板_总数), 面板_PMA(面板_总数)
        ReDim 面板_σYX(面板_总数), 面板_EX(面板_总数), 面板_tX(面板_总数)

        Dim 通用板格板数目 As UShort = 3
        ReDim 板格_Y0(板格_总数), 板格_Z0(板格_总数),
            板格_Y1(板格_总数), 板格_Z1(板格_总数),
            板格_l(板格_总数),
            板格_tp(板格_总数),
            板格_板格板数目(板格_总数),
            板格板_w(板格_总数, 通用板格板数目),
            板格板_t(板格_总数, 通用板格板数目),
            板格板_σY(板格_总数, 通用板格板数目)
        ReDim 板格板_σYX(板格_总数, 通用板格板数目), 板格板_EX(板格_总数, 通用板格板数目), 板格板_tX(板格_总数, 通用板格板数目)

        Dim 通用子对象数目 As UShort = 5
        ReDim 子对象_数目(特别硬角单元_总数),
            子对象_Yc(特别硬角单元_总数, 通用子对象数目),
            子对象_Zc(特别硬角单元_总数, 通用子对象数目),
            子对象_A(特别硬角单元_总数, 通用子对象数目),
            子对象_σY(特别硬角单元_总数, 通用子对象数目)
        ReDim 子对象_σYX(特别硬角单元_总数, 通用子对象数目), 子对象_EX(特别硬角单元_总数, 通用子对象数目), 子对象_AX(特别硬角单元_总数, 通用子对象数目)

        '           导出参数部分
        ReDim 加强筋_Ycw(加强筋_总数), 加强筋_Zcw(加强筋_总数),
            加强筋_Ycf(加强筋_总数), 加强筋_Zcf(加强筋_总数),
            加强筋_Aw(加强筋_总数), 加强筋_Af(加强筋_总数),
            加强筋_Icw(加强筋_总数), 加强筋_Iow(加强筋_总数),
            加强筋_Icf(加强筋_总数), 加强筋_Iof(加强筋_总数),
            加强筋_df(加强筋_总数),
            加强筋_YcS(加强筋_总数), 加强筋_ZcS(加强筋_总数),
            加强筋_AS(加强筋_总数), 加强筋_hcS(加强筋_总数),
            加强筋_IoS(加强筋_总数), 加强筋_IcS(加强筋_总数)
        ReDim 加强筋_IPS(加强筋_总数), 加强筋_ITS(加强筋_总数), 加强筋_IWS(加强筋_总数),
            加强筋_ηS(加强筋_总数), 加强筋_σETS(加强筋_总数),
            加强筋_σELS(加强筋_总数)

        ReDim 面板_Yc(面板_总数), 面板_Zc(面板_总数),
            面板_w(面板_总数), 面板_A(面板_总数)

        ReDim 板格板_Y0(板格_总数, 通用板格板数目), 板格板_Z0(板格_总数, 通用板格板数目),
            板格板_Y1(板格_总数, 通用板格板数目), 板格板_Z1(板格_总数, 通用板格板数目),
            板格板_Yc(板格_总数, 通用板格板数目), 板格板_Zc(板格_总数, 通用板格板数目),
            板格板_A(板格_总数, 通用板格板数目),
            板格板_AYc(板格_总数, 通用板格板数目), 板格板_AZc(板格_总数, 通用板格板数目), 板格板_AσY(板格_总数, 通用板格板数目)
        ReDim 板格_α(板格_总数), 板格_w(板格_总数),
            板格_A(板格_总数), 板格_t(板格_总数),
            板格_AYc(板格_总数), 板格_AZc(板格_总数), 板格_AσY(板格_总数),
            板格_Yc(板格_总数), 板格_Zc(板格_总数), 板格_σY(板格_总数)

        ReDim 子对象_AYc(特别硬角单元_总数, 通用子对象数目), 子对象_AZc(特别硬角单元_总数, 通用子对象数目), 子对象_AσY(特别硬角单元_总数, 通用子对象数目)
        ReDim 特别硬角单元_A(特别硬角单元_总数),
            特别硬角单元_AYc(特别硬角单元_总数), 特别硬角单元_AZc(特别硬角单元_总数), 特别硬角单元_AσY(特别硬角单元_总数),
            特别硬角单元_Yc(特别硬角单元_总数), 特别硬角单元_Zc(特别硬角单元_总数), 特别硬角单元_σY(特别硬角单元_总数)

        '       第三小点    读入弯矩计算控制参数
        χ_初值 = Abs(Val(主窗体.DataGridView4.Rows(0).Cells(1).Value))
        χ_增量 = Abs(Val(主窗体.DataGridView4.Rows(0).Cells(2).Value))
        χ_总数 = Val(主窗体.DataGridView4.Rows(0).Cells(3).Value)
        Select Case Mid(第五部分, 5, 1)
            Case "1"
                ζ_初值 = Val(主窗体.DataGridView4.Rows(1).Cells(1).Value)
                ζ_增量 = Abs(Val(主窗体.DataGridView4.Rows(1).Cells(2).Value))
                ζ_总数 = Val(主窗体.DataGridView4.Rows(1).Cells(3).Value)
                'γ_初值 = Val(主窗体.DataGridView4.Rows(1).Cells(1).Value)
                'γ_增量 = Abs(Val(主窗体.DataGridView4.Rows(1).Cells(2).Value))
                'γ_总数 = Val(主窗体.DataGridView4.Rows(1).Cells(3).Value)
            Case "2"
                'ζ_初值 = Val(主窗体.DataGridView4.Rows(1).Cells(1).Value)
                'ζ_增量 = Abs(Val(主窗体.DataGridView4.Rows(1).Cells(2).Value))
                'ζ_总数 = Val(主窗体.DataGridView4.Rows(1).Cells(3).Value)
                γ_初值 = Val(主窗体.DataGridView4.Rows(1).Cells(1).Value)
                γ_增量 = Abs(Val(主窗体.DataGridView4.Rows(1).Cells(2).Value))
                γ_总数 = Val(主窗体.DataGridView4.Rows(1).Cells(3).Value)
        End Select
        ''ζ_初值 = Val(主窗体.DataGridView4.Rows(1).Cells(1).Value)
        ''ζ_增量 = Abs(Val(主窗体.DataGridView4.Rows(1).Cells(2).Value))
        ''ζ_总数 = Val(主窗体.DataGridView4.Rows(1).Cells(3).Value)
        'γ_初值 = Val(主窗体.DataGridView4.Rows(1).Cells(1).Value)
        'γ_增量 = Abs(Val(主窗体.DataGridView4.Rows(1).Cells(2).Value))
        'γ_总数 = Val(主窗体.DataGridView4.Rows(1).Cells(3).Value)
        α_初值 = Val(主窗体.DataGridView4.Rows(2).Cells(1).Value) * PI / 180
        α_增量 = Abs(Val(主窗体.DataGridView4.Rows(2).Cells(2).Value) * PI / 180)
        α_总数 = Val(主窗体.DataGridView4.Rows(2).Cells(3).Value)
        'Debug.Print(χ_初值 & vbTab & χ_增量 & vbTab & χ_总数 & vbCrLf &
        '            ζ_初值 & vbTab & ζ_增量 & vbTab & ζ_总数 & vbCrLf &
        '            α_初值 & vbTab & α_增量 & vbTab & α_总数 & vbCrLf)

        '       单元划分
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


        '   第二部分    执行动作
        '       第一小点    根据  基本输入对象总数  初始化   DataGridView1
        '           第一分点    节点
        主窗体.DataGridView1.Rows.Add("节点", "", "共有", 节点_总数, "个")
        For i As UShort = 1 To 节点_总数
            节点_序数 = i
            主窗体.DataGridView1.Rows.Add(节点_序数, "", "", "-", "-")
        Next
        '           第二分点    加强筋
        主窗体.DataGridView1.Rows.Add("加强筋", "", "共有", 加强筋_总数, "个")
        For i As UShort = 1 To 加强筋_总数
            加强筋_序数 = i
            主窗体.DataGridView1.Rows.Add(加强筋_序数, "", "", "", "")
        Next
        '           第三分点    面板
        主窗体.DataGridView1.Rows.Add("面板", "", "共有", 面板_总数, "个")
        For i As UShort = 1 To 面板_总数
            面板_序数 = i
            主窗体.DataGridView1.Rows.Add(面板_序数, "", "", "", "")
        Next
        '           第四分点    板格
        主窗体.DataGridView1.Rows.Add("板格", "", "共有", 板格_总数, "个")
        For i As UShort = 1 To 板格_总数
            板格_序数 = i
            主窗体.DataGridView1.Rows.Add(板格_序数, "", "", "", "")
        Next
        '           第五分点    特别硬角单元
        主窗体.DataGridView1.Rows.Add("特别硬角单元", "", "共有", 特别硬角单元_总数, "个")
        For i As UShort = 1 To 特别硬角单元_总数
            特别硬角单元_序数 = i
            主窗体.DataGridView1.Rows.Add(特别硬角单元_序数, "", "", "-", "-")
        Next
        主窗体.DataGridView1.AllowUserToAddRows = False
        主窗体.DataGridView1.AllowUserToDeleteRows = False

    End Sub

    Public Sub 初始化图表()
        主窗体.Chart1.ChartAreas(0).Axes(0).Title = 水平轴_标题_几何结构
        主窗体.Chart1.ChartAreas(0).Axes(1).Title = 垂直轴_标题_几何结构
        主窗体.Chart1.Series.Clear()
        主窗体.Chart2.ChartAreas(0).Axes(0).Title = 水平轴_标题_单元
        主窗体.Chart2.ChartAreas(0).Axes(1).Title = 垂直轴_标题_单元
        主窗体.Chart3.ChartAreas(0).Axes(0).Title = 水平轴_标题_极限承载力
        主窗体.Chart3.ChartAreas(0).Axes(1).Title = 垂直轴_标题_极限承载力
    End Sub

    'Public Sub 填入内部数据()
    '    Select Case InputBox("船型", "船型", "OT")
    '        Case "OT"
    '            Select Case InputBox("工况", "工况", "完整")
    '                Case "完整"

    '                Case "规范碰撞受损"

    '                Case "规范对称搁浅受损"

    '                Case "规范非对称搁浅受损"

    '            End Select

    '            Select Case InputBox("横倾角度", "横倾角度", "0")
    '                Case "0"

    '                Case "15"

    '                Case "30"

    '                Case "45"

    '                Case "60"

    '                Case "75"

    '                Case "90"

    '                Case "105"

    '                Case "120"

    '                Case "135"

    '                Case "150"

    '                Case "165"
    '            End Select
    '        Case "BC"

    '    End Select
    '    Select Case InputBox("模板文件代号", "选择模板")
    '        Case "OT_Collision_-75deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "5.5", "0.01", "200")
    '                .Rows.Add("α/(°)", "-53.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_-60deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "8.1", "0.01", "200")
    '                .Rows.Add("α/(°)", "-38.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_-45deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "9.4", "0.01", "200")
    '                .Rows.Add("α/(°)", "-27.5", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_-30deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "10.2", "0.01", "200")
    '                .Rows.Add("α/(°)", "-19.5", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_-15deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "10.8", "0.01", "200")
    '                .Rows.Add("α/(°)", "-12.5", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_0deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "11.4", "0.01", "200")
    '                .Rows.Add("α/(°)", "-5.8", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_15deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "12", "0.01", "200")
    '                .Rows.Add("α/(°)", "1.4", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_30deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "12.8", "0.01", "200")
    '                .Rows.Add("α/(°)", "10.2", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_45deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "14.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "22.4", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_60deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "16.2", "0.01", "200")
    '                .Rows.Add("α/(°)", "41.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_75deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "25.5", "0.01", "200")
    '                .Rows.Add("α/(°)", "70.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Collision_90deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "-10.0", "0.01", "2000")
    '                .Rows.Add("α/(°)", "100.0", "0.1", "10")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With

    '        Case "OT_Intact_-75deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                '.Rows.Add("ζ or γ/(m)", "13.0", "0.01", "200")
    '                '.Rows.Add("α/(°)", "-31.0", "0.1", "100")
    '                .Rows.Add("ζ or γ/(m)", "10.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "-52.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_-60deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                '.Rows.Add("ζ or γ/(m)", "13.0", "0.01", "200")
    '                '.Rows.Add("α/(°)", "-31.0", "0.1", "100")
    '                .Rows.Add("ζ or γ/(m)", "22.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "-31.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_-45deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "13.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "-19.0", "0.1", "100")
    '                '.Rows.Add("ζ or γ/(m)", "38.0", "0.01", "200")
    '                '.Rows.Add("α/(°)", "-19.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_-30deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "13.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "-11.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_-15deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "13.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "-5.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_0deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "13.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "0.0", "0.1", "1")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_15deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "13.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "5.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_30deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "12.9", "0.01", "200")
    '                .Rows.Add("α/(°)", "11.3", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_45deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "13.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "19.0", "0.1", "100")
    '                '.Rows.Add("ζ or γ/(m)", "-38.0", "0.01", "200")
    '                '.Rows.Add("α/(°)", "19.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_60deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "13.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "31.0", "0.1", "100")
    '                '.Rows.Add("ζ or γ/(m)", "-22.0", "0.01", "200")
    '                '.Rows.Add("α/(°)", "31.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_75deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                '.Rows.Add("ζ or γ/(m)", "-10.0", "0.01", "200")
    '                '.Rows.Add("α/(°)", "52.0", "0.1", "100")
    '                .Rows.Add("ζ or γ/(m)", "-10.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "52.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case "OT_Intact_90deg"
    '            主窗体.跨度_文本框.Text = 4950
    '            主窗体.型宽_文本框.Text = 58000
    '            主窗体.型深_文本框.Text = 32000
    '            主窗体.节点总数_文本框.Text = 422
    '            主窗体.加强筋总数_文本框.Text = 392
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 432
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "0.0", "0.01", "200")
    '                .Rows.Add("α/(°)", "90.0", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With

    '        Case "BC"
    '            主窗体.跨度_文本框.Text = 5220
    '            主窗体.型宽_文本框.Text = 50000
    '            主窗体.型深_文本框.Text = 27000
    '            主窗体.节点总数_文本框.Text = 228
    '            主窗体.加强筋总数_文本框.Text = 188
    '            主窗体.面板总数_文本框.Text = 0
    '            主窗体.板格总数_文本框.Text = 239
    '            主窗体.特别硬角单元总数_文本框.Text = 4
    '            With 主窗体.DataGridView4
    '                .Rows.Add("χ/(1/m)", "0.00000", "0.00001", "50")
    '                .Rows.Add("ζ or γ/(m)", "", "0.01", "200")
    '                .Rows.Add("α/(°)", "", "0.1", "100")
    '                .AllowUserToAddRows = False
    '                .AllowUserToDeleteRows = False
    '            End With
    '        Case Else

    '    End Select
    'End Sub
End Module
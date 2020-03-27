Module 模块_公有定义
    '1.     界面参数
    '1.1.       主尺度
    Public 跨度 As Single, 型宽 As Single, 型深 As Single
    '1.2.       基本输入对象总数
    Public 节点_总数 As UShort, 加强筋_总数 As UShort, 面板_总数 As UShort, 板格_总数 As UShort, 特别硬角单元_总数 As UShort

    '2.     基本对象属性参数
    '2.1.       输入参数
    '2.1.1          节点
    Public 节点_Y0() As Single, 节点_Z0() As Single
    '2.1.2          加强筋
    Public 加强筋_Y0() As Single, 加强筋_Z0() As Single,
        加强筋_l() As Single,
        加强筋_hw() As Single, 加强筋_tw() As Single, 加强筋_αw() As Single,
        加强筋_wf() As Single, 加强筋_tf() As Single, 加强筋_αf() As Single,
        加强筋_dx() As Single,
        加强筋_σY() As Single,
        加强筋_tp() As String, 加强筋_mk() As Boolean
    '2.1.3          面板
    Public 面板_Y0() As Single, 面板_Z0() As Single,
        面板_Y1() As Single, 面板_Z1() As Single,
        面板_YL() As Single, 面板_ZL() As Single,
        面板_l() As Single,
        面板_t() As Single,
        面板_σY() As Single,
        面板_PMA() As Boolean
    '2.1.4          板格
    Public 板格_Y0() As Single, 板格_Z0() As Single,
        板格_Y1() As Single, 板格_Z1() As Single,
        板格_l() As Single,
        板格_tp() As String,
        板格_板格板数目() As UShort,
            板格板_w(,) As Single, 板格板_t(,) As Single,
            板格板_σY(,) As Single
    '2.1.5          特别硬角单元
    Public 子对象_数目() As UShort,
            子对象_Yc(,) As Single, 子对象_Zc(,) As Single,
            子对象_A(,) As Single,
            子对象_σY(,) As Single

    '###随机化参数补充定义
    Public 样本总数 As UShort
    Public 加强筋_EX() As Single, 面板_EX() As Single, 板格板_EX(,) As Single, 子对象_EX(,) As Single
    Public 加强筋_σYX() As Single, 面板_σYX() As Single, 板格板_σYX(,) As Single, 子对象_σYX(,) As Single
    Public 加强筋_twX() As Single, 加强筋_tfX() As Single, 面板_tX() As Single, 板格板_tX(,) As Single, 子对象_AX(,) As Single

    '2.2.       导出参数
    '2.1.1          节点
    '                   [无]
    '2.1.2          加强筋
    Public 加强筋_Ycw() As Single, 加强筋_Zcw() As Single,
        加强筋_Ycf() As Single, 加强筋_Zcf() As Single,
        加强筋_Aw() As Single, 加强筋_Af() As Single,
        加强筋_Icw() As Single, 加强筋_Iow() As Single,
        加强筋_Icf() As Single, 加强筋_Iof() As Single,
        加强筋_df() As Single,
        加强筋_YcS() As Single, 加强筋_ZcS() As Single,
        加强筋_AS() As Single, 加强筋_hcS() As Single,
        加强筋_IoS() As Single, 加强筋_IcS() As Single
    '
    Public 加强筋_IPS() As Single, 加强筋_ITS() As Single, 加强筋_IWS() As Single,
        加强筋_ηS() As Single, 加强筋_σETS() As Single,
        加强筋_σELS() As Single
    '2.1.3          面板
    Public 面板_Yc() As Single, 面板_Zc() As Single,
        面板_w() As Single, 面板_A() As Single
    '2.1.4          板格/板格板
    Public 板格板_Y0(,) As Single, 板格板_Z0(,) As Single,
        板格板_Y1(,) As Single, 板格板_Z1(,) As Single,
        板格板_Yc(,) As Single, 板格板_Zc(,) As Single,
        板格板_A(,) As Single,
        板格板_AYc(,) As Single, 板格板_AZc(,) As Single, 板格板_AσY(,) As Single
    Public 板格_Yc() As Single, 板格_Zc() As Single,
        板格_w() As Single, 板格_t() As Single, 板格_α() As Single,
        板格_A() As Single,
        板格_AYc() As Single, 板格_AZc() As Single, 板格_AσY() As Single,
        板格_σY() As Single
    '2.1.5          特别硬角单元
    Public 子对象_AYc(,) As Single, 子对象_AZc(,) As Single, 子对象_AσY(,) As Single
    'Public 特别硬角单元_A() As Single,
    '    特别硬角单元_Yc() As Single, 特别硬角单元_Zc() As Single, 特别硬角单元_σY() As Single
    Public 特别硬角单元_AYc() As Single, 特别硬角单元_AZc() As Single, 特别硬角单元_AσY() As Single


    '单元划分
    Public 节点_分支_数目(节点_总数) As UShort
    Public 通用分支数目 As UShort = 5
    Public 节点_分支_板格_序数(节点_总数, 通用分支数目) As UShort

    Public 节点_分支_α(节点_总数, 通用分支数目) As Single

    Public 板格_首端节点_序数(板格_总数) As UShort
    Public 板格_首端节点_分支_序数(板格_总数) As UShort
    Public 板格_末端节点_序数(板格_总数) As UShort
    Public 板格_末端节点_分支_序数(板格_总数) As UShort

    Public 节点_tp(节点_总数) As String

    Public 节点_加强筋_序数(节点_总数) As UShort
    Public 加强筋_节点_序数(加强筋_总数) As UShort

    Public 节点_面板_序数(节点_总数) As UShort
    Public 面板_节点_序数(面板_总数) As UShort

    Public 节点_硬角单元_序数(节点_总数) As UShort
    Public 硬角单元_节点_序数() As UShort

    Public 节点_面板硬角单元_序数(节点_总数) As UShort
    Public 面板硬角单元_节点_序数() As UShort

    Public 节点_加强筋单元_序数(节点_总数) As UShort
    Public 加强筋单元_节点_序数() As UShort

    Public 节点_面板加强筋单元_序数(节点_总数) As UShort
    Public 面板加强筋单元_节点_序数() As UShort

    Public 节点_分支_首端节点_序数(节点_总数, 通用分支数目) As UShort, 节点_分支_末端节点_序数(节点_总数, 通用分支数目) As UShort

    Public 板格_首端_w(板格_总数) As Single, 板格_末端_w(板格_总数) As Single

    '横向板格残余部分

    Public χ_初值 As Single, ζ_初值 As Single, γ_初值 As Single, α_初值 As Single    '弯矩计算控制参数
    Public χ_增量 As Single, ζ_增量 As Single, γ_增量 As Single, α_增量 As Single    '弯矩计算控制参数
    Public χ_瞬时 As Single, ζ_瞬时 As Single, γ_瞬时 As Single, α_瞬时 As Single

    Public χ_总数 As UShort, ζ_总数 As UShort, γ_总数 As UShort, α_总数 As UShort    '弯矩计算控制参数
    Public χ_序数 As UShort, ζ_序数 As UShort, γ_序数 As UShort, α_序数 As UShort
    Public χ_临界 As UShort, ζ_临界 As UShort, γ_临界 As UShort, α_临界 As UShort

    Public 节点_序数 As UShort, 加强筋_序数 As UShort, 面板_序数 As UShort, 板格_序数 As UShort, 特别硬角单元_序数 As UShort   '基本输入对象序数

    Public 水平轴_标题_几何结构 As String = "Y轴(船宽方向)/(mm)"   '图表坐标轴标题
    Public 垂直轴_标题_几何结构 As String = "Z轴(吃水方向)/(mm)"   '图表坐标轴标题
    Public 水平轴_标题_单元 As String = "Y轴(船宽方向)/(mm)"   '图表坐标轴标题
    Public 垂直轴_标题_单元 As String = "Z轴(吃水方向)/(mm)"   '图表坐标轴标题
    Public 水平轴_标题_极限承载力 As String = "曲率/(1/m)"     '图表坐标轴标题
    Public 垂直轴_标题_极限承载力 As String = "弯矩/(N.m)"     '图表坐标轴标题

    Public Const 标准弹性模量 As Single = 210000.0   'MPa

    Public 单元_总数 As UShort
    Public 硬角单元_总数 As UShort, 面板硬角单元_总数 As UShort
    Public 加强筋单元_总数 As UShort, 面板加强筋单元_总数 As UShort
    Public 加筋板单元_总数 As UShort

    Public 单元_序数 As UShort
    Public 硬角单元_序数 As UShort, 面板硬角单元_序数 As UShort
    Public 加强筋单元_序数 As UShort, 面板加强筋单元_序数 As UShort
    Public 加筋板单元_序数 As UShort

    Public 硬角单元_Yc() As Single, 硬角单元_Zc() As Single, 硬角单元_L() As Single
    Public 硬角单元_A() As Single
    Public 硬角单元_E() As Single, 硬角单元_σY() As Single
    Public 硬角单元_εO() As Single, 硬角单元_εY() As Single, 硬角单元_εR() As Single

    Public 面板硬角单元_Yc() As Single, 面板硬角单元_Zc() As Single, 面板硬角单元_L() As Single
    Public 面板硬角单元_A() As Single
    Public 面板硬角单元_E() As Single, 面板硬角单元_σY() As Single
    Public 面板硬角单元_εO() As Single, 面板硬角单元_εY() As Single, 面板硬角单元_εR() As Single

    Public 特别硬角单元_Yc() As Single, 特别硬角单元_Zc() As Single, 特别硬角单元_L() As Single
    Public 特别硬角单元_A() As Single
    Public 特别硬角单元_E() As Single, 特别硬角单元_σY() As Single
    Public 特别硬角单元_εO() As Single, 特别硬角单元_εY() As Single, 特别硬角单元_εR() As Single

    Public 加强筋单元_Yc() As Single, 加强筋单元_Zc() As Single, 加强筋单元_L() As Single
    Public 加强筋单元_A() As Single
    Public 加强筋单元_E() As Single, 加强筋单元_σY() As Single
    Public 加强筋单元_εO() As Single, 加强筋单元_εY() As Single, 加强筋单元_εR() As Single

    Public 加强筋单元_σYP() As Single,
        加强筋单元_lP() As Single,
        加强筋单元_wP() As Single, 加强筋单元_tP() As Single,
        加强筋单元_σYS() As Single,
        加强筋单元_lS() As Single,
        加强筋单元_hw() As Single, 加强筋单元_tw() As Single,
        加强筋单元_wf() As Single, 加强筋单元_tf() As Single,
        加强筋单元_dx() As Single,
        加强筋单元_tpS() As String, 加强筋单元_mk() As Boolean

    Public 加强筋单元_ηS() As Single, 加强筋单元_σETS() As Single
    Public 加强筋单元_σELS() As Single

    Public 面板加强筋单元_Yc() As Single, 面板加强筋单元_Zc() As Single, 面板加强筋单元_L() As Single
    Public 面板加强筋单元_A() As Single
    Public 面板加强筋单元_E() As Single, 面板加强筋单元_σY() As Single
    Public 面板加强筋单元_εO() As Single, 面板加强筋单元_εY() As Single, 面板加强筋单元_εR() As Single

    Public 加筋板单元_板格_序数() As Single
    Public 加筋板单元_原始_w() As Single, 加筋板单元_原始_A() As Single,
        加筋板单元_原始_Yc() As Single, 加筋板单元_原始_Zc() As Single,
        加筋板单元_原始_σY() As Single
    Public 加筋板单元_剩余_w() As Single, 加筋板单元_剩余_A() As Single,
        加筋板单元_剩余_Yc() As Single, 加筋板单元_剩余_Zc() As Single,
        加筋板单元_剩余_σY() As Single
    Public 加筋板单元_Yc() As Single, 加筋板单元_Zc() As Single, 加筋板单元_L() As Single
    Public 加筋板单元_A() As Single
    Public 加筋板单元_E() As Single, 加筋板单元_σY() As Single
    Public 加筋板单元_εO() As Single, 加筋板单元_εY() As Single, 加筋板单元_εR() As Single

    Public 硬角单元_D(,) As Single, 面板硬角单元_D(,) As Single, 特别硬角单元_D(,) As Single,
        加强筋单元_屈服_D(,) As Single, 加强筋单元_屈曲_D(,) As Single, 面板加强筋单元_D(,) As Single, 加筋板单元_剩余_D(,) As Single

    Public 单元_单元类型 As String

    Public 单元_拉压状态 As String

    Public 单元_Yc As Single, 单元_Zc As Single, 单元_L
    Public 单元_A As Single
    Public 单元_E As Single, 单元_σY As Single
    Public 单元_εO As Single, 单元_εY As Single, 单元_εR As Single
    Public 单元_D As Single

    Public 屈服应力 As Single, 梁柱应力 As Single, 弯扭应力 As Single, 局部应力 As Single, 板屈应力 As Single

    Public 单元_σO As Single, 单元_FO As Single, 单元_MO As Single
    Public 单元_真实应力 As Single, 单元_真实轴力 As Single, 单元_真实弯矩 As Single

    Public 中性轴_Yc As Single, 中性轴_Zc As Single
    Public 中性轴_z As Single, 中性轴_α As Single
    Public 中性轴_方程 As String

    Public 参数化输入 As String
    Public 第一部分 As String, 第二部分 As String, 第三部分 As String, 第四部分 As String, 第五部分 As String
    Public 水线倾角_α As Single
    Public 合力倾角_α As Single

    Public 全截面_A As Single,
        全截面_Yc As Single, 全截面_Zc As Single,
        全截面_AYc As Single, 全截面_AZc As Single,
        全截面_AσY As Single

    Public 全截面_轴力_和值 As Single
    Public 全截面_轴力_阈值 As Single

    Public 全截面_弯矩_和值 As Single

    Public 全截面_完整性状态 As String

    Public Enum 拉压状态
        拉伸
        压缩
    End Enum
    Public Enum 单元类型
        硬角单元
        面板硬角单元
        特别硬角单元
        加强筋单元
        面板加强筋单元
        加筋板单元
    End Enum
End Module
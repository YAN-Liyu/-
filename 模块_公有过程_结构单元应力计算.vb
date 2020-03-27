Imports System.Math
Module 模块_公有过程_单元应力计算
    Public Sub 单元真实应力计算(χ_瞬时, ζ_瞬时, θ_瞬时, '弯矩计算控制
                   单元类型, 拉压状态, 屈服强度,    '单元
                   Yc坐标, Zc坐标, 横截面积,    '单元
                   真实轴力, 真实弯矩   '结果
                   )
        Dim 相对距离 As Single,
            εO As Single, εY As Single, εR As Single,
            Φ As Single,
            真实应力 As Single

        '相对距离
        相对距离 = (-Yc坐标 * Sin(θ_瞬时)) + ((Zc坐标 - ζ_瞬时) * Cos(θ_瞬时))
        'εO
        εO = 相对距离 * χ_瞬时
        'εY
        εY = 屈服强度 / 标准弹性模量
        'εR
        εR = εO / εY
        'Φ
        Select Case εR
            Case < -1
                Φ = -1
            Case > 1
                Φ = 1
            Case Else
                Φ = εR
        End Select

        '拉压状态
        拉压状态 = If(相对距离 > 0, 拉压状态.拉伸, 拉压状态.压缩)

        受拉屈服(屈服应力, Φ, 屈服强度)
        '确定 真实应力    (根据 单元类型及状态等)
        Select Case 单元类型
            Case 单元类型.硬角单元, 单元类型.面板硬角单元, 单元类型.特别硬角单元
                真实应力 = 屈服应力
            Case 单元类型.加强筋单元, 单元类型.面板加强筋单元
                Select Case 拉压状态
                    Case 拉压状态.拉伸
                        真实应力 = 屈服应力
                    Case 拉压状态.压缩
                        '[待改]梁柱屈曲()

                        '[待改]弯扭屈曲()

                        '[待改]局部屈曲()

                        真实应力 = Min(Min(梁柱应力, 弯扭应力), 局部应力)
                End Select
            Case 单元类型.加筋板单元
                Select Case 拉压状态
                    Case 拉压状态.拉伸
                        真实应力 = 屈服应力
                    Case 拉压状态.压缩
                        '[待改]板材屈曲()

                        真实应力 = 板屈应力
                End Select
        End Select

        '确定 真实轴力/真实弯矩
        真实轴力 = 真实应力 * 横截面积
        真实弯矩 = 真实轴力 * 相对距离

    End Sub

    Public Sub 受拉屈服(屈服应力, Φ, 屈服强度)
        '屈服应力
        屈服应力 = Φ * 屈服强度
    End Sub

    Public Sub 梁柱屈曲()
        '梁柱应力

    End Sub

    Public Sub 弯扭屈曲()
        '弯扭应力

    End Sub

    Public Sub 局部屈曲()
        '局部应力

    End Sub

    Public Sub 板材屈曲()

    End Sub
End Module

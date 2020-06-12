Attribute VB_Name = "TreeViewModule"
Option Explicit

Public Sub loadtreeview()
    With Mainfrm
    
    .LeftTreeView.Nodes.Add , , "Visual Basic 6.0", "Visual Basic 6.0"
    .LeftTreeView.Nodes.Add , , "Support Statement for Visual Basic 6.0", "Windows 上 Visual Basic 6.0 的支持说明"
    .LeftTreeView.Nodes.Add , , "Partner Offers", "合作伙伴优惠"
    .LeftTreeView.Nodes.Add , , "Product Documentation", "产品文档"
        .LeftTreeView.Nodes.Add "Product Documentation", tvwChild, "Visual Basic Documentation Map", "Visual Basic 文档图"
            .LeftTreeView.Nodes.Add "Visual Basic Documentation Map", tvwChild, "Visual Basic Documentation Map2", "Visual Basic 文档图"
            .LeftTreeView.Nodes.Add "Visual Basic Documentation Map", tvwChild, "Visual Basic Editions", "Visual Basic 版本"
            .LeftTreeView.Nodes.Add "Visual Basic Documentation Map", tvwChild, "Visual Basic Enterprise Edition Features", "Visual Basic 企业版特性"
            
    .LeftTreeView.Nodes.Add , , "Controls Reference", "控件参考资料"
        .LeftTreeView.Nodes.Add "Controls Reference", tvwChild, "Intrinsic Controls", "内部控件"
            .LeftTreeView.Nodes.Add "Intrinsic Controls", tvwChild, "CheckBox Control", "复选框控件"
        'LeftTreeView.Nodes.Add "Product Documentation", tvwChild, "Visual Basic Documentation Map", "Visual Basic 文档图"
        'LeftTreeView.Nodes.Add "Product Documentation", tvwChild, "Visual Basic Documentation Map", "Visual Basic 文档图"
        .LeftTreeView.Nodes.Add "Controls Reference", tvwChild, "Others Controls", "其他开源控件"
    

    
    .LeftTreeView.Nodes.Add , , "Form Beauty", "窗体美化"
        .LeftTreeView.Nodes.Add "Form Beauty", tvwChild, "Form Theme", "窗体主题"

    .LeftTreeView.Nodes.Add , , "Article", "文章"
        .LeftTreeView.Nodes.Add "Article", tvwChild, "Why Visual Basic 6 Still Thrives", "为什么 Visual Basic 6 依然经久不衰"
        .LeftTreeView.Nodes.Add "Article", tvwChild, "VB6 and the Art of the Knuckleball", "VB6 和 Knuckleball（关节球）艺术"

    End With
End Sub

Attribute VB_Name = "TreeViewModule"
Option Explicit

Public Sub loadtreeview()
    With Mainfrm
    
    .LeftTreeView.Nodes.Add , , "Visual Basic 6.0", "Visual Basic 6.0"
    .LeftTreeView.Nodes.Add , , "Support Statement for Visual Basic 6.0", "Windows �� Visual Basic 6.0 ��֧��˵��"
    .LeftTreeView.Nodes.Add , , "Partner Offers", "��������Ż�"
    .LeftTreeView.Nodes.Add , , "Product Documentation", "��Ʒ�ĵ�"
        .LeftTreeView.Nodes.Add "Product Documentation", tvwChild, "Visual Basic Documentation Map", "Visual Basic �ĵ�ͼ"
            .LeftTreeView.Nodes.Add "Visual Basic Documentation Map", tvwChild, "Visual Basic Documentation Map2", "Visual Basic �ĵ�ͼ"
            .LeftTreeView.Nodes.Add "Visual Basic Documentation Map", tvwChild, "Visual Basic Editions", "Visual Basic �汾"
            .LeftTreeView.Nodes.Add "Visual Basic Documentation Map", tvwChild, "Visual Basic Enterprise Edition Features", "Visual Basic ��ҵ������"
            
    .LeftTreeView.Nodes.Add , , "Controls Reference", "�ؼ��ο�����"
        .LeftTreeView.Nodes.Add "Controls Reference", tvwChild, "Intrinsic Controls", "�ڲ��ؼ�"
            .LeftTreeView.Nodes.Add "Intrinsic Controls", tvwChild, "CheckBox Control", "��ѡ��ؼ�"
        'LeftTreeView.Nodes.Add "Product Documentation", tvwChild, "Visual Basic Documentation Map", "Visual Basic �ĵ�ͼ"
        'LeftTreeView.Nodes.Add "Product Documentation", tvwChild, "Visual Basic Documentation Map", "Visual Basic �ĵ�ͼ"
        .LeftTreeView.Nodes.Add "Controls Reference", tvwChild, "Others Controls", "������Դ�ؼ�"
    

    
    .LeftTreeView.Nodes.Add , , "Form Beauty", "��������"
        .LeftTreeView.Nodes.Add "Form Beauty", tvwChild, "Form Theme", "��������"

    .LeftTreeView.Nodes.Add , , "Article", "����"
        .LeftTreeView.Nodes.Add "Article", tvwChild, "Why Visual Basic 6 Still Thrives", "Ϊʲô Visual Basic 6 ��Ȼ���ò�˥"
        .LeftTreeView.Nodes.Add "Article", tvwChild, "VB6 and the Art of the Knuckleball", "VB6 �� Knuckleball���ؽ�������"

    End With
End Sub

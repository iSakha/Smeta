Public Class smetaForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim itemExist, subItemExist As Boolean
        itemExist = False
        Dim item As New ToolStripMenuItem(cmb_item.Text)

        If menuStripCat.DropDownItems.Count = 0 Then
            menuStripCat.DropDownItems.Add(item)
        End If

        For Each itm As ToolStripMenuItem In menuStripCat.DropDownItems

            If itm.ToString = cmb_item.Text Then
                itemExist = True
            End If

        Next itm

        If Not itemExist Then
            menuStripCat.DropDownItems.Add(item)
        End If

        itemExist = False

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim subItemExist As Boolean
        subItemExist = False
        Dim subItem As New ToolStripMenuItem(txt_subitem.Text)
        Dim item As New ToolStripMenuItem
        Dim index As Integer

        For Each itm As ToolStripMenuItem In menuStripCat.DropDownItems

            If itm.ToString = cmb_item.Text Then
                index = menuStripCat.DropDownItems.IndexOf(itm)
                item = menuStripCat.DropDownItems.Item(index)
                Console.WriteLine(item)
            End If

        Next itm

        If item.DropDownItems.Count = 0 Then
            item.DropDownItems.Add(subItem)
        End If

        For Each itm As ToolStripMenuItem In item.DropDownItems

            If itm.ToString = txt_subitem.Text Then
                subItemExist = True
            End If

        Next itm

        If Not subItemExist Then
            item.DropDownItems.Add(subItem)
        End If

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim itm As ToolStripMenuItem
        For Each itm In menuStripCat.DropDownItems
            Me.menuStripCat.DropDownItems.Add("test", Nothing, AddressOf MenuItem_Click)
        Next

    End Sub




End Class
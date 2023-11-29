Module Module1
    Private Const client_id As String = "148c24d0-ed6f-48d8-9e30-66ca44216e89"
    Private Const tenant_id As String = "918ed37d-87b5-4d0c-a185-e8e7523e6f8d"
    Private Const client_secret As String = "9yR8Q~86_VD0z4hpgKSFDSgSpeH.-w4hXC5n8byb"
    Private Const sharepointsite_id As String = "606ba5e7-9328-4b07-a52f-02e8c5871ee2"
    Sub Main()
        'GetItem()
        'GetMyDrive()
        'GetDriveItems()
        'UpdateFileName()
        'UploadFile()

        'DeleteFile()
        'RestoreFile() 'not working as not able to retrieve item id in recycle bin

        'MoveFile()
        'PreviewFile()
        GetFileStream()

        Console.ReadKey()
    End Sub
    Async Sub GetItem()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)

        Dim item = Await sp.GetItem(sharepointsite_id, "11839f2f-36e9-4154-b71e-ced88a36e014", "1")

        Console.WriteLine($"Title {item.Name}, Name {item.Title}")
        Console.ReadKey()
    End Sub

    Async Sub GetMyDrive()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)
        Dim driveJson = Await sp.GetMyDrive()
        Console.WriteLine(driveJson)
        Console.ReadKey()

    End Sub
    Async Sub GetDriveItems()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)
        Dim json = Await sp.GetDriveItems()
        Console.WriteLine(json)
        Console.ReadKey()

    End Sub

    Async Sub UpdateFileName()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)
        Dim json = Await sp.UpdateFileName()
        Console.WriteLine(json)
        Console.ReadKey()

    End Sub

    Async Sub UploadFile()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)
        Dim json = Await sp.UploadFile()
        Console.WriteLine(json)
        Console.ReadKey()

    End Sub

    Async Sub DeleteFile()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)
        Dim json = Await sp.DeleteFile()
        Console.WriteLine(json)
        Console.ReadKey()

    End Sub

    Async Sub RestoreFile()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)
        Dim json = Await sp.RestoreDeletedFile()
        Console.WriteLine(json)
        Console.ReadKey()

    End Sub

    Async Sub MoveFile()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)
        Dim json = Await sp.MoveItem()
        Console.WriteLine(json)
        Console.ReadKey()

    End Sub

    Async Sub PreviewFile()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)
        Dim json = Await sp.PreviewItem()
        Console.WriteLine(json)
        Console.ReadKey()

    End Sub

    Async Sub GetFileStream()
        Dim sp As New GraphAPICSLibPoc.SPList(client_id, tenant_id, client_secret)
        Dim json = Await sp.GetFileStream()
        Console.WriteLine(json)
        Console.ReadKey()

    End Sub

End Module

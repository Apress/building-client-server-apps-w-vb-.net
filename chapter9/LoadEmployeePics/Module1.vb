Option Explicit On 
Option Strict On

Imports System.Data.SqlClient
Imports System.IO

Module Module1

    Sub Main()
        Try
            Dim strCn As String = "Data Source=localhost;" & _
                "Initial Catalog=Northwind;Integrated Security=SSPI"
            Dim cn As New SqlConnection(strCn)

            Console.WriteLine("Updating Nancy Davilo's photo...")
            Dim cmd As New SqlCommand("Update Employees Set Photo = @BLOBData " _
                & "Where EmployeeID = 1", cn)
            Dim strBLOBFilePath As String = "nancy.bmp"
            Dim fsBLOBFile As New FileStream(strBLOBFilePath, FileMode.Open, FileAccess.Read)
            Dim bytBLOBData(Convert.ToInt32(fsBLOBFile.Length() - 1)) As Byte
            fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
            fsBLOBFile.Close()
            Dim prm As New SqlParameter("@BLOBData", SqlDbType.VarBinary, _
                bytBLOBData.Length, ParameterDirection.Input, False, _
                0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
            cmd.Parameters.Add(prm)
            cn.Open()
            cmd.ExecuteNonQuery()
            bytBLOBData = Nothing
            fsBLOBFile = Nothing
            prm = Nothing
            cmd = Nothing

            Console.WriteLine("Updating Andrew Fuller's photo...")
            cmd = New SqlCommand("Update Employees Set Photo = @BLOBData " _
                & "Where EmployeeID = 2", cn)
            strBLOBFilePath = "andrew.bmp"
            fsBLOBFile = New FileStream(strBLOBFilePath, FileMode.Open, FileAccess.Read)
            ReDim bytBLOBData(Convert.ToInt32(fsBLOBFile.Length() - 1))
            fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
            fsBLOBFile.Close()
            prm = New SqlParameter("@BLOBData", SqlDbType.VarBinary, _
                bytBLOBData.Length, ParameterDirection.Input, False, _
                0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
            cmd.Parameters.Add(prm)
            cmd.ExecuteNonQuery()
            bytBLOBData = Nothing
            fsBLOBFile = Nothing
            prm = Nothing
            cmd = Nothing

            Console.WriteLine("Updating Janet Leverling's photo...")
            cmd = New SqlCommand("Update Employees Set Photo = @BLOBData " _
                & "Where EmployeeID = 3", cn)
            strBLOBFilePath = "janet.bmp"
            fsBLOBFile = New FileStream(strBLOBFilePath, FileMode.Open, FileAccess.Read)
            ReDim bytBLOBData(Convert.ToInt32(fsBLOBFile.Length() - 1))
            fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
            fsBLOBFile.Close()
            prm = New SqlParameter("@BLOBData", SqlDbType.VarBinary, _
                bytBLOBData.Length, ParameterDirection.Input, False, _
                0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
            cmd.Parameters.Add(prm)
            cmd.ExecuteNonQuery()
            bytBLOBData = Nothing
            fsBLOBFile = Nothing
            prm = Nothing
            cmd = Nothing

            Console.WriteLine("Updating Margaret Peacock's photo...")
            cmd = New SqlCommand("Update Employees Set Photo = @BLOBData " _
                & "Where EmployeeID = 4", cn)
            strBLOBFilePath = "margaret.bmp"
            fsBLOBFile = New FileStream(strBLOBFilePath, FileMode.Open, FileAccess.Read)
            ReDim bytBLOBData(Convert.ToInt32(fsBLOBFile.Length() - 1))
            fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
            fsBLOBFile.Close()
            prm = New SqlParameter("@BLOBData", SqlDbType.VarBinary, _
                bytBLOBData.Length, ParameterDirection.Input, False, _
                0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
            cmd.Parameters.Add(prm)
            cmd.ExecuteNonQuery()
            bytBLOBData = Nothing
            fsBLOBFile = Nothing
            prm = Nothing
            cmd = Nothing

            Console.WriteLine("Updating Steven Buchanan's photo...")
            cmd = New SqlCommand("Update Employees Set Photo = @BLOBData " _
                & "Where EmployeeID = 5", cn)
            strBLOBFilePath = "steven.bmp"
            fsBLOBFile = New FileStream(strBLOBFilePath, FileMode.Open, FileAccess.Read)
            ReDim bytBLOBData(Convert.ToInt32(fsBLOBFile.Length() - 1))
            fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
            fsBLOBFile.Close()
            prm = New SqlParameter("@BLOBData", SqlDbType.VarBinary, _
                bytBLOBData.Length, ParameterDirection.Input, False, _
                0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
            cmd.Parameters.Add(prm)
            cmd.ExecuteNonQuery()
            bytBLOBData = Nothing
            fsBLOBFile = Nothing
            prm = Nothing
            cmd = Nothing

            Console.WriteLine("Updating Michael Suyama's photo...")
            cmd = New SqlCommand("Update Employees Set Photo = @BLOBData " _
                & "Where EmployeeID = 6", cn)
            strBLOBFilePath = "michael.bmp"
            fsBLOBFile = New FileStream(strBLOBFilePath, FileMode.Open, FileAccess.Read)
            ReDim bytBLOBData(Convert.ToInt32(fsBLOBFile.Length() - 1))
            fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
            fsBLOBFile.Close()
            prm = New SqlParameter("@BLOBData", SqlDbType.VarBinary, _
                bytBLOBData.Length, ParameterDirection.Input, False, _
                0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
            cmd.Parameters.Add(prm)
            cmd.ExecuteNonQuery()
            bytBLOBData = Nothing
            fsBLOBFile = Nothing
            prm = Nothing
            cmd = Nothing

            Console.WriteLine("Updating Robert King's photo...")
            cmd = New SqlCommand("Update Employees Set Photo = @BLOBData " _
                & "Where EmployeeID = 7", cn)
            strBLOBFilePath = "robert.bmp"
            fsBLOBFile = New FileStream(strBLOBFilePath, FileMode.Open, FileAccess.Read)
            ReDim bytBLOBData(Convert.ToInt32(fsBLOBFile.Length() - 1))
            fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
            fsBLOBFile.Close()
            prm = New SqlParameter("@BLOBData", SqlDbType.VarBinary, _
                bytBLOBData.Length, ParameterDirection.Input, False, _
                0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
            cmd.Parameters.Add(prm)
            cmd.ExecuteNonQuery()
            bytBLOBData = Nothing
            fsBLOBFile = Nothing
            prm = Nothing
            cmd = Nothing

            Console.WriteLine("Updating Laura Callahan's photo...")
            cmd = New SqlCommand("Update Employees Set Photo = @BLOBData " _
                & "Where EmployeeID = 8", cn)
            strBLOBFilePath = "laura.bmp"
            fsBLOBFile = New FileStream(strBLOBFilePath, FileMode.Open, FileAccess.Read)
            ReDim bytBLOBData(Convert.ToInt32(fsBLOBFile.Length() - 1))
            fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
            fsBLOBFile.Close()
            prm = New SqlParameter("@BLOBData", SqlDbType.VarBinary, _
                bytBLOBData.Length, ParameterDirection.Input, False, _
                0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
            cmd.Parameters.Add(prm)
            cmd.ExecuteNonQuery()
            bytBLOBData = Nothing
            fsBLOBFile = Nothing
            prm = Nothing
            cmd = Nothing

            Console.WriteLine("Updating Anne Dodsworth's photo...")
            cmd = New SqlCommand("Update Employees Set Photo = @BLOBData " _
                & "Where EmployeeID = 9", cn)
            strBLOBFilePath = "anne.bmp"
            fsBLOBFile = New FileStream(strBLOBFilePath, FileMode.Open, FileAccess.Read)
            ReDim bytBLOBData(Convert.ToInt32(fsBLOBFile.Length() - 1))
            fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
            fsBLOBFile.Close()
            prm = New SqlParameter("@BLOBData", SqlDbType.VarBinary, _
                bytBLOBData.Length, ParameterDirection.Input, False, _
                0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
            cmd.Parameters.Add(prm)
            cmd.ExecuteNonQuery()
            bytBLOBData = Nothing
            fsBLOBFile = Nothing
            prm = Nothing
            cmd = Nothing

            cn.Close()

            Console.WriteLine("All photo's successfully updated.")
        Catch exc As Exception
            Console.WriteLine(exc.Message)
            Console.WriteLine("An error occurred during the photo updates.")
        Finally
            Console.ReadLine()
        End Try
    End Sub

End Module

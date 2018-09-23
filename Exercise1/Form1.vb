Option Strict On
Option Explicit On
Public Class Form1
    Dim userName As String
    Dim password As String
    Private Sub LogIn_Click(sender As Object, e As EventArgs) Handles LogIn.Click
        userName = InputBox("Enter student number with prefix s")
        password = InputBox("Enter password")
        OracleConnection()
    End Sub
    Public Sub Status_Upd(ByVal pStr As String)
        lblStatus.Text = pStr
    End Sub
    Public Sub OracleConnection()
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection
        rvConn = CreatConnection()
        Try
            rvConn.Open()
            MessageBox.Show("Connect OK")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            MessageBox.Show("No Connection")
        Finally
            rvConn.Close()
        End Try
    End Sub
    Public Function GetConnectionString() As String
        Dim vConnStr As String
        vConnStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP) "
        vConnStr = vConnStr & "(HOST=feenix-oracle.swin.edu.au) (PORT=1521))"
        vConnStr = vConnStr & "(CONNECT_DATA=(SERVICE_NAME=dms)));"
        vConnStr = vConnStr & "User Id=" & userName & ";"
        vConnStr = vConnStr & "Password=" & password & ";"
        Return vConnStr
    End Function

    Public Function CreatConnection() As Oracle.ManagedDataAccess.Client.OracleConnection
        Dim rvConn As New Oracle.ManagedDataAccess.Client.OracleConnection
        rvConn.ConnectionString = GetConnectionString()
        Return rvConn
    End Function

    Private Sub TestTableCount_Click(sender As Object, e As EventArgs) Handles TestTableCount.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection
        rvConn = CreatConnection()
        Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
        Try
            rvCmd.Connection = rvConn
            rvCmd.CommandText = "SELECT COUNT(*) FROM TABS"
            rvConn.Open()
            rvCmd.CommandType = CommandType.Text
            Dim vStr As String
            vStr = rvCmd.ExecuteScalar.ToString
            MsgBox("Total number of Tables is " & vStr)
        Catch ex As Exception
            MessageBox.Show("ERROR OCCURRED " & ex.Message)
        Finally
            rvConn.Close()
        End Try
    End Sub

    Private Sub AddCustomer_Click(sender As Object, e As EventArgs) Handles AddCustomer.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            Dim vCustID As Double = Val(InputBox("Enter Customer ID"))
            Dim vCustName As String = InputBox("Enter Customer Name")
            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Add_Customer_Procedure(rvConn, rvTran, vCustID, vCustName)

            rvTran.Commit()
            Status_Upd("Add Customer Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Add Customers Rollback " & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub

    Sub Add_Customer_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                               ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction,
                               ByVal pCustID As Double,
                               ByVal pCustName As String)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poCustID As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poCustName As New Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "ADD_CUST_TO_DB"
            rvCmd.CommandType = CommandType.StoredProcedure

            poCustID.ParameterName = "pcustid"
            poCustID.DbType = DbType.Int16
            poCustID.Value = pCustID
            poCustID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poCustID)

            poCustName.ParameterName = "pcustname"
            poCustName.DbType = DbType.String
            poCustName.Value = pCustName
            poCustName.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poCustName)

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub Del_All_Cust_Click(sender As Object, e As EventArgs) Handles Del_All_Cust.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Dim numberCust As String
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            numberCust = Del_All_Cust_Procedure(rvConn, rvTran)

            rvTran.Commit()
            Status_Upd("Delete All Customers Committed, Number deleted: " & numberCust)

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Delete All Customers Rollback " & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub

    Private Function Del_All_Cust_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                               ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction) As String
        Dim rNumber As String
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poStr As Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandType = CommandType.StoredProcedure
            rvCmd.CommandText = "DELETE_ALL_CUSTOMERS_FROM_DB"

            poStr = New Oracle.ManagedDataAccess.Client.OracleParameter
            poStr.ParameterName = "pReturnNumber"
            poStr.DbType = DbType.Int16
            poStr.Direction = ParameterDirection.ReturnValue
            rvCmd.Parameters.Add(poStr)

            rvCmd.ExecuteNonQuery()
            rNumber = rvCmd.Parameters.Item("pReturnNumber").Value.ToString

        Catch ex As Exception
            Throw ex
        End Try
        Return rNumber
    End Function

    Private Sub Get_All_Cust_Click(sender As Object, e As EventArgs) Handles Get_All_Cust.Click
        Get_All_Cust_Procedure()
    End Sub

    Sub Get_All_Cust_Procedure()
        Try
            Dim connOracle As Oracle.ManagedDataAccess.Client.OracleConnection
            Dim commOracle As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim paramOracle As Oracle.ManagedDataAccess.Client.OracleParameter

            connOracle = CreatConnection()
            commOracle.Connection = connOracle
            commOracle.CommandType = CommandType.StoredProcedure
            commOracle.CommandText = "pckg_get_details.GetCusDetails"

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pReturnValue"
            paramOracle.OracleDbType = Oracle.ManagedDataAccess.Client.OracleDbType.RefCursor
            paramOracle.Direction = ParameterDirection.Output
            commOracle.Parameters.Add(paramOracle)
            connOracle.Open()

            Dim readerOracle As Oracle.ManagedDataAccess.Client.OracleDataReader
            readerOracle = commOracle.ExecuteReader()
            If readerOracle.HasRows = True Then
                lblStatus.Text = " "
                Do While readerOracle.Read()
                    MsgBox("Get All Customers - Customer ID: " & readerOracle("custid").ToString & " - Customer name: " & readerOracle("custname").ToString)
                Loop
            Else
                MessageBox.Show("No rows found")
            End If
            connOracle.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Get_All_Product_Click(sender As Object, e As EventArgs) Handles Get_All_Product.Click
        Get_All_Prod_Procedure()
    End Sub

    Sub Get_All_Prod_Procedure()
        Try
            Dim connOracle As Oracle.ManagedDataAccess.Client.OracleConnection
            Dim commOracle As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim paramOracle As Oracle.ManagedDataAccess.Client.OracleParameter

            connOracle = CreatConnection()
            commOracle.Connection = connOracle
            commOracle.CommandType = CommandType.StoredProcedure
            commOracle.CommandText = "pckg_get_details.GetProdDetails"

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pReturnValue"
            paramOracle.OracleDbType = Oracle.ManagedDataAccess.Client.OracleDbType.RefCursor
            paramOracle.Direction = ParameterDirection.Output
            commOracle.Parameters.Add(paramOracle)
            connOracle.Open()

            Dim readerOracle As Oracle.ManagedDataAccess.Client.OracleDataReader
            readerOracle = commOracle.ExecuteReader()
            If readerOracle.HasRows = True Then
                lblStatus.Text = "Test"
                Do While readerOracle.Read()
                    MsgBox("Get All Products - Product ID: " & readerOracle("prodid").ToString & " - Product name: " & readerOracle("prodname").ToString)
                Loop
            Else
                MessageBox.Show("No rows found")
            End If
            connOracle.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Add_Product_Click(sender As Object, e As EventArgs) Handles Add_Product.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            Dim vProdID As Double = Val(InputBox("Enter Product ID"))
            Dim vProdName As String = InputBox("Enter Product Name")
            Dim vProdPrice As Double = Val(InputBox("Enter Product Price"))
            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Add_Product_Procedure(rvConn, rvTran, vProdID, vProdName, vProdPrice)

            rvTran.Commit()
            Status_Upd("Add Product Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Add Product Rollback" & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub

    Sub Add_Product_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                           ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction,
                           ByVal pProdID As Double,
                           ByVal pProdName As String,
                           ByVal pProdPrice As Double)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poProdID As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poProdName As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poProdPrice As New Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "ADD_PROD_TO_DB"
            rvCmd.CommandType = CommandType.StoredProcedure

            poProdID.ParameterName = "pprodid"
            poProdID.DbType = DbType.Int16
            poProdID.Value = pProdID
            poProdID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poProdID)

            poProdName.ParameterName = "pprodname"
            poProdName.DbType = DbType.String
            poProdName.Value = pProdName
            poProdName.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poProdName)


            poProdPrice.ParameterName = "pprice"
            poProdPrice.DbType = DbType.Int16
            poProdPrice.Value = pProdPrice
            poProdPrice.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poProdPrice)

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Add_Simple_Sale_Click(sender As Object, e As EventArgs) Handles Add_Simple_Sale.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            Dim vCustID As Double = Val(InputBox("Enter Customer ID"))
            Dim vProdID As Double = Val(InputBox("Enter Product ID"))
            Dim vQty As Double = Val(InputBox("Enter Quantity"))

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Add_Simple_Sale_Procedure(rvConn, rvTran, vCustID, vProdID, vQty)

            rvTran.Commit()
            Status_Upd("Add Simple Sale Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Add Simple Sale Rollback " & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub
    Sub Add_Simple_Sale_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                       ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction,
                       ByVal pCustID As Double,
                       ByVal pProdId As Double,
                       ByVal pQty As Double)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poCustID As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poProdID As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poQty As New Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "ADD_SIMPLE_SALE_TO_DB"
            rvCmd.CommandType = CommandType.StoredProcedure

            poCustID.ParameterName = "pcustid"
            poCustID.DbType = DbType.Int16
            poCustID.Value = pCustID
            poCustID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poCustID)

            poProdID.ParameterName = "pprodid"
            poProdID.DbType = DbType.Int16
            poProdID.Value = pProdId
            poProdID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poProdID)

            poQty.ParameterName = "pqty"
            poQty.DbType = DbType.Int16
            poQty.Value = pQty
            poQty.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poQty)

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Get_Cust_String_Click(sender As Object, e As EventArgs) Handles Get_Cust_String.Click
        Get_Cust_String_Procedure()
    End Sub

    Private Sub Get_Cust_String_Procedure()
        Try
            Dim connOracle As Oracle.ManagedDataAccess.Client.OracleConnection
            Dim commOracle As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim paramOracle As Oracle.ManagedDataAccess.Client.OracleParameter
            Dim updateString As String

            connOracle = CreatConnection()
            commOracle.Connection = connOracle
            commOracle.CommandType = CommandType.StoredProcedure
            commOracle.CommandText = "GET_CUST_STRING_FROM_DB"

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pReturnValue"
            paramOracle.DbType = DbType.String
            paramOracle.Size = 200
            paramOracle.Direction = ParameterDirection.ReturnValue
            commOracle.Parameters.Add(paramOracle)

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pcustid"
            paramOracle.DbType = DbType.Int16
            paramOracle.Value = InputBox("Enter Customer ID")
            paramOracle.Direction = ParameterDirection.Input
            commOracle.Parameters.Add(paramOracle)

            updateString = "Customer String details: "
            Status_Upd(updateString)

            connOracle.Open()
            commOracle.ExecuteNonQuery()

            Dim returnString As String
            returnString = commOracle.Parameters.Item("pReturnValue").Value.ToString
            updateString = updateString + returnString

            Status_Upd(updateString)
            connOracle.Close()
        Catch ex As Exception
            Status_Upd(ex.Message)
        End Try
    End Sub

    Private Sub Get_Prod_String_Click(sender As Object, e As EventArgs) Handles Get_Prod_String.Click
        Get_Prod_String_Procedure()
    End Sub

    Private Sub Get_Prod_String_Procedure()
        Try
            Dim connOracle As Oracle.ManagedDataAccess.Client.OracleConnection
            Dim commOracle As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim paramOracle As Oracle.ManagedDataAccess.Client.OracleParameter
            Dim updateString As String

            connOracle = CreatConnection()
            commOracle.Connection = connOracle
            commOracle.CommandType = CommandType.StoredProcedure
            commOracle.CommandText = "GET_PROD_STRING_FROM_DB"

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pReturnValue"
            paramOracle.DbType = DbType.String
            paramOracle.Size = 200
            paramOracle.Direction = ParameterDirection.ReturnValue
            commOracle.Parameters.Add(paramOracle)

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pprodid"
            paramOracle.DbType = DbType.Int16
            paramOracle.Value = InputBox("Enter Product ID")
            paramOracle.Direction = ParameterDirection.Input
            commOracle.Parameters.Add(paramOracle)

            updateString = "Product String details: "
            Status_Upd(updateString)

            connOracle.Open()
            commOracle.ExecuteNonQuery()

            Dim returnString As String
            returnString = commOracle.Parameters.Item("pReturnValue").Value.ToString
            updateString = updateString + returnString

            Status_Upd(updateString)
            connOracle.Close()
        Catch ex As Exception
            Status_Upd(ex.Message)
        End Try
    End Sub

    Private Sub Sum_Cust_Sale_Click(sender As Object, e As EventArgs) Handles Sum_Cust_Sale.Click
        Sum_Cust_Sale_Procedure()
    End Sub
    Private Sub Sum_Cust_Sale_Procedure()
        Try
            Dim connOracle As Oracle.ManagedDataAccess.Client.OracleConnection
            Dim commOracle As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim paramOracle As Oracle.ManagedDataAccess.Client.OracleParameter
            Dim updateString As String

            connOracle = CreatConnection()
            commOracle.Connection = connOracle
            commOracle.CommandType = CommandType.StoredProcedure
            commOracle.CommandText = "SUM_CUST_SALESYTD"

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pReturnValue"
            paramOracle.DbType = DbType.Int16
            paramOracle.Direction = ParameterDirection.ReturnValue
            commOracle.Parameters.Add(paramOracle)

            updateString = "Sum Customer Sales YTD: "
            Status_Upd(updateString)

            connOracle.Open()
            commOracle.ExecuteNonQuery()

            Dim returnString As String
            returnString = commOracle.Parameters.Item("pReturnValue").Value.ToString
            updateString = updateString + returnString

            Status_Upd(updateString)
            connOracle.Close()
        Catch ex As Exception
            Status_Upd(ex.Message)
        End Try
    End Sub

    Private Sub Update_Cust_Status_Click(sender As Object, e As EventArgs) Handles Update_Cust_Status.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            Dim vCustID As Double = Val(InputBox("Enter Customer ID"))
            Dim vStatus As String = InputBox("Enter Customer Status")

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Update_Cust_Status_Procedure(rvConn, rvTran, vCustID, vStatus)

            rvTran.Commit()
            Status_Upd("Update Customer Status Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Update Customer Status Rollback " & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub
    Sub Update_Cust_Status_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                   ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction,
                   ByVal pCustID As Double,
                   ByVal pStatus As String)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poCustID As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poStatus As New Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "UPD_CUST_STATUS_IN_DB"
            rvCmd.CommandType = CommandType.StoredProcedure

            poCustID.ParameterName = "pcustid"
            poCustID.DbType = DbType.Int16
            poCustID.Value = pCustID
            poCustID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poCustID)

            poStatus.ParameterName = "pprodid"
            poStatus.DbType = DbType.String
            poStatus.Value = pStatus
            poStatus.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poStatus)

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Upd_Cust_SaleYTD_Click(sender As Object, e As EventArgs) Handles Upd_Cust_SaleYTD.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            Dim vCustID As Double = Val(InputBox("Enter Customer ID"))
            Dim vAmt As Double = Val(InputBox("Enter Amount"))

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Upd_Cust_SaleYTD_Procedure(rvConn, rvTran, vCustID, vAmt)

            rvTran.Commit()
            Status_Upd("Update Customer SaleYTD Committed")


        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Update Customer SaleYTD Rollback " & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub
    Sub Upd_Cust_SaleYTD_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
               ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction,
               ByVal pCustID As Double,
               ByVal pAmt As Double)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poCustID As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poAmt As New Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "UPD_CUST_SALESYTD_IN_DB"
            rvCmd.CommandType = CommandType.StoredProcedure

            poCustID.ParameterName = "pcustid"
            poCustID.DbType = DbType.Int16
            poCustID.Value = pCustID
            poCustID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poCustID)

            poAmt.ParameterName = "pprodid"
            poAmt.DbType = DbType.String
            poAmt.Value = pAmt
            poAmt.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poAmt)

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Upd_Prod_SalesYTD_Click(sender As Object, e As EventArgs) Handles Upd_Prod_SalesYTD.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            Dim vProdID As Double = Val(InputBox("Enter Product ID"))
            Dim vAmt As Double = Val(InputBox("Enter Amount"))

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Upd_Prod_SaleYTD_Procedure(rvConn, rvTran, vProdID, vAmt)

            rvTran.Commit()
            Status_Upd("Update Product SaleYTD Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Update Product SaleYTD  Rollback" & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub
    Sub Upd_Prod_SaleYTD_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
           ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction,
           ByVal pProdID As Double,
           ByVal pAmt As Double)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poProdID As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poAmt As New Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "UPD_PROD_SALESYTD_IN_DB"
            rvCmd.CommandType = CommandType.StoredProcedure

            poProdID.ParameterName = "pprodid"
            poProdID.DbType = DbType.Int16
            poProdID.Value = pProdID
            poProdID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poProdID)

            poAmt.ParameterName = "pprodid"
            poAmt.DbType = DbType.String
            poAmt.Value = pAmt
            poAmt.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poAmt)

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Add_Complex_Sale_Click(sender As Object, e As EventArgs) Handles Add_Complex_Sale.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            Dim vCustID As Double = Val(InputBox("Enter Customer ID"))
            Dim vProdID As Double = Val(InputBox("Enter Product ID"))
            Dim vQty As Double = Val(InputBox("Enter Quantity"))
            Dim vDate As String = InputBox("Enter date format YYYYMMDD")

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Add_Complex_Sale_Procedure(rvConn, rvTran, vCustID, vProdID, vQty, vDate)

            rvTran.Commit()
            Status_Upd("Add Complex Sale Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Add Complex Sale Rollback " & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub

    Sub Add_Complex_Sale_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                       ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction,
                       ByVal pCustID As Double,
                       ByVal pProdId As Double,
                       ByVal pQty As Double,
                       ByVal pDate As String)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poCustID As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poProdID As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poQty As New Oracle.ManagedDataAccess.Client.OracleParameter
            Dim poDate As New Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "ADD_COMPLEX_SALE_TO_DB"
            rvCmd.CommandType = CommandType.StoredProcedure

            poCustID.ParameterName = "pcustid"
            poCustID.DbType = DbType.Int16
            poCustID.Value = pCustID
            poCustID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poCustID)

            poProdID.ParameterName = "pprodid"
            poProdID.DbType = DbType.Int16
            poProdID.Value = pProdId
            poProdID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poProdID)

            poQty.ParameterName = "pqty"
            poQty.DbType = DbType.Int16
            poQty.Value = pQty
            poQty.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poQty)

            poDate.ParameterName = "pdate"
            poDate.DbType = DbType.String
            poDate.Value = pDate
            poDate.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poDate)

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Count_Prod_Sales_Click(sender As Object, e As EventArgs) Handles Count_Prod_Sales.Click
        Count_Prod_Sales_Procedure()
    End Sub

    Private Sub Count_Prod_Sales_Procedure()
        Try
            Dim connOracle As Oracle.ManagedDataAccess.Client.OracleConnection
            Dim commOracle As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim paramOracle As Oracle.ManagedDataAccess.Client.OracleParameter
            Dim updateString As String
            Dim vdays As Double = Val(InputBox("How many days ago?"))

            connOracle = CreatConnection()
            commOracle.Connection = connOracle
            commOracle.CommandType = CommandType.StoredProcedure
            commOracle.CommandText = "COUNT_PRODUCT_SALES_FROM_DB"

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pReturnValue"
            paramOracle.DbType = DbType.Int16
            paramOracle.Direction = ParameterDirection.ReturnValue
            commOracle.Parameters.Add(paramOracle)

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pday"
            paramOracle.DbType = DbType.Int16
            paramOracle.Value = vdays
            paramOracle.Direction = ParameterDirection.Input
            commOracle.Parameters.Add(paramOracle)

            updateString = "Counts sales within " & vdays.ToString & " days: "
            Status_Upd(updateString)

            connOracle.Open()
            commOracle.ExecuteNonQuery()

            Dim returnString As String
            returnString = commOracle.Parameters.Item("pReturnValue").Value.ToString
            updateString = updateString + returnString

            Status_Upd(updateString)
            connOracle.Close()
        Catch ex As Exception
            Status_Upd(ex.Message)
        End Try
    End Sub

    Private Sub Del_Sale_Click(sender As Object, e As EventArgs) Handles Del_Sale.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Del_Sale_Procedure(rvConn, rvTran)

            rvTran.Commit()
            Status_Upd("Delete Sale Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Delete Sale Rollback " & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub
    Sub Del_Sale_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                       ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poSaleID As New Oracle.ManagedDataAccess.Client.OracleParameter


            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "DELETE_SALE_FROM_DB"
            rvCmd.CommandType = CommandType.StoredProcedure

            poSaleID = New Oracle.ManagedDataAccess.Client.OracleParameter
            poSaleID.ParameterName = "pReturnValue"
            poSaleID.DbType = DbType.Int16
            poSaleID.Direction = ParameterDirection.ReturnValue
            rvCmd.Parameters.Add(poSaleID)

            rvCmd.ExecuteNonQuery()

            Dim returnString As String
            returnString = rvCmd.Parameters.Item("pReturnValue").Value.ToString

            Status_Upd("Smallest deleted SaleID is: " & returnString)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Del_All_Prod_Click(sender As Object, e As EventArgs) Handles Del_All_Prod.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Dim numberProd As String
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            numberProd = Del_All_Prod_Procedure(rvConn, rvTran)

            rvTran.Commit()
            Status_Upd("Delete All Products Committed, Number Product deleted: " & numberProd)

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Delete All Product Rollback" & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub
    Private Function Del_All_Prod_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                               ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction) As String

        Dim numberProd As String
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poStr As Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandType = CommandType.StoredProcedure
            rvCmd.CommandText = "DELETE_ALL_PRODUCTS_FROM_DB"

            poStr = New Oracle.ManagedDataAccess.Client.OracleParameter
            poStr.ParameterName = "pReturnNumber"
            poStr.DbType = DbType.Int16
            poStr.Direction = ParameterDirection.ReturnValue
            rvCmd.Parameters.Add(poStr)

            rvCmd.ExecuteNonQuery()
            numberProd = rvCmd.Parameters.Item("pReturnNumber").Value.ToString

        Catch ex As Exception
            Throw ex
        End Try
        Return numberProd
    End Function

    Private Sub Del_All_Sales_Click(sender As Object, e As EventArgs) Handles Del_All_Sales.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Try
            rvConn = CreatConnection()
            rvConn.Open()

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Del_All_Sales_Procedure(rvConn, rvTran)

            rvTran.Commit()
            Status_Upd("Delete All Sales Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Delete All Sales Rollback " & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub
    Sub Del_All_Sales_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                               ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandType = CommandType.StoredProcedure
            rvCmd.CommandText = "DELETE_ALL_SALES_FROM_DB"

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Del_Prod_Click(sender As Object, e As EventArgs) Handles Del_Prod.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Dim vProdID As Double = Val(InputBox("Enter Product ID"))

        Try
            rvConn = CreatConnection()
            rvConn.Open()

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Del_Prod_Procedure(rvConn, rvTran, vProdID)

            rvTran.Commit()
            Status_Upd("Delete Product with Id: " & vProdID & " Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Delete Product with Id: " & vProdID & " Rollback" & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub
    Sub Del_Prod_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                      ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction,
                      ByVal pProdId As Double)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poProdID As New Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "DELETE_PROD_FROM_DB"
            rvCmd.CommandType = CommandType.StoredProcedure

            poProdID.ParameterName = "pprodid"
            poProdID.DbType = DbType.Int16
            poProdID.Value = pProdId
            poProdID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poProdID)

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Del_Cust_Click(sender As Object, e As EventArgs) Handles Del_Cust.Click
        Dim rvConn As Oracle.ManagedDataAccess.Client.OracleConnection = Nothing
        Dim rvTran As Oracle.ManagedDataAccess.Client.OracleTransaction = Nothing
        Dim vCustID As Double = Val(InputBox("Enter Customer ID"))

        Try
            rvConn = CreatConnection()
            rvConn.Open()

            rvTran = rvConn.BeginTransaction(IsolationLevel.ReadCommitted)
            Del_Cust_Procedure(rvConn, rvTran, vCustID)

            rvTran.Commit()
            Status_Upd("Delete Customer with Id: " & vCustID & " Committed")

        Catch ex As Exception
            If rvTran Is Nothing Then
                MsgBox("Please log in")
                Exit Sub
            Else
                rvTran.Rollback()
                Status_Upd(ex.Message)
                Status_Upd("Delete Customer with Id: " & vCustID & " Rollback " & ex.Message)
            End If
        Finally
            rvConn.Close()
        End Try
    End Sub
    Sub Del_Cust_Procedure(ByVal prvConn As Oracle.ManagedDataAccess.Client.OracleConnection,
                      ByVal prvTran As Oracle.ManagedDataAccess.Client.OracleTransaction,
                      ByVal pCustId As Double)
        Try
            Dim rvCmd As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim poCustID As New Oracle.ManagedDataAccess.Client.OracleParameter

            rvCmd.Connection = prvConn
            rvCmd.Transaction = prvTran
            rvCmd.CommandText = "DELETE_CUSTOMER"
            rvCmd.CommandType = CommandType.StoredProcedure

            poCustID.ParameterName = "pcustid"
            poCustID.DbType = DbType.Int16
            poCustID.Value = pCustId
            poCustID.Direction = ParameterDirection.Input
            rvCmd.Parameters.Add(poCustID)

            rvCmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Sign_Out_Click(sender As Object, e As EventArgs) Handles Sign_Out.Click
        userName = ""
        password = ""
        MsgBox("Thank you!")
    End Sub

    Private Sub Get_All_Sales_Click(sender As Object, e As EventArgs) Handles Get_All_Sales.Click
        Get_All_Sales_Procedure()
    End Sub

    Sub Get_All_Sales_Procedure()
        Try
            Dim connOracle As Oracle.ManagedDataAccess.Client.OracleConnection
            Dim commOracle As New Oracle.ManagedDataAccess.Client.OracleCommand
            Dim paramOracle As Oracle.ManagedDataAccess.Client.OracleParameter

            connOracle = CreatConnection()
            commOracle.Connection = connOracle
            commOracle.CommandType = CommandType.StoredProcedure
            commOracle.CommandText = "pckg_get_details.GetSaleDetails"

            paramOracle = New Oracle.ManagedDataAccess.Client.OracleParameter
            paramOracle.ParameterName = "pReturnValue"
            paramOracle.OracleDbType = Oracle.ManagedDataAccess.Client.OracleDbType.RefCursor
            paramOracle.Direction = ParameterDirection.Output
            commOracle.Parameters.Add(paramOracle)
            connOracle.Open()

            Dim readerOracle As Oracle.ManagedDataAccess.Client.OracleDataReader
            readerOracle = commOracle.ExecuteReader()
            If readerOracle.HasRows = True Then
                lblStatus.Text = " "
                Do While readerOracle.Read()
                    MsgBox("Get All Sales - Sale ID: " & readerOracle("saleid").ToString & " Customer ID: " & readerOracle("custid").ToString & " - Product ID: " & readerOracle("prodid").ToString & " - QTY: " & readerOracle("qty").ToString & " - Date: " & readerOracle("saledate").ToString)
                Loop
            Else
                MessageBox.Show("No rows found")
            End If
            connOracle.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class

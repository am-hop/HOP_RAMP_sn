Imports System.IO
Imports System.Data.SQLite
Public Class MainForm

    Dim CurrentFilter As String

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With LvwBidList
            .Columns.Clear()
            .FullRowSelect = True
            .GridLines = True
            .View = View.Details
            .Columns.Add("No")                  '1
            .Columns.Add("US")                  '2
            .Columns.Add("Act.")              '3
            .Columns.Add("Bid ID")              '4
            .Columns.Add("Customer")            '5
            .Columns.Add("Full Bid Name")       '6
            .Columns.Add("BMT")                 '7
            .Columns.Add("COMM Owner")          '8
            .Columns.Add("Tier")                '9
            .Columns.Add("Rnd")                 '10
            .Columns.Add("Status")              '11
            .Columns.Add("%")                   '12
            .Columns.Add("Award")               '13
            .Columns.Add("Lead Region")         '14
            .Columns.Add("Lead GK")             '15
            .Columns.Add("AM GK")               '16
            .Columns.Add("Analyst")             '17
            .Columns.Add("Assigned")            '18
            .Columns.Add("Received")            '19
            .Columns.Add("Port Launch")         '20
        End With
        ClearAll()
        TxtBidID.Focus()
        LoadAllBids()
    End Sub

    Private Sub LvwBidList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LvwBidList.SelectedIndexChanged
        BtnSave.Enabled = False
        BtnUpdate.Enabled = True
        BtnDelete.Enabled = True
        BtnClear.Enabled = True
        BtnCreateFolderBidName.Enabled = True
        BtnCommitChanges.Enabled = True
        BtnExportNotesHTML.Enabled = True
        BtnExportNotesPlainText.Enabled = True
        If LvwBidList.SelectedItems.Count > 0 Then
            OpenCon()
            Try
                query = "SELECT * from rfq Where ID = '" & LvwBidList.SelectedItems(0).SubItems(0).Text & "'"
                cmd = New SQLiteCommand(query, con)
                dr = cmd.ExecuteReader
                While dr.Read
                    LblID.Text = dr.Item(0).ToString
                    ChkUSBid.Checked = dr.Item(1).ToString ' changed all chkboxes to .tostring
                    ChkBidActive.Checked = dr.Item(2).ToString ' changed all chkboxes to .tostring
                    TxtBidID.Text = dr.Item(3).ToString
                    TxtCustomer.Text = dr.Item(4).ToString
                    LblCustomerHeader.Text = dr.Item(4).ToString & " .:. " & dr.Item(10) & " .:. " & dr.Item(11) & "%"
                    TxtBidName.Text = dr.Item(5).ToString
                    TxtBMT.Text = dr.Item(6).ToString
                    TxtCO.Text = dr.Item(7).ToString
                    TxtTier.Text = dr.Item(8).ToString
                    TxtNoOfRounds.Text = dr.Item(9).ToString
                    TxtStatus.Text = dr.Item(10).ToString
                    TxtPercentComplete.Text = dr.Item(11).ToString
                    CboAwardStatus.Text = dr.Item(12).ToString
                    CboLeadRegion.Text = dr.Item(13).ToString
                    CboLeadGK.Text = dr.Item(14).ToString
                    CboAMGK.Text = dr.Item(15).ToString
                    CboAnalyst.Text = dr.Item(16).ToString
                    TxtBidAssigned.Text = dr.Item(17).ToString
                    TxtBidReceived.Text = dr.Item(18).ToString
                    TxtPortLaunch.Text = dr.Item(19).ToString
                    TxtR1_Launch.Text = dr.Item(20).ToString
                    TxtR1_InternalDue.Text = dr.Item(21).ToString
                    TxtR1_CustomerDue.Text = dr.Item(22).ToString
                    TxtR1_Submitted.Text = dr.Item(23).ToString
                    TxtR2_Received.Text = dr.Item(24).ToString
                    TxtR2_Launch.Text = dr.Item(25).ToString
                    TxtR2_InternalDue.Text = dr.Item(26).ToString
                    TxtR2_CustomerDue.Text = dr.Item(27).ToString
                    TxtR2_Submitted.Text = dr.Item(28).ToString
                    TxtR3_Received.Text = dr.Item(29).ToString
                    TxtR3_Launch.Text = dr.Item(30).ToString
                    TxtR3_InternalDue.Text = dr.Item(31).ToString
                    TxtR3_CustomerDue.Text = dr.Item(32).ToString
                    TxtR3_Submitted.Text = dr.Item(33).ToString
                    TxtRateValidity.Text = dr.Item(34).ToString
                    TxtDimFactor.Text = dr.Item(35).ToString
                    ChkStandardFuel.Checked = dr.Item(36).ToString 'changed all chkboxes to .tostring
                    TxtPickupDay.Text = dr.Item(37).ToString
                    RtbStrategy.Text = dr.Item(38).ToString
                    RtbQA.Text = dr.Item(39).ToString
                    RtbToDo.Text = dr.Item(40).ToString
                    RtbJournal.Text = dr.Item(41).ToString
                    ChkUpcoming.Checked = dr.Item(42).ToString ' changed all chkboxes to .tostring
                    'RadGreen.Text = dr.Item(43).ToString 'error shows convertion string G to type boolean is not valid - testing this one
                    LblFolderPath.Text = dr.Item(44).ToString
                    ToolStripStatusFolderPath.Text = dr.Item(44).ToString
                    LblOneNoteFolderPath.Text = dr.Item(45).ToString
                End While
            Catch ex As Exception
                MessageBox.Show(ex.Message & ", ListView function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Public Sub LoadAllBids()
        OpenCon()
        Try
            Dim filterstatus As String
            filterstatus = CboFilter.Text.ToString()
            Select Case filterstatus
                Case "Active"
                    CurrentFilter = "SELECT * FROM rfq where BidActive = 1"
                Case "Inactive"
                    CurrentFilter = "SELECT * FROM rfq where BidActive <> 1"
                Case "US Bids"
                    CurrentFilter = "SELECT * FROM rfq where USBid = 1"
                Case "Upcoming"
                    CurrentFilter = "SELECT * FROM rfq where Upcoming = 1"
                Case "All"
                    CurrentFilter = "SELECT * FROM rfq"
                Case "Pending Award"
                    CurrentFilter = "SELECT * FROM rfq where AwardStatus = 'Pending Award'"
                Case "Won"
                    CurrentFilter = "SELECT * FROM rfq where AwardStatus = 'Won'"
                Case "Lost"
                    CurrentFilter = "SELECT * FROM rfq where AwardStatus = 'Lost'"
            End Select
            query = CurrentFilter
            cmd = New SQLiteCommand(query, con)
            dr = cmd.ExecuteReader
            LvwBidList.Items.Clear()
            While dr.Read
                Dim myitems As ListViewItem = LvwBidList.Items.Add(dr.Item(0).ToString) 'ID (PK)
                myitems.SubItems.Add(1).Text = dr.Item(1).ToString ' to determine the name, count the column headers in the database startign from 0
                myitems.SubItems.Add(2).Text = dr.Item(2).ToString
                myitems.SubItems.Add(3).Text = dr.Item(3).ToString
                myitems.SubItems.Add(4).Text = dr.Item(4).ToString
                myitems.SubItems.Add(5).Text = dr.Item(5).ToString
                myitems.SubItems.Add(6).Text = dr.Item(6).ToString
                myitems.SubItems.Add(7).Text = dr.Item(7).ToString
                myitems.SubItems.Add(8).Text = dr.Item(8).ToString
                myitems.SubItems.Add(9).Text = dr.Item(9).ToString
                myitems.SubItems.Add(10).Text = dr.Item(10).ToString
                myitems.SubItems.Add(11).Text = dr.Item(11).ToString
                myitems.SubItems.Add(12).Text = dr.Item(12).ToString
                myitems.SubItems.Add(13).Text = dr.Item(13).ToString
                myitems.SubItems.Add(14).Text = dr.Item(14).ToString
                myitems.SubItems.Add(15).Text = dr.Item(15).ToString
                myitems.SubItems.Add(16).Text = dr.Item(16).ToString
                myitems.SubItems.Add(17).Text = dr.Item(17).ToString
                myitems.SubItems.Add(18).Text = dr.Item(18).ToString
                myitems.SubItems.Add(19).Text = dr.Item(19).ToString
                myitems.SubItems.Add(20).Text = dr.Item(20).ToString
                myitems.SubItems.Add(21).Text = dr.Item(21).ToString
                myitems.SubItems.Add(22).Text = dr.Item(22).ToString
                myitems.SubItems.Add(23).Text = dr.Item(23).ToString
                myitems.SubItems.Add(24).Text = dr.Item(24).ToString
                myitems.SubItems.Add(25).Text = dr.Item(25).ToString
                myitems.SubItems.Add(26).Text = dr.Item(26).ToString
                myitems.SubItems.Add(27).Text = dr.Item(27).ToString
                myitems.SubItems.Add(28).Text = dr.Item(28).ToString
                myitems.SubItems.Add(29).Text = dr.Item(29).ToString
                myitems.SubItems.Add(30).Text = dr.Item(30).ToString
                myitems.SubItems.Add(31).Text = dr.Item(31).ToString
                myitems.SubItems.Add(32).Text = dr.Item(32).ToString
                myitems.SubItems.Add(33).Text = dr.Item(33).ToString
                myitems.SubItems.Add(34).Text = dr.Item(34).ToString
                myitems.SubItems.Add(35).Text = dr.Item(35).ToString
                myitems.SubItems.Add(36).Text = dr.Item(36).ToString
                myitems.SubItems.Add(37).Text = dr.Item(37).ToString
                myitems.SubItems.Add(38).Text = dr.Item(38).ToString
                myitems.SubItems.Add(39).Text = dr.Item(39).ToString
                myitems.SubItems.Add(40).Text = dr.Item(40).ToString
                myitems.SubItems.Add(41).Text = dr.Item(41).ToString
                myitems.SubItems.Add(42).Text = dr.Item(42).ToString
                'myitems.SubItems.Add(43).Text = dr.Item(43).ToString 'TrafficLight
                myitems.SubItems.Add(44).Text = dr.Item(44).ToString
                myitems.SubItems.Add(44).Text = dr.Item(45).ToString
            End While
            LvwBidList.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", LoadAllBids function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        BtnDelete.Enabled = False
        BtnCreateFolderBidName.Enabled = False
        BtnCommitChanges.Enabled = False
        BtnExportNotesHTML.Enabled = False
        BtnExportNotesPlainText.Enabled = False
        CountActiveBids()
        CountUpcoming()
        CountPendingAward()
        CountWon()
        CountLost()
    End Sub

    Private Sub BtnApplyFilter_Click(sender As Object, e As EventArgs) Handles BtnApplyFilter.Click
        LoadAllBids()
    End Sub

    Public Sub ClearAll()
        For Each gb As GroupBox In Me.Controls.OfType(Of GroupBox)()
            For Each tb As TextBox In gb.Controls.OfType(Of TextBox)()
                tb.Clear()
            Next
        Next
        RtbJournal.ResetText()
        RtbQA.ResetText()
        RtbStrategy.ResetText()
        RtbToDo.ResetText()
        RadCustomer.Checked = True
        BtnDelete.Enabled = False
        BtnCreateFolderBidName.Enabled = False
        BtnCommitChanges.Enabled = False
        BtnExportNotesHTML.Enabled = False
        BtnExportNotesPlainText.Enabled = False
        ChkBidActive.Checked = True
        ChkStandardFuel.Checked = False
        ChkUSBid.Checked = False
        LvwBidList.Items.Clear()
        TxtSearch.Text = String.Empty
        LblCustomerHeader.Text = String.Empty
        LblID.Text = String.Empty
        LblCount.Text = String.Empty
        CboAMGK.Text = String.Empty
        CboLeadGK.Text = String.Empty
        CboAnalyst.Text = String.Empty
        CboAwardStatus.Text = String.Empty
        CboLeadRegion.Text = String.Empty
        LblFolderPath.Text = String.Empty
        LblOneNoteFolderPath.Text = String.Empty
        LoadAllBids()
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        OpenCon()
        Try
            query = "INSERT INTO rfq (USBid,BidActive,BidID,Customer,BidName,BMT,CO,Tier,NoOfRounds,Status,PercentComplete,
                    AwardStatus,LeadRegion,LeadGK,AMGK,Analyst,BidAssigned,BidReceived,PortLaunch,R1_Launch,R1_InternalDue,
                    R1_CustomerDue,	R1_Submitted,R2_Received,R2_Launch,R2_InternalDue,R2_CustomerDue,R2_Submitted,
                    R3_Received,R3_Launch,R3_InternalDue,R3_CustomerDue,R3_Submitted,RateValidity,DimFactor,StandardFuel,
                    PickupDay,Strategy,QA,ToDo,Journal,Upcoming,FolderPath,OneNoteFolderPath) VALUES 
                    (@USBid,@BidActive,@BidID,@Customer,@BidName,@BMT,@CO,@Tier,@NoOfRounds,@Status,@PercentComplete,@AwardStatus,
                    @LeadRegion,@LeadGK,@AMGK,@Analyst,@BidAssigned,@BidReceived,@PortLaunch,@R1_Launch,@R1_InternalDue,@R1_CustomerDue,
                    @R1_Submitted,@R2_Received,@R2_Launch,@R2_InternalDue,@R2_CustomerDue,@R2_Submitted,@R3_Received,@R3_Launch,
                    @R3_InternalDue,@R3_CustomerDue,@R3_Submitted,@RateValidity,@DimFactor,@StandardFuel,@PickupDay,@Strategy,
                    @QA,@ToDo,@Journal,@Upcoming,@FolderPath,@OneNoteFolderPath)"
            cmd = New SQLiteCommand(query, con)
            With cmd
                .Parameters.AddWithValue("@USBid", ChkUSBid.Checked)
                .Parameters.AddWithValue("@BidActive", ChkBidActive.Checked)
                .Parameters.AddWithValue("@BidID", TxtBidID.Text.Trim)
                .Parameters.AddWithValue("@Customer", TxtCustomer.Text.Trim)
                .Parameters.AddWithValue("@BidName", TxtBidName.Text.Trim)
                .Parameters.AddWithValue("@BMT", TxtBMT.Text.Trim)
                .Parameters.AddWithValue("@CO", TxtCO.Text.Trim)
                .Parameters.AddWithValue("@Tier", TxtTier.Text.Trim)
                .Parameters.AddWithValue("@NoOfRounds", TxtNoOfRounds.Text.Trim)
                .Parameters.AddWithValue("@Status", TxtStatus.Text.Trim)
                .Parameters.AddWithValue("@PercentComplete", TxtPercentComplete.Text.Trim)
                .Parameters.AddWithValue("@AwardStatus", CboAwardStatus.Text.Trim)
                .Parameters.AddWithValue("@LeadRegion", CboLeadRegion.Text.Trim)
                .Parameters.AddWithValue("@LeadGK", CboLeadGK.Text.Trim)
                .Parameters.AddWithValue("@AMGK", CboAMGK.Text.Trim)
                .Parameters.AddWithValue("@Analyst", CboAnalyst.Text.Trim)
                .Parameters.AddWithValue("@BidAssigned", TxtBidAssigned.Text.Trim)
                .Parameters.AddWithValue("@BidReceived", TxtBidReceived.Text.Trim)
                .Parameters.AddWithValue("@PortLaunch", TxtPortLaunch.Text.Trim)
                .Parameters.AddWithValue("@R1_Launch", TxtR1_Launch.Text.Trim)
                .Parameters.AddWithValue("@R1_InternalDue", TxtR1_InternalDue.Text.Trim)
                .Parameters.AddWithValue("@R1_CustomerDue", TxtR1_CustomerDue.Text.Trim)
                .Parameters.AddWithValue("@R1_Submitted", TxtR1_Submitted.Text.Trim)
                .Parameters.AddWithValue("@R2_Received", TxtR2_Received.Text.Trim)
                .Parameters.AddWithValue("@R2_Launch", TxtR2_Launch.Text.Trim)
                .Parameters.AddWithValue("@R2_InternalDue", TxtR2_InternalDue.Text.Trim)
                .Parameters.AddWithValue("@R2_CustomerDue", TxtR2_CustomerDue.Text.Trim)
                .Parameters.AddWithValue("@R2_Submitted", TxtR2_Submitted.Text.Trim)
                .Parameters.AddWithValue("@R3_Received", TxtR3_Received.Text.Trim)
                .Parameters.AddWithValue("@R3_Launch", TxtR3_Launch.Text.Trim)
                .Parameters.AddWithValue("@R3_InternalDue", TxtR3_InternalDue.Text.Trim)
                .Parameters.AddWithValue("@R3_CustomerDue", TxtR3_CustomerDue.Text.Trim)
                .Parameters.AddWithValue("@R3_Submitted", TxtR3_Submitted.Text.Trim)
                .Parameters.AddWithValue("@RateValidity", TxtRateValidity.Text.Trim)
                .Parameters.AddWithValue("@DimFactor", TxtDimFactor.Text.Trim)
                .Parameters.AddWithValue("@StandardFuel", ChkStandardFuel.Checked)
                .Parameters.AddWithValue("@PickupDay", TxtPickupDay.Text.Trim)
                .Parameters.AddWithValue("@Strategy", RtbStrategy.Text.Trim)
                .Parameters.AddWithValue("@QA", RtbQA.Text.Trim)
                .Parameters.AddWithValue("@ToDo", RtbToDo.Text.Trim)
                .Parameters.AddWithValue("@Journal", RtbJournal.Text.Trim)
                .Parameters.AddWithValue("@Upcoming", ChkUpcoming.Checked)
                '.Parameters.AddWithValue("@TrafficLight", RadGreen.Checked.ToString)
                .Parameters.AddWithValue("@FolderPath", LblFolderPath.Text.Trim)
                .Parameters.AddWithValue("@OneNoteFolderPath", LblOneNoteFolderPath.Text.Trim)
                .ExecuteNonQuery()
            End With
            Dim SavedMessage As String = "New Bid has been successfully added" & " [" & Now & "]"
            ToolStripStatusSAVED.Text = SavedMessage
            LoadAllBids()
            ClearAll()
            LblTimestamp.Text = Now
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", SAVE function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnUpdate_Click(sender As Object, e As EventArgs) Handles BtnUpdate.Click
        OpenCon()
        Try
            query = "UPDATE rfq SET USBid=@USBid,BidActive=@BidActive,BidID=@BidID,Customer=@Customer,
                BidName=@BidName,BMT=@BMT,CO=@CO,Tier=@Tier,NoOfRounds=@NoOfRounds,Status=@Status,PercentComplete=@PercentComplete,
                AwardStatus=@AwardStatus,LeadRegion=@LeadRegion,LeadGK=@LeadGK,AMGK=@AMGK,Analyst=@Analyst,
                BidAssigned=@BidAssigned,BidReceived=@BidReceived,PortLaunch=@PortLaunch,
                R1_Launch=@R1_Launch,R1_InternalDue=@R1_InternalDue,
                R1_CustomerDue=@R1_CustomerDue,R1_Submitted=@R1_Submitted,
                R2_Received=@R2_Received,R2_Launch=@R2_Launch,R2_InternalDue=@R2_InternalDue,R2_CustomerDue=@R2_CustomerDue,R2_Submitted=@R2_Submitted,
                R3_Received=@R3_Received,R3_Launch=@R3_Launch,R3_InternalDue=@R3_InternalDue,R3_CustomerDue=@R3_CustomerDue,R3_Submitted=@R3_Submitted,
                RateValidity=@RateValidity,
                DimFactor=@DimFactor,
                StandardFuel=@StandardFuel,
                PickupDay=@PickupDay,
                Strategy=@Strategy,QA=@QA,ToDo=@ToDo,Journal=@Journal,Upcoming=@Upcoming,FolderPath=@FolderPath,OneNoteFolderPath=@OneNoteFolderPath 
                    WHERE ID= '" & LblID.Text & "'"
            cmd = New SQLiteCommand(query, con)
            With cmd
                .Parameters.AddWithValue("@USBid", ChkUSBid.Checked)
                .Parameters.AddWithValue("@BidActive", ChkBidActive.Checked)
                .Parameters.AddWithValue("@BidID", TxtBidID.Text.Trim)
                .Parameters.AddWithValue("@Customer", TxtCustomer.Text.Trim)
                .Parameters.AddWithValue("@BidName", TxtBidName.Text.Trim)
                .Parameters.AddWithValue("@BMT", TxtBMT.Text.Trim)
                .Parameters.AddWithValue("@CO", TxtCO.Text.Trim)
                .Parameters.AddWithValue("@Tier", TxtTier.Text.Trim)
                .Parameters.AddWithValue("@NoOfRounds", TxtNoOfRounds.Text.Trim)
                .Parameters.AddWithValue("@Status", TxtStatus.Text.Trim)
                .Parameters.AddWithValue("@PercentComplete", TxtPercentComplete.Text.Trim)
                .Parameters.AddWithValue("@AwardStatus", CboAwardStatus.Text.Trim)
                .Parameters.AddWithValue("@LeadRegion", CboLeadRegion.Text.Trim)
                .Parameters.AddWithValue("@LeadGK", CboLeadGK.Text.Trim)
                .Parameters.AddWithValue("@AMGK", CboAMGK.Text.Trim)
                .Parameters.AddWithValue("@Analyst", CboAnalyst.Text.Trim)
                .Parameters.AddWithValue("@BidAssigned", TxtBidAssigned.Text.Trim)
                .Parameters.AddWithValue("@BidReceived", TxtBidReceived.Text.Trim)
                .Parameters.AddWithValue("@PortLaunch", TxtPortLaunch.Text.Trim)
                .Parameters.AddWithValue("@R1_Launch", TxtR1_Launch.Text.Trim)
                .Parameters.AddWithValue("@R1_InternalDue", TxtR1_InternalDue.Text.Trim)
                .Parameters.AddWithValue("@R1_CustomerDue", TxtR1_CustomerDue.Text.Trim)
                .Parameters.AddWithValue("@R1_Submitted", TxtR1_Submitted.Text.Trim)
                .Parameters.AddWithValue("@R2_Received", TxtR2_Received.Text.Trim)
                .Parameters.AddWithValue("@R2_Launch", TxtR2_Launch.Text.Trim)
                .Parameters.AddWithValue("@R2_InternalDue", TxtR2_InternalDue.Text.Trim)
                .Parameters.AddWithValue("@R2_CustomerDue", TxtR2_CustomerDue.Text.Trim)
                .Parameters.AddWithValue("@R2_Submitted", TxtR2_Submitted.Text.Trim)
                .Parameters.AddWithValue("@R3_Received", TxtR3_Received.Text.Trim)
                .Parameters.AddWithValue("@R3_Launch", TxtR3_Launch.Text.Trim)
                .Parameters.AddWithValue("@R3_InternalDue", TxtR3_InternalDue.Text.Trim)
                .Parameters.AddWithValue("@R3_CustomerDue", TxtR3_CustomerDue.Text.Trim)
                .Parameters.AddWithValue("@R3_Submitted", TxtR3_Submitted.Text.Trim)
                .Parameters.AddWithValue("@RateValidity", TxtRateValidity.Text.Trim)
                .Parameters.AddWithValue("@DimFactor", TxtDimFactor.Text.Trim)
                .Parameters.AddWithValue("@StandardFuel", ChkStandardFuel.Checked)
                .Parameters.AddWithValue("@PickupDay", TxtPickupDay.Text.Trim)
                .Parameters.AddWithValue("@Strategy", RtbStrategy.Text.Trim)
                .Parameters.AddWithValue("@QA", RtbQA.Text.Trim)
                .Parameters.AddWithValue("@ToDo", RtbToDo.Text.Trim)
                .Parameters.AddWithValue("@Journal", RtbJournal.Text.Trim)
                .Parameters.AddWithValue("@Upcoming", ChkUpcoming.Checked)
                '.Parameters.AddWithValue("@TrafficLight", RadGreen.Checked.ToString)
                .Parameters.AddWithValue("@FolderPath", LblFolderPath.Text.Trim)
                .Parameters.AddWithValue("@OneNoteFolderPath", LblOneNoteFolderPath.Text.Trim)
                .ExecuteNonQuery()
            End With
            Dim UpdatedMessage As String = "Bid has been successfully updated" & " [" & Now & "]"
            ToolStripStatusSAVED.Text = UpdatedMessage
            LoadAllBids()
            ClearAll()
            BtnUpdate.Enabled = False
            BtnSave.Enabled = True
            TxtBidID.Focus()
            LblID.Text = String.Empty
            LblTimestamp.Text = Now
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", UPDATE function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        ClearAll()
        BtnUpdate.Enabled = False
        BtnSave.Enabled = True
        TxtBidID.Focus()
        LblID.Text = String.Empty
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs) Handles BtnDelete.Click
        If LvwBidList.SelectedItems.Count = 0 Then
            MessageBox.Show("Please select a customer in the list to delete", "Select item", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            LvwBidList.Focus()
        Else
            OpenCon()
            Try
                query = "DELETE FROM rfq WHERE id = '" & LvwBidList.SelectedItems(0).SubItems(0).Text & "'"
                cmd = New SQLiteCommand(query, con)
                cmd.ExecuteReader()
                Dim DeletedMessage As String = "Bid has been successfully deleted" & " [" & Now & "]"
                ToolStripStatusSAVED.Text = DeletedMessage
            Catch ex As Exception
                MessageBox.Show(ex.Message & ", Delete function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
        ClearAll()
        LoadAllBids()
        TxtBidID.Focus()
        LblTimestamp.Text = Now
    End Sub

    Private Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        OpenCon()
        Try
            query = "SELECT * FROM rfq "
            If RadCustomer.Checked = True Then
                query += "WHERE BIDNAME LIKE '%" & TxtSearch.Text.Trim & "%'"
            ElseIf RadStatus.Checked = True Then
                query += "WHERE STATUS LIKE '%" & TxtSearch.Text.Trim & "%'"
            End If
            cmd = New SQLiteCommand(query, con)
            dr = cmd.ExecuteReader
            LvwBidList.Items.Clear()
            While dr.Read
                Dim myitems As ListViewItem = LvwBidList.Items.Add(dr.Item(0).ToString)
                myitems.SubItems.Add(1).Text = dr.Item(1).ToString ' to determine the name, count the column headers in the database startign from 0
                myitems.SubItems.Add(2).Text = dr.Item(2).ToString
                myitems.SubItems.Add(3).Text = dr.Item(3).ToString
                myitems.SubItems.Add(4).Text = dr.Item(4).ToString
                myitems.SubItems.Add(5).Text = dr.Item(5).ToString
                myitems.SubItems.Add(6).Text = dr.Item(6).ToString
                myitems.SubItems.Add(7).Text = dr.Item(7).ToString
                myitems.SubItems.Add(8).Text = dr.Item(8).ToString
                myitems.SubItems.Add(9).Text = dr.Item(9).ToString
                myitems.SubItems.Add(10).Text = dr.Item(10).ToString
                myitems.SubItems.Add(11).Text = dr.Item(11).ToString
                myitems.SubItems.Add(12).Text = dr.Item(12).ToString
                myitems.SubItems.Add(13).Text = dr.Item(13).ToString
                myitems.SubItems.Add(14).Text = dr.Item(14).ToString
                myitems.SubItems.Add(15).Text = dr.Item(15).ToString
                myitems.SubItems.Add(16).Text = dr.Item(16).ToString
                myitems.SubItems.Add(17).Text = dr.Item(17).ToString
                myitems.SubItems.Add(18).Text = dr.Item(18).ToString
                myitems.SubItems.Add(19).Text = dr.Item(19).ToString
                myitems.SubItems.Add(20).Text = dr.Item(20).ToString
                myitems.SubItems.Add(21).Text = dr.Item(21).ToString
                myitems.SubItems.Add(22).Text = dr.Item(22).ToString
                myitems.SubItems.Add(23).Text = dr.Item(23).ToString
                myitems.SubItems.Add(24).Text = dr.Item(24).ToString
                myitems.SubItems.Add(25).Text = dr.Item(25).ToString
                myitems.SubItems.Add(26).Text = dr.Item(26).ToString
                myitems.SubItems.Add(27).Text = dr.Item(27).ToString
                myitems.SubItems.Add(28).Text = dr.Item(28).ToString
                myitems.SubItems.Add(29).Text = dr.Item(29).ToString
                myitems.SubItems.Add(30).Text = dr.Item(30).ToString
                myitems.SubItems.Add(31).Text = dr.Item(31).ToString
                myitems.SubItems.Add(32).Text = dr.Item(32).ToString
                myitems.SubItems.Add(33).Text = dr.Item(33).ToString
                myitems.SubItems.Add(34).Text = dr.Item(34).ToString
                myitems.SubItems.Add(35).Text = dr.Item(35).ToString
                myitems.SubItems.Add(36).Text = dr.Item(36).ToString
                myitems.SubItems.Add(37).Text = dr.Item(37).ToString
                myitems.SubItems.Add(38).Text = dr.Item(38).ToString
                myitems.SubItems.Add(39).Text = dr.Item(39).ToString
                myitems.SubItems.Add(40).Text = dr.Item(40).ToString
                myitems.SubItems.Add(41).Text = dr.Item(41).ToString
                myitems.SubItems.Add(42).Text = dr.Item(42).ToString
                'myitems.SubItems.Add(43).Text = dr.Item(43).ToString
                myitems.SubItems.Add(44).Text = dr.Item(44).ToString
                myitems.SubItems.Add(44).Text = dr.Item(45).ToString
            End While
            LblCount.Text = LvwBidList.Items.Count & " record(s) found!" & " [" & Now & "]"
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", SEARCH function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnReset_Click(sender As Object, e As EventArgs) Handles BtnReset.Click
        ClearAll()
        BtnUpdate.Enabled = False
        BtnSave.Enabled = True
        TxtBidID.Focus()
        LblID.Text = String.Empty
        LblCount.Text = String.Empty
    End Sub

    Private Sub BtnCreateFolderBidName_Click(sender As Object, e As EventArgs) Handles BtnCreateFolderBidName.Click
        Dim FolderBidName As String = ("\" & TxtBidName.Text).ToString
        Dim fullpath As String = (My.Computer.FileSystem.SpecialDirectories.MyDocuments & FolderBidName).ToString
        Directory.CreateDirectory(fullpath)
        Process.Start(fullpath)
        LblFolderPath.Text = fullpath.ToString
        ToolStripStatusFolderPath.Text = LblFolderPath.Text.ToString
    End Sub

    Private Sub BtnCreateFolderCustomerName_Click(sender As Object, e As EventArgs) Handles BtnCreateFolderCustomerName.Click
        '======== CURRENTLY INVISIBLE ==================
        Dim FolderCustomerName As String = ("\" & TxtCustomer.Text).ToString
        Dim fullpath As String = (My.Computer.FileSystem.SpecialDirectories.MyDocuments & FolderCustomerName).ToString
        Directory.CreateDirectory(fullpath)
        Process.Start(fullpath)
        LblFolderPath.Text = fullpath.ToString
    End Sub

    Private Sub BtnCommitChanges_Click(sender As Object, e As EventArgs) Handles BtnCommitChanges.Click
        OpenCon()
        Try
            query = "UPDATE rfq SET USBid=@USBid,BidActive=@BidActive,BidID=@BidID,Customer=@Customer,
                BidName=@BidName,BMT=@BMT,CO=@CO,Tier=@Tier,NoOfRounds=@NoOfRounds,Status=@Status,PercentComplete=@PercentComplete,
                AwardStatus=@AwardStatus,LeadRegion=@LeadRegion,LeadGK=@LeadGK,AMGK=@AMGK,Analyst=@Analyst,
                BidAssigned=@BidAssigned,BidReceived=@BidReceived,PortLaunch=@PortLaunch,
                R1_Launch=@R1_Launch,R1_InternalDue=@R1_InternalDue,
                R1_CustomerDue=@R1_CustomerDue,R1_Submitted=@R1_Submitted,
                R2_Received=@R2_Received,R2_Launch=@R2_Launch,R2_InternalDue=@R2_InternalDue,R2_CustomerDue=@R2_CustomerDue,R2_Submitted=@R2_Submitted,
                R3_Received=@R3_Received,R3_Launch=@R3_Launch,R3_InternalDue=@R3_InternalDue,R3_CustomerDue=@R3_CustomerDue,R3_Submitted=@R3_Submitted,
                RateValidity=@RateValidity,
                DimFactor=@DimFactor,
                StandardFuel=@StandardFuel,
                PickupDay=@PickupDay,
                Strategy=@Strategy,QA=@QA,ToDo=@ToDo,Journal=@Journal,Upcoming=@Upcoming,FolderPath=@FolderPath,OneNoteFolderPath=@OneNoteFolderPath 
                    WHERE ID= '" & LblID.Text & "'"
            cmd = New SQLiteCommand(query, con)
            With cmd
                .Parameters.AddWithValue("@USBid", ChkUSBid.Checked)
                .Parameters.AddWithValue("@BidActive", ChkBidActive.Checked)
                .Parameters.AddWithValue("@BidID", TxtBidID.Text.Trim)
                .Parameters.AddWithValue("@Customer", TxtCustomer.Text.Trim)
                .Parameters.AddWithValue("@BidName", TxtBidName.Text.Trim)
                .Parameters.AddWithValue("@BMT", TxtBMT.Text.Trim)
                .Parameters.AddWithValue("@CO", TxtCO.Text.Trim)
                .Parameters.AddWithValue("@Tier", TxtTier.Text.Trim)
                .Parameters.AddWithValue("@NoOfRounds", TxtNoOfRounds.Text.Trim)
                .Parameters.AddWithValue("@Status", TxtStatus.Text.Trim)
                .Parameters.AddWithValue("@PercentComplete", TxtPercentComplete.Text.Trim)
                .Parameters.AddWithValue("@AwardStatus", CboAwardStatus.Text.Trim)
                .Parameters.AddWithValue("@LeadRegion", CboLeadRegion.Text.Trim)
                .Parameters.AddWithValue("@LeadGK", CboLeadGK.Text.Trim)
                .Parameters.AddWithValue("@AMGK", CboAMGK.Text.Trim)
                .Parameters.AddWithValue("@Analyst", CboAnalyst.Text.Trim)
                .Parameters.AddWithValue("@BidAssigned", TxtBidAssigned.Text.Trim)
                .Parameters.AddWithValue("@BidReceived", TxtBidReceived.Text.Trim)
                .Parameters.AddWithValue("@PortLaunch", TxtPortLaunch.Text.Trim)
                .Parameters.AddWithValue("@R1_Launch", TxtR1_Launch.Text.Trim)
                .Parameters.AddWithValue("@R1_InternalDue", TxtR1_InternalDue.Text.Trim)
                .Parameters.AddWithValue("@R1_CustomerDue", TxtR1_CustomerDue.Text.Trim)
                .Parameters.AddWithValue("@R1_Submitted", TxtR1_Submitted.Text.Trim)
                .Parameters.AddWithValue("@R2_Received", TxtR2_Received.Text.Trim)
                .Parameters.AddWithValue("@R2_Launch", TxtR2_Launch.Text.Trim)
                .Parameters.AddWithValue("@R2_InternalDue", TxtR2_InternalDue.Text.Trim)
                .Parameters.AddWithValue("@R2_CustomerDue", TxtR2_CustomerDue.Text.Trim)
                .Parameters.AddWithValue("@R2_Submitted", TxtR2_Submitted.Text.Trim)
                .Parameters.AddWithValue("@R3_Received", TxtR3_Received.Text.Trim)
                .Parameters.AddWithValue("@R3_Launch", TxtR3_Launch.Text.Trim)
                .Parameters.AddWithValue("@R3_InternalDue", TxtR3_InternalDue.Text.Trim)
                .Parameters.AddWithValue("@R3_CustomerDue", TxtR3_CustomerDue.Text.Trim)
                .Parameters.AddWithValue("@R3_Submitted", TxtR3_Submitted.Text.Trim)
                .Parameters.AddWithValue("@RateValidity", TxtRateValidity.Text.Trim)
                .Parameters.AddWithValue("@DimFactor", TxtDimFactor.Text.Trim)
                .Parameters.AddWithValue("@StandardFuel", ChkStandardFuel.Checked)
                .Parameters.AddWithValue("@PickupDay", TxtPickupDay.Text.Trim)
                .Parameters.AddWithValue("@Strategy", RtbStrategy.Text.Trim)
                .Parameters.AddWithValue("@QA", RtbQA.Text.Trim)
                .Parameters.AddWithValue("@ToDo", RtbToDo.Text.Trim)
                .Parameters.AddWithValue("@Journal", RtbJournal.Text.Trim)
                .Parameters.AddWithValue("@Upcoming", ChkUpcoming.Checked)
                '.Parameters.AddWithValue("@TrafficLight", RadGreen.Checked.ToString)
                .Parameters.AddWithValue("@FolderPath", LblFolderPath.Text.Trim)
                .Parameters.AddWithValue("@OneNoteFolderPath", LblOneNoteFolderPath.Text.Trim)
                .ExecuteNonQuery()
            End With
            Dim UpdatedMessage As String = "Changes have been successfully committed" & " [" & Now & "]"
            ToolStripStatusSAVED.Text = UpdatedMessage
            BtnUpdate.Enabled = False
            BtnSave.Enabled = False
            RtbJournal.Focus()
            LblTimestamp.Text = Now
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", COMMIT function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MenuBidLog_Click(sender As Object, e As EventArgs) Handles MenuBidLog.Click
        Dim url As String = "https://isharenew.dhl.com/sites/RACC_AMER/Lists/Bid%20Log/Standard%20Bid%20Log.aspx"
        Process.Start(url)
    End Sub

    Private Sub MenuCACCcontactlist_Click(sender As Object, e As EventArgs) Handles MenuCACCcontactlist.Click
        Dim url As String = "https://isharenew.dhl.com/sites/gho_afr_forum/Lists/HoP%20Contact%20List/AllItems.aspx"
        Process.Start(url)
    End Sub

    Private Sub MenuGAF_Click(sender As Object, e As EventArgs) Handles MenuGAF.Click
        Dim url As String = "https://isharenew.dhl.com/sites/gho_afr_forum/default.aspx"
        Process.Start(url)
    End Sub

    Private Sub MenuHOBIT_Click(sender As Object, e As EventArgs) Handles MenuHOBIT.Click
        Dim url As String = "https://isharenew.dhl.com/sites/gho_afr_forum/product/Reports/Forms/AllItems.aspx?RootFolder=%2Fsites%2Fgho%5Fafr%5Fforum%2Fproduct%2FReports&"
        Process.Start(url)
    End Sub

    Private Sub MenuRACC_Click(sender As Object, e As EventArgs) Handles MenuRACC.Click
        Dim url As String = "https://isharenew.dhl.com/sites/RACC_AMER/default.aspx"
        Process.Start(url)
    End Sub

    Private Sub MenuSOP_Click(sender As Object, e As EventArgs) Handles MenuSOP.Click
        Dim url As String = "https://isharenew.dhl.com/sites/RACC_AMER/RACC%20RFQ%20Action%20Tracker/SOP/SOP.html"
        Process.Start(url)
    End Sub

    Private Sub MenuFORWIN_Click(sender As Object, e As EventArgs) Handles MenuFORWIN.Click
        Dim url As String = "https://frp.dhl.com/ibmcognos/cgi-bin/cognosisapi.dll?b_action=xts.run&m=portal/cc.xts&gohome="
        Process.Start(url)
    End Sub

    Private Sub MenuFreightender_Click(sender As Object, e As EventArgs) Handles MenuFreightender.Click
        Dim url As String = "https://app.freightender.com/#/login"
        Process.Start(url)
    End Sub

    Private Sub MenuFRP_Click(sender As Object, e As EventArgs) Handles MenuFRP.Click
        Dim url As String = "https://myfrp.dhl.com/air_freight.html"
        Process.Start(url)
    End Sub

    Private Sub MenuGRIPS_Click(sender As Object, e As EventArgs) Handles MenuGRIPS.Click
        Dim url As String = "https://login.lanetix.com/dhl-dgf"
        Process.Start(url)
    End Sub

    Private Sub MenuMTS_Click(sender As Object, e As EventArgs) Handles MenuMTS.Click
        Dim url As String = "https://mts.dhl.com/AIR/login"
        Process.Start(url)
    End Sub

    Private Sub MenuUSAFRReports_Click(sender As Object, e As EventArgs) Handles MenuUSAFRReports.Click
        Dim url As String = "https://isharenew.dhl.com/sites/USAFR/Steering/Reports/SitePages/Home.aspx"
        Process.Start(url)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        LblTime.Text = TimeOfDay.ToLongTimeString
    End Sub

    Private Sub CountLost()
        OpenCon()
        Try
            query = "SELECT COUNT(AwardStatus) from rfq where AwardStatus = 'Lost'"
            cmd = New SQLiteCommand(query, con)
            LblLost.Text = cmd.ExecuteScalar.ToString & " LOST"
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", Query for Active Bids function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CountWon()
        OpenCon()
        Try
            query = "SELECT COUNT(AwardStatus) from rfq where AwardStatus = 'Won'"
            cmd = New SQLiteCommand(query, con)
            LblWon.Text = cmd.ExecuteScalar.ToString & " WON"
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", Query for Active Bids function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub CountPendingAward()
        OpenCon()
        Try
            query = "SELECT COUNT(AwardStatus) from rfq where AwardStatus = 'Pending Award'"
            cmd = New SQLiteCommand(query, con)
            LblPendingAward.Text = cmd.ExecuteScalar.ToString & " PENDING AWARD"
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", Query for Active Bids function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CountActiveBids()
        'LblActive.Text = String.Empty
        OpenCon()
        Try
            query = "SELECT COUNT(BidActive) from rfq where BidActive = 1"
            cmd = New SQLiteCommand(query, con)
            LblActive.Text = cmd.ExecuteScalar.ToString & " ACTIVE"
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", Query for Active Bids function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CountUpcoming()
        OpenCon()
        Try
            query = "SELECT COUNT(Upcoming) from rfq where Upcoming = 1"
            cmd = New SQLiteCommand(query, con)
            LblUpcoming.Text = cmd.ExecuteScalar.ToString & " UPCOMING"
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", Query for Active Bids function", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        Dim url As String = "Help.md.html"
        Process.Start(url)
    End Sub

    Private Sub BtnExportNotes_Click(sender As Object, e As EventArgs) Handles BtnExportNotesHTML.Click
        ExportNotesViaMarkdown()
    End Sub


    Private Sub ExportNotesViaMarkdown()
        Dim file As StreamWriter
        Dim fileheader As String = "<meta charset=""utf-8"" emacsmode=""-*- markdown -*-""><link rel=""stylesheet"" href=""https://casual-effects.com/markdeep/latest/journal.css?"">"
        Dim filefooter As String = "<style class=""fallback"">body{visibility:hidden}</style><script> markdeepOptions = {tocStyle:'short'};</script><!-- Markdeep: --><script src=""https://casual-effects.com/markdeep/latest/markdeep.min.js?"" charset=""utf-8""></script>"
        Dim path = LblFolderPath.Text & "\"
        Dim markdeep = path & TxtCustomer.Text & "_" & Format(Now, "yyyy-MM-dd_hhmmss") & ".md.html"
        file = My.Computer.FileSystem.OpenTextFileWriter(markdeep, True)
        file.WriteLine(fileheader)
        file.WriteLine(vbCrLf)
        file.WriteLine(vbTab & vbTab & "**" & TxtCustomer.Text & "**")
        file.WriteLine(vbTab & vbTab & FormatDateTime(Now))
        file.WriteLine(vbCrLf)
        file.WriteLine("Current Project Updates")
        file.WriteLine(vbCrLf)
        file.WriteLine(RtbJournal.Text)
        file.WriteLine(vbCrLf)
        file.WriteLine(filefooter)
        file.Close()
    End Sub

    Private Sub ExportNotesViaPlainText()
        Dim file As StreamWriter
        Dim path = LblFolderPath.Text & "\"
        Dim plaintext = path & TxtCustomer.Text & "_" & Format(Now, "yyyy-MM-dd_hhmmss") & ".txt"
        file = My.Computer.FileSystem.OpenTextFileWriter(plaintext, True)
        file.WriteLine(RtbJournal.Text)
        file.Close()
    End Sub

    Private Sub BtnExportNotesPlainText_Click(sender As Object, e As EventArgs) Handles BtnExportNotesPlainText.Click
        ExportNotesViaPlainText()
    End Sub
End Class

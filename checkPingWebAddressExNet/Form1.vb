'
'
''Title: Check if Web Address Destination is Available/Reachable
'
''By: Jason Hensley
'
''Contact: mailto:vbcodesource@gmail.com
'
''Wepage: http://www.vbcodesource.com and http://www.vbforfree.com
'
''Description: This is nothing more than a example of how to use the IsDestinationReachable API 
'Function to check the specified web address to see if the destination was available/reachable 
'or not. It displays the in and out speed for reaching the desctination along with the type of
'network available to be used. I guess in a sense you can consider this a ping to see if the
'network is available but you wouldn't be able to get a accurate roundTripTime since the API
'will also retrieve the In and out speed which can take a few seconds to complete. Still, if
'you only want to know if the address/site is available then this should work just fine for that.
'
'Copyright: 2008, March 7th
'
'
Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnCheck As System.Windows.Forms.Button
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblSuccess As System.Windows.Forms.Label
    Friend WithEvents lblTime As System.Windows.Forms.Label
    Friend WithEvents lblDown As System.Windows.Forms.Label
    Friend WithEvents lblUp As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents lblNetwork As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnCheck = New System.Windows.Forms.Button
        Me.lblTime = New System.Windows.Forms.Label
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.lblAddress = New System.Windows.Forms.Label
        Me.lblDown = New System.Windows.Forms.Label
        Me.lblUp = New System.Windows.Forms.Label
        Me.lblSuccess = New System.Windows.Forms.Label
        Me.lblNetwork = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'btnCheck
        '
        Me.btnCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheck.Location = New System.Drawing.Point(176, 64)
        Me.btnCheck.Name = "btnCheck"
        Me.btnCheck.Size = New System.Drawing.Size(304, 24)
        Me.btnCheck.TabIndex = 0
        Me.btnCheck.Text = "Check the Specified Address"
        '
        'lblTime
        '
        Me.lblTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTime.Location = New System.Drawing.Point(176, 24)
        Me.lblTime.Name = "lblTime"
        Me.lblTime.Size = New System.Drawing.Size(304, 16)
        Me.lblTime.TabIndex = 1
        Me.lblTime.Text = "Time to Reach and Return Data:"
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(8, 64)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(152, 20)
        Me.txtAddress.TabIndex = 2
        Me.txtAddress.Text = "www.vbcodesource.com"
        '
        'lblAddress
        '
        Me.lblAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress.Location = New System.Drawing.Point(8, 8)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(160, 56)
        Me.lblAddress.TabIndex = 3
        Me.lblAddress.Text = "Address to Check/Reach/Ping. You can use IP/URL address or UNC name."
        '
        'lblDown
        '
        Me.lblDown.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDown.Location = New System.Drawing.Point(176, 40)
        Me.lblDown.Name = "lblDown"
        Me.lblDown.Size = New System.Drawing.Size(152, 16)
        Me.lblDown.TabIndex = 4
        Me.lblDown.Text = "Down Speed:"
        '
        'lblUp
        '
        Me.lblUp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblUp.Location = New System.Drawing.Point(328, 40)
        Me.lblUp.Name = "lblUp"
        Me.lblUp.Size = New System.Drawing.Size(152, 16)
        Me.lblUp.TabIndex = 5
        Me.lblUp.Text = "Up Speed:"
        '
        'lblSuccess
        '
        Me.lblSuccess.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSuccess.Location = New System.Drawing.Point(176, 8)
        Me.lblSuccess.Name = "lblSuccess"
        Me.lblSuccess.Size = New System.Drawing.Size(192, 16)
        Me.lblSuccess.TabIndex = 6
        Me.lblSuccess.Text = "Successful:"
        '
        'lblNetwork
        '
        Me.lblNetwork.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNetwork.Location = New System.Drawing.Point(280, 8)
        Me.lblNetwork.Name = "lblNetwork"
        Me.lblNetwork.Size = New System.Drawing.Size(200, 16)
        Me.lblNetwork.TabIndex = 7
        Me.lblNetwork.Text = "Network Type Used:"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(490, 96)
        Me.Controls.Add(Me.lblNetwork)
        Me.Controls.Add(Me.lblSuccess)
        Me.Controls.Add(Me.lblUp)
        Me.Controls.Add(Me.lblDown)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.lblTime)
        Me.Controls.Add(Me.btnCheck)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Check if Web Address is Available using API - 2002/2003"
        Me.ResumeLayout(False)

    End Sub

#End Region
    '
    '    'The type of network the API used to reach its destination.    Enum networkType
        Lan_Network = 1
        Wan_Network = 2
        Aol_Network = 4 'This is pretty much obsolete since it would only be used with Win 95/98.
        No_Network = 0

    End Enum    '    '    'This structure is passed to the 2nd parameter in the isDestination API call ByReference since    'the API will Put info IN the Structure and NOT read data FROM the structure.    Private Structure qocStructure        '        'Used to specifiy the structure's size. The API call will use this value first when you        'call to get the size of the structure being used, and then it will set this value to         'the size of the structure with the data in the structure.        Dim structureSize As Int32        '        'Will hold the value for the type of network.        Dim networkFlags As networkType        '        'Will contain the speed in bytes per second coming from the destination.        Dim inSpeed As Int32        '        'Will contain the speed in bytes per second going to the destination.        Dim outSpeed As Int32    End Structure
    '
    'The API call to get the info we want. This API function is NOT supported by the Windows Vista
    'Operating System. The destination address can be a URL Address, IP Address, or UNC 
    '(Universal Naming Convention) name.
    Private Declare Function IsDestinationReachable Lib "sensapi.dll" Alias _        "IsDestinationReachableA" (ByVal destinationAddress As String, ByRef qocInfoStructure _            As qocStructure) As Int32
    '
    '
    Private Sub btnCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheck.Click
        '
        'Will hold the return value of the API call.
        Dim pingSuccess As Int32
        '
        'Create a variable for the Quality of Connection structure.
        Dim qoc As qocStructure
        '
        'Specify the size of this structure for the API call to read.
        qoc.structureSize = Len(qoc)
        '
        'Get the tick value before calling the API.
        Dim tickStart As Integer = Environment.TickCount
        '
        '
        'Call the API passing the address in the textbox control.
        pingSuccess = IsDestinationReachable(txtAddress.Text, qoc)
        '
        '
        'Get the value after the call.
        Dim tickEnd As Integer = Environment.TickCount
        '
        'Now get the total milli-seconds it took to call the API function and reach the network.
        Dim ticks As Integer = (tickEnd - tickStart)
        '
        '
        'Check if the address was able to be reached before going displaying anymore data.
        If pingSuccess = 1 Then
            '
            'Display "True" since the value was '1'.
            lblSuccess.Text = "Successful: True"

        ElseIf pingSuccess = 0 Then
            '
            'Display "False" since the value was '0'.
            lblSuccess.Text = "Successful: False"

        Else
            '
            'If the return value is not 0 or 1 then the api call failed. It would most likely be
            '-1 which means the api call failed.
            '
            'Display "Actual Call Failed" since the value was not '0 or 1'.
            lblSuccess.Text = "Successful: API Call Failed"

        End If
        '
        'Display the time to call the API call. This is really not that important since the API
        'will also get the up and down speed of the connection before getting back to use. So 
        'this will NOT be the actual ping time to check the address.
        lblTime.Text = "Time to Reach and Return Data: " & ticks.ToString & " ms"
        '
        'Display the UP/Down speed of the connection to reach the network.
        '
        lblDown.Text = "Down Speed: " & (qoc.inSpeed / 1000000).ToString & " mbps"
        lblUp.Text = "Up Speed: " & (qoc.outSpeed / 1000000).ToString & " mbps"
        '
        'Display the network used.
        lblNetwork.Text = "Network Type Used: " & qoc.networkFlags.ToString

    End Sub

End Class
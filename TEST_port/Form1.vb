Imports System.IO.Ports

Public Class Form1
    Dim port1 As New SerialPort()
    Dim port2 As New SerialPort()
    Dim port3 As New SerialPort()
    Dim count As Integer


    Public Function ctiport(ByVal port As String, ByVal co As String)
        Dim odpoved As String
        Dim comport As New SerialPort()
        comport.PortName = port
        If port = "COM1" Then
            comport.BaudRate = 9600
        Else
            comport.BaudRate = 115200
        End If
        comport.Parity = 0
        comport.DataBits = 8
        comport.StopBits = 1
        comport.ReadTimeout = 5000
        comport.NewLine = Chr(13)
        comport.DtrEnable = True
        comport.Open()

        comport.WriteLine(co + Chr(13))
        odpoved = comport.ReadLine()

        comport.Close()

        Return (odpoved)

    End Function


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = True
        Timer1.Interval = 2000

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim st As New Stopwatch()
        Dim buffer As String
        Dim pomoc, resp As Double


        count = count + 1

        st.Start()
        port1.DtrEnable = True
        buffer = "*B1TR1" + Chr(13)
        port1.WriteLine(buffer)
        buffer = port1.ReadLine()
        TextBox1.Text = buffer
        port1.DtrEnable = False
        st.Stop()
        pomoc = st.ElapsedMilliseconds
        Button1.Text = CStr(st.ElapsedMilliseconds)
        st.Reset()

        st.Start()
        port2.DtrEnable = True
        port2.RtsEnable = True
        buffer = "*B1TR1" + Chr(13)
        port2.WriteLine(buffer)
        buffer = port2.ReadLine()
        buffer = Mid(buffer, 7, 4)
        TextBox2.Text = buffer
        port2.DtrEnable = False
        st.Stop()
        pomoc = pomoc + st.ElapsedMilliseconds
        Button2.Text = CStr(st.ElapsedMilliseconds)

        st.Reset()

        st.Start()
        port3.DtrEnable = True
        buffer = "*B1MR0" + Chr(13)
        port3.WriteLine(buffer)
        buffer = port3.ReadLine()
        TextBox3.Text = buffer
        buffer = Mid(buffer, 10, 6) + ";" + Mid(buffer, 20, 6)
        port3.DtrEnable = False
        st.Stop()
        TextBox3.Text = buffer
        pomoc = pomoc + st.ElapsedMilliseconds
        Button4.Text = CStr(st.ElapsedMilliseconds)

        Button3.Text = CStr(count)

        st.Reset()
        st.Start()
        Dim xmtBuf() As Byte = {&H2A, &H61, &H0, &H6, &H31, &H2, &H51, &H1, &HE9, &HD}
        Dim rcvBuf(1024) As Byte
        port2.DtrEnable = True
        port2.Write(xmtBuf, 0, xmtBuf.Length)
        port2.DtrEnable = False
        port2.DtrEnable = True
        port2.Read(rcvBuf, 0, rcvBuf.Length)
        port2.DtrEnable = False
        
        TextBox4.Text = (rcvBuf(8) * 256 + rcvBuf(9) * 1) / 10
        st.Stop()
        pomoc = pomoc + st.ElapsedMilliseconds
        Button5.Text = CStr(st.ElapsedMilliseconds)
        Button6.Text = CStr(pomoc)
        pomoc = 0
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Timer1.Enabled = False
        Me.Close()

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        port1.PortName = "COM1"
        port1.BaudRate = 9600
        port1.Parity = 0
        port1.DataBits = 8
        port1.StopBits = 1
        port1.ReadTimeout = 5000
        port1.NewLine = Chr(13)
        port1.Open()

        port2.PortName = "COM4"
        port2.BaudRate = 115200
        port2.Parity = 0
        port2.DataBits = 8
        port2.StopBits = 1
        port2.ReadTimeout = 5000
        port2.NewLine = Chr(13)
        port2.Open()

        port3.PortName = "COM3"
        port3.BaudRate = 115200
        port3.Parity = 0
        port3.DataBits = 8
        port3.StopBits = 1
        port3.ReadTimeout = 5000
        port3.NewLine = Chr(13)
        port3.Open()


    End Sub
End Class

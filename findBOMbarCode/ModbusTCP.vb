Imports System.Collections
Imports System.Text
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading


Namespace ModbusTCP
    ''' <summary>
    ''' Modbus TCP common driver class. This class implements a modbus TCP master driver.
    ''' It supports the following commands:
    ''' 
    ''' Read coils
    ''' Read discrete inputs
    ''' Write single coil
    ''' Write multiple cooils
    ''' Read holding register
    ''' Read input register
    ''' Write single register
    ''' Write multiple register
    ''' 
    ''' All commands can be sent in synchronous or asynchronous mode. If a value is accessed
    ''' in synchronous mode the program will stop and wait for slave to response. If the 
    ''' slave didn't answer within a specified time a timeout exception is called.
    ''' The class uses multi threading for both synchronous and asynchronous access. For
    ''' the communication two lines are created. This is necessary because the synchronous
    ''' thread has to wait for a previous command to finish.
    ''' 
    ''' </summary>
    Public Class Master
        ' ------------------------------------------------------------------------
        ' Constants for access
        Private Const fctReadCoil As Byte = 1
        Private Const fctReadDiscreteInputs As Byte = 2
        Private Const fctReadHoldingRegister As Byte = 3
        Private Const fctReadInputRegister As Byte = 4
        Private Const fctWriteSingleCoil As Byte = 5
        Private Const fctWriteSingleRegister As Byte = 6
        Private Const fctWriteMultipleCoils As Byte = 15
        Private Const fctWriteMultipleRegister As Byte = 16
        Private Const fctReadWriteMultipleRegister As Byte = 23

        ''' <summary>Constant for exception illegal function.</summary>
        Public Const excIllegalFunction As Byte = 1
        ''' <summary>Constant for exception illegal data address.</summary>
        Public Const excIllegalDataAdr As Byte = 2
        ''' <summary>Constant for exception illegal data value.</summary>
        Public Const excIllegalDataVal As Byte = 3
        ''' <summary>Constant for exception slave device failure.</summary>
        Public Const excSlaveDeviceFailure As Byte = 4
        ''' <summary>Constant for exception acknowledge.</summary>
        Public Const excAck As Byte = 5
        ''' <summary>Constant for exception slave is busy/booting up.</summary>
        Public Const excSlaveIsBusy As Byte = 6
        ''' <summary>Constant for exception gate path unavailable.</summary>
        Public Const excGatePathUnavailable As Byte = 10
        ''' <summary>Constant for exception not connected.</summary>
        Public Const excExceptionNotConnected As Byte = 253
        ''' <summary>Constant for exception connection lost.</summary>
        Public Const excExceptionConnectionLost As Byte = 254
        ''' <summary>Constant for exception response timeout.</summary>
        Public Const excExceptionTimeout As Byte = 255
        ''' <summary>Constant for exception wrong offset.</summary>
        Private Const excExceptionOffset As Byte = 128
        ''' <summary>Constant for exception send failt.</summary>
        Private Const excSendFailt As Byte = 100

        ' ------------------------------------------------------------------------
        ' Private declarations
        Private Shared _timeout As UShort = 10  'OLD 500
        Private Shared _refresh As UShort = 10   'OLD 10
        Private Shared _connected As Boolean = False

        Private tcpAsyCl As Socket
        Private tcpAsyClBuffer As Byte() = New Byte(2047) {}

        Private tcpSynCl As Socket
        Private tcpSynClBuffer As Byte() = New Byte(2047) {}

        ' ------------------------------------------------------------------------
        ''' <summary>Response data event. This event is called when new data arrives</summary>
        Public Delegate Sub ResponseData(ByVal id As UShort, ByVal unit As Byte, ByVal [function] As Byte, ByVal data As Byte())
        ''' <summary>Response data event. This event is called when new data arrives</summary>
        Public Event OnResponseData As ResponseData
        ''' <summary>Exception data event. This event is called when the data is incorrect</summary>
        Public Delegate Sub ExceptionData(ByVal id As UShort, ByVal unit As Byte, ByVal [function] As Byte, ByVal exception As Byte)
        ''' <summary>Exception data event. This event is called when the data is incorrect</summary>
        Public Event OnException As ExceptionData

        ' ------------------------------------------------------------------------
        ''' <summary>Response timeout. If the slave didn't answers within in this time an exception is called.</summary>
        ''' <value>The default value is 500ms.</value>
        Public Property timeout() As UShort
            Get
                Return _timeout
            End Get
            Set(ByVal value As UShort)
                _timeout = value
            End Set
        End Property

        ' ------------------------------------------------------------------------
        ''' <summary>Refresh timer for slave answer. The class is polling for answer every X ms.</summary>
        ''' <value>The default value is 10ms.</value>
        Public Property refresh() As UShort
            Get
                Return _refresh
            End Get
            Set(ByVal value As UShort)
                _refresh = value
            End Set
        End Property

        ' ------------------------------------------------------------------------
        ''' <summary>Shows if a connection is active.</summary>
        Public ReadOnly Property connected() As Boolean
            Get
                Return _connected
            End Get
        End Property

        ' ------------------------------------------------------------------------
        ''' <summary>Create master instance without parameters.</summary>
        Public Sub New()
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Create master instance with parameters.</summary>
        ''' <param name="ip">IP adress of modbus slave.</param>
        ''' <param name="port">Port number of modbus slave. Usually port 502 is used.</param>
        Public Sub New(ByVal ip As String, ByVal port As UShort)
            connect(ip, port)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Start connection to slave.</summary>
        ''' <param name="ip">IP adress of modbus slave.</param>
        ''' <param name="port">Port number of modbus slave. Usually port 502 is used.</param>
        Public Sub connect(ByVal ip As String, ByVal port As UShort)
            Try
                Dim _ip As IPAddress
                If IPAddress.TryParse(ip, _ip) = False Then
                    Dim hst As IPHostEntry = Dns.GetHostEntry(ip)
                    ip = hst.AddressList(0).ToString()
                End If
                ' ----------------------------------------------------------------
                ' Connect asynchronous client
                tcpAsyCl = New Socket(IPAddress.Parse(ip).AddressFamily, SocketType.Stream, ProtocolType.Tcp)
                tcpAsyCl.Connect(New IPEndPoint(IPAddress.Parse(ip), port))
                tcpAsyCl.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendTimeout, _timeout)
                tcpAsyCl.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, _timeout)
                tcpAsyCl.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.NoDelay, 1)
                ' ----------------------------------------------------------------
                '' Connect synchronous client
                'tcpSynCl = New Socket(IPAddress.Parse(ip).AddressFamily, SocketType.Stream, ProtocolType.Tcp)
                'tcpSynCl.Connect(New IPEndPoint(IPAddress.Parse(ip), port))
                'tcpSynCl.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendTimeout, _timeout)
                'tcpSynCl.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, _timeout)
                'tcpSynCl.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.NoDelay, 1)
                _connected = True
            Catch [error] As System.IO.IOException
                _connected = False
                Throw ([error])
            End Try
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Stop connection to slave.</summary>
        Public Sub disconnect()
            Dispose()
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Destroy master instance.</summary>
        Protected Overrides Sub Finalize()
            Try
                Dispose()
            Finally
                MyBase.Finalize()
            End Try
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Destroy master instance</summary>
        Public Sub Dispose()
            If tcpAsyCl IsNot Nothing Then
                If tcpAsyCl.Connected Then
                    Try
                        tcpAsyCl.Shutdown(SocketShutdown.Both)
                    Catch
                    End Try
                    tcpAsyCl.Close()
                End If
                tcpAsyCl = Nothing
            End If
            If tcpSynCl IsNot Nothing Then
                If tcpSynCl.Connected Then
                    Try
                        tcpSynCl.Shutdown(SocketShutdown.Both)
                    Catch
                    End Try
                    tcpSynCl.Close()
                End If
                tcpSynCl = Nothing
            End If
        End Sub

        Friend Sub CallException(ByVal id As UShort, ByVal unit As Byte, ByVal [function] As Byte, ByVal exception As Byte)
            If (tcpAsyCl Is Nothing) OrElse (tcpSynCl Is Nothing) Then
                Return
            End If
            If exception = excExceptionConnectionLost Then
                tcpSynCl = Nothing
                tcpAsyCl = Nothing
            End If
            RaiseEvent OnException(id, unit, [function], exception)
        End Sub

        Friend Shared Function SwapUInt16(ByVal inValue As UInt16) As UInt16
            Return CType(((inValue And &HFF00) >> 8) Or ((inValue And &HFF) << 8), UInt16)
        End Function

        ' ------------------------------------------------------------------------
        ''' <summary>Read coils from slave asynchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        Public Sub ReadCoils(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numInputs As UShort)
            WriteAsyncData(CreateReadHeader(id, unit, startAddress, numInputs, fctReadCoil), id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Read coils from slave synchronous.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        ''' <param name="values">Contains the result of function.</param>
        Public Sub ReadCoils(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numInputs As UShort, ByRef values As Byte())
            values = WriteSyncData(CreateReadHeader(id, unit, startAddress, numInputs, fctReadCoil), id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Read discrete inputs from slave asynchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        Public Sub ReadDiscreteInputs(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numInputs As UShort)
            WriteAsyncData(CreateReadHeader(id, unit, startAddress, numInputs, fctReadDiscreteInputs), id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Read discrete inputs from slave synchronous.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        ''' <param name="values">Contains the result of function.</param>
        Public Sub ReadDiscreteInputs(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numInputs As UShort, ByRef values As Byte())
            values = WriteSyncData(CreateReadHeader(id, unit, startAddress, numInputs, fctReadDiscreteInputs), id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Read holding registers from slave asynchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        Public Sub ReadHoldingRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numInputs As UShort)
            WriteAsyncData(CreateReadHeader(id, unit, startAddress, numInputs, fctReadHoldingRegister), id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Read holding registers from slave synchronous.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        ''' <param name="values">Contains the result of function.</param>
        Public Sub ReadHoldingRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numInputs As UShort, ByRef values As Byte())
            values = WriteSyncData(CreateReadHeader(id, unit, startAddress, numInputs, fctReadHoldingRegister), id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Read input registers from slave asynchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        Public Sub ReadInputRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numInputs As UShort)
            WriteAsyncData(CreateReadHeader(id, unit, startAddress, numInputs, fctReadInputRegister), id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Read input registers from slave synchronous.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        ''' <param name="values">Contains the result of function.</param>
        Public Sub ReadInputRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numInputs As UShort, ByRef values As Byte())
            values = WriteSyncData(CreateReadHeader(id, unit, startAddress, numInputs, fctReadInputRegister), id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Write single coil in slave asynchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="OnOff">Specifys if the coil should be switched on or off.</param>
        Public Sub WriteSingleCoils(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal OnOff As Boolean)
            Dim data As Byte()
            data = CreateWriteHeader(id, unit, startAddress, 1, 1, fctWriteSingleCoil)
            If OnOff = True Then
                data(10) = 255
            Else
                data(10) = 0
            End If
            WriteAsyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Write single coil in slave synchronous.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="OnOff">Specifys if the coil should be switched on or off.</param>
        ''' <param name="result">Contains the result of the synchronous write.</param>
        Public Sub WriteSingleCoils(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal OnOff As Boolean, ByRef result As Byte())
            Dim data As Byte()
            data = CreateWriteHeader(id, unit, startAddress, 1, 1, fctWriteSingleCoil)
            If OnOff = True Then
                data(10) = 255
            Else
                data(10) = 0
            End If
            result = WriteSyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Write multiple coils in slave asynchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numBits">Specifys number of bits.</param>
        ''' <param name="values">Contains the bit information in byte format.</param>
        Public Sub WriteMultipleCoils(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numBits As UShort, ByVal values As Byte())
            Dim numBytes As Byte = Convert.ToByte(values.Length)
            Dim data As Byte()
            data = CreateWriteHeader(id, unit, startAddress, numBits, CByte(numBytes + 2), fctWriteMultipleCoils)
            Array.Copy(values, 0, data, 13, numBytes)
            WriteAsyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Write multiple coils in slave synchronous.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address from where the data read begins.</param>
        ''' <param name="numBits">Specifys number of bits.</param>
        ''' <param name="values">Contains the bit information in byte format.</param>
        ''' <param name="result">Contains the result of the synchronous write.</param>
        Public Sub WriteMultipleCoils(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numBits As UShort, ByVal values As Byte(), ByRef result As Byte())
            Dim numBytes As Byte = Convert.ToByte(values.Length)
            Dim data As Byte()
            data = CreateWriteHeader(id, unit, startAddress, numBits, CByte(numBytes + 2), fctWriteMultipleCoils)
            Array.Copy(values, 0, data, 13, numBytes)
            result = WriteSyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Write single register in slave asynchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address to where the data is written.</param>
        ''' <param name="values">Contains the register information.</param>
        Public Sub WriteSingleRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal values As Byte())
            Dim data As Byte()
            data = CreateWriteHeader(id, unit, startAddress, 1, 1, fctWriteSingleRegister)
            data(10) = values(0)
            data(11) = values(1)
            WriteAsyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Write single register in slave synchronous.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address to where the data is written.</param>
        ''' <param name="values">Contains the register information.</param>
        ''' <param name="result">Contains the result of the synchronous write.</param>
        Public Sub WriteSingleRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal values As Byte(), ByRef result As Byte())
            Dim data As Byte()
            data = CreateWriteHeader(id, unit, startAddress, 1, 1, fctWriteSingleRegister)
            data(10) = values(0)
            data(11) = values(1)
            result = WriteSyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Write multiple registers in slave asynchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address to where the data is written.</param>
        ''' <param name="values">Contains the register information.</param>
        Public Sub WriteMultipleRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal values As Byte())
            Dim numBytes As UShort = Convert.ToUInt16(values.Length)
            If numBytes Mod 2 > 0 Then
                numBytes += 1
            End If
            Dim data As Byte()

            data = CreateWriteHeader(id, unit, startAddress, Convert.ToUInt16(numBytes \ 2), Convert.ToUInt16(numBytes + 2), fctWriteMultipleRegister)
            Array.Copy(values, 0, data, 13, values.Length)
            WriteAsyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Write multiple registers in slave synchronous.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startAddress">Address to where the data is written.</param>
        ''' <param name="values">Contains the register information.</param>
        ''' <param name="result">Contains the result of the synchronous write.</param>
        Public Sub WriteMultipleRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal values As Byte(), ByRef result As Byte())
            Dim numBytes As UShort = Convert.ToUInt16(values.Length)
            If numBytes Mod 2 > 0 Then
                numBytes += 1
            End If
            Dim data As Byte()

            data = CreateWriteHeader(id, unit, startAddress, Convert.ToUInt16(numBytes \ 2), Convert.ToUInt16(numBytes + 2), fctWriteMultipleRegister)
            Array.Copy(values, 0, data, 13, values.Length)
            result = WriteSyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Read/Write multiple registers in slave asynchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startReadAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        ''' <param name="startWriteAddress">Address to where the data is written.</param>
        ''' <param name="values">Contains the register information.</param>
        Public Sub ReadWriteMultipleRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startReadAddress As UShort, ByVal numInputs As UShort, ByVal startWriteAddress As UShort, ByVal values As Byte())
            Dim numBytes As UShort = Convert.ToUInt16(values.Length)
            If numBytes Mod 2 > 0 Then
                numBytes += 1
            End If
            Dim data As Byte()

            data = CreateReadWriteHeader(id, unit, startReadAddress, numInputs, startWriteAddress, Convert.ToUInt16(numBytes \ 2))
            Array.Copy(values, 0, data, 17, values.Length)
            WriteAsyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ''' <summary>Read/Write multiple registers in slave synchronous. The result is given in the response function.</summary>
        ''' <param name="id">Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.</param>
        ''' <param name="unit">Unit identifier (previously slave address). In asynchonous mode this unit is given to the callback function.</param>
        ''' <param name="startReadAddress">Address from where the data read begins.</param>
        ''' <param name="numInputs">Length of data.</param>
        ''' <param name="startWriteAddress">Address to where the data is written.</param>
        ''' <param name="values">Contains the register information.</param>
        ''' <param name="result">Contains the result of the synchronous command.</param>
        Public Sub ReadWriteMultipleRegister(ByVal id As UShort, ByVal unit As Byte, ByVal startReadAddress As UShort, ByVal numInputs As UShort, ByVal startWriteAddress As UShort, ByVal values As Byte(), _
         ByRef result As Byte())
            Dim numBytes As UShort = Convert.ToUInt16(values.Length)
            If numBytes Mod 2 > 0 Then
                numBytes += 1
            End If
            Dim data As Byte()

            data = CreateReadWriteHeader(id, unit, startReadAddress, numInputs, startWriteAddress, Convert.ToUInt16(numBytes \ 2))
            Array.Copy(values, 0, data, 17, values.Length)
            result = WriteSyncData(data, id)
        End Sub

        ' ------------------------------------------------------------------------
        ' Create modbus header for read action
        Private Function CreateReadHeader(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal length As UShort, ByVal [function] As Byte) As Byte()
            Dim data As Byte() = New Byte(11) {}

            Dim _id As Byte() = BitConverter.GetBytes(CShort(id))
            data(0) = _id(1)
            ' Slave id high byte
            data(1) = _id(0)
            ' Slave id low byte
            data(5) = 6
            ' Message size
            data(6) = unit
            ' Slave address
            data(7) = [function]
            ' Function code
            Dim _adr As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(startAddress))))
            data(8) = _adr(0)
            ' Start address
            data(9) = _adr(1)
            ' Start address
            Dim _length As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(length))))
            data(10) = _length(0)
            ' Number of data to read
            data(11) = _length(1)
            ' Number of data to read
            Return data
        End Function

        ' ------------------------------------------------------------------------
        ' Create modbus header for write action
        Private Function CreateWriteHeader(ByVal id As UShort, ByVal unit As Byte, ByVal startAddress As UShort, ByVal numData As UShort, ByVal numBytes As UShort, ByVal [function] As Byte) As Byte()
            Dim data As Byte() = New Byte(numBytes + 10) {}

            Dim _id As Byte() = BitConverter.GetBytes(CShort(id))
            data(0) = _id(1)
            ' Slave id high byte
            data(1) = _id(0)
            ' Slave id low byte
            Dim _size As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(5 + numBytes))))
            data(4) = _size(0)
            ' Complete message size in bytes
            data(5) = _size(1)
            ' Complete message size in bytes
            data(6) = unit
            ' Slave address
            data(7) = [function]
            ' Function code
            Dim _adr As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(startAddress))))
            data(8) = _adr(0)
            ' Start address
            data(9) = _adr(1)
            ' Start address
            If [function] >= fctWriteMultipleCoils Then
                Dim _cnt As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(numData))))
                data(10) = _cnt(0)
                ' Number of bytes
                data(11) = _cnt(1)
                ' Number of bytes
                data(12) = CByte(numBytes - 2)
            End If
            Return data
        End Function

        ' ------------------------------------------------------------------------
        ' Create modbus header for read/write action
        Private Function CreateReadWriteHeader(ByVal id As UShort, ByVal unit As Byte, ByVal startReadAddress As UShort, ByVal numRead As UShort, ByVal startWriteAddress As UShort, ByVal numWrite As UShort) As Byte()
            Dim data As Byte() = New Byte(numWrite * 2 + 16) {}

            Dim _id As Byte() = BitConverter.GetBytes(CShort(id))
            data(0) = _id(1)
            ' Slave id high byte
            data(1) = _id(0)
            ' Slave id low byte
            Dim _size As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(11 + numWrite * 2))))
            data(4) = _size(0)
            ' Complete message size in bytes
            data(5) = _size(1)
            ' Complete message size in bytes
            data(6) = unit
            ' Slave address
            data(7) = fctReadWriteMultipleRegister
            ' Function code
            Dim _adr_read As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(startReadAddress))))
            data(8) = _adr_read(0)
            ' Start read address
            data(9) = _adr_read(1)
            ' Start read address
            Dim _cnt_read As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(numRead))))
            data(10) = _cnt_read(0)
            ' Number of bytes to read
            data(11) = _cnt_read(1)
            ' Number of bytes to read
            Dim _adr_write As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(startWriteAddress))))
            data(12) = _adr_write(0)
            ' Start write address
            data(13) = _adr_write(1)
            ' Start write address
            Dim _cnt_write As Byte() = BitConverter.GetBytes(CShort(IPAddress.HostToNetworkOrder(CShort(numWrite))))
            data(14) = _cnt_write(0)
            ' Number of bytes to write
            data(15) = _cnt_write(1)
            ' Number of bytes to write
            data(16) = CByte(numWrite * 2)

            Return data
        End Function

        ' ------------------------------------------------------------------------
        ' Write asynchronous data
        Private Sub WriteAsyncData(ByVal write_data As Byte(), ByVal id As UShort)
            If (tcpAsyCl IsNot Nothing) AndAlso (tcpAsyCl.Connected) Then
                Try
                    tcpAsyCl.BeginSend(write_data, 0, write_data.Length, SocketFlags.None, New AsyncCallback(AddressOf OnSend), Nothing)
                    tcpAsyCl.BeginReceive(tcpAsyClBuffer, 0, tcpAsyClBuffer.Length, SocketFlags.None, New AsyncCallback(AddressOf OnReceive), tcpAsyCl)
                Catch generatedExceptionName As SystemException
                    CallException(id, write_data(6), write_data(7), excExceptionConnectionLost)
                End Try
            Else
                CallException(id, write_data(6), write_data(7), excExceptionConnectionLost)
            End If
        End Sub

        ' ------------------------------------------------------------------------
        ' Write asynchronous data acknowledge
        Private Sub OnSend(ByVal result As System.IAsyncResult)
            If result.IsCompleted = False Then
                CallException(&HFFFF, &HFF, &HFF, excSendFailt)
            End If
        End Sub

        ' ------------------------------------------------------------------------
        ' Write asynchronous data response
        Private Sub OnReceive(ByVal result As System.IAsyncResult)
            If result.IsCompleted = False Then
                CallException(&HFF, &HFF, &HFF, excExceptionConnectionLost)
            End If

            Dim id As UShort = SwapUInt16(BitConverter.ToUInt16(tcpAsyClBuffer, 0))
            Dim unit As Byte = tcpAsyClBuffer(6)
            Dim [function] As Byte = tcpAsyClBuffer(7)
            Dim data As Byte()

            ' ------------------------------------------------------------
            ' Write response data
            If ([function] >= fctWriteSingleCoil) AndAlso ([function] <> fctReadWriteMultipleRegister) Then
                data = New Byte(1) {}
                Array.Copy(tcpAsyClBuffer, 10, data, 0, 2)
            Else
                ' ------------------------------------------------------------
                ' Read response data
                data = New Byte(tcpAsyClBuffer(8) - 1) {}
                Array.Copy(tcpAsyClBuffer, 9, data, 0, tcpAsyClBuffer(8))
            End If
            ' ------------------------------------------------------------
            ' Response data is slave exception
            If [function] > excExceptionOffset Then
                [function] -= excExceptionOffset
                CallException(id, unit, [function], tcpAsyClBuffer(8))
                ' ------------------------------------------------------------
                ' Response data is regular data
            Else
                'ElseIf OnResponseData IsNot Nothing Then
                RaiseEvent OnResponseData(id, unit, [function], data)
            End If

        End Sub

        ' ------------------------------------------------------------------------
        ' Write data and and wait for response
        Private Function WriteSyncData(ByVal write_data As Byte(), ByVal id As UShort) As Byte()

            If tcpSynCl.Connected Then
                Try
                    tcpSynCl.Send(write_data, 0, write_data.Length, SocketFlags.None)
                    Dim result As Integer = tcpSynCl.Receive(tcpSynClBuffer, 0, tcpSynClBuffer.Length, SocketFlags.None)

                    Dim unit As Byte = tcpSynClBuffer(6)
                    Dim [function] As Byte = tcpSynClBuffer(7)
                    Dim data As Byte()

                    If result = 0 Then
                        CallException(id, unit, write_data(7), excExceptionConnectionLost)
                    End If

                    ' ------------------------------------------------------------
                    ' Response data is slave exception
                    If [function] > excExceptionOffset Then
                        [function] -= excExceptionOffset
                        CallException(id, unit, [function], tcpSynClBuffer(8))
                        Return Nothing
                        ' ------------------------------------------------------------
                        ' Write response data
                    ElseIf ([function] >= fctWriteSingleCoil) AndAlso ([function] <> fctReadWriteMultipleRegister) Then
                        data = New Byte(1) {}
                        Array.Copy(tcpSynClBuffer, 10, data, 0, 2)
                    Else
                        ' ------------------------------------------------------------
                        ' Read response data
                        data = New Byte(tcpSynClBuffer(8) - 1) {}
                        Array.Copy(tcpSynClBuffer, 9, data, 0, tcpSynClBuffer(8))
                    End If
                    Return data
                Catch generatedExceptionName As SystemException
                    CallException(id, write_data(6), write_data(7), excExceptionConnectionLost)
                End Try
            Else
                CallException(id, write_data(6), write_data(7), excExceptionConnectionLost)
            End If
            Return Nothing
        End Function
    End Class
End Namespace
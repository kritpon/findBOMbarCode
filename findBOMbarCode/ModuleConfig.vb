Imports System.Data.OleDb
Module ModuleConfig
    Friend _Directory As String = My.Application.Info.DirectoryPath
    'Dim ini As New Setting.IniFile("setting.ini")
    Friend _ini As New Setting.IniFile(_Directory & "\Config.ini")
    Friend _svIPAddress1, _svIPAddress2, _svIPAddress3, _svIPAddress4, _svIPAddress5, _svIPAddress6 As String
    Friend _svID, _svUnit, _svServer, _svSize, _svStartAddress, _Interval_values, _Interval_readAdd, _svDelay As Integer
    ' ------------------------------------------------------------------------
    ' Connect Database
    ' ------------------------------------------------------------------------
    'Friend Conn As New OleDbConnection
    'Friend da As New OleDbDataAdapter
    'Friend ds As New DataSet
    'Friend tables As DataTableCollection
    'Friend source1 As New BindingSource
    ''Friend StrConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;" '& "Data Source=.\Database\TruckScale_MDB.mdb;Jet OLEDB:Database Password=987654321"
    'Friend StrConn As String = ""

    Sub LoadSetting()
        _svID = _ini.ReadValue("SetSystem", "svID")
        _svUnit = _ini.ReadValue("SetSystem", "svUnit")
        _svStartAddress = _ini.ReadValue("SetSystem", "svStartAddress")
        _svSize = _ini.ReadValue("SetSystem", "svSize")
        _Interval_values = _ini.ReadValue("SetSystem", "svInterval_values_ConnectIP")
        _Interval_readAdd = _ini.ReadValue("SetSystem", "svInterval_values_ReadAddress")
        _svIPAddress1 = _ini.ReadValue("SetSystem", "svIPAddress1")
        _svIPAddress2 = _ini.ReadValue("SetSystem", "svIPAddress2")
        _svIPAddress3 = _ini.ReadValue("SetSystem", "svIPAddress3")
        _svIPAddress4 = _ini.ReadValue("SetSystem", "svIPAddress4")
        _svIPAddress5 = _ini.ReadValue("SetSystem", "svIPAddress5")
        _svIPAddress6 = _ini.ReadValue("SetSystem", "svIPAddress6")
        _svServer = _ini.ReadValue("SetSystem", "svServer")
        _svDelay = _ini.ReadValue("SetSystem", "svDelay")
        strConn = _ini.ReadValue("SetSystem", "svParthDataSource")
    End Sub

    Sub setYear()
        'System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("th-TH")
        System.Threading.Thread.CurrentThread.CurrentUICulture = System.Threading.Thread.CurrentThread.CurrentCulture
        Microsoft.Win32.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\International", "sShortDate", "dd/MM/yyyy")
        Microsoft.Win32.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\International", "sTimeFormat", "HH:mm:ss")
        Microsoft.Win32.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\International", "sShortTime", "HH:mm")
    End Sub
End Module

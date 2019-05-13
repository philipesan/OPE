Attribute VB_Name = "mdPublicVars"
Public fso As New Scripting.FileSystemObject
Public dbtemp As DAO.Database
' #rs é o RecordSet, con é a Conexão com o banco de dados, dbtemp é o Banco de dados temporário.
Public con As New ADODB.Connection, strConn As String, rs As New ADODB.Recordset
'Caminho do DB
Public sFilePath As String
Public sLogPath As String


Attribute VB_Name = "mdPublicVars"
Public fso As New Scripting.FileSystemObject
Public dbtemp As DAO.Database
' #rs � o RecordSet, con � a Conex�o com o banco de dados, dbtemp � o Banco de dados tempor�rio.
Public con As New ADODB.Connection, strConn As String, rs As New ADODB.Recordset
'Caminho do DB
Public sFilePath As String
Public sLogPath As String


'autoria Leonardo Vivas
'Versão 1.8
'criação 03/06/2009
'modificação 21/12/2011
' -----------------------------------------------------------------' 
Option Explicit
Dim objRootLDAP, objContainer, objUser, objShell
Dim objExcel, objSpread, intRow
Dim strUser, strOU, strSheet
Dim strCN, strSam, strFirst, strLast, strPWD

' -----------------------------------------------'
' Important change OU= and strSheet to reflect your domain
' -----------------------------------------------'

strOU = "OU=Accounts7 ," ' Note the comma
strSheet = "cria usuarios no ad.xls"

' Bind to Active Directory, Users container.
Set objRootLDAP = GetObject("LDAP://rootDSE")
Set objContainer = GetObject("LDAP://" & strOU & _
objRootLDAP.Get("defaultNamingContext")) 

' Open the Excel spreadsheet
Set objExcel = CreateObject("Excel.Application")
Set objSpread = objExcel.Workbooks.Open(strSheet)
intRow = 3 'Row 1 often contains headings

' Here is the 'DO...Loop' that cycles through the cells
' Note intRow, x must correspond to the column in strSheet
Do Until objExcel.Cells(intRow,1).Value = ""
   strSam = Trim(objExcel.Cells(intRow, 1).Value)
   strCN = Trim(objExcel.Cells(intRow, 2).Value) 
   strFirst = Trim(objExcel.Cells(intRow, 3).Value)
   strLast = Trim(objExcel.Cells(intRow, 4).Value)
   strPWD = Trim(objExcel.Cells(intRow, 5).Value)

   ' Build the actual User from data in strSheet.
   Set objUser = objContainer.Create("User", "cn=" & strCN)
   objUser.sAMAccountName = strSam
   objUser.givenName = strFirst
   objUser.sn = strLast
   objUser.SetInfo

   ' Separate section to enable account with its password
   objUser.userAccountControl = 512
   objUser.pwdLastSet = 0
   objUser.SetPassword strPWD
   objUser.SetInfo

intRow = intRow + 1
Loop
objExcel.Quit 

WScript.Quit 

' End of free example UserSpreadsheet VBScript.
'!TITLE "Robot program"
#Include "Variant.h"

'MOXAIOLogik API Wrapper-------------------------------------------------------
' @file    MOXAIOLogik.h
' @brief   Wrapper for MOXA Modbus Registers
' @details 

' @version 1.0.1
' @date    2019/11/04
' @author  Carlos A. Lopez (DPAM - R&D)

' Software License Agreement (MIT License)

' @copyright Copyright (c) 2019 DENSO 2DLab

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
'------------------------------------------------------------------------------

'GETING STARTED----------------------------------------------------------------
' Make sure to go over the setup guide in the following document prior to 
' running this program.
'
' @doc
'-------------------------------------------------------------------------------

'///////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////

'---Moxa Settings (EDIT BY USER)------------------------------------------------
Dim MOXA_IP As String = "192.168.127.254"	'MOXA IP Address	
Dim sessionName As String = "MOXA"			'Optional if using multiple devices			
Dim MOXA_DI As Integer = 4					'Number of Digital Inputs on Module
Dim MOXA_DO As Integer = 4					'Number of Digital Output on Module
Dim MOXA_AI As Integer = 4					'Number of Digital Inputs on Module
'-------------------------------------------------------------------------------

'---MOXA_SetDO------------------------------------------------------------------
'Description: Set DO moxa channel to ON status.
'
'Inputs: iVal	->	DO Channel Number (0 to MOXA_DO-1)
'Output: N/A
'-------------------------------------------------------------------------------
Sub MOXA_SetIO(ByVal iVal As Integer)
	Dim objDO As Object, varIndex As String

	'Check that Index is in channel range (based on E1200s model)
	If IndexOutRange("DO", iVal) Then Err.Raise &H83600014
	
	objDO = MOXA_Connect(iVal+1, "DO")

	'Output Value
	objDO.Value = &B1
End Sub

'---MOXA_ResetDO----------------------------------------------------------------
'Description: Set DO moxa channel to OFF status.
'
'Inputs: iVal	->	DO Channel Number (0 to MOXA_DO-1)
'Output: N/A
'-------------------------------------------------------------------------------
Sub MOXA_ResetIO(ByVal iVal As Integer)
	Dim objDO As Object, varIndex As String

	'Check that Index is in channel range (based on E1200s model)
	If IndexOutRange("DO", iVal) Then Err.Raise &H83600014

	objDO = MOXA_Connect(iVal+1, "DO")

	'Output Value
	objDO.Value = &B0
End Sub

'---MOXA_DO_Status--------------------------------------------------------------
'Description: Get DO channel status 
'
'Inputs: iVal			->	DO Channel Number (0 to MOXA_DO-1)
'Output: MOXA_DO_Status	->	DO Channel Status 
'-------------------------------------------------------------------------------
Function MOXA_DO_Status(ByVal iVal As Integer) As Integer
	Dim objDO As Object, varIndex As String

	'Check that Index is in channel range
	If IndexOutRange("DO", iVal) Then Err.Raise &H83600014

	objDO = MOXA_Connect(iVal+1, "DO")

	MOXA_DO_Status = objDO.Value
	
End Function

'---MOXA_DO_ReadAll-------------------------------------------------------------
'Description: Get All DO channel status 
'
'Output: MOXA_DO_ReadAll	->	All DO Channel Status as Integer
'-------------------------------------------------------------------------------
Function MOXA_DO_ReadAll() As Integer
	Dim objDO As Object

	objDO = MOXA_Connect(33, "HR")

	MOXA_DO_ReadAll = objDO.Value
End Function

'---MOXA_DO_WriteAll------------------------------------------------------------
'Description: Write to all outputs by converting Integer value to Binary. 
'
'Input: Write Value
'-------------------------------------------------------------------------------
Sub MOXA_DO_WriteAll(setValue As Integer)
	Dim objDO As Object

	objDO = MOXA_Connect(33, "HR")

	'Output Value
	objDO.Value = setValue
End Sub

'---MOXA_DI_Status--------------------------------------------------------------
'Description: Get DI channel status 
'
'Inputs: iVal			->	DI Channel Number (0 to MOXA_DI-1)
'Output: MOXA_DO_Status	->	DI Channel Status
'-------------------------------------------------------------------------------
Function MOXA_DI_Status(ByVal iVal As Integer) As Integer
	Dim objDI As Object, varIndex As String

	'Check that Index is in channel range (based on E1200s model)
	If IndexOutRange("DI", iVal) Then Err.Raise &H83600014

	objDI = MOXA_Connect(iVal+1, "DI")

	'Output Value
	MOXA_DI_Status = objDI.Value
End Function

'---MOXA_DI_ReadAll-------------------------------------------------------------
'Description: Get all DI channel status 
'
'Output: MOXA_DO_Status	->	All DI Channel Status
'-------------------------------------------------------------------------------
Function MOXA_DI_ReadAll() As Integer
	Dim objDI As Object

	objDI = MOXA_Connect(49, "IR")

	'Output Value
	MOXA_DI_ReadAll = objDI.Value
End Function

'---MOXA_AI_Value---------------------------------------------------------------
'Description: Get AI channel Value (Float)
'
'Inputs: iVal			->	AI Channel Number (0 to MOXA_AI-1)
'Output: MOXA_AI_Value	->	AI Channel Value as Float Type
'-------------------------------------------------------------------------------
Function MOXA_AI_Value(ByVal iVal As Integer) As Single
	Dim objAI As Object, varIndex As String

	'Check that Index is in channel range (based on E1200s model)
	If IndexOutRange("AI", iVal) Then Err.Raise &H83600014

	objAI = MOXA_Connect(iVal*2 + 521, "AI")

	'Output Value
	MOXA_AI_Value = objAI.Value
		
End Function

'---MOXA_Connect----------------------------------------------------------------
'Description: Internal Function to Establish communication with Moxa. User must
'			  edit IP Address and session name to be unique. 
'
'Inputs: varIndex 	->	CAO Variable Name (Internal Use)
'		 IOType		->	DO, DI, or AI (Internal Use)
'
'Output: MOXA_Connect	->	CAO Variable Reference 
'-------------------------------------------------------------------------------
Function MOXA_Connect(varIndex As Integer, IOType As String) As Object

	'---Modbus Provider Connection Settings
	Dim providerName As String = "CaoProv.Modbus.X"
	Dim extensionName AS String = "Moxa_Ex"
	Dim providerOpt As String
	providerOpt = "Conn=eth:" & MOXA_IP & ", @IFNotMember"
	Dim extOpt As String = "UnitAddress=1, @IfNotMember"
	Dim extVarOpt As String

	'---Connect
	Dim objCtrl As Object, objExt As Object, varName As String

	objCtrl = Cao.AddController(sessionName, providerName, "", providerOpt)
	objExt = objCtrl.AddExtension(extensionName, extOpt)

	Select Case IOType
		Case "DO"
			varName = Sprintf("DO%d", varIndex)
			extVarOpt = "UserVarWidth=1, @IfNotMember"
		Case "HR"
			varName = Sprintf("HRI%d", varIndex)
			extVarOpt = "UserVarWidth=16, @IfNotMember"
		Case "DI"
			varName = Sprintf("DI%d", varIndex)
			extVarOpt = "UserVarWidth=1, @IfNotMember"
		Case "IR"
			varName = Sprintf("IRI%d", varIndex)
			extVarOpt = "UserVarWidth=16, @IfNotMember"
		Case "AI"
			varName = Sprintf("IRF%d", varIndex)
			extVarOpt = "UserVarWidth=32, @IfNotMember"
	End Select

	MOXA_Connect = objExt.AddVariable(varIndex, extVarOpt)
End Function

'---IndexOutRange---------------------------------------------------------------
'Description: Check if Moxa Channel is in Range. User should edit MOXA_XX
'			  variables to make sure they match MOXA model.
'
'Inputs: IOType ->	DO, DI, or AI
'		 iVal	->	Channel Number (0 to MOXA_DO-1)
'
'Output: IndexOutRange	->	0 = In Range (OK) 
'							1 = Out Of Range (NG)
'-------------------------------------------------------------------------------
Function IndexOutRange(IOType As String, iVal As Integer) As Integer
	IndexOutRange = 0

	If iVal < 0 Then Exit Function

	Select Case IOType
		Case "DO"
			If iVal > MOXA_DO Then IndexOutRange = 1
		Case "DI"
			If iVal > MOXA_DI Then IndexOutRange = 1
		Case "AI"
			If iVal > MOXA_AI Then IndexOutRange = 1
	End Select
End Function

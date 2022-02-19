// ----------------------------------------------------------------------------------------
//                              COPYRIGHT NOTICE
// ----------------------------------------------------------------------------------------
//
// The Source Code Store LLC
// ACTIVEGANTT SCHEDULER COMPONENT FOR C++ - ActiveGanttVC
// ActiveX Control
// Copyright (c) 2002-2017 The Source Code Store LLC
//
// All Rights Reserved. No parts of this file may be reproduced, modified or transmitted 
// in any form or by any means without the written permission of the author.
//
// ----------------------------------------------------------------------------------------
#pragma once



class clsDB
{
public:
	clsDB(CDatabase &oConn);
	~clsDB(void);

	vCString mp_asFieldNames;
	vCString mp_asParams;
	CDatabase* mp_oConn;
	CRecordset mp_oRecordset;
	BOOL mp_bRecordsetActive;

	void AddParameter(CString v_sFieldName, CString v_vParam, ParamType v_sType);
	void AddParameter(CString v_sFieldName, short v_vParam, ParamType v_sType);
	void AddParameter(CString v_sFieldName, int v_vParam, ParamType v_sType);
	void AddParameter(CString v_sFieldName, float v_vParam, ParamType v_sType);
	void AddParameter(CString v_sFieldName, double v_vParam, ParamType v_sType);
	void AddParameter(CString v_sFieldName, COleDateTime v_vParam, ParamType v_sType);

	CString Insert(CString v_sTableName);
	CString Update(CString v_sTableName, CString v_sWHERE);
	CString ExecuteInsert(CString sTableName);
	void ExecuteUpdate(CString sTableName, CString v_sWHERE);
	void InitReader(CString sSQL);
	void CloseReader();
	void ReadTextBox(CEdit &oControl, CString sFieldName);
	void ReadCheckBox(CButton &oControl, CString sFieldName);
	void ReadComboBox(CComboBox &oControl, CString sFieldName);
	COleDateTime ReadFieldDate(CString sFieldName);
	double ReadFieldDouble(CString sFieldName);
	CString ReadFieldString(CString sFieldName);

};
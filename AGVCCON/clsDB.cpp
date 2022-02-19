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
#include "StdAfx.h"
#include "clsDB.h"

clsDB::clsDB(CDatabase &oConn)
{
	mp_oConn = &oConn;
	mp_oRecordset.m_pDatabase = &oConn;
	mp_bRecordsetActive = FALSE;
}

clsDB::~clsDB(void)
{
	if (mp_bRecordsetActive == TRUE)
	{
		mp_oRecordset.Close();
	}
}

void clsDB::AddParameter(CString v_sFieldName, CString v_vParam, ParamType v_sType)
{
	mp_asFieldNames.push_back(v_sFieldName);
	if (v_sType == PT_STRING)
	{
		if (v_vParam.GetLength() == 0)
		{
			mp_asParams.push_back(_T("''"));
		}
		else
		{
			v_vParam.Replace(_T("'"), _T("''"));
			mp_asParams.push_back(_T("'") + v_vParam + _T("'"));
		}
	}
	else if (v_sType == PT_STRING_EMPTYISNULL)
	{
		if (v_vParam.GetLength() == 0)
		{
			mp_asParams.push_back(_T("NULL"));
		}
		else
		{
			v_vParam.Replace(_T("'"), _T("''"));
			mp_asParams.push_back(_T("'") + v_vParam + _T("'"));
		}
	}
	else
	{
		MessageBox(NULL, _T("Error: clsDB::AddParameter CString"), _T("Error"), MB_OK);
	}
}

void clsDB::AddParameter(CString v_sFieldName, short v_vParam, ParamType v_sType)
{
	mp_asFieldNames.push_back(v_sFieldName);
	if (v_sType == PT_NUMERIC)
	{
		mp_asParams.push_back(CStr(v_vParam));
	}
	else if (v_sType == PT_NUMERIC_ZEROISNULL)
	{
		if (v_vParam == 0)
		{
			mp_asParams.push_back(_T("NULL"));
		}
		else
		{
			mp_asParams.push_back(CStr(v_vParam));
		}
	}
	else
	{
		MessageBox(NULL, _T("Error: clsDB::AddParameter short"), _T("Error"), MB_OK);
	}
}

void clsDB::AddParameter(CString v_sFieldName, int v_vParam, ParamType v_sType)
{
	mp_asFieldNames.push_back(v_sFieldName);
	if (v_sType == PT_NUMERIC)
	{
		mp_asParams.push_back(CStr(v_vParam));
	}
	else if (v_sType == PT_NUMERIC_ZEROISNULL)
	{
		if (v_vParam == 0)
		{
			mp_asParams.push_back(_T("NULL"));
		}
		else
		{
			mp_asParams.push_back(CStr(v_vParam));
		}
	}
	else if (v_sType == PT_BOOL)
	{
		if (v_vParam == TRUE)
		{
			mp_asParams.push_back(_T("-1"));
		}
		else
		{
			mp_asParams.push_back(_T("0"));
		}
	}
	else
	{
		MessageBox(NULL, _T("Error: clsDB::AddParameter int"), _T("Error"), MB_OK);
	}
}

void clsDB::AddParameter(CString v_sFieldName, float v_vParam, ParamType v_sType)
{
	mp_asFieldNames.push_back(v_sFieldName);
	if (v_sType == PT_NUMERIC)
	{
		mp_asParams.push_back(CStr(v_vParam));
	}
	else if (v_sType == PT_NUMERIC_ZEROISNULL)
	{
		if (v_vParam == 0)
		{
			mp_asParams.push_back(_T("NULL"));
		}
		else
		{
			mp_asParams.push_back(CStr(v_vParam));
		}
	}
	else
	{
		MessageBox(NULL, _T("Error: clsDB::AddParameter float"), _T("Error"), MB_OK);
	}
}

void clsDB::AddParameter(CString v_sFieldName, double v_vParam, ParamType v_sType)
{
	mp_asFieldNames.push_back(v_sFieldName);
	if (v_sType == PT_NUMERIC)
	{
		mp_asParams.push_back(CStr(v_vParam));
	}
	else if (v_sType == PT_NUMERIC_ZEROISNULL)
	{
		if (v_vParam == 0)
		{
			mp_asParams.push_back(_T("NULL"));
		}
		else
		{
			mp_asParams.push_back(CStr(v_vParam));
		}
	}
	else
	{
		MessageBox(NULL, _T("Error: clsDB::AddParameter double"), _T("Error"), MB_OK);
	}
}

void clsDB::AddParameter(CString v_sFieldName, COleDateTime v_vParam, ParamType v_sType)
{
	mp_asFieldNames.push_back(v_sFieldName);
	if (v_sType == PT_DATE_EMPTYISNULL)
	{
		if (v_vParam == 0)
		{
			mp_asParams.push_back(_T("NULL"));
		}
		else
		{
			CString sDate;
			sDate = _T("#") + CStr(v_vParam.GetYear()) + _T("-") + CStr(v_vParam.GetMonth()) + _T("-") + CStr(v_vParam.GetDay()) + _T(" ") + CStr(v_vParam.GetHour()) + _T(":") + CStr(v_vParam.GetMinute()) + _T(":") + CStr(v_vParam.GetSecond()) + _T("#");
			mp_asParams.push_back(sDate);
		}
	}
	else if (v_sType == PT_DATE)
	{
		CString sDate;
		sDate = _T("#") + CStr(v_vParam.GetYear()) + _T("-") + CStr(v_vParam.GetMonth()) + _T("-") + CStr(v_vParam.GetDay()) + _T(" ") + CStr(v_vParam.GetHour()) + _T(":") + CStr(v_vParam.GetMinute()) + _T(":") + CStr(v_vParam.GetSecond()) + _T("#");
		mp_asParams.push_back(sDate);
	}
	else
	{
		MessageBox(NULL, _T("Error: clsDB::AddParameter CDateTime"), _T("Error"), MB_OK);
	}
}

CString clsDB::Insert(CString v_sTableName)
{
	UINT iIndex;
	CString sSQL;
    sSQL = _T("INSERT INTO ") + v_sTableName + _T(" (");
	for (iIndex = 0; iIndex <= (mp_asFieldNames.size() - 1); iIndex ++)
	{
		sSQL = sSQL + _T("[") + mp_asFieldNames[iIndex] + _T("],");
	}
    sSQL = sSQL.Mid(0, sSQL.GetLength() - 1);
    sSQL = sSQL + _T(") VALUES (");
	for (iIndex = 0; iIndex <= (mp_asParams.size() - 1); iIndex ++)
	{
		sSQL = sSQL + mp_asParams[iIndex] + _T(",");
	}
    sSQL = sSQL.Mid(0, sSQL.GetLength() - 1);
    sSQL = sSQL + _T(")");
    mp_asFieldNames.clear();
    mp_asParams.clear();
	return sSQL;
}

CString clsDB::Update(CString v_sTableName, CString v_sWHERE)
{
	UINT iIndex;
	CString sSQL;
    sSQL = _T("UPDATE ") + v_sTableName + _T(" SET ");
    for (iIndex = 0; iIndex <= (mp_asParams.size() - 1); iIndex ++)
	{
		sSQL = sSQL + _T("[") + mp_asFieldNames[iIndex] + _T("]=") + mp_asParams[iIndex] + _T(",");
	}
    sSQL = sSQL.Mid(0, sSQL.GetLength() - 1);
    sSQL = sSQL + _T(" WHERE ") + v_sWHERE;
    mp_asFieldNames.clear();
    mp_asParams.clear();
	return sSQL;
}

CString clsDB::ExecuteInsert(CString sTableName)
{
	CRecordset oRecordset(mp_oConn);
	CString sSQL = Insert(sTableName);
	mp_oConn->ExecuteSQL(sSQL);
	CString sReturn;
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT @@IDENTITY AS NewID"), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{	LONG lReturn = GetIntField(oRecordset, _T("NewID"));
		sReturn = CStr(lReturn);
	}
	oRecordset.Close();
	return sReturn;
}

void clsDB::ExecuteUpdate(CString sTableName, CString v_sWHERE)
{
	mp_oConn->ExecuteSQL(Update(sTableName, v_sWHERE));
}

void clsDB::InitReader(CString sSQL)
{
	mp_oRecordset.Open(CRecordset::forwardOnly, sSQL, CRecordset::readOnly);
	mp_bRecordsetActive = TRUE;
}

void clsDB::CloseReader()
{
	mp_oRecordset.Close();
	mp_bRecordsetActive = FALSE;
}

void clsDB::ReadTextBox(CEdit &oControl, CString sFieldName)
{
	CDBVariant vReturn;
	CString sReturn;
	CString sBuff;
	mp_oRecordset.GetFieldValue(sFieldName, vReturn, SQL_C_CHAR);
	if (vReturn.m_dwType == DBVT_WSTRING)
	{
		sBuff = *vReturn.m_pstringW;
		sReturn = sBuff.GetBuffer();
		oControl.SetWindowTextW(sReturn);
	}
	else if (vReturn.m_dwType == DBVT_ASTRING)
	{
		sBuff = *vReturn.m_pstringA;
		sReturn = sBuff.GetBuffer();
		oControl.SetWindowTextW(sReturn);
	}
	else
	{
		oControl.SetWindowTextW(_T(""));
	}
}

void clsDB::ReadCheckBox(CButton &oControl, CString sFieldName)
{
	CDBVariant vReturn;
	mp_oRecordset.GetFieldValue(sFieldName, vReturn);
	if (vReturn.m_dwType == DBVT_BOOL)
	{
		oControl.SetCheck(vReturn.m_boolVal);
	}
	else
	{
		oControl.SetCheck(FALSE);
	}
}

void clsDB::ReadComboBox(CComboBox &oControl, CString sFieldName)
{
	CDBVariant vReturn;
	mp_oRecordset.GetFieldValue(sFieldName, vReturn);
	if (vReturn.m_dwType == DBVT_LONG)
	{
		oControl.SetCurSel(GetListIndex(&oControl, vReturn.m_lVal));
	}
}

COleDateTime clsDB::ReadFieldDate(CString sFieldName)
{
	CDBVariant vReturn;
	COleDateTime dtReturn;
	mp_oRecordset.GetFieldValue(sFieldName, vReturn);
	if (vReturn.m_dwType == DBVT_DATE)
	{
		dtReturn.SetDateTime(vReturn.m_pdate->year, vReturn.m_pdate->month, vReturn.m_pdate->day, vReturn.m_pdate->hour, vReturn.m_pdate->minute, vReturn.m_pdate->second);
		return dtReturn;
	}
	else
	{
		return (DATE)0;
	}
}

double clsDB::ReadFieldDouble(CString sFieldName)
{
	CDBVariant vReturn;
	mp_oRecordset.GetFieldValue(sFieldName, vReturn);
	if (vReturn.m_dwType == DBVT_DOUBLE)
	{
		return vReturn.m_dblVal;
	}
	else if (vReturn.m_dwType == DBVT_WSTRING)
	{
		CString sReturn;
		CString sBuff;
		sBuff = *vReturn.m_pstringW;
		sReturn = sBuff.GetBuffer();
		return CDbl(sReturn);
	}
	else
	{
		return NULL;
	}
}

CString clsDB::ReadFieldString(CString sFieldName)
{
	CDBVariant vReturn;
	CString sReturn;
	CString sBuff;
	mp_oRecordset.GetFieldValue(sFieldName, vReturn, SQL_C_CHAR);
	if (vReturn.m_dwType == DBVT_WSTRING)
	{
		sBuff = *vReturn.m_pstringW;
		sReturn = sBuff.GetBuffer();
		return sReturn;
	}
	else if (vReturn.m_dwType == DBVT_ASTRING)
	{
		sBuff = *vReturn.m_pstringA;
		sReturn = sBuff.GetBuffer();
		return sReturn;
	}
	else
	{
		return _T("");
	}
}
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
#include "stdafx.h"
#include "DST_ACCESS.h"

CString g_DST_ACCESS_GetConnectionString()
{
	return _T("ODBC;DRIVER={MICROSOFT ACCESS DRIVER (*.mdb)};DSN='';DBQ=") + g_GetAppLocation() + _T("\\ActiveGanttExamples.mdb;");
}

COleDateTime GetDateField(CRecordset &oRecordset, CString sFieldName)
{
	CDBVariant vReturn;
	COleDateTime dtReturn;
	oRecordset.GetFieldValue(sFieldName, vReturn);
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

BOOL GetBoolField(CRecordset &oRecordset, CString sFieldName)
{
	CDBVariant vReturn;
	oRecordset.GetFieldValue(sFieldName, vReturn);
	if (vReturn.m_dwType == DBVT_BOOL)
	{
		return vReturn.m_boolVal;
	}
	else
	{
		return -1;
	}
}

float GetFloatField(CRecordset &oRecordset, CString sFieldName)
{
	CDBVariant vReturn;
	oRecordset.GetFieldValue(sFieldName, vReturn);
	if (vReturn.m_dwType == DBVT_SINGLE)
	{
		return vReturn.m_fltVal;
	}
	else
	{
		return -1;
	}
}

CString GetStrField(CRecordset &oRecordset, CString sFieldName)
{
	CDBVariant vReturn;
	CString sReturn;
	CString sBuff;
	oRecordset.GetFieldValue(sFieldName, vReturn, SQL_C_CHAR);
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

int GetIntField(CRecordset &oRecordset, CString sFieldName)
{
	CDBVariant vReturn;
	oRecordset.GetFieldValue(sFieldName, vReturn);
	if (vReturn.m_dwType == DBVT_LONG)
	{
		return vReturn.m_lVal;
	}
	else
	{
		return NULL;
	}
}

short GetShortField(CRecordset &oRecordset, CString sFieldName)
{
	CDBVariant vReturn;
	oRecordset.GetFieldValue(sFieldName, vReturn);
	if (vReturn.m_dwType == DBVT_SHORT)
	{
		return vReturn.m_iVal;
	}
	else
	{
		return NULL;
	}
}

double GetDblField(CRecordset &oRecordset, CString sFieldName)
{
	CDBVariant vReturn;
	oRecordset.GetFieldValue(sFieldName, vReturn);
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

CString g_DST_ACCESS_ConvertDate(COleDateTime dtDate)
{
	CString sReturn;
	sReturn = _T("#") + CStr(dtDate.GetYear()) + _T("-") + CStr(dtDate.GetMonth()) + _T("-") + CStr(dtDate.GetDay()) + _T(" ") + CStr(dtDate.GetHour()) + _T(":") + CStr(dtDate.GetMinute()) + _T(":") + CStr(dtDate.GetSecond()) + _T("#");
	return sReturn;
}

void g_DST_ACCESS_FillComboBox(CComboBox* oControl, CString sSQL, CString sValueMemberField, CString sDisplayMemberField, CDatabase &oConn)
{
	int lIndex = 0;
	CRecordset oRecordset(&oConn);
	oRecordset.Open(CRecordset::forwardOnly, sSQL, CRecordset::readOnly);
	while(oRecordset.IsEOF() == FALSE)
	{
		CString sDisplayMember;
		int lValueMember;
		sDisplayMember = GetStrField(oRecordset, sDisplayMemberField);
		lValueMember = GetIntField(oRecordset, sValueMemberField);
		oControl->AddString(sDisplayMember);
		oControl->SetItemData(lIndex, lValueMember);
		lIndex = lIndex + 1;
		oRecordset.MoveNext();
	}
	oRecordset.Close();
}

void g_DST_ACCESS_FillComboBoxString(CComboBox* oControl, CString sSQL, CString sValueMemberField, CString sDisplayMemberField, CDatabase &oConn)
{
	int lIndex = 0;
	CRecordset oRecordset(&oConn);
	oRecordset.Open(CRecordset::forwardOnly, sSQL, CRecordset::readOnly);
	while(oRecordset.IsEOF() == FALSE)
	{
		CString sDisplayMember;
		CString sValueMember;
		CString* pValueMember;
		sDisplayMember = GetStrField(oRecordset, sDisplayMemberField);
		sValueMember = GetStrField(oRecordset, sValueMemberField);
		pValueMember = new CString(sValueMember);
		oControl->AddString(sValueMember + _T("    ") + sDisplayMember);
		oControl->SetItemDataPtr(lIndex, pValueMember);
		lIndex = lIndex + 1;
		oRecordset.MoveNext();
	}
	oRecordset.Close();
}

//#define DBVT_NULL       0
//#define DBVT_BOOL       1
//#define DBVT_UCHAR      2
//#define DBVT_SHORT      3
//#define DBVT_LONG       4
//#define DBVT_SINGLE     5
//#define DBVT_DOUBLE     6
//#define DBVT_DATE       7
//#define DBVT_STRING     8
//#define DBVT_BINARY     9
//#define DBVT_ASTRING    10
//#define DBVT_WSTRING    11


//void printType ( CDBVariant* var ) 
//{ 
//  if (!var ) 
//  { 
//    cout << "NULL Var Parameter" << endl; 
//   return; 
//  } 
// 
//  switch ( var->m_dwType ) 
//  { 
//  case DBVT_NULL: 
//    { 
//      cout << "var is NULL" << endl; 
//      break; 
//    } 
//  case DBVT_BOOL: 
//    { 
//      bool b = var->m_boolVal ? true : FALSE; 
//      cout << "var is bool:" << b << endl; 
//      break; 
//    } 
//  case DBVT_UCHAR: 
//    { 
//      unsigned char c = var->m_chVal; 
//      cout << "var is unsigned char:" << c << endl; 
//      break; 
//    } 
//  case DBVT_SHORT: 
//    { 
//      short sht = var->m_iVal; 
//      cout << "var is short:" << sht << endl; 
//      break; 
//    } 
//  case DBVT_LONG: 
//    { 
//      long lng = var->m_lVal; 
//      cout << "var is long" << endl; 
//      break; 
//    } 
//  case DBVT_SINGLE: 
//    { 
//      float flt = var->m_fltVal; 
//      cout << "var is single:" << flt << endl; 
//      break; 
//    } 
//  case DBVT_DOUBLE: 
//    { 
//      double vasr = var->m_dblVal; 
//      cout << "var is double: " << vasr << endl; 
//      break; 
//    } 
//  case DBVT_DATE: 
//    { 
//      TIMESTAMP_STRUCT * structdata = var->m_pdate; 
//      cout << "var is a Date" << endl; 
//      break; 
//    } 
//  case DBVT_BINARY: 
//    { 
//      CLongBinary* binarydata = var->m_pbinary; 
//      cout << "var is Binary data" << endl; 
//      break; 
//    } 
//  case DBVT_STRING:
//    { 
//      CString* string = var->m_pstring; 
//      cout << "var is string " << string->GetBuffer() << endl; 
//      break; 
//    } 
//
//  case DBVT_ASTRING: 
//   
//    { 
//      CString* string = var->m_pstringA; 
//      cout << "var is string " << string->GetBuffer() << endl; 
//      break; 
//    } 
//  case DBVT_WSTRING: 
//    { 
//      CStringW * wstring = var->m_pstringW; 
//       
//      cout  << "var is WideString" <<endl; 
//      break; 
//    } 
//  default: 
//    { 
//      cout << "something went wrong" << endl; 
//      break; 
//    } 
//  } 
// 
//} 
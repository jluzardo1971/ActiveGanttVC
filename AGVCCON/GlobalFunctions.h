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
#include <afxdb.h>
#include "stdafx.h"
#include <afxdisp.h>

class CToolTipToolBar;


CString g_GetAppLocation();

void g_Resize(CDialog* oForm, CActiveGanttVCCtl* oControl);

void g_MaximizeWindowsDim(CWnd* oForm);

CString GetText(CWnd &oControl);
CString GetSelText(CComboBox &oControl);
LONG GetTextNum(CWnd &oControl);
LONG GetTextNum(CWnd &oControl, LONG lDefault);
COleDateTime GetTime(CDateTimeCtrl &oControl);

int Ocurrences(CString Expression, CString StringToFind);
CString* Split(CString Expression, CString Delimiter);
void SplitInterval(CString Value, CString* Interval, LONG* Factor);

LONG FromArgb(LONG A, LONG R, LONG G, LONG B);
BOOL IsNumeric(CString sParam);
LONG CInt32(CString sParam);
LONG CInt32(DOUBLE dParam);
DOUBLE CDbl(CString sParam);
CString Replace(CString Expression, CString Find, CString Replace);
CString UCase(CString sParam);
CString LCase(CString sParam);
CString CStr(int lParam);
CString CStr(double dParam);
CString CStr(double dParam, int lDecimalPlaces);
VARIANT CVar(DATE dtDate);
VARIANT CVar(FLOAT fParam);
VARIANT CVar(LONG lParam);
COleDateTime GetDateTimeNow();
COleDateTime GetDateTime(int nYear, int nMonth, int nDay, int nHour, int nMin, int nSec);
COleDateTime GetDateTime(int nYear, int nMonth, int nDay);
COleDateTime GetCurrentDateTime();
CString g_GenerateRandomName(BOOL IncludePrefix, CDatabase &oConn);
CString g_GenerateRandomAddress(CDatabase &oConn);
void g_GenerateRandomCity(CString &sCityName, CString &sState, int &ID, CDatabase &oConn);
CString g_GenerateRandomPhone(CString AreaCode);
CString g_GenerateRandomZIP();
CString g_GenerateRandomLicense();
int GetRnd(int lLowerBound, int lUpperBound);
void g_ShowInBrowser(CString sURL);
void g_OpenFile(CString sName);

int GetListIndex(CComboBox* oComboBox, int lFind);
int GetListIndex(CComboBox* oComboBox, CString sFind);

void InitNumCtrl(CSpinButtonCtrl &numCtrl, CEdit &txtCtrl, int lMin, int lMax, int lValue);
void InitNumCtrlDiv100(CSpinButtonCtrl &numCtrl, CEdit &txtCtrl, int lMin, int lMax, int lValue);

FLOAT CSng(LONG lParam);
DOUBLE CDbl(LONG lParam);

CString FormatDate(COleDateTime dtDate, CString sFormat);
CString Quarter(COleDateTime dtDate);

//Regex: char to wchar_t string literals:

//Text to find: "{(\\"|[^"])*}" (include the double quotes)
//Replace with: _T(\0)
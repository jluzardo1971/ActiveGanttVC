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
#include "GlobalFunctions.h"


#include <math.h>


CString g_GetAppLocation()
{
	CString sReturn = GetCommandLineW();
	sReturn.Replace(_T("\\Debug\\AGVCCON.exe"), _T(""));
	sReturn.Replace(_T("\\Release\\AGVCCON.exe"), _T(""));
	sReturn.Replace(_T("\""), _T(""));
	sReturn = sReturn.Trim();
	return sReturn;
}

void g_Resize(CDialog* oForm, CActiveGanttVCCtl* oControl)
{
	int lToolBarHeight = 0;
	if (oControl->m_hWnd == NULL) 
	{
		return; 
	}
	if (oForm->IsIconic() == FALSE)
	{
		CRect oClientRect;
		oForm->GetClientRect(&oClientRect);
		oForm->ScreenToClient(&oClientRect);
		if (((oClientRect.right - oClientRect.left - 80) > 100) && ((oClientRect.bottom - oClientRect.top - 80) > 100))
		{
			oControl->MoveWindow(40, 40, oClientRect.right - oClientRect.left - 80, oClientRect.bottom - oClientRect.top - 80, TRUE);				
		}
	}
}

void g_MaximizeWindowsDim(CWnd* oForm)
{
	RECT oRect;
	SystemParametersInfo(SPI_GETWORKAREA,0,&oRect,0);
	int lX = -(::GetSystemMetrics(SM_CXBORDER) + 1);
	int lY = -(::GetSystemMetrics(SM_CYBORDER) + 1);
	int lWA_X = oRect.left;
	int lWA_Y = oRect.top;
	int lWAWidth = oRect.right - oRect.left;
	int lWAHeight = oRect.bottom - oRect.top;
	::SetWindowPos(oForm->m_hWnd, HWND_TOPMOST, lX , lY , lWAWidth, lWAHeight, SWP_NOZORDER);
}

CString GetText(CWnd &oControl)
{
	CString sReturn;
	oControl.GetWindowTextW(sReturn);
	return sReturn;
}

CString GetSelText(CComboBox &oControl)
{
	CString sReturn;
	oControl.GetLBText(oControl.GetCurSel(), sReturn);
	return sReturn;
}

LONG GetTextNum(CWnd &oControl)
{
	CString sReturn;
	oControl.GetWindowTextW(sReturn);
	sReturn = sReturn.Trim();
	if (sReturn.GetLength() == 0)
	{
		return 0;
	}
	if (IsNumeric(sReturn) == FALSE)
	{
		return 0;
	}
	return CInt32(sReturn);
}

LONG GetTextNum(CWnd &oControl, LONG lDefault)
{
	CString sReturn;
	oControl.GetWindowTextW(sReturn);
	sReturn = sReturn.Trim();
	if (sReturn.GetLength() == 0)
	{
		return lDefault;
	}
	if (IsNumeric(sReturn) == FALSE)
	{
		return lDefault;
	}
	return CInt32(sReturn);
}


COleDateTime GetTime(CDateTimeCtrl &oControl)
{
	COleDateTime dtReturn;
	oControl.GetTime(dtReturn);
	return dtReturn;
}

COleDateTime GetDateTimeNow()
{
	COleDateTime dtReturn;
	SYSTEMTIME lt;
	GetLocalTime(&lt);
	dtReturn.SetDateTime(lt.wYear, lt.wMonth, lt.wDay, lt.wHour, lt.wMinute, lt.wSecond);
	return dtReturn;
}

COleDateTime GetDateTime(int nYear, int nMonth, int nDay)
{
	COleDateTime dtReturn;
	dtReturn.SetDateTime(nYear,nMonth, nDay, 0, 0, 0);
	return dtReturn;
}

COleDateTime GetDateTime(int nYear, int nMonth, int nDay, int nHour, int nMin, int nSec)
{
	COleDateTime dtReturn;
	dtReturn.SetDateTime(nYear,nMonth, nDay, nHour, nMin, nSec);
	return dtReturn;
}

COleDateTime GetCurrentDateTime()
{
	COleDateTime dtReturn;
	SYSTEMTIME lt;
	GetLocalTime(&lt);
	dtReturn.SetDateTime(lt.wYear, lt.wMonth, lt.wDay, lt.wHour, lt.wMinute, lt.wSecond);
	return dtReturn;
}

int Ocurrences(CString Expression, CString StringToFind)
{
	int lReturn = 0;
	int lPosition = 0;
	while ((lPosition != - 1) && (lPosition <= Expression.GetLength() - 1))
	{
		lPosition = Expression.Find(StringToFind, lPosition);
		if (lPosition != - 1)
		{
			lReturn = lReturn + 1;
			lPosition = lPosition + StringToFind.GetLength();
		}
	}
	return lReturn;
}

CString* Split(CString Expression, CString Delimiter)
{
	CString* aReturn;
	CString sToken=_T("");
	int i = 0;

	int lCount = Ocurrences(Expression, Delimiter);
	aReturn = new CString[lCount + 1];

	while (AfxExtractSubString(sToken, Expression, i, Delimiter[0]))
	{   
		aReturn[i] = sToken;
		i = i + 1;
	}
	return aReturn;

	////Usage:
	//
	//		CString* aArray;
	//		aArray = Split(_T("abc|def|ghi|jkl"), _T("|"));
	//		int lCount = sizeof(aArray); // lCount = 4
	//		CString a1 = aArray[0]; // a1 = "abc"
	//		CString a2 = aArray[1]; // a2 = "def"
	//		CString a3 = aArray[2]; // a3 = "ghi"
	//		CString a4 = aArray[3]; // a4 = "jkl"
	//		delete [] aArray;
}

void SplitInterval(CString Value, CString* Interval, LONG* Factor)
{
	LONG i;
	CString sInterval;
	CString sFactor;
	LONG lFactor;
	sInterval = LCase(Value);
	i = 0;
	while (IsNumeric(sInterval.Mid(i,1)) == TRUE) 
	{
		i = i + 1;
	}
	CString sTest = sInterval.Mid(0, i);
	sFactor = sInterval.Mid(0, i);
	lFactor = CInt32(sFactor);
	sInterval.Replace(sFactor, _T(""));
	if ((!(sInterval == _T("s") || sInterval == _T("n") || sInterval == _T("h") || sInterval == _T("d") || sInterval == _T("w") || sInterval == _T("y") || sInterval == _T("ww") || sInterval == _T("m") || sInterval == _T("q") || sInterval == _T("yyyy"))))
	{
		MessageBox(NULL, _T("Invalid Interval"), _T("SplitInterval"), MB_OK); 
		return;
	}
	*Interval = sInterval;
	*Factor = lFactor;
}

BOOL IsNumeric(CString sParam)
{
	int i;
	BOOL bReturn = TRUE;
	CString sNumeric = _T("0123456789.");
	if (sParam.GetLength() == 0)
	{
		bReturn = FALSE;
	}
	for (i = 0; i <= sParam.GetLength() - 1; i++)
	{
		if (sNumeric.Find(sParam[i]) == -1)
		{
			bReturn = FALSE;
		}
	}
	return bReturn;
}

LONG CInt32(DOUBLE dParam)
{
	LONG lReturn;
	if (dParam < 0.0)
	{
		lReturn = (LONG)ceil(dParam - 0.5);
	}
	else
	{
		lReturn = (LONG)floor(dParam + 0.51);
	}
	return lReturn;
}

LONG CInt32(CString sParam)
{
	LONG lReturn;
	lReturn = _wtoi(sParam);
	return lReturn;
}

DOUBLE CDbl(CString sParam)
{
	DOUBLE dReturn;
	dReturn = _wtof(sParam);
	return dReturn;
}


CString Replace(CString Expression, CString Find, CString Replace)
{
	CString sReturn;
	sReturn = Expression;
	sReturn.Replace(Find, Replace);
	return sReturn;
}

CString CStr(int lParam)
{
	CString sReturn = _T("");
	sReturn.Format(_T("%d"), lParam);
	return sReturn;
}

CString UCase(CString sParam)
{
	CString sReturn;
	sReturn = sParam;
	sReturn.MakeUpper();
	return sReturn;
}

CString LCase(CString sParam)
{
	CString sReturn;
	sReturn = sParam;
	sReturn.MakeLower();
	return sReturn;
}

CString CStr(double dParam)
{
	CString sReturn = _T("");
	sReturn.Format(_T("%.2f"), dParam);
	return sReturn;
}

CString CStr(double dParam, int lDecimalPlaces)
{
	CString sReturn = _T("");
	sReturn.Format(_T("%.") + CStr(lDecimalPlaces) + _T("f"), dParam);
	return sReturn;
}

VARIANT CVar(DATE dtDate)
{
	VARIANT oResult;
	VariantInit(&oResult);
	oResult.vt = VT_DATE;
	oResult.date = dtDate;
	return oResult;
}

VARIANT CVar(FLOAT fParam)
{
	VARIANT oResult;
	VariantInit(&oResult);
	oResult.vt = VT_R4;
	oResult.fltVal = fParam;
	return oResult;
}

VARIANT CVar(LONG lParam)
{
	VARIANT oResult;
	VariantInit(&oResult);
	oResult.vt = VT_I4;
	oResult.lVal = lParam;
	return oResult;
}

CString g_GenerateRandomName(BOOL IncludePrefix, CDatabase &oConn)
{
	CString sReturn = _T("");
	CRecordset oRecordset(&oConn);
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_First_Names WHERE lID = ") + CStr(GetRnd(1, 5490)), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{
		CString sSex;
		CString sName;
		sSex = GetStrField(oRecordset, _T("sSex"));
		sName = GetStrField(oRecordset, _T("sFirstName"));
		if (IncludePrefix == TRUE)
		{
			if (sSex == _T("F"))
			{
				sReturn = _T("Ms. ");
			}
			else
			{
				sReturn = _T("Mr. ");
			}
		}
		sReturn = sReturn + sName;
	}
	oRecordset.Close();
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_Last_Names WHERE lID = ") + CStr(GetRnd(1, 15000)), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{
		CString sLastName;
		sLastName = GetStrField(oRecordset, _T("sLastName"));
		sReturn = sReturn + _T(" ") + sLastName;
	}
	oRecordset.Close();
	return sReturn;
}

CString g_GenerateRandomAddress(CDatabase &oConn)
{
	CString sNumeric = _T("123456789");
	int lNumericMax = sNumeric.GetLength() - 1;
	CString sReturn = _T("");
	int i;

	for (i = 1; i <= 3; i++)
	{
		sReturn = sReturn + sNumeric[GetRnd(0,lNumericMax)];
	}

	CRecordset oRecordset(&oConn);
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_Last_Names WHERE lID=") + CStr(GetRnd(1, 10000)), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{
		CString sLastName;
		sLastName = GetStrField(oRecordset, _T("sLastName"));
		sReturn = sReturn + _T(" ") + sLastName;
	}
	oRecordset.Close();
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_US_Streets WHERE lID=") + CStr(GetRnd(1, 537)), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{
		CString sStreetName;
		sStreetName = GetStrField(oRecordset, _T("sStreetSuffix"));
		sReturn = sReturn + _T(" ") + sStreetName;
	}
	oRecordset.Close();
	return sReturn;
}

void g_GenerateRandomCity(CString &sCityName, CString &sState, int &ID, CDatabase &oConn)
{
	CRecordset oRecordset(&oConn);
	oRecordset.Open(CRecordset::forwardOnly, _T("SELECT * FROM tb_US_Cities WHERE lID=") + CStr(GetRnd(1, 25149)), CRecordset::readOnly);
	if (oRecordset.IsEOF() == FALSE)
	{
		sCityName = GetStrField(oRecordset, _T("sCityName"));
		sState = GetStrField(oRecordset, _T("sStateAbr"));
		ID = GetIntField(oRecordset, _T("lID"));
	}
	oRecordset.Close();
}

CString g_GenerateRandomPhone(CString AreaCode)
{
	CString sNumeric = _T("123456789");
	int lNumericMax = sNumeric.GetLength() - 1;
	CString sReturn = _T("");
	int i;
	if (AreaCode == _T(""))
	{
		for (i = 1; i <= 3; i++)
		{
			sReturn = sReturn + sNumeric[GetRnd(0,lNumericMax)];
		}
		sReturn = _T("(") + sReturn + _T(") ");
	}
	else
	{
		sReturn = AreaCode + _T(" ");
	}
	for (i = 1; i <= 3; i++)
	{
		sReturn = sReturn + sNumeric[GetRnd(0,lNumericMax)];
	}
	sReturn = sReturn + _T("-");
	for (i = 1; i <= 4; i++)
	{
		sReturn = sReturn + sNumeric[GetRnd(0,lNumericMax)];
	}
	return sReturn;
}

CString g_GenerateRandomZIP()
{
	CString sNumeric = _T("123456789");
	int lNumericMax = sNumeric.GetLength() - 1;
	CString sReturn = _T("");
	int i;
	for (i = 1; i <= 5; i++)
	{
		sReturn = sReturn + sNumeric[GetRnd(0,lNumericMax)];
	}
	return sReturn;
}

CString g_GenerateRandomLicense()
{
	CString sAlphabetic = _T("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
	int lAlphabeticMax = sAlphabetic.GetLength() - 1;
	CString sNumeric = _T("123456789");
	int lNumericMax = sNumeric.GetLength() - 1;
	CString sReturn = _T("");
	int i;
	for (i = 1; i <= 3; i++)
	{
		sReturn = sReturn + sAlphabetic[GetRnd(0,lAlphabeticMax)];
	}
	sReturn = sReturn + _T("-");
	for (i = 1; i <= 4; i++)
	{
		sReturn = sReturn + sNumeric[GetRnd(0,lNumericMax)];
	}
	return sReturn;
}

int GetRnd(int lLowerBound, int lUpperBound)
{
	return ((rand() % (int)(((lUpperBound) + 1) - (lLowerBound))) + (lLowerBound));
}

void g_ShowInBrowser(CString sURL)
{
	::ShellExecute(0, _T("open"), sURL, 0, 0, SW_SHOWNORMAL);
}

void g_OpenFile(CString sName)
{
	::ShellExecute(0, _T("open"), sName, 0, g_GetAppLocation(), SW_MAXIMIZE);
}

int GetListIndex(CComboBox* oComboBox, CString sFind)
{
	int lIndex;
	for (lIndex = 0; lIndex <= (oComboBox->GetCount() - 1); lIndex++)
	{
		CString* sData = (CString*) oComboBox->GetItemDataPtr(lIndex);
		if (*sData == sFind)
		{
			return lIndex;
		}
	}
	return -1;
}

int GetListIndex(CComboBox* oComboBox, int lFind)
{
	int lIndex;
	for (lIndex = 0; lIndex <= (oComboBox->GetCount() - 1); lIndex++)
	{
		int lData = (int) oComboBox->GetItemData(lIndex);
		if (lData == lFind)
		{
			return lIndex;
		}
	}
	return -1;
}

LONG FromArgb(LONG A, LONG R, LONG G, LONG B)
{
	return RGB(R,G,B);
}

void InitNumCtrl(CSpinButtonCtrl &numCtrl, CEdit &txtCtrl, int lMin, int lMax, int lValue)
{
	numCtrl.SetRange(lMin, lMax);
	numCtrl.SetPos32(lValue);
	numCtrl.SetBuddy(&txtCtrl);
	txtCtrl.SetWindowTextW(CStr(lValue));
}

void InitNumCtrlDiv100(CSpinButtonCtrl &numCtrl, CEdit &txtCtrl, int lMin, int lMax, int lValue)
{
	numCtrl.SetRange(lMin, lMax);
	numCtrl.SetPos32(lValue);
	numCtrl.SetBuddy(&txtCtrl);
	txtCtrl.SetWindowTextW(CStr((double)lValue / 100, 2));
}

FLOAT CSng(LONG lParam)
{
	return (FLOAT)lParam;
}

DOUBLE CDbl(LONG lParam)
{
	return (DOUBLE)lParam;
}

CString FormatDate(COleDateTime dtDate, CString sFormat)
{
	return dtDate.Format(sFormat);
}

CString Quarter(COleDateTime dtDate)
{
	int lMonth;
	lMonth = dtDate.GetMonth();
	if (lMonth >= 1 && lMonth <= 3)
	{
		return _T("1");
	}
	else if (lMonth >= 4 && lMonth <= 6)
	{
		return _T("2");
	}
	else if (lMonth >= 7 && lMonth <= 9)
	{
		return _T("3");
	}
	else if (lMonth >= 10 && lMonth <= 12)
	{
		return _T("4");
	}
	return _T("-1");
}
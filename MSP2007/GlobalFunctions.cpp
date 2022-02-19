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
#include "MSP2007Ctl.h"

CString F_TOUCSTR(LPCSTR sParam)
{
	int n = ::MultiByteToWideChar(CP_UTF8, 0, sParam, -1, NULL, 0);
	CString sReturn;
	LPWSTR p = sReturn.GetBuffer(n);
	::MultiByteToWideChar(CP_UTF8, 0, sParam, -1, p, n);
	sReturn.ReleaseBuffer();
	return sReturn;
}

BOOL g_IsNumeric(CString sParam)
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

void g_ErrorReport(LONG ErrNumber,CString ErrDescription,CString ErrSource)
{
	g_MSP2007->FireControlError(ErrNumber, ErrDescription, ErrSource);
}

LONG g_CInt(CString sParam)
{
	LONG lReturn;
	lReturn = _wtoi(sParam);
	return lReturn;
}
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
#include "GlobalDateFunctions.h"



CString mp_EscString(CString sParam)
{
	CString sResult = _T("");
	int i;
	for (i = 0; i <= sParam.GetLength() - 1; i ++)
	{
		sResult = sResult + _T("\\") + sParam.GetAt(i);
	}
	return sResult;
}



CString GetLocaleInfoEx(LCTYPE LCType)
{
	LCID Locale;
	LONG iRet1;
	LONG iRet2;
	CString Symbol;
	TCHAR lpLCDataVar[256];
	Locale = GetUserDefaultLCID();
	iRet1 = GetLocaleInfo(Locale, LCType, lpLCDataVar, 0);
	iRet2 = GetLocaleInfo(Locale, LCType, lpLCDataVar, iRet1);
	Symbol = lpLCDataVar;
	return Symbol;
}
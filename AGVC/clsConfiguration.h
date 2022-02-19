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


// clsConfiguration command target

class clsConfiguration : public CCmdTarget
{
	DECLARE_DYNCREATE(clsConfiguration)
	clsConfiguration();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsConfiguration(CActiveGanttVCCtl* oControl);
	virtual ~clsConfiguration();
	virtual void OnFinalRelease();

	LONG mp_yFirstDayOfWeek;
	LONG mp_yCalendarWeekRule;
	BOOL mp_bISO8601Weeks;
	Gdiplus::Color mp_clrRowHighlightedBackColor;
    Gdiplus::Color mp_clrRowEvenBackColor;
    Gdiplus::Color mp_clrRowOddBackColor;
    BOOL mp_bAlternatingRowStyles;
    clsFont mp_oDefaultFont;

	LONG GetFirstDayOfWeek(void);
	void SetFirstDayOfWeek(LONG value);
	LONG GetCalendarWeekRule(void);
	void SetCalendarWeekRule(LONG value);
	BOOL GetISO8601Weeks(void);
	void SetISO8601Weeks(BOOL value);

	clsFont* GetDefaultFont(void);
	Gdiplus::Color GetRowHighlightedBackColor(void);
	void SetRowHighlightedBackColor(Gdiplus::Color Value);
	Gdiplus::Color GetRowEvenBackColor(void);
	void SetRowEvenBackColor(Gdiplus::Color Value);
	Gdiplus::Color GetRowOddBackColor(void);
	void SetRowOddBackColor(Gdiplus::Color Value);
	BOOL GetAlternatingRowStyles(void);
	void SetAlternatingRowStyles(BOOL value);

	clsStyle*  GetDefaultControlStyle(void);
	clsStyle*  GetDefaultTaskStyle(void);
	clsStyle*  GetDefaultRowStyle(void);
	clsStyle*  GetDefaultClientAreaStyle(void);
	clsStyle*  GetDefaultCellStyle(void);
	clsStyle*  GetDefaultColumnStyle(void);
	clsStyle*  GetDefaultPercentageStyle(void);
	clsStyle*  GetDefaultPredecessorStyle(void);
	clsStyle*  GetDefaultTimeLineStyle(void);
	clsStyle*  GetDefaultTimeBlockStyle(void);
	clsStyle*  GetDefaultTickMarkAreaStyle(void);
	clsStyle*  GetDefaultSplitterStyle(void);
	clsStyle*  GetDefaultProgressLineStyle(void);
	clsStyle*  GetDefaultNodeStyle(void);
	clsStyle*  GetDefaultTierStyle(void);
	clsStyle*  GetDefaultScrollBarStyle(void);
	clsStyle*  GetDefaultSBSeparatorStyle(void);
	clsStyle*  GetDefaultSBNormalStyle(void);
	clsStyle*  GetDefaultSBPressedStyle(void);
	clsStyle*  GetDefaultSBDisabledStyle(void);

	CString Format(COleDateTime dtExpression, CString sFormat);



	CString GetXML(void);
	void SetXML(CString sXML);


protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsConfiguration)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()


LONG odl_GetFirstDayOfWeek(void)
{
	
	return GetFirstDayOfWeek();
}

void odl_SetFirstDayOfWeek(LONG value)
{
	
	SetFirstDayOfWeek(value);
}

LONG odl_GetCalendarWeekRule(void)
{
	
	return GetCalendarWeekRule();
}
void odl_SetCalendarWeekRule(LONG value)
{
	
	SetCalendarWeekRule(value);
}

LONG odl_GetISO8601Weeks(void)
{
	
	return GetISO8601Weeks();
}

void odl_SetISO8601Weeks(BOOL value)
{
	
	SetISO8601Weeks(value);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

IDispatch* odl_GetDefaultFont(void)
{
	
	return GetDefaultFont()->GetIDispatch(TRUE);
}

OLE_COLOR odl_GetRowHighlightedBackColor(void)
{
	
	return (OLE_COLOR) GetRowHighlightedBackColor().ToCOLORREF();
}

void odl_SetRowHighlightedBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetRowHighlightedBackColor(oColor);
}

OLE_COLOR odl_GetRowEvenBackColor(void)
{
	
	return (OLE_COLOR) GetRowEvenBackColor().ToCOLORREF();
}

void odl_SetRowEvenBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetRowEvenBackColor(oColor);
}

OLE_COLOR odl_GetRowOddBackColor(void)
{
	
	return (OLE_COLOR) GetRowOddBackColor().ToCOLORREF();
}

void odl_SetRowOddBackColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetRowOddBackColor(oColor);
}

LONG odl_GetAlternatingRowStyles(void)
{
	
	return GetAlternatingRowStyles();
}

void odl_SetAlternatingRowStyles(BOOL value)
{
	
	SetAlternatingRowStyles(value);
}

IDispatch* odl_GetDefaultControlStyle(void)
{
	
	return GetDefaultControlStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultTaskStyle(void)
{
	
	return GetDefaultTaskStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultRowStyle(void)
{
	
	return GetDefaultRowStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultClientAreaStyle(void)
{
	
	return GetDefaultClientAreaStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultCellStyle(void)
{
	return GetDefaultCellStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultColumnStyle(void)
{
	return GetDefaultColumnStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultPercentageStyle(void)
{
	return GetDefaultPercentageStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultPredecessorStyle(void)
{
	return GetDefaultPredecessorStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultTimeLineStyle(void)
{
	return GetDefaultTimeLineStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultTimeBlockStyle(void)
{
	return GetDefaultTimeBlockStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultTickMarkAreaStyle(void)
{
	return GetDefaultTickMarkAreaStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultSplitterStyle(void)
{
	return GetDefaultSplitterStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultProgressLineStyle(void)
{
	return GetDefaultProgressLineStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultNodeStyle(void)
{
	return GetDefaultNodeStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultTierStyle(void)
{
	return GetDefaultTierStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultScrollBarStyle(void)
{
	return GetDefaultScrollBarStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultSBSeparatorStyle(void)
{
	return GetDefaultSBSeparatorStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultSBNormalStyle(void)
{
	return GetDefaultSBNormalStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultSBPressedStyle(void)
{
	return GetDefaultSBPressedStyle()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDefaultSBDisabledStyle(void)
{
	return GetDefaultSBDisabledStyle()->GetIDispatch(TRUE);
}

};
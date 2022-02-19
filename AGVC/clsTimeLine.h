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

class clsStyle;
class clsView;
class clsTimeLineScrollBar;
class clsTierArea;
class clsTickMarkArea;
class clsProgressLine;

class clsTimeLine : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTimeLine)
	clsTimeLine();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTimeLine(CActiveGanttVCCtl* oControl, clsView* oView);
	virtual ~clsTimeLine();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsView* mp_oView;
	CString mp_sStyleIndex;
	Gdiplus::Color mp_clrForeColor;
	COleDateTime mp_dtStartDate;
	LONG mp_lEnd;
	LONG mp_lStart;
	clsStyle* mp_oStyle;
	COleDateTime mp_dtEndDate;				
	LONG mp_lHeight;				
	LONG mp_lTop;				
	LONG mp_lBottom;				
	clsTimeLineScrollBar* mp_oTimeLineScrollBar;				
	clsTierArea* mp_oTierArea;				
	clsTickMarkArea* mp_oTickMarkArea;				
	clsProgressLine* mp_oProgressLine;

//Internal Properties (Extensions to ODL)
public:
	CString GetStyleIndex(void);
	void SetStyleIndex(CString Value);
	clsStyle* GetStyle(void);
	Gdiplus::Color GetForeColor(void);
	void SetForeColor(Gdiplus::Color Value);
	COleDateTime GetEndDate(void);
	COleDateTime GetStartDate(void);
	LONG GetHeight(void);
	LONG GetTop(void);
	LONG GetBottom(void);
	clsTimeLineScrollBar* GetTimeLineScrollBar(void);
	clsTierArea* GetTierArea(void);
	clsTickMarkArea* GetTickMarkArea(void);
	clsProgressLine* GetProgressLine(void);

//Internal Properties
public:
	void Setf_StartDate(COleDateTime Value);
	LONG Getf_lStart(void);
	LONG Getf_lEnd(void);

//Internal Methods
public:
	void Move(LONG Interval, LONG Factor);
	void Position(COleDateTime TimeLineStartDate);
	void Calculate(void);
	LONG TiersTickMarksPosition(CString v_yTier);
	void Draw(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTimeLine* oClone);


protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTimeLine)
	DECLARE_INTERFACE_MAP()

BSTR odl_GetStyleIndex(void)
{
	
	return GetStyleIndex().AllocSysString();
}

void odl_SetStyleIndex(LPCTSTR Value)
{
	
	SetStyleIndex(Value);
}

IDispatch* odl_GetStyle(void)
{
	
	return GetStyle()->GetIDispatch(TRUE);
}

OLE_COLOR odl_GetForeColor(void)
{
	
	return (OLE_COLOR) GetForeColor().ToCOLORREF();
}

void odl_SetForeColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetForeColor(oColor);
}

DATE odl_GetEndDate(void)
{	
	return GetEndDate();
}

DATE odl_GetStartDate(void)
{
	
	return mp_dtStartDate;
}

LONG odl_GetHeight(void)
{
	
	return GetHeight();
}

LONG odl_GetTop(void)
{
	
	return GetTop();
}

LONG odl_GetBottom(void)
{
	
	return GetBottom();
}

IDispatch* odl_GetTimeLineScrollBar(void)
{
	
	return GetTimeLineScrollBar()->GetIDispatch(TRUE);
}

IDispatch* odl_GetTierArea(void)
{
	
	return GetTierArea()->GetIDispatch(TRUE);
}

IDispatch* odl_GetTickMarkArea(void)
{
	
	return GetTickMarkArea()->GetIDispatch(TRUE);
}

IDispatch* odl_GetProgressLine(void)
{
	
	return GetProgressLine()->GetIDispatch(TRUE);
}

void odl_Move(LONG Interval, LONG Factor)
{
	
	Move(Interval, Factor);
}

void odl_Position(DATE TimeLineStartDate)
{
	Position(TimeLineStartDate);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}


};
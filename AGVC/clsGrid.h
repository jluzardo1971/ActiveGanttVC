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

class clsTimeLine;

class clsGrid : public CCmdTarget
{
	DECLARE_DYNCREATE(clsGrid)
	clsGrid();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsGrid(CActiveGanttVCCtl* oControl, clsTimeLine* oTimeLine);

	virtual ~clsGrid();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bVerticalLines;
	BOOL mp_bSnapToGrid;
	BOOL mp_bSnapToGridOnSelection;
	Gdiplus::Color mp_clrColor;
	LONG mp_yInterval;
	LONG mp_lFactor;
	clsTimeLine* mp_oTimeLine;

//Internal Properties (Extensions to ODL)
public:
	BOOL GetVerticalLines(void);
	void SetVerticalLines(BOOL Value);
	BOOL GetSnapToGrid(void);
	void SetSnapToGrid(BOOL Value);
	BOOL GetSnapToGridOnSelection(void);
	void SetSnapToGridOnSelection(BOOL Value);
	Gdiplus::Color GetColor(void);
	void SetColor(Gdiplus::Color Value);
	LONG GetInterval(void);
	void SetInterval(LONG Value);
	LONG GetFactor(void);
	void SetFactor(LONG Value);

//Internal Properties
public:


//Internal Methods
public:
	void Finalize(void);
	void Draw(void);
	void mp_PaintVerticalGridLine(LONG fXCoordinate,LONG v_lDrawStyle);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsGrid* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsGrid)
	DECLARE_INTERFACE_MAP()

BOOL odl_GetVerticalLines(void)
{
	
	return GetVerticalLines();
}

void odl_SetVerticalLines(BOOL Value)
{
	
	
	SetVerticalLines(Value);
}

BOOL odl_GetSnapToGrid(void)
{
	
	return GetSnapToGrid();
}

void odl_SetSnapToGrid(BOOL Value)
{
	
	
	SetSnapToGrid(Value);
}

BOOL odl_GetSnapToGridOnSelection(void)
{
	
	return GetSnapToGridOnSelection();
}

void odl_SetSnapToGridOnSelection(BOOL Value)
{
	
	
	SetSnapToGridOnSelection(Value);
}

OLE_COLOR odl_GetColor(void)
{
	
	return (OLE_COLOR) GetColor().ToCOLORREF();
}

void odl_SetColor(OLE_COLOR Value)
{
	
	Gdiplus::Color oColor;
	oColor.SetFromCOLORREF(g_TranslateColor(Value));
	SetColor(oColor);
}

LONG odl_GetInterval(void)
{
	
	return GetInterval();
}

void odl_SetInterval(LONG Value)
{
	
	SetInterval(Value);
}

LONG odl_GetFactor(void)
{
	
	return GetFactor();
}

void odl_SetFactor(LONG Value)
{
	
	SetFactor(Value);
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
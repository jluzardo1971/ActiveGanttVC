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


// clsGDIGraphics command target

class clsGDIGraphics : public CCmdTarget
{
	DECLARE_DYNCREATE(clsGDIGraphics)
	clsGDIGraphics();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsGDIGraphics(CActiveGanttVCCtl* oControl);
	virtual ~clsGDIGraphics();
	virtual void OnFinalRelease();

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsGDIGraphics)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

public:
	LONG mp_lX1;
	LONG mp_lX2;
	LONG mp_lY1;
	LONG mp_lY2;

	OLE_HANDLE GetHDC(void);
	void ReleaseHDC(OLE_HANDLE hdc);
	LONG StringWidth(LPCTSTR Text, IDispatch* Font, LONG hWnd);
	LONG StringHeight(LPCTSTR Text, IDispatch* Font, LONG hWnd);
	void SetLayoutRect(LONG X1, LONG Y1, LONG X2, LONG Y2);
};
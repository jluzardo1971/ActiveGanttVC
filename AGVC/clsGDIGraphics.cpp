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
#include "ActiveGanttVC.h"
#include "ActiveGanttVCCtl.h"
#include "clsGDIGraphics.h"




// clsGDIGraphics

IMPLEMENT_DYNCREATE(clsGDIGraphics, CCmdTarget)

clsGDIGraphics::clsGDIGraphics()
{
	EnableAutomation();
	AfxOleLockApp();
}


clsGDIGraphics::clsGDIGraphics(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_lX1 = 0;
	mp_lX2 = 0;
	mp_lY1 = 0;
	mp_lY2 = 0;
}

clsGDIGraphics::~clsGDIGraphics()
{
	AfxOleUnlockApp();
}


void clsGDIGraphics::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(clsGDIGraphics, CCmdTarget)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(clsGDIGraphics, CCmdTarget)
	DISP_FUNCTION_ID(clsGDIGraphics, "GetHDC", 1, GetHDC, VT_HANDLE, VTS_NONE)
	DISP_FUNCTION_ID(clsGDIGraphics, "ReleaseHDC", 2, ReleaseHDC, VT_EMPTY, VTS_HANDLE)
	DISP_FUNCTION_ID(clsGDIGraphics, "StringWidth", 3, StringWidth, VT_I4, VTS_BSTR VTS_DISPATCH VTS_I4)
	DISP_FUNCTION_ID(clsGDIGraphics, "StringHeight", 4, StringHeight, VT_I4, VTS_BSTR VTS_DISPATCH VTS_I4)
	DISP_FUNCTION_ID(clsGDIGraphics, "SetLayoutRect", 5, SetLayoutRect, VT_EMPTY, VTS_I4 VTS_I4 VTS_I4 VTS_I4)
END_DISPATCH_MAP()

// {F9E3BAED-B034-45D5-BB53-085B2B301AAC}
static const IID IID_IclsGDIGraphics = { 0xF9E3BAED, 0xB034, 0x45D5, { 0xBB, 0x53, 0x08, 0x5B, 0x2B, 0x30, 0x1A, 0xAC} };

BEGIN_INTERFACE_MAP(clsGDIGraphics, CCmdTarget)
	INTERFACE_PART(clsGDIGraphics, IID_IclsGDIGraphics, Dispatch)
END_INTERFACE_MAP()

// {3D6B8B4A-D9C0-4EF3-B327-C12C115A5EC6}
IMPLEMENT_OLECREATE_FLAGS(clsGDIGraphics, "AGVC.clsGDIGraphics", afxRegApartmentThreading, 0x3D6B8B4A, 0xD9C0, 0x4EF3, 0xB3, 0x27, 0xC1, 0x2C, 0x11, 0x5A, 0x5E, 0xC6)


// clsGDIGraphics message handlers

OLE_HANDLE clsGDIGraphics::GetHDC(void)
{
	
	return (OLE_HANDLE)mp_oControl->mp_oGraphics->GetHDC();
}

void clsGDIGraphics::ReleaseHDC(OLE_HANDLE hdc)
{
	
	mp_oControl->mp_oGraphics->ReleaseHDC((HDC)hdc);
}

LONG clsGDIGraphics::StringWidth(LPCTSTR Text, IDispatch* Font, LONG hWnd)
{
	Gdiplus::Graphics oGraphics((HWND)hWnd);
	clsFont* oFont = g_GetFontFromIDispatch(Font);
	CString sText = Text;
	Gdiplus::RectF layoutRect((REAL)mp_lX1, (REAL)mp_lX2, (REAL)(mp_lX2 - mp_lX1), (REAL)(mp_lY2 - mp_lY1));
	Gdiplus::RectF boundRect;
	oGraphics.MeasureString(sText, sText.GetLength(), oFont->GetFontP(), layoutRect, &boundRect);
	delete oFont;
	oFont = NULL;
	return ((LONG)boundRect.GetRight() - (LONG)boundRect.GetLeft());
}
LONG clsGDIGraphics::StringHeight(LPCTSTR Text, IDispatch* Font, LONG hWnd)
{
	Gdiplus::Graphics oGraphics((HWND)hWnd);
	clsFont* oFont = g_GetFontFromIDispatch(Font);
	CString sText = Text;
	Gdiplus::RectF layoutRect((REAL)mp_lX1, (REAL)mp_lX2, (REAL)(mp_lX2 - mp_lX1), (REAL)(mp_lY2 - mp_lY1));
	Gdiplus::RectF boundRect;
	oGraphics.MeasureString(sText, sText.GetLength(), oFont->GetFontP(), layoutRect, &boundRect);
	delete oFont;
	oFont = NULL;
	return ((LONG)boundRect.GetBottom() - (LONG)boundRect.GetTop());
}

void clsGDIGraphics::SetLayoutRect(LONG X1, LONG Y1, LONG X2, LONG Y2)
{
	mp_lX1 = X1;
	mp_lX2 = X2;
	mp_lY1 = Y1;
	mp_lY2 = Y2;
}
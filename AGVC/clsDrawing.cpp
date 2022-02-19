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
#include "clsXML.h"
#include "clsDrawing.h"
#include "clsTextFlags.h"



IMPLEMENT_DYNCREATE(clsDrawing, CCmdTarget)

//{8D98D82E-E066-48EE-A597-75B65301130A}
static const IID IID_IclsDrawing = { 0x8D98D82E, 0xE066, 0x48EE, { 0xA5, 0x97, 0x75, 0xB6, 0x53, 0x01, 0x13, 0x0A} };

//{F8B85729-648D-43CF-8E88-05A9898A7FE9}
IMPLEMENT_OLECREATE_FLAGS(clsDrawing, "AGVC.clsDrawing", afxRegApartmentThreading, 0xF8B85729, 0x648D, 0x43CF, 0x8E, 0x88, 0x05, 0xA9, 0x89, 0x8A, 0x7F, 0xE9)

BEGIN_DISPATCH_MAP(clsDrawing, CCmdTarget)
	DISP_FUNCTION_ID(clsDrawing, "GraphicsInfo", 1, odl_GraphicsInfo, VT_DISPATCH, VTS_NONE)
	DISP_FUNCTION_ID(clsDrawing, "DrawLine", 2, odl_DrawLine, VT_EMPTY, VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_COLOR VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsDrawing, "DrawBorder", 3, odl_DrawBorder, VT_EMPTY, VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_COLOR VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsDrawing, "DrawRectangle", 4, odl_DrawRectangle, VT_EMPTY, VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_COLOR VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsDrawing, "DrawText", 5, odl_DrawText, VT_EMPTY, VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR VTS_COLOR VTS_DISPATCH) 
	DISP_FUNCTION_ID(clsDrawing, "DrawAlignedText", 6, odl_DrawAlignedText, VT_EMPTY, VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR VTS_I4 VTS_I4 VTS_COLOR VTS_DISPATCH) 
	DISP_FUNCTION_ID(clsDrawing, "PaintImage", 7, odl_PaintImage, VT_EMPTY, VTS_DISPATCH VTS_I4 VTS_I4 VTS_I4 VTS_I4)
	DISP_PROPERTY_EX_ID(clsDrawing, "TextFlags", 8, odl_GetTextFlags, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsDrawing, "Image", 9, odl_GetImage, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsDrawing, "Font", 10, odl_GetFont, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(clsDrawing, "DrawObject", 11, odl_DrawObject, VT_EMPTY, VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR VTS_BSTR VTS_BOOL VTS_DISPATCH VTS_I4)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsDrawing, CCmdTarget)
	INTERFACE_PART(clsDrawing, IID_IclsDrawing, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsDrawing, CCmdTarget)
END_MESSAGE_MAP()

clsDrawing::clsDrawing()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsDrawing. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsDrawing::clsDrawing(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oImage = new clsImage();
	mp_oTextFlags = new clsTextFlags(mp_oControl);
}

clsDrawing::~clsDrawing()
{
	delete mp_oImage;
	mp_oImage = NULL;
	delete mp_oTextFlags;
	mp_oTextFlags = NULL;
	AfxOleUnlockApp();
}

void clsDrawing::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

clsGDIGraphics* clsDrawing::GraphicsInfo(void)
{
	return &mp_oGraphics;
}

void clsDrawing::DrawLine(LONG X1,LONG Y1,LONG X2,LONG Y2,Gdiplus::Color LineColor,LONG LineStyle,LONG LineWidth)
{
	mp_oControl->clsG->mp_DrawLine(X1, Y1, X2, Y2, LT_NORMAL, LineColor, (GRE_LINEDRAWSTYLE)LineStyle);
}

void clsDrawing::DrawBorder(LONG X1,LONG Y1,LONG X2,LONG Y2,Gdiplus::Color LineColor,LONG LineStyle,LONG LineWidth)
{
	mp_oControl->clsG->mp_DrawLine(X1, Y1, X2, Y2, LT_BORDER, LineColor, (GRE_LINEDRAWSTYLE)LineStyle);
}

void clsDrawing::DrawRectangle(LONG X1,LONG Y1,LONG X2,LONG Y2,Gdiplus::Color LineColor,LONG LineStyle,LONG LineWidth)
{
	mp_oControl->clsG->mp_DrawLine(X1, Y1, X2, Y2, LT_FILLED, LineColor, (GRE_LINEDRAWSTYLE)LineStyle);
}

void clsDrawing::DrawText(LONG X1,LONG Y1,LONG X2,LONG Y2,CString Text,StringFormat* oFlags, Gdiplus::Color TextColor,clsFont* TextFont)
{
	mp_oControl->clsG->mp_DrawTextEx(X1, Y1, X2, Y2, Text, oFlags, TextColor, TextFont->GetFontP(), FALSE);
}

void clsDrawing::DrawAlignedText(LONG X1,LONG Y1,LONG X2,LONG Y2,CString Text,LONG HorizontalAlignment,LONG VerticalAlignment,Gdiplus::Color TextColor,clsFont* TextFont)
{
	mp_oControl->clsG->mp_DrawAlignedText(X1,Y1,X2,Y2, Text, (GRE_HORIZONTALALIGNMENT) HorizontalAlignment, (GRE_VERTICALALIGNMENT) VerticalAlignment, TextColor, TextFont, FALSE);
}

void clsDrawing::PaintImage(clsImage* Image,LONG X1,LONG Y1,LONG X2,LONG Y2)
{
	mp_oControl->clsG->mp_PaintImage(Image,X1,Y1,X2,Y2,0,0,FALSE);
}

clsTextFlags* clsDrawing::GetTextFlags(void)
{
	return mp_oTextFlags;
}

clsImage* clsDrawing::GetImage(void)
{
	return mp_oImage;
}

clsFont* clsDrawing::GetFont(void)
{
	return &mp_oFont;
}

void clsDrawing::DrawObject(LONG X1, LONG Y1, LONG X2, LONG Y2, CString StyleIndex, CString Text, BOOL Selected, clsImage* Image, LONG ObjectType)
{ 
	if (ObjectType == DO_GENERAL)
	{
		mp_oControl->clsG->mp_DrawItem(X1, Y1, X2, Y2, StyleIndex, Text, Selected, Image, 0, 0, NULL);
	}
	else if (ObjectType == DO_MILESTONE)
	{
		mp_oControl->clsG->mp_DrawItemI(X1, Y1, X2, Y2, StyleIndex, Text, Selected, Image, 0, 0, NULL);
	}
}
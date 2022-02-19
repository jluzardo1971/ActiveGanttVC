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
#include "PagePrintEventArgs.h"

IMPLEMENT_DYNCREATE(PagePrintEventArgs, CCmdTarget)

//{B744D7D6-DB94-4A0E-A4C1-6A9FADF54A68}
static const IID IID_IPagePrintEventArgs = { 0xB744D7D6, 0xDB94, 0x4A0E, { 0xA4, 0xC1, 0x6A, 0x9F, 0xAD, 0xF5, 0x4A, 0x68} };

//{A6872896-C66A-461D-AD58-728D932E1B38}
IMPLEMENT_OLECREATE_FLAGS(PagePrintEventArgs, "AGVC.PagePrintEventArgs", afxRegApartmentThreading, 0xA6872896, 0xC66A, 0x461D, 0xAD, 0x58, 0x72, 0x8D, 0x93, 0x2E, 0x1B, 0x38)

BEGIN_DISPATCH_MAP(PagePrintEventArgs, CCmdTarget)
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "PageNumber", 1, odl_GetPageNumber, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "Graphics", 2, odl_GetGraphics, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "X1", 3, odl_GetX1, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "Y1", 4, odl_GetY1, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "X2", 5, odl_GetX2, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "Y2", 6, odl_GetY2, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "PageWidth", 7, odl_GetPageWidth, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "PageHeight", 8, odl_GetPageHeight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "ActualPageWidth", 9, odl_GetActualPageWidth, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(PagePrintEventArgs, "ActualPageHeight", 10, odl_GetActualPageHeight, SetNotSupported, VT_I4) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(PagePrintEventArgs, CCmdTarget)
	INTERFACE_PART(PagePrintEventArgs, IID_IPagePrintEventArgs, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(PagePrintEventArgs, CCmdTarget)
END_MESSAGE_MAP()

PagePrintEventArgs::PagePrintEventArgs()
{
	EnableAutomation();
	AfxOleLockApp();
	Clear();
}

PagePrintEventArgs::~PagePrintEventArgs()
{
	AfxOleUnlockApp();
}

void PagePrintEventArgs::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

void PagePrintEventArgs::Clear(void)
{
	mp_lPageNumber = 0;
	mp_lX1 = 0;
	mp_lY1 = 0;
	mp_lX2 = 0;
	mp_lY2 = 0;
	mp_lPageWidth = 0;
	mp_lPageHeight = 0;
	mp_lActualPageWidth = 0;
	mp_lActualPageHeight = 0;
}

LONG PagePrintEventArgs::GetPageNumber(void)
{
	return mp_lPageNumber;
}

void PagePrintEventArgs::SetPageNumber(LONG value)
{
	mp_lPageNumber = value;
}

clsGDIGraphics* PagePrintEventArgs::GetGraphics(void)
{
	return &mp_oGraphics;
}

LONG PagePrintEventArgs::GetX1(void)
{
	return mp_lX1;
}

void PagePrintEventArgs::SetX1(LONG value)
{
	mp_lX1 = value;
}

LONG PagePrintEventArgs::GetY1(void)
{
	return mp_lY1;
}

void PagePrintEventArgs::SetY1(LONG value)
{
	mp_lY1 = value;
}

LONG PagePrintEventArgs::GetX2(void)
{
	return mp_lX2;
}

void PagePrintEventArgs::SetX2(LONG value)
{
	mp_lX2 = value;
}

LONG PagePrintEventArgs::GetY2(void)
{
	return mp_lY2;
}

void PagePrintEventArgs::SetY2(LONG value)
{
	mp_lY2 = value;
}

LONG PagePrintEventArgs::GetPageWidth(void)
{
	return mp_lPageWidth;
}

void PagePrintEventArgs::SetPageWidth(LONG value)
{
	mp_lPageWidth = value;
}

LONG PagePrintEventArgs::GetPageHeight(void)
{
	return mp_lPageHeight;
}

void PagePrintEventArgs::SetPageHeight(LONG value)
{
	mp_lPageHeight = value;
}

LONG PagePrintEventArgs::GetActualPageWidth(void)
{
	return mp_lActualPageWidth;
}

void PagePrintEventArgs::SetActualPageWidth(LONG value)
{
	mp_lActualPageWidth = value;
}

LONG PagePrintEventArgs::GetActualPageHeight(void)
{
	return mp_lActualPageHeight;
}

void PagePrintEventArgs::SetActualPageHeight(LONG value)
{
	mp_lActualPageHeight = value;
}
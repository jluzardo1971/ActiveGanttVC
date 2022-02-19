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

class clsGDIGraphics;

class PagePrintEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(PagePrintEventArgs)

public:
	PagePrintEventArgs();
	virtual ~PagePrintEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lPageNumber;
	clsGDIGraphics mp_oGraphics;
	LONG mp_lX1;
	LONG mp_lY1;
	LONG mp_lX2;
	LONG mp_lY2;
	LONG mp_lPageWidth;
	LONG mp_lPageHeight;
	LONG mp_lActualPageWidth;
	LONG mp_lActualPageHeight;

// Properties
public:
	LONG GetPageNumber(void);
	void SetPageNumber(LONG value);

	clsGDIGraphics* GetGraphics(void);

	LONG GetX1(void);
	void SetX1(LONG value);

	LONG GetY1(void);
	void SetY1(LONG value);

	LONG GetX2(void);
	void SetX2(LONG value);

	LONG GetY2(void);
	void SetY2(LONG value);

	LONG GetPageWidth(void);
	void SetPageWidth(LONG value);

	LONG GetPageHeight(void);
	void SetPageHeight(LONG value);

	LONG GetActualPageWidth(void);
	void SetActualPageWidth(LONG value);

	LONG GetActualPageHeight(void);
	void SetActualPageHeight(LONG value);

//Internal Methods
public:
	void Clear(void);



protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(PagePrintEventArgs)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()


// odl:

LONG odl_GetPageNumber(void)
{
	
	return GetPageNumber();
}

LONG odl_GetX1(void)
{
	
	return GetX1();
}

LONG odl_GetY1(void)
{
	
	return GetY1();
}

LONG odl_GetX2(void)
{
	
	return GetX2();
}

LONG odl_GetY2(void)
{
	
	return GetY2();
}

LONG odl_GetPageWidth(void)
{
	
	return GetPageWidth();
}

LONG odl_GetPageHeight(void)
{
	
	return GetPageHeight();
}

LONG odl_GetActualPageWidth(void)
{
	
	return GetActualPageWidth();
}

LONG odl_GetActualPageHeight(void)
{
	
	return GetActualPageHeight();
}

IDispatch* odl_GetGraphics(void)
{
	
	return GetGraphics()->GetIDispatch(TRUE);
}




};
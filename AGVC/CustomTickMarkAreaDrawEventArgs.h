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

class CustomTickMarkAreaDrawEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(CustomTickMarkAreaDrawEventArgs)

public:
	CustomTickMarkAreaDrawEventArgs(void);
	virtual ~CustomTickMarkAreaDrawEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bCustomDraw;
	clsGDIGraphics mp_oGraphics;
	LONG mp_lLeft;
	LONG mp_lTop;
	LONG mp_lRight;
	LONG mp_lBottom;
	COleDateTime mp_dtDate;
	clsTickMark* mp_oTickMark;
	LONG mp_lX;

//Internal Properties (Extensions to ODL)
public:


//Internal Properties
public:

	BOOL GetCustomDraw(void);
	void SetCustomDraw(BOOL value);
	void SetLeft(LONG value);
	void SetTop(LONG value);
	void SetRight(LONG value);
	void SetBottom(LONG value);
	void SetdtDate(COleDateTime value);
	void SetoTickMark(clsTickMark* value);
	void SetX(LONG value);

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CustomTickMarkAreaDrawEventArgs)
	DECLARE_INTERFACE_MAP()

BOOL odl_GetCustomDraw(void)
{
	
	return mp_bCustomDraw;
}

void odl_SetCustomDraw(BOOL Value)
{
	
	mp_bCustomDraw = Value;
}

IDispatch* odl_GetGraphics(void)
{
	
	return mp_oGraphics.GetIDispatch(TRUE);
}

LONG odl_GetLeft(void)
{
	
	return mp_lLeft;
}

LONG odl_GetTop(void)
{
	
	return mp_lTop;
}

LONG odl_GetRight(void)
{
	
	return mp_lRight;
}

LONG odl_GetBottom(void)
{
	
	return mp_lBottom;
}

DATE odl_GetdtDate(void)
{
	
	return mp_dtDate;
}

IDispatch* odl_GetoTickMark(void)
{
	
	return mp_oTickMark->GetIDispatch(TRUE);
}

LONG odl_GetX(void)
{
	
	return mp_lX;
}


};
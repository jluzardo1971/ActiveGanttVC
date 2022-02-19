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


class MouseWheelEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(MouseWheelEventArgs)

public:
	MouseWheelEventArgs(void);
	virtual ~MouseWheelEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_lX;				
	LONG mp_lY;				
	LONG mp_yButton;				
	LONG mp_lDelta;

//Internal Properties (Extensions to ODL)
public:
	LONG GetX(void);
	void SetX(LONG Value);
	LONG GetY(void);
	void SetY(LONG Value);
	LONG GetButton(void);
	void SetButton(LONG Value);
	LONG GetDelta(void);
	void SetDelta(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(MouseWheelEventArgs)
	DECLARE_INTERFACE_MAP()

LONG odl_GetX(void)
{
	
	return GetX();
}

void odl_SetX(LONG Value)
{
	
	SetX(Value);
}

LONG odl_GetY(void)
{
	
	return GetY();
}

void odl_SetY(LONG Value)
{
	
	SetY(Value);
}

LONG odl_GetButton(void)
{
	
	return GetButton();
}

void odl_SetButton(LONG Value)
{
	
	SetButton(Value);
}

LONG odl_GetDelta(void)
{
	
	return GetDelta();
}

};
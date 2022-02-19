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


class ScrollEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(ScrollEventArgs)

public:
	ScrollEventArgs(void);
	virtual ~ScrollEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_yScrollBarType;				
	LONG mp_lOffset;			

//Internal Properties (Extensions to ODL)
public:
	LONG GetScrollBarType(void);
	void SetScrollBarType(LONG Value);
	LONG GetOffset(void);
	void SetOffset(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(ScrollEventArgs)
	DECLARE_INTERFACE_MAP()


LONG odl_GetScrollBarType(void)
{
	
	return GetScrollBarType();
}

void odl_SetScrollBarType(LONG Value)
{
	
	SetScrollBarType(Value);
}

LONG odl_GetOffset(void)
{
	
	return GetOffset();
}

void odl_SetOffset(LONG Value)
{
	
	SetOffset(Value);
}

};
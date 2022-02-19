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

class DrawEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(DrawEventArgs)

public:
	DrawEventArgs(void);
	virtual ~DrawEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_yEventTarget;				
	BOOL mp_bCustomDraw;				
	LONG mp_lObjectIndex;				
	LONG mp_lParentObjectIndex;				
	clsGDIGraphics mp_oGraphics;				

//Internal Properties (Extensions to ODL)
public:
	LONG GetEventTarget(void);
	void SetEventTarget(LONG Value);
	BOOL GetCustomDraw(void);
	void SetCustomDraw(BOOL Value);
	LONG GetObjectIndex(void);
	void SetObjectIndex(LONG Value);
	LONG GetParentObjectIndex(void);
	void SetParentObjectIndex(LONG Value);
	clsGDIGraphics* GetGraphics(void);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(DrawEventArgs)
	DECLARE_INTERFACE_MAP()

LONG odl_GetEventTarget(void)
{
	
	return GetEventTarget();
}

void odl_SetEventTarget(LONG Value)
{
	
	SetEventTarget(Value);
}

BOOL odl_GetCustomDraw(void)
{
	
	return GetCustomDraw();
}

void odl_SetCustomDraw(BOOL Value)
{
	
	
	SetCustomDraw(Value);
}

LONG odl_GetObjectIndex(void)
{
	
	return GetObjectIndex();
}

void odl_SetObjectIndex(LONG Value)
{
	
	SetObjectIndex(Value);
}

LONG odl_GetParentObjectIndex(void)
{
	
	return GetParentObjectIndex();
}

void odl_SetParentObjectIndex(LONG Value)
{
	
	SetParentObjectIndex(Value);
}

IDispatch* odl_GetGraphics(void)
{
	
	return GetGraphics()->GetIDispatch(TRUE);
}

};
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

class PredecessorDrawEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(PredecessorDrawEventArgs)

public:
	PredecessorDrawEventArgs(void);
	virtual ~PredecessorDrawEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	BOOL mp_bCustomDraw;				
	clsGDIGraphics mp_oGraphics;				
	LONG mp_lPredecessorIndex;			
	LONG mp_lTaskIndex;				
	LONG mp_yPredecessorType;				

//Internal Properties (Extensions to ODL)
public:
	BOOL GetCustomDraw(void);
	void SetCustomDraw(BOOL Value);
	clsGDIGraphics* GetGraphics(void);
	LONG GetPredecessorIndex(void);
	void SetPredecessorIndex(LONG Value);
	LONG GetTaskIndex(void);
	void SetTaskIndex(LONG Value);
	LONG GetPredecessorType(void);
	void SetPredecessorType(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(PredecessorDrawEventArgs)
	DECLARE_INTERFACE_MAP()

BOOL odl_GetCustomDraw(void)
{
	
	return GetCustomDraw();
}

void odl_SetCustomDraw(BOOL Value)
{
	
	
	SetCustomDraw(Value);
}

IDispatch* odl_GetGraphics(void)
{
	
	return GetGraphics()->GetIDispatch(TRUE);
}

LONG odl_GetPredecessorIndex(void)
{
	
	return GetPredecessorIndex();
}

void odl_SetPredecessorIndex(LONG Value)
{
	
	SetPredecessorIndex(Value);
}

LONG odl_GetTaskIndex(void)
{
	
	return GetTaskIndex();
}

void odl_SetTaskIndex(LONG Value)
{
	
	SetTaskIndex(Value);
}

LONG odl_GetPredecessorType(void)
{
	
	return GetPredecessorType();
}

void odl_SetPredecessorType(LONG Value)
{
	
	SetPredecessorType(Value);
}

};
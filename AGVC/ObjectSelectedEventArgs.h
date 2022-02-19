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


class ObjectSelectedEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(ObjectSelectedEventArgs)

public:
	ObjectSelectedEventArgs(void);
	virtual ~ObjectSelectedEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_yEventTarget;				
	LONG mp_lObjectIndex;				
	LONG mp_lParentObjectIndex;				

//Internal Properties (Extensions to ODL)
public:
	LONG GetEventTarget(void);
	void SetEventTarget(LONG Value);
	LONG GetObjectIndex(void);
	void SetObjectIndex(LONG Value);
	LONG GetParentObjectIndex(void);
	void SetParentObjectIndex(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(ObjectSelectedEventArgs)
	DECLARE_INTERFACE_MAP()

LONG odl_GetEventTarget(void)
{
	
	return GetEventTarget();
}

void odl_SetEventTarget(LONG Value)
{
	
	SetEventTarget(Value);
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

};
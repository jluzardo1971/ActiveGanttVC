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


class ObjectStateChangedEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(ObjectStateChangedEventArgs)

public:
	ObjectStateChangedEventArgs(void);
	virtual ~ObjectStateChangedEventArgs();
	virtual void OnFinalRelease();

// Member Variables
public:
	LONG mp_yEventTarget;				
	LONG mp_lIndex;				
	BOOL mp_bCancel;				
	LONG mp_lDestinationIndex;				
	LONG mp_lInitialRowIndex;				
	LONG mp_lFinalRowIndex;				
	LONG mp_lInitialColumnIndex;				
	LONG mp_lFinalColumnIndex;
	COleDateTime mp_dtEndDate;
	COleDateTime mp_dtStartDate;
	LONG mp_lChangeType;

//Internal Properties (Extensions to ODL)
public:
	LONG GetEventTarget(void);
	void SetEventTarget(LONG Value);
	LONG GetIndex(void);
	void SetIndex(LONG Value);
	BOOL GetCancel(void);
	void SetCancel(BOOL Value);
	LONG GetDestinationIndex(void);
	void SetDestinationIndex(LONG Value);
	LONG GetInitialRowIndex(void);
	void SetInitialRowIndex(LONG Value);
	LONG GetFinalRowIndex(void);
	void SetFinalRowIndex(LONG Value);
	LONG GetInitialColumnIndex(void);
	void SetInitialColumnIndex(LONG Value);
	LONG GetFinalColumnIndex(void);
	void SetFinalColumnIndex(LONG Value);
	COleDateTime GetStartDate(void);
	void SetStartDate(COleDateTime Value);
	COleDateTime GetEndDate(void);
	void SetEndDate(COleDateTime Value);
	LONG GetChangeType(void);
	void SetChangeType(LONG Value);

//Internal Properties
public:

//Internal Methods
public:
	void Clear(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(ObjectStateChangedEventArgs)
	DECLARE_INTERFACE_MAP()

LONG odl_GetEventTarget(void)
{
	
	return GetEventTarget();
}

void odl_SetEventTarget(LONG Value)
{
	
	SetEventTarget(Value);
}

LONG odl_GetIndex(void)
{
	
	return GetIndex();
}

void odl_SetIndex(LONG Value)
{
	
	SetIndex(Value);
}

BOOL odl_GetCancel(void)
{
	
	return GetCancel();
}

void odl_SetCancel(BOOL Value)
{
	
	
	SetCancel(Value);
}

LONG odl_GetDestinationIndex(void)
{
	
	return GetDestinationIndex();
}

void odl_SetDestinationIndex(LONG Value)
{
	
	SetDestinationIndex(Value);
}

LONG odl_GetInitialRowIndex(void)
{
	
	return GetInitialRowIndex();
}

void odl_SetInitialRowIndex(LONG Value)
{
	
	SetInitialRowIndex(Value);
}

LONG odl_GetFinalRowIndex(void)
{
	
	return GetFinalRowIndex();
}

void odl_SetFinalRowIndex(LONG Value)
{
	
	SetFinalRowIndex(Value);
}

LONG odl_GetInitialColumnIndex(void)
{
	
	return GetInitialColumnIndex();
}

void odl_SetInitialColumnIndex(LONG Value)
{
	
	SetInitialColumnIndex(Value);
}

LONG odl_GetFinalColumnIndex(void)
{
	
	return GetFinalColumnIndex();
}

void odl_SetFinalColumnIndex(LONG Value)
{
	
	SetFinalColumnIndex(Value);
}

DATE odl_GetStartDate(void)
{
	
	return mp_dtStartDate;
}

DATE odl_GetEndDate(void)
{
	
	return mp_dtEndDate;
}

LONG odl_GetChangeType(void)
{
	return GetChangeType();
}

};
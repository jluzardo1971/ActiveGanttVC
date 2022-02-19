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

class clsCollectionBase;
class clsTimeBlock;
class clsTimeBlocks;

class clsTimeBlocks : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTimeBlocks)
	clsTimeBlocks();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTimeBlocks(CActiveGanttVCCtl* oControl);
	virtual ~clsTimeBlocks();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsCollectionBase* mp_oCollection;
	std::vector <S_TIMEBLOCK> mp_aTimeBlocks;
    COleDateTime mp_dtIntervalStart;
    COleDateTime mp_dtIntervalEnd;
    LONG mp_yIntervalType;

//Internal Properties (Extensions to ODL)
public:
	LONG GetCount(void);

//Internal Properties
public:
	clsCollectionBase* GetCollection(void);

//Internal Methods
public:
	clsTimeBlock* Item(CString Index);
	clsTimeBlock* Add(CString Key);
	void Clear(void);
	void Remove(CString Index);
	void CreateTemporaryTimeBlocks(void);
	void CopyTimeBlock(clsTimeBlock* oDestination,clsTimeBlock* oOriginal);
	void Draw(void);
	void DrawClass(clsTimeBlocks* oTimeBlocks);
	CString GetXML(void);
	void SetXML(CString sXML);
	COleDateTime GetIntervalStart(void);
	void SetIntervalStart(COleDateTime Value);
	COleDateTime GetIntervalEnd(void);
	void SetIntervalEnd(COleDateTime Value);
	LONG GetIntervalType(void);
	void SetIntervalType(LONG Value);
	void CalculateInterval(void);
	CString CP_GetXML(void);
	void CP_SetXML(CString sXML);


protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTimeBlocks)
	DECLARE_INTERFACE_MAP()

LONG odl_GetCount(void)
{
	
	return GetCount();
}

IDispatch* odl_Item(LPCTSTR Index)
{
	
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* odl_Add(LPCTSTR Key)
{
	
	return Add(Key)->GetIDispatch(TRUE);
}

void odl_Clear(void)
{
	
	Clear();
}

void odl_Remove(LPCTSTR Index)
{
	
	Remove(Index);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

BSTR odl_CP_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_CP_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

DATE odl_GetIntervalStart(void)
{
	
	return mp_dtIntervalStart;
}

void odl_SetIntervalStart(DATE Value)
{
	SetIntervalStart(Value);
}

DATE odl_GetIntervalEnd(void)
{
	
	return mp_dtIntervalEnd;
}

void odl_SetIntervalEnd(DATE Value)
{
	SetIntervalEnd(Value);
}

LONG odl_GetIntervalType(void)
{
	
	return GetIntervalType();
}

void odl_SetIntervalType(LONG Value)
{
	
	SetIntervalType(Value);
}

void odl_CalculateInterval(void)
{
	
	CalculateInterval();
}


};
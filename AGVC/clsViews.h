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
class clsView;

class clsViews : public CCmdTarget
{
	DECLARE_DYNCREATE(clsViews)
	clsViews();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsViews(CActiveGanttVCCtl* oControl);
	virtual ~clsViews();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsCollectionBase* mp_oCollection;
	clsView* mp_oDefaultView;




//Internal Properties (Extensions to ODL)
public:
	LONG GetCount(void);

//Internal Properties
public:
	clsCollectionBase* GetCollection(void);

//Internal Methods
public:
	void Finalize(void);
	clsView* Item(CString Index);
	clsView* FItem(CString Index);
	clsView* Add(LONG Interval,LONG Factor,LONG UpperTierType,LONG MiddleTierType,LONG LowerTierType,CString Key);
	void Clear(void);
	void Remove(CString Index);
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsViews)
	DECLARE_INTERFACE_MAP()

LONG odl_GetCount(void)
{
	
	return GetCount();
}

IDispatch* odl_Item(LPCTSTR Index)
{
	
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* odl_Add(LONG Interval,LONG Factor,LONG UpperTierType,LONG MiddleTierType,LONG LowerTierType,LPCTSTR Key)
{
	
	return Add(Interval,Factor,UpperTierType,MiddleTierType,LowerTierType,Key)->GetIDispatch(TRUE);
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


};
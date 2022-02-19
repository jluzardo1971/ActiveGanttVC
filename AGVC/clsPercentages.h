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
class clsPercentage;

class clsPercentages : public CCmdTarget
{
	DECLARE_DYNCREATE(clsPercentages)
	clsPercentages();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsPercentages(CActiveGanttVCCtl* oControl);
	virtual ~clsPercentages();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsCollectionBase* mp_oCollection;




//Internal Properties (Extensions to ODL)
public:
	LONG GetCount(void);

//Internal Properties
public:
	clsCollectionBase* GetCollection(void);

//Internal Methods
public:
	void Finalize(void);
	clsPercentage* Item(CString Index);
	clsPercentage* Add(CString TaskKey,CString StyleIndex,FLOAT Percent,CString Key);
	void Clear(void);
	void Remove(CString Index);
	void Draw(void);
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsPercentages)
	DECLARE_INTERFACE_MAP()

LONG odl_GetCount(void)
{
	
	return GetCount();
}

IDispatch* odl_Item(LPCTSTR Index)
{
	
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* odl_Add(LPCTSTR TaskKey,LPCTSTR StyleIndex,FLOAT Percent,LPCTSTR Key)
{
	
	return Add(TaskKey,StyleIndex,Percent,Key)->GetIDispatch(TRUE);
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
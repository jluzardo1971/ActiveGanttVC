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
class clsTickMark;

class clsTickMarks : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTickMarks)
	clsTickMarks();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTickMarks(CActiveGanttVCCtl* oControl);
	virtual ~clsTickMarks();
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
	clsTickMark* Item(CString Index);
	clsTickMark* Add(LONG Interval,LONG Factor,LONG TickMarkType,BOOL DisplayText,CString TextFormat,CString Key);
	void Clear(void);
	void Remove(CString Index);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clone(clsTickMarks* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTickMarks)
	DECLARE_INTERFACE_MAP()


IDispatch* odl_Item(LPCTSTR Index)
{
	
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* odl_Add(LONG Interval,LONG Factor,LONG TickMarkType,BOOL DisplayText,LPCTSTR TextFormat,LPCTSTR Key)
{
	
	return Add(Interval,Factor,TickMarkType,DisplayText,TextFormat,Key)->GetIDispatch(TRUE);
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

LONG odl_GetCount(void)
{
	
	return GetCount();
}

};
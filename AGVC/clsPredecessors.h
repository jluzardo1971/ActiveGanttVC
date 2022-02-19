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
class clsTask;
class clsPredecessor;
class clsPredecessorStyle;

class clsPredecessors : public CCmdTarget
{
	DECLARE_DYNCREATE(clsPredecessors)
	clsPredecessors();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsPredecessors(CActiveGanttVCCtl* oControl);
	
	virtual ~clsPredecessors();
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
	clsPredecessor* Item(CString Index);
	void Draw(void);
	void mp_DrawEndToEnd(clsPredecessor* oPredecessor);
	void mp_DrawStartToStart(clsPredecessor* oPredecessor);
	void mp_DrawStartToEnd(clsPredecessor* oPredecessor);
	void mp_DrawEndToStart(clsPredecessor* oPredecessor);
	void mp_DrawPredecessorLines(S_Point oPoints[], int Size, clsPredecessor* oPredecessor);
	clsPredecessor* Add(CString SuccessorKey, CString PredecessorKey, LONG PredecessorType, CString Key, CString StyleIndex);
	void Clear(void);
	void Remove(CString Index);
	CString GetXML(void);
	void SetXML(CString sXML);
	BOOL mp_DrawPredecessor(clsTask* oTask1, clsTask* oTask2);
	clsPredecessorStyle* mp_GetPredecessorStyle(clsPredecessor* oPredecessor);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsPredecessors)
	DECLARE_INTERFACE_MAP()


LONG odl_GetCount(void)
{
	
	return GetCount();
}

IDispatch* odl_Item(LPCTSTR Index)
{
	
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* odl_Add(LPCTSTR SuccessorKey, LPCTSTR PredecessorKey, LONG PredecessorType, LPCTSTR Key, LPCTSTR StyleIndex)
{
	
	return Add(SuccessorKey, PredecessorKey, PredecessorType, Key, StyleIndex)->GetIDispatch(TRUE);
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
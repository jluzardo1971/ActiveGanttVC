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
class clsColumn;

class clsColumns : public CCmdTarget
{
	DECLARE_DYNCREATE(clsColumns)
	clsColumns();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsColumns(CActiveGanttVCCtl* oControl);
	virtual ~clsColumns();
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
	LONG GetWidth(void);

//Internal Methods
public:
	void Finalize(void);
	clsColumn* Item(CString Index);
	clsColumn* Add(CString Text,CString Key,LONG Width,CString StyleIndex);
	void Clear(void);
	void Remove(CString Index);
	void MoveColumn(LONG OriginColumnIndex,LONG DestColumnIndex);
	void Position(void);
	void Draw(void);
	CString GetXML(void);
	void SetXML(CString sXML);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsColumns)
	DECLARE_INTERFACE_MAP()

LONG odl_GetCount(void)
{
	
	return GetCount();
}

IDispatch* odl_Item(LPCTSTR Index)
{
	
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* odl_Add(LPCTSTR Text,LPCTSTR Key,LONG Width,LPCTSTR StyleIndex)
{
	
	return Add(Text,Key,Width,StyleIndex)->GetIDispatch(TRUE);
}

void odl_Clear(void)
{
	
	Clear();
}

void odl_Remove(LPCTSTR Index)
{
	
	Remove(Index);
}

void odl_MoveColumn(LONG OriginColumnIndex,LONG DestColumnIndex)
{
	
	MoveColumn(OriginColumnIndex,DestColumnIndex);
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
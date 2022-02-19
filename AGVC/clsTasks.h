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

class clsTasks : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTasks)
	clsTasks();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTasks(CActiveGanttVCCtl* oControl);
	virtual ~clsTasks();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsCollectionBase* mp_oCollection;
	LONG mp_lLoadIndex;

//Internal Properties (Extensions to ODL)
public:
	LONG GetCount(void);

//Internal Properties
public:
	clsCollectionBase* GetCollection(void);

//Internal Methods
public:
	clsTask* Item(CString Index);
	clsTask* Add(CString Text,CString RowKey,COleDateTime StartDate,COleDateTime EndDate,CString Key,CString StyleIndex,CString LayerIndex);
	void Clear(void);
	void Remove(CString Index);
	void Sort(CString PropertyName,BOOL Descending,LONG SortType,LONG StartIndex,LONG EndIndex);
	void Draw(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void BeginLoad(BOOL Preserve);
	clsTask* Load(CString sRowKey,CString sKey);
	void EndLoad(void);
	clsStyle* mp_GetTaskStyle(clsTask* oTask);
	clsTask* DAdd(CString RowKey, COleDateTime StartDate, LONG DurationInterval, LONG DurationFactor, CString Text, CString Key, CString StyleIndex, CString LayerIndex);


protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTasks)
	DECLARE_INTERFACE_MAP()

void odl_BeginLoad(BOOL Preserve)
{
	BeginLoad(Preserve);
}

IDispatch* odl_Load(LPCTSTR sRowKey,LPCTSTR sKey)
{
	return Load(sRowKey, sKey)->GetIDispatch(TRUE);
}

void odl_EndLoad(void)
{
	EndLoad();
}

LONG odl_GetCount(void)
{
	
	return GetCount();
}

IDispatch* odl_Item(LPCTSTR Index)
{
	
	return Item(Index)->GetIDispatch(TRUE);
}

IDispatch* odl_Add(LPCTSTR Text,LPCTSTR RowKey,DATE StartDate,DATE EndDate,LPCTSTR Key,LPCTSTR StyleIndex,LPCTSTR LayerIndex)
{
	return Add(Text,RowKey, StartDate, EndDate,Key,StyleIndex,LayerIndex)->GetIDispatch(TRUE);
}

IDispatch* odl_DAdd(LPCTSTR RowKey, DATE StartDate, LONG DurationInterval, LONG DurationFactor, LPCTSTR Text, LPCTSTR Key, LPCTSTR StyleIndex, LPCTSTR LayerIndex)
{
	return DAdd(RowKey, StartDate,DurationInterval,DurationFactor,Text,Key,StyleIndex,LayerIndex)->GetIDispatch(TRUE);
}

void odl_Clear(void)
{
	
	Clear();
}

void odl_Remove(LPCTSTR Index)
{
	
	Remove(Index);
}

void odl_Sort(LPCTSTR PropertyName,BOOL Descending,LONG SortType,LONG StartIndex,LONG EndIndex)
{
	
	Sort(PropertyName,Descending,SortType,StartIndex,EndIndex);
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
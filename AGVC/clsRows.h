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
class clsRow;
class clsNode;
class clsDictionary;

class clsRows : public CCmdTarget
{
	DECLARE_DYNCREATE(clsRows)
	clsRows();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsRows(CActiveGanttVCCtl* oControl);
	
	virtual ~clsRows();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsCollectionBase* mp_oCollection;
	LONG mp_lTopOffset;
	LONG mp_lRealFirstVisibleRow;
    LONG mp_lLoadIndex;
    CObArray mp_oTempNodeList;

//Internal Properties (Extensions to ODL)
public:
	LONG GetCount(void);

//Internal Properties
public:
	clsCollectionBase* GetCollection(void);
	LONG GetTopOffset(void);
	void SetTopOffset(LONG Value);
	LONG GetRealFirstVisibleRow(void);


//Internal Methods
public:
	clsRow* Item(CString Index);
	clsRow* Add(CString Key,CString Text,BOOL MergeCells,BOOL Container,CString StyleIndex);
	void Clear(void);
	void Remove(CString Index);
	void MoveRow(LONG OriginRowIndex,LONG DestRowIndex);
	void SortRows(CString PropertyName,BOOL Descending,LONG SortType,LONG StartIndex,LONG EndIndex);
	void SortCells(LONG CellIndex,CString PropertyName,BOOL Descending,LONG SortType,LONG StartIndex,LONG EndIndex);
	void ClearCells(void);
	LONG Height(void);
	LONG Height(LONG lRow);
	LONG CalculateHeight(LONG StartIndex,LONG EndIndex);
	LONG CalculateRows(LONG StartIndex,LONG Height);
	void PositionRows(void);
	LONG NodeHeight(LONG lFirstVisibleNode);
	void InitializePosition(void);
	LONG RealIndex(LONG Index);
	void NodesDrawBackground(void);
	void NodesDrawTreeLines(void);
	void NodesDraw(void);
	void NodesDrawElements();
	void mp_DrawSign(BOOL bExpanded,LONG X,LONG Y, clsRow* oRow);
	LONG HiddenRows(void);
	void mp_DrawCheckBox(clsNode* oNode);
	void mp_DrawImage(clsNode* oNode);
	void Draw(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void UpdateTree();
	BOOL VerifyTree();
	void BeginLoad(BOOL Preserve);
	clsRow* Load(CString sKey);
	void EndLoad(void);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsRows)
	DECLARE_INTERFACE_MAP()

void odl_BeginLoad(BOOL Preserve)
{
	BeginLoad(Preserve);
}

IDispatch* odl_Load(LPCTSTR sKey)
{
	
	return Load(sKey)->GetIDispatch(TRUE);
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

IDispatch* odl_Add(LPCTSTR Key,LPCTSTR Text,BOOL MergeCells,BOOL Container,BSTR StyleIndex)
{
	
	return Add(Key,Text,MergeCells,Container,StyleIndex)->GetIDispatch(TRUE);
}

void odl_Clear(void)
{
	
	Clear();
}

void odl_Remove(LPCTSTR Index)
{
	
	Remove(Index);
}

void odl_MoveRow(LONG OriginRowIndex,LONG DestRowIndex)
{
	
	MoveRow(OriginRowIndex,DestRowIndex);
}

void odl_SortRows(LPCTSTR PropertyName,BOOL Descending,LONG SortType,LONG StartIndex,LONG EndIndex)
{
	
	SortRows(PropertyName,Descending,SortType,StartIndex,EndIndex);
}

void odl_SortCells(LONG CellIndex,LPCTSTR PropertyName,BOOL Descending,LONG SortType,LONG StartIndex,LONG EndIndex)
{
	
	SortCells(CellIndex,PropertyName,Descending,SortType,StartIndex,EndIndex);
}

void odl_UpdateTree(void)
{
	
	UpdateTree();
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	SetXML(sXML);
}

BOOL odl_VerifyTree(void)
{
	return VerifyTree();
}


};
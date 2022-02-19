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
class clsStyle;

class clsStyles : public CCmdTarget
{
	DECLARE_DYNCREATE(clsStyles)
	clsStyles();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsStyles(CActiveGanttVCCtl* oControl);
	virtual ~clsStyles();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsCollectionBase* mp_oCollection;
	clsStyle* mp_oDefaultControlStyle;
	clsStyle* mp_oDefaultTaskStyle;
	clsStyle* mp_oDefaultRowStyle;
	clsStyle* mp_oDefaultClientAreaStyle;
	clsStyle* mp_oDefaultCellStyle;
	clsStyle* mp_oDefaultColumnStyle;
	clsStyle* mp_oDefaultPercentageStyle;
	clsStyle* mp_oDefaultPredecessorStyle;
	clsStyle* mp_oDefaultTimeLineStyle;
	clsStyle* mp_oDefaultTimeBlockStyle;
	clsStyle* mp_oDefaultTickMarkAreaStyle;
	clsStyle* mp_oDefaultSplitterStyle;
	clsStyle* mp_oDefaultProgressLineStyle;
	clsStyle* mp_oDefaultNodeStyle;
	clsStyle* mp_oDefaultTierStyle;
	clsStyle* mp_oDefaultScrollBarStyle;
	clsStyle* mp_oDefaultSBSeparatorStyle;
	clsStyle* mp_oDefaultSBNormalStyle;
	clsStyle* mp_oDefaultSBPressedStyle;
	clsStyle* mp_oDefaultSBDisabledStyle;




//Internal Properties (Extensions to ODL)
public:
	LONG GetCount(void);

//Internal Properties
public:
	clsCollectionBase* GetCollection(void);

//Internal Methods
public:
	void Finalize(void);
	clsStyle* Item(CString Index);
	clsStyle* FItem(CString Index);
	clsStyle* Add(CString Key);
	void Clear(void);
	void Remove(CString Index);
	CString GetXML(void);
	void SetXML(CString sXML);
	CString mp_GetNewStyleIndex(CString sKey, CString sIndex, CString sStyleIndex);
	void mp_ResetDefaultStyles(void);
	BOOL ContainsKey(CString Key);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsStyles)
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

BOOL odl_ContainsKey(LPCTSTR Key)
{
	
	return ContainsKey(Key);
}


};
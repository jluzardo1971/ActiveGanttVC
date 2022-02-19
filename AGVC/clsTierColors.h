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
class clsTierColor;

class clsTierColors : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTierColors)
	clsTierColors();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTierColors(CActiveGanttVCCtl* oControl, LONG yTierType);
	virtual ~clsTierColors();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsCollectionBase* mp_oCollection;
	LONG mp_yTierType;




//Internal Properties (Extensions to ODL)
public:
	LONG GetCount(void);

//Internal Properties
public:
	clsCollectionBase* GetCollection(void);

//Internal Methods
public:
	void Finalize(void);
	clsTierColor* Item(CString Index);
	clsTierColor* Add(Gdiplus::Color BackColor, Gdiplus::Color ForeColor, Gdiplus::Color StartGradientColor, Gdiplus::Color EndGradientColor, Gdiplus::Color HatchBackColor, Gdiplus::Color HatchForeColor, CString Key);
	clsTierColor* Add(Gdiplus::Color BackColor, Gdiplus::Color ForeColor, Gdiplus::Color StartGradientColor, Gdiplus::Color EndGradientColor, Gdiplus::Color HatchBackColor, Gdiplus::Color HatchForeColor);
	CString mp_CollectionName(void);
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTierColors* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTierColors)
	DECLARE_INTERFACE_MAP()

IDispatch* odl_Item(LPCTSTR Index)
{
	
	return Item(Index)->GetIDispatch(TRUE);
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
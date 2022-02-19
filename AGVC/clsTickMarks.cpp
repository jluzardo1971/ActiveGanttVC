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
#include "stdafx.h"
#include "ActiveGanttVC.h"
#include "ActiveGanttVCCtl.h"
#include "clsXML.h"
#include "clsTickMarks.h"

IMPLEMENT_DYNCREATE(clsTickMarks, CCmdTarget)

//{35873283-4408-4E22-A471-7FF0D5D89B7F}
static const IID IID_IclsTickMarks = { 0x35873283, 0x4408, 0x4E22, { 0xA4, 0x71, 0x7F, 0xF0, 0xD5, 0xD8, 0x9B, 0x7F} };

//{0A97ECFD-F3DF-4EE5-A524-1F78627DD60C}
IMPLEMENT_OLECREATE_FLAGS(clsTickMarks, "AGVC.clsTickMarks", afxRegApartmentThreading, 0x0A97ECFD, 0xF3DF, 0x4EE5, 0xA5, 0x24, 0x1F, 0x78, 0x62, 0x7D, 0xD6, 0x0C)

BEGIN_DISPATCH_MAP(clsTickMarks, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTickMarks, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsTickMarks, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsTickMarks, "Clear", 3, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsTickMarks, "Remove", 4, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsTickMarks, "GetXML", 5, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTickMarks, "SetXML", 6, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(clsTickMarks, "Add", 7, odl_Add, VT_DISPATCH, VTS_I4 VTS_I4 VTS_I4 VTS_BOOL VTS_BSTR VTS_BSTR)
	DISP_DEFVALUE(clsTickMarks, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTickMarks, CCmdTarget)
	INTERFACE_PART(clsTickMarks, IID_IclsTickMarks, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTickMarks, CCmdTarget)
END_MESSAGE_MAP()

clsTickMarks::clsTickMarks(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "TickMark");
}

clsTickMarks::clsTickMarks()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTickMarks. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTickMarks::~clsTickMarks()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	AfxOleUnlockApp();
}

void clsTickMarks::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsTickMarks::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsTickMarks::GetCollection(void)
{
	return mp_oCollection;
}

void clsTickMarks::Finalize(void)
{
}

clsTickMark* clsTickMarks::Item(CString Index)
{
	clsTickMark *oTickMark;
	oTickMark = (clsTickMark*)mp_oCollection->m_oItem(Index, TICKMARKS_ITEM_1, TICKMARKS_ITEM_2, TICKMARKS_ITEM_3, TICKMARKS_ITEM_4);
	return oTickMark;
}

clsTickMark* clsTickMarks::Add(LONG Interval,LONG Factor,LONG TickMarkType,BOOL DisplayText,CString TextFormat,CString Key)
{
	mp_oCollection->SetAddMode(TRUE);
	clsTickMark* oTickMark = new clsTickMark(mp_oControl, this);
	oTickMark->SetInterval(Interval);
	oTickMark->SetFactor(Factor);
	oTickMark->SetTickMarkType(TickMarkType);
	oTickMark->SetDisplayText(DisplayText);
	oTickMark->SetTextFormat(TextFormat);
	oTickMark->SetKey(Key);
	mp_oCollection->m_Add(oTickMark, Key, TICKMARKS_ADD_1, TICKMARKS_ADD_2, FALSE, TICKMARKS_ADD_3);
	return oTickMark;
}

void clsTickMarks::Clear(void)
{
	mp_oCollection->m_Clear();
}

void clsTickMarks::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, TICKMARKS_REMOVE_1, TICKMARKS_REMOVE_2, TICKMARKS_REMOVE_3, TICKMARKS_REMOVE_4);
}

CString clsTickMarks::GetXML(void)
{
	LONG lIndex;
	clsTickMark* oTickMark; 
	clsXML oXML(mp_oControl, "TickMarks");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTickMark = (clsTickMark*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTickMark->GetXML());
	}
	return oXML.GetXML();
}

void clsTickMarks::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "TickMarks");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsTickMark* oTickMark = new clsTickMark(mp_oControl, this);
		oTickMark->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oTickMark, oTickMark->mp_sKey, TICKMARKS_ADD_1, TICKMARKS_ADD_2, FALSE, TICKMARKS_ADD_3);
	}
}

void clsTickMarks::Clone(clsTickMarks* oClone)
{
    oClone->Clear();
    LONG i;
    for (i = 1; i <= GetCount(); i++)
	{
        clsTickMark* oTickMark = Item(CStr(i));
        clsTickMark* oTickMarkClone;
        oTickMarkClone = oClone->Add(oTickMark->GetInterval(), oTickMark->GetFactor(), oTickMark->GetTickMarkType(), oTickMark->GetDisplayText(), oTickMark->GetTextFormat(), oTickMark->GetKey());
        oTickMark->Clone(oTickMarkClone);
	}
}
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
#include "clsPercentages.h"



IMPLEMENT_DYNCREATE(clsPercentages, CCmdTarget)

//{95620627-D756-4697-8217-D8190DBACA1C}
static const IID IID_IclsPercentages = { 0x95620627, 0xD756, 0x4697, { 0x82, 0x17, 0xD8, 0x19, 0x0D, 0xBA, 0xCA, 0x1C} };

//{A4A469B6-A527-426F-A955-7B43DD5F0C01}
IMPLEMENT_OLECREATE_FLAGS(clsPercentages, "AGVC.clsPercentages", afxRegApartmentThreading, 0xA4A469B6, 0xA527, 0x426F, 0xA9, 0x55, 0x7B, 0x43, 0xDD, 0x5F, 0x0C, 0x01)

BEGIN_DISPATCH_MAP(clsPercentages, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsPercentages, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsPercentages, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsPercentages, "Add", 3, odl_Add, VT_DISPATCH, VTS_BSTR VTS_BSTR VTS_R4 VTS_BSTR) 
	DISP_FUNCTION_ID(clsPercentages, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsPercentages, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsPercentages, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsPercentages, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_DEFVALUE(clsPercentages, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsPercentages, CCmdTarget)
	INTERFACE_PART(clsPercentages, IID_IclsPercentages, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsPercentages, CCmdTarget)
END_MESSAGE_MAP()

clsPercentages::clsPercentages(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "Percentage");
}

clsPercentages::clsPercentages()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsPercentages. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsPercentages::~clsPercentages()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	AfxOleUnlockApp();
}

void clsPercentages::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsPercentages::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsPercentages::GetCollection(void)
{
	return mp_oCollection;
}

void clsPercentages::Finalize(void)
{
}

clsPercentage* clsPercentages::Item(CString Index)
{
	clsPercentage *oPercentage;
	oPercentage = (clsPercentage*)mp_oCollection->m_oItem(Index, PERCENTAGES_ITEM_1, PERCENTAGES_ITEM_2, PERCENTAGES_ITEM_3, PERCENTAGES_ITEM_4);
	return oPercentage;
}

clsPercentage* clsPercentages::Add(CString TaskKey,CString StyleIndex,FLOAT Percent,CString Key)
{
	mp_oCollection->SetAddMode(TRUE);
	clsPercentage* oPercentage = new clsPercentage(mp_oControl);
	Key = g_Trim(Key);
	TaskKey = g_Trim(TaskKey);
	oPercentage->SetKey(Key);
	oPercentage->SetTaskKey(TaskKey);
	oPercentage->SetPercent(Percent);
	oPercentage->SetStyleIndex(StyleIndex);
	mp_oCollection->m_Add(oPercentage, Key, PERCENTAGES_ADD_1, PERCENTAGES_ADD_2, FALSE, PERCENTAGES_ADD_3);
	return oPercentage;
}

void clsPercentages::Clear(void)
{
	mp_oCollection->m_Clear();
}

void clsPercentages::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, PERCENTAGES_REMOVE_1, PERCENTAGES_REMOVE_2, PERCENTAGES_REMOVE_3, PERCENTAGES_REMOVE_4);
}

void clsPercentages::Draw(void)
{
	LONG lIndex;
	clsPercentage* oPercentage;
	clsTask* oTask;
	clsRow* oRow;
	clsStyle* oStyle;
	BOOL bDraw;
	if (GetCount() == 0)
	{
		return;
	}
	if (GetCount() == 0)
	{
		return;
	}
	for (lIndex = 1;lIndex <= GetCount();lIndex++) 
	{
		oPercentage = (clsPercentage*) mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oPercentage->GetVisible() == TRUE)
		{
			mp_oControl->clsG->mp_ClipRegion(oPercentage->GetLeftTrim(), oPercentage->GetTop(), oPercentage->GetRightTrim(), oPercentage->GetBottom(), TRUE);
			oTask = (clsTask*) mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElementKey(oPercentage->GetTaskKey());
			oRow = (clsRow*) mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElementKey(oTask->GetRowKey());
			oStyle = mp_oControl->GetStyles()->FItem(oPercentage->GetStyleIndex());
			bDraw = FALSE;
			//*mp_oControl->FireDraw(EVT_PERCENTAGE, bDraw, lIndex, 0, mp_oControl->clsG->oGraphics);
			if (bDraw == FALSE)
			{
				mp_oControl->clsG->mp_DrawItem(oPercentage->GetLeft(), oPercentage->GetTop(), oPercentage->GetRight(), oPercentage->GetBottom(), oPercentage->GetStyleIndex(), oPercentage->ToString(), FALSE, NULL, oPercentage->GetLeftTrim(), oPercentage->GetRightTrim(), NULL);
			}
		}
	}
}

CString clsPercentages::GetXML(void)
{
	LONG lIndex;
	clsPercentage* oPercentage; 
	clsXML oXML(mp_oControl, "Percentages");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oPercentage = (clsPercentage*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oPercentage->GetXML());
	}
	return oXML.GetXML();
}

void clsPercentages::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "Percentages");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsPercentage* oPercentage = new clsPercentage(mp_oControl);
		oPercentage->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oPercentage, oPercentage->mp_sKey, PERCENTAGES_ADD_1, PERCENTAGES_ADD_2, FALSE, PERCENTAGES_ADD_3);
	}
}
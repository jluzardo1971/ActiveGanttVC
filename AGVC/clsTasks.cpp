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
#include "clsTasks.h"



IMPLEMENT_DYNCREATE(clsTasks, CCmdTarget)

//{BDE36FA1-AF28-4F71-8785-8DF7C541BF7F}
static const IID IID_IclsTasks = { 0xBDE36FA1, 0xAF28, 0x4F71, { 0x87, 0x85, 0x8D, 0xF7, 0xC5, 0x41, 0xBF, 0x7F} };

//{7ED57F62-F65C-4B5B-AF57-EF0488B0ABD9}
IMPLEMENT_OLECREATE_FLAGS(clsTasks, "AGVC.clsTasks", afxRegApartmentThreading, 0x7ED57F62, 0xF65C, 0x4B5B, 0xAF, 0x57, 0xEF, 0x04, 0x88, 0xB0, 0xAB, 0xD9)

BEGIN_DISPATCH_MAP(clsTasks, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTasks, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsTasks, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsTasks, "Add", 3, odl_Add, VT_DISPATCH, VTS_BSTR VTS_BSTR VTS_DATE VTS_DATE VTS_BSTR VTS_BSTR VTS_BSTR) 
	DISP_FUNCTION_ID(clsTasks, "Clear", 4, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsTasks, "Remove", 5, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsTasks, "Sort", 6, odl_Sort, VT_EMPTY, VTS_BSTR VTS_BOOL VTS_I4 VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsTasks, "GetXML", 7, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTasks, "SetXML", 8, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(clsTasks, "BeginLoad", 9, odl_BeginLoad, VT_EMPTY, VTS_BOOL) 
	DISP_FUNCTION_ID(clsTasks, "Load", 10, odl_Load, VT_DISPATCH, VTS_BSTR VTS_BSTR) 
	DISP_FUNCTION_ID(clsTasks, "EndLoad", 11, odl_EndLoad, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(clsTasks, "DAdd", 12, odl_DAdd, VT_DISPATCH, VTS_BSTR VTS_DATE VTS_I4 VTS_I4 VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR) 
	DISP_DEFVALUE(clsTasks, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTasks, CCmdTarget)
	INTERFACE_PART(clsTasks, IID_IclsTasks, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTasks, CCmdTarget)
END_MESSAGE_MAP()

clsTasks::clsTasks(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "Task");
	mp_lLoadIndex = -1;
}

clsTasks::clsTasks()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTasks. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTasks::~clsTasks()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	AfxOleUnlockApp();
}

void clsTasks::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsTasks::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsTasks::GetCollection(void)
{
	return mp_oCollection;
}

clsTask* clsTasks::Item(CString Index)
{
	clsTask *oTask;
	oTask = (clsTask*)mp_oCollection->m_oItem(Index, TASKS_ITEM_1, TASKS_ITEM_2, TASKS_ITEM_3, TASKS_ITEM_4);
	return oTask;
}

clsTask* clsTasks::Add(CString Text,CString RowKey,COleDateTime StartDate, COleDateTime EndDate,CString Key,CString StyleIndex,CString LayerIndex)
{
	mp_oCollection->SetAddMode(TRUE);
	clsTask* oTask = new clsTask(mp_oControl);
	Key = g_Trim(Key);
	Text = g_Trim(Text);
	RowKey = g_Trim(RowKey);
	oTask->SetText(Text);
	oTask->SetRowKey(RowKey);
	oTask->SetStartDate(StartDate);
	oTask->SetEndDate(EndDate);
	oTask->SetKey(Key);
	oTask->SetStyleIndex(StyleIndex);
	oTask->SetLayerIndex(LayerIndex);
	mp_oCollection->m_Add(oTask, Key, TASKS_ADD_1, TASKS_ADD_2, FALSE, TASKS_ADD_3);
	return oTask;
}

clsTask* clsTasks::DAdd(CString RowKey, COleDateTime StartDate, LONG DurationInterval, LONG DurationFactor,CString Text, CString Key, CString StyleIndex, CString LayerIndex)
{
	mp_oCollection->SetAddMode(TRUE);
	clsTask* oTask = new clsTask(mp_oControl);
	Key = g_Trim(Key);
	Text = g_Trim(Text);
	RowKey = g_Trim(RowKey);
	oTask->SetText(Text);
	oTask->SetRowKey(RowKey);
	oTask->SetStartDate(StartDate);
    oTask->SetDurationInterval(DurationInterval);
    oTask->SetDurationFactor(DurationFactor);
	oTask->SetTaskType(TT_DURATION);
	oTask->SetKey(Key);
	oTask->SetStyleIndex(StyleIndex);
	oTask->SetLayerIndex(LayerIndex);
	mp_oCollection->m_Add(oTask, Key, TASKS_ADD_1, TASKS_ADD_2, FALSE, TASKS_ADD_3);
	return oTask;
}

void clsTasks::Clear(void)
{
	mp_oControl->GetPredecessors()->Clear();
	mp_oControl->GetPercentages()->Clear();
	mp_oCollection->m_Clear();
	mp_oControl->SetSelectedTaskIndex(0);
}

void clsTasks::Remove(CString Index)
{
	LONG lIndex = 0;
	CString sRIndex = "";
	CString sRKey = "";
	clsPredecessor* oPredecessor = NULL;
	mp_oCollection->m_GetKeyAndIndex(Index, &sRKey, &sRIndex);
	mp_oControl->GetPercentages()->mp_oCollection->m_CollectionRemoveWhere("TaskKey", sRKey, sRIndex);
	if (sRKey.GetLength() > 0)
	{
		for (lIndex = mp_oControl->GetPredecessors()->GetCount(); lIndex >= 1; lIndex--)
		{
			oPredecessor = (clsPredecessor*)mp_oControl->GetPredecessors()->mp_oCollection->m_oReturnArrayElement(lIndex);
			if (oPredecessor->GetSuccessorKey() == sRKey || oPredecessor->GetPredecessorKey() == sRKey)
			{
				mp_oControl->GetPredecessors()->Remove(CStr(lIndex));
			}
		}
	}
	mp_oCollection->m_Remove(Index, TASKS_REMOVE_1, TASKS_REMOVE_2, TASKS_REMOVE_3, TASKS_REMOVE_4);
	mp_oControl->SetSelectedTaskIndex(0);
}

void clsTasks::Sort(CString PropertyName,BOOL Descending,LONG SortType,LONG StartIndex,LONG EndIndex)
{
	if (StartIndex == -1) 
	{
		StartIndex = 1;
	}
	if (EndIndex == -1) 
	{
		EndIndex = GetCount();
	}
	if (GetCount() == 0) return; 
	if (StartIndex < 1 || StartIndex > GetCount()) 
	{
		return;
	}
	if (EndIndex < 1 || EndIndex > GetCount()) 
	{
		return;
	}
	if (EndIndex == StartIndex) 
	{
		return;
	}
	mp_oCollection->m_Sort(PropertyName, Descending, (E_SORTTYPE)SortType, StartIndex, EndIndex);
}

void clsTasks::Draw(void)
{
	LONG lIndex = 0;
	clsTask* oTask = NULL;
	if (GetCount() == 0) 
	{
		return;
	}
	for (lIndex = 1; lIndex <= GetCount(); lIndex++) 
	{
		oTask = (clsTask*) mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oTask->GetVisible() == TRUE && oTask->GetInsideVisibleTimeLineArea() == TRUE) 
		{
			mp_oControl->GetDrawEventArgs()->Clear();
			mp_oControl->GetDrawEventArgs()->SetCustomDraw(FALSE);
			mp_oControl->GetDrawEventArgs()->SetEventTarget((LONG)EVT_TASK);
			mp_oControl->GetDrawEventArgs()->SetObjectIndex(lIndex);
			mp_oControl->GetDrawEventArgs()->SetParentObjectIndex(0);
			mp_oControl->FireDraw();
			if (mp_oControl->GetDrawEventArgs()->GetCustomDraw() == FALSE) 
			{
				if (oTask->GetType() == OT_MILESTONE) 
				{
					mp_oControl->clsG->mp_DrawItemI(oTask, "", lIndex == mp_oControl->GetSelectedTaskIndex(), mp_GetTaskStyle(oTask));
				}
				else 
				{
					mp_oControl->clsG->mp_DrawItem(oTask->GetLeft(), oTask->GetTop(), oTask->GetRight(), oTask->GetBottom(), "", oTask->GetText(), (lIndex == mp_oControl->GetSelectedTaskIndex()), oTask->GetImage(), oTask->GetLeftTrim(), oTask->GetRightTrim(), mp_GetTaskStyle(oTask));
				}
				if (oTask->GetText().GetLength() > 0)
				{
					oTask->mp_lTextLeft = mp_oControl->clsG->mp_oTextFinalLayout.Left;
					oTask->mp_lTextTop = mp_oControl->clsG->mp_oTextFinalLayout.Top;
					oTask->mp_lTextRight = mp_oControl->clsG->mp_oTextFinalLayout.Left + mp_oControl->clsG->mp_oTextFinalLayout.Width - 1;
					oTask->mp_lTextBottom = mp_oControl->clsG->mp_oTextFinalLayout.Top + mp_oControl->clsG->mp_oTextFinalLayout.Height - 1;
				}
				else
				{
					if (oTask->GetStyle()->GetTextPlacement() == SCP_EXTERIORPLACEMENT)
					{
						if (oTask->GetStyle()->GetTextAlignmentHorizontal() == HAL_LEFT)
						{
							oTask->mp_lTextLeft = oTask->GetLeft() - mp_oControl->clsG->mp_lStrWidth("WWWWW", oTask->GetStyle()->GetFont()) - oTask->GetStyle()->GetTextXMargin();
							oTask->mp_lTextRight = oTask->GetLeft() - oTask->GetStyle()->GetTextXMargin() + 1;
						}
						if (oTask->GetStyle()->GetTextAlignmentHorizontal() == HAL_RIGHT)
						{
							oTask->mp_lTextLeft = oTask->GetRight() + oTask->GetStyle()->GetTextXMargin();
							oTask->mp_lTextRight = oTask->GetRight() + mp_oControl->clsG->mp_lStrWidth("WWWWW", oTask->GetStyle()->GetFont()) + oTask->GetStyle()->GetTextXMargin() + 1;
						}
						if (oTask->GetStyle()->GetTextAlignmentHorizontal() == HAL_CENTER)
						{
							oTask->mp_lTextLeft = oTask->GetLeft();
							oTask->mp_lTextRight = oTask->GetRight() + 1;
						}
						if (oTask->GetStyle()->GetTextAlignmentVertical() == VAL_TOP)
						{
							oTask->mp_lTextTop = oTask->GetTop() - mp_oControl->clsG->mp_lStrHeight("WWWWW", oTask->GetStyle()->GetFont()) - oTask->GetStyle()->GetTextYMargin();
							oTask->mp_lTextBottom = oTask->GetTop() - oTask->GetStyle()->GetTextYMargin() + 3;
						}
						if (oTask->GetStyle()->GetTextAlignmentVertical() == VAL_BOTTOM)
						{
							oTask->mp_lTextTop = oTask->GetBottom() + oTask->GetStyle()->GetTextYMargin();
							oTask->mp_lTextBottom = oTask->GetBottom() + mp_oControl->clsG->mp_lStrHeight("WWWWW", oTask->GetStyle()->GetFont()) + oTask->GetStyle()->GetTextYMargin() + 3;
						}
						if (oTask->GetStyle()->GetTextAlignmentVertical() == VAL_CENTER)
						{
							oTask->mp_lTextTop = oTask->GetTop();
							oTask->mp_lTextBottom = oTask->GetBottom() + 3;
						}
					}
					else
					{
						oTask->mp_lTextLeft = oTask->GetLeft();
						oTask->mp_lTextTop = oTask->GetTop();
						oTask->mp_lTextRight = oTask->GetRight();
						oTask->mp_lTextBottom = oTask->GetBottom();
					}

				}
			}
		}
	}
}

clsStyle* clsTasks::mp_GetTaskStyle(clsTask* oTask)
{
    clsStyle* oStyle;
    if (oTask->mp_bWarning == TRUE)
    {
        oStyle = oTask->GetWarningStyle();
    }
    else
    {
        oStyle = oTask->GetStyle();
    }
    return oStyle;
}

CString clsTasks::GetXML(void)
{
	LONG lIndex;
	clsTask* oTask; 
	clsXML oXML(mp_oControl, "Tasks");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTask = (clsTask*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTask->GetXML());
	}
	return oXML.GetXML();
}

void clsTasks::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "Tasks");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsTask* oTask = new clsTask(mp_oControl);
		oTask->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oTask, oTask->mp_sKey, TASKS_ADD_1, TASKS_ADD_2, FALSE, TASKS_ADD_3);
	}
}

void clsTasks::BeginLoad(BOOL Preserve)
{
    if (Preserve == FALSE)
    {
        mp_lLoadIndex = 1;
		mp_oCollection->mp_aoCollection_Clear();
		mp_oCollection->mp_oKeys_Clear();
    }
    else
    {
        mp_lLoadIndex = mp_oCollection->mp_aoCollection.GetCount() + 1;
    }
}

clsTask* clsTasks::Load(CString sRowKey,CString sKey)
{
	if (mp_lLoadIndex == -1)
	{
		return NULL;
	}
	clsTask* oTask = new clsTask(mp_oControl);
	if (sKey.GetLength() > 0)
	{
		oTask->mp_sKey = sKey;
	}
	oTask->mp_sRowKey = sRowKey;
	oTask->mp_oRow = mp_oControl->GetRows()->Item(sRowKey);
	oTask->mp_lIndex = mp_lLoadIndex;
	mp_oCollection->mp_aoCollection.SetAt(mp_lLoadIndex, oTask);
	if (sKey.GetLength() > 0)
	{
		int* pIndex;
		pIndex = new int(mp_lLoadIndex);
		mp_oCollection->mp_oKeys.SetAt(sKey, pIndex);
	}
	mp_lLoadIndex = mp_lLoadIndex + 1;
	return oTask;
}

void clsTasks::EndLoad(void)
{
	mp_lLoadIndex = -1;
}
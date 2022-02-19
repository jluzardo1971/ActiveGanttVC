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
#include "clsPredecessor.h"



IMPLEMENT_DYNCREATE(clsPredecessor, clsItemBase)

//{71546057-2C77-4259-939D-752EEE5312DA}
static const IID IID_IclsPredecessor = { 0x71546057, 0x2C77, 0x4259, { 0x93, 0x9D, 0x75, 0x2E, 0xEE, 0x53, 0x12, 0xDA} };

//{43E2E679-CFBE-4A99-8C74-20B3FCF540C4}
IMPLEMENT_OLECREATE_FLAGS(clsPredecessor, "AGVC.clsPredecessor", afxRegApartmentThreading, 0x43E2E679, 0xCFBE, 0x4A99, 0x8C, 0x74, 0x20, 0xB3, 0xFC, 0xF5, 0x40, 0xC4)

BEGIN_DISPATCH_MAP(clsPredecessor, clsItemBase)
	DISP_PROPERTY_EX_ID(clsPredecessor, "Key", 1, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "Visible", 2, odl_GetVisible, odl_SetVisible, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "PredecessorKey", 3, odl_GetPredecessorKey, odl_SetPredecessorKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "PredecessorTask", 4, odl_GetPredecessorTask, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "PredecessorType", 5, odl_GetPredecessorType, odl_SetPredecessorType, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "StyleIndex", 6, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "Style", 7, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "Tag", 8, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_FUNCTION_ID(clsPredecessor, "GetXML", 9, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsPredecessor, "SetXML", 10, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsPredecessor, "Index", 11, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsPredecessor, "SuccessorKey", 12, odl_GetSuccessorKey, odl_SetSuccessorKey, VT_BSTR)
	DISP_PROPERTY_EX_ID(clsPredecessor, "SuccessorTask", 13, odl_GetSuccessorTask, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX_ID(clsPredecessor, "LagInterval", 14, odl_GetLagInterval, odl_SetLagInterval, VT_I4) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "LagFactor", 15, odl_GetLagFactor, odl_SetLagFactor, VT_I4)
	DISP_PROPERTY_EX_ID(clsPredecessor, "Warning", 16, odl_GetWarning, SetNotSupported, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsPredecessor, "WarningStyleIndex", 17, odl_GetWarningStyleIndex, odl_SetWarningStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "WarningStyle", 18, odl_GetWarningStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "SelectedStyleIndex", 19, odl_GetSelectedStyleIndex, odl_SetSelectedStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsPredecessor, "SelectedStyle", 20, odl_GetSelectedStyle, SetNotSupported, VT_DISPATCH) 
END_DISPATCH_MAP()




BEGIN_INTERFACE_MAP(clsPredecessor, clsItemBase)
	INTERFACE_PART(clsPredecessor, IID_IclsPredecessor, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsPredecessor, clsItemBase)
END_MESSAGE_MAP()

clsPredecessor::clsPredecessor()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsPredecessor. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsPredecessor::clsPredecessor(CActiveGanttVCCtl* oControl, clsPredecessors* oPredecessors)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_bVisible = TRUE;
	mp_clsPredecessors = oPredecessors;
	mp_sPredecessorKey = "";
	mp_sSuccessorKey = "";
	mp_sStyleIndex = "DS_PREDECESSOR";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_PREDECESSOR");
	mp_sTag = "";
    mp_yLagInterval = IL_DAY;
    mp_lLagFactor = 0;
    mp_bWarning = FALSE;
    mp_sWarningStyleIndex = "";
    mp_sSelectedStyleIndex = "";
}

clsPredecessor::~clsPredecessor()
{
	AfxOleUnlockApp();
}

void clsPredecessor::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

CString clsPredecessor::GetKey(void)
{
	return mp_sKey;
}

void clsPredecessor::SetKey(CString Value)
{
	mp_clsPredecessors->mp_oCollection->mp_SetKey(&mp_sKey, Value, PREDECESSORS_SET_KEY);
}

BOOL clsPredecessor::GetVisible(void)
{
	return mp_bVisible;
}

void clsPredecessor::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

CString clsPredecessor::GetPredecessorKey(void)
{
	return mp_sPredecessorKey;
}

void clsPredecessor::SetPredecessorKey(CString Value)
{
	mp_sPredecessorKey = Value;
	mp_oPredecessorTask = mp_oControl->GetTasks()->Item(Value);
}

clsTask* clsPredecessor::GetPredecessorTask(void)
{
	return mp_oPredecessorTask;
}

LONG clsPredecessor::GetPredecessorType(void)
{
	return mp_yPredecessorType;
}

void clsPredecessor::SetPredecessorType(LONG Value)
{
	mp_yPredecessorType = Value;
}

CString clsPredecessor::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_PREDECESSOR") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsPredecessor::SetStyleIndex(CString Value)
{
	if (Value.Trim() == "") Value = "DS_PREDECESSOR"; 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsPredecessor::GetStyle(void)
{
	return mp_oStyle;
}

CString clsPredecessor::GetTag(void)
{
	return mp_sTag;
}

void clsPredecessor::SetTag(CString Value)
{
	mp_sTag = Value;
}

LONG clsPredecessor::GetIndex(void)
{
	return mp_lIndex;
}

void clsPredecessor::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

CString clsPredecessor::GetSuccessorKey(void)
{
	return mp_sSuccessorKey;
}

void clsPredecessor::SetSuccessorKey(CString Value)
{
	mp_sSuccessorKey = Value;
	mp_oSuccessorTask = mp_oControl->GetTasks()->Item(Value);
}

clsTask* clsPredecessor::GetSuccessorTask(void)
{
	return mp_oSuccessorTask;
}

LONG clsPredecessor::GetLagInterval(void)
{
	return mp_yLagInterval;
}

void clsPredecessor::SetLagInterval(LONG Value)
{
	mp_yLagInterval = Value;
}

LONG clsPredecessor::GetLagFactor(void)
{
	return mp_lLagFactor;
}

void clsPredecessor::SetLagFactor(LONG Value)
{
	mp_lLagFactor = Value;
}

void clsPredecessor::AddRectangle(S_Rectangle oRectangle)
{
	mp_oRectangles.Add(oRectangle);
}

void clsPredecessor::ClearRectangles()
{
	mp_oRectangles.RemoveAll();
}

BOOL clsPredecessor::HitTest(int X, int Y)
{
	int i = 0;
	for (i = 0; i <= mp_oRectangles.GetCount() - 1; i++) 
	{
		S_Rectangle oRectangle;
		oRectangle = mp_oRectangles.GetAt(i);
		if (oRectangle.mp_bInRect(X, Y) == TRUE) 
		{
			return TRUE;
		}
	}
	return FALSE;
}

CString clsPredecessor::GetWarningStyleIndex(void)
{
	return mp_sWarningStyleIndex;
}

void clsPredecessor::SetWarningStyleIndex(CString Value)
{
	Value = Value.Trim();
	mp_sWarningStyleIndex = Value;
	if (Value.GetLength() > 0) 
	{
		mp_oWarningStyle = mp_oControl->GetStyles()->FItem(Value);
	}
	else 
	{
		mp_oWarningStyle = NULL;
	}
}

clsStyle* clsPredecessor::GetWarningStyle(void)
{
	if (mp_sWarningStyleIndex.GetLength() == 0) 
	{
		return mp_oStyle;
	}
	else 
	{
		return mp_oWarningStyle;
	}
}

CString clsPredecessor::GetSelectedStyleIndex(void)
{
	return mp_sSelectedStyleIndex;
}

void clsPredecessor::SetSelectedStyleIndex(CString Value)
{
	Value = Value.Trim();
	mp_sSelectedStyleIndex = Value;
	if (Value.GetLength() > 0) 
	{
		mp_oSelectedStyle = mp_oControl->GetStyles()->FItem(Value);
	} 
	else 
	{
		mp_oSelectedStyle = NULL;
	}
}

clsStyle* clsPredecessor::GetSelectedStyle(void)
{
	if (mp_sSelectedStyleIndex.GetLength() == 0) 
	{
		return mp_oStyle;
	}
	else 
	{
		return mp_oSelectedStyle;
	}
}

void clsPredecessor::Check(LONG Mode)
{
	COleDateTime dtPredecessor;
	COleDateTime dtSuccessor;
	int lDiff = 0;
	int lDuration = 0;
	mp_bWarning = FALSE;
	switch (mp_yPredecessorType) 
	{
	case PCT_START_TO_START:
		dtPredecessor = mp_oPredecessorTask->GetStartDate();
		dtSuccessor = mp_oSuccessorTask->GetStartDate();
		lDiff = mp_oControl->GetMathLib()->DateTimeDiff(mp_yLagInterval, dtPredecessor, dtSuccessor);
		if (lDiff != mp_lLagFactor && Mode == PM_FORCE) 
		{
			lDuration = mp_oControl->GetMathLib()->DateTimeDiff(mp_oControl->GetCurrentViewObject()->GetInterval(), mp_oSuccessorTask->GetStartDate(), mp_oSuccessorTask->GetEndDate());
			mp_oSuccessorTask->SetStartDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yLagInterval, mp_lLagFactor, mp_oPredecessorTask->GetStartDate()));
			mp_oSuccessorTask->SetEndDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yLagInterval, lDuration, mp_oSuccessorTask->GetStartDate()));
		}
		break;
	case PCT_END_TO_END:
		dtPredecessor = mp_oPredecessorTask->GetEndDate();
		dtSuccessor = mp_oSuccessorTask->GetEndDate();
		lDiff = mp_oControl->GetMathLib()->DateTimeDiff(mp_yLagInterval, dtPredecessor, dtSuccessor);
		if (lDiff != mp_lLagFactor && Mode == PM_FORCE) 
		{
			lDuration = mp_oControl->GetMathLib()->DateTimeDiff(mp_oControl->GetCurrentViewObject()->GetInterval(), mp_oSuccessorTask->GetStartDate(), mp_oSuccessorTask->GetEndDate());
			mp_oSuccessorTask->SetEndDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yLagInterval, mp_lLagFactor, mp_oPredecessorTask->GetEndDate()));
			mp_oSuccessorTask->SetStartDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yLagInterval, -lDuration, mp_oSuccessorTask->GetEndDate()));
		}
		break;
	case PCT_START_TO_END:
		dtPredecessor = mp_oPredecessorTask->GetStartDate();
		dtSuccessor = mp_oSuccessorTask->GetEndDate();
		lDiff = mp_oControl->GetMathLib()->DateTimeDiff(mp_yLagInterval, dtPredecessor, dtSuccessor);
		if (lDiff != mp_lLagFactor && Mode == PM_FORCE) 
		{
			lDuration = mp_oControl->GetMathLib()->DateTimeDiff(mp_oControl->GetCurrentViewObject()->GetInterval(), mp_oSuccessorTask->GetStartDate(), mp_oSuccessorTask->GetEndDate());
			mp_oSuccessorTask->SetEndDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yLagInterval, mp_lLagFactor, mp_oPredecessorTask->GetStartDate()));
			mp_oSuccessorTask->SetStartDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yLagInterval, -lDuration, mp_oSuccessorTask->GetEndDate()));
		}
		break;
	case PCT_END_TO_START:
		dtPredecessor = mp_oPredecessorTask->GetEndDate();
		dtSuccessor = mp_oSuccessorTask->GetStartDate();
		lDiff = mp_oControl->GetMathLib()->DateTimeDiff(mp_yLagInterval, dtPredecessor, dtSuccessor);
		if (lDiff != mp_lLagFactor && Mode == PM_FORCE) 
		{
			lDuration = mp_oControl->GetMathLib()->DateTimeDiff(mp_oControl->GetCurrentViewObject()->GetInterval(), mp_oSuccessorTask->GetStartDate(), mp_oSuccessorTask->GetEndDate());
			mp_oSuccessorTask->SetStartDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yLagInterval, mp_lLagFactor, mp_oPredecessorTask->GetEndDate()));
			mp_oSuccessorTask->SetEndDate(mp_oControl->GetMathLib()->DateTimeAdd(mp_yLagInterval, lDuration, mp_oSuccessorTask->GetStartDate()));
		}
		break;
	}
	if (lDiff != mp_lLagFactor && Mode == PM_CREATEWARNINGFLAG) 
	{
		mp_bWarning = TRUE;
		mp_oSuccessorTask->mp_bWarning = TRUE;
	} 
	else if (lDiff != mp_lLagFactor && Mode == PM_RAISEEVENT) 
	{
		mp_oControl->GetPredecessorExceptionEventArgs()->Clear();
		mp_oControl->GetPredecessorExceptionEventArgs()->SetPredecessorIndex(GetIndex());
		mp_oControl->GetPredecessorExceptionEventArgs()->SetPredecessorType(GetPredecessorType());
		mp_oControl->FirePredecessorException();
	}
}

CString clsPredecessor::GetXML(void)
{
	clsXML oXML(mp_oControl, "Predecessor");
	oXML.InitializeWriter();
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("SuccessorKey", mp_sSuccessorKey);
	oXML.WriteProperty("PredecessorKey", mp_sPredecessorKey);
	oXML.WriteProperty("PredecessorType", mp_yPredecessorType);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("Visible", mp_bVisible);
	oXML.WriteProperty("LagInterval", mp_yLagInterval);
	oXML.WriteProperty("LagFactor", mp_lLagFactor);
    oXML.WriteProperty("WarningStyleIndex", mp_sWarningStyleIndex);
    oXML.WriteProperty("SelectedStyleIndex", mp_sSelectedStyleIndex);
	return oXML.GetXML();
}

void clsPredecessor::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "Predecessor");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("SuccessorKey", mp_sSuccessorKey);
	mp_oSuccessorTask = mp_oControl->mp_oTasks->Item(mp_sSuccessorKey);
	oXML.ReadProperty("PredecessorKey", mp_sPredecessorKey);
	mp_oPredecessorTask = mp_oControl->mp_oTasks->Item(mp_sPredecessorKey);
	oXML.ReadProperty("PredecessorType", mp_yPredecessorType);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("Visible", mp_bVisible);
	oXML.ReadProperty("LagInterval", mp_yLagInterval);
	oXML.ReadProperty("LagFactor", mp_lLagFactor);
    oXML.ReadProperty("WarningStyleIndex", mp_sWarningStyleIndex);
    SetWarningStyleIndex(mp_sWarningStyleIndex);
    oXML.ReadProperty("SelectedStyleIndex", mp_sSelectedStyleIndex);
    SetSelectedStyleIndex(mp_sSelectedStyleIndex);
}

BOOL clsPredecessor::GetWarning(void)
{
	return mp_bWarning;
}
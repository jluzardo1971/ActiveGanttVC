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
#include "clsTimeBlock.h"



IMPLEMENT_DYNCREATE(clsTimeBlock, clsItemBase)

//{F1B83CF4-5590-4DD5-A125-E587310A7B99}
static const IID IID_IclsTimeBlock = { 0xF1B83CF4, 0x5590, 0x4DD5, { 0xA1, 0x25, 0xE5, 0x87, 0x31, 0x0A, 0x7B, 0x99} };

//{05E7915D-7920-432D-A576-EDD3AEC25EFD}
IMPLEMENT_OLECREATE_FLAGS(clsTimeBlock, "AGVC.clsTimeBlock", afxRegApartmentThreading, 0x05E7915D, 0x7920, 0x432D, 0xA5, 0x76, 0xED, 0xD3, 0xAE, 0xC2, 0x5E, 0xFD)

BEGIN_DISPATCH_MAP(clsTimeBlock, clsItemBase)
	DISP_PROPERTY_EX_ID(clsTimeBlock, "Key", 1, odl_GetKey, odl_SetKey, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "TimeBlockType", 2, odl_GetTimeBlockType, odl_SetTimeBlockType, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "RecurringType", 3, odl_GetRecurringType, odl_SetRecurringType, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "StyleIndex", 4, odl_GetStyleIndex, odl_SetStyleIndex, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "Style", 5, odl_GetStyle, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "Tag", 6, odl_GetTag, odl_SetTag, VT_BSTR) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "GenerateConflict", 7, odl_GetGenerateConflict, odl_SetGenerateConflict, VT_BOOL) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "LeftTrim", 8, odl_GetLeftTrim, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "RightTrim", 9, odl_GetRightTrim, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "Left", 10, odl_GetLeft, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "Top", 11, odl_GetTop, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "Right", 12, odl_GetRight, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "Bottom", 13, odl_GetBottom, SetNotSupported, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "Visible", 14, odl_GetVisible, odl_SetVisible, VT_BOOL)
	DISP_PROPERTY_EX_ID(clsTimeBlock, "Index", 15, odl_GetIndex, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX_ID(clsTimeBlock, "BaseDate", 16, odl_GetBaseDate, odl_SetBaseDate, VT_DATE)
	DISP_PROPERTY_EX_ID(clsTimeBlock, "BaseWeekDay", 17, odl_GetBaseWeekDay, odl_SetBaseWeekDay, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "DurationInterval", 18, odl_GetDurationInterval, odl_SetDurationInterval, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "DurationFactor", 19, odl_GetDurationFactor, odl_SetDurationFactor, VT_I4) 
	DISP_PROPERTY_EX_ID(clsTimeBlock, "NonWorking", 20, odl_GetNonWorking, odl_SetNonWorking, VT_BOOL)  
	DISP_FUNCTION_ID(clsTimeBlock, "GetXML", 21, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTimeBlock, "SetXML", 22, odl_SetXML, VT_EMPTY, VTS_BSTR)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTimeBlock, clsItemBase)
	INTERFACE_PART(clsTimeBlock, IID_IclsTimeBlock, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTimeBlock, clsItemBase)
END_MESSAGE_MAP()

clsTimeBlock::clsTimeBlock()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTimeBlock. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTimeBlock::clsTimeBlock(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_sStyleIndex = "DS_TIMEBLOCK";
	mp_oStyle = mp_oControl->GetStyles()->FItem("DS_TIMEBLOCK");
	mp_sTag = "";
	mp_bGenerateConflict = FALSE;
	mp_yTimeBlockType = TBT_SINGLE_OCURRENCE;
	mp_yRecurringType = RCT_DAY;
	mp_bVisible = TRUE;
    mp_yBaseWeekDay = WD_SUNDAY;
    mp_dtBaseDate = (DATE)0;
    mp_yDurationInterval = IL_HOUR;
    mp_lDurationFactor = 1;
    mp_bNonWorking = FALSE;
}

clsTimeBlock::~clsTimeBlock()
{
	AfxOleUnlockApp();
}

void clsTimeBlock::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

CString clsTimeBlock::GetKey(void)
{
	return mp_sKey;
}

void clsTimeBlock::SetKey(CString Value)
{
	mp_oControl->GetTimeBlocks()->mp_oCollection->mp_SetKey(&mp_sKey, Value, TIMEBLOCKS_SET_KEY);
}

LONG clsTimeBlock::GetTimeBlockType(void)
{
	return mp_yTimeBlockType;
}

void clsTimeBlock::SetTimeBlockType(LONG Value)
{
	mp_yTimeBlockType = Value;
}

LONG clsTimeBlock::GetRecurringType(void)
{
	return mp_yRecurringType;
}

void clsTimeBlock::SetRecurringType(LONG Value)
{
	mp_yRecurringType = Value;
}

COleDateTime clsTimeBlock::GetEndDate(void)
{
	if (mp_lDurationFactor > 0)
	{
		return mp_oControl->GetMathLib()->DateTimeAdd(mp_yDurationInterval, mp_lDurationFactor, mp_dtBaseDate);
	}
	else
	{
		return mp_dtBaseDate;
	}
}

COleDateTime clsTimeBlock::GetStartDate(void)
{
	if (mp_lDurationFactor > 0)
	{
		return mp_dtBaseDate;
	}
	else
	{
		return mp_oControl->GetMathLib()->DateTimeAdd(mp_yDurationInterval, mp_lDurationFactor, mp_dtBaseDate);
	}
}

CString clsTimeBlock::GetStyleIndex(void)
{
	if (mp_sStyleIndex == "DS_TIMEBLOCK") 
	{
		return "";
	}
	else 
	{
		return mp_sStyleIndex;
	}
}

void clsTimeBlock::SetStyleIndex(CString Value)
{
	if (Value.Trim() == "") Value = "DS_TIMEBLOCK"; 
	mp_sStyleIndex = Value;
	mp_oStyle = mp_oControl->GetStyles()->FItem(Value);
}

clsStyle* clsTimeBlock::GetStyle(void)
{
	return mp_oStyle;
}

CString clsTimeBlock::GetTag(void)
{
	return mp_sTag;
}

void clsTimeBlock::SetTag(CString Value)
{
	mp_sTag = Value;
}

BOOL clsTimeBlock::GetGenerateConflict(void)
{
	return mp_bGenerateConflict;
}

void clsTimeBlock::SetGenerateConflict(BOOL Value)
{
	mp_bGenerateConflict = Value;
}

LONG clsTimeBlock::GetLeftTrim(void)
{
	if (mp_yTimeBlockType == TBT_SINGLE_OCURRENCE) 
	{
		if (GetLeft() < mp_oControl->GetSplitter()->GetRight()) 
		{
			return mp_oControl->GetSplitter()->GetRight();
		}
		else 
		{
			return GetLeft();
		}
	}
	else 
	{
		return 0;
	}
}

LONG clsTimeBlock::GetRightTrim(void)
{
	if (mp_yTimeBlockType == TBT_SINGLE_OCURRENCE) 
	{
		if (GetRight() > mp_oControl->Getmt_RightMargin()) 
		{
			return mp_oControl->Getmt_RightMargin();
		}
		else 
		{
			return GetRight();
		}
	}
	else 
	{
		return 0;
	}
}

LONG clsTimeBlock::GetLeft(void)
{

    if (mp_yTimeBlockType == TBT_SINGLE_OCURRENCE)
    {
        return mp_oControl->GetMathLib()->GetXCoordinateFromDate(GetStartDate());
    }
    else
    {
        return 0;
    }
}

LONG clsTimeBlock::GetTop(void)
{
	return mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop();
}

LONG clsTimeBlock::GetRight(void)
{

    if (mp_yTimeBlockType == TBT_SINGLE_OCURRENCE)
    {
        return mp_oControl->GetMathLib()->GetXCoordinateFromDate(GetEndDate());
    }
    else
    {
        return 0;
    }
}

LONG clsTimeBlock::GetBottom(void)
{
    if (mp_oControl->GetTimeBlockBehaviour() == TBB_CONTROLEXTENTS)
    {
        return mp_oControl->Getmt_BottomMargin();
    }
    else if (mp_oControl->GetRows()->GetCount() > 0)
    {
        return mp_oControl->GetRows()->Item(CStr(mp_oControl->GetCurrentViewObject()->GetClientArea()->GetLastVisibleRow()))->GetBottom();
    }
    else
    {
        return mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop();
    }
}

BOOL clsTimeBlock::GetVisible(void)
{
	if (mp_oControl->GetRows()->GetCount() == 0) 
	{
		return FALSE;
	}
	if (mp_yTimeBlockType == TBT_RECURRING) 
	{
		return FALSE;
	}
    COleDateTime dtStartDate;
    COleDateTime dtEndDate;
    dtStartDate = GetStartDate();
    dtEndDate = GetEndDate();
	if ((((dtStartDate >= mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate() && dtStartDate <= mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetEndDate()) || (dtEndDate >= mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate() && dtEndDate <= mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetEndDate())) || (dtStartDate < mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate() && dtEndDate > mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetEndDate()))) 
	{
		return mp_bVisible;
	}
	else 
	{
		return FALSE;
	}
}

void clsTimeBlock::SetVisible(BOOL Value)
{
	mp_bVisible = Value;
}

COleDateTime clsTimeBlock::GetBaseDate(void)
{
	return mp_dtBaseDate;
}

void clsTimeBlock::SetBaseDate(COleDateTime Value)
{
	mp_dtBaseDate = Value;
}

LONG clsTimeBlock::GetBaseWeekDay(void)
{
	return mp_yBaseWeekDay;
}

void clsTimeBlock::SetBaseWeekDay(LONG Value)
{
	mp_yBaseWeekDay = Value;
}

LONG clsTimeBlock::GetDurationInterval(void)
{
	return mp_yDurationInterval;
}

void clsTimeBlock::SetDurationInterval(LONG Value)
{
	mp_yDurationInterval = Value;
}

LONG clsTimeBlock::GetDurationFactor(void)
{
	return mp_lDurationFactor;
}

void clsTimeBlock::SetDurationFactor(LONG Value)
{
	mp_lDurationFactor = Value;
}

BOOL clsTimeBlock::GetNonWorking(void)
{
	return mp_bNonWorking;
}

void clsTimeBlock::SetNonWorking(BOOL Value)
{
	mp_bNonWorking = Value;
}

LONG clsTimeBlock::GetIndex(void)
{
	return mp_lIndex;
}

void clsTimeBlock::SetIndex(LONG Value)
{
	mp_lIndex = Value;
}

BOOL clsTimeBlock::Getf_Visible(void)
{
	return mp_bVisible;
}

void clsTimeBlock::Setf_Visible(BOOL Value)
{
	mp_bVisible = Value;
}

CString clsTimeBlock::GetXML(void)
{
	clsXML oXML(mp_oControl, "TimeBlock");
	oXML.InitializeWriter();
	oXML.WriteProperty("GenerateConflict", mp_bGenerateConflict);
	oXML.WriteProperty("Key", mp_sKey);
	oXML.WriteProperty("RecurringType", mp_yRecurringType);
	oXML.WriteProperty("StyleIndex", mp_sStyleIndex);
	oXML.WriteProperty("Tag", mp_sTag);
	oXML.WriteProperty("TimeBlockType", mp_yTimeBlockType);
	oXML.WriteProperty("Visible", mp_bVisible);
    oXML.WriteProperty("BaseDate", mp_dtBaseDate);
    oXML.WriteProperty("BaseWeekDay", mp_yBaseWeekDay);
    oXML.WriteProperty("DurationInterval", mp_yDurationInterval);
    oXML.WriteProperty("DurationFactor", mp_lDurationFactor);
    oXML.WriteProperty("NonWorking", mp_bNonWorking);
	return oXML.GetXML();
}

void clsTimeBlock::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TimeBlock");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("GenerateConflict", mp_bGenerateConflict);
	oXML.ReadProperty("Key", mp_sKey);
	oXML.ReadProperty("RecurringType", mp_yRecurringType);
	oXML.ReadProperty("StyleIndex", mp_sStyleIndex);
	SetStyleIndex(mp_sStyleIndex);
	oXML.ReadProperty("Tag", mp_sTag);
	oXML.ReadProperty("TimeBlockType", mp_yTimeBlockType);
	oXML.ReadProperty("Visible", mp_bVisible);
    oXML.ReadProperty("BaseDate", mp_dtBaseDate);
    oXML.ReadProperty("BaseWeekDay", mp_yBaseWeekDay);
    oXML.ReadProperty("DurationInterval", mp_yDurationInterval);
    oXML.ReadProperty("DurationFactor", mp_lDurationFactor);
    oXML.ReadProperty("NonWorking", mp_bNonWorking);
}
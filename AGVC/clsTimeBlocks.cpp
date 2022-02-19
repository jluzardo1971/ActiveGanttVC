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
#include "clsTimeBlocks.h"



IMPLEMENT_DYNCREATE(clsTimeBlocks, CCmdTarget)

//{EF7A5D9F-E35D-421F-AC1F-43934AE75DE5}
static const IID IID_IclsTimeBlocks = { 0xEF7A5D9F, 0xE35D, 0x421F, { 0xAC, 0x1F, 0x43, 0x93, 0x4A, 0xE7, 0x5D, 0xE5} };

//{6CE166AD-2CC0-45C9-B733-45E54154B817}
IMPLEMENT_OLECREATE_FLAGS(clsTimeBlocks, "AGVC.clsTimeBlocks", afxRegApartmentThreading, 0x6CE166AD, 0x2CC0, 0x45C9, 0xB7, 0x33, 0x45, 0xE5, 0x41, 0x54, 0xB8, 0x17)

BEGIN_DISPATCH_MAP(clsTimeBlocks, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTimeBlocks, "Count", 1, odl_GetCount, SetNotSupported, VT_I4) 
	DISP_PROPERTY_PARAM_ID(clsTimeBlocks, "Item", 2, odl_Item, SetNotSupported, VT_DISPATCH, VTS_BSTR) 
	DISP_FUNCTION_ID(clsTimeBlocks, "Clear", 3, odl_Clear, VT_EMPTY, VTS_NONE) 
	DISP_FUNCTION_ID(clsTimeBlocks, "Remove", 4, odl_Remove, VT_EMPTY, VTS_BSTR) 
	DISP_FUNCTION_ID(clsTimeBlocks, "GetXML", 5, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTimeBlocks, "SetXML", 6, odl_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(clsTimeBlocks, "Add", 7, odl_Add, VT_DISPATCH, VTS_BSTR)
	DISP_PROPERTY_EX_ID(clsTimeBlocks, "IntervalStart", 8, odl_GetIntervalStart, odl_SetIntervalStart, VT_DATE) 
	DISP_PROPERTY_EX_ID(clsTimeBlocks, "IntervalEnd", 9, odl_GetIntervalEnd, odl_SetIntervalEnd, VT_DATE)
	DISP_PROPERTY_EX_ID(clsTimeBlocks, "IntervalType", 10, odl_GetIntervalType, odl_SetIntervalType, VT_I4)
	DISP_FUNCTION_ID(clsTimeBlocks, "CalculateInterval", 11, odl_CalculateInterval, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(clsTimeBlocks, "CP_GetXML", 12, odl_CP_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTimeBlocks, "CP_SetXML", 13, odl_CP_SetXML, VT_EMPTY, VTS_BSTR)
	DISP_DEFVALUE(clsTimeBlocks, "Item")
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTimeBlocks, CCmdTarget)
	INTERFACE_PART(clsTimeBlocks, IID_IclsTimeBlocks, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTimeBlocks, CCmdTarget)
END_MESSAGE_MAP()

clsTimeBlocks::clsTimeBlocks(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oCollection = new clsCollectionBase(mp_oControl, "TimeBlock");
	mp_dtIntervalStart = (DATE)0;
	mp_dtIntervalEnd = (DATE)0;
	mp_yIntervalType = TBIT_AUTOMATIC;
}

clsTimeBlocks::clsTimeBlocks()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTimeBlocks. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTimeBlocks::~clsTimeBlocks()
{
	delete mp_oCollection;
	mp_oCollection = NULL;
	AfxOleUnlockApp();
}

void clsTimeBlocks::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

LONG clsTimeBlocks::GetCount(void)
{
	return mp_oCollection->m_lCount();
}

clsCollectionBase* clsTimeBlocks::GetCollection(void)
{
	return mp_oCollection;
}

clsTimeBlock* clsTimeBlocks::Item(CString Index)
{
	clsTimeBlock *oTimeBlock;
	oTimeBlock = (clsTimeBlock*)mp_oCollection->m_oItem(Index, TIMEBLOCKS_ITEM_1, TIMEBLOCKS_ITEM_2, TIMEBLOCKS_ITEM_3, TIMEBLOCKS_ITEM_4);
	return oTimeBlock;
}

clsTimeBlock* clsTimeBlocks::Add(CString Key)
{
	mp_oCollection->SetAddMode(TRUE);
	clsTimeBlock* oTimeBlock = new clsTimeBlock(mp_oControl);
	Key = g_Trim(Key);
	oTimeBlock->SetKey(Key);
	mp_oCollection->m_Add(oTimeBlock, Key, TIMEBLOCKS_ADD_1, TIMEBLOCKS_ADD_2, FALSE, TIMEBLOCKS_ADD_3);
	return oTimeBlock;
}

void clsTimeBlocks::Clear(void)
{
	mp_oCollection->m_Clear();
}

void clsTimeBlocks::Remove(CString Index)
{
	mp_oCollection->m_Remove(Index, TIMEBLOCKS_REMOVE_1, TIMEBLOCKS_REMOVE_2, TIMEBLOCKS_REMOVE_3, TIMEBLOCKS_REMOVE_4);
}

void clsTimeBlocks::CreateTemporaryTimeBlocks(void)
{
	int lIndex = 0;
	mp_oControl->TempTimeBlocks()->Clear();

	for (lIndex = 1; lIndex <= GetCount(); lIndex++) 
	{
		clsTimeBlock* oTimeBlock = NULL;
		clsTimeBlock* oTempTimeBlock = NULL;
		COleDateTime dtTimeLineStart;
		COleDateTime dtTimeLineEnd;
		COleDateTime dtCurrent;
		COleDateTime dtStartBuff;
		COleDateTime dtEndBuff;
		COleDateTime dtBase;
		COleDateTime dtStartDate;
		COleDateTime dtEndDate;
		oTimeBlock = (clsTimeBlock*)mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oTimeBlock->GetTimeBlockType() == TBT_RECURRING) 
		{
			dtTimeLineStart = mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate();
			dtTimeLineEnd = mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetEndDate();
			switch (oTimeBlock->GetRecurringType())
			{
			case RCT_DAY:
				dtTimeLineStart = mp_oControl->GetMathLib()->DateTimeAdd(IL_DAY, -1, dtTimeLineStart);
				dtCurrent.SetDateTime(dtTimeLineStart.GetYear(), dtTimeLineStart.GetMonth(), dtTimeLineStart.GetDay(), 0, 0, 0);
				while (dtCurrent < dtTimeLineEnd)
				{
					dtBase.SetDateTime(dtCurrent.GetYear(), dtCurrent.GetMonth(), dtCurrent.GetDay(), oTimeBlock->GetBaseDate().GetHour(), oTimeBlock->GetBaseDate().GetMinute(), oTimeBlock->GetBaseDate().GetSecond());
					dtCurrent = mp_oControl->GetMathLib()->DateTimeAdd(IL_DAY, 1, dtCurrent);
					if (mp_oControl->GetMathLib()->mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor()) == TRUE)
					{
						oTempTimeBlock = mp_oControl->TempTimeBlocks()->Add("");
						oTempTimeBlock->SetBaseDate(dtBase);
						oTempTimeBlock->SetDurationInterval(oTimeBlock->GetDurationInterval());
						oTempTimeBlock->SetDurationFactor(oTimeBlock->GetDurationFactor());
						CopyTimeBlock(oTempTimeBlock, oTimeBlock);
					}
				}
				break;
			case RCT_WEEK:
				dtTimeLineStart = mp_oControl->GetMathLib()->DateTimeAdd(IL_DAY, -7, dtTimeLineStart);
				dtCurrent.SetDateTime(dtTimeLineStart.GetYear(), dtTimeLineStart.GetMonth(), dtTimeLineStart.GetDay(), 0, 0, 0);
				while (dtCurrent < dtTimeLineEnd)
				{
					dtBase.SetDateTime(dtCurrent.GetYear(), dtCurrent.GetMonth(), dtCurrent.GetDay(), oTimeBlock->GetBaseDate().GetHour(), oTimeBlock->GetBaseDate().GetMinute(), oTimeBlock->GetBaseDate().GetSecond());
					dtCurrent = mp_oControl->GetMathLib()->DateTimeAdd(IL_DAY, 1, dtCurrent);
					if (oTimeBlock->GetBaseWeekDay() == dtBase.GetDayOfWeek() - 1)
					{
						if (mp_oControl->GetMathLib()->mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor()) == TRUE)
						{
							oTempTimeBlock = mp_oControl->TempTimeBlocks()->Add("");
							oTempTimeBlock->SetBaseDate(dtBase);
							oTempTimeBlock->SetDurationInterval(oTimeBlock->GetDurationInterval());
							oTempTimeBlock->SetDurationFactor(oTimeBlock->GetDurationFactor());
							CopyTimeBlock(oTempTimeBlock, oTimeBlock);
						}
					}
				}
				break;
			case RCT_MONTH:
				dtTimeLineStart = mp_oControl->GetMathLib()->DateTimeAdd(IL_MONTH, -1, dtTimeLineStart);
				dtCurrent.SetDateTime(dtTimeLineStart.GetYear(), dtTimeLineStart.GetMonth(), dtTimeLineStart.GetDay(), 0, 0, 0);
				while (dtCurrent < dtTimeLineEnd)
				{
					if (oTimeBlock->GetBaseDate().GetDay() == dtCurrent.GetDay())
					{
						dtBase.SetDateTime(dtCurrent.GetYear(), dtCurrent.GetMonth(), dtCurrent.GetDay(), oTimeBlock->GetBaseDate().GetHour(), oTimeBlock->GetBaseDate().GetMinute(), oTimeBlock->GetBaseDate().GetSecond());
						dtCurrent = mp_oControl->GetMathLib()->DateTimeAdd(IL_DAY, 1, dtCurrent);
						if (mp_oControl->GetMathLib()->mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor()) == TRUE)
						{
							oTempTimeBlock = mp_oControl->TempTimeBlocks()->Add("");
							oTempTimeBlock->SetBaseDate(dtBase);
							oTempTimeBlock->SetDurationInterval(oTimeBlock->GetDurationInterval());
							oTempTimeBlock->SetDurationFactor(oTimeBlock->GetDurationFactor());
							CopyTimeBlock(oTempTimeBlock, oTimeBlock);
						}
					}
					else
					{
						dtCurrent = mp_oControl->GetMathLib()->DateTimeAdd(IL_DAY, 1, dtCurrent);
					}
				}
				break;
			case RCT_YEAR:
				dtTimeLineStart = mp_oControl->GetMathLib()->DateTimeAdd(IL_YEAR, -1, dtTimeLineStart);
				dtCurrent.SetDateTime(dtTimeLineStart.GetYear(), dtTimeLineStart.GetMonth(), dtTimeLineStart.GetDay(), 0, 0, 0);
				while (dtCurrent < dtTimeLineEnd)
				{
					if (oTimeBlock->GetBaseDate().GetMonth() == dtCurrent.GetMonth())
					{
						if (oTimeBlock->GetBaseDate().GetDay() == dtCurrent.GetDay())
						{
							dtBase.SetDateTime(dtCurrent.GetYear(), dtCurrent.GetMonth(), dtCurrent.GetDay(), oTimeBlock->GetBaseDate().GetHour(), oTimeBlock->GetBaseDate().GetMinute(), oTimeBlock->GetBaseDate().GetSecond());
							dtCurrent = mp_oControl->GetMathLib()->DateTimeAdd(IL_DAY, 1, dtCurrent);
							if (mp_oControl->GetMathLib()->mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor()) == TRUE)
							{
								oTempTimeBlock = mp_oControl->TempTimeBlocks()->Add("");
								oTempTimeBlock->SetBaseDate(dtBase);
								oTempTimeBlock->SetDurationInterval(oTimeBlock->GetDurationInterval());
								oTempTimeBlock->SetDurationFactor(oTimeBlock->GetDurationFactor());
								CopyTimeBlock(oTempTimeBlock, oTimeBlock);
							}
						}
						else
						{
							dtCurrent = mp_oControl->GetMathLib()->DateTimeAdd(IL_DAY, 1, dtCurrent);
						}
					}
					else
					{
						dtCurrent = mp_oControl->GetMathLib()->DateTimeAdd(IL_MONTH, 1, dtCurrent);
					}
				}
				break;
			}
		}
	}
}

void clsTimeBlocks::CopyTimeBlock(clsTimeBlock* oDestination,clsTimeBlock* oOriginal)
{
	oDestination->SetTimeBlockType((LONG)TBT_SINGLE_OCURRENCE);
	oDestination->SetStyleIndex(oOriginal->GetStyleIndex());
	oDestination->SetGenerateConflict(oOriginal->GetGenerateConflict());
	oDestination->SetTag(oOriginal->GetTag());
	oDestination->SetNonWorking(oOriginal->GetNonWorking());
	oDestination->Setf_Visible(oOriginal->Getf_Visible());
}

void clsTimeBlocks::Draw(void)
{
	DrawClass(this);
	DrawClass(mp_oControl->TempTimeBlocks());
}

void clsTimeBlocks::DrawClass(clsTimeBlocks* oTimeBlocks)
{
	LONG lIndex = 0;
	clsTimeBlock* oTimeBlock = NULL;
	if (oTimeBlocks->GetCount() == 0) 
	{
		return;
	}
	mp_oControl->clsG->mp_ClipRegion(mp_oControl->GetSplitter()->GetRight(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetTop(), mp_oControl->Getmt_RightMargin(), mp_oControl->GetCurrentViewObject()->GetClientArea()->GetBottom(), TRUE);
	for (lIndex = 1; lIndex <= oTimeBlocks->GetCount(); lIndex++) 
	{
		oTimeBlock = (clsTimeBlock*)oTimeBlocks->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oTimeBlock->GetVisible() == TRUE) 
		{
			mp_oControl->GetDrawEventArgs()->Clear();
			mp_oControl->GetDrawEventArgs()->SetCustomDraw(FALSE);
			mp_oControl->GetDrawEventArgs()->SetEventTarget((LONG)EVT_TIMEBLOCK);
			mp_oControl->GetDrawEventArgs()->SetObjectIndex(lIndex);
			mp_oControl->GetDrawEventArgs()->SetParentObjectIndex(0);
			mp_oControl->FireDraw();
			if (mp_oControl->GetDrawEventArgs()->GetCustomDraw() == FALSE) 
			{
				if ((oTimeBlock->GetRight() - oTimeBlock->GetLeft()) >= 1)
				{
					mp_oControl->clsG->mp_DrawItem(oTimeBlock->GetLeft(), oTimeBlock->GetTop(), oTimeBlock->GetRight(), oTimeBlock->GetBottom(), "", "", FALSE, NULL, oTimeBlock->GetLeftTrim(), oTimeBlock->GetRightTrim(), oTimeBlock->GetStyle());
				}
			}
		}
	}
}

COleDateTime clsTimeBlocks::GetIntervalStart(void)
{
	return mp_dtIntervalStart;
}

void clsTimeBlocks::SetIntervalStart(COleDateTime Value)
{
	mp_dtIntervalStart = Value;
}

COleDateTime clsTimeBlocks::GetIntervalEnd(void)
{
	return mp_dtIntervalEnd;
}

void clsTimeBlocks::SetIntervalEnd(COleDateTime Value)
{
	mp_dtIntervalEnd = Value;
}

LONG clsTimeBlocks::GetIntervalType(void)
{
	return mp_yIntervalType;
}

void clsTimeBlocks::SetIntervalType(LONG Value)
{
	mp_yIntervalType = Value;
}

void clsTimeBlocks::CalculateInterval(void)
{
    if (mp_yIntervalType == TBIT_AUTOMATIC)
    {
        return;
    }
    mp_aTimeBlocks.clear();
    mp_oControl->GetMathLib()->mp_GenerateTimeBlocks(mp_aTimeBlocks, mp_dtIntervalStart, mp_dtIntervalEnd);
}

CString clsTimeBlocks::CP_GetXML(void)
{
	clsXML oXML(mp_oControl, "CP_TimeBlocks");
	oXML.InitializeWriter();
	oXML.WriteProperty("IntervalStart", mp_dtIntervalStart);
	oXML.WriteProperty("IntervalEnd", mp_dtIntervalEnd);
	oXML.WriteProperty("IntervalType", mp_yIntervalType);
	return oXML.GetXML();
}

void clsTimeBlocks::CP_SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "CP_TimeBlocks");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("IntervalStart", mp_dtIntervalStart);
	oXML.ReadProperty("IntervalEnd", mp_dtIntervalEnd);
	oXML.ReadProperty("IntervalType", mp_yIntervalType);
}

CString clsTimeBlocks::GetXML(void)
{
	LONG lIndex;
	clsTimeBlock* oTimeBlock; 
	clsXML oXML(mp_oControl, "TimeBlocks");
	oXML.InitializeWriter();
	for (lIndex = 1; lIndex <= GetCount(); lIndex++)
	{
		oTimeBlock = (clsTimeBlock*) mp_oCollection->m_oReturnArrayElement(lIndex);
		oXML.WriteObject(oTimeBlock->GetXML());
	}
	return oXML.GetXML();
}

void clsTimeBlocks::SetXML(CString sXML)
{
	LONG lIndex;
	clsXML oXML(mp_oControl, "TimeBlocks");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	mp_oCollection->m_Clear();
	for (lIndex = 1;lIndex <= oXML.ReadCollectionCount();lIndex++)
	{
		clsTimeBlock* oTimeBlock = new clsTimeBlock(mp_oControl);
		oTimeBlock->SetXML(oXML.ReadCollectionObject(lIndex));
		mp_oCollection->SetAddMode(TRUE);
		mp_oCollection->m_Add(oTimeBlock, oTimeBlock->mp_sKey, TIMEBLOCKS_ADD_1, TIMEBLOCKS_ADD_2, FALSE, TIMEBLOCKS_ADD_3);
	}
}
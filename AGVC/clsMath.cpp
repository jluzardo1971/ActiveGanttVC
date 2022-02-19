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
#include "clsMath.h"
#include "math.h"



IMPLEMENT_DYNCREATE(clsMath, CCmdTarget)

//{FA672F01-E3CD-468A-98D2-617FA388F7E1}
static const IID IID_IclsMath = { 0xFA672F01, 0xE3CD, 0x468A, { 0x98, 0xD2, 0x61, 0x7F, 0xA3, 0x88, 0xF7, 0xE1} };

//{05A22D3B-8501-45A4-905B-D5A173D92B04}
IMPLEMENT_OLECREATE_FLAGS(clsMath, "AGVC.clsMath", afxRegApartmentThreading, 0x05A22D3B, 0x8501, 0x45A4, 0x90, 0x5B, 0xD5, 0xA1, 0x73, 0xD9, 0x2B, 0x04)

BEGIN_DISPATCH_MAP(clsMath, CCmdTarget)
	DISP_FUNCTION_ID(clsMath, "DateTimeAdd", 1, odl_DateTimeAdd, VT_DATE, VTS_I4 VTS_I4 VTS_DATE) 
	DISP_FUNCTION_ID(clsMath, "DateTimeDiff", 2, odl_DateTimeDiff, VT_I4, VTS_I4 VTS_DATE VTS_DATE) 
	DISP_FUNCTION_ID(clsMath, "GetXCoordinateFromDate", 3, odl_GetXCoordinateFromDate, VT_I4, VTS_DATE) 
	DISP_FUNCTION_ID(clsMath, "GetDateFromXCoordinate", 4, odl_GetDateFromXCoordinate, VT_DATE, VTS_I4) 
	DISP_FUNCTION_ID(clsMath, "GetRowIndexByPosition", 5, odl_GetRowIndexByPosition, VT_I4, VTS_I4) 
	DISP_FUNCTION_ID(clsMath, "GetCellIndexByPosition", 6, odl_GetCellIndexByPosition, VT_I4, VTS_I4) 
	DISP_FUNCTION_ID(clsMath, "GetColumnIndexByPosition", 7, odl_GetColumnIndexByPosition, VT_I4, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsMath, "GetTaskIndexByPosition", 8, odl_GetTaskIndexByPosition, VT_I4, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsMath, "GetPercentageIndexByPosition", 9, odl_GetPercentageIndexByPosition, VT_I4, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsMath, "GetNodeIndexByCheckBoxPosition", 10, odl_GetNodeIndexByCheckBoxPosition, VT_I4, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsMath, "GetNodeIndexBySignPosition", 11, odl_GetNodeIndexBySignPosition, VT_I4, VTS_I4 VTS_I4) 
	DISP_FUNCTION_ID(clsMath, "DetectConflict", 12, odl_DetectConflict, VT_BOOL, VTS_DATE VTS_DATE VTS_BSTR VTS_I4 VTS_BSTR) 
	DISP_FUNCTION_ID(clsMath, "RoundDate", 13, odl_RoundDate, VT_DATE, VTS_I4 VTS_I4 VTS_DATE) 
	DISP_FUNCTION_ID(clsMath, "RoundDouble", 14, odl_RoundDouble, VT_I4, VTS_R8)
	DISP_FUNCTION_ID(clsMath, "GetPredecessorIndexByPosition", 15, odl_GetPredecessorIndexByPosition, VT_I4, VTS_I4 VTS_I4)
	DISP_FUNCTION_ID(clsMath, "CalculateDuration", 16, odl_CalculateDuration, VT_I4, VTS_DATE VTS_DATE VTS_I4)
	DISP_FUNCTION_ID(clsMath, "GetEndDate", 17, odl_GetEndDate, VT_DATE, VTS_DATE VTS_I4 VTS_I4)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsMath, CCmdTarget)
	INTERFACE_PART(clsMath, IID_IclsMath, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsMath, CCmdTarget)
END_MESSAGE_MAP()

clsMath::clsMath()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsMath. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsMath::clsMath(CActiveGanttVCCtl* oControl)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
}

clsMath::~clsMath()
{
	AfxOleUnlockApp();
}

void clsMath::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

void clsMath::Finalize(void)
{
}

COleDateTime clsMath::DateTimeAdd(LONG Interval,LONG Number, COleDateTime dtDate)
{
	COleDateTimeSpan dtSpan;
	COleDateTime dtReturn;
	LONG lYear = 0;
	LONG lMonth = 0;
	LONG lDay = 0;
	LONG lHour = 0;
	LONG lMinute = 0;
	LONG lSecond = 0;

	BOOL bNegative = FALSE;
	if (Number < 0)
	{
		bNegative = TRUE;
		Number = abs(Number);
	}

	if (Interval == IL_SECOND)
	{
		lDay = (Number / 86400);
		Number = Number - (lDay * 86400);
		lHour = (Number / 3600);
		Number = Number - (lHour * 3600);
		lMinute = (Number / 60);
		Number = Number - (lMinute * 60);
		lSecond = Number;
		dtSpan.SetDateTimeSpan(lDay,lHour,lMinute, lSecond);
	}
	else if (Interval == IL_MINUTE)
	{
		lDay = (Number / 1440);
		Number = Number - (lDay * 1440);
		lHour = (Number / 60);
		Number = Number - (lHour * 60);
		lMinute = Number;
		dtSpan.SetDateTimeSpan(lDay,lHour,lMinute,0);
	}
	else if (Interval == IL_HOUR)
	{
		lDay = (Number / 24);
		Number = Number - (lDay * 24);
		lHour = Number;
		dtSpan.SetDateTimeSpan(lDay,lHour,0,0);
	}
	else if (Interval == IL_DAY)
	{
		dtSpan.SetDateTimeSpan(Number,0,0,0);
	}
	else if (Interval == IL_WEEK)
	{
		dtSpan.SetDateTimeSpan(Number * 7,0,0,0);
	}
	else if (Interval == IL_MONTH)
	{
		if (bNegative == FALSE)
		{
			lYear = (Number / 12);
			lMonth = Number - (lYear * 12);
			if (lMonth + dtDate.GetMonth() > 12)
			{
				lYear = lYear + 1;
				lMonth = lMonth + dtDate.GetMonth()  - 12;
			}
			else
			{
				lMonth = dtDate.GetMonth() + lMonth;
			}
			lYear = dtDate.GetYear() + lYear;
		}
		else
		{
			lYear = Number / 12;
			lMonth = Number - (lYear * 12);
			if (dtDate.GetMonth() - lMonth < 1)
			{
				lYear = lYear + 1;
				lMonth = dtDate.GetMonth() - lMonth + 12;
			}
			else
			{
				lMonth = dtDate.GetMonth() - lMonth;
			}
			lYear = dtDate.GetYear() - lYear;
		}

		lDay = dtDate.GetDay();
		lHour = dtDate.GetHour();
		lMinute = dtDate.GetMinute();
		lSecond = dtDate.GetSecond();
		dtReturn.SetDateTime(lYear,lMonth,lDay,lHour,lMinute,lSecond);
		while (dtReturn.GetStatus() != 0)
		{
			lDay = lDay - 1;
			dtReturn.SetDateTime(lYear,lMonth,lDay,lHour,lMinute,lSecond);
		}
		return dtReturn;
	}
	else if (Interval == IL_QUARTER)
	{
		if (bNegative == FALSE)
		{
			Number = Number * 3;
			lYear = (Number / 12);
			lMonth = Number - (lYear * 12);
			if (lMonth + dtDate.GetMonth() > 12)
			{
				lYear = lYear + 1;
				lMonth = lMonth + dtDate.GetMonth()  - 12;
			}
			else
			{
				lMonth = dtDate.GetMonth() + lMonth;
			}
			lYear = dtDate.GetYear() + lYear;
		}
		else
		{
			Number = Number * 3;
			lYear = -(Number / 12);
			lMonth = Number - (lYear * 12);
			if (dtDate.GetMonth() - lMonth < 1)
			{
				lYear = lYear - 1;
				lMonth = dtDate.GetMonth() - lMonth  + 12;
			}
			else
			{
				lMonth = dtDate.GetMonth() - lMonth;
			}
			lYear = dtDate.GetYear() + lYear;
		}

		lDay = dtDate.GetDay();
		lHour = dtDate.GetHour();
		lMinute = dtDate.GetMinute();
		lSecond = dtDate.GetSecond();
		dtReturn.SetDateTime(lYear,lMonth,lDay,lHour,lMinute,lSecond);
		while (dtReturn.GetStatus() != 0)
		{
			lDay = lDay - 1;
			dtReturn.SetDateTime(lYear,lMonth,lDay,lHour,lMinute,lSecond);
		}
		return dtReturn;
	}
	else if (Interval == IL_YEAR)
	{
		if (bNegative == FALSE)
		{
			lYear = dtDate.GetYear() + Number;
		}
		else
		{
			lYear = dtDate.GetYear() - Number;
		}
		lMonth = dtDate.GetMonth();
		lDay = dtDate.GetDay();
		lHour = dtDate.GetHour();
		lMinute = dtDate.GetMinute();
		lSecond = dtDate.GetSecond();
		dtReturn.SetDateTime(lYear,lMonth,lDay,lHour,lMinute,lSecond);
		while (dtReturn.GetStatus() != 0)
		{
			lDay = lDay - 1;
			dtReturn.SetDateTime(lYear,lMonth,lDay,lHour,lMinute,lSecond);
		}
		return dtReturn;
	}

	if (bNegative == FALSE)
	{
		dtReturn = dtDate + dtSpan;
	}
	else
	{
		dtReturn = dtDate - dtSpan;
		
	}
	while (dtReturn.GetStatus() != 0)
	{
		lDay = lDay - 1;
		dtReturn.SetDateTime(lYear,lMonth,lDay,lHour,lMinute,lSecond);
	}
	return dtReturn;
}

LONG clsMath::DateTimeDiff(LONG Interval,COleDateTime dtDate1,COleDateTime dtDate2)
{
	DOUBLE dDateSpan;
	dDateSpan = dtDate2 - dtDate1;
	LONG lReturn;


	if (Interval == IL_SECOND)
	{
		DOUBLE dSeconds = dDateSpan * 86400;
		lReturn = CInt32(dSeconds);
		return lReturn;
	}
	else if (Interval == IL_MINUTE)
	{
		DOUBLE dMinutes = dDateSpan * 1440;
		lReturn = CInt32(dMinutes);
		return lReturn;
	}
	else if (Interval == IL_HOUR)
	{
		DOUBLE dHours = dDateSpan * 24;
		lReturn = CInt32(dHours);
		return lReturn;
	}
	else if (Interval == IL_DAY)
	{
		lReturn = CInt32(dDateSpan);
		return lReturn;
	}
	else if (Interval == IL_WEEK)
	{

		DOUBLE dWeeks = dDateSpan / 7;
		lReturn = CInt32(dWeeks);
		return lReturn;
	}
	else if (Interval == IL_MONTH)
	{
		LONG lYearDiff = dtDate2.GetYear() - dtDate1.GetYear();
		LONG lMonthDiff = dtDate2.GetMonth() - dtDate1.GetMonth();
		lYearDiff = (lYearDiff * 12);
		return lYearDiff + lMonthDiff;	
	}
	else if (Interval == IL_YEAR)
	{
		LONG lYearDiff = dtDate2.GetYear() - dtDate1.GetYear();
		return lYearDiff;
	}
	else
	{
		return 0;
	}
}

COleDateTime clsMath::NewDateTime(LONG Year,LONG Month,LONG Day,LONG Hour,LONG Minute,LONG Second)
{
	COleDateTime dtReturn;
	dtReturn.SetDateTime(Year, Month, Day, Hour, Minute, Second);
	return dtReturn;
}

LONG clsMath::GetXCoordinateFromDate(COleDateTime dtCoordinate)
{
	return (DateTimeDiff(mp_oControl->GetCurrentViewObject()->GetInterval(), mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate(), dtCoordinate) / mp_oControl->GetCurrentViewObject()->GetFactor()) + mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart();
}

COleDateTime clsMath::GetDateFromXCoordinate(LONG v_lXCoordinate)
{
	return DateTimeAdd(mp_oControl->GetCurrentViewObject()->GetInterval(), (v_lXCoordinate - mp_oControl->GetCurrentViewObject()->GetTimeLine()->Getf_lStart()) * mp_oControl->GetCurrentViewObject()->GetFactor(), mp_oControl->GetCurrentViewObject()->GetTimeLine()->GetStartDate());
}

LONG clsMath::GetRowIndexByPosition(LONG Y)
{
	clsRow* oRow;
	LONG lVisRowIndex;
	if ((mp_oControl->GetRows()->GetCount() == 0))
	{
		return -1;
	}
	for (lVisRowIndex = mp_oControl->GetVerticalScrollBar()->GetValue();lVisRowIndex <= mp_oControl->GetCurrentViewObject()->GetClientArea()->GetLastVisibleRow();lVisRowIndex++)
	{
		oRow = (clsRow*) mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lVisRowIndex);
		if (Y >= oRow->GetTop() && Y <= oRow->GetBottom() && oRow->GetVisible() == TRUE)
		{
			return lVisRowIndex;
		}
	}
	return -1;
}

LONG clsMath::GetCellIndexByPosition(LONG X)
{
	clsColumn* oColumn;
	LONG lIndex;
	for (lIndex = 1;lIndex <= mp_oControl->GetColumns()->GetCount();lIndex++)
	{
		oColumn = (clsColumn*) mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (X > oColumn->GetLeft() && X < oColumn->GetRight())
		{
			return lIndex;
		}
	}
	return -1;
}

LONG clsMath::GetColumnIndexByPosition(LONG X,LONG Y)
{
	clsColumn* oColumn;
	LONG lIndex;
	for (lIndex = 1;lIndex <= mp_oControl->GetColumns()->GetCount();lIndex++)
	{
		oColumn = (clsColumn*) mp_oControl->GetColumns()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (X >= oColumn->GetLeft() && X <= oColumn->GetRight() && Y >= oColumn->GetTop() && Y <= oColumn->GetBottom())
		{
			return lIndex;
		}
	}
	return -1;
}

LONG clsMath::GetTaskIndexByPosition(LONG X,LONG Y)
{
	clsTask* oTask;
	LONG lIndex;
	for (lIndex = mp_oControl->GetTasks()->GetCount();lIndex >= 1 ;lIndex--)
	{
		oTask = (clsTask*) mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oTask->GetVisible() == TRUE && InCurrentLayer(oTask->GetLayerIndex()))
		{
			if (X >= oTask->GetLeft() && X <= oTask->GetRight() && Y >= oTask->GetTop() && Y <= oTask->GetBottom())
			{
				return lIndex;
			}
		}
	}
	return -1;
}

LONG clsMath::GetPredecessorIndexByPosition(LONG X,LONG Y)
{
	clsPredecessor* oPredecessor;
	LONG lIndex = 0;
	for (lIndex = mp_oControl->GetPredecessors()->GetCount(); lIndex >= 1; lIndex--)
	{
		oPredecessor = (clsPredecessor*)mp_oControl->GetPredecessors()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oPredecessor->GetVisible() == TRUE && oPredecessor->HitTest(X, Y) == TRUE)
		{
			return lIndex;
		}
	}
	return -1;
}

LONG clsMath::GetPercentageIndexByPosition(LONG X,LONG Y)
{
	clsPercentage* oPercentage;
	LONG lIndex;
	for (lIndex = mp_oControl->GetPercentages()->GetCount();lIndex >= 1 ;lIndex--)
	{
		oPercentage = (clsPercentage*) mp_oControl->GetPercentages()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (oPercentage->GetVisible() == TRUE)
		{
			if (X >= oPercentage->GetLeft() && X <= oPercentage->GetRight() && Y >= oPercentage->GetTop() && Y <= oPercentage->GetBottom())
			{
				return lIndex;
			}
		}
	}
	return -1;
}

LONG clsMath::GetNodeIndexByCheckBoxPosition(LONG X,LONG Y)
{
	LONG lIndex;
	clsNode* oNode;
	clsRow* oRow;
	LONG lReturn;
	if (mp_oControl->GetTreeview()->GetTreeviewMode() == TM_NONE)
	{
		return -1;
	}
	lReturn = -1;
	for (lIndex = 1;lIndex <= mp_oControl->GetRows()->GetCount();lIndex++)
	{
		oRow = (clsRow*) mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (oRow->GetClientAreaVisibility() == VS_INSIDEVISIBLEAREA && X >= (oNode->GetCheckBoxLeft()) && X <= (oNode->GetCheckBoxLeft() + 13) && Y <= (oNode->GetYCenter() + 6) && Y >= (oNode->GetYCenter() - 7))
		{
			lReturn = oNode->GetIndex();
		}
	}
	return lReturn;
}

LONG clsMath::GetNodeIndexBySignPosition(LONG X,LONG Y)
{
	LONG lIndex;
	clsNode* oNode;
	clsRow* oRow;
	LONG lReturn;
	if (mp_oControl->GetTreeview()->GetTreeviewMode() == TM_NONE)
	{
		return -1;
	}
	lReturn = -1;
	for (lIndex = 1;lIndex <= mp_oControl->GetRows()->GetCount();lIndex++)
	{
		oRow = (clsRow*) mp_oControl->GetRows()->mp_oCollection->m_oReturnArrayElement(lIndex);
		oNode = oRow->GetNode();
		if (oRow->GetClientAreaVisibility() == VS_INSIDEVISIBLEAREA && X >= (oNode->GetLeft() - 5) && X <= (oNode->GetLeft() + 5) && Y <= (oNode->GetYCenter() + 5) && Y >= (oNode->GetYCenter() - 5))
		{
			lReturn = oNode->GetIndex();
		}
	}
	return lReturn;
}

BOOL clsMath::InCurrentLayer(CString sLayer)
{
	if (mp_oControl->GetLayerEnableObjects() == EC_INALLLAYERS)
	{
		return TRUE;
	}
	else
	{
		LONG lLayerIndex;
		LONG lCurrentLayerIndex;
		lLayerIndex = mp_oControl->GetLayers()->mp_oCollection->m_lReturnIndex(sLayer, TRUE);
		lCurrentLayerIndex = mp_oControl->GetLayers()->mp_oCollection->m_lReturnIndex(mp_oControl->GetCurrentLayer(), TRUE);
		if ((lLayerIndex != lCurrentLayerIndex))
		{
			return FALSE;
		}
		else
		{
			return TRUE;
		}
	}
}

BOOL clsMath::DetectConflict(COleDateTime StartDate, COleDateTime EndDate,CString RowKey,LONG ExcludeIndex,CString LayerIndex)
{
	if (mp_oControl->GetCurrentViewObject()->GetClientArea()->GetDetectConflicts() == FALSE)
	{
		return FALSE;
	}
	clsTask* oTask;
	clsTimeBlock* oTimeBlock;
	LONG lIndex = 0;
	LONG lLayerIndex = 0;
	LONG lRowIndex = 0;
	if (EndDate < StartDate) 
	{
		COleDateTime dtEndDate = StartDate;
		COleDateTime dtStartDate = EndDate;
		StartDate = dtStartDate;
		EndDate = dtEndDate;
	}
	lLayerIndex = mp_oControl->GetLayers()->mp_oCollection->m_lReturnIndex(LayerIndex, TRUE);
	if ((lLayerIndex == -1)) 
	{
		mp_oControl->mp_ErrorReport(INVALID_LAYER_INDEX, "Invalid Layer Index", "ActiveGanttVCCtl.mp_bDetectConflict");
		return FALSE;
	}
	//// Compare to other Task Objects
	for (lIndex = 1; lIndex <= mp_oControl->GetTasks()->GetCount(); lIndex++) 
	{
		oTask = (clsTask*) mp_oControl->GetTasks()->mp_oCollection->m_oReturnArrayElement(lIndex);
		if (RowKey == oTask->GetRowKey() && ExcludeIndex != lIndex) 
		{
			if ((mp_oControl->GetLayers()->mp_oCollection->m_lReturnIndex(oTask->GetLayerIndex(), TRUE) == lLayerIndex)) 
			{
				//// mp_aoTasks S------------------E
				//// interval S------------------E
				if ((StartDate == oTask->GetStartDate() || EndDate == oTask->GetEndDate()) && (StartDate != EndDate)) 
				{
					return TRUE;
				}
				//// mp_aoTasks S------------------E
				//// interval S------------------E
				if (StartDate > oTask->GetStartDate() && StartDate < oTask->GetEndDate()) 
				{
					return TRUE;
				}
				//// mp_aoTasks S------------------E
				//// interval S------------------E
				if (EndDate > oTask->GetStartDate() && EndDate < oTask->GetEndDate()) 
				{
					return TRUE;
				}
				//// mp_aoTasks S------------------E
				//// interval S-------------------------E
				if (StartDate < oTask->GetStartDate() && EndDate > oTask->GetEndDate()) 
				{
					return TRUE;
				}
				//// mp_aoTasks S--------------------------E
				//// interval S------------------E
				if (StartDate > oTask->GetStartDate() && EndDate < oTask->GetEndDate()) 
				{
					return TRUE;
				}
			}
		}
	}
	//// Compare to TimeBlock Objects 
	for (lIndex = 1; lIndex <= mp_oControl->GetTimeBlocks()->GetCount(); lIndex++) 
	{
		oTimeBlock = (clsTimeBlock*) mp_oControl->GetTimeBlocks()->mp_oCollection->m_oReturnArrayElement(lIndex);
		lRowIndex = mp_oControl->GetRows()->mp_oCollection->m_lFindIndexByKey(RowKey);

		if (oTimeBlock->GetGenerateConflict() == TRUE) 
		{
			//// mp_aoTimeBlocks S------------------E
			//// interval S------------------E
			if ((StartDate == oTimeBlock->GetStartDate() || EndDate == oTimeBlock->GetEndDate()) && (StartDate != EndDate)) 
			{
				return TRUE;
			}
			//// mp_aoTimeBlocks S------------------E
			//// interval S------------------E
			if (StartDate > oTimeBlock->GetStartDate() && StartDate < oTimeBlock->GetEndDate()) 
			{
				return TRUE;
			}
			//// mp_aoTimeBlocks S------------------E
			//// interval S------------------E
			if (EndDate > oTimeBlock->GetStartDate() && EndDate < oTimeBlock->GetEndDate()) 
			{
				return TRUE;
			}
			//// mp_aoTimeBlocks S------------------E
			//// interval S-------------------------E
			if (StartDate < oTimeBlock->GetStartDate() && EndDate > oTimeBlock->GetEndDate()) 
			{
				return TRUE;
			}
			//// mp_aoTimeBlocks S--------------------------E
			//// interval S------------------E
			if (StartDate > oTimeBlock->GetStartDate() && EndDate < oTimeBlock->GetEndDate()) 
			{
				return TRUE;
			}
		}
	}
	//// Compare to Temporary TimeBlock Objects 
	for (lIndex = 1; lIndex <= mp_oControl->TempTimeBlocks()->GetCount(); lIndex++) 
	{
		oTimeBlock = (clsTimeBlock*) mp_oControl->TempTimeBlocks()->mp_oCollection->m_oReturnArrayElement(lIndex);
		lRowIndex = mp_oControl->GetRows()->mp_oCollection->m_lFindIndexByKey(RowKey);
		if (oTimeBlock->GetGenerateConflict() == TRUE) 
		{
			//// mp_aoTimeBlocks S------------------E
			//// interval S------------------E
			if ((StartDate == oTimeBlock->GetStartDate() || EndDate == oTimeBlock->GetEndDate()) && (StartDate != EndDate)) 
			{
				return TRUE;
			}
			//// mp_aoTimeBlocks S------------------E
			//// interval S------------------E
			if (StartDate > oTimeBlock->GetStartDate() && StartDate < oTimeBlock->GetEndDate()) 
			{
				return TRUE;
			}
			//// mp_aoTimeBlocks S------------------E
			//// interval S------------------E
			if (EndDate > oTimeBlock->GetStartDate() && EndDate < oTimeBlock->GetEndDate()) 
			{
				return TRUE;
			}
			//// mp_aoTimeBlocks S------------------E
			//// interval S-------------------------E
			if (StartDate < oTimeBlock->GetStartDate() && EndDate > oTimeBlock->GetEndDate()) 
			{
				return TRUE;
			}
			//// mp_aoTimeBlocks S--------------------------E
			//// interval S------------------E
			if (StartDate > oTimeBlock->GetStartDate() && EndDate < oTimeBlock->GetEndDate()) 
			{
				return TRUE;
			}
		}
	}
	return FALSE;
}

FLOAT clsMath::PercentageComplete(LONG X1,LONG X2,LONG X)
{
	X2 = X2 - X1;
	X = X - X1;
	if (X == 0)
	{
		return 0;
	}
	else if (X == X2)
	{
		return 1;
	}
	else
	{
		return (FLOAT)X / (FLOAT)X2;
	}
}

COleDateTime clsMath::RoundDate(LONG Interval,LONG Number,COleDateTime dtDate)
{
	LONG lBuffer;
	LONG lBuffer2;
	if (Interval == IL_SECOND)
	{
		lBuffer = dtDate.GetSecond();
		lBuffer2 = Round(lBuffer, Number);
		return DateTimeAdd(IL_SECOND, lBuffer2 - lBuffer, dtDate);
	}
	else if (Interval == IL_MINUTE)
	{
		if (Number == 1)
		{
			dtDate = RoundDate(IL_SECOND, 60, dtDate);
			lBuffer = dtDate.GetSecond();
			lBuffer2 = Round(lBuffer, 60);
			return DateTimeAdd(IL_SECOND, lBuffer2 - lBuffer, dtDate);
		}
		else
		{
			dtDate = RoundDate(IL_MINUTE, 1, dtDate);
			lBuffer = dtDate.GetMinute();
			lBuffer2 = Round(lBuffer, Number);
			return DateTimeAdd(IL_MINUTE, lBuffer2 - lBuffer, dtDate);
		}
	}
	else if (Interval == IL_HOUR)
	{
		if (Number == 1)
		{
			dtDate = RoundDate(IL_MINUTE, 1, dtDate);
			lBuffer = dtDate.GetMinute();
			lBuffer2 = Round(lBuffer, 60);
			return DateTimeAdd(IL_MINUTE, lBuffer2 - lBuffer, dtDate);
		}
		else
		{
			dtDate = RoundDate(IL_HOUR, 1, dtDate);
			lBuffer = dtDate.GetHour();
			lBuffer2 = Round(lBuffer, Number);
			return DateTimeAdd(IL_HOUR, lBuffer2 - lBuffer, dtDate);
		}
	}
	else if (Interval == IL_DAY)
	{
		if (Number == 1)
		{
			dtDate = RoundDate(IL_HOUR, 1, dtDate);
			lBuffer = dtDate.GetHour();
			lBuffer2 = Round(lBuffer, 24);
			return DateTimeAdd(IL_HOUR, lBuffer2 - lBuffer, dtDate);
		}
		else
		{
			dtDate = RoundDate(IL_DAY, 1, dtDate);
			lBuffer = dtDate.GetDay();
			lBuffer2 = Round(lBuffer, Number);
			return DateTimeAdd(IL_DAY, lBuffer2 - lBuffer, dtDate);
		}
	}
	else if (Interval == IL_WEEK)
	{
		if (Number == 1)
		{
			dtDate = RoundDate(IL_DAY, 1, dtDate);
			lBuffer = NumericDayOfWeek(dtDate);
			if (lBuffer <= 3)
			{
				dtDate = DateTimeAdd(IL_DAY, -(lBuffer - 1), dtDate);
			}
			else if (lBuffer >= 4)
			{
				dtDate = DateTimeAdd(IL_DAY, 8 - lBuffer, dtDate);
			}
			return dtDate;
		}
		else
		{
			dtDate = RoundDate(IL_WEEK, 1, dtDate);
			lBuffer = dtDate.GetDay();
			lBuffer2 = Round(lBuffer, Number);
			return DateTimeAdd(IL_WEEK, lBuffer2 - lBuffer, dtDate);
		}
	}
	else if (Interval == IL_MONTH)
	{
		if (Number == 1)
		{
			COleDateTime dtNextMonth;
			dtDate = RoundDate(IL_DAY, 1, dtDate);
			lBuffer = dtDate.GetDay();
			if ( lBuffer == 1 )
			{
				return dtDate;
			}
			else if ( lBuffer >= 15 )
			{
				dtNextMonth = DateTimeAdd(IL_MONTH, 1, dtDate);
				return NewDateTime(dtNextMonth.GetYear(), dtNextMonth.GetMonth(), 1, 0, 0, 0);
			}
			else
			{
				dtNextMonth = DateTimeAdd(IL_MONTH, -1, dtDate);
				return NewDateTime(dtNextMonth.GetYear(), dtNextMonth.GetMonth(), 1, 0, 0, 0);
			}
		}
		else
		{
			dtDate = RoundDate(IL_MONTH, 1, dtDate);
			lBuffer = dtDate.GetMonth();
			lBuffer2 = Round(lBuffer - 1, Number) + 1;
			return DateTimeAdd(IL_MONTH, lBuffer2 - lBuffer, dtDate);
		}
	}
	else if (Interval == IL_QUARTER)
	{
		dtDate = RoundDate(IL_DAY, 1, dtDate);
		return RoundDate(IL_MONTH, 3, dtDate);
	}
	else if (Interval == IL_YEAR)
	{
		if (Number == 1)
		{
			dtDate = RoundDate(IL_MONTH, 1, dtDate);
			lBuffer = dtDate.GetMonth();
			lBuffer2 = Round(lBuffer, 12);
			if (lBuffer2 == 1)
			{
				return NewDateTime(dtDate.GetYear(), 1, 1, 0, 0, 0);
			}
			else if (lBuffer2 == 12)
			{
				return NewDateTime(dtDate.GetYear() + 1, 1, 1, 0, 0, 0);
			}
		}
		else
		{
			dtDate = RoundDate(IL_YEAR, 1, dtDate);
			lBuffer = dtDate.GetYear();
			lBuffer2 = Round(lBuffer, Number);
			return DateTimeAdd(IL_YEAR, lBuffer2 - lBuffer, dtDate);
		}
	}
	return dtDate;
}

LONG clsMath::RoundDouble(DOUBLE dParam)
{
	return CInt32(dParam);
}

LONG clsMath::Round(LONG v_lNumberToRound,LONG v_lRoundTo)
{
    int lRoundToHalf = 0;
    int lMultiplier = 0;
    while (v_lNumberToRound > v_lRoundTo)
	{
        v_lNumberToRound = v_lNumberToRound - v_lRoundTo;
        lMultiplier = lMultiplier + 1;
    }
    lRoundToHalf = CInt32( (double) v_lRoundTo / (double) 2);
    if (v_lNumberToRound >= lRoundToHalf)
	{
        v_lNumberToRound = v_lRoundTo;
	}
    else
	{
        v_lNumberToRound = 0;
    }
    return (v_lRoundTo * lMultiplier) + v_lNumberToRound;
}

LONG clsMath::lAbs(LONG Number)
{
	return abs(Number);
}

BYTE clsMath::Rnd()
{
	int yResult;
	int lowest=0, highest=255;
	int range=(highest-lowest)+1;
	yResult = lowest+int(range*rand()/(RAND_MAX + 1.0));
	return (BYTE)yResult;
}

LONG clsMath::Quarter(COleDateTime dtDate)
{
	int lMonth;
	lMonth = dtDate.GetMonth();
	if (lMonth >= 1 && lMonth <= 3)
	{
		return 1;
	}
	else if (lMonth >= 4 && lMonth <= 6)
	{
		return 2;
	}
	else if (lMonth >= 7 && lMonth <= 9)
	{
		return 3;
	}
	else if (lMonth >= 10 && lMonth <= 12)
	{
		return 4;
	}
	return -1;
}

LONG clsMath::Week(COleDateTime dtDate)
{
	if (mp_oControl->GetConfiguration()->GetISO8601Weeks() == FALSE)
	{
		return GetWeekOfYear(dtDate, mp_oControl->GetConfiguration()->GetCalendarWeekRule(), mp_oControl->GetConfiguration()->GetFirstDayOfWeek());
	}
	else
	{
		LONG yDay = dtDate.GetDayOfWeek() - 1;
		COleDateTime dtBuff = dtDate;
		if (yDay >= WD_MONDAY && yDay <= WD_WEDNESDAY)
		{
			COleDateTimeSpan oSpan(3, 0, 0, 0);
			dtBuff = dtDate + oSpan;
		}
		return GetWeekOfYear(dtBuff, CWR_FIRSTFOURDAYWEEK, WD_MONDAY);
	}
}

LONG clsMath::NumericDayOfWeek(COleDateTime dtDate)
{
	int lWeekStartDay = mp_oControl->GetConfiguration()->GetFirstDayOfWeek();
	int lWeekDay = dtDate.GetDayOfWeek() - 1;
	if ((lWeekDay - lWeekStartDay) >= 0)
	{
		return lWeekDay - lWeekStartDay + 1;
	}
	else
	{
		return lWeekDay + 8 - lWeekStartDay;
	}
}


//////////////////////////////////

BOOL clsMath::mp_DateBlockVisible(COleDateTime dtIntStart, COleDateTime dtIntEnd, COleDateTime dtBaseDate, LONG yInterval, LONG lFactor)
{
	COleDateTime dtStart;
	COleDateTime dtEnd;
	if (lFactor > 0)
	{
		dtStart = dtBaseDate;
		dtEnd = DateTimeAdd(yInterval, lFactor, dtBaseDate);
	}
	else
	{
		dtStart = DateTimeAdd(yInterval, lFactor, dtBaseDate);
		dtEnd = dtBaseDate;
	}
	if (dtStart > dtIntStart && dtStart < dtIntEnd)
	{
		return TRUE;
	}
	if (dtEnd > dtIntStart && dtEnd < dtIntEnd)
	{
		return TRUE;
	}
	if (dtStart <= dtIntStart && dtEnd >= dtIntEnd)
	{
		return TRUE;
	}
	return FALSE;
}

void clsMath::mp_GenerateTimeBlocks(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtIntStart, COleDateTime dtIntEnd)
{
	int i = 0;
	clsTimeBlock* oTimeBlock;
	for (i = 1; i <= mp_oControl->GetTimeBlocks()->GetCount(); i++)
	{
		oTimeBlock = (clsTimeBlock*)mp_oControl->GetTimeBlocks()->mp_oCollection->m_oReturnArrayElement(i);
		if (mp_bIsValidTimeBlock(oTimeBlock) == TRUE)
		{
			if (oTimeBlock->GetTimeBlockType() == TBT_SINGLE_OCURRENCE)
			{
				//If mp_DateBlockVisible(dtIntStart, dtIntEnd, oTimeBlock.StartDate, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) = True Then
				mp_AddTimeBlock(aTimeBlocks, oTimeBlock->GetStartDate(), oTimeBlock);
				//End If
			}
			else if (oTimeBlock->GetTimeBlockType() == TBT_RECURRING)
			{
				COleDateTime dtCurrent;
				COleDateTime dtBase;
				COleDateTime dtTimeLineStart;
				COleDateTime dtTimeLineEnd;
				switch (oTimeBlock->GetRecurringType())
				{
				case RCT_DAY:
					dtTimeLineStart = dtIntStart;
					dtTimeLineEnd = dtIntEnd;
					dtTimeLineStart = DateTimeAdd(IL_DAY, -7, dtTimeLineStart);
					dtTimeLineEnd = DateTimeAdd(IL_DAY, 7, dtTimeLineEnd);
					dtCurrent.SetDateTime(dtTimeLineStart.GetYear(), dtTimeLineStart.GetMonth(), dtTimeLineStart.GetDay(), 0, 0, 0);
					while (dtCurrent < dtTimeLineEnd)
					{
						dtBase.SetDateTime(dtCurrent.GetYear(), dtCurrent.GetMonth(), dtCurrent.GetDay(), oTimeBlock->GetBaseDate().GetHour(), oTimeBlock->GetBaseDate().GetMinute(), oTimeBlock->GetBaseDate().GetSecond());
						dtCurrent = DateTimeAdd(IL_DAY, 1, dtCurrent);
						if (mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor()))
						{
							mp_AddTimeBlock(aTimeBlocks, dtBase, oTimeBlock);
						}
					}
					break;
				case RCT_WEEK:
					dtTimeLineStart = dtIntStart;
					dtTimeLineEnd = dtIntEnd;
					dtTimeLineStart = DateTimeAdd(IL_DAY, -7, dtTimeLineStart);
					dtTimeLineEnd = DateTimeAdd(IL_DAY, 7, dtTimeLineEnd);
					dtCurrent.SetDateTime(dtTimeLineStart.GetYear(), dtTimeLineStart.GetMonth(), dtTimeLineStart.GetDay(), 0, 0, 0);
					while (dtCurrent < dtTimeLineEnd)
					{
						dtBase.SetDateTime(dtCurrent.GetYear(), dtCurrent.GetMonth(), dtCurrent.GetDay(), oTimeBlock->GetBaseDate().GetHour(), oTimeBlock->GetBaseDate().GetMinute(), oTimeBlock->GetBaseDate().GetSecond());
						dtCurrent = DateTimeAdd(IL_DAY, 1, dtCurrent);
						if (oTimeBlock->GetBaseWeekDay() == dtBase.GetDayOfWeek() - 1)
						{
							if (mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor()))
							{
								mp_AddTimeBlock(aTimeBlocks, dtBase, oTimeBlock);
							}
						}
					}
					break;
				case RCT_MONTH:
					dtTimeLineStart = dtIntStart;
					dtTimeLineEnd = dtIntEnd;
					dtTimeLineStart = DateTimeAdd(IL_MONTH, -1, dtTimeLineStart);
					dtTimeLineEnd = DateTimeAdd(IL_MONTH, 1, dtTimeLineEnd);
					dtCurrent.SetDateTime(dtTimeLineStart.GetYear(), dtTimeLineStart.GetMonth(), dtTimeLineStart.GetDay(), 0, 0, 0);
					while (dtCurrent < dtTimeLineEnd)
					{
						if (oTimeBlock->GetBaseDate().GetDay() == dtCurrent.GetDay())
						{
							dtBase.SetDateTime(dtCurrent.GetYear(), dtCurrent.GetMonth(), dtCurrent.GetDay(), oTimeBlock->GetBaseDate().GetHour(), oTimeBlock->GetBaseDate().GetMinute(), oTimeBlock->GetBaseDate().GetSecond());
							dtCurrent = DateTimeAdd(IL_DAY, 1, dtCurrent);
							if (mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor()))
							{
								mp_AddTimeBlock(aTimeBlocks, dtBase, oTimeBlock);
							}
						}
						else
						{
							dtCurrent = DateTimeAdd(IL_DAY, 1, dtCurrent);
						}
					}
					break;
				case RCT_YEAR:
					dtTimeLineStart = dtIntStart;
					dtTimeLineEnd = dtIntEnd;
					dtTimeLineStart = DateTimeAdd(IL_YEAR, -1, dtTimeLineStart);
					dtTimeLineEnd = DateTimeAdd(IL_YEAR, 1, dtTimeLineEnd);
					dtCurrent.SetDateTime(dtTimeLineStart.GetYear(), dtTimeLineStart.GetMonth(), dtTimeLineStart.GetDay(), 0, 0, 0);
					while (dtCurrent < dtTimeLineEnd)
					{
						if (oTimeBlock->GetBaseDate().GetMonth() == dtCurrent.GetMonth())
						{
							if (oTimeBlock->GetBaseDate().GetDay() == dtCurrent.GetDay())
							{
								dtBase.SetDateTime(dtCurrent.GetYear(), dtCurrent.GetMonth(), dtCurrent.GetDay(), oTimeBlock->GetBaseDate().GetHour(), oTimeBlock->GetBaseDate().GetMinute(), oTimeBlock->GetBaseDate().GetSecond());
								dtCurrent = DateTimeAdd(IL_DAY, 1, dtCurrent);
								if (mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor()))
								{
									mp_AddTimeBlock(aTimeBlocks, dtBase, oTimeBlock);
								}
							}
							else
							{
								dtCurrent = DateTimeAdd(IL_DAY, 1, dtCurrent);
							}
						}
						else
						{
							dtCurrent = DateTimeAdd(IL_MONTH, 1, dtCurrent);
						}
					}
					break;
				}
			}
		}
	}
	if (aTimeBlocks.size() > 0)
	{
		mp_QuickSortTB(aTimeBlocks, 0, aTimeBlocks.size() - 1);
		mp_MergeTB(aTimeBlocks);
	}
}

void clsMath::mp_AddTimeBlock(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtBase, clsTimeBlock* oTimeBlock)
{
	S_TIMEBLOCK oTB;
	if (oTimeBlock->GetDurationFactor() > 0)
	{
		oTB.dtStart = dtBase;
		oTB.dtEnd = DateTimeAdd(oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor(), dtBase);
	}
	else
	{
		oTB.dtEnd = dtBase;
		oTB.dtStart = DateTimeAdd(oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor(), dtBase);
	}
	aTimeBlocks.push_back(oTB);
}

BOOL clsMath::mp_bIsValidTimeBlock(clsTimeBlock* oTimeBlock)
{
	if (oTimeBlock->GetDurationFactor() == 0)
	{
		return FALSE;
	}
	if (oTimeBlock->GetNonWorking() == TRUE)
	{
		return TRUE;
	}
	return FALSE;
}

void clsMath::mp_MergeTB(std::vector <S_TIMEBLOCK> &aTimeBlocks)
{
	unsigned i = 0;
	int lStart = 0;
	bool bFinished = FALSE;
	while (bFinished == FALSE)
	{
		for (i = lStart; i <= aTimeBlocks.size() - 2; i++)
		{
			S_TIMEBLOCK oTB1;
			S_TIMEBLOCK oTB2;
			int lTB1 = 0;
			int lTB2 = 0;
			lTB1 = i;
			lTB2 = i + 1;
			oTB1 = aTimeBlocks[lTB1];
			oTB2 = aTimeBlocks[lTB2];
			if ((oTB2.dtStart > oTB1.dtStart) && (oTB2.dtEnd < oTB1.dtEnd))
			{
				//Case 1
				//xxxxxxxxxxxxxxxxx           TB1
				//   xxxxxxxxxxxx             TB2
				aTimeBlocks.erase(aTimeBlocks.begin() + lTB2);
				bFinished = FALSE;
				if (i >= 1)
					lStart = i - 1;
				break; // TODO: might not be correct. Was : Exit For
			}
			else if ((oTB2.dtStart == oTB1.dtStart) && (oTB2.dtEnd == oTB1.dtEnd))
			{
				//Case 2
				//xxxxxxxxxxxxxxxxx           TB1
				//xxxxxxxxxxxxxxxxx           TB2
				aTimeBlocks.erase(aTimeBlocks.begin() + lTB2);
				bFinished = FALSE;
				if (i >= 1)
					lStart = i - 1;
				break; // TODO: might not be correct. Was : Exit For
			}
			else if ((oTB2.dtStart == oTB1.dtStart) && (oTB2.dtEnd > oTB1.dtEnd))
			{
				//Case 3
				//xxxxxxxxxxxx                TB1
				//xxxxxxxxxxxxxxxxx           TB2
				oTB1.dtEnd = oTB2.dtEnd;
				aTimeBlocks[lTB1] = oTB1;
				aTimeBlocks.erase(aTimeBlocks.begin() + lTB2);
				bFinished = FALSE;
				if (i >= 1)
					lStart = i - 1;
				break; // TODO: might not be correct. Was : Exit For
			}
			else if ((oTB2.dtStart == oTB1.dtStart) && (oTB2.dtEnd < oTB1.dtEnd))
			{
				//Case 4
				//xxxxxxxxxxxxxxxxx           TB1
				//xxxxxxxxxxx                 TB2
				aTimeBlocks.erase(aTimeBlocks.begin() + lTB2);
				bFinished = FALSE;
				if (i >= 1)
					lStart = i - 1;
				break; // TODO: might not be correct. Was : Exit For
			}
			else if ((oTB2.dtStart > oTB1.dtStart) && (oTB2.dtStart < oTB1.dtEnd))
			{
				//Case 5
				//xxxxxxxxxxxxxxxxx            TB1
				//              xxxxxxxxxxxxxx TB2
				oTB1.dtEnd = oTB2.dtEnd;
				aTimeBlocks[lTB1] = oTB1;
				aTimeBlocks.erase(aTimeBlocks.begin() + lTB2);
				bFinished = FALSE;
				if (i >= 1)
					lStart = i - 1;
				break; // TODO: might not be correct. Was : Exit For
			}
			else if ((oTB1.dtEnd == oTB2.dtStart))
			{
				//Case 6
				//xxxxxxxxxxxxxx               TB1
				//              xxxxxxxxxxxxxx TB2
				oTB1.dtEnd = oTB2.dtEnd;
				aTimeBlocks[lTB1] = oTB1;
				aTimeBlocks.erase(aTimeBlocks.begin() + lTB2);
				bFinished = FALSE;
				if (i >= 1)
					lStart = i - 1;
				break; // TODO: might not be correct. Was : Exit For
			}
			bFinished = TRUE;
		}
	}
}

LONG clsMath::mp_CheckDuration(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtStartDate)
{
	unsigned i = 0;
	int lSeconds = 0;
	bool bInside = FALSE;
	for (i = 0; i <= aTimeBlocks.size() - 2; i++)
	{
		S_TIMEBLOCK oTB1;
		S_TIMEBLOCK oTB2;
		oTB1 = aTimeBlocks[i];
		oTB2 = aTimeBlocks[i + 1];
		int lSecondDiff = 0;
		if (dtStartDate >= oTB1.dtEnd && dtStartDate < oTB2.dtStart)
		{
			lSecondDiff = DateTimeDiff(IL_SECOND, dtStartDate, oTB2.dtStart);
			bInside = TRUE;
		}
		else if (bInside == TRUE)
		{
			lSecondDiff = DateTimeDiff(IL_SECOND, oTB1.dtEnd, oTB2.dtStart);
		}
		if (bInside == TRUE && lSecondDiff <= 0)
		{
			mp_oControl->mp_ErrorReport(CHECK_DURATION_ERROR, _T("Inconsistent State in Check Duration"), _T("clsMath.mp_CheckDuration"));
			return -1;
		}
		lSeconds = lSeconds + lSecondDiff;
	}
	return lSeconds;
}

BOOL clsMath::mp_GetStartDate(std::vector <S_TIMEBLOCK> &aTimeBlocks, BOOL &bStartDateVerified, COleDateTime dtStartDate)
{
	if (bStartDateVerified == TRUE)
	{
		return TRUE;
	}
	unsigned i = 0;
	for (i = 0; i <= aTimeBlocks.size() - 1; i++)
	{
		S_TIMEBLOCK oTB1;
		oTB1 = aTimeBlocks[i];
		if (dtStartDate >= oTB1.dtStart && dtStartDate < oTB1.dtEnd)
		{
			dtStartDate = oTB1.dtEnd;
			bStartDateVerified = TRUE;
			return FALSE;
		}
	}
	bStartDateVerified = TRUE;
	return TRUE;
}

void clsMath::mp_GetEndDate(std::vector <S_TIMEBLOCK> &aTimeBlocks, LONG lDurationInSeconds, COleDateTime dtStartDate, COleDateTime dtEndDate)
{
	unsigned i = 0;
	bool bInside = FALSE;
	for (i = 0; i <= aTimeBlocks.size() - 2; i++)
	{
		S_TIMEBLOCK oTB1;
		S_TIMEBLOCK oTB2;
		oTB1 = aTimeBlocks[i];
		oTB2 = aTimeBlocks[i + 1];
		int lSecondDiff = 0;
		if (dtStartDate >= oTB1.dtEnd && dtStartDate < oTB2.dtStart && bInside == FALSE)
		{
			lSecondDiff = DateTimeDiff(IL_SECOND, dtStartDate, oTB2.dtStart);
			bInside = TRUE;
			if (lDurationInSeconds <= lSecondDiff)
            {
                dtEndDate = DateTimeAdd(IL_SECOND, lDurationInSeconds, dtStartDate);
                return;
            }
		}
		else if (bInside == TRUE)
		{
			lSecondDiff = DateTimeDiff(IL_SECOND, oTB1.dtEnd, oTB2.dtStart);
		}
		if (bInside == TRUE)
		{
			lDurationInSeconds = lDurationInSeconds - lSecondDiff;
			if (lDurationInSeconds <= 0)
			{
				lDurationInSeconds = lDurationInSeconds + lSecondDiff;
				dtEndDate = DateTimeAdd(IL_SECOND, lDurationInSeconds, oTB1.dtEnd);
				return;
			}
		}
	}
}

void clsMath::mp_ValidateStartDate(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtStartDate)
{
	unsigned i = 0;
	for (i = 0; i <= aTimeBlocks.size() - 1; i++)
	{
		S_TIMEBLOCK oTB1;
		oTB1 = aTimeBlocks[i];
		if (dtStartDate >= oTB1.dtStart && dtStartDate < oTB1.dtEnd)
		{
			dtStartDate = oTB1.dtEnd;
			return;
		}
	}
}

void clsMath::mp_ValidateEndDate(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtEndDate)
{
	unsigned i = 0;
	for (i = 0; i <= aTimeBlocks.size() - 1; i++)
	{
		S_TIMEBLOCK oTB1;
		oTB1 = aTimeBlocks[i];
		if (dtEndDate >= oTB1.dtStart && dtEndDate < oTB1.dtEnd)
		{
			dtEndDate = oTB1.dtStart;
			return;
		}
	}
}

void clsMath::mp_StandarizeInterval(COleDateTime dtIntStart, COleDateTime dtIntEnd)
{
	if (dtIntStart < dtIntEnd)
	{
		return;
	}
	COleDateTime dtIntBuff;
	dtIntBuff = dtIntStart;
	dtIntStart = dtIntEnd;
	dtIntEnd = dtIntBuff;
}

COleDateTime clsMath::GetEndDate(COleDateTime StartDate, LONG DurationInterval, LONG DurationFactor)
{
	COleDateTime EndDate;
	COleDateTime dtIntStart;
	COleDateTime dtIntEnd;
	BOOL bFinished = FALSE;
	BOOL bStartDateVerified = FALSE;
	int lIntFactor = 2;
	std::vector <S_TIMEBLOCK> aTimeBlocks;
	int lDurationInSeconds = 0;
	int lCheckDuration = 0;
	int lPass = 0;
	int i = 0;
	clsTimeBlock* oTimeBlock;
	int lSecondsInDay = 86400;
	int lSecondsInWeek = 604800;
	int lSecondsInMonth = 2419200;
	int lSecondsInYear = 31449600;
	int lEstimatedDuration = 0;
	if (!(DurationInterval == IL_SECOND || DurationInterval == IL_MINUTE || DurationInterval == IL_HOUR || DurationInterval == IL_DAY))
	{
		mp_oControl->mp_ErrorReport(INVALID_DURATION_INTERVAL, _T("Interval is invalid for a duration"), _T("clsMath.GetEndDate"));
		return EndDate;
	}
	lDurationInSeconds = mp_oControl->GetMathLib()->mp_GetSeconds(DurationInterval, DurationFactor);
	lEstimatedDuration = lDurationInSeconds;
	for (i = 1; i <= mp_oControl->GetTimeBlocks()->GetCount(); i++)
	{
		oTimeBlock = (clsTimeBlock*)mp_oControl->GetTimeBlocks()->mp_oCollection->m_oReturnArrayElement(i);
		if (mp_oControl->GetMathLib()->mp_bIsValidTimeBlock(oTimeBlock))
		{
			if (oTimeBlock->GetTimeBlockType() == TBT_RECURRING)
			{
				switch (oTimeBlock->GetRecurringType())
				{
				case RCT_DAY:
					lSecondsInDay = lSecondsInDay - mp_oControl->GetMathLib()->mp_GetSeconds(oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor());
					break;
				case RCT_WEEK:
					lSecondsInWeek = lSecondsInWeek - mp_oControl->GetMathLib()->mp_GetSeconds(oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor());
					break;
				case RCT_MONTH:
					lSecondsInMonth = lSecondsInMonth - mp_oControl->GetMathLib()->mp_GetSeconds(oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor());
					break;
				case RCT_YEAR:
					lSecondsInYear = lSecondsInYear - mp_oControl->GetMathLib()->mp_GetSeconds(oTimeBlock->GetDurationInterval(), oTimeBlock->GetDurationFactor());
					break;
				}
			}
		}
	}
	if (lDurationInSeconds > 31449600)
	{
		lEstimatedDuration = lEstimatedDuration + (CInt32(ceil((double)lDurationInSeconds / (double)lSecondsInYear)) * 31449600);
	}
	if (lDurationInSeconds > 2419200)
	{
		lEstimatedDuration = lEstimatedDuration + (CInt32(ceil((double)lDurationInSeconds / (double)lSecondsInMonth)) * 2419200);
	}
	if (lDurationInSeconds > 604800)
	{
		lEstimatedDuration = lEstimatedDuration + (CInt32(ceil((double)lDurationInSeconds / (double)lSecondsInWeek)) * 604800);
	}
	if (lDurationInSeconds > 86400)
	{
		lEstimatedDuration = lEstimatedDuration + (CInt32(ceil((double)lDurationInSeconds / (double)lSecondsInDay)) * 86400);
	}
	while (bFinished == FALSE)
	{
		lPass = lPass + 1;
		dtIntStart = StartDate;
		dtIntEnd = mp_oControl->GetMathLib()->DateTimeAdd(IL_SECOND, lEstimatedDuration * lIntFactor, StartDate);
		mp_oControl->GetMathLib()->mp_StandarizeInterval(dtIntStart, dtIntEnd);
		if (mp_oControl->GetTimeBlocks()->GetIntervalType() == TBIT_AUTOMATIC)
		{
			aTimeBlocks.clear();
			mp_oControl->GetMathLib()->mp_GenerateTimeBlocks(aTimeBlocks, dtIntStart, dtIntEnd);
		}
		else
		{
			aTimeBlocks = mp_oControl->GetTimeBlocks()->mp_aTimeBlocks;
		}
		if (aTimeBlocks.size() == 0)
		{
			EndDate = mp_oControl->GetMathLib()->DateTimeAdd(DurationInterval, DurationFactor, StartDate);
			return EndDate;
		}
		else
		{
			if (mp_oControl->GetMathLib()->mp_GetStartDate(aTimeBlocks, bStartDateVerified, StartDate) == TRUE)
			{
				lCheckDuration = mp_oControl->GetMathLib()->mp_CheckDuration(aTimeBlocks, StartDate);
				if (lCheckDuration < lDurationInSeconds)
				{
					if (lCheckDuration == 0)
					{
						lCheckDuration = lDurationInSeconds;
					}
					lIntFactor = (CInt32(ceil((double)lDurationInSeconds / (double)lCheckDuration)) * 2) + lIntFactor;
				}
				else
				{
					mp_oControl->GetMathLib()->mp_GetEndDate(aTimeBlocks, lDurationInSeconds, StartDate, EndDate);
					bFinished = TRUE;
				}
			}
		}
	}
	return EndDate;
}



void clsMath::mp_GetTimeBlocks(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtStartDate, COleDateTime dtEndDate)
{
	if (mp_oControl->GetTimeBlocks()->GetIntervalType() == TBIT_AUTOMATIC)
	{
		aTimeBlocks.clear();
		mp_GenerateTimeBlocks(aTimeBlocks, dtStartDate, dtEndDate);
	}
	else
	{
		aTimeBlocks = mp_oControl->GetTimeBlocks()->mp_aTimeBlocks;
	}
}
	
LONG clsMath::CalculateDuration(COleDateTime dtStartDate, COleDateTime dtEndDate, LONG DurationInterval)
{
	std::vector <S_TIMEBLOCK> aTimeBlocks;
	int lDurationInSeconds = 0;
	int lReturn = 0;
	mp_StandarizeInterval(dtStartDate, dtEndDate);
	if (!(DurationInterval == IL_SECOND || DurationInterval == IL_MINUTE || DurationInterval == IL_HOUR || DurationInterval == IL_DAY))
	{
		mp_oControl->mp_ErrorReport(INVALID_DURATION_INTERVAL, _T("Interval is invalid for a duration"), _T("clsMath.CalculateDuration"));
		return -1;
	}
	mp_GetTimeBlocks(aTimeBlocks, dtStartDate, dtEndDate);
	if (aTimeBlocks.size() == 0)
	{
		return DateTimeDiff(DurationInterval, dtStartDate, dtEndDate);
	}
	else
	{
		mp_ValidateStartDate(aTimeBlocks, dtStartDate);
		mp_ValidateEndDate(aTimeBlocks, dtEndDate);
		lDurationInSeconds = mp_GetDuration(aTimeBlocks, dtStartDate, dtEndDate);
		switch (DurationInterval)
		{
		case IL_SECOND:
			lReturn = lDurationInSeconds;
			break;
		case IL_MINUTE:
			lReturn = CInt32(floor((double)lDurationInSeconds / 60));
			break;
		case IL_HOUR:
			lReturn = CInt32(floor((double)lDurationInSeconds / 3600));
			break;
		case IL_DAY:
			lReturn = CInt32(floor((double)lDurationInSeconds / 86400));
			break;
		}
		return lReturn;
	}
}

LONG clsMath::mp_GetDuration(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtStartDate, COleDateTime dtEndDate)
{
	unsigned i = 0;
	bool bInside = FALSE;
	int lReturn = 0;
	for (i = 0; i <= aTimeBlocks.size() - 2; i++)
	{
		S_TIMEBLOCK oTB1;
		S_TIMEBLOCK oTB2;
		oTB1 = aTimeBlocks[i];
		oTB2 = aTimeBlocks[i + 1];
		if (dtStartDate >= oTB1.dtEnd && dtStartDate <= oTB2.dtStart && dtEndDate >= oTB1.dtEnd && dtEndDate <= oTB2.dtStart)
		{
			lReturn = DateTimeDiff(IL_SECOND, dtStartDate, dtEndDate);
		}
		else if (dtStartDate >= oTB1.dtEnd && dtStartDate <= oTB2.dtStart)
		{
			lReturn = lReturn + DateTimeDiff(IL_SECOND, dtStartDate, oTB2.dtStart);
			bInside = TRUE;
		}
		else if (dtEndDate >= oTB1.dtEnd && dtEndDate <= oTB2.dtStart && bInside == TRUE)
		{
			lReturn = lReturn + DateTimeDiff(IL_SECOND, oTB1.dtEnd, dtEndDate);
			break; // TODO: might not be correct. Was : Exit For
		}
		else if (bInside == TRUE)
		{
			lReturn = lReturn + DateTimeDiff(IL_SECOND, oTB1.dtEnd, oTB2.dtStart);
		}
	}
	return lReturn;
}

void clsMath::mp_DumpTB(S_TIMEBLOCK oTB)
{
	//TRACE("StartDate: " + oTB.dtStart.ToString("yyyy/MM/dd HH:mm:ss") + _T("\r\n"));
	//TRACE("EndDate: " + oTB.dtEnd.ToString("yyyy/MM/dd HH:mm:ss") + _T("\r\n"));
}

void clsMath::mp_DumpTimeBlocks(std::vector <S_TIMEBLOCK> &aTimeBlocks, CString sCaption)
{
	TRACE(sCaption + _T(" *************************** Dumping TimeBlocks:\r\n"));
	unsigned i = 0;
	for (i = 0; i <= aTimeBlocks.size() - 1; i++)
	{
		S_TIMEBLOCK oTB;
		oTB = aTimeBlocks[i];
		//TRACE(CStr(i) + ":\r\n");
		//TRACE("StartDate: " + oTB.dtStart.ToString("yyyy/MM/dd HH:mm:ss") + _T("\r\n"));
		//TRACE("EndDate: " + oTB.dtEnd.ToString("yyyy/MM/dd HH:mm:ss") + _T("\r\n"));
	}
}

LONG clsMath::mp_GetSeconds(LONG yInterval, LONG lFactor)
{
	if (lFactor < 0)
	{
		lFactor = lFactor * -1;
	}
	switch (yInterval)
	{
	case IL_SECOND:
		return lFactor;
	case IL_MINUTE:
		return lFactor * 60;
	case IL_HOUR:
		return lFactor * 3600;
	case IL_DAY:
		return lFactor * 86400;
	}
	return -1;
}

void clsMath::mp_QuickSortTB(std::vector <S_TIMEBLOCK> &aTimeBlocks, int StartIndex, int EndIndex)
{
	// StartIndex = subscript of beginning of array
	// EndIndex = subscript of end of array
      
	int MiddleIndex;
	if (StartIndex < EndIndex)
    {
          MiddleIndex = mp_QSPartitionTB(aTimeBlocks, StartIndex, EndIndex);
          mp_QuickSortTB(aTimeBlocks, StartIndex, MiddleIndex);   // sort first section
          mp_QuickSortTB(aTimeBlocks, MiddleIndex + 1, EndIndex);    // sort second section
     }
     return;
}

int clsMath::mp_QSPartitionTB(std::vector <S_TIMEBLOCK> &aTimeBlocks, int StartIndex, int EndIndex)
{
	COleDateTime x = aTimeBlocks[StartIndex].dtStart;
	int i = StartIndex - 1;
	int j = EndIndex + 1;
	S_TIMEBLOCK temp;
	do
	{
		do   
		{
			j--;
		}while (x < aTimeBlocks[j].dtStart); //Change to > for descending

		do  
		{
			i++;
		}while (x > aTimeBlocks[i].dtStart); //Change to < for descending

		if (i < j)
		{ 
			 temp = aTimeBlocks[i];    
			 aTimeBlocks[i] = aTimeBlocks[j];
			 aTimeBlocks[j] = temp;
		}

     }while (i < j);    
     return j;           // returns middle subscript  
}

CString clsMath::GetTierIndex(LONG Number)
{
    CString sNumber;
    LONG lLastDigit = 0;
    sNumber = CStr(Number);
    lLastDigit = CInt32(sNumber.Right(1));
    if (lLastDigit == 0)
    {
        return _T("10");
    }
    else
    {
        return CStr(lLastDigit);
    }
}

int clsMath::GetWeekOfYear(COleDateTime time, int rule, int firstDayOfWeek)
{
    int lReturn = 0;
    COleDateTime dtBase(time.GetYear(), 1, 1, 0, 0, 0);
    COleDateTime dtBuff;
    int lDay = 0;
    int lWeek = 0;
    switch (rule) 
    {
    case CWR_FIRSTDAY:
        for (lDay = 0; lDay <= 6; lDay++) 
        {
            COleDateTimeSpan oSpan(lDay, 0, 0, 0);
            dtBuff = dtBase + oSpan;
            if ((dtBuff.GetDayOfWeek() - 1) == firstDayOfWeek) 
            {
                break;
            }
        }
        if ((dtBase.GetDayOfWeek() - 1) == firstDayOfWeek) 
        {
            lReturn = (int)floor(((double)time.GetDayOfYear() - (double)dtBuff.GetDayOfYear()) / 7.0) + 1;
        } 
        else 
        {
            lReturn = (int)floor(((double)time.GetDayOfYear() - (double)dtBuff.GetDayOfYear()) / 7.0) + 2;
        }
        break;
    case CWR_FIRSTFULLWEEK:
        for (lDay = 0; lDay <= 6; lDay++) 
        {
            COleDateTimeSpan oSpan(lDay, 0, 0, 0);
            dtBuff = dtBase + oSpan;
            if ((dtBuff.GetDayOfWeek() - 1) == firstDayOfWeek) 
            {
                break;
            }
        }
        lWeek = (int)floor(((double)time.GetDayOfYear() - (double)dtBuff.GetDayOfYear()) / 7) + 1;
        if (lWeek == 0) 
        {
            COleDateTime dtParam(time.GetYear() - 1, 12, 31, 0, 0, 0);
            lReturn = GetWeekOfYear(dtParam, CWR_FIRSTFULLWEEK, firstDayOfWeek);
        } else 
        {
            lReturn = lWeek;
        }
        break;
    case CWR_FIRSTFOURDAYWEEK:
        for (lDay = 0; lDay <= 12; lDay++) 
        {
            COleDateTimeSpan oSpan(lDay, 0, 0, 0);
            dtBuff = dtBase + oSpan;
            if ((dtBuff.GetDayOfWeek() - 1) == firstDayOfWeek && dtBuff.GetDayOfYear() >= 5) 
            {
                break;
            }
        }
        lWeek = (int)floor(((double)time.GetDayOfYear() - (double)dtBuff.GetDayOfYear()) / 7.0) + 2;
        if (lWeek == 0) 
        {
            COleDateTime dtParam(time.GetYear() - 1, 12, 31, 0, 0, 0);
            lReturn = GetWeekOfYear(dtParam, CWR_FIRSTFOURDAYWEEK, firstDayOfWeek);
        } 
        else 
        {
            lReturn = lWeek;
        }
        break;
    }
    return lReturn;
}
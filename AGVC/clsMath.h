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


class clsMath : public CCmdTarget
{
	DECLARE_DYNCREATE(clsMath)
	clsMath();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsMath(CActiveGanttVCCtl* oControl);
	virtual ~clsMath();
	virtual void OnFinalRelease();

// Member Variables
public:
	COleDateTime mp_dtDate;

//Internal Properties (Extensions to ODL)
public:

//Internal Properties
public:

//Internal Methods
public:
	void Finalize(void);
	COleDateTime DateTimeAdd(LONG Interval,LONG Number, COleDateTime dtDate);
	LONG DateTimeDiff(LONG Interval, COleDateTime dtDate1, COleDateTime dtDate2);
	COleDateTime NewDateTime(LONG Year,LONG Month,LONG Day,LONG Hour,LONG Minute,LONG Second);
	LONG GetXCoordinateFromDate(COleDateTime dtCoordinate);
	COleDateTime GetDateFromXCoordinate(LONG v_lXCoordinate);
	LONG GetRowIndexByPosition(LONG Y);
	LONG GetCellIndexByPosition(LONG X);
	LONG GetColumnIndexByPosition(LONG X,LONG Y);
	LONG GetTaskIndexByPosition(LONG X,LONG Y);
	LONG GetPredecessorIndexByPosition(LONG X,LONG Y);
	LONG GetPercentageIndexByPosition(LONG X,LONG Y);
	LONG GetNodeIndexByCheckBoxPosition(LONG X,LONG Y);
	LONG GetNodeIndexBySignPosition(LONG X,LONG Y);
	BOOL InCurrentLayer(CString sLayer);
	BOOL DetectConflict(COleDateTime StartDate, COleDateTime EndDate,CString RowKey,LONG ExcludeIndex,CString LayerIndex);
	FLOAT PercentageComplete(LONG X1,LONG X2,LONG X);
	COleDateTime RoundDate(LONG Interval,LONG Number, COleDateTime dtDate);
	LONG RoundDouble(DOUBLE dParam);
	LONG Round(LONG v_lNumberToRound,LONG v_lRoundTo);
	LONG lAbs(LONG Number);
	BYTE Rnd();
	LONG Quarter(COleDateTime dtDate);
	LONG Week(COleDateTime dtDate);
	LONG NumericDayOfWeek(COleDateTime dtDate);

	BOOL mp_DateBlockVisible(COleDateTime dtIntStart, COleDateTime dtIntEnd, COleDateTime dtBaseDate, LONG yInterval, LONG lFactor);
	void mp_GenerateTimeBlocks(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtIntStart, COleDateTime dtIntEnd);
	void mp_AddTimeBlock(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtBase, clsTimeBlock* oTimeBlock);
	BOOL mp_bIsValidTimeBlock(clsTimeBlock* oTimeBlock);
	void mp_MergeTB(std::vector <S_TIMEBLOCK> &aTimeBlocks);
	LONG mp_CheckDuration(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtStartDate);
	BOOL mp_GetStartDate(std::vector <S_TIMEBLOCK> &aTimeBlocks, BOOL &bStartDateVerified, COleDateTime dtStartDate);
	void mp_GetEndDate(std::vector <S_TIMEBLOCK> &aTimeBlocks, LONG lDurationInSeconds, COleDateTime dtStartDate, COleDateTime dtEndDate);
	void mp_ValidateStartDate(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtStartDate);
	void mp_ValidateEndDate(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtEndDate);
	void mp_StandarizeInterval(COleDateTime dtIntStart, COleDateTime dtIntEnd);
	COleDateTime GetEndDate(COleDateTime StartDate, LONG DurationInterval, LONG DurationFactor);
	void mp_GetTimeBlocks(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtStartDate, COleDateTime dtEndDate);
	LONG CalculateDuration(COleDateTime dtStartDate, COleDateTime dtEndDate, LONG DurationInterval);
	LONG mp_GetDuration(std::vector <S_TIMEBLOCK> &aTimeBlocks, COleDateTime dtStartDate, COleDateTime dtEndDate);
	void mp_DumpTB(S_TIMEBLOCK oTB);
	void mp_DumpTimeBlocks(std::vector <S_TIMEBLOCK> &aTimeBlocks, CString sCaption);
	LONG mp_GetSeconds(LONG yInterval, LONG lFactor);

	void mp_QuickSortTB(std::vector <S_TIMEBLOCK> &aTimeBlocks, int StartIndex, int EndIndex);
	int mp_QSPartitionTB(std::vector <S_TIMEBLOCK> &aTimeBlocks, int StartIndex, int EndIndex);

	CString GetTierIndex(LONG Number);

	int clsMath::GetWeekOfYear(COleDateTime time, int rule, int firstDayOfWeek);



protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsMath)
	DECLARE_INTERFACE_MAP()


DATE odl_DateTimeAdd(LONG Interval,LONG Number, DATE dtDate)
{	
	mp_dtDate = DateTimeAdd(Interval,Number, dtDate);
	return mp_dtDate;
}

LONG odl_DateTimeDiff(LONG Interval,DATE dtDate1,DATE dtDate2)
{
	return DateTimeDiff(Interval, dtDate1, dtDate2);
}

LONG odl_GetXCoordinateFromDate(DATE dtCoordinate)
{
	return GetXCoordinateFromDate(dtCoordinate);
}

DATE odl_GetDateFromXCoordinate(LONG v_lXCoordinate)
{
	
	mp_dtDate = GetDateFromXCoordinate(v_lXCoordinate);
	return mp_dtDate;
}

LONG odl_GetRowIndexByPosition(LONG Y)
{
	
	return GetRowIndexByPosition(Y);
}

LONG odl_GetCellIndexByPosition(LONG X)
{
	
	return GetCellIndexByPosition(X);
}

LONG odl_GetColumnIndexByPosition(LONG X,LONG Y)
{
	
	return GetColumnIndexByPosition(X,Y);
}

LONG odl_GetTaskIndexByPosition(LONG X,LONG Y)
{
	
	return GetTaskIndexByPosition(X,Y);
}

LONG odl_GetPredecessorIndexByPosition(LONG X,LONG Y)
{
	
	return GetPredecessorIndexByPosition(X,Y);
}

LONG odl_GetPercentageIndexByPosition(LONG X,LONG Y)
{
	
	return GetPercentageIndexByPosition(X,Y);
}

LONG odl_GetNodeIndexByCheckBoxPosition(LONG X,LONG Y)
{
	
	return GetNodeIndexByCheckBoxPosition(X,Y);
}

LONG odl_GetNodeIndexBySignPosition(LONG X,LONG Y)
{
	
	return GetNodeIndexBySignPosition(X,Y);
}

BOOL odl_DetectConflict(DATE StartDate, DATE EndDate,LPCTSTR RowKey,LONG ExcludeIndex,LPCTSTR LayerIndex)
{
	return DetectConflict(StartDate, EndDate,RowKey,ExcludeIndex,LayerIndex);
}

DATE odl_RoundDate(LONG Interval,LONG Number,DATE dtDate)
{	
	return RoundDate(Interval,Number, dtDate);
}

LONG odl_RoundDouble(DOUBLE dParam)
{
	
	return RoundDouble(dParam);
}

DATE odl_GetEndDate(DATE StartDate, LONG DurationInterval, LONG DurationFactor)
{
	

	DATE odtReturn = GetEndDate(StartDate, DurationInterval, DurationFactor);
	return odtReturn;
}

LONG odl_CalculateDuration(DATE dtStartDate, DATE dtEndDate, LONG DurationInterval)
{	
	LONG lReturn;
	lReturn = CalculateDuration(dtStartDate, dtEndDate, DurationInterval);
	return lReturn;
}



};
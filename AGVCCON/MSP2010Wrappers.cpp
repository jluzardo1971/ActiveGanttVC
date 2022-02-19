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
#include "MSP2010Wrappers.h"

namespace MSP2010
{
IMPLEMENT_DYNCREATE(CMSP2010Ctl, CWnd)

void CDuration::SetYear(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CDuration::GetYear()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CDuration::SetMonth(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CDuration::GetMonth()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CDuration::SetDay(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CDuration::GetDay()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CDuration::SetHour(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CDuration::GetHour()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CDuration::SetMinute(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CDuration::GetMinute()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CDuration::SetSecond(LONG propval)
{
	SetProperty(0x9, VT_I4, propval);
}

LONG CDuration::GetSecond()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

BOOL CDuration::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x1, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

CString CDuration::ToString(void)
{
	CString oReturn;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

CString CDuration::FromString(LPCTSTR sString)
{
	CString oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, (void*)&oReturn, params, sString);
	return oReturn;
}

void CTime::SetHour(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CTime::GetHour()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CTime::SetMinute(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CTime::GetMinute()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CTime::SetSecond(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CTime::GetSecond()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

BOOL CTime::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x1, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

CString CTime::ToString(void)
{
	CString oReturn;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

CString CTime::FromString(LPCTSTR sString)
{
	CString oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, (void*)&oReturn, params, sString);
	return oReturn;
}

void CTimephasedData::SetyType(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CTimephasedData::GetyType()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CTimephasedData::SetlUID(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CTimephasedData::GetlUID()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CTimephasedData::SetdtStart(DATE propval)
{
	SetProperty(0x3, VT_DATE, propval);
}

DATE CTimephasedData::GetdtStart()
{
	DATE result;
	GetProperty(0x3, VT_DATE, (void*)&result);
	return result;
}

void CTimephasedData::SetdtFinish(DATE propval)
{
	SetProperty(0x4, VT_DATE, propval);
}

DATE CTimephasedData::GetdtFinish()
{
	DATE result;
	GetProperty(0x4, VT_DATE, (void*)&result);
	return result;
}

void CTimephasedData::SetyUnit(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CTimephasedData::GetyUnit()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CTimephasedData::SetsValue(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CTimephasedData::GetsValue()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CTimephasedData::SetKey(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CTimephasedData::GetKey()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

CString CTimephasedData::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTimephasedData::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTimephasedData::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTimephasedData::Initialize(void)
{
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CTimephasedData_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CTimephasedData CTimephasedData_C::Item(LPCTSTR Index)
{
	CTimephasedData oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CTimephasedData CTimephasedData_C::Add(void)
{
	CTimephasedData oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CTimephasedData_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTimephasedData_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CTimephasedData_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTimephasedData_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CTimephasedData_C CAssignmentBaseline::GetoTimephasedData_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1, VT_DISPATCH, (void*)&pDispatch);
	return CTimephasedData_C(pDispatch);
}

void CAssignmentBaseline::SetsNumber(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CAssignmentBaseline::GetsNumber()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CAssignmentBaseline::SetsStart(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CAssignmentBaseline::GetsStart()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CAssignmentBaseline::SetsFinish(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CAssignmentBaseline::GetsFinish()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

CDuration CAssignmentBaseline::GetoWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignmentBaseline::SetsCost(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CAssignmentBaseline::GetsCost()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CAssignmentBaseline::SetfBCWS(FLOAT propval)
{
	SetProperty(0x7, VT_R4, propval);
}

FLOAT CAssignmentBaseline::GetfBCWS()
{
	FLOAT result;
	GetProperty(0x7, VT_R4, (void*)&result);
	return result;
}

void CAssignmentBaseline::SetfBCWP(FLOAT propval)
{
	SetProperty(0x8, VT_R4, propval);
}

FLOAT CAssignmentBaseline::GetfBCWP()
{
	FLOAT result;
	GetProperty(0x8, VT_R4, (void*)&result);
	return result;
}

void CAssignmentBaseline::SetKey(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CAssignmentBaseline::GetKey()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

CString CAssignmentBaseline::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentBaseline::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CAssignmentBaseline::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0xc, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentBaseline::Initialize(void)
{
	InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CAssignmentBaseline_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CAssignmentBaseline CAssignmentBaseline_C::Item(LPCTSTR Index)
{
	CAssignmentBaseline oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CAssignmentBaseline CAssignmentBaseline_C::Add(void)
{
	CAssignmentBaseline oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentBaseline_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CAssignmentBaseline_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CAssignmentBaseline_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentBaseline_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CAssignmentExtendedAttribute::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CAssignmentExtendedAttribute::GetsFieldID()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CAssignmentExtendedAttribute::SetsValue(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CAssignmentExtendedAttribute::GetsValue()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CAssignmentExtendedAttribute::SetlValueGUID(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CAssignmentExtendedAttribute::GetlValueGUID()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CAssignmentExtendedAttribute::SetyDurationFormat(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CAssignmentExtendedAttribute::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CAssignmentExtendedAttribute::SetKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CAssignmentExtendedAttribute::GetKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CString CAssignmentExtendedAttribute::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentExtendedAttribute::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CAssignmentExtendedAttribute::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentExtendedAttribute::Initialize(void)
{
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CAssignmentExtendedAttribute_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CAssignmentExtendedAttribute CAssignmentExtendedAttribute_C::Item(LPCTSTR Index)
{
	CAssignmentExtendedAttribute oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CAssignmentExtendedAttribute CAssignmentExtendedAttribute_C::Add(void)
{
	CAssignmentExtendedAttribute oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentExtendedAttribute_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CAssignmentExtendedAttribute_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CAssignmentExtendedAttribute_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentExtendedAttribute_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CAssignment::SetlUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CAssignment::GetlUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetlTaskUID(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CAssignment::GetlTaskUID()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetlResourceUID(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CAssignment::GetlResourceUID()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetlPercentWorkComplete(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CAssignment::GetlPercentWorkComplete()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetcActualCost(DOUBLE propval)
{
	SetProperty(0x5, VT_R8, propval);
}

DOUBLE CAssignment::GetcActualCost()
{
	DOUBLE result;
	GetProperty(0x5, VT_R8, (void*)&result);
	return result;
}

void CAssignment::SetdtActualFinish(DATE propval)
{
	SetProperty(0x6, VT_DATE, propval);
}

DATE CAssignment::GetdtActualFinish()
{
	DATE result;
	GetProperty(0x6, VT_DATE, (void*)&result);
	return result;
}

void CAssignment::SetcActualOvertimeCost(DOUBLE propval)
{
	SetProperty(0x7, VT_R8, propval);
}

DOUBLE CAssignment::GetcActualOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x7, VT_R8, (void*)&result);
	return result;
}

CDuration CAssignment::GetoActualOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x8, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetdtActualStart(DATE propval)
{
	SetProperty(0x9, VT_DATE, propval);
}

DATE CAssignment::GetdtActualStart()
{
	DATE result;
	GetProperty(0x9, VT_DATE, (void*)&result);
	return result;
}

CDuration CAssignment::GetoActualWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0xa, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetfACWP(FLOAT propval)
{
	SetProperty(0xb, VT_R4, propval);
}

FLOAT CAssignment::GetfACWP()
{
	FLOAT result;
	GetProperty(0xb, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetbConfirmed(BOOL propval)
{
	SetProperty(0xc, VT_BOOL, propval);
}

BOOL CAssignment::GetbConfirmed()
{
	BOOL result;
	GetProperty(0xc, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetcCost(DOUBLE propval)
{
	SetProperty(0xd, VT_R8, propval);
}

DOUBLE CAssignment::GetcCost()
{
	DOUBLE result;
	GetProperty(0xd, VT_R8, (void*)&result);
	return result;
}

void CAssignment::SetyCostRateTable(LONG propval)
{
	SetProperty(0xe, VT_I4, propval);
}

LONG CAssignment::GetyCostRateTable()
{
	LONG result;
	GetProperty(0xe, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetfCostVariance(FLOAT propval)
{
	SetProperty(0xf, VT_R4, propval);
}

FLOAT CAssignment::GetfCostVariance()
{
	FLOAT result;
	GetProperty(0xf, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetfCV(FLOAT propval)
{
	SetProperty(0x10, VT_R4, propval);
}

FLOAT CAssignment::GetfCV()
{
	FLOAT result;
	GetProperty(0x10, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetlDelay(LONG propval)
{
	SetProperty(0x11, VT_I4, propval);
}

LONG CAssignment::GetlDelay()
{
	LONG result;
	GetProperty(0x11, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetdtFinish(DATE propval)
{
	SetProperty(0x12, VT_DATE, propval);
}

DATE CAssignment::GetdtFinish()
{
	DATE result;
	GetProperty(0x12, VT_DATE, (void*)&result);
	return result;
}

void CAssignment::SetlFinishVariance(LONG propval)
{
	SetProperty(0x13, VT_I4, propval);
}

LONG CAssignment::GetlFinishVariance()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetsHyperlink(LPCTSTR propval)
{
	SetProperty(0x14, VT_BSTR, propval);
}

CString CAssignment::GetsHyperlink()
{
	CString result;
	GetProperty(0x14, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::SetsHyperlinkAddress(LPCTSTR propval)
{
	SetProperty(0x15, VT_BSTR, propval);
}

CString CAssignment::GetsHyperlinkAddress()
{
	CString result;
	GetProperty(0x15, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::SetsHyperlinkSubAddress(LPCTSTR propval)
{
	SetProperty(0x16, VT_BSTR, propval);
}

CString CAssignment::GetsHyperlinkSubAddress()
{
	CString result;
	GetProperty(0x16, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::SetfWorkVariance(FLOAT propval)
{
	SetProperty(0x17, VT_R4, propval);
}

FLOAT CAssignment::GetfWorkVariance()
{
	FLOAT result;
	GetProperty(0x17, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetbHasFixedRateUnits(BOOL propval)
{
	SetProperty(0x18, VT_BOOL, propval);
}

BOOL CAssignment::GetbHasFixedRateUnits()
{
	BOOL result;
	GetProperty(0x18, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetbFixedMaterial(BOOL propval)
{
	SetProperty(0x19, VT_BOOL, propval);
}

BOOL CAssignment::GetbFixedMaterial()
{
	BOOL result;
	GetProperty(0x19, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetlLevelingDelay(LONG propval)
{
	SetProperty(0x1a, VT_I4, propval);
}

LONG CAssignment::GetlLevelingDelay()
{
	LONG result;
	GetProperty(0x1a, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetyLevelingDelayFormat(LONG propval)
{
	SetProperty(0x1b, VT_I4, propval);
}

LONG CAssignment::GetyLevelingDelayFormat()
{
	LONG result;
	GetProperty(0x1b, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetbLinkedFields(BOOL propval)
{
	SetProperty(0x1c, VT_BOOL, propval);
}

BOOL CAssignment::GetbLinkedFields()
{
	BOOL result;
	GetProperty(0x1c, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetbMilestone(BOOL propval)
{
	SetProperty(0x1d, VT_BOOL, propval);
}

BOOL CAssignment::GetbMilestone()
{
	BOOL result;
	GetProperty(0x1d, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetsNotes(LPCTSTR propval)
{
	SetProperty(0x1e, VT_BSTR, propval);
}

CString CAssignment::GetsNotes()
{
	CString result;
	GetProperty(0x1e, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::SetbOverallocated(BOOL propval)
{
	SetProperty(0x1f, VT_BOOL, propval);
}

BOOL CAssignment::GetbOverallocated()
{
	BOOL result;
	GetProperty(0x1f, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetcOvertimeCost(DOUBLE propval)
{
	SetProperty(0x20, VT_R8, propval);
}

DOUBLE CAssignment::GetcOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x20, VT_R8, (void*)&result);
	return result;
}

CDuration CAssignment::GetoOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x21, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetfPeakUnits(FLOAT propval)
{
	SetProperty(0x22, VT_R4, propval);
}

FLOAT CAssignment::GetfPeakUnits()
{
	FLOAT result;
	GetProperty(0x22, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetlRateScale(LONG propval)
{
	SetProperty(0x23, VT_I4, propval);
}

LONG CAssignment::GetlRateScale()
{
	LONG result;
	GetProperty(0x23, VT_I4, (void*)&result);
	return result;
}

CDuration CAssignment::GetoRegularWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x24, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetcRemainingCost(DOUBLE propval)
{
	SetProperty(0x25, VT_R8, propval);
}

DOUBLE CAssignment::GetcRemainingCost()
{
	DOUBLE result;
	GetProperty(0x25, VT_R8, (void*)&result);
	return result;
}

void CAssignment::SetcRemainingOvertimeCost(DOUBLE propval)
{
	SetProperty(0x26, VT_R8, propval);
}

DOUBLE CAssignment::GetcRemainingOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x26, VT_R8, (void*)&result);
	return result;
}

CDuration CAssignment::GetoRemainingOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x27, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CAssignment::GetoRemainingWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x28, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetbResponsePending(BOOL propval)
{
	SetProperty(0x29, VT_BOOL, propval);
}

BOOL CAssignment::GetbResponsePending()
{
	BOOL result;
	GetProperty(0x29, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetdtStart(DATE propval)
{
	SetProperty(0x2a, VT_DATE, propval);
}

DATE CAssignment::GetdtStart()
{
	DATE result;
	GetProperty(0x2a, VT_DATE, (void*)&result);
	return result;
}

void CAssignment::SetdtStop(DATE propval)
{
	SetProperty(0x2b, VT_DATE, propval);
}

DATE CAssignment::GetdtStop()
{
	DATE result;
	GetProperty(0x2b, VT_DATE, (void*)&result);
	return result;
}

void CAssignment::SetdtResume(DATE propval)
{
	SetProperty(0x2c, VT_DATE, propval);
}

DATE CAssignment::GetdtResume()
{
	DATE result;
	GetProperty(0x2c, VT_DATE, (void*)&result);
	return result;
}

void CAssignment::SetlStartVariance(LONG propval)
{
	SetProperty(0x2d, VT_I4, propval);
}

LONG CAssignment::GetlStartVariance()
{
	LONG result;
	GetProperty(0x2d, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetbSummary(BOOL propval)
{
	SetProperty(0x2e, VT_BOOL, propval);
}

BOOL CAssignment::GetbSummary()
{
	BOOL result;
	GetProperty(0x2e, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetfSV(FLOAT propval)
{
	SetProperty(0x2f, VT_R4, propval);
}

FLOAT CAssignment::GetfSV()
{
	FLOAT result;
	GetProperty(0x2f, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetfUnits(FLOAT propval)
{
	SetProperty(0x30, VT_R4, propval);
}

FLOAT CAssignment::GetfUnits()
{
	FLOAT result;
	GetProperty(0x30, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetbUpdateNeeded(BOOL propval)
{
	SetProperty(0x31, VT_BOOL, propval);
}

BOOL CAssignment::GetbUpdateNeeded()
{
	BOOL result;
	GetProperty(0x31, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetfVAC(FLOAT propval)
{
	SetProperty(0x32, VT_R4, propval);
}

FLOAT CAssignment::GetfVAC()
{
	FLOAT result;
	GetProperty(0x32, VT_R4, (void*)&result);
	return result;
}

CDuration CAssignment::GetoWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x33, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetyWorkContour(LONG propval)
{
	SetProperty(0x34, VT_I4, propval);
}

LONG CAssignment::GetyWorkContour()
{
	LONG result;
	GetProperty(0x34, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetfBCWS(FLOAT propval)
{
	SetProperty(0x35, VT_R4, propval);
}

FLOAT CAssignment::GetfBCWS()
{
	FLOAT result;
	GetProperty(0x35, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetfBCWP(FLOAT propval)
{
	SetProperty(0x36, VT_R4, propval);
}

FLOAT CAssignment::GetfBCWP()
{
	FLOAT result;
	GetProperty(0x36, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetyBookingType(LONG propval)
{
	SetProperty(0x37, VT_I4, propval);
}

LONG CAssignment::GetyBookingType()
{
	LONG result;
	GetProperty(0x37, VT_I4, (void*)&result);
	return result;
}

CDuration CAssignment::GetoActualWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x38, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CAssignment::GetoActualOvertimeWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x39, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetdtCreationDate(DATE propval)
{
	SetProperty(0x3a, VT_DATE, propval);
}

DATE CAssignment::GetdtCreationDate()
{
	DATE result;
	GetProperty(0x3a, VT_DATE, (void*)&result);
	return result;
}

void CAssignment::SetsAssnOwner(LPCTSTR propval)
{
	SetProperty(0x3b, VT_BSTR, propval);
}

CString CAssignment::GetsAssnOwner()
{
	CString result;
	GetProperty(0x3b, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::SetsAssnOwnerGuid(LPCTSTR propval)
{
	SetProperty(0x3c, VT_BSTR, propval);
}

CString CAssignment::GetsAssnOwnerGuid()
{
	CString result;
	GetProperty(0x3c, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::SetcBudgetCost(DOUBLE propval)
{
	SetProperty(0x3d, VT_R8, propval);
}

DOUBLE CAssignment::GetcBudgetCost()
{
	DOUBLE result;
	GetProperty(0x3d, VT_R8, (void*)&result);
	return result;
}

CDuration CAssignment::GetoBudgetWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3e, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CAssignmentExtendedAttribute_C CAssignment::GetoExtendedAttribute_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3f, VT_DISPATCH, (void*)&pDispatch);
	return CAssignmentExtendedAttribute_C(pDispatch);
}

CAssignmentBaseline_C CAssignment::GetoBaseline_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x40, VT_DISPATCH, (void*)&pDispatch);
	return CAssignmentBaseline_C(pDispatch);
}

void CAssignment::Setsf404000(LPCTSTR propval)
{
	SetProperty(0x41, VT_BSTR, propval);
}

CString CAssignment::Getsf404000()
{
	CString result;
	GetProperty(0x41, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404001(LPCTSTR propval)
{
	SetProperty(0x42, VT_BSTR, propval);
}

CString CAssignment::Getsf404001()
{
	CString result;
	GetProperty(0x42, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404002(LPCTSTR propval)
{
	SetProperty(0x43, VT_BSTR, propval);
}

CString CAssignment::Getsf404002()
{
	CString result;
	GetProperty(0x43, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404003(LPCTSTR propval)
{
	SetProperty(0x44, VT_BSTR, propval);
}

CString CAssignment::Getsf404003()
{
	CString result;
	GetProperty(0x44, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404004(LPCTSTR propval)
{
	SetProperty(0x45, VT_BSTR, propval);
}

CString CAssignment::Getsf404004()
{
	CString result;
	GetProperty(0x45, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404005(LPCTSTR propval)
{
	SetProperty(0x46, VT_BSTR, propval);
}

CString CAssignment::Getsf404005()
{
	CString result;
	GetProperty(0x46, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404006(LPCTSTR propval)
{
	SetProperty(0x47, VT_BSTR, propval);
}

CString CAssignment::Getsf404006()
{
	CString result;
	GetProperty(0x47, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404007(LPCTSTR propval)
{
	SetProperty(0x48, VT_BSTR, propval);
}

CString CAssignment::Getsf404007()
{
	CString result;
	GetProperty(0x48, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404008(LPCTSTR propval)
{
	SetProperty(0x49, VT_BSTR, propval);
}

CString CAssignment::Getsf404008()
{
	CString result;
	GetProperty(0x49, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404009(LPCTSTR propval)
{
	SetProperty(0x4a, VT_BSTR, propval);
}

CString CAssignment::Getsf404009()
{
	CString result;
	GetProperty(0x4a, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40400a(LPCTSTR propval)
{
	SetProperty(0x4b, VT_BSTR, propval);
}

CString CAssignment::Getsf40400a()
{
	CString result;
	GetProperty(0x4b, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40400b(LPCTSTR propval)
{
	SetProperty(0x4c, VT_BSTR, propval);
}

CString CAssignment::Getsf40400b()
{
	CString result;
	GetProperty(0x4c, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40400c(LPCTSTR propval)
{
	SetProperty(0x4d, VT_BSTR, propval);
}

CString CAssignment::Getsf40400c()
{
	CString result;
	GetProperty(0x4d, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40400d(LPCTSTR propval)
{
	SetProperty(0x4e, VT_BSTR, propval);
}

CString CAssignment::Getsf40400d()
{
	CString result;
	GetProperty(0x4e, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40400e(LPCTSTR propval)
{
	SetProperty(0x4f, VT_BSTR, propval);
}

CString CAssignment::Getsf40400e()
{
	CString result;
	GetProperty(0x4f, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40400f(LPCTSTR propval)
{
	SetProperty(0x50, VT_BSTR, propval);
}

CString CAssignment::Getsf40400f()
{
	CString result;
	GetProperty(0x50, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404010(LPCTSTR propval)
{
	SetProperty(0x51, VT_BSTR, propval);
}

CString CAssignment::Getsf404010()
{
	CString result;
	GetProperty(0x51, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404011(LPCTSTR propval)
{
	SetProperty(0x52, VT_BSTR, propval);
}

CString CAssignment::Getsf404011()
{
	CString result;
	GetProperty(0x52, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404012(LPCTSTR propval)
{
	SetProperty(0x53, VT_BSTR, propval);
}

CString CAssignment::Getsf404012()
{
	CString result;
	GetProperty(0x53, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404013(LPCTSTR propval)
{
	SetProperty(0x54, VT_BSTR, propval);
}

CString CAssignment::Getsf404013()
{
	CString result;
	GetProperty(0x54, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404014(LPCTSTR propval)
{
	SetProperty(0x55, VT_BSTR, propval);
}

CString CAssignment::Getsf404014()
{
	CString result;
	GetProperty(0x55, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404015(LPCTSTR propval)
{
	SetProperty(0x56, VT_BSTR, propval);
}

CString CAssignment::Getsf404015()
{
	CString result;
	GetProperty(0x56, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404016(LPCTSTR propval)
{
	SetProperty(0x57, VT_BSTR, propval);
}

CString CAssignment::Getsf404016()
{
	CString result;
	GetProperty(0x57, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404017(LPCTSTR propval)
{
	SetProperty(0x58, VT_BSTR, propval);
}

CString CAssignment::Getsf404017()
{
	CString result;
	GetProperty(0x58, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404018(LPCTSTR propval)
{
	SetProperty(0x59, VT_BSTR, propval);
}

CString CAssignment::Getsf404018()
{
	CString result;
	GetProperty(0x59, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404019(LPCTSTR propval)
{
	SetProperty(0x5a, VT_BSTR, propval);
}

CString CAssignment::Getsf404019()
{
	CString result;
	GetProperty(0x5a, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40401a(LPCTSTR propval)
{
	SetProperty(0x5b, VT_BSTR, propval);
}

CString CAssignment::Getsf40401a()
{
	CString result;
	GetProperty(0x5b, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40401b(LPCTSTR propval)
{
	SetProperty(0x5c, VT_BSTR, propval);
}

CString CAssignment::Getsf40401b()
{
	CString result;
	GetProperty(0x5c, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40401c(LPCTSTR propval)
{
	SetProperty(0x5d, VT_BSTR, propval);
}

CString CAssignment::Getsf40401c()
{
	CString result;
	GetProperty(0x5d, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40401d(LPCTSTR propval)
{
	SetProperty(0x5e, VT_BSTR, propval);
}

CString CAssignment::Getsf40401d()
{
	CString result;
	GetProperty(0x5e, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40401e(LPCTSTR propval)
{
	SetProperty(0x5f, VT_BSTR, propval);
}

CString CAssignment::Getsf40401e()
{
	CString result;
	GetProperty(0x5f, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40401f(LPCTSTR propval)
{
	SetProperty(0x60, VT_BSTR, propval);
}

CString CAssignment::Getsf40401f()
{
	CString result;
	GetProperty(0x60, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404020(LPCTSTR propval)
{
	SetProperty(0x61, VT_BSTR, propval);
}

CString CAssignment::Getsf404020()
{
	CString result;
	GetProperty(0x61, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404021(LPCTSTR propval)
{
	SetProperty(0x62, VT_BSTR, propval);
}

CString CAssignment::Getsf404021()
{
	CString result;
	GetProperty(0x62, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404022(LPCTSTR propval)
{
	SetProperty(0x63, VT_BSTR, propval);
}

CString CAssignment::Getsf404022()
{
	CString result;
	GetProperty(0x63, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404023(LPCTSTR propval)
{
	SetProperty(0x64, VT_BSTR, propval);
}

CString CAssignment::Getsf404023()
{
	CString result;
	GetProperty(0x64, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404024(LPCTSTR propval)
{
	SetProperty(0x65, VT_BSTR, propval);
}

CString CAssignment::Getsf404024()
{
	CString result;
	GetProperty(0x65, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404025(LPCTSTR propval)
{
	SetProperty(0x66, VT_BSTR, propval);
}

CString CAssignment::Getsf404025()
{
	CString result;
	GetProperty(0x66, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404026(LPCTSTR propval)
{
	SetProperty(0x67, VT_BSTR, propval);
}

CString CAssignment::Getsf404026()
{
	CString result;
	GetProperty(0x67, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404027(LPCTSTR propval)
{
	SetProperty(0x68, VT_BSTR, propval);
}

CString CAssignment::Getsf404027()
{
	CString result;
	GetProperty(0x68, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404028(LPCTSTR propval)
{
	SetProperty(0x69, VT_BSTR, propval);
}

CString CAssignment::Getsf404028()
{
	CString result;
	GetProperty(0x69, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404029(LPCTSTR propval)
{
	SetProperty(0x6a, VT_BSTR, propval);
}

CString CAssignment::Getsf404029()
{
	CString result;
	GetProperty(0x6a, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40402a(LPCTSTR propval)
{
	SetProperty(0x6b, VT_BSTR, propval);
}

CString CAssignment::Getsf40402a()
{
	CString result;
	GetProperty(0x6b, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40402b(LPCTSTR propval)
{
	SetProperty(0x6c, VT_BSTR, propval);
}

CString CAssignment::Getsf40402b()
{
	CString result;
	GetProperty(0x6c, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40402c(LPCTSTR propval)
{
	SetProperty(0x6d, VT_BSTR, propval);
}

CString CAssignment::Getsf40402c()
{
	CString result;
	GetProperty(0x6d, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40402d(LPCTSTR propval)
{
	SetProperty(0x6e, VT_BSTR, propval);
}

CString CAssignment::Getsf40402d()
{
	CString result;
	GetProperty(0x6e, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40402e(LPCTSTR propval)
{
	SetProperty(0x6f, VT_BSTR, propval);
}

CString CAssignment::Getsf40402e()
{
	CString result;
	GetProperty(0x6f, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40402f(LPCTSTR propval)
{
	SetProperty(0x70, VT_BSTR, propval);
}

CString CAssignment::Getsf40402f()
{
	CString result;
	GetProperty(0x70, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404030(LPCTSTR propval)
{
	SetProperty(0x71, VT_BSTR, propval);
}

CString CAssignment::Getsf404030()
{
	CString result;
	GetProperty(0x71, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404031(LPCTSTR propval)
{
	SetProperty(0x72, VT_BSTR, propval);
}

CString CAssignment::Getsf404031()
{
	CString result;
	GetProperty(0x72, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404032(LPCTSTR propval)
{
	SetProperty(0x73, VT_BSTR, propval);
}

CString CAssignment::Getsf404032()
{
	CString result;
	GetProperty(0x73, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404033(LPCTSTR propval)
{
	SetProperty(0x74, VT_BSTR, propval);
}

CString CAssignment::Getsf404033()
{
	CString result;
	GetProperty(0x74, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404034(LPCTSTR propval)
{
	SetProperty(0x75, VT_BSTR, propval);
}

CString CAssignment::Getsf404034()
{
	CString result;
	GetProperty(0x75, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404035(LPCTSTR propval)
{
	SetProperty(0x76, VT_BSTR, propval);
}

CString CAssignment::Getsf404035()
{
	CString result;
	GetProperty(0x76, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404036(LPCTSTR propval)
{
	SetProperty(0x77, VT_BSTR, propval);
}

CString CAssignment::Getsf404036()
{
	CString result;
	GetProperty(0x77, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404037(LPCTSTR propval)
{
	SetProperty(0x78, VT_BSTR, propval);
}

CString CAssignment::Getsf404037()
{
	CString result;
	GetProperty(0x78, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404038(LPCTSTR propval)
{
	SetProperty(0x79, VT_BSTR, propval);
}

CString CAssignment::Getsf404038()
{
	CString result;
	GetProperty(0x79, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404039(LPCTSTR propval)
{
	SetProperty(0x7a, VT_BSTR, propval);
}

CString CAssignment::Getsf404039()
{
	CString result;
	GetProperty(0x7a, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40403a(LPCTSTR propval)
{
	SetProperty(0x7b, VT_BSTR, propval);
}

CString CAssignment::Getsf40403a()
{
	CString result;
	GetProperty(0x7b, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40403b(LPCTSTR propval)
{
	SetProperty(0x7c, VT_BSTR, propval);
}

CString CAssignment::Getsf40403b()
{
	CString result;
	GetProperty(0x7c, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40403c(LPCTSTR propval)
{
	SetProperty(0x7d, VT_BSTR, propval);
}

CString CAssignment::Getsf40403c()
{
	CString result;
	GetProperty(0x7d, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40403d(LPCTSTR propval)
{
	SetProperty(0x7e, VT_BSTR, propval);
}

CString CAssignment::Getsf40403d()
{
	CString result;
	GetProperty(0x7e, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40403e(LPCTSTR propval)
{
	SetProperty(0x7f, VT_BSTR, propval);
}

CString CAssignment::Getsf40403e()
{
	CString result;
	GetProperty(0x7f, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40403f(LPCTSTR propval)
{
	SetProperty(0x80, VT_BSTR, propval);
}

CString CAssignment::Getsf40403f()
{
	CString result;
	GetProperty(0x80, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404040(LPCTSTR propval)
{
	SetProperty(0x81, VT_BSTR, propval);
}

CString CAssignment::Getsf404040()
{
	CString result;
	GetProperty(0x81, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404041(LPCTSTR propval)
{
	SetProperty(0x82, VT_BSTR, propval);
}

CString CAssignment::Getsf404041()
{
	CString result;
	GetProperty(0x82, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404042(LPCTSTR propval)
{
	SetProperty(0x83, VT_BSTR, propval);
}

CString CAssignment::Getsf404042()
{
	CString result;
	GetProperty(0x83, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404043(LPCTSTR propval)
{
	SetProperty(0x84, VT_BSTR, propval);
}

CString CAssignment::Getsf404043()
{
	CString result;
	GetProperty(0x84, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404044(LPCTSTR propval)
{
	SetProperty(0x85, VT_BSTR, propval);
}

CString CAssignment::Getsf404044()
{
	CString result;
	GetProperty(0x85, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404045(LPCTSTR propval)
{
	SetProperty(0x86, VT_BSTR, propval);
}

CString CAssignment::Getsf404045()
{
	CString result;
	GetProperty(0x86, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404046(LPCTSTR propval)
{
	SetProperty(0x87, VT_BSTR, propval);
}

CString CAssignment::Getsf404046()
{
	CString result;
	GetProperty(0x87, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404047(LPCTSTR propval)
{
	SetProperty(0x88, VT_BSTR, propval);
}

CString CAssignment::Getsf404047()
{
	CString result;
	GetProperty(0x88, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404048(LPCTSTR propval)
{
	SetProperty(0x89, VT_BSTR, propval);
}

CString CAssignment::Getsf404048()
{
	CString result;
	GetProperty(0x89, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404049(LPCTSTR propval)
{
	SetProperty(0x8a, VT_BSTR, propval);
}

CString CAssignment::Getsf404049()
{
	CString result;
	GetProperty(0x8a, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40404a(LPCTSTR propval)
{
	SetProperty(0x8b, VT_BSTR, propval);
}

CString CAssignment::Getsf40404a()
{
	CString result;
	GetProperty(0x8b, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40404b(LPCTSTR propval)
{
	SetProperty(0x8c, VT_BSTR, propval);
}

CString CAssignment::Getsf40404b()
{
	CString result;
	GetProperty(0x8c, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40404c(LPCTSTR propval)
{
	SetProperty(0x8d, VT_BSTR, propval);
}

CString CAssignment::Getsf40404c()
{
	CString result;
	GetProperty(0x8d, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40404d(LPCTSTR propval)
{
	SetProperty(0x8e, VT_BSTR, propval);
}

CString CAssignment::Getsf40404d()
{
	CString result;
	GetProperty(0x8e, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40404e(LPCTSTR propval)
{
	SetProperty(0x8f, VT_BSTR, propval);
}

CString CAssignment::Getsf40404e()
{
	CString result;
	GetProperty(0x8f, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40404f(LPCTSTR propval)
{
	SetProperty(0x90, VT_BSTR, propval);
}

CString CAssignment::Getsf40404f()
{
	CString result;
	GetProperty(0x90, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404050(LPCTSTR propval)
{
	SetProperty(0x91, VT_BSTR, propval);
}

CString CAssignment::Getsf404050()
{
	CString result;
	GetProperty(0x91, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404051(LPCTSTR propval)
{
	SetProperty(0x92, VT_BSTR, propval);
}

CString CAssignment::Getsf404051()
{
	CString result;
	GetProperty(0x92, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404052(LPCTSTR propval)
{
	SetProperty(0x93, VT_BSTR, propval);
}

CString CAssignment::Getsf404052()
{
	CString result;
	GetProperty(0x93, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404053(LPCTSTR propval)
{
	SetProperty(0x94, VT_BSTR, propval);
}

CString CAssignment::Getsf404053()
{
	CString result;
	GetProperty(0x94, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404054(LPCTSTR propval)
{
	SetProperty(0x95, VT_BSTR, propval);
}

CString CAssignment::Getsf404054()
{
	CString result;
	GetProperty(0x95, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404055(LPCTSTR propval)
{
	SetProperty(0x96, VT_BSTR, propval);
}

CString CAssignment::Getsf404055()
{
	CString result;
	GetProperty(0x96, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404056(LPCTSTR propval)
{
	SetProperty(0x97, VT_BSTR, propval);
}

CString CAssignment::Getsf404056()
{
	CString result;
	GetProperty(0x97, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404057(LPCTSTR propval)
{
	SetProperty(0x98, VT_BSTR, propval);
}

CString CAssignment::Getsf404057()
{
	CString result;
	GetProperty(0x98, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404058(LPCTSTR propval)
{
	SetProperty(0x99, VT_BSTR, propval);
}

CString CAssignment::Getsf404058()
{
	CString result;
	GetProperty(0x99, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404059(LPCTSTR propval)
{
	SetProperty(0x9a, VT_BSTR, propval);
}

CString CAssignment::Getsf404059()
{
	CString result;
	GetProperty(0x9a, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40405a(LPCTSTR propval)
{
	SetProperty(0x9b, VT_BSTR, propval);
}

CString CAssignment::Getsf40405a()
{
	CString result;
	GetProperty(0x9b, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40405b(LPCTSTR propval)
{
	SetProperty(0x9c, VT_BSTR, propval);
}

CString CAssignment::Getsf40405b()
{
	CString result;
	GetProperty(0x9c, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40405c(LPCTSTR propval)
{
	SetProperty(0x9d, VT_BSTR, propval);
}

CString CAssignment::Getsf40405c()
{
	CString result;
	GetProperty(0x9d, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40405d(LPCTSTR propval)
{
	SetProperty(0x9e, VT_BSTR, propval);
}

CString CAssignment::Getsf40405d()
{
	CString result;
	GetProperty(0x9e, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40405e(LPCTSTR propval)
{
	SetProperty(0x9f, VT_BSTR, propval);
}

CString CAssignment::Getsf40405e()
{
	CString result;
	GetProperty(0x9f, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40405f(LPCTSTR propval)
{
	SetProperty(0xa0, VT_BSTR, propval);
}

CString CAssignment::Getsf40405f()
{
	CString result;
	GetProperty(0xa0, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404060(LPCTSTR propval)
{
	SetProperty(0xa1, VT_BSTR, propval);
}

CString CAssignment::Getsf404060()
{
	CString result;
	GetProperty(0xa1, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404061(LPCTSTR propval)
{
	SetProperty(0xa2, VT_BSTR, propval);
}

CString CAssignment::Getsf404061()
{
	CString result;
	GetProperty(0xa2, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404062(LPCTSTR propval)
{
	SetProperty(0xa3, VT_BSTR, propval);
}

CString CAssignment::Getsf404062()
{
	CString result;
	GetProperty(0xa3, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404063(LPCTSTR propval)
{
	SetProperty(0xa4, VT_BSTR, propval);
}

CString CAssignment::Getsf404063()
{
	CString result;
	GetProperty(0xa4, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404064(LPCTSTR propval)
{
	SetProperty(0xa5, VT_BSTR, propval);
}

CString CAssignment::Getsf404064()
{
	CString result;
	GetProperty(0xa5, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404065(LPCTSTR propval)
{
	SetProperty(0xa6, VT_BSTR, propval);
}

CString CAssignment::Getsf404065()
{
	CString result;
	GetProperty(0xa6, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404066(LPCTSTR propval)
{
	SetProperty(0xa7, VT_BSTR, propval);
}

CString CAssignment::Getsf404066()
{
	CString result;
	GetProperty(0xa7, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404067(LPCTSTR propval)
{
	SetProperty(0xa8, VT_BSTR, propval);
}

CString CAssignment::Getsf404067()
{
	CString result;
	GetProperty(0xa8, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404068(LPCTSTR propval)
{
	SetProperty(0xa9, VT_BSTR, propval);
}

CString CAssignment::Getsf404068()
{
	CString result;
	GetProperty(0xa9, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404069(LPCTSTR propval)
{
	SetProperty(0xaa, VT_BSTR, propval);
}

CString CAssignment::Getsf404069()
{
	CString result;
	GetProperty(0xaa, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40406a(LPCTSTR propval)
{
	SetProperty(0xab, VT_BSTR, propval);
}

CString CAssignment::Getsf40406a()
{
	CString result;
	GetProperty(0xab, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40406b(LPCTSTR propval)
{
	SetProperty(0xac, VT_BSTR, propval);
}

CString CAssignment::Getsf40406b()
{
	CString result;
	GetProperty(0xac, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40406c(LPCTSTR propval)
{
	SetProperty(0xad, VT_BSTR, propval);
}

CString CAssignment::Getsf40406c()
{
	CString result;
	GetProperty(0xad, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40406d(LPCTSTR propval)
{
	SetProperty(0xae, VT_BSTR, propval);
}

CString CAssignment::Getsf40406d()
{
	CString result;
	GetProperty(0xae, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40406e(LPCTSTR propval)
{
	SetProperty(0xaf, VT_BSTR, propval);
}

CString CAssignment::Getsf40406e()
{
	CString result;
	GetProperty(0xaf, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40406f(LPCTSTR propval)
{
	SetProperty(0xb0, VT_BSTR, propval);
}

CString CAssignment::Getsf40406f()
{
	CString result;
	GetProperty(0xb0, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404070(LPCTSTR propval)
{
	SetProperty(0xb1, VT_BSTR, propval);
}

CString CAssignment::Getsf404070()
{
	CString result;
	GetProperty(0xb1, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404071(LPCTSTR propval)
{
	SetProperty(0xb2, VT_BSTR, propval);
}

CString CAssignment::Getsf404071()
{
	CString result;
	GetProperty(0xb2, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404072(LPCTSTR propval)
{
	SetProperty(0xb3, VT_BSTR, propval);
}

CString CAssignment::Getsf404072()
{
	CString result;
	GetProperty(0xb3, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404073(LPCTSTR propval)
{
	SetProperty(0xb4, VT_BSTR, propval);
}

CString CAssignment::Getsf404073()
{
	CString result;
	GetProperty(0xb4, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404074(LPCTSTR propval)
{
	SetProperty(0xb5, VT_BSTR, propval);
}

CString CAssignment::Getsf404074()
{
	CString result;
	GetProperty(0xb5, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404075(LPCTSTR propval)
{
	SetProperty(0xb6, VT_BSTR, propval);
}

CString CAssignment::Getsf404075()
{
	CString result;
	GetProperty(0xb6, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404076(LPCTSTR propval)
{
	SetProperty(0xb7, VT_BSTR, propval);
}

CString CAssignment::Getsf404076()
{
	CString result;
	GetProperty(0xb7, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404077(LPCTSTR propval)
{
	SetProperty(0xb8, VT_BSTR, propval);
}

CString CAssignment::Getsf404077()
{
	CString result;
	GetProperty(0xb8, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404078(LPCTSTR propval)
{
	SetProperty(0xb9, VT_BSTR, propval);
}

CString CAssignment::Getsf404078()
{
	CString result;
	GetProperty(0xb9, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404079(LPCTSTR propval)
{
	SetProperty(0xba, VT_BSTR, propval);
}

CString CAssignment::Getsf404079()
{
	CString result;
	GetProperty(0xba, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40407a(LPCTSTR propval)
{
	SetProperty(0xbb, VT_BSTR, propval);
}

CString CAssignment::Getsf40407a()
{
	CString result;
	GetProperty(0xbb, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40407b(LPCTSTR propval)
{
	SetProperty(0xbc, VT_BSTR, propval);
}

CString CAssignment::Getsf40407b()
{
	CString result;
	GetProperty(0xbc, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40407c(LPCTSTR propval)
{
	SetProperty(0xbd, VT_BSTR, propval);
}

CString CAssignment::Getsf40407c()
{
	CString result;
	GetProperty(0xbd, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40407d(LPCTSTR propval)
{
	SetProperty(0xbe, VT_BSTR, propval);
}

CString CAssignment::Getsf40407d()
{
	CString result;
	GetProperty(0xbe, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40407e(LPCTSTR propval)
{
	SetProperty(0xbf, VT_BSTR, propval);
}

CString CAssignment::Getsf40407e()
{
	CString result;
	GetProperty(0xbf, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40407f(LPCTSTR propval)
{
	SetProperty(0xc0, VT_BSTR, propval);
}

CString CAssignment::Getsf40407f()
{
	CString result;
	GetProperty(0xc0, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404080(LPCTSTR propval)
{
	SetProperty(0xc1, VT_BSTR, propval);
}

CString CAssignment::Getsf404080()
{
	CString result;
	GetProperty(0xc1, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404081(LPCTSTR propval)
{
	SetProperty(0xc2, VT_BSTR, propval);
}

CString CAssignment::Getsf404081()
{
	CString result;
	GetProperty(0xc2, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404082(LPCTSTR propval)
{
	SetProperty(0xc3, VT_BSTR, propval);
}

CString CAssignment::Getsf404082()
{
	CString result;
	GetProperty(0xc3, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404083(LPCTSTR propval)
{
	SetProperty(0xc4, VT_BSTR, propval);
}

CString CAssignment::Getsf404083()
{
	CString result;
	GetProperty(0xc4, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404084(LPCTSTR propval)
{
	SetProperty(0xc5, VT_BSTR, propval);
}

CString CAssignment::Getsf404084()
{
	CString result;
	GetProperty(0xc5, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404085(LPCTSTR propval)
{
	SetProperty(0xc6, VT_BSTR, propval);
}

CString CAssignment::Getsf404085()
{
	CString result;
	GetProperty(0xc6, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404086(LPCTSTR propval)
{
	SetProperty(0xc7, VT_BSTR, propval);
}

CString CAssignment::Getsf404086()
{
	CString result;
	GetProperty(0xc7, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404087(LPCTSTR propval)
{
	SetProperty(0xc8, VT_BSTR, propval);
}

CString CAssignment::Getsf404087()
{
	CString result;
	GetProperty(0xc8, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404088(LPCTSTR propval)
{
	SetProperty(0xc9, VT_BSTR, propval);
}

CString CAssignment::Getsf404088()
{
	CString result;
	GetProperty(0xc9, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404089(LPCTSTR propval)
{
	SetProperty(0xca, VT_BSTR, propval);
}

CString CAssignment::Getsf404089()
{
	CString result;
	GetProperty(0xca, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40408a(LPCTSTR propval)
{
	SetProperty(0xcb, VT_BSTR, propval);
}

CString CAssignment::Getsf40408a()
{
	CString result;
	GetProperty(0xcb, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40408b(LPCTSTR propval)
{
	SetProperty(0xcc, VT_BSTR, propval);
}

CString CAssignment::Getsf40408b()
{
	CString result;
	GetProperty(0xcc, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40408c(LPCTSTR propval)
{
	SetProperty(0xcd, VT_BSTR, propval);
}

CString CAssignment::Getsf40408c()
{
	CString result;
	GetProperty(0xcd, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40408d(LPCTSTR propval)
{
	SetProperty(0xce, VT_BSTR, propval);
}

CString CAssignment::Getsf40408d()
{
	CString result;
	GetProperty(0xce, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40408e(LPCTSTR propval)
{
	SetProperty(0xcf, VT_BSTR, propval);
}

CString CAssignment::Getsf40408e()
{
	CString result;
	GetProperty(0xcf, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40408f(LPCTSTR propval)
{
	SetProperty(0xd0, VT_BSTR, propval);
}

CString CAssignment::Getsf40408f()
{
	CString result;
	GetProperty(0xd0, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404090(LPCTSTR propval)
{
	SetProperty(0xd1, VT_BSTR, propval);
}

CString CAssignment::Getsf404090()
{
	CString result;
	GetProperty(0xd1, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404091(LPCTSTR propval)
{
	SetProperty(0xd2, VT_BSTR, propval);
}

CString CAssignment::Getsf404091()
{
	CString result;
	GetProperty(0xd2, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404092(LPCTSTR propval)
{
	SetProperty(0xd3, VT_BSTR, propval);
}

CString CAssignment::Getsf404092()
{
	CString result;
	GetProperty(0xd3, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404093(LPCTSTR propval)
{
	SetProperty(0xd4, VT_BSTR, propval);
}

CString CAssignment::Getsf404093()
{
	CString result;
	GetProperty(0xd4, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404094(LPCTSTR propval)
{
	SetProperty(0xd5, VT_BSTR, propval);
}

CString CAssignment::Getsf404094()
{
	CString result;
	GetProperty(0xd5, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404095(LPCTSTR propval)
{
	SetProperty(0xd6, VT_BSTR, propval);
}

CString CAssignment::Getsf404095()
{
	CString result;
	GetProperty(0xd6, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404096(LPCTSTR propval)
{
	SetProperty(0xd7, VT_BSTR, propval);
}

CString CAssignment::Getsf404096()
{
	CString result;
	GetProperty(0xd7, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404097(LPCTSTR propval)
{
	SetProperty(0xd8, VT_BSTR, propval);
}

CString CAssignment::Getsf404097()
{
	CString result;
	GetProperty(0xd8, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404098(LPCTSTR propval)
{
	SetProperty(0xd9, VT_BSTR, propval);
}

CString CAssignment::Getsf404098()
{
	CString result;
	GetProperty(0xd9, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf404099(LPCTSTR propval)
{
	SetProperty(0xda, VT_BSTR, propval);
}

CString CAssignment::Getsf404099()
{
	CString result;
	GetProperty(0xda, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40409a(LPCTSTR propval)
{
	SetProperty(0xdb, VT_BSTR, propval);
}

CString CAssignment::Getsf40409a()
{
	CString result;
	GetProperty(0xdb, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40409b(LPCTSTR propval)
{
	SetProperty(0xdc, VT_BSTR, propval);
}

CString CAssignment::Getsf40409b()
{
	CString result;
	GetProperty(0xdc, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40409c(LPCTSTR propval)
{
	SetProperty(0xdd, VT_BSTR, propval);
}

CString CAssignment::Getsf40409c()
{
	CString result;
	GetProperty(0xdd, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40409d(LPCTSTR propval)
{
	SetProperty(0xde, VT_BSTR, propval);
}

CString CAssignment::Getsf40409d()
{
	CString result;
	GetProperty(0xde, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40409e(LPCTSTR propval)
{
	SetProperty(0xdf, VT_BSTR, propval);
}

CString CAssignment::Getsf40409e()
{
	CString result;
	GetProperty(0xdf, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf40409f(LPCTSTR propval)
{
	SetProperty(0xe0, VT_BSTR, propval);
}

CString CAssignment::Getsf40409f()
{
	CString result;
	GetProperty(0xe0, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a0(LPCTSTR propval)
{
	SetProperty(0xe1, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a0()
{
	CString result;
	GetProperty(0xe1, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a1(LPCTSTR propval)
{
	SetProperty(0xe2, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a1()
{
	CString result;
	GetProperty(0xe2, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a2(LPCTSTR propval)
{
	SetProperty(0xe3, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a2()
{
	CString result;
	GetProperty(0xe3, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a3(LPCTSTR propval)
{
	SetProperty(0xe4, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a3()
{
	CString result;
	GetProperty(0xe4, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a4(LPCTSTR propval)
{
	SetProperty(0xe5, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a4()
{
	CString result;
	GetProperty(0xe5, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a5(LPCTSTR propval)
{
	SetProperty(0xe6, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a5()
{
	CString result;
	GetProperty(0xe6, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a6(LPCTSTR propval)
{
	SetProperty(0xe7, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a6()
{
	CString result;
	GetProperty(0xe7, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a7(LPCTSTR propval)
{
	SetProperty(0xe8, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a7()
{
	CString result;
	GetProperty(0xe8, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a8(LPCTSTR propval)
{
	SetProperty(0xe9, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a8()
{
	CString result;
	GetProperty(0xe9, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040a9(LPCTSTR propval)
{
	SetProperty(0xea, VT_BSTR, propval);
}

CString CAssignment::Getsf4040a9()
{
	CString result;
	GetProperty(0xea, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040aa(LPCTSTR propval)
{
	SetProperty(0xeb, VT_BSTR, propval);
}

CString CAssignment::Getsf4040aa()
{
	CString result;
	GetProperty(0xeb, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040ab(LPCTSTR propval)
{
	SetProperty(0xec, VT_BSTR, propval);
}

CString CAssignment::Getsf4040ab()
{
	CString result;
	GetProperty(0xec, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040ac(LPCTSTR propval)
{
	SetProperty(0xed, VT_BSTR, propval);
}

CString CAssignment::Getsf4040ac()
{
	CString result;
	GetProperty(0xed, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040ad(LPCTSTR propval)
{
	SetProperty(0xee, VT_BSTR, propval);
}

CString CAssignment::Getsf4040ad()
{
	CString result;
	GetProperty(0xee, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040ae(LPCTSTR propval)
{
	SetProperty(0xef, VT_BSTR, propval);
}

CString CAssignment::Getsf4040ae()
{
	CString result;
	GetProperty(0xef, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040af(LPCTSTR propval)
{
	SetProperty(0xf0, VT_BSTR, propval);
}

CString CAssignment::Getsf4040af()
{
	CString result;
	GetProperty(0xf0, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b0(LPCTSTR propval)
{
	SetProperty(0xf1, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b0()
{
	CString result;
	GetProperty(0xf1, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b1(LPCTSTR propval)
{
	SetProperty(0xf2, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b1()
{
	CString result;
	GetProperty(0xf2, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b2(LPCTSTR propval)
{
	SetProperty(0xf3, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b2()
{
	CString result;
	GetProperty(0xf3, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b3(LPCTSTR propval)
{
	SetProperty(0xf4, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b3()
{
	CString result;
	GetProperty(0xf4, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b4(LPCTSTR propval)
{
	SetProperty(0xf5, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b4()
{
	CString result;
	GetProperty(0xf5, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b5(LPCTSTR propval)
{
	SetProperty(0xf6, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b5()
{
	CString result;
	GetProperty(0xf6, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b6(LPCTSTR propval)
{
	SetProperty(0xf7, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b6()
{
	CString result;
	GetProperty(0xf7, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b7(LPCTSTR propval)
{
	SetProperty(0xf8, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b7()
{
	CString result;
	GetProperty(0xf8, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b8(LPCTSTR propval)
{
	SetProperty(0xf9, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b8()
{
	CString result;
	GetProperty(0xf9, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040b9(LPCTSTR propval)
{
	SetProperty(0xfa, VT_BSTR, propval);
}

CString CAssignment::Getsf4040b9()
{
	CString result;
	GetProperty(0xfa, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040ba(LPCTSTR propval)
{
	SetProperty(0xfb, VT_BSTR, propval);
}

CString CAssignment::Getsf4040ba()
{
	CString result;
	GetProperty(0xfb, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040bb(LPCTSTR propval)
{
	SetProperty(0xfc, VT_BSTR, propval);
}

CString CAssignment::Getsf4040bb()
{
	CString result;
	GetProperty(0xfc, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040bc(LPCTSTR propval)
{
	SetProperty(0xfd, VT_BSTR, propval);
}

CString CAssignment::Getsf4040bc()
{
	CString result;
	GetProperty(0xfd, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040bd(LPCTSTR propval)
{
	SetProperty(0xfe, VT_BSTR, propval);
}

CString CAssignment::Getsf4040bd()
{
	CString result;
	GetProperty(0xfe, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040be(LPCTSTR propval)
{
	SetProperty(0xff, VT_BSTR, propval);
}

CString CAssignment::Getsf4040be()
{
	CString result;
	GetProperty(0xff, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040bf(LPCTSTR propval)
{
	SetProperty(0x100, VT_BSTR, propval);
}

CString CAssignment::Getsf4040bf()
{
	CString result;
	GetProperty(0x100, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040c0(LPCTSTR propval)
{
	SetProperty(0x101, VT_BSTR, propval);
}

CString CAssignment::Getsf4040c0()
{
	CString result;
	GetProperty(0x101, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040c1(LPCTSTR propval)
{
	SetProperty(0x102, VT_BSTR, propval);
}

CString CAssignment::Getsf4040c1()
{
	CString result;
	GetProperty(0x102, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040c2(LPCTSTR propval)
{
	SetProperty(0x103, VT_BSTR, propval);
}

CString CAssignment::Getsf4040c2()
{
	CString result;
	GetProperty(0x103, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040c3(LPCTSTR propval)
{
	SetProperty(0x104, VT_BSTR, propval);
}

CString CAssignment::Getsf4040c3()
{
	CString result;
	GetProperty(0x104, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040c4(LPCTSTR propval)
{
	SetProperty(0x105, VT_BSTR, propval);
}

CString CAssignment::Getsf4040c4()
{
	CString result;
	GetProperty(0x105, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040c5(LPCTSTR propval)
{
	SetProperty(0x106, VT_BSTR, propval);
}

CString CAssignment::Getsf4040c5()
{
	CString result;
	GetProperty(0x106, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040c6(LPCTSTR propval)
{
	SetProperty(0x107, VT_BSTR, propval);
}

CString CAssignment::Getsf4040c6()
{
	CString result;
	GetProperty(0x107, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040c7(LPCTSTR propval)
{
	SetProperty(0x108, VT_BSTR, propval);
}

CString CAssignment::Getsf4040c7()
{
	CString result;
	GetProperty(0x108, VT_BSTR, (void*)&result);
	return result;
}

void CAssignment::Setsf4040c8(LPCTSTR propval)
{
	SetProperty(0x109, VT_BSTR, propval);
}

CString CAssignment::Getsf4040c8()
{
	CString result;
	GetProperty(0x109, VT_BSTR, (void*)&result);
	return result;
}

CTimephasedData_C CAssignment::GetoTimephasedData_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x10a, VT_DISPATCH, (void*)&pDispatch);
	return CTimephasedData_C(pDispatch);
}

void CAssignment::SetKey(LPCTSTR propval)
{
	SetProperty(0x10b, VT_BSTR, propval);
}

CString CAssignment::GetKey()
{
	CString result;
	GetProperty(0x10b, VT_BSTR, (void*)&result);
	return result;
}

CString CAssignment::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x10c, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignment::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x10d, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CAssignment::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x10e, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignment::Initialize(void)
{
	InvokeHelper(0x10f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CAssignments::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CAssignment CAssignments::Item(LPCTSTR Index)
{
	CAssignment oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CAssignment CAssignments::Add(void)
{
	CAssignment oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignments::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CAssignments::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CAssignments::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignments::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CAssignments::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CAssignments::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceRate::SetdtRatesFrom(DATE propval)
{
	SetProperty(0x1, VT_DATE, propval);
}

DATE CResourceRate::GetdtRatesFrom()
{
	DATE result;
	GetProperty(0x1, VT_DATE, (void*)&result);
	return result;
}

void CResourceRate::SetdtRatesTo(DATE propval)
{
	SetProperty(0x2, VT_DATE, propval);
}

DATE CResourceRate::GetdtRatesTo()
{
	DATE result;
	GetProperty(0x2, VT_DATE, (void*)&result);
	return result;
}

void CResourceRate::SetyRateTable(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CResourceRate::GetyRateTable()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CResourceRate::SetcStandardRate(DOUBLE propval)
{
	SetProperty(0x4, VT_R8, propval);
}

DOUBLE CResourceRate::GetcStandardRate()
{
	DOUBLE result;
	GetProperty(0x4, VT_R8, (void*)&result);
	return result;
}

void CResourceRate::SetyStandardRateFormat(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CResourceRate::GetyStandardRateFormat()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CResourceRate::SetcOvertimeRate(DOUBLE propval)
{
	SetProperty(0x6, VT_R8, propval);
}

DOUBLE CResourceRate::GetcOvertimeRate()
{
	DOUBLE result;
	GetProperty(0x6, VT_R8, (void*)&result);
	return result;
}

void CResourceRate::SetyOvertimeRateFormat(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CResourceRate::GetyOvertimeRateFormat()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CResourceRate::SetcCostPerUse(DOUBLE propval)
{
	SetProperty(0x8, VT_R8, propval);
}

DOUBLE CResourceRate::GetcCostPerUse()
{
	DOUBLE result;
	GetProperty(0x8, VT_R8, (void*)&result);
	return result;
}

void CResourceRate::SetKey(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CResourceRate::GetKey()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

CString CResourceRate::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceRate::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CResourceRate::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0xc, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceRate::Initialize(void)
{
	InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CResourceRates::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CResourceRate CResourceRates::Item(LPCTSTR Index)
{
	CResourceRate oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CResourceRate CResourceRates::Add(void)
{
	CResourceRate oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceRates::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResourceRates::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CResourceRates::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceRates::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResourceRates::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CResourceRates::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceAvailabilityPeriod::SetdtAvailableFrom(DATE propval)
{
	SetProperty(0x1, VT_DATE, propval);
}

DATE CResourceAvailabilityPeriod::GetdtAvailableFrom()
{
	DATE result;
	GetProperty(0x1, VT_DATE, (void*)&result);
	return result;
}

void CResourceAvailabilityPeriod::SetdtAvailableTo(DATE propval)
{
	SetProperty(0x2, VT_DATE, propval);
}

DATE CResourceAvailabilityPeriod::GetdtAvailableTo()
{
	DATE result;
	GetProperty(0x2, VT_DATE, (void*)&result);
	return result;
}

void CResourceAvailabilityPeriod::SetfAvailableUnits(FLOAT propval)
{
	SetProperty(0x3, VT_R4, propval);
}

FLOAT CResourceAvailabilityPeriod::GetfAvailableUnits()
{
	FLOAT result;
	GetProperty(0x3, VT_R4, (void*)&result);
	return result;
}

void CResourceAvailabilityPeriod::SetKey(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CResourceAvailabilityPeriod::GetKey()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

CString CResourceAvailabilityPeriod::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceAvailabilityPeriod::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CResourceAvailabilityPeriod::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceAvailabilityPeriod::Initialize(void)
{
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CResourceAvailabilityPeriods::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CResourceAvailabilityPeriod CResourceAvailabilityPeriods::Item(LPCTSTR Index)
{
	CResourceAvailabilityPeriod oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CResourceAvailabilityPeriod CResourceAvailabilityPeriods::Add(void)
{
	CResourceAvailabilityPeriod oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceAvailabilityPeriods::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResourceAvailabilityPeriods::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CResourceAvailabilityPeriods::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceAvailabilityPeriods::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResourceAvailabilityPeriods::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CResourceAvailabilityPeriods::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceOutlineCode::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CResourceOutlineCode::GetsFieldID()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CResourceOutlineCode::SetlValueID(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CResourceOutlineCode::GetlValueID()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CResourceOutlineCode::SetlValueGUID(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CResourceOutlineCode::GetlValueGUID()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CResourceOutlineCode::SetKey(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CResourceOutlineCode::GetKey()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

CString CResourceOutlineCode::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceOutlineCode::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CResourceOutlineCode::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceOutlineCode::Initialize(void)
{
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CResourceOutlineCode_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CResourceOutlineCode CResourceOutlineCode_C::Item(LPCTSTR Index)
{
	CResourceOutlineCode oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CResourceOutlineCode CResourceOutlineCode_C::Add(void)
{
	CResourceOutlineCode oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceOutlineCode_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResourceOutlineCode_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CResourceOutlineCode_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceOutlineCode_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResourceBaseline::SetlNumber(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CResourceBaseline::GetlNumber()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CDuration CResourceBaseline::GetoWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CResourceBaseline::SetfCost(FLOAT propval)
{
	SetProperty(0x3, VT_R4, propval);
}

FLOAT CResourceBaseline::GetfCost()
{
	FLOAT result;
	GetProperty(0x3, VT_R4, (void*)&result);
	return result;
}

void CResourceBaseline::SetfBCWS(FLOAT propval)
{
	SetProperty(0x4, VT_R4, propval);
}

FLOAT CResourceBaseline::GetfBCWS()
{
	FLOAT result;
	GetProperty(0x4, VT_R4, (void*)&result);
	return result;
}

void CResourceBaseline::SetfBCWP(FLOAT propval)
{
	SetProperty(0x5, VT_R4, propval);
}

FLOAT CResourceBaseline::GetfBCWP()
{
	FLOAT result;
	GetProperty(0x5, VT_R4, (void*)&result);
	return result;
}

void CResourceBaseline::SetKey(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CResourceBaseline::GetKey()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

CString CResourceBaseline::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceBaseline::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CResourceBaseline::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceBaseline::Initialize(void)
{
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CResourceBaseline_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CResourceBaseline CResourceBaseline_C::Item(LPCTSTR Index)
{
	CResourceBaseline oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CResourceBaseline CResourceBaseline_C::Add(void)
{
	CResourceBaseline oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceBaseline_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResourceBaseline_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CResourceBaseline_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceBaseline_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResourceExtendedAttribute::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CResourceExtendedAttribute::GetsFieldID()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CResourceExtendedAttribute::SetsValue(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CResourceExtendedAttribute::GetsValue()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CResourceExtendedAttribute::SetlValueGUID(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CResourceExtendedAttribute::GetlValueGUID()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CResourceExtendedAttribute::SetyDurationFormat(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CResourceExtendedAttribute::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CResourceExtendedAttribute::SetKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CResourceExtendedAttribute::GetKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CString CResourceExtendedAttribute::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceExtendedAttribute::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CResourceExtendedAttribute::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceExtendedAttribute::Initialize(void)
{
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CResourceExtendedAttribute_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CResourceExtendedAttribute CResourceExtendedAttribute_C::Item(LPCTSTR Index)
{
	CResourceExtendedAttribute oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CResourceExtendedAttribute CResourceExtendedAttribute_C::Add(void)
{
	CResourceExtendedAttribute oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceExtendedAttribute_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResourceExtendedAttribute_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CResourceExtendedAttribute_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceExtendedAttribute_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResource::SetlUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CResource::GetlUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CResource::SetlID(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CResource::GetlID()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CResource::SetsName(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CResource::GetsName()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetyType(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CResource::GetyType()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CResource::SetbIsNull(BOOL propval)
{
	SetProperty(0x5, VT_BOOL, propval);
}

BOOL CResource::GetbIsNull()
{
	BOOL result;
	GetProperty(0x5, VT_BOOL, (void*)&result);
	return result;
}

void CResource::SetsInitials(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CResource::GetsInitials()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetsPhonetics(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CResource::GetsPhonetics()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetsNTAccount(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CResource::GetsNTAccount()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetsMaterialLabel(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CResource::GetsMaterialLabel()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetsCode(LPCTSTR propval)
{
	SetProperty(0xa, VT_BSTR, propval);
}

CString CResource::GetsCode()
{
	CString result;
	GetProperty(0xa, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetsGroup(LPCTSTR propval)
{
	SetProperty(0xb, VT_BSTR, propval);
}

CString CResource::GetsGroup()
{
	CString result;
	GetProperty(0xb, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetyWorkGroup(LONG propval)
{
	SetProperty(0xc, VT_I4, propval);
}

LONG CResource::GetyWorkGroup()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

void CResource::SetsEmailAddress(LPCTSTR propval)
{
	SetProperty(0xd, VT_BSTR, propval);
}

CString CResource::GetsEmailAddress()
{
	CString result;
	GetProperty(0xd, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetsHyperlink(LPCTSTR propval)
{
	SetProperty(0xe, VT_BSTR, propval);
}

CString CResource::GetsHyperlink()
{
	CString result;
	GetProperty(0xe, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetsHyperlinkAddress(LPCTSTR propval)
{
	SetProperty(0xf, VT_BSTR, propval);
}

CString CResource::GetsHyperlinkAddress()
{
	CString result;
	GetProperty(0xf, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetsHyperlinkSubAddress(LPCTSTR propval)
{
	SetProperty(0x10, VT_BSTR, propval);
}

CString CResource::GetsHyperlinkSubAddress()
{
	CString result;
	GetProperty(0x10, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetfMaxUnits(FLOAT propval)
{
	SetProperty(0x11, VT_R4, propval);
}

FLOAT CResource::GetfMaxUnits()
{
	FLOAT result;
	GetProperty(0x11, VT_R4, (void*)&result);
	return result;
}

void CResource::SetfPeakUnits(FLOAT propval)
{
	SetProperty(0x12, VT_R4, propval);
}

FLOAT CResource::GetfPeakUnits()
{
	FLOAT result;
	GetProperty(0x12, VT_R4, (void*)&result);
	return result;
}

void CResource::SetbOverAllocated(BOOL propval)
{
	SetProperty(0x13, VT_BOOL, propval);
}

BOOL CResource::GetbOverAllocated()
{
	BOOL result;
	GetProperty(0x13, VT_BOOL, (void*)&result);
	return result;
}

void CResource::SetdtAvailableFrom(DATE propval)
{
	SetProperty(0x14, VT_DATE, propval);
}

DATE CResource::GetdtAvailableFrom()
{
	DATE result;
	GetProperty(0x14, VT_DATE, (void*)&result);
	return result;
}

void CResource::SetdtAvailableTo(DATE propval)
{
	SetProperty(0x15, VT_DATE, propval);
}

DATE CResource::GetdtAvailableTo()
{
	DATE result;
	GetProperty(0x15, VT_DATE, (void*)&result);
	return result;
}

void CResource::SetdtStart(DATE propval)
{
	SetProperty(0x16, VT_DATE, propval);
}

DATE CResource::GetdtStart()
{
	DATE result;
	GetProperty(0x16, VT_DATE, (void*)&result);
	return result;
}

void CResource::SetdtFinish(DATE propval)
{
	SetProperty(0x17, VT_DATE, propval);
}

DATE CResource::GetdtFinish()
{
	DATE result;
	GetProperty(0x17, VT_DATE, (void*)&result);
	return result;
}

void CResource::SetbCanLevel(BOOL propval)
{
	SetProperty(0x18, VT_BOOL, propval);
}

BOOL CResource::GetbCanLevel()
{
	BOOL result;
	GetProperty(0x18, VT_BOOL, (void*)&result);
	return result;
}

void CResource::SetyAccrueAt(LONG propval)
{
	SetProperty(0x19, VT_I4, propval);
}

LONG CResource::GetyAccrueAt()
{
	LONG result;
	GetProperty(0x19, VT_I4, (void*)&result);
	return result;
}

CDuration CResource::GetoWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1a, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CResource::GetoRegularWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1b, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CResource::GetoOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1c, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CResource::GetoActualWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1d, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CResource::GetoRemainingWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1e, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CResource::GetoActualOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1f, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CResource::GetoRemainingOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x20, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CResource::SetlPercentWorkComplete(LONG propval)
{
	SetProperty(0x21, VT_I4, propval);
}

LONG CResource::GetlPercentWorkComplete()
{
	LONG result;
	GetProperty(0x21, VT_I4, (void*)&result);
	return result;
}

void CResource::SetcStandardRate(DOUBLE propval)
{
	SetProperty(0x22, VT_R8, propval);
}

DOUBLE CResource::GetcStandardRate()
{
	DOUBLE result;
	GetProperty(0x22, VT_R8, (void*)&result);
	return result;
}

void CResource::SetyStandardRateFormat(LONG propval)
{
	SetProperty(0x23, VT_I4, propval);
}

LONG CResource::GetyStandardRateFormat()
{
	LONG result;
	GetProperty(0x23, VT_I4, (void*)&result);
	return result;
}

void CResource::SetcCost(DOUBLE propval)
{
	SetProperty(0x24, VT_R8, propval);
}

DOUBLE CResource::GetcCost()
{
	DOUBLE result;
	GetProperty(0x24, VT_R8, (void*)&result);
	return result;
}

void CResource::SetcOvertimeRate(DOUBLE propval)
{
	SetProperty(0x25, VT_R8, propval);
}

DOUBLE CResource::GetcOvertimeRate()
{
	DOUBLE result;
	GetProperty(0x25, VT_R8, (void*)&result);
	return result;
}

void CResource::SetyOvertimeRateFormat(LONG propval)
{
	SetProperty(0x26, VT_I4, propval);
}

LONG CResource::GetyOvertimeRateFormat()
{
	LONG result;
	GetProperty(0x26, VT_I4, (void*)&result);
	return result;
}

void CResource::SetcOvertimeCost(DOUBLE propval)
{
	SetProperty(0x27, VT_R8, propval);
}

DOUBLE CResource::GetcOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x27, VT_R8, (void*)&result);
	return result;
}

void CResource::SetcCostPerUse(DOUBLE propval)
{
	SetProperty(0x28, VT_R8, propval);
}

DOUBLE CResource::GetcCostPerUse()
{
	DOUBLE result;
	GetProperty(0x28, VT_R8, (void*)&result);
	return result;
}

void CResource::SetcActualCost(DOUBLE propval)
{
	SetProperty(0x29, VT_R8, propval);
}

DOUBLE CResource::GetcActualCost()
{
	DOUBLE result;
	GetProperty(0x29, VT_R8, (void*)&result);
	return result;
}

void CResource::SetcActualOvertimeCost(DOUBLE propval)
{
	SetProperty(0x2a, VT_R8, propval);
}

DOUBLE CResource::GetcActualOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x2a, VT_R8, (void*)&result);
	return result;
}

void CResource::SetcRemainingCost(DOUBLE propval)
{
	SetProperty(0x2b, VT_R8, propval);
}

DOUBLE CResource::GetcRemainingCost()
{
	DOUBLE result;
	GetProperty(0x2b, VT_R8, (void*)&result);
	return result;
}

void CResource::SetcRemainingOvertimeCost(DOUBLE propval)
{
	SetProperty(0x2c, VT_R8, propval);
}

DOUBLE CResource::GetcRemainingOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x2c, VT_R8, (void*)&result);
	return result;
}

void CResource::SetfWorkVariance(FLOAT propval)
{
	SetProperty(0x2d, VT_R4, propval);
}

FLOAT CResource::GetfWorkVariance()
{
	FLOAT result;
	GetProperty(0x2d, VT_R4, (void*)&result);
	return result;
}

void CResource::SetfCostVariance(FLOAT propval)
{
	SetProperty(0x2e, VT_R4, propval);
}

FLOAT CResource::GetfCostVariance()
{
	FLOAT result;
	GetProperty(0x2e, VT_R4, (void*)&result);
	return result;
}

void CResource::SetfSV(FLOAT propval)
{
	SetProperty(0x2f, VT_R4, propval);
}

FLOAT CResource::GetfSV()
{
	FLOAT result;
	GetProperty(0x2f, VT_R4, (void*)&result);
	return result;
}

void CResource::SetfCV(FLOAT propval)
{
	SetProperty(0x30, VT_R4, propval);
}

FLOAT CResource::GetfCV()
{
	FLOAT result;
	GetProperty(0x30, VT_R4, (void*)&result);
	return result;
}

void CResource::SetfACWP(FLOAT propval)
{
	SetProperty(0x31, VT_R4, propval);
}

FLOAT CResource::GetfACWP()
{
	FLOAT result;
	GetProperty(0x31, VT_R4, (void*)&result);
	return result;
}

void CResource::SetlCalendarUID(LONG propval)
{
	SetProperty(0x32, VT_I4, propval);
}

LONG CResource::GetlCalendarUID()
{
	LONG result;
	GetProperty(0x32, VT_I4, (void*)&result);
	return result;
}

void CResource::SetsNotes(LPCTSTR propval)
{
	SetProperty(0x33, VT_BSTR, propval);
}

CString CResource::GetsNotes()
{
	CString result;
	GetProperty(0x33, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetfBCWS(FLOAT propval)
{
	SetProperty(0x34, VT_R4, propval);
}

FLOAT CResource::GetfBCWS()
{
	FLOAT result;
	GetProperty(0x34, VT_R4, (void*)&result);
	return result;
}

void CResource::SetfBCWP(FLOAT propval)
{
	SetProperty(0x35, VT_R4, propval);
}

FLOAT CResource::GetfBCWP()
{
	FLOAT result;
	GetProperty(0x35, VT_R4, (void*)&result);
	return result;
}

void CResource::SetbIsGeneric(BOOL propval)
{
	SetProperty(0x36, VT_BOOL, propval);
}

BOOL CResource::GetbIsGeneric()
{
	BOOL result;
	GetProperty(0x36, VT_BOOL, (void*)&result);
	return result;
}

void CResource::SetbIsInactive(BOOL propval)
{
	SetProperty(0x37, VT_BOOL, propval);
}

BOOL CResource::GetbIsInactive()
{
	BOOL result;
	GetProperty(0x37, VT_BOOL, (void*)&result);
	return result;
}

void CResource::SetbIsEnterprise(BOOL propval)
{
	SetProperty(0x38, VT_BOOL, propval);
}

BOOL CResource::GetbIsEnterprise()
{
	BOOL result;
	GetProperty(0x38, VT_BOOL, (void*)&result);
	return result;
}

void CResource::SetyBookingType(LONG propval)
{
	SetProperty(0x39, VT_I4, propval);
}

LONG CResource::GetyBookingType()
{
	LONG result;
	GetProperty(0x39, VT_I4, (void*)&result);
	return result;
}

CDuration CResource::GetoActualWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3a, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CResource::GetoActualOvertimeWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3b, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CResource::SetsActiveDirectoryGUID(LPCTSTR propval)
{
	SetProperty(0x3c, VT_BSTR, propval);
}

CString CResource::GetsActiveDirectoryGUID()
{
	CString result;
	GetProperty(0x3c, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetdtCreationDate(DATE propval)
{
	SetProperty(0x3d, VT_DATE, propval);
}

DATE CResource::GetdtCreationDate()
{
	DATE result;
	GetProperty(0x3d, VT_DATE, (void*)&result);
	return result;
}

CResourceExtendedAttribute_C CResource::GetoExtendedAttribute_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3e, VT_DISPATCH, (void*)&pDispatch);
	return CResourceExtendedAttribute_C(pDispatch);
}

CResourceBaseline_C CResource::GetoBaseline_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3f, VT_DISPATCH, (void*)&pDispatch);
	return CResourceBaseline_C(pDispatch);
}

CResourceOutlineCode_C CResource::GetoOutlineCode_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x40, VT_DISPATCH, (void*)&pDispatch);
	return CResourceOutlineCode_C(pDispatch);
}

void CResource::SetbIsCostResource(BOOL propval)
{
	SetProperty(0x41, VT_BOOL, propval);
}

BOOL CResource::GetbIsCostResource()
{
	BOOL result;
	GetProperty(0x41, VT_BOOL, (void*)&result);
	return result;
}

void CResource::SetsAssnOwner(LPCTSTR propval)
{
	SetProperty(0x42, VT_BSTR, propval);
}

CString CResource::GetsAssnOwner()
{
	CString result;
	GetProperty(0x42, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetsAssnOwnerGuid(LPCTSTR propval)
{
	SetProperty(0x43, VT_BSTR, propval);
}

CString CResource::GetsAssnOwnerGuid()
{
	CString result;
	GetProperty(0x43, VT_BSTR, (void*)&result);
	return result;
}

void CResource::SetbIsBudget(BOOL propval)
{
	SetProperty(0x44, VT_BOOL, propval);
}

BOOL CResource::GetbIsBudget()
{
	BOOL result;
	GetProperty(0x44, VT_BOOL, (void*)&result);
	return result;
}

CResourceAvailabilityPeriods CResource::GetoAvailabilityPeriods()
{
	LPDISPATCH pDispatch;
	GetProperty(0x45, VT_DISPATCH, (void*)&pDispatch);
	return CResourceAvailabilityPeriods(pDispatch);
}

CResourceRates CResource::GetoRates()
{
	LPDISPATCH pDispatch;
	GetProperty(0x46, VT_DISPATCH, (void*)&pDispatch);
	return CResourceRates(pDispatch);
}

CTimephasedData_C CResource::GetoTimephasedData_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x47, VT_DISPATCH, (void*)&pDispatch);
	return CTimephasedData_C(pDispatch);
}

void CResource::SetKey(LPCTSTR propval)
{
	SetProperty(0x48, VT_BSTR, propval);
}

CString CResource::GetKey()
{
	CString result;
	GetProperty(0x48, VT_BSTR, (void*)&result);
	return result;
}

CString CResource::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x49, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResource::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4a, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CResource::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x4b, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResource::Initialize(void)
{
	InvokeHelper(0x4c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CResources::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CResource CResources::Item(LPCTSTR Index)
{
	CResource oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CResource CResources::Add(void)
{
	CResource oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CResources::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResources::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CResources::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResources::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CResources::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CResources::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskOutlineCode::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CTaskOutlineCode::GetsFieldID()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CTaskOutlineCode::SetlValueID(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CTaskOutlineCode::GetlValueID()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CTaskOutlineCode::SetlValueGUID(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CTaskOutlineCode::GetlValueGUID()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CTaskOutlineCode::SetKey(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CTaskOutlineCode::GetKey()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

CString CTaskOutlineCode::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskOutlineCode::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTaskOutlineCode::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskOutlineCode::Initialize(void)
{
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CTaskOutlineCode_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CTaskOutlineCode CTaskOutlineCode_C::Item(LPCTSTR Index)
{
	CTaskOutlineCode oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CTaskOutlineCode CTaskOutlineCode_C::Add(void)
{
	CTaskOutlineCode oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskOutlineCode_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTaskOutlineCode_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CTaskOutlineCode_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskOutlineCode_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CTimephasedData_C CTaskBaseline::GetoTimephasedData_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1, VT_DISPATCH, (void*)&pDispatch);
	return CTimephasedData_C(pDispatch);
}

void CTaskBaseline::SetlNumber(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CTaskBaseline::GetlNumber()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CTaskBaseline::SetbInterim(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CTaskBaseline::GetbInterim()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CTaskBaseline::SetdtStart(DATE propval)
{
	SetProperty(0x4, VT_DATE, propval);
}

DATE CTaskBaseline::GetdtStart()
{
	DATE result;
	GetProperty(0x4, VT_DATE, (void*)&result);
	return result;
}

void CTaskBaseline::SetdtFinish(DATE propval)
{
	SetProperty(0x5, VT_DATE, propval);
}

DATE CTaskBaseline::GetdtFinish()
{
	DATE result;
	GetProperty(0x5, VT_DATE, (void*)&result);
	return result;
}

CDuration CTaskBaseline::GetoDuration()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTaskBaseline::SetyDurationFormat(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CTaskBaseline::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CTaskBaseline::SetbEstimatedDuration(BOOL propval)
{
	SetProperty(0x8, VT_BOOL, propval);
}

BOOL CTaskBaseline::GetbEstimatedDuration()
{
	BOOL result;
	GetProperty(0x8, VT_BOOL, (void*)&result);
	return result;
}

CDuration CTaskBaseline::GetoWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x9, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTaskBaseline::SetcCost(DOUBLE propval)
{
	SetProperty(0xa, VT_R8, propval);
}

DOUBLE CTaskBaseline::GetcCost()
{
	DOUBLE result;
	GetProperty(0xa, VT_R8, (void*)&result);
	return result;
}

void CTaskBaseline::SetfBCWS(FLOAT propval)
{
	SetProperty(0xb, VT_R4, propval);
}

FLOAT CTaskBaseline::GetfBCWS()
{
	FLOAT result;
	GetProperty(0xb, VT_R4, (void*)&result);
	return result;
}

void CTaskBaseline::SetfBCWP(FLOAT propval)
{
	SetProperty(0xc, VT_R4, propval);
}

FLOAT CTaskBaseline::GetfBCWP()
{
	FLOAT result;
	GetProperty(0xc, VT_R4, (void*)&result);
	return result;
}

void CTaskBaseline::SetfFixedCost(FLOAT propval)
{
	SetProperty(0xd, VT_R4, propval);
}

FLOAT CTaskBaseline::GetfFixedCost()
{
	FLOAT result;
	GetProperty(0xd, VT_R4, (void*)&result);
	return result;
}

void CTaskBaseline::SetKey(LPCTSTR propval)
{
	SetProperty(0xe, VT_BSTR, propval);
}

CString CTaskBaseline::GetKey()
{
	CString result;
	GetProperty(0xe, VT_BSTR, (void*)&result);
	return result;
}

CString CTaskBaseline::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskBaseline::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x10, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTaskBaseline::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x11, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskBaseline::Initialize(void)
{
	InvokeHelper(0x12, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CTaskBaseline_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CTaskBaseline CTaskBaseline_C::Item(LPCTSTR Index)
{
	CTaskBaseline oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CTaskBaseline CTaskBaseline_C::Add(void)
{
	CTaskBaseline oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskBaseline_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTaskBaseline_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CTaskBaseline_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskBaseline_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTaskExtendedAttribute::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CTaskExtendedAttribute::GetsFieldID()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CTaskExtendedAttribute::SetsValue(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CTaskExtendedAttribute::GetsValue()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CTaskExtendedAttribute::SetlValueGUID(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CTaskExtendedAttribute::GetlValueGUID()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CTaskExtendedAttribute::SetyDurationFormat(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CTaskExtendedAttribute::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CTaskExtendedAttribute::SetKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CTaskExtendedAttribute::GetKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CString CTaskExtendedAttribute::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskExtendedAttribute::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTaskExtendedAttribute::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskExtendedAttribute::Initialize(void)
{
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CTaskExtendedAttribute_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CTaskExtendedAttribute CTaskExtendedAttribute_C::Item(LPCTSTR Index)
{
	CTaskExtendedAttribute oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CTaskExtendedAttribute CTaskExtendedAttribute_C::Add(void)
{
	CTaskExtendedAttribute oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskExtendedAttribute_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTaskExtendedAttribute_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CTaskExtendedAttribute_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskExtendedAttribute_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTaskPredecessorLink::SetlPredecessorUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CTaskPredecessorLink::GetlPredecessorUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CTaskPredecessorLink::SetyType(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CTaskPredecessorLink::GetyType()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CTaskPredecessorLink::SetbCrossProject(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CTaskPredecessorLink::GetbCrossProject()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CTaskPredecessorLink::SetsCrossProjectName(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CTaskPredecessorLink::GetsCrossProjectName()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CTaskPredecessorLink::SetlLinkLag(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CTaskPredecessorLink::GetlLinkLag()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CTaskPredecessorLink::SetyLagFormat(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CTaskPredecessorLink::GetyLagFormat()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CTaskPredecessorLink::SetKey(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CTaskPredecessorLink::GetKey()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

CString CTaskPredecessorLink::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskPredecessorLink::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTaskPredecessorLink::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskPredecessorLink::Initialize(void)
{
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CTaskPredecessorLink_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CTaskPredecessorLink CTaskPredecessorLink_C::Item(LPCTSTR Index)
{
	CTaskPredecessorLink oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CTaskPredecessorLink CTaskPredecessorLink_C::Add(void)
{
	CTaskPredecessorLink oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskPredecessorLink_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTaskPredecessorLink_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CTaskPredecessorLink_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskPredecessorLink_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTask::SetlUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CTask::GetlUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlID(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CTask::GetlID()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CTask::SetsName(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CTask::GetsName()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetyType(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CTask::GetyType()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CTask::SetbIsNull(BOOL propval)
{
	SetProperty(0x5, VT_BOOL, propval);
}

BOOL CTask::GetbIsNull()
{
	BOOL result;
	GetProperty(0x5, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetdtCreateDate(DATE propval)
{
	SetProperty(0x6, VT_DATE, propval);
}

DATE CTask::GetdtCreateDate()
{
	DATE result;
	GetProperty(0x6, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetsContact(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CTask::GetsContact()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetsWBS(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CTask::GetsWBS()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetsWBSLevel(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CTask::GetsWBSLevel()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetsOutlineNumber(LPCTSTR propval)
{
	SetProperty(0xa, VT_BSTR, propval);
}

CString CTask::GetsOutlineNumber()
{
	CString result;
	GetProperty(0xa, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetlOutlineLevel(LONG propval)
{
	SetProperty(0xb, VT_I4, propval);
}

LONG CTask::GetlOutlineLevel()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlPriority(LONG propval)
{
	SetProperty(0xc, VT_I4, propval);
}

LONG CTask::GetlPriority()
{
	LONG result;
	GetProperty(0xc, VT_I4, (void*)&result);
	return result;
}

void CTask::SetdtStart(DATE propval)
{
	SetProperty(0xd, VT_DATE, propval);
}

DATE CTask::GetdtStart()
{
	DATE result;
	GetProperty(0xd, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtFinish(DATE propval)
{
	SetProperty(0xe, VT_DATE, propval);
}

DATE CTask::GetdtFinish()
{
	DATE result;
	GetProperty(0xe, VT_DATE, (void*)&result);
	return result;
}

CDuration CTask::GetoDuration()
{
	LPDISPATCH pDispatch;
	GetProperty(0xf, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetyDurationFormat(LONG propval)
{
	SetProperty(0x10, VT_I4, propval);
}

LONG CTask::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

CDuration CTask::GetoWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x11, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetdtStop(DATE propval)
{
	SetProperty(0x12, VT_DATE, propval);
}

DATE CTask::GetdtStop()
{
	DATE result;
	GetProperty(0x12, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtResume(DATE propval)
{
	SetProperty(0x13, VT_DATE, propval);
}

DATE CTask::GetdtResume()
{
	DATE result;
	GetProperty(0x13, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetbResumeValid(BOOL propval)
{
	SetProperty(0x14, VT_BOOL, propval);
}

BOOL CTask::GetbResumeValid()
{
	BOOL result;
	GetProperty(0x14, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbEffortDriven(BOOL propval)
{
	SetProperty(0x15, VT_BOOL, propval);
}

BOOL CTask::GetbEffortDriven()
{
	BOOL result;
	GetProperty(0x15, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbRecurring(BOOL propval)
{
	SetProperty(0x16, VT_BOOL, propval);
}

BOOL CTask::GetbRecurring()
{
	BOOL result;
	GetProperty(0x16, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbOverAllocated(BOOL propval)
{
	SetProperty(0x17, VT_BOOL, propval);
}

BOOL CTask::GetbOverAllocated()
{
	BOOL result;
	GetProperty(0x17, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbEstimated(BOOL propval)
{
	SetProperty(0x18, VT_BOOL, propval);
}

BOOL CTask::GetbEstimated()
{
	BOOL result;
	GetProperty(0x18, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbMilestone(BOOL propval)
{
	SetProperty(0x19, VT_BOOL, propval);
}

BOOL CTask::GetbMilestone()
{
	BOOL result;
	GetProperty(0x19, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbSummary(BOOL propval)
{
	SetProperty(0x1a, VT_BOOL, propval);
}

BOOL CTask::GetbSummary()
{
	BOOL result;
	GetProperty(0x1a, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbDisplayAsSummary(BOOL propval)
{
	SetProperty(0x1b, VT_BOOL, propval);
}

BOOL CTask::GetbDisplayAsSummary()
{
	BOOL result;
	GetProperty(0x1b, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbCritical(BOOL propval)
{
	SetProperty(0x1c, VT_BOOL, propval);
}

BOOL CTask::GetbCritical()
{
	BOOL result;
	GetProperty(0x1c, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbIsSubproject(BOOL propval)
{
	SetProperty(0x1d, VT_BOOL, propval);
}

BOOL CTask::GetbIsSubproject()
{
	BOOL result;
	GetProperty(0x1d, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbIsSubprojectReadOnly(BOOL propval)
{
	SetProperty(0x1e, VT_BOOL, propval);
}

BOOL CTask::GetbIsSubprojectReadOnly()
{
	BOOL result;
	GetProperty(0x1e, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetsSubprojectName(LPCTSTR propval)
{
	SetProperty(0x1f, VT_BSTR, propval);
}

CString CTask::GetsSubprojectName()
{
	CString result;
	GetProperty(0x1f, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetbExternalTask(BOOL propval)
{
	SetProperty(0x20, VT_BOOL, propval);
}

BOOL CTask::GetbExternalTask()
{
	BOOL result;
	GetProperty(0x20, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetsExternalTaskProject(LPCTSTR propval)
{
	SetProperty(0x21, VT_BSTR, propval);
}

CString CTask::GetsExternalTaskProject()
{
	CString result;
	GetProperty(0x21, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetdtEarlyStart(DATE propval)
{
	SetProperty(0x22, VT_DATE, propval);
}

DATE CTask::GetdtEarlyStart()
{
	DATE result;
	GetProperty(0x22, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtEarlyFinish(DATE propval)
{
	SetProperty(0x23, VT_DATE, propval);
}

DATE CTask::GetdtEarlyFinish()
{
	DATE result;
	GetProperty(0x23, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtLateStart(DATE propval)
{
	SetProperty(0x24, VT_DATE, propval);
}

DATE CTask::GetdtLateStart()
{
	DATE result;
	GetProperty(0x24, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtLateFinish(DATE propval)
{
	SetProperty(0x25, VT_DATE, propval);
}

DATE CTask::GetdtLateFinish()
{
	DATE result;
	GetProperty(0x25, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetlStartVariance(LONG propval)
{
	SetProperty(0x26, VT_I4, propval);
}

LONG CTask::GetlStartVariance()
{
	LONG result;
	GetProperty(0x26, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlFinishVariance(LONG propval)
{
	SetProperty(0x27, VT_I4, propval);
}

LONG CTask::GetlFinishVariance()
{
	LONG result;
	GetProperty(0x27, VT_I4, (void*)&result);
	return result;
}

void CTask::SetfWorkVariance(FLOAT propval)
{
	SetProperty(0x28, VT_R4, propval);
}

FLOAT CTask::GetfWorkVariance()
{
	FLOAT result;
	GetProperty(0x28, VT_R4, (void*)&result);
	return result;
}

void CTask::SetlFreeSlack(LONG propval)
{
	SetProperty(0x29, VT_I4, propval);
}

LONG CTask::GetlFreeSlack()
{
	LONG result;
	GetProperty(0x29, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlStartSlack(LONG propval)
{
	SetProperty(0x2a, VT_I4, propval);
}

LONG CTask::GetlStartSlack()
{
	LONG result;
	GetProperty(0x2a, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlFinishSlack(LONG propval)
{
	SetProperty(0x2b, VT_I4, propval);
}

LONG CTask::GetlFinishSlack()
{
	LONG result;
	GetProperty(0x2b, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlTotalSlack(LONG propval)
{
	SetProperty(0x2c, VT_I4, propval);
}

LONG CTask::GetlTotalSlack()
{
	LONG result;
	GetProperty(0x2c, VT_I4, (void*)&result);
	return result;
}

void CTask::SetfFixedCost(FLOAT propval)
{
	SetProperty(0x2d, VT_R4, propval);
}

FLOAT CTask::GetfFixedCost()
{
	FLOAT result;
	GetProperty(0x2d, VT_R4, (void*)&result);
	return result;
}

void CTask::SetyFixedCostAccrual(LONG propval)
{
	SetProperty(0x2e, VT_I4, propval);
}

LONG CTask::GetyFixedCostAccrual()
{
	LONG result;
	GetProperty(0x2e, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlPercentComplete(LONG propval)
{
	SetProperty(0x2f, VT_I4, propval);
}

LONG CTask::GetlPercentComplete()
{
	LONG result;
	GetProperty(0x2f, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlPercentWorkComplete(LONG propval)
{
	SetProperty(0x30, VT_I4, propval);
}

LONG CTask::GetlPercentWorkComplete()
{
	LONG result;
	GetProperty(0x30, VT_I4, (void*)&result);
	return result;
}

void CTask::SetcCost(DOUBLE propval)
{
	SetProperty(0x31, VT_R8, propval);
}

DOUBLE CTask::GetcCost()
{
	DOUBLE result;
	GetProperty(0x31, VT_R8, (void*)&result);
	return result;
}

void CTask::SetcOvertimeCost(DOUBLE propval)
{
	SetProperty(0x32, VT_R8, propval);
}

DOUBLE CTask::GetcOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x32, VT_R8, (void*)&result);
	return result;
}

CDuration CTask::GetoOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x33, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetdtActualStart(DATE propval)
{
	SetProperty(0x34, VT_DATE, propval);
}

DATE CTask::GetdtActualStart()
{
	DATE result;
	GetProperty(0x34, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtActualFinish(DATE propval)
{
	SetProperty(0x35, VT_DATE, propval);
}

DATE CTask::GetdtActualFinish()
{
	DATE result;
	GetProperty(0x35, VT_DATE, (void*)&result);
	return result;
}

CDuration CTask::GetoActualDuration()
{
	LPDISPATCH pDispatch;
	GetProperty(0x36, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetcActualCost(DOUBLE propval)
{
	SetProperty(0x37, VT_R8, propval);
}

DOUBLE CTask::GetcActualCost()
{
	DOUBLE result;
	GetProperty(0x37, VT_R8, (void*)&result);
	return result;
}

void CTask::SetcActualOvertimeCost(DOUBLE propval)
{
	SetProperty(0x38, VT_R8, propval);
}

DOUBLE CTask::GetcActualOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x38, VT_R8, (void*)&result);
	return result;
}

CDuration CTask::GetoActualWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x39, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CTask::GetoActualOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3a, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CTask::GetoRegularWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3b, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CTask::GetoRemainingDuration()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3c, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetcRemainingCost(DOUBLE propval)
{
	SetProperty(0x3d, VT_R8, propval);
}

DOUBLE CTask::GetcRemainingCost()
{
	DOUBLE result;
	GetProperty(0x3d, VT_R8, (void*)&result);
	return result;
}

CDuration CTask::GetoRemainingWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3e, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetcRemainingOvertimeCost(DOUBLE propval)
{
	SetProperty(0x3f, VT_R8, propval);
}

DOUBLE CTask::GetcRemainingOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x3f, VT_R8, (void*)&result);
	return result;
}

CDuration CTask::GetoRemainingOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x40, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetfACWP(FLOAT propval)
{
	SetProperty(0x41, VT_R4, propval);
}

FLOAT CTask::GetfACWP()
{
	FLOAT result;
	GetProperty(0x41, VT_R4, (void*)&result);
	return result;
}

void CTask::SetfCV(FLOAT propval)
{
	SetProperty(0x42, VT_R4, propval);
}

FLOAT CTask::GetfCV()
{
	FLOAT result;
	GetProperty(0x42, VT_R4, (void*)&result);
	return result;
}

void CTask::SetyConstraintType(LONG propval)
{
	SetProperty(0x43, VT_I4, propval);
}

LONG CTask::GetyConstraintType()
{
	LONG result;
	GetProperty(0x43, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlCalendarUID(LONG propval)
{
	SetProperty(0x44, VT_I4, propval);
}

LONG CTask::GetlCalendarUID()
{
	LONG result;
	GetProperty(0x44, VT_I4, (void*)&result);
	return result;
}

void CTask::SetdtConstraintDate(DATE propval)
{
	SetProperty(0x45, VT_DATE, propval);
}

DATE CTask::GetdtConstraintDate()
{
	DATE result;
	GetProperty(0x45, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtDeadline(DATE propval)
{
	SetProperty(0x46, VT_DATE, propval);
}

DATE CTask::GetdtDeadline()
{
	DATE result;
	GetProperty(0x46, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetbLevelAssignments(BOOL propval)
{
	SetProperty(0x47, VT_BOOL, propval);
}

BOOL CTask::GetbLevelAssignments()
{
	BOOL result;
	GetProperty(0x47, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbLevelingCanSplit(BOOL propval)
{
	SetProperty(0x48, VT_BOOL, propval);
}

BOOL CTask::GetbLevelingCanSplit()
{
	BOOL result;
	GetProperty(0x48, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetlLevelingDelay(LONG propval)
{
	SetProperty(0x49, VT_I4, propval);
}

LONG CTask::GetlLevelingDelay()
{
	LONG result;
	GetProperty(0x49, VT_I4, (void*)&result);
	return result;
}

void CTask::SetyLevelingDelayFormat(LONG propval)
{
	SetProperty(0x4a, VT_I4, propval);
}

LONG CTask::GetyLevelingDelayFormat()
{
	LONG result;
	GetProperty(0x4a, VT_I4, (void*)&result);
	return result;
}

void CTask::SetdtPreLeveledStart(DATE propval)
{
	SetProperty(0x4b, VT_DATE, propval);
}

DATE CTask::GetdtPreLeveledStart()
{
	DATE result;
	GetProperty(0x4b, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtPreLeveledFinish(DATE propval)
{
	SetProperty(0x4c, VT_DATE, propval);
}

DATE CTask::GetdtPreLeveledFinish()
{
	DATE result;
	GetProperty(0x4c, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetsHyperlink(LPCTSTR propval)
{
	SetProperty(0x4d, VT_BSTR, propval);
}

CString CTask::GetsHyperlink()
{
	CString result;
	GetProperty(0x4d, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetsHyperlinkAddress(LPCTSTR propval)
{
	SetProperty(0x4e, VT_BSTR, propval);
}

CString CTask::GetsHyperlinkAddress()
{
	CString result;
	GetProperty(0x4e, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetsHyperlinkSubAddress(LPCTSTR propval)
{
	SetProperty(0x4f, VT_BSTR, propval);
}

CString CTask::GetsHyperlinkSubAddress()
{
	CString result;
	GetProperty(0x4f, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetbIgnoreResourceCalendar(BOOL propval)
{
	SetProperty(0x50, VT_BOOL, propval);
}

BOOL CTask::GetbIgnoreResourceCalendar()
{
	BOOL result;
	GetProperty(0x50, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetsNotes(LPCTSTR propval)
{
	SetProperty(0x51, VT_BSTR, propval);
}

CString CTask::GetsNotes()
{
	CString result;
	GetProperty(0x51, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetbHideBar(BOOL propval)
{
	SetProperty(0x52, VT_BOOL, propval);
}

BOOL CTask::GetbHideBar()
{
	BOOL result;
	GetProperty(0x52, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbRollup(BOOL propval)
{
	SetProperty(0x53, VT_BOOL, propval);
}

BOOL CTask::GetbRollup()
{
	BOOL result;
	GetProperty(0x53, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetfBCWS(FLOAT propval)
{
	SetProperty(0x54, VT_R4, propval);
}

FLOAT CTask::GetfBCWS()
{
	FLOAT result;
	GetProperty(0x54, VT_R4, (void*)&result);
	return result;
}

void CTask::SetfBCWP(FLOAT propval)
{
	SetProperty(0x55, VT_R4, propval);
}

FLOAT CTask::GetfBCWP()
{
	FLOAT result;
	GetProperty(0x55, VT_R4, (void*)&result);
	return result;
}

void CTask::SetlPhysicalPercentComplete(LONG propval)
{
	SetProperty(0x56, VT_I4, propval);
}

LONG CTask::GetlPhysicalPercentComplete()
{
	LONG result;
	GetProperty(0x56, VT_I4, (void*)&result);
	return result;
}

void CTask::SetyEarnedValueMethod(LONG propval)
{
	SetProperty(0x57, VT_I4, propval);
}

LONG CTask::GetyEarnedValueMethod()
{
	LONG result;
	GetProperty(0x57, VT_I4, (void*)&result);
	return result;
}

CTaskPredecessorLink_C CTask::GetoPredecessorLink_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x58, VT_DISPATCH, (void*)&pDispatch);
	return CTaskPredecessorLink_C(pDispatch);
}

CDuration CTask::GetoActualWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x59, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CTask::GetoActualOvertimeWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5a, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CTaskExtendedAttribute_C CTask::GetoExtendedAttribute_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5b, VT_DISPATCH, (void*)&pDispatch);
	return CTaskExtendedAttribute_C(pDispatch);
}

CTaskBaseline_C CTask::GetoBaseline_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5c, VT_DISPATCH, (void*)&pDispatch);
	return CTaskBaseline_C(pDispatch);
}

CTaskOutlineCode_C CTask::GetoOutlineCode_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5d, VT_DISPATCH, (void*)&pDispatch);
	return CTaskOutlineCode_C(pDispatch);
}

void CTask::SetbIsPublished(BOOL propval)
{
	SetProperty(0x5e, VT_BOOL, propval);
}

BOOL CTask::GetbIsPublished()
{
	BOOL result;
	GetProperty(0x5e, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetsStatusManager(LPCTSTR propval)
{
	SetProperty(0x5f, VT_BSTR, propval);
}

CString CTask::GetsStatusManager()
{
	CString result;
	GetProperty(0x5f, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetdtCommitmentStart(DATE propval)
{
	SetProperty(0x60, VT_DATE, propval);
}

DATE CTask::GetdtCommitmentStart()
{
	DATE result;
	GetProperty(0x60, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtCommitmentFinish(DATE propval)
{
	SetProperty(0x61, VT_DATE, propval);
}

DATE CTask::GetdtCommitmentFinish()
{
	DATE result;
	GetProperty(0x61, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetyCommitmentType(LONG propval)
{
	SetProperty(0x62, VT_I4, propval);
}

LONG CTask::GetyCommitmentType()
{
	LONG result;
	GetProperty(0x62, VT_I4, (void*)&result);
	return result;
}

void CTask::SetbActive(BOOL propval)
{
	SetProperty(0x63, VT_BOOL, propval);
}

BOOL CTask::GetbActive()
{
	BOOL result;
	GetProperty(0x63, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbPinned(BOOL propval)
{
	SetProperty(0x64, VT_BOOL, propval);
}

BOOL CTask::GetbPinned()
{
	BOOL result;
	GetProperty(0x64, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetsPinnedStart(LPCTSTR propval)
{
	SetProperty(0x65, VT_BSTR, propval);
}

CString CTask::GetsPinnedStart()
{
	CString result;
	GetProperty(0x65, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetsPinnedFinish(LPCTSTR propval)
{
	SetProperty(0x66, VT_BSTR, propval);
}

CString CTask::GetsPinnedFinish()
{
	CString result;
	GetProperty(0x66, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetsPinnedDuration(LPCTSTR propval)
{
	SetProperty(0x67, VT_BSTR, propval);
}

CString CTask::GetsPinnedDuration()
{
	CString result;
	GetProperty(0x67, VT_BSTR, (void*)&result);
	return result;
}

CTimephasedData_C CTask::GetoTimephasedData_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x68, VT_DISPATCH, (void*)&pDispatch);
	return CTimephasedData_C(pDispatch);
}

void CTask::SetKey(LPCTSTR propval)
{
	SetProperty(0x69, VT_BSTR, propval);
}

CString CTask::GetKey()
{
	CString result;
	GetProperty(0x69, VT_BSTR, (void*)&result);
	return result;
}

CString CTask::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6a, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTask::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6b, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTask::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6c, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTask::Initialize(void)
{
	InvokeHelper(0x6d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CTasks::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CTask CTasks::Item(LPCTSTR Index)
{
	CTask oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CTask CTasks::Add(void)
{
	CTask oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CTasks::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTasks::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CTasks::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTasks::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTasks::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CTasks::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

CTime CWorkingTime::GetoFromTime()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1, VT_DISPATCH, (void*)&pDispatch);
	return CTime(pDispatch);
}

CTime CWorkingTime::GetoToTime()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CTime(pDispatch);
}

void CWorkingTime::SetKey(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CWorkingTime::GetKey()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

CString CWorkingTime::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CWorkingTime::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CWorkingTime::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CWorkingTime::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CWorkingTimes::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CWorkingTime CWorkingTimes::Item(LPCTSTR Index)
{
	CWorkingTime oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CWorkingTime CWorkingTimes::Add(void)
{
	CWorkingTime oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CWorkingTimes::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CWorkingTimes::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CWorkingTimes::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CWorkingTimes::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CWorkingTimes::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CWorkingTimes::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTimePeriod::SetdtFromDate(DATE propval)
{
	SetProperty(0x1, VT_DATE, propval);
}

DATE CTimePeriod::GetdtFromDate()
{
	DATE result;
	GetProperty(0x1, VT_DATE, (void*)&result);
	return result;
}

void CTimePeriod::SetdtToDate(DATE propval)
{
	SetProperty(0x2, VT_DATE, propval);
}

DATE CTimePeriod::GetdtToDate()
{
	DATE result;
	GetProperty(0x2, VT_DATE, (void*)&result);
	return result;
}

void CTimePeriod::SetKey(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CTimePeriod::GetKey()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

CString CTimePeriod::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTimePeriod::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTimePeriod::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTimePeriod::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CTimePeriods::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CTimePeriod CTimePeriods::Item(LPCTSTR Index)
{
	CTimePeriod oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CTimePeriod CTimePeriods::Add(void)
{
	CTimePeriod oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CTimePeriods::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTimePeriods::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CTimePeriods::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTimePeriods::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CTimePeriods::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CTimePeriods::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CWeekDay::SetyDayType(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CWeekDay::GetyDayType()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CWeekDay::SetbDayWorking(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CWeekDay::GetbDayWorking()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

void CWeekDay::SetKey(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CWeekDay::GetKey()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

CString CWeekDay::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x4, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CWeekDay::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CWeekDay::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CWeekDay::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CWeekDay_C::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CWeekDay CWeekDay_C::Item(LPCTSTR Index)
{
	CWeekDay oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CWeekDay CWeekDay_C::Add(void)
{
	CWeekDay oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CWeekDay_C::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CWeekDay_C::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CWeekDay_C::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CWeekDay_C::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CTimePeriod CCalendarWorkWeek::GetoTimePeriod()
{
	LPDISPATCH pDispatch;
	GetProperty(0x1, VT_DISPATCH, (void*)&pDispatch);
	return CTimePeriod(pDispatch);
}

void CCalendarWorkWeek::SetsName(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CCalendarWorkWeek::GetsName()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

CWeekDay_C CCalendarWorkWeek::GetoWeekDay_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3, VT_DISPATCH, (void*)&pDispatch);
	return CWeekDay_C(pDispatch);
}

void CCalendarWorkWeek::SetKey(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CCalendarWorkWeek::GetKey()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

CString CCalendarWorkWeek::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarWorkWeek::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CCalendarWorkWeek::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarWorkWeek::Initialize(void)
{
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CCalendarWorkWeeks::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CCalendarWorkWeek CCalendarWorkWeeks::Item(LPCTSTR Index)
{
	CCalendarWorkWeek oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CCalendarWorkWeek CCalendarWorkWeeks::Add(void)
{
	CCalendarWorkWeek oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarWorkWeeks::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCalendarWorkWeeks::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CCalendarWorkWeeks::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarWorkWeeks::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCalendarWorkWeeks::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CCalendarWorkWeeks::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarException::SetbEnteredByOccurrences(BOOL propval)
{
	SetProperty(0x1, VT_BOOL, propval);
}

BOOL CCalendarException::GetbEnteredByOccurrences()
{
	BOOL result;
	GetProperty(0x1, VT_BOOL, (void*)&result);
	return result;
}

CTimePeriod CCalendarException::GetoTimePeriod()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CTimePeriod(pDispatch);
}

void CCalendarException::SetlOccurrences(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CCalendarException::GetlOccurrences()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CCalendarException::SetsName(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CCalendarException::GetsName()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CCalendarException::SetyType(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CCalendarException::GetyType()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CCalendarException::SetlPeriod(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CCalendarException::GetlPeriod()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CCalendarException::SetlDaysOfWeek(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG CCalendarException::GetlDaysOfWeek()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void CCalendarException::SetyMonthItem(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG CCalendarException::GetyMonthItem()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void CCalendarException::SetyMonthPosition(LONG propval)
{
	SetProperty(0x9, VT_I4, propval);
}

LONG CCalendarException::GetyMonthPosition()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

void CCalendarException::SetyMonth(LONG propval)
{
	SetProperty(0xa, VT_I4, propval);
}

LONG CCalendarException::GetyMonth()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

void CCalendarException::SetlMonthDay(LONG propval)
{
	SetProperty(0xb, VT_I4, propval);
}

LONG CCalendarException::GetlMonthDay()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

void CCalendarException::SetbDayWorking(BOOL propval)
{
	SetProperty(0xc, VT_BOOL, propval);
}

BOOL CCalendarException::GetbDayWorking()
{
	BOOL result;
	GetProperty(0xc, VT_BOOL, (void*)&result);
	return result;
}

CWorkingTimes CCalendarException::GetoWorkingTimes()
{
	LPDISPATCH pDispatch;
	GetProperty(0xd, VT_DISPATCH, (void*)&pDispatch);
	return CWorkingTimes(pDispatch);
}

void CCalendarException::SetKey(LPCTSTR propval)
{
	SetProperty(0xe, VT_BSTR, propval);
}

CString CCalendarException::GetKey()
{
	CString result;
	GetProperty(0xe, VT_BSTR, (void*)&result);
	return result;
}

CString CCalendarException::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarException::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x10, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CCalendarException::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x11, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarException::Initialize(void)
{
	InvokeHelper(0x12, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CCalendarExceptions::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CCalendarException CCalendarExceptions::Item(LPCTSTR Index)
{
	CCalendarException oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CCalendarException CCalendarExceptions::Add(void)
{
	CCalendarException oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarExceptions::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCalendarExceptions::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CCalendarExceptions::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarExceptions::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCalendarExceptions::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CCalendarExceptions::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarWeekDay::SetyDayType(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CCalendarWeekDay::GetyDayType()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CCalendarWeekDay::SetbDayWorking(BOOL propval)
{
	SetProperty(0x2, VT_BOOL, propval);
}

BOOL CCalendarWeekDay::GetbDayWorking()
{
	BOOL result;
	GetProperty(0x2, VT_BOOL, (void*)&result);
	return result;
}

CTimePeriod CCalendarWeekDay::GetoTimePeriod()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3, VT_DISPATCH, (void*)&pDispatch);
	return CTimePeriod(pDispatch);
}

CWorkingTimes CCalendarWeekDay::GetoWorkingTimes()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CWorkingTimes(pDispatch);
}

void CCalendarWeekDay::SetKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CCalendarWeekDay::GetKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CString CCalendarWeekDay::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarWeekDay::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CCalendarWeekDay::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarWeekDay::Initialize(void)
{
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CCalendarWeekDays::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CCalendarWeekDay CCalendarWeekDays::Item(LPCTSTR Index)
{
	CCalendarWeekDay oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CCalendarWeekDay CCalendarWeekDays::Add(void)
{
	CCalendarWeekDay oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarWeekDays::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCalendarWeekDays::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CCalendarWeekDays::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendarWeekDays::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCalendarWeekDays::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CCalendarWeekDays::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendar::SetlUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CCalendar::GetlUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CCalendar::SetsName(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CCalendar::GetsName()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CCalendar::SetbIsBaseCalendar(BOOL propval)
{
	SetProperty(0x3, VT_BOOL, propval);
}

BOOL CCalendar::GetbIsBaseCalendar()
{
	BOOL result;
	GetProperty(0x3, VT_BOOL, (void*)&result);
	return result;
}

void CCalendar::SetbIsBaselineCalendar(BOOL propval)
{
	SetProperty(0x4, VT_BOOL, propval);
}

BOOL CCalendar::GetbIsBaselineCalendar()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

void CCalendar::SetlBaseCalendarUID(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CCalendar::GetlBaseCalendarUID()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

CCalendarWeekDays CCalendar::GetoWeekDays()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return CCalendarWeekDays(pDispatch);
}

CCalendarExceptions CCalendar::GetoExceptions()
{
	LPDISPATCH pDispatch;
	GetProperty(0x7, VT_DISPATCH, (void*)&pDispatch);
	return CCalendarExceptions(pDispatch);
}

CCalendarWorkWeeks CCalendar::GetoWorkWeeks()
{
	LPDISPATCH pDispatch;
	GetProperty(0x8, VT_DISPATCH, (void*)&pDispatch);
	return CCalendarWorkWeeks(pDispatch);
}

void CCalendar::SetKey(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CCalendar::GetKey()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

CString CCalendar::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendar::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CCalendar::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0xc, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendar::Initialize(void)
{
	InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CCalendars::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CCalendar CCalendars::Item(LPCTSTR Index)
{
	CCalendar oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CCalendar CCalendars::Add(void)
{
	CCalendar oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendars::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCalendars::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CCalendars::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendars::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCalendars::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CCalendars::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttributeValue::SetlID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CExtendedAttributeValue::GetlID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttributeValue::SetsValue(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CExtendedAttributeValue::GetsValue()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttributeValue::SetsDescription(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CExtendedAttributeValue::GetsDescription()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttributeValue::SetsPhonetic(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CExtendedAttributeValue::GetsPhonetic()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttributeValue::SetKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CExtendedAttributeValue::GetKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CString CExtendedAttributeValue::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttributeValue::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CExtendedAttributeValue::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttributeValue::Initialize(void)
{
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CExtendedAttributeValueList::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CExtendedAttributeValue CExtendedAttributeValueList::Item(LPCTSTR Index)
{
	CExtendedAttributeValue oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CExtendedAttributeValue CExtendedAttributeValueList::Add(void)
{
	CExtendedAttributeValue oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttributeValueList::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CExtendedAttributeValueList::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CExtendedAttributeValueList::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttributeValueList::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CExtendedAttributeValueList::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CExtendedAttributeValueList::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttribute::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsFieldID()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsFieldName(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsFieldName()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetyCFType(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CExtendedAttribute::GetyCFType()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsGuid(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsGuid()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetyElemType(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CExtendedAttribute::GetyElemType()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttribute::SetlMaxMultiValues(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CExtendedAttribute::GetlMaxMultiValues()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttribute::SetbUserDef(BOOL propval)
{
	SetProperty(0x7, VT_BOOL, propval);
}

BOOL CExtendedAttribute::GetbUserDef()
{
	BOOL result;
	GetProperty(0x7, VT_BOOL, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsAlias(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsAlias()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsSecondaryPID(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsSecondaryPID()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetbAutoRollDown(BOOL propval)
{
	SetProperty(0xa, VT_BOOL, propval);
}

BOOL CExtendedAttribute::GetbAutoRollDown()
{
	BOOL result;
	GetProperty(0xa, VT_BOOL, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsDefaultGuid(LPCTSTR propval)
{
	SetProperty(0xb, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsDefaultGuid()
{
	CString result;
	GetProperty(0xb, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsLtuid(LPCTSTR propval)
{
	SetProperty(0xc, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsLtuid()
{
	CString result;
	GetProperty(0xc, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsSecondaryGuid(LPCTSTR propval)
{
	SetProperty(0xd, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsSecondaryGuid()
{
	CString result;
	GetProperty(0xd, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsPhoneticAlias(LPCTSTR propval)
{
	SetProperty(0xe, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsPhoneticAlias()
{
	CString result;
	GetProperty(0xe, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetyRollupType(LONG propval)
{
	SetProperty(0xf, VT_I4, propval);
}

LONG CExtendedAttribute::GetyRollupType()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttribute::SetyCalculationType(LONG propval)
{
	SetProperty(0x10, VT_I4, propval);
}

LONG CExtendedAttribute::GetyCalculationType()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsFormula(LPCTSTR propval)
{
	SetProperty(0x11, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsFormula()
{
	CString result;
	GetProperty(0x11, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetbRestrictValues(BOOL propval)
{
	SetProperty(0x12, VT_BOOL, propval);
}

BOOL CExtendedAttribute::GetbRestrictValues()
{
	BOOL result;
	GetProperty(0x12, VT_BOOL, (void*)&result);
	return result;
}

void CExtendedAttribute::SetyValuelistSortOrder(LONG propval)
{
	SetProperty(0x13, VT_I4, propval);
}

LONG CExtendedAttribute::GetyValuelistSortOrder()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttribute::SetbAppendNewValues(BOOL propval)
{
	SetProperty(0x14, VT_BOOL, propval);
}

BOOL CExtendedAttribute::GetbAppendNewValues()
{
	BOOL result;
	GetProperty(0x14, VT_BOOL, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsDefault(LPCTSTR propval)
{
	SetProperty(0x15, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsDefault()
{
	CString result;
	GetProperty(0x15, VT_BSTR, (void*)&result);
	return result;
}

CExtendedAttributeValueList CExtendedAttribute::GetoValueList()
{
	LPDISPATCH pDispatch;
	GetProperty(0x16, VT_DISPATCH, (void*)&pDispatch);
	return CExtendedAttributeValueList(pDispatch);
}

void CExtendedAttribute::SetKey(LPCTSTR propval)
{
	SetProperty(0x17, VT_BSTR, propval);
}

CString CExtendedAttribute::GetKey()
{
	CString result;
	GetProperty(0x17, VT_BSTR, (void*)&result);
	return result;
}

CString CExtendedAttribute::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x18, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttribute::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x19, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CExtendedAttribute::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x1a, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttribute::Initialize(void)
{
	InvokeHelper(0x1b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CExtendedAttributes::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CExtendedAttribute CExtendedAttributes::Item(LPCTSTR Index)
{
	CExtendedAttribute oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CExtendedAttribute CExtendedAttributes::Add(void)
{
	CExtendedAttribute oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttributes::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CExtendedAttributes::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CExtendedAttributes::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttributes::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CExtendedAttributes::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CExtendedAttributes::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CWBSMask::SetlLevel(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CWBSMask::GetlLevel()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CWBSMask::SetyType(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG CWBSMask::GetyType()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void CWBSMask::SetsLength(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CWBSMask::GetsLength()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CWBSMask::SetsSeparator(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CWBSMask::GetsSeparator()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CWBSMask::SetKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CWBSMask::GetKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CString CWBSMask::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CWBSMask::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CWBSMask::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CWBSMask::Initialize(void)
{
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CWBSMasks::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CWBSMask CWBSMasks::Item(LPCTSTR Index)
{
	CWBSMask oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CWBSMask CWBSMasks::Add(void)
{
	CWBSMask oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CWBSMasks::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CWBSMasks::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CWBSMasks::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CWBSMasks::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CWBSMasks::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CWBSMasks::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeMask::SetlLevel(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG COutlineCodeMask::GetlLevel()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void COutlineCodeMask::SetyType(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG COutlineCodeMask::GetyType()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void COutlineCodeMask::SetlLength(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG COutlineCodeMask::GetlLength()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void COutlineCodeMask::SetsSeparator(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString COutlineCodeMask::GetsSeparator()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCodeMask::SetKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString COutlineCodeMask::GetKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CString COutlineCodeMask::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeMask::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL COutlineCodeMask::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeMask::Initialize(void)
{
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG COutlineCodeMasks::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

COutlineCodeMask COutlineCodeMasks::Item(LPCTSTR Index)
{
	COutlineCodeMask oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

COutlineCodeMask COutlineCodeMasks::Add(void)
{
	COutlineCodeMask oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeMasks::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void COutlineCodeMasks::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL COutlineCodeMasks::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeMasks::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void COutlineCodeMasks::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString COutlineCodeMasks::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeValue::SetlValueID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG COutlineCodeValue::GetlValueID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void COutlineCodeValue::SetsFieldGUID(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString COutlineCodeValue::GetsFieldGUID()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCodeValue::SetyType(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG COutlineCodeValue::GetyType()
{
	LONG result;
	GetProperty(0x3, VT_I4, (void*)&result);
	return result;
}

void COutlineCodeValue::SetbIsCollapsed(BOOL propval)
{
	SetProperty(0x4, VT_BOOL, propval);
}

BOOL COutlineCodeValue::GetbIsCollapsed()
{
	BOOL result;
	GetProperty(0x4, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCodeValue::SetlParentValueID(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG COutlineCodeValue::GetlParentValueID()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void COutlineCodeValue::SetsValue(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString COutlineCodeValue::GetsValue()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCodeValue::SetsDescription(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString COutlineCodeValue::GetsDescription()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCodeValue::SetKey(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString COutlineCodeValue::GetKey()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

CString COutlineCodeValue::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeValue::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL COutlineCodeValue::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0xb, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeValue::Initialize(void)
{
	InvokeHelper(0xc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG COutlineCodeValues::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

COutlineCodeValue COutlineCodeValues::Item(LPCTSTR Index)
{
	COutlineCodeValue oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

COutlineCodeValue COutlineCodeValues::Add(void)
{
	COutlineCodeValue oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeValues::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void COutlineCodeValues::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL COutlineCodeValues::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeValues::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void COutlineCodeValues::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString COutlineCodeValues::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCode::SetsGuid(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString COutlineCode::GetsGuid()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCode::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString COutlineCode::GetsFieldID()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCode::SetsFieldName(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString COutlineCode::GetsFieldName()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCode::SetsAlias(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString COutlineCode::GetsAlias()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCode::SetsPhoneticAlias(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString COutlineCode::GetsPhoneticAlias()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

COutlineCodeValues COutlineCode::GetoValues()
{
	LPDISPATCH pDispatch;
	GetProperty(0x6, VT_DISPATCH, (void*)&pDispatch);
	return COutlineCodeValues(pDispatch);
}

void COutlineCode::SetbEnterprise(BOOL propval)
{
	SetProperty(0x7, VT_BOOL, propval);
}

BOOL COutlineCode::GetbEnterprise()
{
	BOOL result;
	GetProperty(0x7, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCode::SetlEnterpriseOutlineCodeAlias(LONG propval)
{
	SetProperty(0x8, VT_I4, propval);
}

LONG COutlineCode::GetlEnterpriseOutlineCodeAlias()
{
	LONG result;
	GetProperty(0x8, VT_I4, (void*)&result);
	return result;
}

void COutlineCode::SetbResourceSubstitutionEnabled(BOOL propval)
{
	SetProperty(0x9, VT_BOOL, propval);
}

BOOL COutlineCode::GetbResourceSubstitutionEnabled()
{
	BOOL result;
	GetProperty(0x9, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCode::SetbLeafOnly(BOOL propval)
{
	SetProperty(0xa, VT_BOOL, propval);
}

BOOL COutlineCode::GetbLeafOnly()
{
	BOOL result;
	GetProperty(0xa, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCode::SetbAllLevelsRequired(BOOL propval)
{
	SetProperty(0xb, VT_BOOL, propval);
}

BOOL COutlineCode::GetbAllLevelsRequired()
{
	BOOL result;
	GetProperty(0xb, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCode::SetbOnlyTableValuesAllowed(BOOL propval)
{
	SetProperty(0xc, VT_BOOL, propval);
}

BOOL COutlineCode::GetbOnlyTableValuesAllowed()
{
	BOOL result;
	GetProperty(0xc, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCode::SetbShowIndent(BOOL propval)
{
	SetProperty(0xd, VT_BOOL, propval);
}

BOOL COutlineCode::GetbShowIndent()
{
	BOOL result;
	GetProperty(0xd, VT_BOOL, (void*)&result);
	return result;
}

COutlineCodeMasks COutlineCode::GetoMasks()
{
	LPDISPATCH pDispatch;
	GetProperty(0xe, VT_DISPATCH, (void*)&pDispatch);
	return COutlineCodeMasks(pDispatch);
}

void COutlineCode::SetKey(LPCTSTR propval)
{
	SetProperty(0xf, VT_BSTR, propval);
}

CString COutlineCode::GetKey()
{
	CString result;
	GetProperty(0xf, VT_BSTR, (void*)&result);
	return result;
}

CString COutlineCode::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x10, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCode::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x11, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL COutlineCode::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x12, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCode::Initialize(void)
{
	InvokeHelper(0x13, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG COutlineCodes::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

COutlineCode COutlineCodes::Item(LPCTSTR Index)
{
	COutlineCode oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

COutlineCode COutlineCodes::Add(void)
{
	COutlineCode oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodes::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void COutlineCodes::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL COutlineCodes::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodes::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void COutlineCodes::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString COutlineCodes::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CMP14::SetlSaveVersion(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CMP14::GetlSaveVersion()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetsUID(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CMP14::GetsUID()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetsName(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CMP14::GetsName()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetsTitle(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CMP14::GetsTitle()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetsSubject(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CMP14::GetsSubject()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetsCategory(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CMP14::GetsCategory()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetsCompany(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CMP14::GetsCompany()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetsManager(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CMP14::GetsManager()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetsAuthor(LPCTSTR propval)
{
	SetProperty(0x9, VT_BSTR, propval);
}

CString CMP14::GetsAuthor()
{
	CString result;
	GetProperty(0x9, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetdtCreationDate(DATE propval)
{
	SetProperty(0xa, VT_DATE, propval);
}

DATE CMP14::GetdtCreationDate()
{
	DATE result;
	GetProperty(0xa, VT_DATE, (void*)&result);
	return result;
}

void CMP14::SetlRevision(LONG propval)
{
	SetProperty(0xb, VT_I4, propval);
}

LONG CMP14::GetlRevision()
{
	LONG result;
	GetProperty(0xb, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetdtLastSaved(DATE propval)
{
	SetProperty(0xc, VT_DATE, propval);
}

DATE CMP14::GetdtLastSaved()
{
	DATE result;
	GetProperty(0xc, VT_DATE, (void*)&result);
	return result;
}

void CMP14::SetbScheduleFromStart(BOOL propval)
{
	SetProperty(0xd, VT_BOOL, propval);
}

BOOL CMP14::GetbScheduleFromStart()
{
	BOOL result;
	GetProperty(0xd, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetdtStartDate(DATE propval)
{
	SetProperty(0xe, VT_DATE, propval);
}

DATE CMP14::GetdtStartDate()
{
	DATE result;
	GetProperty(0xe, VT_DATE, (void*)&result);
	return result;
}

void CMP14::SetdtFinishDate(DATE propval)
{
	SetProperty(0xf, VT_DATE, propval);
}

DATE CMP14::GetdtFinishDate()
{
	DATE result;
	GetProperty(0xf, VT_DATE, (void*)&result);
	return result;
}

void CMP14::SetyFYStartDate(LONG propval)
{
	SetProperty(0x10, VT_I4, propval);
}

LONG CMP14::GetyFYStartDate()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetlCriticalSlackLimit(LONG propval)
{
	SetProperty(0x11, VT_I4, propval);
}

LONG CMP14::GetlCriticalSlackLimit()
{
	LONG result;
	GetProperty(0x11, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetlCurrencyDigits(LONG propval)
{
	SetProperty(0x12, VT_I4, propval);
}

LONG CMP14::GetlCurrencyDigits()
{
	LONG result;
	GetProperty(0x12, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetsCurrencySymbol(LPCTSTR propval)
{
	SetProperty(0x13, VT_BSTR, propval);
}

CString CMP14::GetsCurrencySymbol()
{
	CString result;
	GetProperty(0x13, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetsCurrencyCode(LPCTSTR propval)
{
	SetProperty(0x14, VT_BSTR, propval);
}

CString CMP14::GetsCurrencyCode()
{
	CString result;
	GetProperty(0x14, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetyCurrencySymbolPosition(LONG propval)
{
	SetProperty(0x15, VT_I4, propval);
}

LONG CMP14::GetyCurrencySymbolPosition()
{
	LONG result;
	GetProperty(0x15, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetlCalendarUID(LONG propval)
{
	SetProperty(0x16, VT_I4, propval);
}

LONG CMP14::GetlCalendarUID()
{
	LONG result;
	GetProperty(0x16, VT_I4, (void*)&result);
	return result;
}

CTime CMP14::GetoDefaultStartTime()
{
	LPDISPATCH pDispatch;
	GetProperty(0x17, VT_DISPATCH, (void*)&pDispatch);
	return CTime(pDispatch);
}

CTime CMP14::GetoDefaultFinishTime()
{
	LPDISPATCH pDispatch;
	GetProperty(0x18, VT_DISPATCH, (void*)&pDispatch);
	return CTime(pDispatch);
}

void CMP14::SetlMinutesPerDay(LONG propval)
{
	SetProperty(0x19, VT_I4, propval);
}

LONG CMP14::GetlMinutesPerDay()
{
	LONG result;
	GetProperty(0x19, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetlMinutesPerWeek(LONG propval)
{
	SetProperty(0x1a, VT_I4, propval);
}

LONG CMP14::GetlMinutesPerWeek()
{
	LONG result;
	GetProperty(0x1a, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetlDaysPerMonth(LONG propval)
{
	SetProperty(0x1b, VT_I4, propval);
}

LONG CMP14::GetlDaysPerMonth()
{
	LONG result;
	GetProperty(0x1b, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetyDefaultTaskType(LONG propval)
{
	SetProperty(0x1c, VT_I4, propval);
}

LONG CMP14::GetyDefaultTaskType()
{
	LONG result;
	GetProperty(0x1c, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetyDefaultFixedCostAccrual(LONG propval)
{
	SetProperty(0x1d, VT_I4, propval);
}

LONG CMP14::GetyDefaultFixedCostAccrual()
{
	LONG result;
	GetProperty(0x1d, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetfDefaultStandardRate(FLOAT propval)
{
	SetProperty(0x1e, VT_R4, propval);
}

FLOAT CMP14::GetfDefaultStandardRate()
{
	FLOAT result;
	GetProperty(0x1e, VT_R4, (void*)&result);
	return result;
}

void CMP14::SetfDefaultOvertimeRate(FLOAT propval)
{
	SetProperty(0x1f, VT_R4, propval);
}

FLOAT CMP14::GetfDefaultOvertimeRate()
{
	FLOAT result;
	GetProperty(0x1f, VT_R4, (void*)&result);
	return result;
}

void CMP14::SetyDurationFormat(LONG propval)
{
	SetProperty(0x20, VT_I4, propval);
}

LONG CMP14::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x20, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetyWorkFormat(LONG propval)
{
	SetProperty(0x21, VT_I4, propval);
}

LONG CMP14::GetyWorkFormat()
{
	LONG result;
	GetProperty(0x21, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetbEditableActualCosts(BOOL propval)
{
	SetProperty(0x22, VT_BOOL, propval);
}

BOOL CMP14::GetbEditableActualCosts()
{
	BOOL result;
	GetProperty(0x22, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbHonorConstraints(BOOL propval)
{
	SetProperty(0x23, VT_BOOL, propval);
}

BOOL CMP14::GetbHonorConstraints()
{
	BOOL result;
	GetProperty(0x23, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetyEarnedValueMethod(LONG propval)
{
	SetProperty(0x24, VT_I4, propval);
}

LONG CMP14::GetyEarnedValueMethod()
{
	LONG result;
	GetProperty(0x24, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetbInsertedProjectsLikeSummary(BOOL propval)
{
	SetProperty(0x25, VT_BOOL, propval);
}

BOOL CMP14::GetbInsertedProjectsLikeSummary()
{
	BOOL result;
	GetProperty(0x25, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbMultipleCriticalPaths(BOOL propval)
{
	SetProperty(0x26, VT_BOOL, propval);
}

BOOL CMP14::GetbMultipleCriticalPaths()
{
	BOOL result;
	GetProperty(0x26, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbNewTasksEffortDriven(BOOL propval)
{
	SetProperty(0x27, VT_BOOL, propval);
}

BOOL CMP14::GetbNewTasksEffortDriven()
{
	BOOL result;
	GetProperty(0x27, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbNewTasksEstimated(BOOL propval)
{
	SetProperty(0x28, VT_BOOL, propval);
}

BOOL CMP14::GetbNewTasksEstimated()
{
	BOOL result;
	GetProperty(0x28, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbSplitsInProgressTasks(BOOL propval)
{
	SetProperty(0x29, VT_BOOL, propval);
}

BOOL CMP14::GetbSplitsInProgressTasks()
{
	BOOL result;
	GetProperty(0x29, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbSpreadActualCost(BOOL propval)
{
	SetProperty(0x2a, VT_BOOL, propval);
}

BOOL CMP14::GetbSpreadActualCost()
{
	BOOL result;
	GetProperty(0x2a, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbSpreadPercentComplete(BOOL propval)
{
	SetProperty(0x2b, VT_BOOL, propval);
}

BOOL CMP14::GetbSpreadPercentComplete()
{
	BOOL result;
	GetProperty(0x2b, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbTaskUpdatesResource(BOOL propval)
{
	SetProperty(0x2c, VT_BOOL, propval);
}

BOOL CMP14::GetbTaskUpdatesResource()
{
	BOOL result;
	GetProperty(0x2c, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbFiscalYearStart(BOOL propval)
{
	SetProperty(0x2d, VT_BOOL, propval);
}

BOOL CMP14::GetbFiscalYearStart()
{
	BOOL result;
	GetProperty(0x2d, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetyWeekStartDay(LONG propval)
{
	SetProperty(0x2e, VT_I4, propval);
}

LONG CMP14::GetyWeekStartDay()
{
	LONG result;
	GetProperty(0x2e, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetbMoveCompletedEndsBack(BOOL propval)
{
	SetProperty(0x2f, VT_BOOL, propval);
}

BOOL CMP14::GetbMoveCompletedEndsBack()
{
	BOOL result;
	GetProperty(0x2f, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbMoveRemainingStartsBack(BOOL propval)
{
	SetProperty(0x30, VT_BOOL, propval);
}

BOOL CMP14::GetbMoveRemainingStartsBack()
{
	BOOL result;
	GetProperty(0x30, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbMoveRemainingStartsForward(BOOL propval)
{
	SetProperty(0x31, VT_BOOL, propval);
}

BOOL CMP14::GetbMoveRemainingStartsForward()
{
	BOOL result;
	GetProperty(0x31, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbMoveCompletedEndsForward(BOOL propval)
{
	SetProperty(0x32, VT_BOOL, propval);
}

BOOL CMP14::GetbMoveCompletedEndsForward()
{
	BOOL result;
	GetProperty(0x32, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetyBaselineForEarnedValue(LONG propval)
{
	SetProperty(0x33, VT_I4, propval);
}

LONG CMP14::GetyBaselineForEarnedValue()
{
	LONG result;
	GetProperty(0x33, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetbAutoAddNewResourcesAndTasks(BOOL propval)
{
	SetProperty(0x34, VT_BOOL, propval);
}

BOOL CMP14::GetbAutoAddNewResourcesAndTasks()
{
	BOOL result;
	GetProperty(0x34, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetdtStatusDate(DATE propval)
{
	SetProperty(0x35, VT_DATE, propval);
}

DATE CMP14::GetdtStatusDate()
{
	DATE result;
	GetProperty(0x35, VT_DATE, (void*)&result);
	return result;
}

void CMP14::SetdtCurrentDate(DATE propval)
{
	SetProperty(0x36, VT_DATE, propval);
}

DATE CMP14::GetdtCurrentDate()
{
	DATE result;
	GetProperty(0x36, VT_DATE, (void*)&result);
	return result;
}

void CMP14::SetbMicrosoftProjectServerURL(BOOL propval)
{
	SetProperty(0x37, VT_BOOL, propval);
}

BOOL CMP14::GetbMicrosoftProjectServerURL()
{
	BOOL result;
	GetProperty(0x37, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbAutolink(BOOL propval)
{
	SetProperty(0x38, VT_BOOL, propval);
}

BOOL CMP14::GetbAutolink()
{
	BOOL result;
	GetProperty(0x38, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetyNewTaskStartDate(LONG propval)
{
	SetProperty(0x39, VT_I4, propval);
}

LONG CMP14::GetyNewTaskStartDate()
{
	LONG result;
	GetProperty(0x39, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetyDefaultTaskEVMethod(LONG propval)
{
	SetProperty(0x3a, VT_I4, propval);
}

LONG CMP14::GetyDefaultTaskEVMethod()
{
	LONG result;
	GetProperty(0x3a, VT_I4, (void*)&result);
	return result;
}

void CMP14::SetbProjectExternallyEdited(BOOL propval)
{
	SetProperty(0x3b, VT_BOOL, propval);
}

BOOL CMP14::GetbProjectExternallyEdited()
{
	BOOL result;
	GetProperty(0x3b, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetdtExtendedCreationDate(DATE propval)
{
	SetProperty(0x3c, VT_DATE, propval);
}

DATE CMP14::GetdtExtendedCreationDate()
{
	DATE result;
	GetProperty(0x3c, VT_DATE, (void*)&result);
	return result;
}

void CMP14::SetbActualsInSync(BOOL propval)
{
	SetProperty(0x3d, VT_BOOL, propval);
}

BOOL CMP14::GetbActualsInSync()
{
	BOOL result;
	GetProperty(0x3d, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbRemoveFileProperties(BOOL propval)
{
	SetProperty(0x3e, VT_BOOL, propval);
}

BOOL CMP14::GetbRemoveFileProperties()
{
	BOOL result;
	GetProperty(0x3e, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbAdminProject(BOOL propval)
{
	SetProperty(0x3f, VT_BOOL, propval);
}

BOOL CMP14::GetbAdminProject()
{
	BOOL result;
	GetProperty(0x3f, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetsBaslineCalendar(LPCTSTR propval)
{
	SetProperty(0x40, VT_BSTR, propval);
}

CString CMP14::GetsBaslineCalendar()
{
	CString result;
	GetProperty(0x40, VT_BSTR, (void*)&result);
	return result;
}

void CMP14::SetbNewTasksAreManual(BOOL propval)
{
	SetProperty(0x41, VT_BOOL, propval);
}

BOOL CMP14::GetbNewTasksAreManual()
{
	BOOL result;
	GetProperty(0x41, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbUpdateManuallyScheduledTasksWhenEditingLinks(BOOL propval)
{
	SetProperty(0x42, VT_BOOL, propval);
}

BOOL CMP14::GetbUpdateManuallyScheduledTasksWhenEditingLinks()
{
	BOOL result;
	GetProperty(0x42, VT_BOOL, (void*)&result);
	return result;
}

void CMP14::SetbKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled(BOOL propval)
{
	SetProperty(0x43, VT_BOOL, propval);
}

BOOL CMP14::GetbKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled()
{
	BOOL result;
	GetProperty(0x43, VT_BOOL, (void*)&result);
	return result;
}

COutlineCodes CMP14::GetoOutlineCodes()
{
	LPDISPATCH pDispatch;
	GetProperty(0x44, VT_DISPATCH, (void*)&pDispatch);
	return COutlineCodes(pDispatch);
}

CWBSMasks CMP14::GetoWBSMasks()
{
	LPDISPATCH pDispatch;
	GetProperty(0x45, VT_DISPATCH, (void*)&pDispatch);
	return CWBSMasks(pDispatch);
}

CExtendedAttributes CMP14::GetoExtendedAttributes()
{
	LPDISPATCH pDispatch;
	GetProperty(0x46, VT_DISPATCH, (void*)&pDispatch);
	return CExtendedAttributes(pDispatch);
}

CCalendars CMP14::GetoCalendars()
{
	LPDISPATCH pDispatch;
	GetProperty(0x47, VT_DISPATCH, (void*)&pDispatch);
	return CCalendars(pDispatch);
}

CTasks CMP14::GetoTasks()
{
	LPDISPATCH pDispatch;
	GetProperty(0x48, VT_DISPATCH, (void*)&pDispatch);
	return CTasks(pDispatch);
}

CResources CMP14::GetoResources()
{
	LPDISPATCH pDispatch;
	GetProperty(0x49, VT_DISPATCH, (void*)&pDispatch);
	return CResources(pDispatch);
}

CAssignments CMP14::GetoAssignments()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4a, VT_DISPATCH, (void*)&pDispatch);
	return CAssignments(pDispatch);
}

void CMP14::WriteXML(LPCTSTR url)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4b, DISPATCH_METHOD, VT_EMPTY, NULL, params, url);
}

void CMP14::ReadXML(LPCTSTR url)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4c, DISPATCH_METHOD, VT_EMPTY, NULL, params, url);
}

CString CMP14::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x4d, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CMP14::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x4e, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CMP14::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x4f, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CMP14::Initialize(void)
{
	InvokeHelper(0x50, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CMP14 CMSP2010Ctl::GetMP14()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CMP14(pDispatch);
}

void CMSP2010Ctl::AboutBox(void)
{
	InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CMSP2010Ctl::Clear(void)
{
	InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}
}
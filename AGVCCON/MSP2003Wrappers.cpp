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
#include "MSP2003Wrappers.h"

namespace MSP2003
{
IMPLEMENT_DYNCREATE(CMSP2003Ctl, CWnd)

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

void CAssignmentExtendedAttribute::SetlUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CAssignmentExtendedAttribute::GetlUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CAssignmentExtendedAttribute::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CAssignmentExtendedAttribute::GetsFieldID()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CAssignmentExtendedAttribute::SetsValue(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CAssignmentExtendedAttribute::GetsValue()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CAssignmentExtendedAttribute::SetlValueID(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CAssignmentExtendedAttribute::GetlValueID()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CAssignmentExtendedAttribute::SetyDurationFormat(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CAssignmentExtendedAttribute::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CAssignmentExtendedAttribute::SetKey(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CAssignmentExtendedAttribute::GetKey()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

CString CAssignmentExtendedAttribute::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentExtendedAttribute::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CAssignmentExtendedAttribute::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignmentExtendedAttribute::Initialize(void)
{
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

CDuration CAssignment::GetoRegularWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x22, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetcRemainingCost(DOUBLE propval)
{
	SetProperty(0x23, VT_R8, propval);
}

DOUBLE CAssignment::GetcRemainingCost()
{
	DOUBLE result;
	GetProperty(0x23, VT_R8, (void*)&result);
	return result;
}

void CAssignment::SetcRemainingOvertimeCost(DOUBLE propval)
{
	SetProperty(0x24, VT_R8, propval);
}

DOUBLE CAssignment::GetcRemainingOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x24, VT_R8, (void*)&result);
	return result;
}

CDuration CAssignment::GetoRemainingOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x25, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CAssignment::GetoRemainingWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x26, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetbResponsePending(BOOL propval)
{
	SetProperty(0x27, VT_BOOL, propval);
}

BOOL CAssignment::GetbResponsePending()
{
	BOOL result;
	GetProperty(0x27, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetdtStart(DATE propval)
{
	SetProperty(0x28, VT_DATE, propval);
}

DATE CAssignment::GetdtStart()
{
	DATE result;
	GetProperty(0x28, VT_DATE, (void*)&result);
	return result;
}

void CAssignment::SetdtStop(DATE propval)
{
	SetProperty(0x29, VT_DATE, propval);
}

DATE CAssignment::GetdtStop()
{
	DATE result;
	GetProperty(0x29, VT_DATE, (void*)&result);
	return result;
}

void CAssignment::SetdtResume(DATE propval)
{
	SetProperty(0x2a, VT_DATE, propval);
}

DATE CAssignment::GetdtResume()
{
	DATE result;
	GetProperty(0x2a, VT_DATE, (void*)&result);
	return result;
}

void CAssignment::SetlStartVariance(LONG propval)
{
	SetProperty(0x2b, VT_I4, propval);
}

LONG CAssignment::GetlStartVariance()
{
	LONG result;
	GetProperty(0x2b, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetfUnits(FLOAT propval)
{
	SetProperty(0x2c, VT_R4, propval);
}

FLOAT CAssignment::GetfUnits()
{
	FLOAT result;
	GetProperty(0x2c, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetbUpdateNeeded(BOOL propval)
{
	SetProperty(0x2d, VT_BOOL, propval);
}

BOOL CAssignment::GetbUpdateNeeded()
{
	BOOL result;
	GetProperty(0x2d, VT_BOOL, (void*)&result);
	return result;
}

void CAssignment::SetfVAC(FLOAT propval)
{
	SetProperty(0x2e, VT_R4, propval);
}

FLOAT CAssignment::GetfVAC()
{
	FLOAT result;
	GetProperty(0x2e, VT_R4, (void*)&result);
	return result;
}

CDuration CAssignment::GetoWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2f, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetyWorkContour(LONG propval)
{
	SetProperty(0x30, VT_I4, propval);
}

LONG CAssignment::GetyWorkContour()
{
	LONG result;
	GetProperty(0x30, VT_I4, (void*)&result);
	return result;
}

void CAssignment::SetfBCWS(FLOAT propval)
{
	SetProperty(0x31, VT_R4, propval);
}

FLOAT CAssignment::GetfBCWS()
{
	FLOAT result;
	GetProperty(0x31, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetfBCWP(FLOAT propval)
{
	SetProperty(0x32, VT_R4, propval);
}

FLOAT CAssignment::GetfBCWP()
{
	FLOAT result;
	GetProperty(0x32, VT_R4, (void*)&result);
	return result;
}

void CAssignment::SetyBookingType(LONG propval)
{
	SetProperty(0x33, VT_I4, propval);
}

LONG CAssignment::GetyBookingType()
{
	LONG result;
	GetProperty(0x33, VT_I4, (void*)&result);
	return result;
}

CDuration CAssignment::GetoActualWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x34, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CAssignment::GetoActualOvertimeWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x35, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CAssignment::SetdtCreationDate(DATE propval)
{
	SetProperty(0x36, VT_DATE, propval);
}

DATE CAssignment::GetdtCreationDate()
{
	DATE result;
	GetProperty(0x36, VT_DATE, (void*)&result);
	return result;
}

CAssignmentExtendedAttribute_C CAssignment::GetoExtendedAttribute_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x37, VT_DISPATCH, (void*)&pDispatch);
	return CAssignmentExtendedAttribute_C(pDispatch);
}

CAssignmentBaseline_C CAssignment::GetoBaseline_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x38, VT_DISPATCH, (void*)&pDispatch);
	return CAssignmentBaseline_C(pDispatch);
}

CTimephasedData_C CAssignment::GetoTimephasedData_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x39, VT_DISPATCH, (void*)&pDispatch);
	return CTimephasedData_C(pDispatch);
}

void CAssignment::SetKey(LPCTSTR propval)
{
	SetProperty(0x3a, VT_BSTR, propval);
}

CString CAssignment::GetKey()
{
	CString result;
	GetProperty(0x3a, VT_BSTR, (void*)&result);
	return result;
}

CString CAssignment::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x3b, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignment::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x3c, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CAssignment::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x3d, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CAssignment::Initialize(void)
{
	InvokeHelper(0x3e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

void CResourceOutlineCode::SetlUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CResourceOutlineCode::GetlUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CResourceOutlineCode::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CResourceOutlineCode::GetsFieldID()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CResourceOutlineCode::SetlValueID(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CResourceOutlineCode::GetlValueID()
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

void CResourceExtendedAttribute::SetlUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CResourceExtendedAttribute::GetlUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CResourceExtendedAttribute::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CResourceExtendedAttribute::GetsFieldID()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CResourceExtendedAttribute::SetsValue(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CResourceExtendedAttribute::GetsValue()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CResourceExtendedAttribute::SetlValueID(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CResourceExtendedAttribute::GetlValueID()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CResourceExtendedAttribute::SetyDurationFormat(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CResourceExtendedAttribute::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CResourceExtendedAttribute::SetKey(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CResourceExtendedAttribute::GetKey()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

CString CResourceExtendedAttribute::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceExtendedAttribute::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CResourceExtendedAttribute::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResourceExtendedAttribute::Initialize(void)
{
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

CResourceAvailabilityPeriods CResource::GetoAvailabilityPeriods()
{
	LPDISPATCH pDispatch;
	GetProperty(0x41, VT_DISPATCH, (void*)&pDispatch);
	return CResourceAvailabilityPeriods(pDispatch);
}

CResourceRates CResource::GetoRates()
{
	LPDISPATCH pDispatch;
	GetProperty(0x42, VT_DISPATCH, (void*)&pDispatch);
	return CResourceRates(pDispatch);
}

CTimephasedData_C CResource::GetoTimephasedData_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x43, VT_DISPATCH, (void*)&pDispatch);
	return CTimephasedData_C(pDispatch);
}

void CResource::SetKey(LPCTSTR propval)
{
	SetProperty(0x44, VT_BSTR, propval);
}

CString CResource::GetKey()
{
	CString result;
	GetProperty(0x44, VT_BSTR, (void*)&result);
	return result;
}

CString CResource::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x45, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CResource::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x46, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CResource::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x47, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CResource::Initialize(void)
{
	InvokeHelper(0x48, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

void CTaskOutlineCode::SetlUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CTaskOutlineCode::GetlUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CTaskOutlineCode::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CTaskOutlineCode::GetsFieldID()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CTaskOutlineCode::SetlValueID(LONG propval)
{
	SetProperty(0x3, VT_I4, propval);
}

LONG CTaskOutlineCode::GetlValueID()
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

void CTaskBaseline::SetKey(LPCTSTR propval)
{
	SetProperty(0xd, VT_BSTR, propval);
}

CString CTaskBaseline::GetKey()
{
	CString result;
	GetProperty(0xd, VT_BSTR, (void*)&result);
	return result;
}

CString CTaskBaseline::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xe, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskBaseline::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTaskBaseline::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x10, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskBaseline::Initialize(void)
{
	InvokeHelper(0x11, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

void CTaskExtendedAttribute::SetlUID(LONG propval)
{
	SetProperty(0x1, VT_I4, propval);
}

LONG CTaskExtendedAttribute::GetlUID()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

void CTaskExtendedAttribute::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CTaskExtendedAttribute::GetsFieldID()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CTaskExtendedAttribute::SetsValue(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CTaskExtendedAttribute::GetsValue()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CTaskExtendedAttribute::SetlValueID(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CTaskExtendedAttribute::GetlValueID()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

void CTaskExtendedAttribute::SetyDurationFormat(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CTaskExtendedAttribute::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CTaskExtendedAttribute::SetKey(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CTaskExtendedAttribute::GetKey()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

CString CTaskExtendedAttribute::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskExtendedAttribute::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTaskExtendedAttribute::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTaskExtendedAttribute::Initialize(void)
{
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

void CTask::SetbCritical(BOOL propval)
{
	SetProperty(0x1b, VT_BOOL, propval);
}

BOOL CTask::GetbCritical()
{
	BOOL result;
	GetProperty(0x1b, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbIsSubproject(BOOL propval)
{
	SetProperty(0x1c, VT_BOOL, propval);
}

BOOL CTask::GetbIsSubproject()
{
	BOOL result;
	GetProperty(0x1c, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbIsSubprojectReadOnly(BOOL propval)
{
	SetProperty(0x1d, VT_BOOL, propval);
}

BOOL CTask::GetbIsSubprojectReadOnly()
{
	BOOL result;
	GetProperty(0x1d, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetsSubprojectName(LPCTSTR propval)
{
	SetProperty(0x1e, VT_BSTR, propval);
}

CString CTask::GetsSubprojectName()
{
	CString result;
	GetProperty(0x1e, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetbExternalTask(BOOL propval)
{
	SetProperty(0x1f, VT_BOOL, propval);
}

BOOL CTask::GetbExternalTask()
{
	BOOL result;
	GetProperty(0x1f, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetsExternalTaskProject(LPCTSTR propval)
{
	SetProperty(0x20, VT_BSTR, propval);
}

CString CTask::GetsExternalTaskProject()
{
	CString result;
	GetProperty(0x20, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetdtEarlyStart(DATE propval)
{
	SetProperty(0x21, VT_DATE, propval);
}

DATE CTask::GetdtEarlyStart()
{
	DATE result;
	GetProperty(0x21, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtEarlyFinish(DATE propval)
{
	SetProperty(0x22, VT_DATE, propval);
}

DATE CTask::GetdtEarlyFinish()
{
	DATE result;
	GetProperty(0x22, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtLateStart(DATE propval)
{
	SetProperty(0x23, VT_DATE, propval);
}

DATE CTask::GetdtLateStart()
{
	DATE result;
	GetProperty(0x23, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtLateFinish(DATE propval)
{
	SetProperty(0x24, VT_DATE, propval);
}

DATE CTask::GetdtLateFinish()
{
	DATE result;
	GetProperty(0x24, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetlStartVariance(LONG propval)
{
	SetProperty(0x25, VT_I4, propval);
}

LONG CTask::GetlStartVariance()
{
	LONG result;
	GetProperty(0x25, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlFinishVariance(LONG propval)
{
	SetProperty(0x26, VT_I4, propval);
}

LONG CTask::GetlFinishVariance()
{
	LONG result;
	GetProperty(0x26, VT_I4, (void*)&result);
	return result;
}

void CTask::SetfWorkVariance(FLOAT propval)
{
	SetProperty(0x27, VT_R4, propval);
}

FLOAT CTask::GetfWorkVariance()
{
	FLOAT result;
	GetProperty(0x27, VT_R4, (void*)&result);
	return result;
}

void CTask::SetlFreeSlack(LONG propval)
{
	SetProperty(0x28, VT_I4, propval);
}

LONG CTask::GetlFreeSlack()
{
	LONG result;
	GetProperty(0x28, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlTotalSlack(LONG propval)
{
	SetProperty(0x29, VT_I4, propval);
}

LONG CTask::GetlTotalSlack()
{
	LONG result;
	GetProperty(0x29, VT_I4, (void*)&result);
	return result;
}

void CTask::SetfFixedCost(FLOAT propval)
{
	SetProperty(0x2a, VT_R4, propval);
}

FLOAT CTask::GetfFixedCost()
{
	FLOAT result;
	GetProperty(0x2a, VT_R4, (void*)&result);
	return result;
}

void CTask::SetyFixedCostAccrual(LONG propval)
{
	SetProperty(0x2b, VT_I4, propval);
}

LONG CTask::GetyFixedCostAccrual()
{
	LONG result;
	GetProperty(0x2b, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlPercentComplete(LONG propval)
{
	SetProperty(0x2c, VT_I4, propval);
}

LONG CTask::GetlPercentComplete()
{
	LONG result;
	GetProperty(0x2c, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlPercentWorkComplete(LONG propval)
{
	SetProperty(0x2d, VT_I4, propval);
}

LONG CTask::GetlPercentWorkComplete()
{
	LONG result;
	GetProperty(0x2d, VT_I4, (void*)&result);
	return result;
}

void CTask::SetcCost(DOUBLE propval)
{
	SetProperty(0x2e, VT_R8, propval);
}

DOUBLE CTask::GetcCost()
{
	DOUBLE result;
	GetProperty(0x2e, VT_R8, (void*)&result);
	return result;
}

void CTask::SetcOvertimeCost(DOUBLE propval)
{
	SetProperty(0x2f, VT_R8, propval);
}

DOUBLE CTask::GetcOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x2f, VT_R8, (void*)&result);
	return result;
}

CDuration CTask::GetoOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x30, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetdtActualStart(DATE propval)
{
	SetProperty(0x31, VT_DATE, propval);
}

DATE CTask::GetdtActualStart()
{
	DATE result;
	GetProperty(0x31, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtActualFinish(DATE propval)
{
	SetProperty(0x32, VT_DATE, propval);
}

DATE CTask::GetdtActualFinish()
{
	DATE result;
	GetProperty(0x32, VT_DATE, (void*)&result);
	return result;
}

CDuration CTask::GetoActualDuration()
{
	LPDISPATCH pDispatch;
	GetProperty(0x33, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetcActualCost(DOUBLE propval)
{
	SetProperty(0x34, VT_R8, propval);
}

DOUBLE CTask::GetcActualCost()
{
	DOUBLE result;
	GetProperty(0x34, VT_R8, (void*)&result);
	return result;
}

void CTask::SetcActualOvertimeCost(DOUBLE propval)
{
	SetProperty(0x35, VT_R8, propval);
}

DOUBLE CTask::GetcActualOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x35, VT_R8, (void*)&result);
	return result;
}

CDuration CTask::GetoActualWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x36, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CTask::GetoActualOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x37, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CTask::GetoRegularWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x38, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CTask::GetoRemainingDuration()
{
	LPDISPATCH pDispatch;
	GetProperty(0x39, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetcRemainingCost(DOUBLE propval)
{
	SetProperty(0x3a, VT_R8, propval);
}

DOUBLE CTask::GetcRemainingCost()
{
	DOUBLE result;
	GetProperty(0x3a, VT_R8, (void*)&result);
	return result;
}

CDuration CTask::GetoRemainingWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3b, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetcRemainingOvertimeCost(DOUBLE propval)
{
	SetProperty(0x3c, VT_R8, propval);
}

DOUBLE CTask::GetcRemainingOvertimeCost()
{
	DOUBLE result;
	GetProperty(0x3c, VT_R8, (void*)&result);
	return result;
}

CDuration CTask::GetoRemainingOvertimeWork()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3d, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

void CTask::SetfACWP(FLOAT propval)
{
	SetProperty(0x3e, VT_R4, propval);
}

FLOAT CTask::GetfACWP()
{
	FLOAT result;
	GetProperty(0x3e, VT_R4, (void*)&result);
	return result;
}

void CTask::SetfCV(FLOAT propval)
{
	SetProperty(0x3f, VT_R4, propval);
}

FLOAT CTask::GetfCV()
{
	FLOAT result;
	GetProperty(0x3f, VT_R4, (void*)&result);
	return result;
}

void CTask::SetyConstraintType(LONG propval)
{
	SetProperty(0x40, VT_I4, propval);
}

LONG CTask::GetyConstraintType()
{
	LONG result;
	GetProperty(0x40, VT_I4, (void*)&result);
	return result;
}

void CTask::SetlCalendarUID(LONG propval)
{
	SetProperty(0x41, VT_I4, propval);
}

LONG CTask::GetlCalendarUID()
{
	LONG result;
	GetProperty(0x41, VT_I4, (void*)&result);
	return result;
}

void CTask::SetdtConstraintDate(DATE propval)
{
	SetProperty(0x42, VT_DATE, propval);
}

DATE CTask::GetdtConstraintDate()
{
	DATE result;
	GetProperty(0x42, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtDeadline(DATE propval)
{
	SetProperty(0x43, VT_DATE, propval);
}

DATE CTask::GetdtDeadline()
{
	DATE result;
	GetProperty(0x43, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetbLevelAssignments(BOOL propval)
{
	SetProperty(0x44, VT_BOOL, propval);
}

BOOL CTask::GetbLevelAssignments()
{
	BOOL result;
	GetProperty(0x44, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbLevelingCanSplit(BOOL propval)
{
	SetProperty(0x45, VT_BOOL, propval);
}

BOOL CTask::GetbLevelingCanSplit()
{
	BOOL result;
	GetProperty(0x45, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetlLevelingDelay(LONG propval)
{
	SetProperty(0x46, VT_I4, propval);
}

LONG CTask::GetlLevelingDelay()
{
	LONG result;
	GetProperty(0x46, VT_I4, (void*)&result);
	return result;
}

void CTask::SetyLevelingDelayFormat(LONG propval)
{
	SetProperty(0x47, VT_I4, propval);
}

LONG CTask::GetyLevelingDelayFormat()
{
	LONG result;
	GetProperty(0x47, VT_I4, (void*)&result);
	return result;
}

void CTask::SetdtPreLeveledStart(DATE propval)
{
	SetProperty(0x48, VT_DATE, propval);
}

DATE CTask::GetdtPreLeveledStart()
{
	DATE result;
	GetProperty(0x48, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetdtPreLeveledFinish(DATE propval)
{
	SetProperty(0x49, VT_DATE, propval);
}

DATE CTask::GetdtPreLeveledFinish()
{
	DATE result;
	GetProperty(0x49, VT_DATE, (void*)&result);
	return result;
}

void CTask::SetsHyperlink(LPCTSTR propval)
{
	SetProperty(0x4a, VT_BSTR, propval);
}

CString CTask::GetsHyperlink()
{
	CString result;
	GetProperty(0x4a, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetsHyperlinkAddress(LPCTSTR propval)
{
	SetProperty(0x4b, VT_BSTR, propval);
}

CString CTask::GetsHyperlinkAddress()
{
	CString result;
	GetProperty(0x4b, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetsHyperlinkSubAddress(LPCTSTR propval)
{
	SetProperty(0x4c, VT_BSTR, propval);
}

CString CTask::GetsHyperlinkSubAddress()
{
	CString result;
	GetProperty(0x4c, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetbIgnoreResourceCalendar(BOOL propval)
{
	SetProperty(0x4d, VT_BOOL, propval);
}

BOOL CTask::GetbIgnoreResourceCalendar()
{
	BOOL result;
	GetProperty(0x4d, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetsNotes(LPCTSTR propval)
{
	SetProperty(0x4e, VT_BSTR, propval);
}

CString CTask::GetsNotes()
{
	CString result;
	GetProperty(0x4e, VT_BSTR, (void*)&result);
	return result;
}

void CTask::SetbHideBar(BOOL propval)
{
	SetProperty(0x4f, VT_BOOL, propval);
}

BOOL CTask::GetbHideBar()
{
	BOOL result;
	GetProperty(0x4f, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetbRollup(BOOL propval)
{
	SetProperty(0x50, VT_BOOL, propval);
}

BOOL CTask::GetbRollup()
{
	BOOL result;
	GetProperty(0x50, VT_BOOL, (void*)&result);
	return result;
}

void CTask::SetfBCWS(FLOAT propval)
{
	SetProperty(0x51, VT_R4, propval);
}

FLOAT CTask::GetfBCWS()
{
	FLOAT result;
	GetProperty(0x51, VT_R4, (void*)&result);
	return result;
}

void CTask::SetfBCWP(FLOAT propval)
{
	SetProperty(0x52, VT_R4, propval);
}

FLOAT CTask::GetfBCWP()
{
	FLOAT result;
	GetProperty(0x52, VT_R4, (void*)&result);
	return result;
}

void CTask::SetlPhysicalPercentComplete(LONG propval)
{
	SetProperty(0x53, VT_I4, propval);
}

LONG CTask::GetlPhysicalPercentComplete()
{
	LONG result;
	GetProperty(0x53, VT_I4, (void*)&result);
	return result;
}

void CTask::SetyEarnedValueMethod(LONG propval)
{
	SetProperty(0x54, VT_I4, propval);
}

LONG CTask::GetyEarnedValueMethod()
{
	LONG result;
	GetProperty(0x54, VT_I4, (void*)&result);
	return result;
}

CTaskPredecessorLink_C CTask::GetoPredecessorLink_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x55, VT_DISPATCH, (void*)&pDispatch);
	return CTaskPredecessorLink_C(pDispatch);
}

CDuration CTask::GetoActualWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x56, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CDuration CTask::GetoActualOvertimeWorkProtected()
{
	LPDISPATCH pDispatch;
	GetProperty(0x57, VT_DISPATCH, (void*)&pDispatch);
	return CDuration(pDispatch);
}

CTaskExtendedAttribute_C CTask::GetoExtendedAttribute_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x58, VT_DISPATCH, (void*)&pDispatch);
	return CTaskExtendedAttribute_C(pDispatch);
}

CTaskBaseline_C CTask::GetoBaseline_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x59, VT_DISPATCH, (void*)&pDispatch);
	return CTaskBaseline_C(pDispatch);
}

CTaskOutlineCode_C CTask::GetoOutlineCode_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5a, VT_DISPATCH, (void*)&pDispatch);
	return CTaskOutlineCode_C(pDispatch);
}

CTimephasedData_C CTask::GetoTimephasedData_C()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5b, VT_DISPATCH, (void*)&pDispatch);
	return CTimephasedData_C(pDispatch);
}

void CTask::SetKey(LPCTSTR propval)
{
	SetProperty(0x5c, VT_BSTR, propval);
}

CString CTask::GetKey()
{
	CString result;
	GetProperty(0x5c, VT_BSTR, (void*)&result);
	return result;
}

CString CTask::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5d, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CTask::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5e, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CTask::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x5f, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CTask::Initialize(void)
{
	InvokeHelper(0x60, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

CTimePeriod CWeekDay::GetoTimePeriod()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3, VT_DISPATCH, (void*)&pDispatch);
	return CTimePeriod(pDispatch);
}

CWorkingTimes CWeekDay::GetoWorkingTimes()
{
	LPDISPATCH pDispatch;
	GetProperty(0x4, VT_DISPATCH, (void*)&pDispatch);
	return CWorkingTimes(pDispatch);
}

void CWeekDay::SetKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CWeekDay::GetKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CString CWeekDay::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CWeekDay::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CWeekDay::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CWeekDay::Initialize(void)
{
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LONG CWeekDays::GetCount()
{
	LONG result;
	GetProperty(0x1, VT_I4, (void*)&result);
	return result;
}

CWeekDay CWeekDays::Item(LPCTSTR Index)
{
	CWeekDay oReturn;
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x2, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, params, Index);
	return oReturn;
}

CWeekDay CWeekDays::Add(void)
{
	CWeekDay oReturn;
	InvokeHelper(0x3, DISPATCH_METHOD, VT_DISPATCH, (void*)&oReturn, NULL);
	return oReturn;
}

void CWeekDays::Clear(void)
{
	InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CWeekDays::Remove(LPCTSTR Index)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, params, Index);
}

BOOL CWeekDays::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CWeekDays::Initialize(void)
{
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CWeekDays::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

CString CWeekDays::GetXML(void)
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

void CCalendar::SetlBaseCalendarUID(LONG propval)
{
	SetProperty(0x4, VT_I4, propval);
}

LONG CCalendar::GetlBaseCalendarUID()
{
	LONG result;
	GetProperty(0x4, VT_I4, (void*)&result);
	return result;
}

CWeekDays CCalendar::GetoWeekDays()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return CWeekDays(pDispatch);
}

void CCalendar::SetKey(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CCalendar::GetKey()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

CString CCalendar::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendar::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CCalendar::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x9, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CCalendar::Initialize(void)
{
	InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

void CExtendedAttributeValue::SetKey(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CExtendedAttributeValue::GetKey()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

CString CExtendedAttributeValue::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x5, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttributeValue::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CExtendedAttributeValue::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttributeValue::Initialize(void)
{
	InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

void CExtendedAttribute::SetsAlias(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsAlias()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsPhoneticAlias(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsPhoneticAlias()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetyRollupType(LONG propval)
{
	SetProperty(0x5, VT_I4, propval);
}

LONG CExtendedAttribute::GetyRollupType()
{
	LONG result;
	GetProperty(0x5, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttribute::SetyCalculationType(LONG propval)
{
	SetProperty(0x6, VT_I4, propval);
}

LONG CExtendedAttribute::GetyCalculationType()
{
	LONG result;
	GetProperty(0x6, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsFormula(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsFormula()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

void CExtendedAttribute::SetbRestrictValues(BOOL propval)
{
	SetProperty(0x8, VT_BOOL, propval);
}

BOOL CExtendedAttribute::GetbRestrictValues()
{
	BOOL result;
	GetProperty(0x8, VT_BOOL, (void*)&result);
	return result;
}

void CExtendedAttribute::SetyValuelistSortOrder(LONG propval)
{
	SetProperty(0x9, VT_I4, propval);
}

LONG CExtendedAttribute::GetyValuelistSortOrder()
{
	LONG result;
	GetProperty(0x9, VT_I4, (void*)&result);
	return result;
}

void CExtendedAttribute::SetbAppendNewValues(BOOL propval)
{
	SetProperty(0xa, VT_BOOL, propval);
}

BOOL CExtendedAttribute::GetbAppendNewValues()
{
	BOOL result;
	GetProperty(0xa, VT_BOOL, (void*)&result);
	return result;
}

void CExtendedAttribute::SetsDefault(LPCTSTR propval)
{
	SetProperty(0xb, VT_BSTR, propval);
}

CString CExtendedAttribute::GetsDefault()
{
	CString result;
	GetProperty(0xb, VT_BSTR, (void*)&result);
	return result;
}

CExtendedAttributeValueList CExtendedAttribute::GetoValueList()
{
	LPDISPATCH pDispatch;
	GetProperty(0xc, VT_DISPATCH, (void*)&pDispatch);
	return CExtendedAttributeValueList(pDispatch);
}

void CExtendedAttribute::SetKey(LPCTSTR propval)
{
	SetProperty(0xd, VT_BSTR, propval);
}

CString CExtendedAttribute::GetKey()
{
	CString result;
	GetProperty(0xd, VT_BSTR, (void*)&result);
	return result;
}

CString CExtendedAttribute::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xe, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttribute::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CExtendedAttribute::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x10, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CExtendedAttribute::Initialize(void)
{
	InvokeHelper(0x11, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

void COutlineCodeValue::SetlParentValueID(LONG propval)
{
	SetProperty(0x2, VT_I4, propval);
}

LONG COutlineCodeValue::GetlParentValueID()
{
	LONG result;
	GetProperty(0x2, VT_I4, (void*)&result);
	return result;
}

void COutlineCodeValue::SetsValue(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString COutlineCodeValue::GetsValue()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCodeValue::SetsDescription(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString COutlineCodeValue::GetsDescription()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCodeValue::SetKey(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString COutlineCodeValue::GetKey()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

CString COutlineCodeValue::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x6, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeValue::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL COutlineCodeValue::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x8, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCodeValue::Initialize(void)
{
	InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

void COutlineCode::SetsFieldID(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString COutlineCode::GetsFieldID()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCode::SetsFieldName(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString COutlineCode::GetsFieldName()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCode::SetsAlias(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString COutlineCode::GetsAlias()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void COutlineCode::SetsPhoneticAlias(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString COutlineCode::GetsPhoneticAlias()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

COutlineCodeValues COutlineCode::GetoValues()
{
	LPDISPATCH pDispatch;
	GetProperty(0x5, VT_DISPATCH, (void*)&pDispatch);
	return COutlineCodeValues(pDispatch);
}

void COutlineCode::SetbEnterprise(BOOL propval)
{
	SetProperty(0x6, VT_BOOL, propval);
}

BOOL COutlineCode::GetbEnterprise()
{
	BOOL result;
	GetProperty(0x6, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCode::SetlEnterpriseOutlineCodeAlias(LONG propval)
{
	SetProperty(0x7, VT_I4, propval);
}

LONG COutlineCode::GetlEnterpriseOutlineCodeAlias()
{
	LONG result;
	GetProperty(0x7, VT_I4, (void*)&result);
	return result;
}

void COutlineCode::SetbResourceSubstitutionEnabled(BOOL propval)
{
	SetProperty(0x8, VT_BOOL, propval);
}

BOOL COutlineCode::GetbResourceSubstitutionEnabled()
{
	BOOL result;
	GetProperty(0x8, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCode::SetbLeafOnly(BOOL propval)
{
	SetProperty(0x9, VT_BOOL, propval);
}

BOOL COutlineCode::GetbLeafOnly()
{
	BOOL result;
	GetProperty(0x9, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCode::SetbAllLevelsRequired(BOOL propval)
{
	SetProperty(0xa, VT_BOOL, propval);
}

BOOL COutlineCode::GetbAllLevelsRequired()
{
	BOOL result;
	GetProperty(0xa, VT_BOOL, (void*)&result);
	return result;
}

void COutlineCode::SetbOnlyTableValuesAllowed(BOOL propval)
{
	SetProperty(0xb, VT_BOOL, propval);
}

BOOL COutlineCode::GetbOnlyTableValuesAllowed()
{
	BOOL result;
	GetProperty(0xb, VT_BOOL, (void*)&result);
	return result;
}

COutlineCodeMasks COutlineCode::GetoMasks()
{
	LPDISPATCH pDispatch;
	GetProperty(0xc, VT_DISPATCH, (void*)&pDispatch);
	return COutlineCodeMasks(pDispatch);
}

void COutlineCode::SetKey(LPCTSTR propval)
{
	SetProperty(0xd, VT_BSTR, propval);
}

CString COutlineCode::GetKey()
{
	CString result;
	GetProperty(0xd, VT_BSTR, (void*)&result);
	return result;
}

CString COutlineCode::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0xe, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCode::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0xf, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL COutlineCode::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x10, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void COutlineCode::Initialize(void)
{
	InvokeHelper(0x11, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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

void CMP11::SetsUID(LPCTSTR propval)
{
	SetProperty(0x1, VT_BSTR, propval);
}

CString CMP11::GetsUID()
{
	CString result;
	GetProperty(0x1, VT_BSTR, (void*)&result);
	return result;
}

void CMP11::SetsName(LPCTSTR propval)
{
	SetProperty(0x2, VT_BSTR, propval);
}

CString CMP11::GetsName()
{
	CString result;
	GetProperty(0x2, VT_BSTR, (void*)&result);
	return result;
}

void CMP11::SetsTitle(LPCTSTR propval)
{
	SetProperty(0x3, VT_BSTR, propval);
}

CString CMP11::GetsTitle()
{
	CString result;
	GetProperty(0x3, VT_BSTR, (void*)&result);
	return result;
}

void CMP11::SetsSubject(LPCTSTR propval)
{
	SetProperty(0x4, VT_BSTR, propval);
}

CString CMP11::GetsSubject()
{
	CString result;
	GetProperty(0x4, VT_BSTR, (void*)&result);
	return result;
}

void CMP11::SetsCategory(LPCTSTR propval)
{
	SetProperty(0x5, VT_BSTR, propval);
}

CString CMP11::GetsCategory()
{
	CString result;
	GetProperty(0x5, VT_BSTR, (void*)&result);
	return result;
}

void CMP11::SetsCompany(LPCTSTR propval)
{
	SetProperty(0x6, VT_BSTR, propval);
}

CString CMP11::GetsCompany()
{
	CString result;
	GetProperty(0x6, VT_BSTR, (void*)&result);
	return result;
}

void CMP11::SetsManager(LPCTSTR propval)
{
	SetProperty(0x7, VT_BSTR, propval);
}

CString CMP11::GetsManager()
{
	CString result;
	GetProperty(0x7, VT_BSTR, (void*)&result);
	return result;
}

void CMP11::SetsAuthor(LPCTSTR propval)
{
	SetProperty(0x8, VT_BSTR, propval);
}

CString CMP11::GetsAuthor()
{
	CString result;
	GetProperty(0x8, VT_BSTR, (void*)&result);
	return result;
}

void CMP11::SetdtCreationDate(DATE propval)
{
	SetProperty(0x9, VT_DATE, propval);
}

DATE CMP11::GetdtCreationDate()
{
	DATE result;
	GetProperty(0x9, VT_DATE, (void*)&result);
	return result;
}

void CMP11::SetlRevision(LONG propval)
{
	SetProperty(0xa, VT_I4, propval);
}

LONG CMP11::GetlRevision()
{
	LONG result;
	GetProperty(0xa, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetdtLastSaved(DATE propval)
{
	SetProperty(0xb, VT_DATE, propval);
}

DATE CMP11::GetdtLastSaved()
{
	DATE result;
	GetProperty(0xb, VT_DATE, (void*)&result);
	return result;
}

void CMP11::SetbScheduleFromStart(BOOL propval)
{
	SetProperty(0xc, VT_BOOL, propval);
}

BOOL CMP11::GetbScheduleFromStart()
{
	BOOL result;
	GetProperty(0xc, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetdtStartDate(DATE propval)
{
	SetProperty(0xd, VT_DATE, propval);
}

DATE CMP11::GetdtStartDate()
{
	DATE result;
	GetProperty(0xd, VT_DATE, (void*)&result);
	return result;
}

void CMP11::SetdtFinishDate(DATE propval)
{
	SetProperty(0xe, VT_DATE, propval);
}

DATE CMP11::GetdtFinishDate()
{
	DATE result;
	GetProperty(0xe, VT_DATE, (void*)&result);
	return result;
}

void CMP11::SetyFYStartDate(LONG propval)
{
	SetProperty(0xf, VT_I4, propval);
}

LONG CMP11::GetyFYStartDate()
{
	LONG result;
	GetProperty(0xf, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetlCriticalSlackLimit(LONG propval)
{
	SetProperty(0x10, VT_I4, propval);
}

LONG CMP11::GetlCriticalSlackLimit()
{
	LONG result;
	GetProperty(0x10, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetlCurrencyDigits(LONG propval)
{
	SetProperty(0x11, VT_I4, propval);
}

LONG CMP11::GetlCurrencyDigits()
{
	LONG result;
	GetProperty(0x11, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetsCurrencySymbol(LPCTSTR propval)
{
	SetProperty(0x12, VT_BSTR, propval);
}

CString CMP11::GetsCurrencySymbol()
{
	CString result;
	GetProperty(0x12, VT_BSTR, (void*)&result);
	return result;
}

void CMP11::SetyCurrencySymbolPosition(LONG propval)
{
	SetProperty(0x13, VT_I4, propval);
}

LONG CMP11::GetyCurrencySymbolPosition()
{
	LONG result;
	GetProperty(0x13, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetlCalendarUID(LONG propval)
{
	SetProperty(0x14, VT_I4, propval);
}

LONG CMP11::GetlCalendarUID()
{
	LONG result;
	GetProperty(0x14, VT_I4, (void*)&result);
	return result;
}

CTime CMP11::GetoDefaultStartTime()
{
	LPDISPATCH pDispatch;
	GetProperty(0x15, VT_DISPATCH, (void*)&pDispatch);
	return CTime(pDispatch);
}

CTime CMP11::GetoDefaultFinishTime()
{
	LPDISPATCH pDispatch;
	GetProperty(0x16, VT_DISPATCH, (void*)&pDispatch);
	return CTime(pDispatch);
}

void CMP11::SetlMinutesPerDay(LONG propval)
{
	SetProperty(0x17, VT_I4, propval);
}

LONG CMP11::GetlMinutesPerDay()
{
	LONG result;
	GetProperty(0x17, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetlMinutesPerWeek(LONG propval)
{
	SetProperty(0x18, VT_I4, propval);
}

LONG CMP11::GetlMinutesPerWeek()
{
	LONG result;
	GetProperty(0x18, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetlDaysPerMonth(LONG propval)
{
	SetProperty(0x19, VT_I4, propval);
}

LONG CMP11::GetlDaysPerMonth()
{
	LONG result;
	GetProperty(0x19, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetyDefaultTaskType(LONG propval)
{
	SetProperty(0x1a, VT_I4, propval);
}

LONG CMP11::GetyDefaultTaskType()
{
	LONG result;
	GetProperty(0x1a, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetyDefaultFixedCostAccrual(LONG propval)
{
	SetProperty(0x1b, VT_I4, propval);
}

LONG CMP11::GetyDefaultFixedCostAccrual()
{
	LONG result;
	GetProperty(0x1b, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetfDefaultStandardRate(FLOAT propval)
{
	SetProperty(0x1c, VT_R4, propval);
}

FLOAT CMP11::GetfDefaultStandardRate()
{
	FLOAT result;
	GetProperty(0x1c, VT_R4, (void*)&result);
	return result;
}

void CMP11::SetfDefaultOvertimeRate(FLOAT propval)
{
	SetProperty(0x1d, VT_R4, propval);
}

FLOAT CMP11::GetfDefaultOvertimeRate()
{
	FLOAT result;
	GetProperty(0x1d, VT_R4, (void*)&result);
	return result;
}

void CMP11::SetyDurationFormat(LONG propval)
{
	SetProperty(0x1e, VT_I4, propval);
}

LONG CMP11::GetyDurationFormat()
{
	LONG result;
	GetProperty(0x1e, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetyWorkFormat(LONG propval)
{
	SetProperty(0x1f, VT_I4, propval);
}

LONG CMP11::GetyWorkFormat()
{
	LONG result;
	GetProperty(0x1f, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetbEditableActualCosts(BOOL propval)
{
	SetProperty(0x20, VT_BOOL, propval);
}

BOOL CMP11::GetbEditableActualCosts()
{
	BOOL result;
	GetProperty(0x20, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbHonorConstraints(BOOL propval)
{
	SetProperty(0x21, VT_BOOL, propval);
}

BOOL CMP11::GetbHonorConstraints()
{
	BOOL result;
	GetProperty(0x21, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetyEarnedValueMethod(LONG propval)
{
	SetProperty(0x22, VT_I4, propval);
}

LONG CMP11::GetyEarnedValueMethod()
{
	LONG result;
	GetProperty(0x22, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetbInsertedProjectsLikeSummary(BOOL propval)
{
	SetProperty(0x23, VT_BOOL, propval);
}

BOOL CMP11::GetbInsertedProjectsLikeSummary()
{
	BOOL result;
	GetProperty(0x23, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbMultipleCriticalPaths(BOOL propval)
{
	SetProperty(0x24, VT_BOOL, propval);
}

BOOL CMP11::GetbMultipleCriticalPaths()
{
	BOOL result;
	GetProperty(0x24, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbNewTasksEffortDriven(BOOL propval)
{
	SetProperty(0x25, VT_BOOL, propval);
}

BOOL CMP11::GetbNewTasksEffortDriven()
{
	BOOL result;
	GetProperty(0x25, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbNewTasksEstimated(BOOL propval)
{
	SetProperty(0x26, VT_BOOL, propval);
}

BOOL CMP11::GetbNewTasksEstimated()
{
	BOOL result;
	GetProperty(0x26, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbSplitsInProgressTasks(BOOL propval)
{
	SetProperty(0x27, VT_BOOL, propval);
}

BOOL CMP11::GetbSplitsInProgressTasks()
{
	BOOL result;
	GetProperty(0x27, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbSpreadActualCost(BOOL propval)
{
	SetProperty(0x28, VT_BOOL, propval);
}

BOOL CMP11::GetbSpreadActualCost()
{
	BOOL result;
	GetProperty(0x28, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbSpreadPercentComplete(BOOL propval)
{
	SetProperty(0x29, VT_BOOL, propval);
}

BOOL CMP11::GetbSpreadPercentComplete()
{
	BOOL result;
	GetProperty(0x29, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbTaskUpdatesResource(BOOL propval)
{
	SetProperty(0x2a, VT_BOOL, propval);
}

BOOL CMP11::GetbTaskUpdatesResource()
{
	BOOL result;
	GetProperty(0x2a, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbFiscalYearStart(BOOL propval)
{
	SetProperty(0x2b, VT_BOOL, propval);
}

BOOL CMP11::GetbFiscalYearStart()
{
	BOOL result;
	GetProperty(0x2b, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetyWeekStartDay(LONG propval)
{
	SetProperty(0x2c, VT_I4, propval);
}

LONG CMP11::GetyWeekStartDay()
{
	LONG result;
	GetProperty(0x2c, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetbMoveCompletedEndsBack(BOOL propval)
{
	SetProperty(0x2d, VT_BOOL, propval);
}

BOOL CMP11::GetbMoveCompletedEndsBack()
{
	BOOL result;
	GetProperty(0x2d, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbMoveRemainingStartsBack(BOOL propval)
{
	SetProperty(0x2e, VT_BOOL, propval);
}

BOOL CMP11::GetbMoveRemainingStartsBack()
{
	BOOL result;
	GetProperty(0x2e, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbMoveRemainingStartsForward(BOOL propval)
{
	SetProperty(0x2f, VT_BOOL, propval);
}

BOOL CMP11::GetbMoveRemainingStartsForward()
{
	BOOL result;
	GetProperty(0x2f, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbMoveCompletedEndsForward(BOOL propval)
{
	SetProperty(0x30, VT_BOOL, propval);
}

BOOL CMP11::GetbMoveCompletedEndsForward()
{
	BOOL result;
	GetProperty(0x30, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetyBaselineForEarnedValue(LONG propval)
{
	SetProperty(0x31, VT_I4, propval);
}

LONG CMP11::GetyBaselineForEarnedValue()
{
	LONG result;
	GetProperty(0x31, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetbAutoAddNewResourcesAndTasks(BOOL propval)
{
	SetProperty(0x32, VT_BOOL, propval);
}

BOOL CMP11::GetbAutoAddNewResourcesAndTasks()
{
	BOOL result;
	GetProperty(0x32, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetdtStatusDate(DATE propval)
{
	SetProperty(0x33, VT_DATE, propval);
}

DATE CMP11::GetdtStatusDate()
{
	DATE result;
	GetProperty(0x33, VT_DATE, (void*)&result);
	return result;
}

void CMP11::SetdtCurrentDate(DATE propval)
{
	SetProperty(0x34, VT_DATE, propval);
}

DATE CMP11::GetdtCurrentDate()
{
	DATE result;
	GetProperty(0x34, VT_DATE, (void*)&result);
	return result;
}

void CMP11::SetbMicrosoftProjectServerURL(BOOL propval)
{
	SetProperty(0x35, VT_BOOL, propval);
}

BOOL CMP11::GetbMicrosoftProjectServerURL()
{
	BOOL result;
	GetProperty(0x35, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbAutolink(BOOL propval)
{
	SetProperty(0x36, VT_BOOL, propval);
}

BOOL CMP11::GetbAutolink()
{
	BOOL result;
	GetProperty(0x36, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetyNewTaskStartDate(LONG propval)
{
	SetProperty(0x37, VT_I4, propval);
}

LONG CMP11::GetyNewTaskStartDate()
{
	LONG result;
	GetProperty(0x37, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetyDefaultTaskEVMethod(LONG propval)
{
	SetProperty(0x38, VT_I4, propval);
}

LONG CMP11::GetyDefaultTaskEVMethod()
{
	LONG result;
	GetProperty(0x38, VT_I4, (void*)&result);
	return result;
}

void CMP11::SetbProjectExternallyEdited(BOOL propval)
{
	SetProperty(0x39, VT_BOOL, propval);
}

BOOL CMP11::GetbProjectExternallyEdited()
{
	BOOL result;
	GetProperty(0x39, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetdtExtendedCreationDate(DATE propval)
{
	SetProperty(0x3a, VT_DATE, propval);
}

DATE CMP11::GetdtExtendedCreationDate()
{
	DATE result;
	GetProperty(0x3a, VT_DATE, (void*)&result);
	return result;
}

void CMP11::SetbActualsInSync(BOOL propval)
{
	SetProperty(0x3b, VT_BOOL, propval);
}

BOOL CMP11::GetbActualsInSync()
{
	BOOL result;
	GetProperty(0x3b, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbRemoveFileProperties(BOOL propval)
{
	SetProperty(0x3c, VT_BOOL, propval);
}

BOOL CMP11::GetbRemoveFileProperties()
{
	BOOL result;
	GetProperty(0x3c, VT_BOOL, (void*)&result);
	return result;
}

void CMP11::SetbAdminProject(BOOL propval)
{
	SetProperty(0x3d, VT_BOOL, propval);
}

BOOL CMP11::GetbAdminProject()
{
	BOOL result;
	GetProperty(0x3d, VT_BOOL, (void*)&result);
	return result;
}

COutlineCodes CMP11::GetoOutlineCodes()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3e, VT_DISPATCH, (void*)&pDispatch);
	return COutlineCodes(pDispatch);
}

CWBSMasks CMP11::GetoWBSMasks()
{
	LPDISPATCH pDispatch;
	GetProperty(0x3f, VT_DISPATCH, (void*)&pDispatch);
	return CWBSMasks(pDispatch);
}

CExtendedAttributes CMP11::GetoExtendedAttributes()
{
	LPDISPATCH pDispatch;
	GetProperty(0x40, VT_DISPATCH, (void*)&pDispatch);
	return CExtendedAttributes(pDispatch);
}

CCalendars CMP11::GetoCalendars()
{
	LPDISPATCH pDispatch;
	GetProperty(0x41, VT_DISPATCH, (void*)&pDispatch);
	return CCalendars(pDispatch);
}

CTasks CMP11::GetoTasks()
{
	LPDISPATCH pDispatch;
	GetProperty(0x42, VT_DISPATCH, (void*)&pDispatch);
	return CTasks(pDispatch);
}

CResources CMP11::GetoResources()
{
	LPDISPATCH pDispatch;
	GetProperty(0x43, VT_DISPATCH, (void*)&pDispatch);
	return CResources(pDispatch);
}

CAssignments CMP11::GetoAssignments()
{
	LPDISPATCH pDispatch;
	GetProperty(0x44, VT_DISPATCH, (void*)&pDispatch);
	return CAssignments(pDispatch);
}

void CMP11::WriteXML(LPCTSTR url)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x45, DISPATCH_METHOD, VT_EMPTY, NULL, params, url);
}

void CMP11::ReadXML(LPCTSTR url)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x46, DISPATCH_METHOD, VT_EMPTY, NULL, params, url);
}

CString CMP11::GetXML(void)
{
	CString oReturn;
	InvokeHelper(0x47, DISPATCH_METHOD, VT_BSTR, (void*)&oReturn, NULL);
	return oReturn;
}

void CMP11::SetXML(LPCTSTR sXML)
{
	static BYTE params[] = VTS_BSTR;
	InvokeHelper(0x48, DISPATCH_METHOD, VT_EMPTY, NULL, params, sXML);
}

BOOL CMP11::IsNull(void)
{
	BOOL oReturn;
	InvokeHelper(0x49, DISPATCH_METHOD, VT_BOOL, (void*)&oReturn, NULL);
	return oReturn;
}

void CMP11::Initialize(void)
{
	InvokeHelper(0x4a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CMP11 CMSP2003Ctl::GetMP11()
{
	LPDISPATCH pDispatch;
	GetProperty(0x2, VT_DISPATCH, (void*)&pDispatch);
	return CMP11(pDispatch);
}

void CMSP2003Ctl::AboutBox(void)
{
	InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CMSP2003Ctl::Clear(void)
{
	InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}
}
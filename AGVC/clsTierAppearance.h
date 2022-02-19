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

class clsTierColors;

class clsTierAppearance : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTierAppearance)

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTierAppearance(CActiveGanttVCCtl* oControl);
	clsTierAppearance();
	virtual ~clsTierAppearance();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsTierColors* mp_oMicrosecondColors;
	clsTierColors* mp_oMillisecondColors;
	clsTierColors* mp_oSecondColors;				
	clsTierColors* mp_oMinuteColors;			
	clsTierColors* mp_oHourColors;				
	clsTierColors* mp_oDayColors;				
	clsTierColors* mp_oDayOfWeekColors;				
	clsTierColors* mp_oDayOfYearColors;				
	clsTierColors* mp_oWeekColors;				
	clsTierColors* mp_oMonthColors;				
	clsTierColors* mp_oQuarterColors;				
	clsTierColors* mp_oYearColors;			

//Internal Properties (Extensions to ODL)
public:
	clsTierColors* GetMicrosecondColors(void);
	clsTierColors* GetMillisecondColors(void);
	clsTierColors* GetSecondColors(void);
	clsTierColors* GetMinuteColors(void);
	clsTierColors* GetHourColors(void);
	clsTierColors* GetDayColors(void);
	clsTierColors* GetDayOfWeekColors(void);
	clsTierColors* GetDayOfYearColors(void);
	clsTierColors* GetWeekColors(void);
	clsTierColors* GetMonthColors(void);
	clsTierColors* GetQuarterColors(void);
	clsTierColors* GetYearColors(void);

//Internal Properties
public:

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTierAppearance* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTierAppearance)
	DECLARE_INTERFACE_MAP()

IDispatch* odl_GetMicrosecondColors(void)
{
	
	return GetMicrosecondColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetMillisecondColors(void)
{
	
	return GetMillisecondColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetSecondColors(void)
{
	
	return GetSecondColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetMinuteColors(void)
{
	
	return GetMinuteColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetHourColors(void)
{
	
	return GetHourColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDayColors(void)
{
	
	return GetDayColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDayOfWeekColors(void)
{
	
	return GetDayOfWeekColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetDayOfYearColors(void)
{
	
	return GetDayOfYearColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetWeekColors(void)
{
	
	return GetWeekColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetMonthColors(void)
{
	
	return GetMonthColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetQuarterColors(void)
{
	
	return GetQuarterColors()->GetIDispatch(TRUE);
}

IDispatch* odl_GetYearColors(void)
{
	
	return GetYearColors()->GetIDispatch(TRUE);
}

BSTR odl_GetXML(void)
{
	
	return GetXML().AllocSysString();
}

void odl_SetXML(LPCTSTR sXML)
{
	
	SetXML(sXML);
}

};
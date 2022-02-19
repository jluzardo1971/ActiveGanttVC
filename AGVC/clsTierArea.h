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

class clsTier;
class clsTierFormat;
class clsTierAppearance;
class clsTimeLine;

class clsTierArea : public CCmdTarget
{
	DECLARE_DYNCREATE(clsTierArea)
	clsTierArea();

public:
	CActiveGanttVCCtl* mp_oControl;

	clsTierArea(CActiveGanttVCCtl* oControl, clsTimeLine* oTimeLine);
	
	virtual ~clsTierArea();
	virtual void OnFinalRelease();

// Member Variables
public:
	clsTier* mp_oUpperTier;
	clsTier* mp_oMiddleTier;
	clsTier* mp_oLowerTier;
	clsTierFormat* mp_oTierFormat;
	clsTierAppearance* mp_oTierAppearance;
	clsTimeLine* mp_oTimeLine;

//Internal Properties (Extensions to ODL)
public:
	clsTier* GetUpperTier(void);
	clsTier* GetMiddleTier(void);
	clsTier* GetLowerTier(void);
	clsTierFormat* GetTierFormat(void);
	clsTierAppearance* GetTierAppearance(void);
	clsTimeLine* GetTimeLine(void);

//Internal Properties
public:

//Internal Methods
public:
	CString GetXML(void);
	void SetXML(CString sXML);
	void Clear();
	void Clone(clsTierArea* oClone);

protected:
	DECLARE_DISPATCH_MAP()
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsTierArea)
	DECLARE_INTERFACE_MAP()

IDispatch* odl_GetUpperTier(void)
{
	
	return GetUpperTier()->GetIDispatch(TRUE);
}

IDispatch* odl_GetMiddleTier(void)
{
	
	return GetMiddleTier()->GetIDispatch(TRUE);
}

IDispatch* odl_GetLowerTier(void)
{
	
	return GetLowerTier()->GetIDispatch(TRUE);
}

IDispatch* odl_GetTierFormat(void)
{
	
	return GetTierFormat()->GetIDispatch(TRUE);
}

IDispatch* odl_GetTierAppearance(void)
{
	
	return GetTierAppearance()->GetIDispatch(TRUE);
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
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
#include "clsTierArea.h"

IMPLEMENT_DYNCREATE(clsTierArea, CCmdTarget)

//{2CCB666A-3795-4A7C-A17A-CAD772E1775A}
static const IID IID_IclsTierArea = { 0x2CCB666A, 0x3795, 0x4A7C, { 0xA1, 0x7A, 0xCA, 0xD7, 0x72, 0xE1, 0x77, 0x5A} };

//{3B201B8A-7C68-4004-A63D-535495DE9BE5}
IMPLEMENT_OLECREATE_FLAGS(clsTierArea, "AGVC.clsTierArea", afxRegApartmentThreading, 0x3B201B8A, 0x7C68, 0x4004, 0xA6, 0x3D, 0x53, 0x54, 0x95, 0xDE, 0x9B, 0xE5)

BEGIN_DISPATCH_MAP(clsTierArea, CCmdTarget)
	DISP_PROPERTY_EX_ID(clsTierArea, "UpperTier", 1, odl_GetUpperTier, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierArea, "MiddleTier", 2, odl_GetMiddleTier, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierArea, "LowerTier", 3, odl_GetLowerTier, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierArea, "TierFormat", 4, odl_GetTierFormat, SetNotSupported, VT_DISPATCH) 
	DISP_PROPERTY_EX_ID(clsTierArea, "TierAppearance", 5, odl_GetTierAppearance, SetNotSupported, VT_DISPATCH) 
	DISP_FUNCTION_ID(clsTierArea, "GetXML", 6, odl_GetXML, VT_BSTR, VTS_NONE) 
	DISP_FUNCTION_ID(clsTierArea, "SetXML", 7, odl_SetXML, VT_EMPTY, VTS_BSTR) 
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(clsTierArea, CCmdTarget)
	INTERFACE_PART(clsTierArea, IID_IclsTierArea, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(clsTierArea, CCmdTarget)
END_MESSAGE_MAP()

clsTierArea::clsTierArea(CActiveGanttVCCtl* oControl, clsTimeLine* oTimeLine)
{
	EnableAutomation();
	AfxOleLockApp();
	mp_oControl = oControl;
	mp_oTimeLine = oTimeLine;
	mp_oUpperTier = new clsTier(mp_oControl, this, SP_UPPER);
	mp_oMiddleTier = new clsTier(mp_oControl, this, SP_MIDDLE);
	mp_oLowerTier = new clsTier(mp_oControl, this, SP_LOWER);
	mp_oTierFormat = new clsTierFormat(mp_oControl);
	mp_oTierAppearance = new clsTierAppearance(mp_oControl);
	Clear();
}

clsTierArea::clsTierArea()
{
	AfxMessageBox(_T("Invalid constructor has been invoked for clsTierArea. Please inform support@sourcecodestore.com of this error"), MB_OK | MB_ICONSTOP | MB_APPLMODAL);
}

clsTierArea::~clsTierArea()
{
	delete mp_oUpperTier;
	mp_oUpperTier = NULL;
	delete mp_oMiddleTier;
	mp_oMiddleTier = NULL;
	delete mp_oLowerTier;
	mp_oLowerTier = NULL;
	delete mp_oTierFormat;
	mp_oTierFormat = NULL;
	delete mp_oTierAppearance;
	mp_oTierAppearance;
	AfxOleUnlockApp();
}

void clsTierArea::OnFinalRelease()
{
	CCmdTarget::OnFinalRelease();
}

clsTier* clsTierArea::GetUpperTier(void)
{
	return mp_oUpperTier;
}

clsTier* clsTierArea::GetMiddleTier(void)
{
	return mp_oMiddleTier;
}

clsTier* clsTierArea::GetLowerTier(void)
{
	return mp_oLowerTier;
}

clsTierFormat* clsTierArea::GetTierFormat(void)
{
	return mp_oTierFormat;
}

clsTierAppearance* clsTierArea::GetTierAppearance(void)
{
	return mp_oTierAppearance;
}

clsTimeLine* clsTierArea::GetTimeLine(void)
{
	return mp_oTimeLine;
}

CString clsTierArea::GetXML(void)
{
	clsXML oXML(mp_oControl, "TierArea");
	oXML.InitializeWriter();
	oXML.WriteObject(mp_oLowerTier->GetXML());
	oXML.WriteObject(mp_oMiddleTier->GetXML());
	oXML.WriteObject(mp_oTierAppearance->GetXML());
	oXML.WriteObject(mp_oTierFormat->GetXML());
	oXML.WriteObject(mp_oUpperTier->GetXML());
	return oXML.GetXML();
}

void clsTierArea::SetXML(CString sXML)
{
	clsXML oXML(mp_oControl, "TierArea");
	oXML.SetXML(sXML);
	oXML.InitializeReader();
    mp_oLowerTier->SetXML(oXML.ReadObject("LowerTier"));
    mp_oMiddleTier->SetXML(oXML.ReadObject("MiddleTier"));
    mp_oTierAppearance->SetXML(oXML.ReadObject("TierAppearance"));
    mp_oTierFormat->SetXML(oXML.ReadObject("TierFormat"));
    mp_oUpperTier->SetXML(oXML.ReadObject("UpperTier"));
}

void clsTierArea::Clear()
{
    GetUpperTier()->Clear();
    GetMiddleTier()->Clear();
    GetLowerTier()->Clear();
    GetTierFormat()->Clear();
    GetTierAppearance()->Clear();
}

void clsTierArea::Clone(clsTierArea* oClone)
{
    GetUpperTier()->Clone(oClone->GetUpperTier());
    GetMiddleTier()->Clone(oClone->GetMiddleTier());
    GetLowerTier()->Clone(oClone->GetLowerTier());
    GetTierFormat()->Clone(oClone->GetTierFormat());
    GetTierAppearance()->Clone(oClone->GetTierAppearance());
}
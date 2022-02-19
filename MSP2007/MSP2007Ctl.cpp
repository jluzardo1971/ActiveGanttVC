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
#include "MSP2007.h"
#include "MSP2007Ctl.h"
#include "MSP2007Ppg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CMSP2007Ctrl, COleControl)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CMSP2007Ctrl, COleControl)
	//{{AFX_MSG_MAP(CMSP2007Ctrl)
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CMSP2007Ctrl, COleControl)
	//{{AFX_DISPATCH_MAP(CMSP2007Ctrl)
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CMSP2007Ctrl, "AboutBox", 1, AboutBox, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_EX_ID(CMSP2007Ctrl, "MP12", 2, odl_GetMP12, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(CMSP2007Ctrl, "Clear", 3, Clear, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CMSP2007Ctrl, COleControl)
	//{{AFX_EVENT_MAP(CMSP2007Ctrl)
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

BEGIN_PROPPAGEIDS(CMSP2007Ctrl, 1)
	PROPPAGEID(CMSP2007PropPage::guid)
END_PROPPAGEIDS(CMSP2007Ctrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CMSP2007Ctrl, "MSP2007.MSP2007Ctrl",0x3366d66a,0x4f20,0x40c1,0x95,0x80,0x51,0x31,0xab,0x08,0xe2,0x18)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CMSP2007Ctrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DMSP2007 = {0x33e638bf,0x08b6,0x4331,{0xbc,0x28,0xab,0x4d,0x9a,0x6d,0x36,0x09}};
const IID BASED_CODE IID_DMSP2007Events = {0x3442204f,0x4444,0x4056,{0xbf,0x7f,0x47,0xe6,0xe0,0x75,0xbc,0x18}};


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwMSP2007OleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_INVISIBLEATRUNTIME |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CMSP2007Ctrl, IDS_MSP2007, _dwMSP2007OleMisc)


/////////////////////////////////////////////////////////////////////////////
// CMSP2007Ctrl::CMSP2007CtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CMSP2007Ctrl

BOOL CMSP2007Ctrl::CMSP2007CtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_MSP2007,
			IDB_MSP2007_ICON,
			afxRegApartmentThreading,
			_dwMSP2007OleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2007Ctrl::CMSP2007Ctrl - Constructor

CMSP2007Ctrl::CMSP2007Ctrl()
{
	InitializeIIDs(&IID_DMSP2007, &IID_DMSP2007Events);
	mp_oMP12 = new MP12();
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2007Ctrl::~CMSP2007Ctrl - Destructor

CMSP2007Ctrl::~CMSP2007Ctrl()
{
	delete mp_oMP12;
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2007Ctrl::OnDraw - Drawing function

void CMSP2007Ctrl::OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	CBitmap bmpMSP2007Logo;
	bmpMSP2007Logo.LoadBitmap(IDB_MSP2007_LOGO);
	 
	// Calculate bitmap size using a BITMAP structure.
	BITMAP bm;
	bmpMSP2007Logo.GetObject( sizeof(BITMAP), &bm );
	 
	// Create a memory DC, select the bitmap into the
	// memory DC, and BitBlt it into the view.
	CDC dcMem;
	dcMem.CreateCompatibleDC(pdc);
	CBitmap* pbmpOld = dcMem.SelectObject(&bmpMSP2007Logo);
	pdc->BitBlt(0, 0, bm.bmWidth, bm.bmHeight, &dcMem, 0,0, SRCCOPY );
	 
	// Reselect the original bitmap into the memory DC.
	dcMem.SelectObject(pbmpOld);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2007Ctrl::DoPropExchange - Persistence support

void CMSP2007Ctrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

}


/////////////////////////////////////////////////////////////////////////////
// CMSP2007Ctrl::OnResetState - Reset control to default state

void CMSP2007Ctrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2007Ctrl::AboutBox - Display an "About" box to the user

void CMSP2007Ctrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_MSP2007);
	dlgAbout.DoModal();
}


void CMSP2007Ctrl::Clear(void)
{
	delete mp_oMP12;
	mp_oMP12 = new MP12();
}

MP12* CMSP2007Ctrl::GetMP12(void)
{
	return mp_oMP12;
}

IDispatch* CMSP2007Ctrl::odl_GetMP12(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMP12()->GetIDispatch(TRUE);
}
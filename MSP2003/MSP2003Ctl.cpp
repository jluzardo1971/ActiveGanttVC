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
#include "MSP2003.h"
#include "MSP2003Ctl.h"
#include "MSP2003Ppg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CMSP2003Ctrl, COleControl)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CMSP2003Ctrl, COleControl)
	//{{AFX_MSG_MAP(CMSP2003Ctrl)
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CMSP2003Ctrl, COleControl)
	//{{AFX_DISPATCH_MAP(CMSP2003Ctrl)
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CMSP2003Ctrl, "AboutBox", 1, AboutBox, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_EX_ID(CMSP2003Ctrl, "MP11", 2, odl_GetMP11, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(CMSP2003Ctrl, "Clear", 3, Clear, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CMSP2003Ctrl, COleControl)
	//{{AFX_EVENT_MAP(CMSP2003Ctrl)
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

BEGIN_PROPPAGEIDS(CMSP2003Ctrl, 1)
	PROPPAGEID(CMSP2003PropPage::guid)
END_PROPPAGEIDS(CMSP2003Ctrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CMSP2003Ctrl, "MSP2003.MSP2003Ctrl",0xaf1509e7,0x4a8c,0x4306,0x81,0x16,0xe5,0x11,0xd4,0xff,0x55,0x82)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CMSP2003Ctrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DMSP2003 = {0xb0b2a085,0x471b,0x4b3f,{0xbc,0xa1,0xd1,0x73,0xfd,0x22,0x42,0x0f}};
const IID BASED_CODE IID_DMSP2003Events = {0x64a23dec,0xe58e,0x4ea0,{0xa6,0x46,0x79,0xa8,0xde,0x16,0x2f,0x7c}};


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwMSP2003OleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_INVISIBLEATRUNTIME |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CMSP2003Ctrl, IDS_MSP2003, _dwMSP2003OleMisc)


/////////////////////////////////////////////////////////////////////////////
// CMSP2003Ctrl::CMSP2003CtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CMSP2003Ctrl

BOOL CMSP2003Ctrl::CMSP2003CtrlFactory::UpdateRegistry(BOOL bRegister)
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
			IDS_MSP2003,
			IDB_MSP2003_ICON,
			afxRegApartmentThreading,
			_dwMSP2003OleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2003Ctrl::CMSP2003Ctrl - Constructor

CMSP2003Ctrl::CMSP2003Ctrl()
{
	InitializeIIDs(&IID_DMSP2003, &IID_DMSP2003Events);
	mp_oMP11 = new MP11();
	g_MSP2003 = this;
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2003Ctrl::~CMSP2003Ctrl - Destructor

CMSP2003Ctrl::~CMSP2003Ctrl()
{
	delete mp_oMP11;
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2003Ctrl::OnDraw - Drawing function

void CMSP2003Ctrl::OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	CBitmap bmpMSP2003Logo;
	bmpMSP2003Logo.LoadBitmap(IDB_MSP2003_LOGO);
	 
	// Calculate bitmap size using a BITMAP structure.
	BITMAP bm;
	bmpMSP2003Logo.GetObject( sizeof(BITMAP), &bm );
	 
	// Create a memory DC, select the bitmap into the
	// memory DC, and BitBlt it into the view.
	CDC dcMem;
	dcMem.CreateCompatibleDC(pdc);
	CBitmap* pbmpOld = dcMem.SelectObject(&bmpMSP2003Logo);
	pdc->BitBlt(0, 0, bm.bmWidth, bm.bmHeight, &dcMem, 0,0, SRCCOPY );
	 
	// Reselect the original bitmap into the memory DC.
	dcMem.SelectObject(pbmpOld);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2003Ctrl::DoPropExchange - Persistence support

void CMSP2003Ctrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

}


/////////////////////////////////////////////////////////////////////////////
// CMSP2003Ctrl::OnResetState - Reset control to default state

void CMSP2003Ctrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2003Ctrl::AboutBox - Display an "About" box to the user

void CMSP2003Ctrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_MSP2003);
	dlgAbout.DoModal();
}


void CMSP2003Ctrl::Clear(void)
{
	delete mp_oMP11;
	mp_oMP11 = new MP11();
}

MP11* CMSP2003Ctrl::GetMP11(void)
{
	return mp_oMP11;
}

IDispatch* CMSP2003Ctrl::odl_GetMP11(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMP11()->GetIDispatch(TRUE);
}
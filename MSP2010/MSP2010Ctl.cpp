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
#include "MSP2010.h"
#include "MSP2010Ctl.h"
#include "MSP2010Ppg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CMSP2010Ctrl, COleControl)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CMSP2010Ctrl, COleControl)
	//{{AFX_MSG_MAP(CMSP2010Ctrl)
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CMSP2010Ctrl, COleControl)
	//{{AFX_DISPATCH_MAP(CMSP2010Ctrl)
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CMSP2010Ctrl, "AboutBox", 1, AboutBox, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_EX_ID(CMSP2010Ctrl, "MP14", 2, odl_GetMP14, SetNotSupported, VT_DISPATCH)
	DISP_FUNCTION_ID(CMSP2010Ctrl, "Clear", 3, Clear, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CMSP2010Ctrl, COleControl)
	//{{AFX_EVENT_MAP(CMSP2010Ctrl)
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

BEGIN_PROPPAGEIDS(CMSP2010Ctrl, 1)
	PROPPAGEID(CMSP2010PropPage::guid)
END_PROPPAGEIDS(CMSP2010Ctrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CMSP2010Ctrl, "MSP2010.MSP2010Ctrl",0xc56e2638,0x476c,0x44c6,0xb2,0x1f,0x91,0xfd,0x43,0x48,0xb8,0xa5)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CMSP2010Ctrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DMSP2010 = {0x7faf00a1,0x838b,0x4834,{0x97,0xc1,0x1b,0x35,0x20,0xbb,0x05,0x8e}};
const IID BASED_CODE IID_DMSP2010Events = {0xd5155f70,0x1c8f,0x4091,{0xbc,0x84,0xee,0x66,0x94,0x30,0x7b,0x10}};


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwMSP2010OleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_INVISIBLEATRUNTIME |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CMSP2010Ctrl, IDS_MSP2010, _dwMSP2010OleMisc)


/////////////////////////////////////////////////////////////////////////////
// CMSP2010Ctrl::CMSP2010CtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CMSP2010Ctrl

BOOL CMSP2010Ctrl::CMSP2010CtrlFactory::UpdateRegistry(BOOL bRegister)
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
			IDS_MSP2010,
			IDB_MSP2010_ICON,
			afxRegApartmentThreading,
			_dwMSP2010OleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2010Ctrl::CMSP2010Ctrl - Constructor

CMSP2010Ctrl::CMSP2010Ctrl()
{
	InitializeIIDs(&IID_DMSP2010, &IID_DMSP2010Events);
	mp_oMP14 = new MP14();
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2010Ctrl::~CMSP2010Ctrl - Destructor

CMSP2010Ctrl::~CMSP2010Ctrl()
{
	delete mp_oMP14;
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2010Ctrl::OnDraw - Drawing function

void CMSP2010Ctrl::OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	CBitmap bmpMSP2010Logo;
	bmpMSP2010Logo.LoadBitmap(IDB_MSP2010_LOGO);
	 
	// Calculate bitmap size using a BITMAP structure.
	BITMAP bm;
	bmpMSP2010Logo.GetObject( sizeof(BITMAP), &bm );
	 
	// Create a memory DC, select the bitmap into the
	// memory DC, and BitBlt it into the view.
	CDC dcMem;
	dcMem.CreateCompatibleDC(pdc);
	CBitmap* pbmpOld = dcMem.SelectObject(&bmpMSP2010Logo);
	pdc->BitBlt(0, 0, bm.bmWidth, bm.bmHeight, &dcMem, 0,0, SRCCOPY );
	 
	// Reselect the original bitmap into the memory DC.
	dcMem.SelectObject(pbmpOld);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2010Ctrl::DoPropExchange - Persistence support

void CMSP2010Ctrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

}


/////////////////////////////////////////////////////////////////////////////
// CMSP2010Ctrl::OnResetState - Reset control to default state

void CMSP2010Ctrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2010Ctrl::AboutBox - Display an "About" box to the user

void CMSP2010Ctrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_MSP2010);
	dlgAbout.DoModal();
}


void CMSP2010Ctrl::Clear(void)
{
	delete mp_oMP14;
	mp_oMP14 = new MP14();
}

MP14* CMSP2010Ctrl::GetMP14(void)
{
	return mp_oMP14;
}

IDispatch* CMSP2010Ctrl::odl_GetMP14(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMP14()->GetIDispatch(TRUE);
}
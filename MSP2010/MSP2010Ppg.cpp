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
#include "MSP2010Ppg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CMSP2010PropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CMSP2010PropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CMSP2010PropPage)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CMSP2010PropPage, "MSP2010.MSP2010PropPage.1",
	0xf09e5580, 0xb002, 0x4e00, 0xab, 0x8e, 0xc7, 0x29, 0x19, 0x1f, 0x98, 0xfa)


/////////////////////////////////////////////////////////////////////////////
// CMSP2010PropPage::CMSP2010PropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CMSP2010PropPage

BOOL CMSP2010PropPage::CMSP2010PropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_MSP2010_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2010PropPage::CMSP2010PropPage - Constructor

CMSP2010PropPage::CMSP2010PropPage() :
	COlePropertyPage(IDD, IDS_MSP2010_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CMSP2010PropPage)
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2010PropPage::DoDataExchange - Moves data between page and properties

void CMSP2010PropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CMSP2010PropPage)
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2010PropPage message handlers
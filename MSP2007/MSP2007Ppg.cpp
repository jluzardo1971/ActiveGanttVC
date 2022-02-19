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
#include "MSP2007Ppg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CMSP2007PropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CMSP2007PropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CMSP2007PropPage)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CMSP2007PropPage, "MSP2007.MSP2007PropPage.1",
	0xb00f005e, 0x637c, 0x4bed, 0x8f, 0x10, 0x99, 0x64, 0x71, 0xd2, 0xd9, 0x37)


/////////////////////////////////////////////////////////////////////////////
// CMSP2007PropPage::CMSP2007PropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CMSP2007PropPage

BOOL CMSP2007PropPage::CMSP2007PropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_MSP2007_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2007PropPage::CMSP2007PropPage - Constructor

CMSP2007PropPage::CMSP2007PropPage() :
	COlePropertyPage(IDD, IDS_MSP2007_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CMSP2007PropPage)
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2007PropPage::DoDataExchange - Moves data between page and properties

void CMSP2007PropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CMSP2007PropPage)
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2007PropPage message handlers
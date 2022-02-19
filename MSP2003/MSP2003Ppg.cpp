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
#include "MSP2003Ppg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CMSP2003PropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CMSP2003PropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CMSP2003PropPage)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CMSP2003PropPage, "MSP2003.MSP2003PropPage.1",
	0x7b01b562, 0xeae5, 0x4179, 0x88, 0x3d, 0xbb, 0x63, 0xc3, 0x10, 0x97, 0xd5)


/////////////////////////////////////////////////////////////////////////////
// CMSP2003PropPage::CMSP2003PropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CMSP2003PropPage

BOOL CMSP2003PropPage::CMSP2003PropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_MSP2003_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2003PropPage::CMSP2003PropPage - Constructor

CMSP2003PropPage::CMSP2003PropPage() :
	COlePropertyPage(IDD, IDS_MSP2003_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CMSP2003PropPage)
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2003PropPage::DoDataExchange - Moves data between page and properties

void CMSP2003PropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CMSP2003PropPage)
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CMSP2003PropPage message handlers
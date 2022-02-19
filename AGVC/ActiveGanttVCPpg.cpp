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
#include "ActiveGanttVCPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

IMPLEMENT_DYNCREATE(CActiveGanttVCPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CActiveGanttVCPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CActiveGanttVCPropPage)
	// NOTE - ClassWizard will add and remove message map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CActiveGanttVCPropPage, "ActiveGanttVC.ActiveGanttVCPropPage",
	0x87ee9621, 0xe6a5, 0x4ab7, 0xa8, 0xe, 0xcd, 0x28, 0x3e, 0xbd, 0x19, 0x89)


/////////////////////////////////////////////////////////////////////////////
// CActiveGanttVCPropPage::CActiveGanttVCPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CActiveGanttVCPropPage

BOOL CActiveGanttVCPropPage::CActiveGanttVCPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_ACTIVEGANTTVC_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CActiveGanttVCPropPage::CActiveGanttVCPropPage - Constructor

CActiveGanttVCPropPage::CActiveGanttVCPropPage() :
	COlePropertyPage(IDD, IDS_ACTIVEGANTTVC_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CActiveGanttVCPropPage)
	// NOTE: ClassWizard will add member initialization here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CActiveGanttVCPropPage::DoDataExchange - Moves data between page and properties

void CActiveGanttVCPropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CActiveGanttVCPropPage)
	// NOTE: ClassWizard will add DDP, DDX, and DDV calls here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}
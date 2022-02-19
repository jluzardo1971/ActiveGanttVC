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
#if !defined(AFX_MSP2003PPG_H__B4C70048_3E7E_485E_BA18_64AF0C062FA1__INCLUDED_)
#define AFX_MSP2003PPG_H__B4C70048_3E7E_485E_BA18_64AF0C062FA1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// MSP2003Ppg.h : Declaration of the CMSP2003PropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CMSP2003PropPage : See MSP2003Ppg.cpp.cpp for implementation.

class CMSP2003PropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CMSP2003PropPage)
	DECLARE_OLECREATE_EX(CMSP2003PropPage)

// Constructor
public:
	CMSP2003PropPage();

// Dialog Data
	//{{AFX_DATA(CMSP2003PropPage)
	enum { IDD = IDD_PROPPAGE_MSP2003 };
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CMSP2003PropPage)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MSP2003PPG_H__B4C70048_3E7E_485E_BA18_64AF0C062FA1__INCLUDED)
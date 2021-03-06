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
#if !defined(AFX_MSP2010PPG_H__B1C74834_DDC6_4AF7_90D5_48541DACED20__INCLUDED_)
#define AFX_MSP2010PPG_H__B1C74834_DDC6_4AF7_90D5_48541DACED20__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// MSP2010Ppg.h : Declaration of the CMSP2010PropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CMSP2010PropPage : See MSP2010Ppg.cpp.cpp for implementation.

class CMSP2010PropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CMSP2010PropPage)
	DECLARE_OLECREATE_EX(CMSP2010PropPage)

// Constructor
public:
	CMSP2010PropPage();

// Dialog Data
	//{{AFX_DATA(CMSP2010PropPage)
	enum { IDD = IDD_PROPPAGE_MSP2010 };
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CMSP2010PropPage)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MSP2010PPG_H__B1C74834_DDC6_4AF7_90D5_48541DACED20__INCLUDED)
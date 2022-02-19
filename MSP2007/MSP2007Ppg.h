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
#if !defined(AFX_MSP2007PPG_H__5BE3DF8A_1123_452F_83F3_39BF48C516A2__INCLUDED_)
#define AFX_MSP2007PPG_H__5BE3DF8A_1123_452F_83F3_39BF48C516A2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// MSP2007Ppg.h : Declaration of the CMSP2007PropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CMSP2007PropPage : See MSP2007Ppg.cpp.cpp for implementation.

class CMSP2007PropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CMSP2007PropPage)
	DECLARE_OLECREATE_EX(CMSP2007PropPage)

// Constructor
public:
	CMSP2007PropPage();

// Dialog Data
	//{{AFX_DATA(CMSP2007PropPage)
	enum { IDD = IDD_PROPPAGE_MSP2007 };
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CMSP2007PropPage)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MSP2007PPG_H__5BE3DF8A_1123_452F_83F3_39BF48C516A2__INCLUDED)
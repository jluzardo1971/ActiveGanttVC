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
#pragma once


// MSP2007Ctl.h : Declaration of the CMSP2007Ctrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CMSP2007Ctrl : See MSP2007Ctl.cpp for implementation.

class CMSP2007Ctrl : public COleControl
{
	DECLARE_DYNCREATE(CMSP2007Ctrl)

// Constructor
public:
	CMSP2007Ctrl();

	MP12* mp_oMP12;
	void Clear(void);
	MP12* GetMP12(void);
	IDispatch* odl_GetMP12(void);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMSP2007Ctrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	//}}AFX_VIRTUAL


// Implementation
protected:
	~CMSP2007Ctrl();

	DECLARE_OLECREATE_EX(CMSP2007Ctrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CMSP2007Ctrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CMSP2007Ctrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CMSP2007Ctrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CMSP2007Ctrl)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CMSP2007Ctrl)
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

public:

	void FireControlError(long Number, LPCTSTR Description, LPCTSTR Source)
	{
		FireEvent(1, EVENT_PARAM(VTS_I4  VTS_BSTR  VTS_BSTR), Number, Description, Source);
	}
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	enum {
	//{{AFX_DISP_ID(CMSP2007Ctrl)
	//}}AFX_DISP_ID
	};
};
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


// MSP2010Ctl.h : Declaration of the CMSP2010Ctrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CMSP2010Ctrl : See MSP2010Ctl.cpp for implementation.

class CMSP2010Ctrl : public COleControl
{
	DECLARE_DYNCREATE(CMSP2010Ctrl)

// Constructor
public:
	CMSP2010Ctrl();

	MP14* mp_oMP14;
	void Clear(void);
	MP14* GetMP14(void);
	IDispatch* odl_GetMP14(void);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMSP2010Ctrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	//}}AFX_VIRTUAL

// Implementation
protected:
	~CMSP2010Ctrl();

	DECLARE_OLECREATE_EX(CMSP2010Ctrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CMSP2010Ctrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CMSP2010Ctrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CMSP2010Ctrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CMSP2010Ctrl)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CMSP2010Ctrl)
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
	//{{AFX_DISP_ID(CMSP2010Ctrl)
	//}}AFX_DISP_ID
	};
};
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
#include "afxwin.h"

class fWBSProject;


// fWBSPProperties dialog

class fWBSPProperties : public CDialog
{
	DECLARE_DYNAMIC(fWBSPProperties)

public:
	fWBSPProperties(fWBSProject* oParent);   // standard constructor
	virtual ~fWBSPProperties();

	fWBSProject* mp_oParent;

// Dialog Data
	enum { IDD = IDD_FWBSPPROPERTIES };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CButton chkEnforcePredecessors;
	CComboBox cboPredecessorMode;
	virtual BOOL OnInitDialog();
protected:
	virtual void OnOK();
};
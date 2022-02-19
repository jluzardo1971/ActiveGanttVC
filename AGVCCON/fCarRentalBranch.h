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

class fCarRental;

class fCarRentalBranch : public CDialog
{
	DECLARE_DYNAMIC(fCarRentalBranch)

public:
	fCarRentalBranch(CWnd* pParent = NULL);   
	virtual ~fCarRentalBranch();

	PRG_DIALOGMODE mp_yDialogMode;
	CString mp_sRowID;
	fCarRental* mp_oParent;

	CEdit txtBranchName;
	CEdit txtAddress;
	CEdit txtCity;
	CEdit txtState;
	CEdit txtZIP;
	CEdit txtPhone;
	CEdit txtManagerName;
	CEdit txtManagerMobile;


	enum { IDD = IDD_FCARRENTALBRANCH };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
protected:
	virtual void OnOK();
};
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

#include "atlimage.h" //CImage

class fCarRental;

class fCarRentalVehicle : public CDialog
{
	DECLARE_DYNAMIC(fCarRentalVehicle)

public:
	fCarRentalVehicle(CWnd* pParent = NULL);   
	virtual ~fCarRentalVehicle();

	PRG_DIALOGMODE mp_yDialogMode;
	CString mp_sRowID;
	fCarRental* mp_oParent;

	CImage mp_oImage;

	void UpdateACRISSCode(CString sACRISSCode);

	void DeleteComboBoxItems();
	void UpdatePicture();


	enum { IDD = IDD_FCARRENTALVEHICLE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    

	DECLARE_MESSAGE_MAP()
public:
	CComboBox drpCarTypeID;
	CEdit txtLicensePlates;
	CComboBox drpACRISS1;
	CComboBox drpACRISS2;
	CComboBox drpACRISS3;
	CComboBox drpACRISS4;
	CEdit txtRate;
	CStatic lblACRISS1;
	CStatic lblACRISS2;
	CStatic lblACRISS3;
	CStatic lblACRISS4;
	CEdit txtNotes;
	CStatic picMake;
	virtual BOOL OnInitDialog();
protected:
	virtual void OnOK();
public:

protected:
	virtual void OnCancel();
public:
	afx_msg void OnCbnSelchangeDrpcartypeid();
	afx_msg void OnCbnSelchangeDrpacriss1();
	afx_msg void OnCbnSelchangeDrpacriss2();
	afx_msg void OnCbnSelchangeDrpacriss3();
	afx_msg void OnCbnSelchangeDrpacriss4();
};
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

class CActiveGanttVCPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CActiveGanttVCPropPage)
	DECLARE_OLECREATE_EX(CActiveGanttVCPropPage)

// Constructor
public:
	CActiveGanttVCPropPage();

	enum { IDD = IDD_PROPPAGE_ACTIVEGANTTVC };


// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);

// Message maps
protected:
	DECLARE_MESSAGE_MAP()

};
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

class clsItemBase : public CCmdTarget
{
public:
	DECLARE_DYNCREATE(clsItemBase)
	clsItemBase();
	virtual ~clsItemBase();

public:


public:
	long mp_lIndex;
	CString mp_sKey;

	//{{AFX_VIRTUAL(clsItemBase)
	//}}AFX_VIRTUAL

protected:
	

	//{{AFX_MSG(clsItemBase)

	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};


//{{AFX_INSERT_LOCATION}}
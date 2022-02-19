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


// clsCultureInfo command target

class clsCultureInfo : public CCmdTarget
{
	DECLARE_DYNCREATE(clsCultureInfo)

public:
	clsCultureInfo();
	virtual ~clsCultureInfo();

	virtual void OnFinalRelease();

protected:
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(clsCultureInfo)
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};
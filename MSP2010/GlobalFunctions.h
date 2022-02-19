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

class CMSP2010Ctrl;

static CMSP2010Ctrl* g_MSP2010;

CString F_TOUCSTR(LPCSTR sParam);
BOOL g_IsNumeric(CString sParam);
void g_ErrorReport(LONG ErrNumber,CString ErrDescription,CString ErrSource);
LONG g_CInt(CString sParam);
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

//MSVC++ 9.0  _MSC_VER = 1500 
//MSVC++ 8.0  _MSC_VER = 1400 
//MSVC++ 7.1  _MSC_VER = 1310 
//MSVC++ 7.0  _MSC_VER = 1300 
//MSVC++ 6.0  _MSC_VER = 1200
//MSVC++ 5.0  _MSC_VER = 1100

#if _MSC_VER <= 1310 //Control Operating System Compatibility
	#ifndef WINVER
	#define  WINVER  0x0400 //Windows 95 & Windows NT
	#endif
#else
	#ifndef WINVER    
	#define WINVER 0x0501  // Windows XP
	#endif
#endif

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afxctl.h>         // MFC support for ActiveX Controls
#include <afxext.h>         // MFC extensions
#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Comon Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

#define _CRTDBG_MAP_ALLOC
#include <stdlib.h>
#include <crtdbg.h>

#include <MsXml6.h>




#pragma comment( lib, "gdiplus.lib" )
#include "initgdiplus.h"
#include <gdiplus.h>
#include <afxtempl.h> //CArray Support

#include <vector>			// Standard Library std::vector 

#include "ActiveGanttVCWrappers.h"

#include "Globals.h"
#include "PrivateVars.h"
#include "GlobalFunctions.h"
#include "GlobalDateFunctions.h"

#include "resource.h"
#include "fAbout.h"




#include "clsCollectionBase.h"
#include "clsItemBase.h"
#include "clsMouseKeyboardEvents.h"
#include "clsImage.h"
#include "clsFont.h"
#include "clsGDIGraphics.h"
#include "clsGraphics.h"
#include "clsCultureInfo.h"

#include "clsTaskStyle.h"
#include "clsMilestoneStyle.h"
#include "clsPredecessorStyle.h"
#include "clsTextFlags.h"
#include "clsCustomBorderStyle.h"
#include "clsSelectionRectangleStyle.h"
#include "clsButtonBorderStyle.h"
#include "clsScrollBarStyle.h"
#include "clsStyle.h"
#include "clsStyles.h"

#include "ControlTemplateGradient.h"
#include "ControlTemplateSolid.h"

#include "Templates.h"


#include "clsPrinter.h"
#include "clsSplitter.h"
#include "clsTreeview.h"
#include "clsToolTip.h"

#include "clsLayer.h"
#include "clsLayers.h"

#include "clsNode.h"
#include "clsCell.h"
#include "clsCells.h"
#include "clsRow.h"
#include "clsRows.h"

#include "clsTask.h"
#include "clsTasks.h"
#include "clsPredecessor.h"
#include "clsPredecessors.h"
#include "clsTimeBlock.h"
#include "clsTimeBlocks.h"




#include "clsColumn.h"
#include "clsColumns.h"






#include "clsButtonState.h"
#include "clsHScrollBarTemplate.h"
#include "clsVScrollBarTemplate.h"
#include "clsVerticalScrollBar.h"
#include "clsHorizontalScrollBar.h"


#include "clsScrollBarSeparator.h"
#include "clsPercentage.h"
#include "clsPercentages.h"

#include "clsGrid.h"
#include "clsClientArea.h"

#include "clsTier.h"
#include "clsTierFormat.h"
#include "clsTierColor.h"
#include "clsTierColors.h"
#include "clsTierAppearance.h"
#include "clsTierArea.h"
#include "clsTimeLineScrollBar.h"
#include "clsTickMark.h"
#include "clsTickMarks.h"
#include "clsTickMarkArea.h"
#include "clsProgressLine.h"
#include "clsTimeLine.h"








#include "clsView.h"
#include "clsViews.h"



#include "CustomTierDrawEventArgs.h"
#include "DrawEventArgs.h"
#include "ErrorEventArgs.h"
#include "KeyEventArgs.h"
#include "MouseEventArgs.h"
#include "MouseWheelEventArgs.h"
#include "NodeEventArgs.h"
#include "ObjectAddedEventArgs.h"
#include "ObjectSelectedEventArgs.h"
#include "ObjectStateChangedEventArgs.h"
#include "PredecessorDrawEventArgs.h"
#include "ScrollEventArgs.h"
#include "ToolTipEventArgs.h"
#include "TextEditEventArgs.h"
#include "PredecessorExceptionEventArgs.h"
#include "TickMarkTextDrawEventArgs.h"
#include "CustomTickMarkAreaDrawEventArgs.h"
#include "PagePrintEventArgs.h"

#include "clsMath.h"
#include "clsDrawing.h"

#include "clsTextBox.h"

#include "clsConfiguration.h"


#include <afxwin.h>
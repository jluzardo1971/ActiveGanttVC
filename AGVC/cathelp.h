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
      #include <comcat.h>

      // Helper function to create a component category and associated
      // description
      HRESULT CreateComponentCategory(CATID catid, WCHAR* catDescription);

      // Helper function to register a CLSID as belonging to a component
      // category
      HRESULT RegisterCLSIDInCategory(REFCLSID clsid, CATID catid);
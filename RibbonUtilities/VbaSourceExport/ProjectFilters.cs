////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [CLSCompliant(false)]
    public class ProjectFilters : List<ProjectFilter> {
        public ProjectFilters(IApplication application) {
            Add(new ProjectFilterExcel(application, 
                    "MS-Excel Projects", "*.xlsm;*.xlsb;*.xlam;*.xls;*.xla"));

            if (AccessWrapper.IsAccessSupported) {
                Add(new ProjectFilterAccess(application, 
                        "MS-Access Projects", "*.accdb;*.accda;*.mdb;*.mda"));
            }
        }
    }
}

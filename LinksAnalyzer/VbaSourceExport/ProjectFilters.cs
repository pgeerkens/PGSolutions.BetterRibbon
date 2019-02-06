////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    internal class ProjectFilters : List<ProjectFilter> {
        public ProjectFilters(Excel.Application application) {
            Add(new ProjectFilterExcel(application, 
                    "MS-Excel Projects", "*.xlsm;*.xlsb;*.xlam;*.xls;*.xla"));

            if (AccessWrapper.IsAccessSupported) {
                Add(new ProjectFilterAccess(application, 
                        "MS-Access Projects", "*.accdb;*.accda;*.mdb;*.mda"));
            }
        }
    }
}

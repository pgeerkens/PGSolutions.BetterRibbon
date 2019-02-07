////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
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

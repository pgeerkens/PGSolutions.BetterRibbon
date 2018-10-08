using System.Collections.Generic;

namespace PGSolutions.ExcelRibbon2013 {
    internal class ProjectFilters : List<ProjectFilter> {
        public ProjectFilters() {
            Add(new ProjectFilterExcel("MS-Excel Workbooks", "*.xlsm;*.xlsb;*.xls"));
            Add(new ProjectFilterExcel("MS-Excel Add-Ins",   "*.xlam;*.xla"));


        //    if (new AccessWrapper()).IsAccessSupported {
        //       Add(new ProjectFilterAccess("MS-Access Databases", "*.mdb;*.accdb"));
        //       Add(new ProjectFilterAccess("MS-Access Add-Ins",   "*.mda;*.accda"));
        //    }
        }
    }
}

using System.Collections.Generic;

namespace PGSolutions.BetterRibbon.VbaSourceExport {
    internal class ProjectFilters : List<ProjectFilter> {
        public ProjectFilters() {
            Add(new ProjectFilterExcel("MS-Excel Projects", "*.xlsm;*.xlsb;*.xlam;*.xls;*.xla"));

            if (AccessWrapper.IsAccessSupported) {
                Add(new ProjectFilterAccess("MS-Access Projects", "*.accdb;*.accda;*.mdb;*.mda"));
            }
        }
    }
}

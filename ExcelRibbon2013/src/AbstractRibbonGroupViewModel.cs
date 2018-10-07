////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.AbstractCOM;

namespace PGSolutions.ExcelRibbon2013 {
    internal abstract class AbstractRibbonGroupViewModel {
        protected AbstractRibbonGroupViewModel(IRibbonFactory factory) => Factory = factory;

        public IRibbonFactory Factory { get; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")]
        public string NoImage => "MacroSecurity";
    }
}

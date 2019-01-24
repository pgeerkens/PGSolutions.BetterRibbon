////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.BetterRibbon {
    internal abstract class AbstractRibbonGroupViewModel {
        protected AbstractRibbonGroupViewModel(IRibbonFactory factory) => Factory = factory;

        protected IRibbonFactory Factory { get; }

        protected static string NoImage => "MacroSecurity";
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using static Microsoft.Office.Core.RibbonControlSize;

namespace PGSolutions.BetterRibbon {
    internal abstract class AbstractRibbonGroupViewModel {
        protected AbstractRibbonGroupViewModel(IRibbonFactory factory) => Factory = factory;

        public IRibbonFactory Factory { get; }

        public static string NoImage => "MacroSecurity";

        protected static bool ToggleButtonSize(bool isLarge, IList<IRibbonButton> buttons) {
            foreach (var b in buttons) { b.Size = isLarge ? RibbonControlSizeLarge : RibbonControlSizeRegular; }
            return !isLarge;
        }
    }
}

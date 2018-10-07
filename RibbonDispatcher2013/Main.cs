////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher2013.AbstractCOM;
using PGSolutions.RibbonDispatcher2013.ConcreteCOM;
using PGSolutions.RibbonDispatcher2013.Utilities;

namespace PGSolutions.RibbonDispatcher2013 {

    /// <summary>The publicly available entry points to the library.</summary>
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMain))]
    [Guid(Guids.Main)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    public class Main : IMain {
        private static Lazy<Dictionary<string,IRibbonUI>> RibbonCollection =
                new Lazy<Dictionary<string, IRibbonUI>>( () => new Dictionary<string, IRibbonUI>() );

        /// <inheritdoc/>
        public IRibbonViewModel NewRibbonViewModel(IRibbonUI ribbonUI)
            => new RibbonViewModel(ribbonUI);

        /// <inheritdoc/>
        public IRibbonUI SetRibbonUI(IRibbonUI ribbonUI, string workbookPath) {
            RibbonCollection.Value.AddNotNull(workbookPath,ribbonUI);
            return ribbonUI;
        }

        /// <inheritdoc/>
        public IRibbonUI GetRibbonUI(string WorkbookPath) =>
            RibbonCollection.Value.GetOrDefault(WorkbookPath);
    }
}

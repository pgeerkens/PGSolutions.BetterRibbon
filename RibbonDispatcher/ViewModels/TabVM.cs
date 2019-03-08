////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System.Collections.Generic;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public class TabVM: AbstractContainerVM<IControlSource,ITabVM>, ITabVM, 
            IActivatable<IControlSource,ITabVM> {
        internal TabVM(string itemId, IEnumerable<IControlVM> controls)
        : base(itemId,controls) { }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override ITabVM Attach(IControlSource source)
        => Attach<TabVM>(source);
    }
}

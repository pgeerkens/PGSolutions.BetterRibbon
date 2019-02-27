////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    internal class TabVM: AbstractContainerVM<IControlSource>, ITabVM, 
            IActivatable<IControlSource, TabVM> {
        public TabVM(ViewModelFactory factory, string itemId) : base(factory, itemId) { }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        TabVM IActivatable<IControlSource, TabVM>.Attach(IControlSource source)
        => Attach<TabVM>(source);
    }
}

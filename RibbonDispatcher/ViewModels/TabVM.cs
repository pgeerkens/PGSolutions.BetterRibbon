////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    public class TabVM: AbstractContainerVM<IControlSource,ITabVM>, ITabVM, 
            IActivatable<IControlSource,ITabVM> {
        internal TabVM(ViewModelFactory factory, string itemId) : base(itemId) { }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override ITabVM Attach(IControlSource source)
        => Attach<TabVM>(source);
    }
}

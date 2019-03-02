////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ViewModels {

    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [CLSCompliant(true)]
    public class GroupVM : AbstractContainerVM<IControlSource,IGroupVM>, IGroupVM, 
            IActivatable<IControlSource,IGroupVM> {
        internal protected GroupVM(ViewModelFactory factory, string itemId)
        : base(factory, itemId) { }

        public TControl GetControl<TControl>(string controlId) where TControl : class, IControlVM
        => Factory.GetControl<TControl>(controlId);

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override IGroupVM Attach(IControlSource source) => Attach<GroupVM>(source);
    }

    internal class KeyedControls: KeyedCollection<string, IControlVM> {
        public KeyedControls() : base() { }
        protected override string GetKeyForItem(IControlVM control) => control?.Id;
    }
}

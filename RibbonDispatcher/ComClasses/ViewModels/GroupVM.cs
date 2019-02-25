////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.ObjectModel;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {

    internal class GroupVM : AbstractContainerVM<IControlSource>, IGroupVM, 
            IActivatable<IControlSource,GroupVM> {
        public GroupVM(IViewModelFactory factory, string itemId)
        : base(factory, itemId) { }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        GroupVM IActivatable<IControlSource,GroupVM>.Attach(IControlSource source)
        => Attach<GroupVM>(source);

        public override string Description
        => throw new InvalidOperationException("Attribute Description not supported on a Group.");
    }

    internal class KeyedControls: KeyedCollection<string, IActivatable> {
        public KeyedControls() : base() { }
        protected override string GetKeyForItem(IActivatable control) => control?.Id;
    }
}

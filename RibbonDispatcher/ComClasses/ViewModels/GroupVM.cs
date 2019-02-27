////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.ObjectModel;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {

    public class GroupVM : AbstractContainerVM<IControlSource>, IGroupVM, 
            IActivatable<IControlSource,GroupVM> {
        internal GroupVM(ViewModelFactory factory, string itemId)
        : base(factory, itemId) { }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        GroupVM IActivatable<IControlSource,GroupVM>.Attach(IControlSource source)
        => Attach<GroupVM>(source);
    }

    internal class KeyedControls: KeyedCollection<string, IControlVM> {
        public KeyedControls() : base() { }
        protected override string GetKeyForItem(IControlVM control) => control?.Id;
    }
}

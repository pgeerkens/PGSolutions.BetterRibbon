////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    internal class SplitButtonVM: AbstractContainerVM<ISplitButtonSource>, ISplitButtonVM,
             IActivatable<ISplitButtonSource, ISplitButtonVM> {
        public SplitButtonVM(IViewModelFactory factory, string itemId, IButtonVM button, IMenuVM menu)
        : base(factory,itemId) {
            ButtonVM = button;
            MenuVM   = menu;
        }

        public IButtonVM ButtonVM { get; }
        public IMenuVM   MenuVM   { get; }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public new ISplitButtonVM Attach(ISplitButtonSource source) => Attach<SplitButtonVM>(source);

        public override void Invalidate() {
            ButtonVM?.Invalidate();
            MenuVM?.Invalidate();
            base.Invalidate();
        }

        public override string Description
        => throw new InvalidOperationException("Attribute Description not supported on a Split Button.");
    }
}

﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    internal class SplitVM: AbstractContainerVM<IControlSource>, ILabelVM,
             IActivatable<IControlSource, SplitVM> {
        public SplitVM(IViewModelFactory factory, string itemId, IButtonVM button, IMenuVM menu)
        : base(factory,itemId) {
            SplitButtonVM = button;
            SplitMenuVM   = menu;
        }

        public IButtonVM SplitButtonVM { get; }
        public IMenuVM   SplitMenuVM   { get; }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        SplitVM IActivatable<IControlSource, SplitVM>.Attach(IControlSource source)
        => Attach<SplitVM>(source);

        public override string Description
        => throw new InvalidOperationException("Attribute Description not supported on a Split Button.");
    }
}
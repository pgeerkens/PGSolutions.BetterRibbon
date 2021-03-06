﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IModels  = IReadOnlyList<ICanInvalidate>;

    [CLSCompliant(true)]
    public abstract class AbstractRibbonTabModel {
        protected AbstractRibbonTabModel(IRibbonViewModel viewModel, IReadOnlyList<ICanInvalidate> models) {
            ViewModel = viewModel;
            Models    = models;
        }

        public    IRibbonViewModel ViewModel { get; }

        protected IModels          Models    { get; }

        public void Invalidate() => Models.ToList().ForEach(model => model.Invalidate());

        /// <inheritdoc/>
        public void DetachProxy(string controlId) => ViewModel.GetControl<IControlVM>(controlId)?.Detach();
    }
}

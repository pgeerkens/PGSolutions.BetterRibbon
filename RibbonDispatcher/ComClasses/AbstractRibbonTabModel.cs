////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IControlStrings;

    [CLSCompliant(false)]
    public abstract class AbstractRibbonTabModel {
        protected AbstractRibbonTabModel(AbstractRibbonViewModel viewModel, IReadOnlyList<IInvalidate> models) {
            ViewModel = viewModel;
            Models    = models;
        }

        public    AbstractRibbonViewModel    ViewModel { get; }

        protected IReadOnlyList<IInvalidate> Models    { get; private set; }

        private List<CustomButtonsGroupModel> CustomButtonsModel
        => Models.OfType<CustomButtonsGroupModel>().ToList();

        public void Invalidate() { foreach (var model in Models) { model?.Invalidate(); } }

        /// <inheritdoc/>
        internal void DetachProxy(string controlId) => GetControl<IRibbonControlVM>(controlId)?.Detach();

        private TControl GetControl<TControl>(string controlId) where TControl : class, IRibbonControlVM
        => ViewModel.RibbonFactory.GetControl<TControl>(controlId);

        public void DetachCustomControls()
        => CustomButtonsModel.ForEach(model => model.DetachControls());

        protected IStrings GetStrings(string id) => ViewModel.RibbonFactory.ResourceManager.GetControlStrings(id);
    }
}

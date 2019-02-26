////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using System.Linq;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IModels  = IReadOnlyList<ICanInvalidate>;

    public abstract class AbstractRibbonTabModel {
        protected AbstractRibbonTabModel(IRibbonViewModel viewModel, IReadOnlyList<ICanInvalidate> models) {
            ViewModel = viewModel;
            Models    = models;
        }

        public  IRibbonViewModel ViewModel { get; }

        private IModels          Models    { get; }

        public void Invalidate()
        => Models.ToList().ForEach(model => model.Invalidate());

        /// <inheritdoc/>
        internal void DetachProxy(string controlId)
        => ViewModel.ViewModelFactory.GetControl<IControlVM>(controlId)?.Detach();

        public void DetachCustomControls()
        => Models.OfType<CustomButtonsGroupModel>().ToList().ForEach(model => model.DetachControls());
    }
}

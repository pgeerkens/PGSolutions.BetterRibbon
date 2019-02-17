////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IRibbonControlStrings;

    [CLSCompliant(false)]
    public abstract class AbstractRibbonTabModel {
        protected AbstractRibbonTabModel(AbstractRibbonViewModel viewModel, IReadOnlyList<IInvalidate> models) {
            ViewModel = viewModel;
            Models    = models;
        }

        public    AbstractRibbonViewModel    ViewModel          { get; }

        protected IReadOnlyList<IInvalidate> Models             { get; }

        public void Invalidate() { foreach (var model in Models) { model?.Invalidate(); } }

        protected abstract AbstractRibbonGroupModel CustomButtons1Model { get; }

        /// <inheritdoc/>
        public IRibbonButtonModel NewRibbonButtonModel(IStrings strings,
                IPictureDisp image, bool isEnabled, bool isVisible)
        => CustomButtons1Model.NewButtonModel(strings, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        public IRibbonButtonModel NewRibbonButtonModel(IStrings strings,
                string imageMso, bool isEnabled, bool isVisible)
        => CustomButtons1Model.NewButtonModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        public IRibbonToggleModel NewRibbonToggleModel(IStrings strings, IPictureDisp image, bool isEnabled, bool isVisible)
        => CustomButtons1Model.NewToggleModel(strings, new ImageObject(image), isEnabled, isVisible);

        /// <inheritdoc/>
        public IRibbonToggleModel NewRibbonToggleModel(IStrings strings, string imageMso, bool isEnabled, bool isVisible)
        => CustomButtons1Model.NewToggleModel(strings, imageMso, isEnabled, isVisible);

        /// <inheritdoc/>
        public IRibbonDropDownModel NewRibbonDropDownModel(IStrings strings, bool isEnabled, bool isVisible)
        => CustomButtons1Model.NewDropDownModel(strings, isEnabled, isVisible);

        /// <inheritdoc/>
        public RibbonGroupModel NewRibbonGroupModel(IStrings strings, bool isEnabled, bool isVisible)
        => new RibbonGroupModel(CustomButtons1Model.GetControl<RibbonGroupViewModel>, strings, isEnabled, isVisible, CustomButtons1Model);
    }
}

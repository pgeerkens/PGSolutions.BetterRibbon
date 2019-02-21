////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    /// <summary>The (top-level) TabModel for the ribbon interface.</summary>
    [CLSCompliant(false)]
    public sealed class BetterRibbonModel : AbstractRibbonTabModel {
        internal BetterRibbonModel(BetterRibbonViewModel viewModel)
        : base(viewModel) {
            Initialize(new List<IInvalidate> {
                new BrandingModel(viewModel.GroupViewModels.FirstOrDefault(vm => vm.Id == "BrandingGroup"), this),
                new LinksAnalysisModel(viewModel.GroupViewModels.FirstOrDefault(vm => vm.Id == "LinksAnalysisGroup"), this),
                new VbaSourceExportModel( new List<VbaSourceExportGroupModel>() {
                    new VbaSourceExportGroupModel(viewModel.GroupViewModels.FirstOrDefault(vm => vm.Id == "VbaExportGroupMS"), this, "MS"),
                    new VbaSourceExportGroupModel(viewModel.GroupViewModels.FirstOrDefault(vm => vm.Id == "VbaExportGroupPG"), this, "PG")
                } ),
                new CustomButtonsGroup1Model(viewModel.GroupViewModels.FirstOrDefault(vm => vm.Id == "CustomizableGroup"))
            }.AsReadOnly());
            CustomButtons1Model = Models.OfType<CustomButtonsGroup1Model>().FirstOrDefault();
        }

        /// <summary>.</summary>
        protected override AbstractRibbonGroupModel CustomButtons1Model { get; }

        internal void     DetachCustomControls() => CustomButtons1Model?.DetachControls();
    }
}

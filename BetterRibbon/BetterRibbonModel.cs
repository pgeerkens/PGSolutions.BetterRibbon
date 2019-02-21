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
        : base(viewModel, new List<IInvalidate> {
                new BrandingModel(viewModel.GroupViewModels.FirstOrDefault(
                            vm => vm.Id == "BrandingGroup"), viewModel.RibbonFactory),
                new LinksAnalysisModel(viewModel.GroupViewModels.FirstOrDefault(
                            vm => vm.Id == "LinksAnalysisGroup"), viewModel.RibbonFactory),
                new VbaSourceExportModel( new List<VbaSourceExportGroupModel>() {
                    new VbaSourceExportGroupModel(viewModel.GroupViewModels.FirstOrDefault(
                            vm => vm.Id == "VbaExportGroupMS"), viewModel.RibbonFactory, "MS"),
                    new VbaSourceExportGroupModel(viewModel.GroupViewModels.FirstOrDefault(
                            vm => vm.Id == "VbaExportGroupPG"), viewModel.RibbonFactory, "PG")
                } ),
                new CustomButtonsGroupModel(viewModel.GroupViewModels.FirstOrDefault(
                            vm => vm.Id == "CustomizableGroup"))
            }.AsReadOnly()) { }
    }
}

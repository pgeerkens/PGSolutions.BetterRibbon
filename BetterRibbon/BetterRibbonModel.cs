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
        internal BetterRibbonModel(AbstractDispatcher viewModel)
        : base(viewModel, new List<IInvalidatible> {
                new BrandingModel(viewModel, "BrandingGroup"),
                new LinksAnalysisModel(viewModel, "LinksAnalysisGroup"),
                new VbaSourceExportModel( new List<VbaSourceExportGroupModel>() {
                    new VbaSourceExportGroupModel(viewModel, "VbaExportGroupMS", "MS"),
                    new VbaSourceExportGroupModel(viewModel, "VbaExportGroupPG", "PG")
                } ),
                new CustomButtonsGroupModel(viewModel, "CustomizableGroup")
            }.AsReadOnly()) { }
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    /// <summary>The (top-level) TabModel for the ribbon interface.</summary>
    [CLSCompliant(false)]
    public sealed class BetterRibbonModel : AbstractRibbonTabModel {
        internal BetterRibbonModel(IRibbonViewModel viewModel, AbstractModelFactory factory)
        : base(viewModel, new List<ICanInvalidate> {
                new BrandingModel(viewModel, factory, "BrandingGroup"),
                new LinksAnalysisModel(viewModel, factory, "LinksAnalysisGroup"),
                new VbaSourceExportModel( new List<VbaSourceExportGroupModel>() {
                    new VbaSourceExportGroupModel(viewModel, factory, "VbaExportGroupMS", "MS"),
                    new VbaSourceExportGroupModel(viewModel, factory, "VbaExportGroupPG", "PG")
                } ),
                new CustomButtonsGroupModel(viewModel, "CustomizableGroup")
            }.AsReadOnly()) { }
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.UtilityRibbon {
    /// <summary>The (top-level) TabModel for the ribbon interface.</summary>
    [CLSCompliant(false)]
    internal sealed class Model : AbstractRibbonTabModel {
        public Model(ViewModel viewModel, IModelFactory factory)
        : base(viewModel, new List<ICanInvalidate> {
                new BrandingModel(factory, viewModel.BrandingGroupVM),
                new LinksAnalysisModel(factory, viewModel.LinkedAnalysisGroupVM),
                new VbaSourceExportModel( new List<VbaSourceExportGroupModel>() {
                    new VbaSourceExportGroupModel(factory, viewModel.VbaExportGroupVM_MS, "MS"),
                    new VbaSourceExportGroupModel(factory, viewModel.VbaExportGroupVM_PG, "PG")
                } )
            }.AsReadOnly())
        { }
    }
}

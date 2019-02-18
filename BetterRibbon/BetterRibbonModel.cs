////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    /// <summary>The (top-level) Model for the ribbon interface.</summary>
    [CLSCompliant(false)]
    public sealed class BetterRibbonModel : AbstractRibbonTabModel {
        internal BetterRibbonModel(BetterRibbonViewModel viewModel)
        : base(viewModel, new List<IInvalidate> {
                new BrandingModel(viewModel.AddGroupViewModel("BrandingGroup")),
                new LinksAnalysisModel(viewModel.AddGroupViewModel("LinksAnalysisGroup")),
                new VbaSourceExportModel( new List<VbaSourceExportGroupModel>() {
                    new VbaSourceExportGroupModel(viewModel.AddGroupViewModel("VbaExportGroupMS"),"MS"),
                    new VbaSourceExportGroupModel(viewModel.AddGroupViewModel("VbaExportGroupPG"),"PG")
                } ),
                new CustomButtonsGroup1Model(viewModel.AddGroupViewModel(NewCustomButtonsViewModel))
        }.AsReadOnly())
        => CustomButtons1Model = Models.OfType<CustomButtonsGroup1Model>().FirstOrDefault();

        /// <summary>.</summary>
        protected override AbstractRibbonGroupModel CustomButtons1Model { get; }

        private static RibbonGroupViewModel NewCustomButtonsViewModel(IRibbonFactory factory)
        => factory.NewRibbonGroup("CustomizableGroup")
                .Add<IRibbonToggleSource>(factory.NewRibbonToggle("CustomVbaToggle1"))
                .Add<IRibbonToggleSource>(factory.NewRibbonToggle("CustomVbaToggle2"))
                .Add<IRibbonToggleSource>(factory.NewRibbonToggle("CustomVbaToggle3"))

                .Add<IRibbonToggleSource>(factory.NewRibbonCheckBox("CustomVbaCheckBox1"))
                .Add<IRibbonToggleSource>(factory.NewRibbonCheckBox("CustomVbaCheckBox2"))
                .Add<IRibbonToggleSource>(factory.NewRibbonCheckBox("CustomVbaCheckBox3"))

                .Add<IRibbonDropDownSource>(factory.NewRibbonDropDown("CustomVbaDropDown1"))
                .Add<IRibbonDropDownSource>(factory.NewRibbonDropDown("CustomVbaDropDown2"))
                .Add<IRibbonDropDownSource>(factory.NewRibbonDropDown("CustomVbaDropDown3"))

                .Add<IRibbonButtonSource>(factory.NewRibbonButton("CustomizableButton1"))
                .Add<IRibbonButtonSource>(factory.NewRibbonButton("CustomizableButton2"))
                .Add<IRibbonButtonSource>(factory.NewRibbonButton("CustomizableButton3"));

        internal TControl GetCustomControl<TControl>(string controlId) where TControl : class, IRibbonCommon
        => CustomButtons1Model.GetControl<TControl>(controlId);

        internal void     DetachCustomControls()
        => CustomButtons1Model?.DetachControls();
    }
}

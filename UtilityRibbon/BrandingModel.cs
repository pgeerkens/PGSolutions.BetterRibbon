////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Text;

using PGSolutions.RibbonDispatcher;
using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

using PGSolutions.RibbonUtilities.VbaSourceExport;

using PGSolutions.UtilityRibbon.Properties;

namespace PGSolutions.UtilityRibbon {
    internal sealed class BrandingModel : AbstractRibbonGroupModel {
        public BrandingModel(IModelFactory factory, IGroupVM viewModel)
        : base(viewModel, factory.GetStrings(viewModel.ControlId)) {
            BrandingButtonModel = factory.NewButtonModel("BrandingButton", ButtonClicked,
                factory.GetImage(Resources.PGeerkens.ImageToPictureDisp()));

            Invalidate();
        }

        private IButtonModel BrandingButtonModel { get; }

        private void ButtonClicked(object sender) => new StringBuilder()
            .AppendLine($"PGSolutions Better Ribbon")
            .AppendLine()
            .AppendLine($"Better Ribbon V {ThisAddIn.VersionNo3}")
            .AppendLine($"Ribbon Utilities V {UtilitiesVersion.Format2()}")
            .AppendLine($"Ribbon ModelFactory V {DispatcherVersion.Format2()}")
            .AppendLine()
            .AppendLine($"{BrandingButtonModel.SuperTip}")
        #if DEBUG
            .AppendLine()
            .AppendLine("***  DEBUG build  ***")
        #endif
            .ToString().MsgBoxShow();

        static Version DispatcherVersion => typeof(ViewModelFactory).Assembly.GetName().Version;
        static Version UtilitiesVersion  => typeof(VbaExportEventArgs).Assembly.GetName().Version;
    }
}

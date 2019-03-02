////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Text;
using stdole;

using PGSolutions.RibbonDispatcher;
using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

using PGSolutions.RibbonUtilities.VbaSourceExport;

using PGSolutions.BetterRibbon.Properties;

namespace PGSolutions.BetterRibbon {
    internal sealed class BrandingModel : AbstractRibbonGroupModel {
        public BrandingModel(IModelFactory factory, IGroupVM viewModel)
        : base(viewModel, factory.GetStrings(viewModel.Id)) {
            BrandingButtonModel = factory.NewButtonModel("BrandingButton", ButtonClicked,
                new ImageObject(BrandingIcon));

            Invalidate();
        }

        private IButtonModel BrandingButtonModel { get; }

        private void ButtonClicked(object sender) => new StringBuilder()
            .AppendLine($"PGSolutions Better Ribbon")
            .AppendLine()
            .AppendLine($"Better Ribbon V {Globals.ThisAddIn.VersionNo3}")
            .AppendLine($"Ribbon Utilities V {UtilitiesVersion.Format2()}")
            .AppendLine($"Ribbon ModelFactory V {DispatcherVersion.Format2()}")
            .AppendLine()
            .AppendLine($"{BrandingButtonModel.Strings.SuperTip}")
        #if DEBUG
            .AppendLine()
            .AppendLine("***  DEBUG build  ***")
        #endif
            .ToString().MsgBoxShow();

        static Version DispatcherVersion => typeof(ViewModelFactory).Assembly.GetName().Version;
        static Version UtilitiesVersion  => typeof(VbaExportEventArgs).Assembly.GetName().Version;

        private static IPictureDisp BrandingIcon => Resources.PGeerkens.ImageToPictureDisp();
    }
}

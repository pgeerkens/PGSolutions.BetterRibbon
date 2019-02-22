﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Text;
using stdole;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.VbaSourceExport;
using PGSolutions.BetterRibbon.Properties;

namespace PGSolutions.BetterRibbon {
    internal sealed class BrandingModel : AbstractRibbonGroupModel {
        public BrandingModel(GroupVM viewModel, IRibbonFactory factory) : base(viewModel) {
            BrandingButtonModel = factory.NewRibbonButtonModel("BrandingButton", ButtonClicked,
                new ImageObject(BrandingIcon));

            Invalidate();
        }

        private IRibbonButtonModel BrandingButtonModel { get; }

        private void ButtonClicked(object sender, EventArgs e) => new StringBuilder()
            .AppendLine($"PGSolutions Better Ribbon")
            .AppendLine()
            .AppendLine($"Better Ribbon V {Globals.ThisAddIn.VersionNo3}")
            .AppendLine($"Ribbon Utilities V {UtilitiesVersion.Format2()}")
            .AppendLine($"Ribbon Dispatcher V {DispatcherVersion.Format2()}")
            .AppendLine()
            .AppendLine($"{BrandingButtonModel.Strings.SuperTip}")
        #if DEBUG
            .AppendLine()
            .AppendLine("***  DEBUG build  ***")
        #endif
            .ToString().MsgBoxShow();

        static Version DispatcherVersion => new RibbonFactory().GetType().Assembly.GetName().Version;
        static Version UtilitiesVersion  => new VbaExportEventArgs(null).GetType().Assembly.GetName().Version;

        private static IPictureDisp BrandingIcon => Resources.PGeerkens.ImageToPictureDisp();
    }
}

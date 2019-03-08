////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Deployment.Application;
using System.Text;

using PGSolutions.RibbonDispatcher;
using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

using PGSolutions.RibbonUtilities.VbaSourceExport;

namespace PGSolutions.ToolsRibbon {
    internal sealed class BrandingModel : AbstractRibbonGroupModel {
        public BrandingModel(IModelFactory factory, IGroupVM viewModel)
        : base(viewModel, factory.GetStrings(viewModel.ControlId)) {
            BrandingButtonModel = factory.NewButtonModel("BrandingButton", ButtonClicked,
                factory.GetImage(Properties.Resources.BrandingImage.ImageToPictureDisp()));

            Invalidate();
        }

        private IButtonModel BrandingButtonModel { get; }

        private void ButtonClicked(object sender) => new StringBuilder()
            .AppendLine($"PGSolutions Better Ribbon")
            .AppendLine()
            .AppendLine($"Better Ribbon V {ThisVersion?.Format()}")
            .AppendLine($"RibbonUtilities V {UtilitiesVersion.Format()}")
            .AppendLine($"RibbonDispatcher V {DispatcherVersion.Format()}")
            .AppendLine()
            .AppendLine($"{BrandingButtonModel.SuperTip}")
        #if DEBUG
            .AppendLine()
            .AppendLine("***  DEBUG build  ***")
        #endif
            .ToString().MsgBoxShow();

        static Version ThisVersion       => typeof(BrandingModel).Assembly.GetName().Version;
        static Version DispatcherVersion => typeof(ViewModelFactory).Assembly.GetName().Version;
        static Version UtilitiesVersion  => typeof(VbaExportEventArgs).Assembly.GetName().Version;

        /// <summary>.</summary>
        static string VersionNo => ApplicationDeployment.IsNetworkDeployed
            ? ApplicationDeployment.CurrentDeployment.CurrentVersion?.Format()
            : new Version(0,0,0,0).Format();

        /// <summary>.</summary>
        static string WindowsFormsVersionNo => System.Windows.Forms.Application.ProductVersion;
    }
}

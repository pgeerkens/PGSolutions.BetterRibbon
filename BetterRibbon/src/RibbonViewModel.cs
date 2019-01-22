////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.Utilities;
using PGSolutions.BetterRibbon.VbaSourceExport;
using BetterRibbon.Properties;

namespace PGSolutions.BetterRibbon {
    /// <summary>The (top-level) ViewModel for the ribbon interface.</summary>
    /// <remarks>
    /// <a href=" https://go.microsoft.com/fwlink/?LinkID=271226">For more information about adding callback methods.</a>
    /// 
    /// Take care renaming this class, or its namespace; and coordinate any such with the content of the (hidden)
    /// ThisAddIn.Designer.xml file. Commit frequently. Excel is very tempermental on the naming of ribbon objects
    /// and provides poor, and very minimal, diagnostic information.
    /// 
    /// This class MUST be ComVisible for the ribbon to launch properly.
    /// </remarks>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [Guid("A8ED8DFB-C422-4F03-93BF-FB5453D8F213")]
    public sealed class RibbonViewModel : AbstractRibbonViewModel, IRibbonExtensibility {
        const string _assemblyName  = "BetterRibbon";

        internal RibbonViewModel() : base(new LocalResourceManager(_assemblyName)) {}

        internal BrandingViewModel            BrandingViewModel        { get; private set; }
        internal VbaSourceExportModel         VbaSourceExportModel     { get; private set; }
        //internal DemonstrationModel           DemonstrationModel       { get; private set; }
        internal DemonstrationViewModel       DemonstrationModel       { get; private set; }
        internal CustomizableButtonsViewModel CustomButtonsViewModel   { get; private set; }

        internal IReadOnlyDictionary<string, IActivatable> AdaptorControls =>
                CustomButtonsViewModel.AdaptorControls;

        public string GetCustomUI(string RibbonID) => Resources.Ribbon;

        [SuppressMessage("Microsoft.Design", "CA1061:DoNotHideBaseClassMethods",
                Justification="False positive - parameter types are identical.")]
        [CLSCompliant(false)]
        public sealed override void OnRibbonLoad(IRibbonUI ribbonUI) {
            base.OnRibbonLoad(ribbonUI);

            BrandingViewModel    = new BrandingViewModel(RibbonFactory, GetBrandingIcon);
            VbaSourceExportModel = new VbaSourceExportModel(
                new List<IVbaSourceExportGroupModel> {
                    new VbaSourceExportViewModel(RibbonFactory, "MS"),
                    new VbaSourceExportViewModel(RibbonFactory, "PG")
                } );

            CustomButtonsViewModel = new CustomizableButtonsViewModel(RibbonFactory);

            //DemonstrationModel = new DemonstrationModel( new DemonstrationViewModel(RibbonFactory) );
            DemonstrationModel = new DemonstrationViewModel(RibbonFactory);

            Invalidate();
        }

        private static IPictureDisp GetBrandingIcon() => Resources.PGeerkens.ImageToPictureDisp();

        public static string MsgBoxTitle => Resources.ApplicationName;
    }
}

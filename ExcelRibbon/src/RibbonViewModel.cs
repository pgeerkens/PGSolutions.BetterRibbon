////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using stdole;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.Utilities;
using PGSolutions.ExcelRibbon.VbaSourceExport;
using ExcelRibbon.Properties;

namespace PGSolutions.ExcelRibbon {
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
    [ComVisible(true)]
    [CLSCompliant(true)]
    [Guid("A8ED8DFB-C422-4F03-93BF-FB5453D8F213")]
    public sealed class RibbonViewModel : AbstractRibbonViewModel, IRibbonExtensibility {
        const string _assemblyName  = "ExcelRibbon";

        public RibbonViewModel() :base(new LocalResourceManager(_assemblyName)) { }

        internal BrandingViewModel        BrandingViewModel        { get; private set; }
        internal VbaSourceExportViewModel VbaExportGroupMS         { get; private set; }
        internal VbaSourceExportViewModel VbaExportGroupPG         { get; private set; }
        internal VbaSourceExportViewModel StandardButtonsViewModel { get; private set; }
        internal CustomButtonsViewModel   CustomButtonsViewModel   { get; private set; }

        public string GetCustomUI(string RibbonID) => Resources.Ribbon;

        [CLSCompliant(false)]
        public override void OnRibbonLoad(IRibbonUI ribbonUI) {
            base.OnRibbonLoad(ribbonUI);

            BrandingViewModel       = new BrandingViewModel(RibbonFactory, GetBrandingIcon);
            CustomButtonsViewModel  = new CustomButtonsViewModel(RibbonFactory);
            VbaExportGroupMS        = new VbaSourceExportViewModel(RibbonFactory, "MS");
            VbaExportGroupPG        = new VbaSourceExportViewModel(RibbonFactory, "PG");
            Invalidate();
        }

        private static IPictureDisp GetBrandingIcon() => Resources.RD_AboutWindow.ImageToPictureDisp();

        public static string MsgBoxTitle => Resources.ApplicationName;
    }
}

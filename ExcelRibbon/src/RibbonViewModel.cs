////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using stdole;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ConcreteCOM;
using PGSolutions.RibbonDispatcher.Utilities;
using PGSolutions.ExcelRibbon.VbaSourceExport;
using PGSolutions.ExcelRibbon.Properties;

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
        const string _AssemblyName  = "ExcelRibbon";

        public RibbonViewModel() {;}

        internal BrandingViewModel        BrandingViewModel        { get; private set; }
        internal VbaSourceExportViewModel StandardButtonsViewModel { get; private set; }
        internal CustomButtonsViewModel   CustomButtonsViewModel   { get; private set; }

        public string GetCustomUI(string RibbonID) => Resources.Ribbon;

        [CLSCompliant(false)]
        public  void OnRibbonLoad(IRibbonUI ribbonUI) {
            Initialize(ribbonUI, this);

            BrandingViewModel        = new BrandingViewModel(RibbonFactory, GetBrandingIcon);
            CustomButtonsViewModel   = new CustomButtonsViewModel(RibbonFactory);
            StandardButtonsViewModel = new VbaSourceExportViewModel(RibbonFactory);
            Invalidate();
        }

        private static IPictureDisp GetBrandingIcon() => Resources.RD_AboutWindow.ImageToPictureDisp();

        public static string MsgBoxTitle => Resources.ApplicationName;

        protected override Lazy<ResourceManager> ResourceManager => new Lazy<ResourceManager>(
            () => new ResourceManager($"{_AssemblyName}.Properties.Resources", Assembly.GetExecutingAssembly())
        );
    }
}

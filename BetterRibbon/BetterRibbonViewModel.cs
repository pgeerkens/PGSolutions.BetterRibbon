////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using PGSolutions.RibbonDispatcher.ComClasses;
using BetterRibbon.Properties;
using PGSolutions.RibbonDispatcher.ComInterfaces;

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
    //[Guid("A8ED8DFB-C422-4F03-93BF-FB5453D8F213")]
    public sealed class BetterRibbonViewModel : AbstractRibbonViewModel, IRibbonExtensibility {
        const string _assemblyName  = "BetterRibbon";

        internal BetterRibbonViewModel() : base(new LocalResourceManager(_assemblyName)) {
            Id = "TabPGSolutions";

            BrandingViewModel      = RibbonFactory.NewRibbonGroup("BrandingGroup");
            LinksAnalysisViewModel = RibbonFactory.NewRibbonGroup("LinksAnalysisGroup");
            VbaExportViewModel_PG  = RibbonFactory.NewRibbonGroup("VbaExportGroupPG");
            VbaExportViewModel_MS  = RibbonFactory.NewRibbonGroup("VbaExportGroupMS");
            CustomButtonsViewModel = RibbonFactory.NewCustomButtonsViewModel();

            DemonstrationViewModel = null;//RibbonFactory.Add(new DemonstrationViewModel(RibbonFactory));
        }

        internal RibbonGroupViewModel BrandingViewModel      { get; private set; }
        internal RibbonGroupViewModel CustomButtonsViewModel { get; private set; }
        internal RibbonGroupViewModel LinksAnalysisViewModel { get; private set; }
        internal RibbonGroupViewModel VbaExportViewModel_MS  { get; private set; }
        internal RibbonGroupViewModel VbaExportViewModel_PG  { get; private set; }

        internal DemonstrationViewModel   DemonstrationViewModel { get; private set; }

        internal void DetachControls() => CustomButtonsViewModel?.DetachControls();

        /// <summary>.</summary>
        public string GetCustomUI(string RibbonID) => Resources.Ribbon;

        /// <summary>.</summary>
        public event EventHandler Initialized;

         /// <summary>.</summary>
       public bool IsInitialized => RibbonUI != null;

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1061:DoNotHideBaseClassMethods",
                Justification="False positive - parameter types are identical.")]
        [CLSCompliant(false)]
        public sealed override void OnRibbonLoad(IRibbonUI ribbonUI) {
            base.OnRibbonLoad(ribbonUI);

            Initialized?.Invoke(this, EventArgs.Empty);

            Invalidate();
        }

        /// <summary>.</summary>
        protected override string Id { get; }

        /// <summary>.</summary>
        public static string MsgBoxTitle => Resources.ApplicationName;

    }

    internal static partial class Extensions {
        public static RibbonGroupViewModel NewCustomButtonsViewModel(this IRibbonFactory factory)
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

                .Add<IRibbonButtonSource>(factory.NewRibbonButtonMso("CustomizableButton1"))
                .Add<IRibbonButtonSource>(factory.NewRibbonButtonMso("CustomizableButton2"))
                .Add<IRibbonButtonSource>(factory.NewRibbonButtonMso("CustomizableButton3"));
    }
}

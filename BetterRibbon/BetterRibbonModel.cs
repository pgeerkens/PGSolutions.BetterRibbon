////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
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
    [Serializable, CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonDispatcher))]
    [Guid(RibbonDispatcher.Guids.BetterRibbon)]
    [ProgId(ProgIds.RibbonDispatcherProgId)]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported Events.")]
    public sealed class BetterRibbonModel {
        internal BetterRibbonModel(BetterRibbonViewModel viewModel) {
            ViewModel   = viewModel;

            BrandingModel        = new BrandingModel(ViewModel.AddGroupViewModel("BrandingGroup"), BrandingIcon);
            LinksAnalysisModel   = new LinksAnalysisModel(ViewModel.AddGroupViewModel("LinksAnalysisGroup"));
            VbaSourceExportModel = new VbaSourceExportModel(
                new List<VbaSourceExportGroupModel>() {
                    new VbaSourceExportGroupModel(ViewModel.AddGroupViewModel("VbaExportGroupMS"),"MS"),
                    new VbaSourceExportGroupModel(ViewModel.AddGroupViewModel("VbaExportGroupPG"),"PG")
                });

            CustomButtonsModel   = new CustomButtonsModel(ViewModel.AddGroupViewModel(NewCustomButtonsViewModel));
        }

        internal BetterRibbonViewModel ViewModel            { get; }

        internal BrandingModel         BrandingModel        { get; private set; }
        internal LinksAnalysisModel    LinksAnalysisModel   { get; private set; }
        internal VbaSourceExportModel  VbaSourceExportModel { get; private set; }
        internal CustomButtonsModel    CustomButtonsModel   { get; private set; }

        private static IPictureDisp BrandingIcon => Resources.PGeerkens.ImageToPictureDisp();

        internal IRibbonDispatcher Dispatcher => new Dispatcher(this);

        public static RibbonGroupViewModel NewCustomButtonsViewModel(IRibbonFactory factory)
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
    }
}
